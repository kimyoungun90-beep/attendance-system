import { json, requireAuth } from "../_lib/auth.js";
import { recalculateRoute, runBatch } from "../_lib/recalculate.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const result = await context.env.DB.prepare(`
    SELECT c.id, c.company AS route, c.month, c.cutoff_date, c.plan_file_name, c.attendance_file_name,
           c.plan_people, c.attendance_people, c.matched_people, c.match_rate,
           c.missing_count, c.missing_people, c.unexpected_count, c.unexpected_people,
           c.mismatch_count, c.mismatch_people, c.dayoff_excess_people,
           c.substitute_shortage_people, c.compensation_shortage_people,
           c.annual_leave_people, c.created_at
    FROM attendance_closures c
    WHERE EXISTS (
      SELECT 1 FROM attendance_monthly_employee_facts f WHERE f.closure_id = c.id
    )
    ORDER BY c.month DESC, c.company ASC, c.created_at DESC
    LIMIT 240
  `).all();

  return json({ items: result.results || [] });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const body = await context.request.json().catch(() => null);
  const route = String(body?.route || "");
  const month = String(body?.month || "");
  if (!body || !VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month)) {
    return json({ error: "경로와 대상 월을 확인해 주세요." }, 400);
  }

  const issueRows = Array.isArray(body.issueRows) ? body.issueRows.slice(0, 10000) : [];
  const mismatchRows = Array.isArray(body.mismatchRows) ? body.mismatchRows.slice(0, 10000) : [];
  const employeeFacts = Array.isArray(body.employeeFacts) ? body.employeeFacts.slice(0, 3000) : [];
  if (!employeeFacts.length) return json({ error: "저장할 직원별 월 마감 원본이 없습니다." }, 400);

  const existing = await context.env.DB.prepare(
    "SELECT id FROM attendance_closures WHERE company = ? AND month = ? LIMIT 1"
  ).bind(route, month).first();
  const replaced = Boolean(existing?.id);
  const closureId = existing?.id || crypto.randomUUID();
  const preservedOpeningRows = replaced
    ? await context.env.DB.prepare(`
        SELECT employee_id, imported_opening_balance_present,
               imported_opening_substitute, imported_opening_compensation
        FROM attendance_monthly_employee_facts
        WHERE closure_id = ?
      `).bind(closureId).all()
    : { results: [] };
  const preservedOpeningByEmployee = new Map((preservedOpeningRows.results || []).map((row) => [
    String(row.employee_id || "").trim(),
    {
      present: Number(row.imported_opening_balance_present || 0) > 0,
      substitute: roundHalf(row.imported_opening_substitute),
      compensation: roundHalf(row.imported_opening_compensation),
    },
  ]));

  const missingRows = issueRows.filter((row) => row.issueType === "missing_clock_in");
  const unexpectedRows = issueRows.filter((row) => row.issueType === "unexpected_clock_in");
  const values = [
    body.cutoffDate || endOfMonth(month),
    body.planFileName || "",
    body.attendanceFileName || "",
    toNumber(body.planPeople),
    toNumber(body.attendancePeople),
    toNumber(body.matchedPeople),
    toNumber(body.matchRate),
    missingRows.length,
    uniquePeople(missingRows),
    unexpectedRows.length,
    uniquePeople(unexpectedRows),
    mismatchRows.length,
    uniquePeople(mismatchRows),
    employeeFacts.filter((row) => Number(row.baseExcess || 0) > 0).length,
    0,
    0,
    employeeFacts.filter((row) => Number(row.annualLeaveUsed || 0) > 0).length,
  ];

  const initialStatements = [];
  if (replaced) {
    initialStatements.push(
      context.env.DB.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare("DELETE FROM route_substitute_allocations WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare("DELETE FROM attendance_monthly_employee_facts WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare("DELETE FROM attendance_mismatch_items WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare("DELETE FROM attendance_issue_items WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(closureId),
      context.env.DB.prepare(`
        UPDATE attendance_closures
        SET cutoff_date = ?, check_mode = 'auto', plan_file_name = ?, attendance_file_name = ?,
            plan_people = ?, attendance_people = ?, matched_people = ?, match_rate = ?,
            missing_count = ?, missing_people = ?, unexpected_count = ?, unexpected_people = ?,
            mismatch_count = ?, mismatch_people = ?, dayoff_excess_people = ?,
            substitute_shortage_people = ?, compensation_shortage_people = ?, annual_leave_people = ?,
            created_at = datetime('now')
        WHERE id = ?
      `).bind(...values, closureId)
    );
  } else {
    initialStatements.push(context.env.DB.prepare(`
      INSERT INTO attendance_closures
      (id, company, month, cutoff_date, check_mode, plan_file_name, attendance_file_name,
       plan_people, attendance_people, matched_people, match_rate,
       missing_count, missing_people, unexpected_count, unexpected_people,
       mismatch_count, mismatch_people, dayoff_excess_people,
       substitute_shortage_people, compensation_shortage_people, annual_leave_people)
      VALUES (?, ?, ?, ?, 'auto', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(closureId, route, month, ...values));
  }

  try {
    await context.env.DB.batch(initialStatements);

    const issueStatements = issueRows.map((row) => context.env.DB.prepare(`
      INSERT OR IGNORE INTO attendance_issue_items
      (closure_id, issue_type, route, store, employee_id, employee_name, issue_date, weekday,
       plan_status, actual_status, actual_in, changed_in, clock_status, result, reason, duplicate_plan_note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      closureId,
      row.issueType === "unexpected_clock_in" ? "unexpected_clock_in" : "missing_clock_in",
      route,
      row.store || "",
      row.employeeId || "",
      row.name || "",
      row.date || "",
      row.weekday || "",
      row.planStatus || "",
      row.actualStatus || "",
      row.actualIn || "",
      row.changedIn || "",
      row.clockStatus || "",
      row.result || "",
      row.reason || "",
      row.duplicatePlanNote || ""
    ));

    const missingStatements = missingRows.map((row) => context.env.DB.prepare(`
      INSERT OR IGNORE INTO attendance_missing_items
      (closure_id, company, store, employee_id, employee_name, missing_date, weekday,
       plan_status, actual_in, changed_in, clock_status, result, reason, duplicate_plan_note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      closureId, route, row.store || "", row.employeeId || "", row.name || "",
      row.date || "", row.weekday || "", row.planStatus || "", row.actualIn || "",
      row.changedIn || "", row.clockStatus || "", row.result || "", row.reason || "",
      row.duplicatePlanNote || ""
    ));

    const mismatchStatements = mismatchRows.map((row) => context.env.DB.prepare(`
      INSERT OR REPLACE INTO attendance_mismatch_items
      (closure_id, route, store, employee_id, employee_name, issue_date, weekday,
       plan_status, actual_status, actual_in, changed_in, result, reason, duplicate_plan_note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      closureId, route, row.store || "", row.employeeId || "", row.name || "",
      row.date || "", row.weekday || "", row.planStatus || "", row.actualStatus || "",
      row.actualIn || "", row.changedIn || "", row.result || "", row.reason || "",
      row.duplicatePlanNote || ""
    ));

    const factStatements = employeeFacts.map((row) => {
      const opening = preservedOpeningByEmployee.get(String(row.employeeId || "").trim()) || { present: false, substitute: 0, compensation: 0 };
      return context.env.DB.prepare(`
      INSERT INTO attendance_monthly_employee_facts
      (closure_id, route, month, store, employee_id, employee_name,
       base_allowance, basic_dayoff_used, explicit_sub_dayoff_used, base_excess,
       substitute_needed, compensation_leave_used, compensation_needed,
       annual_leave_used, substitute_events_json, compensation_events_json,
       annual_leave_events_json, worked_dates_json, occurrence_substitute_dates_json,
       base_allowance_raw, occurrence_rest_days, occurrence_rest_allowances_json,
       daily_statuses_json, evidence_dates_json,
       imported_opening_balance_present, imported_opening_substitute, imported_opening_compensation)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      closureId,
      route,
      month,
      row.store || "",
      String(row.employeeId || "").trim(),
      row.name || "",
      roundHalf(row.baseAllowance),
      roundHalf(row.basicDayoffUsed),
      roundHalf(row.explicitSubDayoffUsed),
      roundHalf(row.baseExcess),
      roundHalf(row.substituteNeeded),
      roundHalf(row.compensationLeaveUsed),
      roundHalf(row.compensationNeeded),
      roundHalf(row.annualLeaveUsed),
      JSON.stringify(Array.isArray(row.substituteEvents) ? row.substituteEvents : []),
      JSON.stringify(Array.isArray(row.compensationEvents) ? row.compensationEvents : []),
      JSON.stringify(Array.isArray(row.annualLeaveEvents) ? row.annualLeaveEvents : []),
      JSON.stringify(Array.isArray(row.workedDates) ? row.workedDates : []),
      JSON.stringify(Array.isArray(row.occurrenceSubstituteDates) ? row.occurrenceSubstituteDates : []),
      roundHalf(row.baseAllowanceRaw ?? row.baseAllowance),
      roundHalf(row.occurrenceRestDays),
      JSON.stringify(Array.isArray(row.occurrenceRestAllowances) ? row.occurrenceRestAllowances : []),
      JSON.stringify(Array.isArray(row.dailyStatuses) ? row.dailyStatuses : []),
      JSON.stringify(Array.isArray(row.evidenceDates) ? row.evidenceDates : []),
      opening.present ? 1 : 0,
      opening.substitute,
      opening.compensation
    );
    });

    await runBatch(context.env.DB, issueStatements);
    await runBatch(context.env.DB, missingStatements);
    await runBatch(context.env.DB, mismatchStatements);
    await runBatch(context.env.DB, factStatements);

    // 월 마감 원본 저장과 누적 재계산은 실패 범위를 분리합니다.
    // 재계산 오류 때문에 정상 저장된 마감 원본까지 삭제되는 현상을 방지합니다.
    let affectedMonths = 0;
    let recalculateWarning = "";
    let summaries = { results: [] };
    try {
      const recalculated = await recalculateRoute(context.env.DB, route);
      affectedMonths = recalculated.affectedMonths || 0;
      summaries = await context.env.DB.prepare(`
        SELECT *
        FROM attendance_monthly_summaries_v3
        WHERE closure_id = ?
        ORDER BY MAX(shortage, compensation_shortage) DESC, base_excess DESC, store ASC, employee_name ASC
      `).bind(closureId).all();
    } catch (error) {
      recalculateWarning = error?.message || "누적 재계산은 다음 저장 또는 부여 수정 시 다시 실행됩니다.";
      console.error("closure recalculation failed", route, month, error);
    }

    return json({
      ok: true,
      id: closureId,
      replaced,
      affectedMonths,
      recalculateWarning,
      summaries: summaries.results || [],
    }, 201);
  } catch (error) {
    await cleanupFailedClosure(context.env.DB, closureId, replaced);
    return json({ error: `월 마감 저장 중 오류가 발생했습니다: ${error.message || "알 수 없는 오류"}` }, 500);
  }
}

async function cleanupFailedClosure(db, closureId, replaced = false) {
  try {
    await db.batch([
      db.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM route_substitute_allocations WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_monthly_employee_facts WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_mismatch_items WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_issue_items WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(closureId),
      // 신규 저장 실패 때만 껍데기 행을 제거합니다. 교체 저장 중 오류라면 월 식별 행은 남겨
      // 사용자가 같은 월 수정본을 다시 올릴 수 있게 합니다.
      ...(replaced ? [] : [db.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(closureId)]),
    ]);
  } catch {
    // 원래 오류를 우선 반환합니다.
  }
}

function uniquePeople(rows) {
  return new Set(rows.map((row) => String(row.employeeId || "").trim()).filter(Boolean)).size;
}

function toNumber(value) {
  return Number(value) || 0;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}
