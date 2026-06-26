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

    const factStatements = employeeFacts.map((row) => context.env.DB.prepare(`
      INSERT INTO attendance_monthly_employee_facts
      (closure_id, route, month, store, employee_id, employee_name,
       base_allowance, basic_dayoff_used, explicit_sub_dayoff_used, base_excess,
       substitute_needed, compensation_leave_used, compensation_needed,
       annual_leave_used, substitute_events_json, compensation_events_json,
       annual_leave_events_json, worked_dates_json)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
      JSON.stringify(Array.isArray(row.workedDates) ? row.workedDates : [])
    ));

    await runBatch(context.env.DB, issueStatements);
    await runBatch(context.env.DB, missingStatements);
    await runBatch(context.env.DB, mismatchStatements);
    await runBatch(context.env.DB, factStatements);

    const recalculated = await recalculateRoute(context.env.DB, route);
    const summaries = await context.env.DB.prepare(`
      SELECT *
      FROM attendance_monthly_summaries_v3
      WHERE closure_id = ?
      ORDER BY MAX(shortage, compensation_shortage) DESC, base_excess DESC, store ASC, employee_name ASC
    `).bind(closureId).all();

    return json({
      ok: true,
      id: closureId,
      replaced,
      affectedMonths: recalculated.affectedMonths,
      summaries: summaries.results || [],
    }, 201);
  } catch (error) {
    await cleanupFailedClosure(context.env.DB, closureId);
    return json({ error: `월 마감 저장 중 오류가 발생했습니다: ${error.message || "알 수 없는 오류"}` }, 500);
  }
}

async function cleanupFailedClosure(db, closureId) {
  try {
    await db.batch([
      db.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM route_substitute_allocations WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_monthly_employee_facts WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_mismatch_items WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_issue_items WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(closureId),
      db.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(closureId),
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
