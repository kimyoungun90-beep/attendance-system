import { json, requireAuth } from "../_lib/auth.js";
import { normalizeEmployeeId } from "../_lib/leave-eligibility.js";
import { recalculateRoute, runBatch } from "../_lib/recalculate.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const SYNTHETIC_WORKFORCE_FILE = "이전 최종본 상담사근태 자동매칭";

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const body = await context.request.json().catch(() => null);
  const route = String(body?.route || "");
  const month = String(body?.month || "");
  const fileName = String(body?.fileName || "이전 최종본").slice(0, 240);
  if (!body || !VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month)) {
    return json({ error: "경로와 대상 월을 확인해 주세요." }, 400);
  }

  const employeeFacts = normalizeFacts(body.employeeFacts, route, month);
  if (!employeeFacts.length) {
    return json({ error: "상담사근태 시트에서 저장할 사번·날짜 자료를 찾지 못했습니다." }, 400);
  }

  const existing = await context.env.DB.prepare(`
    SELECT id, check_mode FROM attendance_closures
    WHERE company = ? AND month = ? LIMIT 1
  `).bind(route, month).first();
  const closureId = existing?.id || crypto.randomUUID();
  const replaced = Boolean(existing?.id);
  const cutoffDate = endOfMonth(month);

  try {
    if (!existing) {
      await context.env.DB.prepare(`
        INSERT INTO attendance_closures
        (id, company, month, cutoff_date, check_mode, plan_file_name, attendance_file_name,
         plan_people, attendance_people, matched_people, match_rate,
         missing_count, missing_people, unexpected_count, unexpected_people,
         mismatch_count, mismatch_people, dayoff_excess_people,
         substitute_shortage_people, compensation_shortage_people, annual_leave_people)
        VALUES (?, ?, ?, ?, 'final_import', '이전 최종본 상담사근태', ?, ?, ?, ?, 100,
                0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      `).bind(
        closureId, route, month, cutoffDate, fileName,
        employeeFacts.length, employeeFacts.length, employeeFacts.length
      ).run();
    } else {
      await context.env.DB.prepare(`
        UPDATE attendance_closures
        SET cutoff_date = ?,
            check_mode = CASE WHEN check_mode = 'auto' THEN 'auto' ELSE 'final_import' END,
            attendance_file_name = ?,
            plan_people = CASE WHEN check_mode = 'auto' THEN plan_people ELSE ? END,
            attendance_people = CASE WHEN check_mode = 'auto' THEN attendance_people ELSE ? END,
            matched_people = CASE WHEN check_mode = 'auto' THEN matched_people ELSE ? END,
            match_rate = CASE WHEN check_mode = 'auto' THEN match_rate ELSE 100 END,
            created_at = datetime('now')
        WHERE id = ?
      `).bind(cutoffDate, fileName, employeeFacts.length, employeeFacts.length, employeeFacts.length, closureId).run();
    }

    // 전용 가져오기 월은 상담사근태 시트 내용으로 완전히 교체합니다.
    // 정상 월 마감이 이미 있으면 오류·연차·원본은 유지하고, 이번 파일에 포함된 사번의
    // 대체휴무/보상휴가 사용내역과 일별 최종표시만 갱신합니다.
    if (existing?.check_mode === "final_import") {
      await context.env.DB.prepare("DELETE FROM attendance_monthly_employee_facts WHERE closure_id = ?").bind(closureId).run();
    } else if (existing) {
      const resetStatements = employeeFacts.map((row) => context.env.DB.prepare(`
        UPDATE attendance_monthly_employee_facts
        SET basic_dayoff_used = 0,
            explicit_sub_dayoff_used = 0,
            base_excess = 0,
            substitute_needed = 0,
            compensation_leave_used = 0,
            compensation_needed = 0,
            substitute_events_json = '[]',
            compensation_events_json = '[]',
            worked_dates_json = '[]',
            daily_statuses_json = '[]',
            occurrence_rest_days = 0,
            occurrence_rest_allowances_json = '[]',
            imported_opening_balance_present = 0,
            imported_opening_substitute = 0,
            imported_opening_compensation = 0
        WHERE closure_id = ? AND employee_id = ?
      `).bind(closureId, row.employeeId));
      await runBatch(context.env.DB, resetStatements);
    }

    const factStatements = employeeFacts.map((row) => context.env.DB.prepare(`
      INSERT INTO attendance_monthly_employee_facts
      (closure_id, route, month, store, employee_id, employee_name,
       base_allowance, basic_dayoff_used, explicit_sub_dayoff_used, base_excess,
       substitute_needed, compensation_leave_used, compensation_needed,
       annual_leave_used, substitute_events_json, compensation_events_json,
       annual_leave_events_json, worked_dates_json, occurrence_substitute_dates_json,
       base_allowance_raw, occurrence_rest_days, occurrence_rest_allowances_json,
       daily_statuses_json, evidence_dates_json,
       imported_opening_balance_present, imported_opening_substitute, imported_opening_compensation)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, '[]', ?, '[]', ?, 0, '[]', ?, '[]', ?, ?, ?)
      ON CONFLICT(closure_id, employee_id) DO UPDATE SET
        route = excluded.route,
        month = excluded.month,
        store = excluded.store,
        employee_name = excluded.employee_name,
        base_allowance = excluded.base_allowance,
        basic_dayoff_used = excluded.basic_dayoff_used,
        explicit_sub_dayoff_used = excluded.explicit_sub_dayoff_used,
        base_excess = excluded.base_excess,
        substitute_needed = excluded.substitute_needed,
        compensation_leave_used = excluded.compensation_leave_used,
        compensation_needed = excluded.compensation_needed,
        substitute_events_json = excluded.substitute_events_json,
        compensation_events_json = excluded.compensation_events_json,
        worked_dates_json = excluded.worked_dates_json,
        base_allowance_raw = excluded.base_allowance_raw,
        occurrence_rest_days = 0,
        occurrence_rest_allowances_json = '[]',
        daily_statuses_json = excluded.daily_statuses_json,
        imported_opening_balance_present = excluded.imported_opening_balance_present,
        imported_opening_substitute = excluded.imported_opening_substitute,
        imported_opening_compensation = excluded.imported_opening_compensation
    `).bind(
      closureId, route, month, row.store, row.employeeId, row.name,
      row.baseAllowance, row.basicDayoffUsed, row.explicitSubDayoffUsed, row.baseExcess,
      row.substituteNeeded, row.compensationNeeded, row.compensationNeeded,
      JSON.stringify(row.substituteEvents), JSON.stringify(row.compensationEvents),
      JSON.stringify(row.workedDates), row.baseAllowanceRaw,
      JSON.stringify(row.dailyStatuses),
      0,
      0,
      0
    ));
    await runBatch(context.env.DB, factStatements);

    // 발생일 부여 대상(입사일)과 지역/매장 표시가 과거 최종본만으로도 복원되도록
    // 상담사근태 시트의 인적사항을 월별 인력 스냅샷에 비파괴 방식으로 보완합니다.
    const workforce = await upsertImportedWorkforce(context.env.DB, employeeFacts, route, month);
    const recalculated = await recalculateRoute(context.env.DB, route);

    return json({
      ok: true,
      replaced,
      closureId,
      employeeCount: employeeFacts.length,
      workforceCount: workforce.memberCount,
      substituteEvents: employeeFacts.reduce((sum, row) => sum + row.substituteEvents.length, 0),
      substituteDays: roundHalf(employeeFacts.reduce((sum, row) => sum + row.substituteNeeded, 0)),
      compensationEvents: employeeFacts.reduce((sum, row) => sum + row.compensationEvents.length, 0),
      compensationDays: roundHalf(employeeFacts.reduce((sum, row) => sum + row.compensationNeeded, 0)),
      affectedMonths: recalculated.affectedMonths || 0,
    }, replaced ? 200 : 201);
  } catch (error) {
    if (!existing) {
      await context.env.DB.prepare("DELETE FROM attendance_monthly_employee_facts WHERE closure_id = ?").bind(closureId).run().catch(() => {});
      await context.env.DB.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(closureId).run().catch(() => {});
    }
    return json({ error: `이전 최종본 휴가 사용내역 저장 중 오류가 발생했습니다: ${error.message || "알 수 없는 오류"}` }, 500);
  }
}

async function upsertImportedWorkforce(db, employeeFacts, route, month) {
  let upload = await db.prepare(`
    SELECT id, file_name FROM attendance_workforce_uploads WHERE month = ? LIMIT 1
  `).bind(month).first();

  if (!upload?.id) {
    upload = { id: crypto.randomUUID(), file_name: SYNTHETIC_WORKFORCE_FILE };
    await db.prepare(`
      INSERT INTO attendance_workforce_uploads
      (id, month, file_name, content_type, size_bytes, storage_type, object_key,
       employee_count, electroland_count, homeplus_count, portal_count, created_at, updated_at)
      VALUES (?, ?, ?, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              0, 'd1', NULL, 0, 0, 0, 0, datetime('now'), datetime('now'))
    `).bind(upload.id, month, SYNTHETIC_WORKFORCE_FILE).run();
  } else if (String(upload.file_name || "").startsWith(SYNTHETIC_WORKFORCE_FILE)) {
    // 자동 생성 스냅샷만 같은 경로·같은 월 재등록 시 교체합니다.
    // 사용자가 별도로 올린 인력자료는 절대 삭제하지 않습니다.
    await db.prepare(`
      DELETE FROM attendance_workforce_members WHERE upload_id = ? AND route = ?
    `).bind(upload.id, route).run();
  }

  const usable = employeeFacts.filter((row) => row.employeeId && row.name);
  const portalStatements = usable
    .filter((row) => row.portalId)
    .map((row) => db.prepare(`
      INSERT INTO attendance_portal_ids (employee_id, portal_id, employee_name, updated_at)
      VALUES (?, ?, ?, datetime('now'))
      ON CONFLICT(employee_id) DO UPDATE SET
        portal_id = excluded.portal_id,
        employee_name = COALESCE(NULLIF(excluded.employee_name, ''), attendance_portal_ids.employee_name),
        updated_at = datetime('now')
    `).bind(row.employeeId, row.portalId, row.name));
  await runBatch(db, portalStatements);

  const memberStatements = usable.map((row) => db.prepare(`
    INSERT INTO attendance_workforce_members
    (upload_id, month, route, regional_manager, manager, region1, region2,
     store_code, store_name, portal_id, employee_id, employee_name,
     hire_date, group_hire_date, note)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(month, route, employee_id, store_code) DO UPDATE SET
      upload_id = excluded.upload_id,
      regional_manager = COALESCE(NULLIF(excluded.regional_manager, ''), attendance_workforce_members.regional_manager),
      manager = COALESCE(NULLIF(excluded.manager, ''), attendance_workforce_members.manager),
      region1 = COALESCE(NULLIF(excluded.region1, ''), attendance_workforce_members.region1),
      region2 = COALESCE(NULLIF(excluded.region2, ''), attendance_workforce_members.region2),
      store_name = COALESCE(NULLIF(excluded.store_name, ''), attendance_workforce_members.store_name),
      portal_id = COALESCE(NULLIF(excluded.portal_id, ''), attendance_workforce_members.portal_id),
      employee_name = COALESCE(NULLIF(excluded.employee_name, ''), attendance_workforce_members.employee_name),
      hire_date = COALESCE(NULLIF(excluded.hire_date, ''), attendance_workforce_members.hire_date),
      group_hire_date = COALESCE(NULLIF(excluded.group_hire_date, ''), attendance_workforce_members.group_hire_date),
      note = COALESCE(NULLIF(excluded.note, ''), attendance_workforce_members.note)
  `).bind(
    upload.id, month, route, row.regionalManager, row.manager, row.region1, row.region2,
    row.storeCode, row.store, row.portalId, row.employeeId, row.name,
    row.hireDate || null, row.groupHireDate || null, row.note
  ));
  await runBatch(db, memberStatements);

  const counts = await db.prepare(`
    SELECT
      COUNT(*) AS employee_count,
      SUM(CASE WHEN route = 'electroland' THEN 1 ELSE 0 END) AS electroland_count,
      SUM(CASE WHEN route = 'homeplus' THEN 1 ELSE 0 END) AS homeplus_count,
      SUM(CASE WHEN COALESCE(portal_id, '') <> '' THEN 1 ELSE 0 END) AS portal_count
    FROM attendance_workforce_members
    WHERE upload_id = ?
  `).bind(upload.id).first();
  await db.prepare(`
    UPDATE attendance_workforce_uploads
    SET employee_count = ?, electroland_count = ?, homeplus_count = ?, portal_count = ?, updated_at = datetime('now')
    WHERE id = ?
  `).bind(
    Number(counts?.employee_count || 0),
    Number(counts?.electroland_count || 0),
    Number(counts?.homeplus_count || 0),
    Number(counts?.portal_count || 0),
    upload.id
  ).run();

  return { uploadId: upload.id, memberCount: usable.length };
}

function normalizeFacts(items, route, month) {
  const baseAllowance = route === "homeplus" ? 6 : countWeekendDays(month);
  const map = new Map();
  for (const source of Array.isArray(items) ? items.slice(0, 4000) : []) {
    const employeeId = normalizeEmployeeId(source?.employeeId);
    if (!employeeId) continue;
    const substituteEvents = normalizeEvents(source?.substituteEvents, month, "substitute");
    const compensationEvents = normalizeEvents(source?.compensationEvents, month, "compensation");
    const dailyStatuses = normalizeDailyStatuses(source?.dailyStatuses, month);
    const workedDates = [...new Set((Array.isArray(source?.workedDates) ? source.workedDates : [])
      .map((date) => String(date || "").slice(0, 10))
      .filter((date) => date.startsWith(`${month}-`)))].sort();
    const basicDayoffUsed = roundHalf(source?.basicDayoffUsed);
    const substituteNeeded = roundHalf(substituteEvents.reduce((sum, event) => sum + event.days, 0));
    const compensationNeeded = roundHalf(compensationEvents.reduce((sum, event) => sum + event.days, 0));
    map.set(employeeId, {
      employeeId,
      name: clean(source?.name).slice(0, 120),
      store: clean(source?.store).slice(0, 120),
      regionalManager: clean(source?.regionalManager).slice(0, 120),
      manager: clean(source?.manager).slice(0, 120),
      region1: clean(source?.region1).slice(0, 120),
      region2: clean(source?.region2).slice(0, 120),
      storeCode: clean(source?.storeCode).replace(/,/g, "").replace(/\.0+$/, "").slice(0, 60),
      portalId: clean(source?.portalId).replace(/,/g, "").replace(/\.0+$/, "").slice(0, 80),
      hireDate: normalizeDate(source?.hireDate),
      groupHireDate: normalizeDate(source?.groupHireDate),
      note: clean(source?.note).slice(0, 500),
      baseAllowanceRaw: baseAllowance,
      baseAllowance,
      basicDayoffUsed,
      explicitSubDayoffUsed: substituteNeeded,
      baseExcess: roundHalf(Math.max(0, basicDayoffUsed - baseAllowance)),
      substituteNeeded,
      compensationNeeded,
      substituteEvents,
      compensationEvents,
      workedDates,
      dailyStatuses,
      // 과거 최종본의 수기 이월칸은 참고값일 뿐입니다.
      // 실제 잔여는 부여 설정의 발생일·사용기간과 월별 사용내역으로 다시 계산합니다.
      importedOpeningBalancePresent: false,
      importedOpeningSubstitute: 0,
      importedOpeningCompensation: 0,
    });
  }
  return [...map.values()];
}

function normalizeEvents(items, month, type) {
  const map = new Map();
  for (const source of Array.isArray(items) ? items : []) {
    const date = String(source?.date || "").slice(0, 10);
    const days = roundHalf(source?.days);
    if (!date.startsWith(`${month}-`) || !(days > 0)) continue;
    const planStatus = String(source?.planStatus || (type === "compensation" ? "보상휴가(1일)" : "대체휴일(1일)")).trim();
    const key = `${date}|${planStatus}`;
    if (!map.has(key)) map.set(key, {
      date,
      days,
      source: "이전 최종본 상담사근태",
      planStatus,
    });
  }
  return [...map.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function normalizeDailyStatuses(items, month) {
  const map = new Map();
  for (const source of Array.isArray(items) ? items : []) {
    const date = String(source?.date || "").slice(0, 10);
    if (!date.startsWith(`${month}-`)) continue;
    map.set(date, {
      date,
      planStatus: String(source?.planStatus || "공백").trim(),
      hasClockIn: Boolean(source?.hasClockIn),
      actualStatus: String(source?.actualStatus || "").trim(),
      importedFinal: true,
    });
  }
  return [...map.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function normalizeDate(value) {
  const raw = String(value || "").trim().slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : "";
}

function clean(value) {
  return String(value ?? "").trim();
}

function countWeekendDays(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  const days = new Date(year, month, 0).getDate();
  let count = 0;
  for (let day = 1; day <= days; day += 1) {
    const weekday = new Date(year, month - 1, day).getDay();
    if (weekday === 0 || weekday === 6) count += 1;
  }
  return count;
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
