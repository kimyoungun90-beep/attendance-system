import { json, requireAuth } from "../_lib/auth.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const result = await context.env.DB.prepare(`
    SELECT id, company, month, cutoff_date, check_mode, plan_file_name, attendance_file_name,
           plan_people, attendance_people, matched_people, match_rate,
           missing_count, missing_people, unexpected_count, unexpected_people,
           dayoff_excess_people, substitute_shortage_people, created_at
    FROM attendance_closures
    ORDER BY month DESC, created_at DESC LIMIT 120
  `).all();
  return json({ items: result.results || [] });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const body = await context.request.json().catch(() => null);
  if (!body || !["homeplus", "electroland"].includes(body.company) || !/^\d{4}-\d{2}$/.test(body.month || "")) {
    return json({ error: "저장 데이터 형식이 올바르지 않습니다." }, 400);
  }

  const issueRows = Array.isArray(body.issueRows) ? body.issueRows.slice(0, 10000) : [];
  const employeeSummaries = Array.isArray(body.employeeSummaries) ? body.employeeSummaries.slice(0, 1000) : [];
  const old = await context.env.DB.prepare("SELECT id FROM attendance_closures WHERE company = ? AND month = ?")
    .bind(body.company, body.month).all();
  const oldIds = (old.results || []).map((item) => item.id);

  for (const oldId of oldIds) {
    await runBatch(context.env.DB, [
      context.env.DB.prepare("DELETE FROM substitute_dayoff_allocations WHERE closure_id = ?").bind(oldId),
      context.env.DB.prepare("DELETE FROM attendance_employee_summaries WHERE closure_id = ?").bind(oldId),
      context.env.DB.prepare("DELETE FROM attendance_issue_items WHERE closure_id = ?").bind(oldId),
      context.env.DB.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(oldId),
      context.env.DB.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(oldId),
    ]);
  }

  const id = crypto.randomUUID();
  await context.env.DB.prepare(`
    INSERT INTO attendance_closures
    (id, company, month, cutoff_date, check_mode, plan_file_name, attendance_file_name,
     plan_people, attendance_people, matched_people, match_rate, missing_count, missing_people,
     unexpected_count, unexpected_people, dayoff_excess_people, substitute_shortage_people)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).bind(
    id, body.company, body.month, body.cutoffDate || `${body.month}-01`, body.checkMode || "strict",
    body.planFileName || "", body.attendanceFileName || "", Number(body.planPeople || 0),
    Number(body.attendancePeople || 0), Number(body.matchedPeople || 0), Number(body.matchRate || 0),
    Number(body.missingCount || 0), Number(body.missingPeople || 0), Number(body.unexpectedCount || 0),
    Number(body.unexpectedPeople || 0), Number(body.dayoffExcessPeople || 0), Number(body.substituteShortagePeople || 0)
  ).run();

  if (issueRows.length) {
    await context.env.DB.prepare(`
      INSERT OR IGNORE INTO attendance_issue_items
      (closure_id, issue_type, company, store, employee_id, employee_name, issue_date, weekday,
       plan_status, actual_in, changed_in, clock_status, result, reason, duplicate_plan_note)
      SELECT ?,
             COALESCE(json_extract(value, '$.issueType'), 'missing_clock_in'), ?,
             COALESCE(json_extract(value, '$.store'), ''),
             COALESCE(json_extract(value, '$.employeeId'), ''),
             COALESCE(json_extract(value, '$.name'), ''),
             COALESCE(json_extract(value, '$.date'), ''),
             COALESCE(json_extract(value, '$.weekday'), ''),
             COALESCE(json_extract(value, '$.planStatus'), ''),
             COALESCE(json_extract(value, '$.actualIn'), ''),
             COALESCE(json_extract(value, '$.changedIn'), ''),
             COALESCE(json_extract(value, '$.clockStatus'), ''),
             COALESCE(json_extract(value, '$.result'), ''),
             COALESCE(json_extract(value, '$.reason'), ''),
             COALESCE(json_extract(value, '$.duplicatePlanNote'), '')
      FROM json_each(?)
    `).bind(id, body.company, JSON.stringify(issueRows)).run();

    const missingRows = issueRows.filter((row) => row.issueType === "missing_clock_in");
    if (missingRows.length) {
      await context.env.DB.prepare(`
        INSERT OR IGNORE INTO attendance_missing_items
        (closure_id, company, store, employee_id, employee_name, missing_date, weekday, plan_status,
         actual_in, changed_in, clock_status, result, reason, duplicate_plan_note)
        SELECT ?, ?,
               COALESCE(json_extract(value, '$.store'), ''),
               COALESCE(json_extract(value, '$.employeeId'), ''),
               COALESCE(json_extract(value, '$.name'), ''),
               COALESCE(json_extract(value, '$.date'), ''),
               COALESCE(json_extract(value, '$.weekday'), ''),
               COALESCE(json_extract(value, '$.planStatus'), ''),
               COALESCE(json_extract(value, '$.actualIn'), ''),
               COALESCE(json_extract(value, '$.changedIn'), ''),
               COALESCE(json_extract(value, '$.clockStatus'), ''),
               COALESCE(json_extract(value, '$.result'), ''),
               COALESCE(json_extract(value, '$.reason'), ''),
               COALESCE(json_extract(value, '$.duplicatePlanNote'), '')
        FROM json_each(?)
      `).bind(id, body.company, JSON.stringify(missingRows)).run();
    }
  }

  const monthStart = `${body.month}-01`;
  const [year, monthNumber] = body.month.split("-").map(Number);
  const monthEnd = `${body.month}-${String(new Date(year, monthNumber, 0).getDate()).padStart(2, "0")}`;
  const summaryStatements = [];
  const allocationStatements = [];
  let actualShortagePeople = 0;

  for (const source of employeeSummaries) {
    const employeeId = String(source.employeeId || "").trim();
    if (!employeeId) continue;
    const substituteNeeded = roundHalf(Number(source.substituteNeeded || 0));
    const grantsResult = await context.env.DB.prepare(`
      SELECT g.id, g.granted_days, g.valid_from, g.valid_to,
             COALESCE(SUM(a.used_days), 0) AS used_days
      FROM substitute_dayoff_grants g
      LEFT JOIN substitute_dayoff_allocations a ON a.grant_id = g.id
      WHERE g.company = ? AND g.employee_id = ? AND g.valid_from <= ? AND g.valid_to >= ?
      GROUP BY g.id
      ORDER BY g.valid_to ASC, g.valid_from ASC, g.created_at ASC
    `).bind(body.company, employeeId, monthEnd, monthStart).all();

    const grants = (grantsResult.results || []).map((grant) => ({
      ...grant,
      remaining: roundHalf(Math.max(0, Number(grant.granted_days || 0) - Number(grant.used_days || 0))),
    }));
    const available = roundHalf(grants.reduce((sum, grant) => sum + grant.remaining, 0));
    let needLeft = substituteNeeded;
    let applied = 0;

    for (const grant of grants) {
      if (needLeft <= 0) break;
      const use = roundHalf(Math.min(needLeft, grant.remaining));
      if (use <= 0) continue;
      applied = roundHalf(applied + use);
      needLeft = roundHalf(needLeft - use);
      allocationStatements.push(context.env.DB.prepare(`
        INSERT INTO substitute_dayoff_allocations
        (grant_id, closure_id, company, employee_id, month, used_days)
        VALUES (?, ?, ?, ?, ?, ?)
      `).bind(grant.id, id, body.company, employeeId, body.month, use));
    }

    const shortage = roundHalf(Math.max(0, substituteNeeded - applied));
    const remaining = roundHalf(Math.max(0, available - applied));
    if (shortage > 0) actualShortagePeople += 1;
    const judgment = buildJudgment({
      baseAllowance: Number(source.baseAllowance || 0),
      basicDayoffUsed: Number(source.basicDayoffUsed || 0),
      baseExcess: Number(source.baseExcess || 0),
      explicitSubDayoffUsed: Number(source.explicitSubDayoffUsed || 0),
      substituteNeeded, available, shortage, remaining,
    });

    summaryStatements.push(context.env.DB.prepare(`
      INSERT INTO attendance_employee_summaries
      (closure_id, company, month, store, employee_id, employee_name, base_allowance,
       basic_dayoff_used, explicit_sub_dayoff_used, base_excess, substitute_needed,
       available_substitute, substitute_applied, remaining_substitute, shortage, judgment)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      id, body.company, body.month, source.store || "", employeeId, source.name || "",
      Number(source.baseAllowance || 0), Number(source.basicDayoffUsed || 0),
      Number(source.explicitSubDayoffUsed || 0), Number(source.baseExcess || 0), substituteNeeded,
      available, applied, remaining, shortage, judgment
    ));
  }

  await runBatch(context.env.DB, allocationStatements);
  await runBatch(context.env.DB, summaryStatements);
  await context.env.DB.prepare("UPDATE attendance_closures SET substitute_shortage_people = ? WHERE id = ?")
    .bind(actualShortagePeople, id).run();

  return json({ ok: true, id, substituteShortagePeople: actualShortagePeople }, 201);
}

async function runBatch(db, statements) {
  for (let index = 0; index < statements.length; index += 80) {
    const chunk = statements.slice(index, index + 80);
    if (chunk.length) await db.batch(chunk);
  }
}

function roundHalf(value) { return Math.round((Number(value) || 0) * 2) / 2; }
function formatDays(value) { const numberValue = roundHalf(value); return `${Number.isInteger(numberValue) ? numberValue : numberValue.toFixed(1)}일`; }
function buildJudgment({ baseAllowance, basicDayoffUsed, baseExcess, explicitSubDayoffUsed, substituteNeeded, available, shortage, remaining }) {
  if (shortage > 0) return `휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 대체휴무 필요 ${formatDays(substituteNeeded)}, 가용 ${formatDays(available)} → ${formatDays(shortage)} 초과 사용`;
  if (baseExcess > 0) return `휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 초과 ${formatDays(baseExcess)}은 대체휴무 여분 활용 · 잔여 ${formatDays(remaining)}`;
  if (explicitSubDayoffUsed > 0) return `기본 휴무 정상 · 대체휴무 ${formatDays(explicitSubDayoffUsed)} 사용 · 잔여 ${formatDays(remaining)}`;
  return `정상 · 기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)}`;
}
