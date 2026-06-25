import { json, requireAuth } from "../_lib/auth.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const result = await context.env.DB.prepare(`
    SELECT id, company, month, cutoff_date, check_mode, plan_file_name, attendance_file_name,
           plan_people, attendance_people, matched_people, match_rate, missing_count, missing_people, created_at
    FROM attendance_closures ORDER BY month DESC, created_at DESC LIMIT 120
  `).all();
  return json({ items: result.results || [] });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const body = await context.request.json().catch(() => null);
  if (!body || !["homeplus", "electroland"].includes(body.company) || !/^\d{4}-\d{2}$/.test(body.month || "")) return json({ error: "저장 데이터 형식이 올바르지 않습니다." }, 400);
  const id = crypto.randomUUID();
  const rows = Array.isArray(body.rows) ? body.rows.slice(0, 5000) : [];

  const closureStatement = context.env.DB.prepare(`
    INSERT INTO attendance_closures
    (id, company, month, cutoff_date, check_mode, plan_file_name, attendance_file_name,
     plan_people, attendance_people, matched_people, match_rate, missing_count, missing_people)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).bind(
    id, body.company, body.month, body.cutoffDate, body.checkMode || "strict",
    body.planFileName || "", body.attendanceFileName || "", Number(body.planPeople || 0),
    Number(body.attendancePeople || 0), Number(body.matchedPeople || 0), Number(body.matchRate || 0),
    Number(body.missingCount || rows.length), Number(body.missingPeople || 0)
  );

  const statements = [closureStatement];
  if (rows.length) {
    statements.push(context.env.DB.prepare(`
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
    `).bind(id, body.company, JSON.stringify(rows)));
  }

  await context.env.DB.batch(statements);
  return json({ ok: true, id }, 201);
}
