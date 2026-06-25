import { json, requireAuth } from "../_lib/auth.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const url = new URL(context.request.url);
  const company = url.searchParams.get("company");
  const params = [];
  let where = "";
  if (company) {
    if (!["homeplus", "electroland"].includes(company)) return json({ error: "회사 구분이 올바르지 않습니다." }, 400);
    where = "WHERE g.company = ?";
    params.push(company);
  }
  const statement = context.env.DB.prepare(`
    SELECT g.id, g.company, g.employee_id, g.employee_name, g.store, g.grant_month,
           g.granted_days, g.valid_from, g.valid_to, g.reason, g.note, g.created_at,
           COALESCE(SUM(a.used_days), 0) AS used_days,
           MAX(0, g.granted_days - COALESCE(SUM(a.used_days), 0)) AS remaining_days
    FROM substitute_dayoff_grants g
    LEFT JOIN substitute_dayoff_allocations a ON a.grant_id = g.id
    ${where}
    GROUP BY g.id
    ORDER BY g.valid_to ASC, g.employee_name ASC, g.created_at ASC
  `);
  const result = params.length ? await statement.bind(...params).all() : await statement.all();
  return json({ items: result.results || [] });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const body = await context.request.json().catch(() => null);
  if (!body || !["homeplus", "electroland"].includes(body.company)) return json({ error: "회사 구분이 올바르지 않습니다." }, 400);
  const employeeId = String(body.employeeId || "").trim().toUpperCase();
  const employeeName = String(body.employeeName || "").trim();
  const grantMonth = String(body.grantMonth || "");
  const grantedDays = Math.round(Number(body.grantedDays || 0) * 2) / 2;
  const validFrom = String(body.validFrom || "");
  const validTo = String(body.validTo || "");
  if (!employeeId || !employeeName || !/^\d{4}-\d{2}$/.test(grantMonth) || !(grantedDays > 0) || !/^\d{4}-\d{2}-\d{2}$/.test(validFrom) || !/^\d{4}-\d{2}-\d{2}$/.test(validTo) || validFrom > validTo) {
    return json({ error: "사번·이름·부여 일수·사용기간을 확인해 주세요." }, 400);
  }
  const id = crypto.randomUUID();
  await context.env.DB.prepare(`
    INSERT INTO substitute_dayoff_grants
    (id, company, employee_id, employee_name, store, grant_month, granted_days, valid_from, valid_to, reason, note)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).bind(
    id, body.company, employeeId, employeeName, String(body.store || "").trim(), grantMonth,
    grantedDays, validFrom, validTo, String(body.reason || "").trim(), String(body.note || "").trim()
  ).run();
  return json({ ok: true, id }, 201);
}
