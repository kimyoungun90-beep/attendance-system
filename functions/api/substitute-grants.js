import { json, requireAuth } from "../_lib/auth.js";
import { recalculateRoute } from "../_lib/recalculate.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const route = url.searchParams.get("route") || "";
  if (route && !VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);

  const where = route ? "WHERE g.route = ?" : "";
  const statement = context.env.DB.prepare(`
    SELECT g.id, g.route, g.grant_month, g.granted_days, g.valid_from, g.valid_to,
           g.reason, g.note, g.created_at, g.updated_at,
           COALESCE(e.eligible_people, 0) AS eligible_people,
           ROUND(g.granted_days * COALESCE(e.eligible_people, 0), 1) AS assigned_days,
           COALESCE(u.used_days, 0) AS used_days,
           MAX(0, ROUND(g.granted_days * COALESCE(e.eligible_people, 0), 1) - COALESCE(u.used_days, 0)) AS unused_days
    FROM route_substitute_grants g
    LEFT JOIN (
      SELECT rg.id AS grant_id, COUNT(f.employee_id) AS eligible_people
      FROM route_substitute_grants rg
      LEFT JOIN attendance_closures c
        ON c.company = rg.route AND c.month = rg.grant_month
      LEFT JOIN attendance_monthly_employee_facts f
        ON f.closure_id = c.id
      GROUP BY rg.id
    ) e ON e.grant_id = g.id
    LEFT JOIN (
      SELECT grant_id, ROUND(SUM(used_days), 1) AS used_days
      FROM route_substitute_allocations
      GROUP BY grant_id
    ) u ON u.grant_id = g.id
    ${where}
    ORDER BY g.grant_month DESC, g.route ASC
  `);

  const result = route ? await statement.bind(route).all() : await statement.all();
  return json({ items: result.results || [] });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const body = await context.request.json().catch(() => null);
  const route = String(body?.route || "");
  const grantMonth = String(body?.grantMonth || "");
  const grantedDays = roundHalf(body?.grantedDays);
  const validFrom = String(body?.validFrom || "");
  const validTo = String(body?.validTo || "");

  if (!body || !VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);
  if (!/^\d{4}-\d{2}$/.test(grantMonth) || !(grantedDays > 0)) {
    return json({ error: "발생 월과 전 직원 공통 부여 일수를 확인해 주세요." }, 400);
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(validFrom) || !/^\d{4}-\d{2}-\d{2}$/.test(validTo) || validFrom > validTo) {
    return json({ error: "사용 시작일과 종료일을 확인해 주세요." }, 400);
  }

  const existing = await context.env.DB.prepare(
    "SELECT id FROM route_substitute_grants WHERE route = ? AND grant_month = ?"
  ).bind(route, grantMonth).first();
  const id = existing?.id || crypto.randomUUID();

  if (existing?.id) {
    await context.env.DB.prepare(`
      UPDATE route_substitute_grants
      SET granted_days = ?, valid_from = ?, valid_to = ?, reason = ?, note = ?, updated_at = datetime('now')
      WHERE id = ?
    `).bind(
      grantedDays,
      validFrom,
      validTo,
      String(body.reason || "").trim(),
      String(body.note || "").trim(),
      id
    ).run();
  } else {
    await context.env.DB.prepare(`
      INSERT INTO route_substitute_grants
      (id, route, grant_month, granted_days, valid_from, valid_to, reason, note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      id,
      route,
      grantMonth,
      grantedDays,
      validFrom,
      validTo,
      String(body.reason || "").trim(),
      String(body.note || "").trim()
    ).run();
  }

  const recalculated = await recalculateRoute(context.env.DB, route);
  return json({ ok: true, id, replaced: Boolean(existing?.id), affectedMonths: recalculated.affectedMonths }, 201);
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
