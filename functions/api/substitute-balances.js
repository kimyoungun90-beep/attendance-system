import { json, requireAuth } from "../_lib/auth.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const route = url.searchParams.get("route") || "";
  const month = url.searchParams.get("month") || "";
  if (!VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month)) {
    return json({ error: "경로와 대상 월을 확인해 주세요." }, 400);
  }

  const monthStart = `${month}-01`;
  const monthEnd = endOfMonth(month);

  // 대상 월 이전에 발생한 공통 부여분만 기존 마감 대상자를 기준으로 불러옵니다.
  // 대상 월에 발생한 부여분은 현재 업로드한 마감 인원에게 적용해야 하므로 currentGrant로 별도 반환합니다.
  const lotsResult = await context.env.DB.prepare(`
    SELECT g.id AS grant_id, f.employee_id, g.grant_month, g.granted_days,
           g.valid_from, g.valid_to,
           ROUND(g.granted_days - COALESCE(SUM(a.used_days), 0), 1) AS remaining_days
    FROM route_substitute_grants g
    JOIN attendance_closures c
      ON c.company = g.route AND c.month = g.grant_month
    JOIN attendance_monthly_employee_facts f
      ON f.closure_id = c.id
    LEFT JOIN route_substitute_allocations a
      ON a.grant_id = g.id
     AND a.employee_id = f.employee_id
     AND a.month < ?
    WHERE g.route = ?
      AND g.grant_month < ?
      AND g.valid_from <= ?
      AND g.valid_to >= ?
    GROUP BY g.id, f.employee_id
    HAVING remaining_days > 0
    ORDER BY f.employee_id, g.valid_to, g.valid_from, g.grant_month
  `).bind(month, route, month, monthEnd, monthStart).all();

  const annualResult = await context.env.DB.prepare(`
    SELECT f.employee_id, ROUND(SUM(f.annual_leave_used), 1) AS cumulative_days
    FROM attendance_monthly_employee_facts f
    WHERE f.route = ? AND f.month < ?
    GROUP BY f.employee_id
  `).bind(route, month).all();

  const currentGrant = await context.env.DB.prepare(`
    SELECT id, route, grant_month, granted_days, valid_from, valid_to, reason, note
    FROM route_substitute_grants
    WHERE route = ? AND grant_month = ?
    LIMIT 1
  `).bind(route, month).first();

  const lotsByEmployee = {};
  for (const row of lotsResult.results || []) {
    const employeeId = row.employee_id;
    if (!lotsByEmployee[employeeId]) lotsByEmployee[employeeId] = [];
    lotsByEmployee[employeeId].push({
      grantId: row.grant_id,
      grantMonth: row.grant_month,
      grantedDays: roundHalf(row.granted_days),
      remaining: roundHalf(row.remaining_days),
      validFrom: row.valid_from,
      validTo: row.valid_to,
    });
  }

  const annualLeaveBefore = {};
  for (const row of annualResult.results || []) {
    annualLeaveBefore[row.employee_id] = roundHalf(row.cumulative_days);
  }

  const balances = {};
  for (const [employeeId, lots] of Object.entries(lotsByEmployee)) {
    balances[employeeId] = roundHalf(lots.reduce((sum, lot) => sum + Number(lot.remaining || 0), 0));
  }

  return json({ lotsByEmployee, balances, annualLeaveBefore, currentGrant: currentGrant || null });
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
