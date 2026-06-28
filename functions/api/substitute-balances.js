import { json, requireAuth } from "../_lib/auth.js";
import { normalizeEmployeeId, resolveEligibleEmployees } from "../_lib/leave-eligibility.js";
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
  const [grantsResult, factsResult, allocationsResult, annualUsageResult, workforceResult, annualEmployeeResult] = await Promise.all([
    context.env.DB.prepare(`
      SELECT * FROM attendance_leave_grants_v5
      WHERE route = ? AND valid_from <= ? AND valid_to >= ?
      ORDER BY grant_month, valid_to, valid_from, created_at
    `).bind(route, monthEnd, monthStart).all(),
    context.env.DB.prepare(`
      SELECT * FROM attendance_monthly_employee_facts
      WHERE route = ?
    `).bind(route).all(),
    context.env.DB.prepare(`
      SELECT grant_id, employee_id, ROUND(SUM(used_days), 1) AS used_days
      FROM attendance_leave_allocations_v5
      WHERE route = ? AND month < ?
      GROUP BY grant_id, employee_id
    `).bind(route, month).all(),
    context.env.DB.prepare(`
      SELECT employee_id, ROUND(SUM(annual_leave_used), 1) AS cumulative_days
      FROM attendance_monthly_employee_facts
      WHERE route = ? AND month < ?
      GROUP BY employee_id
    `).bind(route, month).all(),
    context.env.DB.prepare(`
      SELECT * FROM attendance_workforce_members
      WHERE route = ?
      ORDER BY month, id
    `).bind(route).all(),
    context.env.DB.prepare(`
      SELECT * FROM annual_leave_employees
      WHERE route = ?
    `).bind(route).all(),
  ]);

  const factsByMonth = new Map();
  for (const fact of factsResult.results || []) {
    if (!factsByMonth.has(fact.month)) factsByMonth.set(fact.month, []);
    factsByMonth.get(fact.month).push(fact);
  }
  const usedMap = new Map((allocationsResult.results || []).map((row) => [
    `${row.grant_id}|${normalizeEmployeeId(row.employee_id)}`,
    roundHalf(row.used_days),
  ]));

  // 대상 월에 사용 가능한 모든 부여분을 직원별 lot로 바로 내려보냅니다.
  // grant_month가 7월이어도 사용 시작일이 5월이면 6월 분석에서 사용할 수 있습니다.
  const lotsByEmployee = {};
  for (const grant of grantsResult.results || []) {
    const eligibility = resolveEligibleEmployees(grant, {
      workforceRows: workforceResult.results || [],
      annualRows: annualEmployeeResult.results || [],
      factsByMonth,
    });
    for (const employeeId of eligibility.employeeIds) {
      const remaining = roundHalf(Number(grant.granted_days || 0) - (usedMap.get(`${grant.id}|${employeeId}`) || 0));
      if (!(remaining > 0)) continue;
      if (!lotsByEmployee[employeeId]) lotsByEmployee[employeeId] = [];
      lotsByEmployee[employeeId].push({
        grantId: grant.id,
        grantType: grant.grant_type || "substitute",
        grantMonth: grant.grant_month,
        occurrenceDate: grant.occurrence_date || "",
        grantedDays: roundHalf(grant.granted_days),
        remaining,
        validFrom: grant.valid_from,
        validTo: grant.valid_to,
      });
    }
  }

  const annualLeaveBefore = {};
  for (const row of annualUsageResult.results || []) {
    annualLeaveBefore[normalizeEmployeeId(row.employee_id)] = roundHalf(row.cumulative_days);
  }

  const balances = {};
  for (const [employeeId, lots] of Object.entries(lotsByEmployee)) {
    balances[employeeId] = {
      substitute: roundHalf(lots.filter((lot) => lot.grantType === "substitute").reduce((sum, lot) => sum + Number(lot.remaining || 0), 0)),
      compensation: roundHalf(lots.filter((lot) => lot.grantType === "compensation").reduce((sum, lot) => sum + Number(lot.remaining || 0), 0)),
    };
  }

  // currentGrants는 기존 브라우저 이중 합산 방지를 위해 비웁니다.
  return json({ lotsByEmployee, balances, annualLeaveBefore, currentGrants: [] });
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
