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
  const [grantsResult, factsResult, allocationsResult, annualResult] = await Promise.all([
    context.env.DB.prepare(`
      SELECT * FROM attendance_leave_grants_v5
      WHERE route = ? AND grant_month <= ? AND valid_from <= ? AND valid_to >= ?
      ORDER BY grant_month, valid_to, valid_from, created_at
    `).bind(route, month, monthEnd, monthStart).all(),
    context.env.DB.prepare(`
      SELECT * FROM attendance_monthly_employee_facts
      WHERE route = ? AND month <= ?
    `).bind(route, month).all(),
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

  const lotsByEmployee = {};
  const currentGrants = [];
  for (const grant of grantsResult.results || []) {
    if (grant.grant_month === month) {
      currentGrants.push(normalizeGrant(grant));
      continue;
    }
    const grantFacts = factsByMonth.get(grant.grant_month) || [];
    for (const fact of grantFacts) {
      if (!isEligibleForGrant(grant, fact)) continue;
      const employeeId = normalizeEmployeeId(fact.employee_id);
      const remaining = roundHalf(Number(grant.granted_days || 0) - (usedMap.get(`${grant.id}|${employeeId}`) || 0));
      if (!(remaining > 0)) continue;
      if (!lotsByEmployee[employeeId]) lotsByEmployee[employeeId] = [];
      lotsByEmployee[employeeId].push({
        grantId: grant.id,
        grantType: grant.grant_type || "substitute",
        grantMonth: grant.grant_month,
        grantedDays: roundHalf(grant.granted_days),
        remaining,
        validFrom: grant.valid_from,
        validTo: grant.valid_to,
      });
    }
  }

  const annualLeaveBefore = {};
  for (const row of annualResult.results || []) {
    annualLeaveBefore[normalizeEmployeeId(row.employee_id)] = roundHalf(row.cumulative_days);
  }

  const balances = {};
  for (const [employeeId, lots] of Object.entries(lotsByEmployee)) {
    balances[employeeId] = {
      substitute: roundHalf(lots.filter((lot) => lot.grantType === "substitute").reduce((sum, lot) => sum + Number(lot.remaining || 0), 0)),
      compensation: roundHalf(lots.filter((lot) => lot.grantType === "compensation").reduce((sum, lot) => sum + Number(lot.remaining || 0), 0)),
    };
  }

  return json({ lotsByEmployee, balances, annualLeaveBefore, currentGrants });
}

function normalizeGrant(grant) {
  return {
    id: grant.id,
    route: grant.route,
    grantType: grant.grant_type || "substitute",
    grantScope: grant.grant_scope || "route",
    grantMonth: grant.grant_month,
    grantedDays: roundHalf(grant.granted_days),
    validFrom: grant.valid_from,
    validTo: grant.valid_to,
    eligibilityMode: grant.eligibility_mode || "all",
    criterionDate: grant.criterion_date || "",
    employeeId: normalizeEmployeeId(grant.employee_id),
    excludedEmployeeIds: parseEmployeeIds(grant.excluded_employee_ids_json),
  };
}

function isEligibleForGrant(grant, fact) {
  const employeeId = normalizeEmployeeId(fact.employee_id);
  if (!employeeId) return false;
  if (grant.grant_scope === "employee") {
    if (employeeId !== normalizeEmployeeId(grant.employee_id)) return false;
  } else if (new Set(parseEmployeeIds(grant.excluded_employee_ids_json)).has(employeeId)) {
    return false;
  }
  if (grant.eligibility_mode === "worked_on_date") {
    const workedDates = new Set(parseJsonArray(fact.worked_dates_json).map((value) => typeof value === "string" ? value : value?.date).filter(Boolean));
    return workedDates.has(grant.criterion_date);
  }
  return true;
}

function parseEmployeeIds(value) {
  return parseJsonArray(value).map(normalizeEmployeeId).filter(Boolean);
}

function parseJsonArray(value) {
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function normalizeEmployeeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
