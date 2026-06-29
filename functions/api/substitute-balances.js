import { json, requireAuth } from "../_lib/auth.js";
import { normalizeEmployeeId, parseEmployeeIds, resolveEligibleEmployees, resolveOccurrenceFact } from "../_lib/leave-eligibility.js";
import { ensureSchema } from "../_lib/schema.js";
import { recalculateRoute } from "../_lib/recalculate.js";

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

  // v34 배포 전 저장된 월도 새 휴무초과 자동대체 기준으로 즉시 재계산합니다.
  await recalculateRoute(context.env.DB, route);

  const monthStart = `${month}-01`;
  const monthEnd = endOfMonth(month);
  const previousMonth = shiftMonth(month, -1);
  const [grantsResult, factsResult, allocationsResult, annualUsageResult, workforceResult, annualEmployeeResult, closuresResult, previousSummaryResult] = await Promise.all([
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
    context.env.DB.prepare(`
      SELECT month, cutoff_date FROM attendance_closures WHERE company = ?
    `).bind(route).all(),
    context.env.DB.prepare(`
      SELECT employee_id, base_excess, dayoff_replacement_used, dayoff_replacement_shortage,
             dayoff_replacement_substitute_used, dayoff_replacement_compensation_used
      FROM attendance_monthly_summaries_v3
      WHERE route = ? AND month = ?
    `).bind(route, previousMonth).all(),
  ]);

  const factsByMonth = new Map();
  for (const fact of factsResult.results || []) {
    if (!factsByMonth.has(fact.month)) factsByMonth.set(fact.month, []);
    factsByMonth.get(fact.month).push(fact);
  }
  const factByMonthEmployee = new Map();
  for (const fact of factsResult.results || []) {
    const employeeId = normalizeEmployeeId(fact.employee_id);
    if (employeeId) factByMonthEmployee.set(`${fact.month}|${employeeId}`, fact);
  }
  const closureCutoffByMonth = new Map((closuresResult.results || []).map((row) => [String(row.month || ""), String(row.cutoff_date || "")]));

  const usedMap = new Map((allocationsResult.results || []).map((row) => [
    `${row.grant_id}|${normalizeEmployeeId(row.employee_id)}`,
    roundHalf(row.used_days),
  ]));

  // 대상 월에 사용 가능한 모든 부여분을 직원별 lot로 바로 내려보냅니다.
  // grant_month가 7월이어도 사용 시작일이 5월이면 6월 분석에서 사용할 수 있습니다.
  const grants = grantsResult.results || [];
  const lotsByEmployee = {};
  const settlementGrants = [];
  for (const grant of grants) {
    const occurrenceDate = String(grant.occurrence_date || grant.criterion_date || "");
    const occurrenceMonth = occurrenceDate.slice(0, 7);
    const grantType = grant.grant_type || "substitute";
    const settlementMode = (grant.grant_scope || "route") === "route"
      && /^\d{4}-\d{2}-\d{2}$/.test(occurrenceDate)
      && ["substitute", "compensation"].includes(grantType);
    const baseGrant = settlementMode
      ? { ...grant, eligibility_mode: "all", criterion_date: null }
      : grant;
    const baseEligibility = resolveEligibleEmployees(baseGrant, {
      workforceRows: workforceResult.results || [],
      annualRows: annualEmployeeResult.results || [],
      factsByMonth,
    });

    // 대체휴무는 발생일 근태와 무관하게 동일 부여합니다.
    // 보상휴가는 발생일이 마감된 뒤 실제 출근자만 lot를 만들고, 미출근자는 기본휴무 +1 대상이 됩니다.
    const finalized = closureCutoffByMonth.get(occurrenceMonth) >= occurrenceDate;
    const eligibility = settlementMode && grantType === "compensation"
      ? {
          ...baseEligibility,
          employeeIds: finalized
            ? baseEligibility.employeeIds.filter((employeeId) => {
                const fact = factByMonthEmployee.get(`${occurrenceMonth}|${normalizeEmployeeId(employeeId)}`);
                return fact ? resolveOccurrenceFact(fact, occurrenceDate).hasClockIn : false;
              })
            : [],
        }
      : baseEligibility;

    if (settlementMode && occurrenceMonth === month) {
      settlementGrants.push({
        id: grant.id,
        grantType,
        occurrenceDate,
        grantedDays: roundHalf(grant.granted_days),
        eligibleEmployeeIds: baseEligibility.employeeIds,
      });
    }

    for (const employeeId of eligibility.employeeIds) {
      const remaining = roundHalf(Number(grant.granted_days || 0) - (usedMap.get(`${grant.id}|${employeeId}`) || 0));
      if (!(remaining > 0)) continue;
      if (!lotsByEmployee[employeeId]) lotsByEmployee[employeeId] = [];
      lotsByEmployee[employeeId].push({
        grantId: grant.id,
        grantType: grant.grant_type || "substitute",
        grantMonth: grant.grant_month,
        occurrenceDate,
        settlementMode,
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

  const autoUseDates = [...new Set(grants
    .filter((grant) => (grant.grant_type || "substitute") === "substitute")
    .map((grant) => String(grant.occurrence_date || grant.criterion_date || ""))
    .filter((date) => date >= monthStart && date <= monthEnd))].sort();

  // 같은 월 부여분도 계획·근태 판정 없이 대상자 전원에게 적용할 수 있도록 원본 설정을 함께 내려보냅니다.
  const currentGrants = grants
    .filter((grant) => String(grant.occurrence_date || grant.criterion_date || "").startsWith(month))
    .map((grant) => ({
      id: grant.id,
      grantType: grant.grant_type || "substitute",
      grantScope: grant.grant_scope || "route",
      grantMonth: grant.grant_month,
      occurrenceDate: grant.occurrence_date || grant.criterion_date || "",
      settlementMode: ["substitute", "compensation"].includes(grant.grant_type || "substitute") && (grant.grant_scope || "route") === "route",
      grantedDays: roundHalf(grant.granted_days),
      validFrom: grant.valid_from,
      validTo: grant.valid_to,
      eligibilityMode: grant.eligibility_mode || "worked_on_date",
      criterionDate: grant.criterion_date || grant.occurrence_date || "",
      employeeId: grant.employee_id || "",
      excludedEmployeeIds: parseEmployeeIds(grant.excluded_employee_ids_json),
    }));

  const previousMonthFacts = {};
  for (const row of previousSummaryResult.results || []) {
    previousMonthFacts[normalizeEmployeeId(row.employee_id)] = {
      baseExcess: roundHalf(row.base_excess),
      dayoffReplacementUsed: roundHalf(row.dayoff_replacement_used),
      dayoffReplacementShortage: roundHalf(row.dayoff_replacement_shortage),
      dayoffReplacementSubstituteUsed: roundHalf(row.dayoff_replacement_substitute_used),
      dayoffReplacementCompensationUsed: roundHalf(row.dayoff_replacement_compensation_used),
    };
  }
  return json({ lotsByEmployee, balances, annualLeaveBefore, currentGrants, settlementGrants, autoUseDates, previousMonth, previousMonthFacts });
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}

function shiftMonth(monthText, amount) {
  const [year, month] = monthText.split("-").map(Number);
  const date = new Date(year, month - 1 + amount, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}
