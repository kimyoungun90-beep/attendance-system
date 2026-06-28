const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

export function normalizeEmployeeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}

export function parseEmployeeIds(value) {
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed.map(normalizeEmployeeId).filter(Boolean) : [];
  } catch {
    return [];
  }
}

export function parseEvents(value) {
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

/**
 * 경로 일괄 부여의 대상자를 계산합니다.
 *
 * - all: 발생 월 1일보다 입사일이 빠른 직원
 * - worked_on_date: 발생 월 마감자료에서 기준일에 실제 출근한 직원
 * - employee: 지정 사번 1명
 *
 * workforceRows는 여러 월 자료가 섞여 있어도 됩니다. 직원별로 발생 월과 가장 가까운
 * 인력 스냅샷을 선택하고, 연차 초본의 입사일을 보조값으로 사용합니다.
 */
export function resolveEligibleEmployees(grant, {
  workforceRows = [],
  annualRows = [],
  factsByMonth = new Map(),
} = {}) {
  const excluded = new Set(parseEmployeeIds(grant?.excluded_employee_ids_json));
  const grantScope = String(grant?.grant_scope || "route");
  const eligibilityMode = String(grant?.eligibility_mode || "all");
  const cutoffDate = `${String(grant?.grant_month || "")}-01`;

  if (grantScope === "employee") {
    const employeeId = normalizeEmployeeId(grant?.employee_id);
    return {
      employeeIds: employeeId ? [employeeId] : [],
      cutoffDate,
      missingHireCount: 0,
    };
  }

  if (eligibilityMode === "worked_on_date") {
    const facts = factsByMonth.get(String(grant?.grant_month || "")) || [];
    const employeeIds = [];
    const seen = new Set();
    for (const fact of facts) {
      const employeeId = normalizeEmployeeId(fact?.employee_id);
      if (!employeeId || seen.has(employeeId) || excluded.has(employeeId)) continue;
      const workedDates = new Set(parseEvents(fact?.worked_dates_json)
        .map((value) => typeof value === "string" ? value : value?.date)
        .filter(Boolean));
      if (!workedDates.has(grant?.criterion_date)) continue;
      seen.add(employeeId);
      employeeIds.push(employeeId);
    }
    return { employeeIds, cutoffDate, missingHireCount: 0 };
  }

  const roster = buildRosterForMonth(workforceRows, annualRows, String(grant?.grant_month || ""));
  const employeeIds = [];
  let missingHireCount = 0;
  for (const [employeeId, member] of roster.entries()) {
    if (!employeeId || excluded.has(employeeId)) continue;
    const hireDate = firstValidDate(member.hire_date, member.group_hire_date, member.basis_hire_date);
    if (!hireDate) {
      missingHireCount += 1;
      continue;
    }
    // “5월 1일 이전 입사자”처럼 발생 월 1일 당일 입사자는 제외합니다.
    if (hireDate >= cutoffDate) continue;
    const terminationDate = firstValidDate(member.termination_date);
    if (terminationDate && terminationDate < String(grant?.valid_from || cutoffDate)) continue;
    employeeIds.push(employeeId);
  }
  employeeIds.sort();
  return { employeeIds, cutoffDate, missingHireCount };
}

export function buildRosterForMonth(workforceRows = [], annualRows = [], targetMonth = "") {
  const annualByEmployee = new Map();
  for (const row of annualRows || []) {
    const employeeId = normalizeEmployeeId(row?.employee_id);
    if (employeeId) annualByEmployee.set(employeeId, row);
  }

  // 인력자료는 월별 “전체 명단 스냅샷”이므로 직원별 과거 행을 무기한 이월하지 않습니다.
  // 발생 월 이하의 최신 스냅샷을 사용하고, 없으면 발생 월 이후의 가장 가까운 스냅샷을 사용합니다.
  const availableMonths = [...new Set((workforceRows || []).map((row) => String(row?.month || "")).filter(Boolean))].sort();
  const selectedMonth = selectClosestMonth(availableMonths, targetMonth);
  const selectedRows = selectedMonth
    ? (workforceRows || []).filter((row) => String(row?.month || "") === selectedMonth)
    : [];

  const selectedByEmployee = new Map();
  for (const row of selectedRows) {
    const employeeId = normalizeEmployeeId(row?.employee_id);
    if (employeeId) selectedByEmployee.set(employeeId, row);
  }

  const employeeIds = selectedByEmployee.size
    ? new Set(selectedByEmployee.keys())
    : new Set(annualByEmployee.keys());
  const roster = new Map();
  for (const employeeId of employeeIds) {
    const selected = selectedByEmployee.get(employeeId) || {};
    const annual = annualByEmployee.get(employeeId) || {};
    roster.set(employeeId, {
      ...(annual || {}),
      ...(selected || {}),
      employee_id: employeeId,
      hire_date: selected?.hire_date || annual?.hire_date || "",
      group_hire_date: selected?.group_hire_date || "",
      basis_hire_date: annual?.basis_hire_date || "",
      termination_date: annual?.termination_date || "",
    });
  }
  return roster;
}

function selectClosestMonth(months, targetMonth) {
  if (!months?.length) return "";
  const atOrBefore = months.filter((month) => month <= targetMonth);
  if (atOrBefore.length) return atOrBefore[atOrBefore.length - 1];
  return months[0];
}

function firstValidDate(...values) {
  for (const value of values) {
    const date = String(value || "").trim().slice(0, 10);
    if (ISO_DATE_RE.test(date)) return date;
  }
  return "";
}
