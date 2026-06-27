const VALID_ROUTES = new Set(["homeplus", "electroland"]);

export async function loadGrantRecipientContext(db, route) {
  if (!VALID_ROUTES.has(route)) throw new Error("경로 구분이 올바르지 않습니다.");

  const [workforceResult, factsResult] = await Promise.all([
    db.prepare(`
      SELECT month, route, regional_manager, manager, region1, region2,
             store_code, store_name, employee_id, employee_name,
             hire_date, group_hire_date
      FROM attendance_workforce_members
      WHERE route = ?
      ORDER BY month ASC, employee_id ASC
    `).bind(route).all(),
    db.prepare(`
      SELECT month, employee_id, employee_name, store, worked_dates_json
      FROM attendance_monthly_employee_facts
      WHERE route = ?
      ORDER BY month ASC, employee_id ASC
    `).bind(route).all(),
  ]);

  const candidates = new Map();
  const snapshotsByMonth = new Map();
  const workedDatesByEmployeeMonth = new Map();

  for (const row of workforceResult.results || []) {
    const employeeId = normalizeEmployeeId(row.employee_id);
    const month = normalizeMonth(row.month);
    if (!employeeId || !month) continue;

    const candidate = {
      ...emptyCandidate(employeeId),
      employeeName: clean(row.employee_name),
      regionalManager: clean(row.regional_manager),
      manager: clean(row.manager),
      region1: clean(row.region1),
      region2: clean(row.region2),
      storeCode: clean(row.store_code),
      storeName: clean(row.store_name),
      hireDate: normalizeDate(row.hire_date),
      groupHireDate: normalizeDate(row.group_hire_date),
      firstSeenMonth: month,
      lastSeenMonth: month,
      sourceMonth: month,
    };
    mergeCandidate(candidates, candidate, month);
    addSnapshotCandidate(snapshotsByMonth, month, candidate, true);
  }

  for (const row of factsResult.results || []) {
    const employeeId = normalizeEmployeeId(row.employee_id);
    const month = normalizeMonth(row.month);
    if (!employeeId || !month) continue;

    const candidate = {
      ...emptyCandidate(employeeId),
      employeeName: clean(row.employee_name),
      storeName: clean(row.store),
      firstSeenMonth: month,
      lastSeenMonth: month,
      sourceMonth: month,
    };
    mergeCandidate(candidates, candidate, month);
    // 월 마감 사실은 해당 월 실제 관리 인원으로 보므로 같은 월 인력자료와 합칩니다.
    addSnapshotCandidate(snapshotsByMonth, month, candidate, false);

    const key = `${employeeId}|${month}`;
    const dates = parseJsonArray(row.worked_dates_json)
      .map((value) => typeof value === "string" ? value : value?.date)
      .filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value || "")));
    if (!workedDatesByEmployeeMonth.has(key)) workedDatesByEmployeeMonth.set(key, new Set());
    const dateSet = workedDatesByEmployeeMonth.get(key);
    dates.forEach((date) => dateSet.add(date));
  }

  // 월별 스냅샷에 최신 인적정보(입사일·관리자 등)를 보강합니다.
  for (const snapshot of snapshotsByMonth.values()) {
    for (const [employeeId, person] of snapshot.entries()) {
      snapshot.set(employeeId, mergePerson(person, candidates.get(employeeId) || emptyCandidate(employeeId)));
    }
  }

  return {
    route,
    candidates: [...candidates.values()].sort(candidateSort),
    candidatesById: candidates,
    snapshotsByMonth,
    snapshotMonths: [...snapshotsByMonth.keys()].sort(),
    workedDatesByEmployeeMonth,
  };
}

export function resolveGrantRecipients(grant, context) {
  const grantMonth = normalizeMonth(grant.grant_month ?? grant.grantMonth);
  const grantScope = String((grant.grant_scope ?? grant.grantScope) || "route");
  const employeeId = normalizeEmployeeId(grant.employee_id ?? grant.employeeId);
  const exclusions = new Set(parseEmployeeIds(grant.excluded_employee_ids_json ?? grant.excludedEmployeeIds));
  const eligibilityMode = String((grant.eligibility_mode ?? grant.eligibilityMode) || "all");
  const criterionDate = String((grant.criterion_date ?? grant.criterionDate) || "");
  const cohortEnd = endOfMonth(grantMonth);

  let sourceCandidates;
  let sourceMonth = "";
  if (grantScope === "employee") {
    sourceCandidates = employeeId
      ? [context.candidatesById.get(employeeId) || emptyCandidate(employeeId)]
      : [];
  } else {
    sourceMonth = selectClosestSnapshotMonth(context.snapshotMonths || [], grantMonth);
    const snapshot = sourceMonth ? context.snapshotsByMonth?.get(sourceMonth) : null;
    sourceCandidates = snapshot ? [...snapshot.values()] : context.candidates;
  }

  const unique = new Map();
  for (const candidate of sourceCandidates || []) {
    if (!candidate?.employeeId || exclusions.has(candidate.employeeId)) continue;
    if (grantScope !== "employee" && !isInGrantCohort(candidate, grantMonth, cohortEnd)) continue;
    if (eligibilityMode === "worked_on_date") {
      const criterionMonth = criterionDate.slice(0, 7);
      const worked = context.workedDatesByEmployeeMonth.get(`${candidate.employeeId}|${criterionMonth}`);
      if (!worked?.has(criterionDate)) continue;
    }
    unique.set(candidate.employeeId, { ...candidate, recipientSourceMonth: sourceMonth });
  }
  return [...unique.values()].sort(candidateSort);
}

export function parseEmployeeIds(value) {
  const raw = Array.isArray(value) ? value : parseJsonArray(value);
  const values = raw.length ? raw : String(value || "").split(/[\s,;]+/);
  return [...new Set(values.map(normalizeEmployeeId).filter(Boolean))];
}

export function normalizeEmployeeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}

export function endOfMonth(monthText) {
  if (!/^\d{4}-\d{2}$/.test(String(monthText || ""))) return "";
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function selectClosestSnapshotMonth(months, targetMonth) {
  if (!targetMonth || !months?.length) return "";
  if (months.includes(targetMonth)) return targetMonth;
  const targetIndex = monthIndex(targetMonth);
  return [...months].sort((a, b) => {
    const distanceA = Math.abs(monthIndex(a) - targetIndex);
    const distanceB = Math.abs(monthIndex(b) - targetIndex);
    if (distanceA !== distanceB) return distanceA - distanceB;
    // 같은 거리라면 이후 월 자료를 우선해 퇴사자를 과다 포함하지 않도록 합니다.
    return b.localeCompare(a);
  })[0] || "";
}

function monthIndex(month) {
  const [year, monthNo] = String(month || "").split("-").map(Number);
  return Number.isFinite(year) && Number.isFinite(monthNo) ? year * 12 + monthNo : Number.MAX_SAFE_INTEGER;
}

function isInGrantCohort(candidate, grantMonth, cohortEnd) {
  const hireDate = normalizeDate(candidate.hireDate) || normalizeDate(candidate.groupHireDate);
  if (hireDate) return hireDate <= cohortEnd;
  const firstSeenMonth = normalizeMonth(candidate.firstSeenMonth);
  return Boolean(firstSeenMonth && firstSeenMonth <= grantMonth);
}

function addSnapshotCandidate(snapshotsByMonth, month, candidate, overwrite) {
  if (!snapshotsByMonth.has(month)) snapshotsByMonth.set(month, new Map());
  const snapshot = snapshotsByMonth.get(month);
  const current = snapshot.get(candidate.employeeId);
  if (!current) snapshot.set(candidate.employeeId, { ...candidate });
  else snapshot.set(candidate.employeeId, overwrite ? mergePerson(current, candidate) : mergePerson(candidate, current));
}

function mergeCandidate(candidates, candidate, month) {
  const existing = candidates.get(candidate.employeeId) || emptyCandidate(candidate.employeeId);
  const preferNew = !existing.sourceMonth || month >= existing.sourceMonth;
  const merged = preferNew ? mergePerson(existing, candidate) : mergePerson(candidate, existing);
  merged.firstSeenMonth = minMonth(existing.firstSeenMonth, month);
  merged.lastSeenMonth = maxMonth(existing.lastSeenMonth, month);
  merged.sourceMonth = preferNew ? month : existing.sourceMonth;
  candidates.set(candidate.employeeId, merged);
}

function mergePerson(base, preferred) {
  return {
    ...base,
    employeeId: preferred.employeeId || base.employeeId,
    employeeName: preferred.employeeName || base.employeeName,
    regionalManager: preferred.regionalManager || base.regionalManager,
    manager: preferred.manager || base.manager,
    region1: preferred.region1 || base.region1,
    region2: preferred.region2 || base.region2,
    storeCode: preferred.storeCode || base.storeCode,
    storeName: preferred.storeName || base.storeName,
    hireDate: preferred.hireDate || base.hireDate,
    groupHireDate: preferred.groupHireDate || base.groupHireDate,
    firstSeenMonth: minMonth(base.firstSeenMonth, preferred.firstSeenMonth),
    lastSeenMonth: maxMonth(base.lastSeenMonth, preferred.lastSeenMonth),
    sourceMonth: preferred.sourceMonth || base.sourceMonth,
  };
}

function emptyCandidate(employeeId) {
  return {
    employeeId,
    employeeName: "",
    regionalManager: "",
    manager: "",
    region1: "",
    region2: "",
    storeCode: "",
    storeName: "",
    hireDate: "",
    groupHireDate: "",
    firstSeenMonth: "",
    lastSeenMonth: "",
    sourceMonth: "",
  };
}

function candidateSort(a, b) {
  return String(a.regionalManager || "").localeCompare(String(b.regionalManager || ""), "ko")
    || String(a.manager || "").localeCompare(String(b.manager || ""), "ko")
    || String(a.storeName || "").localeCompare(String(b.storeName || ""), "ko")
    || String(a.employeeName || "").localeCompare(String(b.employeeName || ""), "ko")
    || String(a.employeeId || "").localeCompare(String(b.employeeId || ""));
}

function parseJsonArray(value) {
  if (Array.isArray(value)) return value;
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function normalizeDate(value) {
  const raw = String(value || "").trim().slice(0, 10).replace(/[./]/g, "-");
  const match = raw.match(/^(20\d{2})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return "";
  return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}-${String(Number(match[3])).padStart(2, "0")}`;
}

function normalizeMonth(value) {
  const raw = String(value || "").trim().slice(0, 7);
  return /^\d{4}-\d{2}$/.test(raw) ? raw : "";
}

function minMonth(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a <= b ? a : b;
}

function maxMonth(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a >= b ? a : b;
}

function clean(value) {
  return String(value ?? "").trim().slice(0, 500);
}
