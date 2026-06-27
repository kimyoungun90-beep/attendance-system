import { loadGrantRecipientContext, resolveGrantRecipients } from "./grant-recipients.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const VALID_GRANT_TYPES = new Set(["substitute", "compensation"]);

export async function recalculateRoute(db, route) {
  if (!VALID_ROUTES.has(route)) throw new Error("경로 구분이 올바르지 않습니다.");

  const [closuresResult, factsResult, grantsResult] = await Promise.all([
    db.prepare(`
      SELECT id, company AS route, month
      FROM attendance_closures
      WHERE company = ?
      ORDER BY month ASC, created_at ASC
    `).bind(route).all(),
    db.prepare(`
      SELECT *
      FROM attendance_monthly_employee_facts
      WHERE route = ?
      ORDER BY month ASC, employee_id ASC
    `).bind(route).all(),
    db.prepare(`
      SELECT *
      FROM attendance_leave_grants_v5
      WHERE route = ?
      ORDER BY grant_month ASC, valid_to ASC, valid_from ASC, created_at ASC
    `).bind(route).all(),
  ]);

  const closures = closuresResult.results || [];
  const facts = factsResult.results || [];
  const grants = grantsResult.results || [];

  await runBatch(db, [
    db.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE route = ?").bind(route),
    db.prepare("DELETE FROM route_substitute_allocations WHERE route = ?").bind(route),
    db.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE route = ?").bind(route),
  ]);

  if (!closures.length) return { affectedMonths: 0, affectedEmployees: 0 };

  const factsByClosure = new Map();
  for (const fact of facts) {
    if (!factsByClosure.has(fact.closure_id)) factsByClosure.set(fact.closure_id, []);
    factsByClosure.get(fact.closure_id).push(fact);
  }

  const recipientContext = await loadGrantRecipientContext(db, route);
  const lotsByEmployee = new Map();
  for (const grant of grants) {
    const recipients = resolveGrantRecipients(grant, recipientContext);
    for (const recipient of recipients) {
      const employeeId = normalizeEmployeeId(recipient.employeeId);
      if (!employeeId) continue;
      if (!lotsByEmployee.has(employeeId)) lotsByEmployee.set(employeeId, []);
      lotsByEmployee.get(employeeId).push({
        grantId: grant.id,
        grantType: VALID_GRANT_TYPES.has(grant.grant_type) ? grant.grant_type : "substitute",
        grantMonth: grant.grant_month,
        validFrom: grant.valid_from,
        validTo: grant.valid_to,
        remaining: roundHalf(grant.granted_days),
      });
    }
  }
  const cumulativeAnnualByEmployee = new Map();
  const allocationTotals = new Map();
  const summaryStatements = [];
  const closureUpdates = [];
  let affectedEmployees = 0;

  for (const closure of closures) {
    const monthFacts = factsByClosure.get(closure.id) || [];

    let dayoffExcessPeople = 0;
    let substituteShortagePeople = 0;
    let compensationShortagePeople = 0;
    let annualLeavePeople = 0;

    for (const fact of monthFacts) {
      const employeeId = normalizeEmployeeId(fact.employee_id);
      if (!employeeId) continue;
      affectedEmployees += 1;

      const lots = lotsByEmployee.get(employeeId) || [];
      const monthStart = `${closure.month}-01`;
      const monthEnd = endOfMonth(closure.month);
      const nextMonthStart = startOfNextMonth(closure.month);

      const legacyEvents = parseEvents(fact.substitute_events_json);
      const compensationEvents = mergeEvents(
        parseEvents(fact.compensation_events_json),
        legacyEvents.filter((event) => String(event.planStatus || "").startsWith("보상휴가"))
      );
      const substituteEvents = legacyEvents
        .filter((event) => !String(event.planStatus || "").startsWith("보상휴가"))
        .filter(validUsageEvent)
        .sort(eventSort);

      const substitute = consumePool({
        lots,
        grantType: "substitute",
        events: substituteEvents,
        monthStart,
        monthEnd,
        nextMonthStart,
        closureId: closure.id,
        employeeId,
        allocationTotals,
      });
      const compensation = consumePool({
        lots,
        grantType: "compensation",
        events: compensationEvents,
        monthStart,
        monthEnd,
        nextMonthStart,
        closureId: closure.id,
        employeeId,
        allocationTotals,
      });

      const currentAnnual = roundHalf(fact.annual_leave_used);
      const cumulativeAnnual = roundHalf((cumulativeAnnualByEmployee.get(employeeId) || 0) + currentAnnual);
      cumulativeAnnualByEmployee.set(employeeId, cumulativeAnnual);

      const baseAllowance = roundHalf(fact.base_allowance);
      const basicDayoffUsed = roundHalf(fact.basic_dayoff_used);
      const baseExcess = roundHalf(fact.base_excess);
      const explicitSubDayoffUsed = roundHalf(substituteEvents
        .filter((event) => String(event.planStatus || "").startsWith("대체휴일"))
        .reduce((sum, event) => sum + Number(event.days || 0), 0));
      const substituteNeeded = roundHalf(substituteEvents.reduce((sum, event) => sum + Number(event.days || 0), 0));
      const compensationNeeded = roundHalf(compensationEvents.reduce((sum, event) => sum + Number(event.days || 0), 0));

      const judgment = buildJudgment({
        label: "대체휴무",
        baseAllowance,
        basicDayoffUsed,
        explicitUsed: explicitSubDayoffUsed,
        baseExcess,
        needed: substituteNeeded,
        ...substitute,
      });
      const compensationJudgment = buildJudgment({
        label: "보상휴가",
        baseAllowance: 0,
        basicDayoffUsed: 0,
        explicitUsed: compensationNeeded,
        baseExcess: 0,
        needed: compensationNeeded,
        ...compensation,
      });

      if (baseExcess > 0) dayoffExcessPeople += 1;
      if (substitute.shortage > 0) substituteShortagePeople += 1;
      if (compensation.shortage > 0) compensationShortagePeople += 1;
      if (currentAnnual > 0) annualLeavePeople += 1;

      summaryStatements.push(db.prepare(`
        INSERT INTO attendance_monthly_summaries_v3
        (closure_id, route, month, store, employee_id, employee_name,
         base_allowance, basic_dayoff_used, explicit_sub_dayoff_used, base_excess,
         substitute_needed, available_substitute, substitute_applied,
         remaining_substitute, expired_substitute, shortage,
         compensation_needed, available_compensation, compensation_applied,
         remaining_compensation, expired_compensation, compensation_shortage,
         current_annual_leave, cumulative_annual_leave, judgment, compensation_judgment)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).bind(
        closure.id, route, closure.month, fact.store || "", employeeId, fact.employee_name || "",
        baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess,
        substituteNeeded, substitute.available, substitute.applied, substitute.remaining, substitute.expired, substitute.shortage,
        compensationNeeded, compensation.available, compensation.applied, compensation.remaining, compensation.expired, compensation.shortage,
        currentAnnual, cumulativeAnnual, judgment, compensationJudgment
      ));
    }

    closureUpdates.push(db.prepare(`
      UPDATE attendance_closures
      SET dayoff_excess_people = ?, substitute_shortage_people = ?,
          compensation_shortage_people = ?, annual_leave_people = ?
      WHERE id = ?
    `).bind(dayoffExcessPeople, substituteShortagePeople, compensationShortagePeople, annualLeavePeople, closure.id));
  }

  const allocationStatements = [];
  for (const [key, usedDays] of allocationTotals.entries()) {
    const [grantId, closureId, employeeId, grantType] = key.split("|");
    const closure = closures.find((item) => item.id === closureId);
    allocationStatements.push(db.prepare(`
      INSERT INTO attendance_leave_allocations_v5
      (grant_id, closure_id, route, grant_type, employee_id, month, used_days)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).bind(grantId, closureId, route, grantType, employeeId, closure?.month || "", usedDays));
  }

  await runBatch(db, allocationStatements);
  await runBatch(db, summaryStatements);
  await runBatch(db, closureUpdates);

  return { affectedMonths: closures.length, affectedEmployees };
}

function consumePool({ lots, grantType, events, monthStart, monthEnd, nextMonthStart, closureId, employeeId, allocationTotals }) {
  const typedLots = lots.filter((lot) => lot.grantType === grantType);
  const available = roundHalf(typedLots
    .filter((lot) => lot.remaining > 0 && lot.validFrom <= monthEnd && lot.validTo >= monthStart)
    .reduce((sum, lot) => sum + lot.remaining, 0));

  let applied = 0;
  let shortage = 0;
  for (const event of [...events].filter(validUsageEvent).sort(eventSort)) {
    let need = roundHalf(event.days);
    const candidates = typedLots
      .filter((lot) => lot.remaining > 0 && lot.validFrom <= event.date && lot.validTo >= event.date)
      .sort((a, b) => a.validTo.localeCompare(b.validTo) || a.validFrom.localeCompare(b.validFrom) || a.grantMonth.localeCompare(b.grantMonth));

    for (const lot of candidates) {
      if (need <= 0) break;
      const used = roundHalf(Math.min(need, lot.remaining));
      if (used <= 0) continue;
      lot.remaining = roundHalf(lot.remaining - used);
      need = roundHalf(need - used);
      applied = roundHalf(applied + used);
      const key = `${lot.grantId}|${closureId}|${employeeId}|${grantType}`;
      allocationTotals.set(key, roundHalf((allocationTotals.get(key) || 0) + used));
    }
    shortage = roundHalf(shortage + Math.max(0, need));
  }

  const remaining = roundHalf(typedLots
    .filter((lot) => lot.remaining > 0 && lot.validTo >= nextMonthStart)
    .reduce((sum, lot) => sum + lot.remaining, 0));
  const expired = roundHalf(typedLots
    .filter((lot) => lot.remaining > 0 && lot.validTo >= monthStart && lot.validTo < nextMonthStart)
    .reduce((sum, lot) => sum + lot.remaining, 0));

  return { available, applied, remaining, expired, shortage };
}

function mergeEvents(...groups) {
  const map = new Map();
  for (const event of groups.flat()) {
    if (!validUsageEvent(event)) continue;
    const key = `${event.date}|${roundHalf(event.days)}|${event.planStatus || ""}`;
    if (!map.has(key)) map.set(key, { ...event, days: roundHalf(event.days) });
  }
  return [...map.values()].sort(eventSort);
}

function validUsageEvent(event) {
  return Boolean(event?.date && Number(event.days) > 0);
}

function eventSort(a, b) {
  return String(a.date || "").localeCompare(String(b.date || ""));
}

export async function runBatch(db, statements, chunkSize = 80) {
  for (let index = 0; index < statements.length; index += chunkSize) {
    const chunk = statements.slice(index, index + chunkSize);
    if (chunk.length) await db.batch(chunk);
  }
}

function parseEvents(value) {
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(year, month, 0).getDate()).padStart(2, "0")}`;
}

function startOfNextMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  const date = new Date(year, month, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-01`;
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}

function buildJudgment({ label, baseAllowance, basicDayoffUsed, explicitUsed, baseExcess, needed, remaining, expired, shortage }) {
  if (shortage > 0) return `${label} ${formatDays(shortage)} 초과 사용`;
  if (label === "대체휴무" && baseExcess > 0) {
    const explicit = explicitUsed > 0 ? ` · 표기 대체휴무 ${formatDays(explicitUsed)} 사용` : "";
    const expiredText = expired > 0 ? ` · ${formatDays(expired)} 만료` : "";
    return `휴무 개수 초과 ${formatDays(baseExcess)} · 대체휴무 여분 활용${explicit} · 잔여 ${formatDays(remaining)}${expiredText}`;
  }
  if (needed > 0) {
    const expiredText = expired > 0 ? ` · ${formatDays(expired)} 만료` : "";
    return `${label} ${formatDays(needed)} 사용 · 잔여 ${formatDays(remaining)}${expiredText}`;
  }
  if (expired > 0) return `미사용 ${label} ${formatDays(expired)} 만료`;
  if (label === "대체휴무") return `정상 · 기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 잔여 ${formatDays(remaining)}`;
  return `${label} 사용 없음 · 잔여 ${formatDays(remaining)}`;
}

function formatDays(value) {
  const number = roundHalf(value);
  return `${Number.isInteger(number) ? number : number.toFixed(1)}일`;
}
