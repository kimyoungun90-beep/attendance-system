import { normalizeEmployeeId, parseEvents, resolveEligibleEmployees, resolveOccurrenceFact } from "./leave-eligibility.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const VALID_GRANT_TYPES = new Set(["substitute", "compensation"]);

export async function recalculateRoute(db, route) {
  if (!VALID_ROUTES.has(route)) throw new Error("경로 구분이 올바르지 않습니다.");

  const [closuresResult, factsResult, grantsResult, workforceResult, annualEmployeeResult] = await Promise.all([
    db.prepare(`
      SELECT id, company AS route, month, cutoff_date
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
    db.prepare(`
      SELECT *
      FROM attendance_workforce_members
      WHERE route = ?
      ORDER BY month ASC, id ASC
    `).bind(route).all(),
    db.prepare(`
      SELECT *
      FROM annual_leave_employees
      WHERE route = ?
    `).bind(route).all(),
  ]);

  const closures = closuresResult.results || [];
  const facts = factsResult.results || [];
  const grants = grantsResult.results || [];
  const workforceRows = workforceResult.results || [];
  const annualEmployeeRows = annualEmployeeResult.results || [];

  if (!closures.length) return { affectedMonths: 0, affectedEmployees: 0 };

  const factsByClosure = new Map();
  for (const fact of facts) {
    if (!factsByClosure.has(fact.closure_id)) factsByClosure.set(fact.closure_id, []);
    factsByClosure.get(fact.closure_id).push(fact);
  }

  const closureByMonth = new Map(closures.map((closure) => [String(closure.month || ""), closure]));
  const factsByMonth = new Map();
  const factByMonthEmployee = new Map();
  for (const fact of facts) {
    if (!factsByMonth.has(fact.month)) factsByMonth.set(fact.month, []);
    factsByMonth.get(fact.month).push(fact);
    const employeeId = normalizeEmployeeId(fact.employee_id);
    if (employeeId) factByMonthEmployee.set(`${fact.month}|${employeeId}`, fact);
  }

  const lotsByEmployee = new Map();
  const compensationRestByMonthEmployee = new Map();
  // 부여분은 월 마감자료가 아니라 인력·매장매칭의 입사일을 기준으로 직원별 lot를 만듭니다.
  // 예: 5월 발생분은 5월 1일 이전 입사자, 6월 발생분은 6월 1일 이전 입사자입니다.
  for (const grant of grants) {
    const occurrenceDate = String(grant.occurrence_date || grant.criterion_date || "");
    const occurrenceMonth = occurrenceDate.slice(0, 7);
    const grantType = VALID_GRANT_TYPES.has(grant.grant_type) ? grant.grant_type : "substitute";
    const settlementMode = (grant.grant_scope || "route") === "route"
      && /^\d{4}-\d{2}-\d{2}$/.test(occurrenceDate)
      && ["substitute", "compensation"].includes(grantType);
    const baseGrant = settlementMode
      ? { ...grant, eligibility_mode: "all", criterion_date: null }
      : grant;
    const baseEligibility = resolveEligibleEmployees(baseGrant, {
      workforceRows,
      annualRows: annualEmployeeRows,
      factsByMonth,
    });
    const occurrenceClosure = closureByMonth.get(occurrenceMonth);
    const finalized = Boolean(
      settlementMode
      && occurrenceClosure
      && String(occurrenceClosure.cutoff_date || "") >= occurrenceDate
    );

    for (const employeeId of baseEligibility.employeeIds) {
      let entitled = true;
      let restEligible = false;
      if (settlementMode && grantType === "compensation") {
        const occurrenceFact = factByMonthEmployee.get(`${occurrenceMonth}|${normalizeEmployeeId(employeeId)}`);
        if (!finalized || !occurrenceFact) {
          entitled = false;
        } else {
          const resolved = resolveOccurrenceFact(occurrenceFact, occurrenceDate);
          entitled = Boolean(resolved.hasClockIn);
          restEligible = !resolved.hasClockIn;
        }
      }

      if (restEligible) {
        const key = `${occurrenceMonth}|${normalizeEmployeeId(employeeId)}`;
        if (!compensationRestByMonthEmployee.has(key)) compensationRestByMonthEmployee.set(key, []);
        compensationRestByMonthEmployee.get(key).push({
          date: occurrenceDate,
          days: 1,
          source: "보상휴가 발생일 미출근 기본휴무 추가",
          grantId: grant.id,
        });
      }
      if (!entitled) continue;

      const grantedDays = roundHalf(grant.granted_days);
      if (!lotsByEmployee.has(employeeId)) lotsByEmployee.set(employeeId, []);
      lotsByEmployee.get(employeeId).push({
        grantId: grant.id,
        grantType,
        grantMonth: grant.grant_month,
        occurrenceDate,
        settlementMode,
        validFrom: grant.valid_from,
        validTo: grant.valid_to,
        grantedDays,
        finalized,
        finalEntitled: true,
        finalRestEligible: restEligible,
        remaining: grantedDays,
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
      // 대체휴무는 기본 휴무를 가감하지 않습니다.
      // 보상휴가 발생일 미출근자는 발생일 1건당 기본 휴무를 1일 추가합니다.
      const occurrenceRestAllowances = compensationRestByMonthEmployee.get(`${closure.month}|${employeeId}`) || [];
      const occurrenceRestDays = roundHalf(occurrenceRestAllowances.reduce((sum, row) => sum + Number(row.days || 0), 0));

      const legacyEvents = parseEvents(fact.substitute_events_json);
      const compensationEvents = mergeEvents(
        parseEvents(fact.compensation_events_json),
        legacyEvents.filter((event) => String(event.planStatus || "").startsWith("보상휴가"))
      );
      const substituteEvents = legacyEvents
        .filter((event) => !String(event.planStatus || "").startsWith("보상휴가"))
        .filter((event) => !["기본 휴무 초과", "발생일 지정 자동 대체휴무"].includes(String(event.source || "")))
        .filter(validUsageEvent)
        .sort(eventSort);

      const combinedPools = consumeCombinedPools({
        lots,
        substituteEvents,
        compensationEvents,
        monthStart,
        monthEnd,
        nextMonthStart,
        closureId: closure.id,
        employeeId,
        allocationTotals,
      });
      const substitute = combinedPools.substitute;
      const compensation = combinedPools.compensation;

      const currentAnnual = roundHalf(fact.annual_leave_used);
      const cumulativeAnnual = roundHalf((cumulativeAnnualByEmployee.get(employeeId) || 0) + currentAnnual);
      cumulativeAnnualByEmployee.set(employeeId, cumulativeAnnual);

      const storedOccurrenceRestDays = roundHalf(fact.occurrence_rest_days);
      const savedBaseAllowance = roundHalf(fact.base_allowance);
      const storedRawAllowance = roundHalf(fact.base_allowance_raw);
      const baseAllowanceRaw = storedRawAllowance > 0
        ? storedRawAllowance
        : roundHalf(Math.max(0, savedBaseAllowance - storedOccurrenceRestDays));
      const baseAllowance = roundHalf(baseAllowanceRaw + occurrenceRestDays);
      const basicDayoffUsed = roundHalf(fact.basic_dayoff_used);
      const baseExcess = roundHalf(Math.max(0, basicDayoffUsed - baseAllowance));
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

  // 모든 계산이 정상 완료된 뒤 기존 계산 결과를 지웁니다.
  // 조회·계산 오류 때문에 기존 부여/잔여 표시가 갑자기 사라지는 현상을 방지합니다.
  await db.batch([
    db.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE route = ?").bind(route),
    db.prepare("DELETE FROM route_substitute_allocations WHERE route = ?").bind(route),
    db.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE route = ?").bind(route),
  ]);
  await runBatch(db, allocationStatements);
  await runBatch(db, summaryStatements);
  await runBatch(db, closureUpdates);

  return { affectedMonths: closures.length, affectedEmployees };
}

function consumeCombinedPools({ lots, substituteEvents, compensationEvents, monthStart, monthEnd, nextMonthStart, closureId, employeeId, allocationTotals }) {
  const availableByType = {
    substitute: roundHalf(lots
      .filter((lot) => lot.grantType === "substitute" && lot.remaining > 0 && lot.validFrom <= monthEnd && lot.validTo >= monthStart)
      .reduce((sum, lot) => sum + lot.remaining, 0)),
    compensation: roundHalf(lots
      .filter((lot) => lot.grantType === "compensation" && lot.remaining > 0 && lot.validFrom <= monthEnd && lot.validTo >= monthStart)
      .reduce((sum, lot) => sum + lot.remaining, 0)),
  };
  const appliedByType = { substitute: 0, compensation: 0 };
  const shortageByOrigin = { substitute: 0, compensation: 0 };
  const events = [
    ...substituteEvents.map((event) => ({ ...event, originType: "substitute" })),
    ...compensationEvents.map((event) => ({ ...event, originType: "compensation" })),
  ].filter(validUsageEvent).sort(eventSort);

  for (const event of events) {
    let need = roundHalf(event.days);
    const candidates = lots
      .filter((lot) => lot.remaining > 0 && lot.validFrom <= event.date && lot.validTo >= event.date)
      .sort((a, b) => {
        const aPriority = a.grantType === event.originType ? 0 : 1;
        const bPriority = b.grantType === event.originType ? 0 : 1;
        return aPriority - bPriority
          || a.validTo.localeCompare(b.validTo)
          || a.validFrom.localeCompare(b.validFrom)
          || a.grantMonth.localeCompare(b.grantMonth);
      });

    for (const lot of candidates) {
      if (need <= 0) break;
      const used = roundHalf(Math.min(need, lot.remaining));
      if (used <= 0) continue;
      lot.remaining = roundHalf(lot.remaining - used);
      need = roundHalf(need - used);
      appliedByType[lot.grantType] = roundHalf((appliedByType[lot.grantType] || 0) + used);
      const key = `${lot.grantId}|${closureId}|${employeeId}|${lot.grantType}`;
      allocationTotals.set(key, roundHalf((allocationTotals.get(key) || 0) + used));
    }
    shortageByOrigin[event.originType] = roundHalf((shortageByOrigin[event.originType] || 0) + Math.max(0, need));
  }

  const poolResult = (grantType, shortage) => ({
    available: availableByType[grantType],
    applied: appliedByType[grantType],
    remaining: roundHalf(lots
      .filter((lot) => lot.grantType === grantType && lot.remaining > 0 && lot.validTo >= nextMonthStart)
      .reduce((sum, lot) => sum + lot.remaining, 0)),
    expired: roundHalf(lots
      .filter((lot) => lot.grantType === grantType && lot.remaining > 0 && lot.validTo >= monthStart && lot.validTo < nextMonthStart)
      .reduce((sum, lot) => sum + lot.remaining, 0)),
    shortage: roundHalf(shortage),
  });

  return {
    substitute: poolResult("substitute", shortageByOrigin.substitute),
    compensation: poolResult("compensation", shortageByOrigin.compensation),
  };
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
  if (label === "대체휴무") {
    const parts = [];
    if (baseExcess > 0) parts.push(`휴무 개수 초과 ${formatDays(baseExcess)}`);
    else parts.push(`기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)}`);
    if (shortage > 0) parts.push(`대체휴무 ${formatDays(shortage)} 초과 사용`);
    else if (needed > 0) parts.push(`대체휴무 ${formatDays(needed)} 사용 · 잔여 ${formatDays(remaining)}`);
    else parts.push(`대체휴무 사용 없음 · 잔여 ${formatDays(remaining)}`);
    if (expired > 0) parts.push(`${formatDays(expired)} 만료`);
    return parts.join(" · ");
  }
  if (shortage > 0) return `${label} ${formatDays(shortage)} 초과 사용`;
  if (needed > 0) {
    const expiredText = expired > 0 ? ` · ${formatDays(expired)} 만료` : "";
    return `${label} ${formatDays(needed)} 사용 · 잔여 ${formatDays(remaining)}${expiredText}`;
  }
  if (expired > 0) return `미사용 ${label} ${formatDays(expired)} 만료`;
  return `${label} 사용 없음 · 잔여 ${formatDays(remaining)}`;
}

function formatDays(value) {
  const number = roundHalf(value);
  return `${Number.isInteger(number) ? number : number.toFixed(1)}일`;
}
