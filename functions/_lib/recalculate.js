const VALID_ROUTES = new Set(["homeplus", "electroland"]);

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
      FROM route_substitute_grants
      WHERE route = ?
      ORDER BY grant_month ASC, valid_to ASC, valid_from ASC, created_at ASC
    `).bind(route).all(),
  ]);

  const closures = closuresResult.results || [];
  const facts = factsResult.results || [];
  const grants = grantsResult.results || [];

  await runBatch(db, [
    db.prepare("DELETE FROM route_substitute_allocations WHERE route = ?").bind(route),
    db.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE route = ?").bind(route),
  ]);

  if (!closures.length) return { affectedMonths: 0, affectedEmployees: 0 };

  const factsByClosure = new Map();
  for (const fact of facts) {
    if (!factsByClosure.has(fact.closure_id)) factsByClosure.set(fact.closure_id, []);
    factsByClosure.get(fact.closure_id).push(fact);
  }

  const grantsByMonth = new Map();
  for (const grant of grants) {
    if (!grantsByMonth.has(grant.grant_month)) grantsByMonth.set(grant.grant_month, []);
    grantsByMonth.get(grant.grant_month).push(grant);
  }

  const lotsByEmployee = new Map();
  const cumulativeAnnualByEmployee = new Map();
  const allocationTotals = new Map();
  const summaryStatements = [];
  const closureUpdates = [];
  let affectedEmployees = 0;

  for (const closure of closures) {
    const monthFacts = factsByClosure.get(closure.id) || [];
    const monthGrants = grantsByMonth.get(closure.month) || [];

    // 발생 월 마감에 포함된 직원에게만 경로 공통 부여분을 생성합니다.
    for (const grant of monthGrants) {
      for (const fact of monthFacts) {
        const employeeId = String(fact.employee_id || "").trim();
        if (!employeeId) continue;
        if (!lotsByEmployee.has(employeeId)) lotsByEmployee.set(employeeId, []);
        lotsByEmployee.get(employeeId).push({
          grantId: grant.id,
          grantMonth: grant.grant_month,
          validFrom: grant.valid_from,
          validTo: grant.valid_to,
          remaining: roundHalf(grant.granted_days),
        });
      }
    }

    let dayoffExcessPeople = 0;
    let shortagePeople = 0;
    let annualLeavePeople = 0;

    for (const fact of monthFacts) {
      const employeeId = String(fact.employee_id || "").trim();
      if (!employeeId) continue;
      affectedEmployees += 1;

      const lots = lotsByEmployee.get(employeeId) || [];
      const monthStart = `${closure.month}-01`;
      const monthEnd = endOfMonth(closure.month);
      const nextMonthStart = startOfNextMonth(closure.month);
      const usageEvents = parseEvents(fact.substitute_events_json)
        .filter((event) => event.date && Number(event.days) > 0)
        .sort((a, b) => a.date.localeCompare(b.date));

      const available = roundHalf(lots
        .filter((lot) => lot.remaining > 0 && lot.validFrom <= monthEnd && lot.validTo >= monthStart)
        .reduce((sum, lot) => sum + lot.remaining, 0));

      let applied = 0;
      let shortage = 0;

      for (const event of usageEvents) {
        let need = roundHalf(event.days);
        const candidates = lots
          .filter((lot) => lot.remaining > 0 && lot.validFrom <= event.date && lot.validTo >= event.date)
          .sort((a, b) => a.validTo.localeCompare(b.validTo) || a.validFrom.localeCompare(b.validFrom) || a.grantMonth.localeCompare(b.grantMonth));

        for (const lot of candidates) {
          if (need <= 0) break;
          const used = roundHalf(Math.min(need, lot.remaining));
          if (used <= 0) continue;
          lot.remaining = roundHalf(lot.remaining - used);
          need = roundHalf(need - used);
          applied = roundHalf(applied + used);
          const key = `${lot.grantId}|${closure.id}|${employeeId}`;
          allocationTotals.set(key, roundHalf((allocationTotals.get(key) || 0) + used));
        }
        shortage = roundHalf(shortage + Math.max(0, need));
      }

      const remaining = roundHalf(lots
        .filter((lot) => lot.remaining > 0 && lot.validTo >= nextMonthStart)
        .reduce((sum, lot) => sum + lot.remaining, 0));
      const expired = roundHalf(lots
        .filter((lot) => lot.remaining > 0 && lot.validTo >= monthStart && lot.validTo < nextMonthStart)
        .reduce((sum, lot) => sum + lot.remaining, 0));

      const currentAnnual = roundHalf(fact.annual_leave_used);
      const cumulativeAnnual = roundHalf((cumulativeAnnualByEmployee.get(employeeId) || 0) + currentAnnual);
      cumulativeAnnualByEmployee.set(employeeId, cumulativeAnnual);

      const baseAllowance = roundHalf(fact.base_allowance);
      const basicDayoffUsed = roundHalf(fact.basic_dayoff_used);
      const explicitSubDayoffUsed = roundHalf(fact.explicit_sub_dayoff_used);
      const baseExcess = roundHalf(fact.base_excess);
      const substituteNeeded = roundHalf(fact.substitute_needed);
      const judgment = buildJudgment({
        baseAllowance,
        basicDayoffUsed,
        explicitSubDayoffUsed,
        baseExcess,
        substituteNeeded,
        available,
        applied,
        remaining,
        expired,
        shortage,
      });

      if (baseExcess > 0) dayoffExcessPeople += 1;
      if (shortage > 0) shortagePeople += 1;
      if (currentAnnual > 0) annualLeavePeople += 1;

      summaryStatements.push(db.prepare(`
        INSERT INTO attendance_monthly_summaries_v3
        (closure_id, route, month, store, employee_id, employee_name,
         base_allowance, basic_dayoff_used, explicit_sub_dayoff_used, base_excess,
         substitute_needed, available_substitute, substitute_applied,
         remaining_substitute, expired_substitute, shortage,
         current_annual_leave, cumulative_annual_leave, judgment)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).bind(
        closure.id, route, closure.month, fact.store || "", employeeId, fact.employee_name || "",
        baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess,
        substituteNeeded, available, applied, remaining, expired, shortage,
        currentAnnual, cumulativeAnnual, judgment
      ));
    }

    closureUpdates.push(db.prepare(`
      UPDATE attendance_closures
      SET dayoff_excess_people = ?, substitute_shortage_people = ?, annual_leave_people = ?
      WHERE id = ?
    `).bind(dayoffExcessPeople, shortagePeople, annualLeavePeople, closure.id));
  }

  const allocationStatements = [];
  for (const [key, usedDays] of allocationTotals.entries()) {
    const [grantId, closureId, employeeId] = key.split("|");
    const closure = closures.find((item) => item.id === closureId);
    allocationStatements.push(db.prepare(`
      INSERT INTO route_substitute_allocations
      (grant_id, closure_id, route, employee_id, month, used_days)
      VALUES (?, ?, ?, ?, ?, ?)
    `).bind(grantId, closureId, route, employeeId, closure?.month || "", usedDays));
  }

  await runBatch(db, allocationStatements);
  await runBatch(db, summaryStatements);
  await runBatch(db, closureUpdates);

  return { affectedMonths: closures.length, affectedEmployees };
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

function formatDays(value) {
  const numberValue = roundHalf(value);
  return `${Number.isInteger(numberValue) ? numberValue : numberValue.toFixed(1)}일`;
}

function buildJudgment({ baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess, substituteNeeded, remaining, expired, shortage }) {
  if (shortage > 0) return `대체휴무 ${formatDays(shortage)} 초과 사용`;
  if (baseExcess > 0) {
    const extra = explicitSubDayoffUsed > 0 ? ` · 표기 대체휴무 ${formatDays(explicitSubDayoffUsed)} 사용` : "";
    const expiry = expired > 0 ? ` · ${formatDays(expired)} 만료` : "";
    return `휴무 개수 초과 ${formatDays(baseExcess)} · 대체휴무 여분 활용${extra} · 잔여 ${formatDays(remaining)}${expiry}`;
  }
  if (substituteNeeded > 0) {
    const expiry = expired > 0 ? ` · ${formatDays(expired)} 만료` : "";
    return `대체휴무 ${formatDays(substituteNeeded)} 사용 · 잔여 ${formatDays(remaining)}${expiry}`;
  }
  if (expired > 0) return `기본 휴무 정상 · 미사용 대체휴무 ${formatDays(expired)} 만료`;
  return `정상 · 기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 잔여 ${formatDays(remaining)}`;
}
