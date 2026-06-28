import { json, requireAuth } from "../_lib/auth.js";
import { resolveEligibleEmployees, normalizeEmployeeId, parseEmployeeIds, parseEvents } from "../_lib/leave-eligibility.js";
import { recalculateRoute } from "../_lib/recalculate.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const VALID_TYPES = new Set(["substitute", "compensation"]);
const VALID_SCOPES = new Set(["route", "employee"]);
const VALID_ELIGIBILITY = new Set(["all", "worked_on_date"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const route = url.searchParams.get("route") || "";
  if (route && !VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);

  const routeWhere = route ? " WHERE route = ?" : "";
  const routeBind = (statement) => route ? statement.bind(route) : statement;
  const [grantsResult, factsResult, allocationsResult, workforceResult, annualResult] = await Promise.all([
    routeBind(context.env.DB.prepare(`SELECT * FROM attendance_leave_grants_v5${routeWhere} ORDER BY grant_month DESC, created_at DESC`)).all(),
    routeBind(context.env.DB.prepare(`SELECT * FROM attendance_monthly_employee_facts${routeWhere}`)).all(),
    routeBind(context.env.DB.prepare(`
      SELECT grant_id, ROUND(SUM(used_days), 1) AS used_days
      FROM attendance_leave_allocations_v5${routeWhere}
      GROUP BY grant_id
    `)).all(),
    routeBind(context.env.DB.prepare(`SELECT * FROM attendance_workforce_members${routeWhere} ORDER BY month, id`)).all(),
    routeBind(context.env.DB.prepare(`SELECT * FROM annual_leave_employees${routeWhere}`)).all(),
  ]);

  const grants = grantsResult.results || [];
  if (!grants.length) return json({ items: [] });

  const factsByRouteMonth = new Map();
  for (const fact of factsResult.results || []) {
    const key = `${fact.route}|${fact.month}`;
    if (!factsByRouteMonth.has(key)) factsByRouteMonth.set(key, []);
    factsByRouteMonth.get(key).push(fact);
  }
  const workforceByRoute = groupByRoute(workforceResult.results || []);
  const annualByRoute = groupByRoute(annualResult.results || []);
  const usedByGrant = new Map((allocationsResult.results || []).map((row) => [row.grant_id, roundHalf(row.used_days)]));

  const items = grants.map((grant) => {
    const factsByMonth = new Map();
    for (const [key, rows] of factsByRouteMonth.entries()) {
      const [factRoute, month] = key.split("|");
      if (factRoute === grant.route) factsByMonth.set(month, rows);
    }
    const occurrenceDate = String(grant.occurrence_date || grant.criterion_date || "");
    const occurrenceMonth = occurrenceDate.slice(0, 7);
    const settlementMode = (grant.grant_type || "substitute") === "substitute"
      && (grant.grant_scope || "route") === "route"
      && /^\d{4}-\d{2}-\d{2}$/.test(occurrenceDate);
    const baseEligibility = resolveEligibleEmployees(settlementMode ? { ...grant, eligibility_mode: "all", criterion_date: null } : grant, {
      workforceRows: workforceByRoute.get(grant.route) || [],
      annualRows: annualByRoute.get(grant.route) || [],
      factsByMonth,
    });
    let eligibility = baseEligibility;
    const occurrenceFacts = factsByMonth.get(occurrenceMonth) || [];
    if (settlementMode && occurrenceFacts.length) {
      const entitled = new Set();
      for (const fact of occurrenceFacts) {
        const employeeId = normalizeEmployeeId(fact.employee_id);
        if (!employeeId) continue;
        const saved = parseEvents(fact.occurrence_substitute_dates_json)
          .map((value) => typeof value === "string" ? value : value?.date)
          .filter(Boolean);
        const fallback = parseEvents(fact.worked_dates_json)
          .map((value) => typeof value === "string" ? value : value?.date)
          .filter(Boolean);
        if ((saved.length ? saved : fallback).includes(occurrenceDate)) entitled.add(employeeId);
      }
      eligibility = {
        ...baseEligibility,
        employeeIds: baseEligibility.employeeIds.filter((employeeId) => entitled.has(normalizeEmployeeId(employeeId))),
      };
    }
    const usedDays = usedByGrant.get(grant.id) || 0;
    const assignedDays = roundHalf(roundHalf(grant.granted_days) * eligibility.employeeIds.length);
    return {
      ...grant,
      settlement_mode: settlementMode ? 1 : 0,
      provisional_people: baseEligibility.employeeIds.length,
      eligible_people: eligibility.employeeIds.length,
      assigned_days: assignedDays,
      used_days: usedDays,
      unused_days: roundHalf(Math.max(0, assignedDays - usedDays)),
      excluded_count: parseEmployeeIds(grant.excluded_employee_ids_json).length,
      eligibility_cutoff: eligibility.cutoffDate,
      missing_hire_count: eligibility.missingHireCount,
    };
  });

  return json({ items });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const body = await context.request.json().catch(() => null);
  const route = String(body?.route || "");
  const grantType = String(body?.grantType || "substitute");
  const grantScope = String(body?.grantScope || "route");
  const grantMonth = String(body?.grantMonth || "");
  const grantedDays = roundHalf(body?.grantedDays);
  const validFrom = String(body?.validFrom || "");
  const validTo = String(body?.validTo || "");
  const eligibilityMode = String(body?.eligibilityMode || "all");
  const employeeId = normalizeEmployeeId(body?.employeeId);
  const excludedEmployeeIds = normalizeEmployeeIdList(body?.excludedEmployeeIds);
  const editingId = String(body?.id || "").trim();
  const occurrenceDates = normalizeOccurrenceDates(body?.occurrenceDates ?? body?.occurrenceDate, grantMonth);

  if (!body || !VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);
  if (!VALID_TYPES.has(grantType)) return json({ error: "부여 구분을 확인해 주세요." }, 400);
  if (!VALID_SCOPES.has(grantScope)) return json({ error: "부여 방식을 확인해 주세요." }, 400);
  if (!VALID_ELIGIBILITY.has(eligibilityMode)) return json({ error: "부여 대상 판정 방식을 확인해 주세요." }, 400);
  if (!/^\d{4}-\d{2}$/.test(grantMonth) || !(grantedDays > 0)) {
    return json({ error: "발생 월과 부여 일수를 확인해 주세요. 0.5일 단위로 입력할 수 있습니다." }, 400);
  }
  if (!occurrenceDates.length || occurrenceDates.some((date) => !/^\d{4}-\d{2}-\d{2}$/.test(date) || !date.startsWith(grantMonth))) {
    return json({ error: "발생일은 발생 월 안의 날짜로 입력해 주세요. 여러 날짜는 줄바꿈이나 쉼표로 구분할 수 있습니다." }, 400);
  }
  if (editingId && occurrenceDates.length !== 1) {
    return json({ error: "기존 부여 기록 수정 시에는 발생일을 하나만 입력해 주세요." }, 400);
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(validFrom) || !/^\d{4}-\d{2}-\d{2}$/.test(validTo) || validFrom > validTo) {
    return json({ error: "사용 시작일과 종료일을 확인해 주세요." }, 400);
  }
  if (grantScope === "employee" && !employeeId) {
    return json({ error: "사번별 부여는 대상 사번을 입력해야 합니다." }, 400);
  }

  const explicitCriterionDate = String(body?.criterionDate || "").trim();
  if (eligibilityMode === "worked_on_date" && explicitCriterionDate && !/^\d{4}-\d{2}-\d{2}$/.test(explicitCriterionDate)) {
    return json({ error: "실제 출근 기준일을 날짜 형식으로 입력해 주세요." }, 400);
  }

  const existing = editingId
    ? await context.env.DB.prepare("SELECT id, route FROM attendance_leave_grants_v5 WHERE id = ?").bind(editingId).first()
    : null;
  if (editingId && !existing) return json({ error: "수정할 부여 기록을 찾지 못했습니다." }, 404);

  const saveOne = async (occurrenceDate, id) => {
    const criterionDate = eligibilityMode === "worked_on_date" ? (occurrenceDates.length > 1 ? occurrenceDate : (explicitCriterionDate || occurrenceDate)) : null;
    if (criterionDate && !criterionDate.startsWith(grantMonth)) {
      throw new Error("실제 출근 기준일은 발생 월 안의 날짜로 입력해 주세요.");
    }
    const values = [
      route,
      grantType,
      grantScope,
      grantMonth,
      occurrenceDate,
      grantedDays,
      validFrom,
      validTo,
      eligibilityMode,
      criterionDate,
      grantScope === "employee" ? employeeId : null,
      JSON.stringify(grantScope === "route" ? excludedEmployeeIds : []),
      String(body.reason || "").trim(),
      String(body.note || "").trim(),
    ];
    if (existing) {
      await context.env.DB.prepare(`
        UPDATE attendance_leave_grants_v5
        SET route = ?, grant_type = ?, grant_scope = ?, grant_month = ?, occurrence_date = ?, granted_days = ?,
            valid_from = ?, valid_to = ?, eligibility_mode = ?, criterion_date = ?, employee_id = ?,
            excluded_employee_ids_json = ?, reason = ?, note = ?, updated_at = datetime('now')
        WHERE id = ?
      `).bind(...values, id).run();
    } else {
      await context.env.DB.prepare(`
        INSERT INTO attendance_leave_grants_v5
        (id, route, grant_type, grant_scope, grant_month, occurrence_date, granted_days, valid_from, valid_to,
         eligibility_mode, criterion_date, employee_id, excluded_employee_ids_json, reason, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).bind(id, ...values).run();
    }
  };

  const savedIds = [];
  if (existing) {
    await saveOne(occurrenceDates[0], existing.id);
    savedIds.push(existing.id);
  } else {
    for (const occurrenceDate of occurrenceDates) {
      const id = crypto.randomUUID();
      await saveOne(occurrenceDate, id);
      savedIds.push(id);
    }
  }

  const affectedRoutes = new Set([route]);
  if (existing?.route) affectedRoutes.add(existing.route);
  let affectedMonths = 0;
  let recalculateWarning = "";
  for (const affectedRoute of affectedRoutes) {
    try {
      const recalculated = await recalculateRoute(context.env.DB, affectedRoute);
      affectedMonths += recalculated.affectedMonths;
    } catch (error) {
      // 부여 기록 저장 자체는 성공했으므로 기록을 숨기거나 실패로 돌리지 않습니다.
      recalculateWarning = error?.message || "누적 재계산은 다음 월 마감 저장 시 다시 실행됩니다.";
      console.error("leave grant recalculation failed", affectedRoute, error);
    }
  }

  return json({ ok: true, id: savedIds[0], ids: savedIds, createdCount: savedIds.length, replaced: Boolean(existing), affectedMonths, recalculateWarning }, 201);
}

function groupByRoute(rows) {
  const map = new Map();
  for (const row of rows) {
    if (!map.has(row.route)) map.set(row.route, []);
    map.get(row.route).push(row);
  }
  return map;
}

function normalizeEmployeeIdList(value) {
  const values = Array.isArray(value) ? value : String(value || "").split(/[\s,;]+/);
  return [...new Set(values.map(normalizeEmployeeId).filter(Boolean))];
}


function normalizeOccurrenceDates(value, grantMonth) {
  const values = Array.isArray(value) ? value : String(value || "").split(/[\s,;]+/);
  const normalized = [...new Set(values.map((item) => String(item || "").trim()).filter(Boolean))];
  return normalized.length ? normalized : (grantMonth ? [`${grantMonth}-01`] : []);
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
