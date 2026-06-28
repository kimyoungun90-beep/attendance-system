import { json, requireAuth } from "../_lib/auth.js";
import { resolveEligibleEmployees, normalizeEmployeeId, parseEmployeeIds } from "../_lib/leave-eligibility.js";
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

  // 배포 전에 저장된 0명 부여·0일 잔여도 새 입사일 기준으로 즉시 다시 계산합니다.
  const routesToRecalculate = route ? [route] : [...VALID_ROUTES];
  for (const targetRoute of routesToRecalculate) {
    await recalculateRoute(context.env.DB, targetRoute);
  }

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
    const eligibility = resolveEligibleEmployees(grant, {
      workforceRows: workforceByRoute.get(grant.route) || [],
      annualRows: annualByRoute.get(grant.route) || [],
      factsByMonth,
    });
    const usedDays = usedByGrant.get(grant.id) || 0;
    const assignedDays = roundHalf(roundHalf(grant.granted_days) * eligibility.employeeIds.length);
    return {
      ...grant,
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
  const criterionDate = String(body?.criterionDate || "").trim();
  const employeeId = normalizeEmployeeId(body?.employeeId);
  const excludedEmployeeIds = normalizeEmployeeIdList(body?.excludedEmployeeIds);
  const editingId = String(body?.id || "").trim();

  if (!body || !VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);
  if (!VALID_TYPES.has(grantType)) return json({ error: "부여 구분을 확인해 주세요." }, 400);
  if (!VALID_SCOPES.has(grantScope)) return json({ error: "부여 방식을 확인해 주세요." }, 400);
  if (!VALID_ELIGIBILITY.has(eligibilityMode)) return json({ error: "부여 대상 판정 방식을 확인해 주세요." }, 400);
  if (!/^\d{4}-\d{2}$/.test(grantMonth) || !(grantedDays > 0)) {
    return json({ error: "발생 월과 부여 일수를 확인해 주세요. 0.5일 단위로 입력할 수 있습니다." }, 400);
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(validFrom) || !/^\d{4}-\d{2}-\d{2}$/.test(validTo) || validFrom > validTo) {
    return json({ error: "사용 시작일과 종료일을 확인해 주세요." }, 400);
  }
  if (grantScope === "employee" && !employeeId) {
    return json({ error: "사번별 부여는 대상 사번을 입력해야 합니다." }, 400);
  }
  if (eligibilityMode === "worked_on_date") {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(criterionDate)) return json({ error: "실제 출근자만 부여하려면 출근 기준일을 입력해 주세요." }, 400);
    if (!criterionDate.startsWith(grantMonth)) return json({ error: "출근 기준일은 발생 월 안의 날짜로 입력해 주세요." }, 400);
  }

  const existing = editingId
    ? await context.env.DB.prepare("SELECT id, route FROM attendance_leave_grants_v5 WHERE id = ?").bind(editingId).first()
    : null;
  if (editingId && !existing) return json({ error: "수정할 부여 기록을 찾지 못했습니다." }, 404);
  const id = existing?.id || crypto.randomUUID();

  const values = [
    route,
    grantType,
    grantScope,
    grantMonth,
    grantedDays,
    validFrom,
    validTo,
    eligibilityMode,
    eligibilityMode === "worked_on_date" ? criterionDate : null,
    grantScope === "employee" ? employeeId : null,
    JSON.stringify(grantScope === "route" ? excludedEmployeeIds : []),
    String(body.reason || "").trim(),
    String(body.note || "").trim(),
  ];

  if (existing) {
    await context.env.DB.prepare(`
      UPDATE attendance_leave_grants_v5
      SET route = ?, grant_type = ?, grant_scope = ?, grant_month = ?, granted_days = ?,
          valid_from = ?, valid_to = ?, eligibility_mode = ?, criterion_date = ?, employee_id = ?,
          excluded_employee_ids_json = ?, reason = ?, note = ?, updated_at = datetime('now')
      WHERE id = ?
    `).bind(...values, id).run();
  } else {
    await context.env.DB.prepare(`
      INSERT INTO attendance_leave_grants_v5
      (id, route, grant_type, grant_scope, grant_month, granted_days, valid_from, valid_to,
       eligibility_mode, criterion_date, employee_id, excluded_employee_ids_json, reason, note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(id, ...values).run();
  }

  const affectedRoutes = new Set([route]);
  if (existing?.route) affectedRoutes.add(existing.route);
  let affectedMonths = 0;
  for (const affectedRoute of affectedRoutes) {
    const recalculated = await recalculateRoute(context.env.DB, affectedRoute);
    affectedMonths += recalculated.affectedMonths;
  }

  return json({ ok: true, id, replaced: Boolean(existing), affectedMonths }, 201);
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

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}
