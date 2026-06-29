import { json, requireAuth } from "../_lib/auth.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const VALID_STATUSES = new Set(["확인 요청", "재직·포함", "퇴사", "경로이동", "육아휴직", "기타휴직", "제외"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);
  const url = new URL(context.request.url);
  const month = String(url.searchParams.get("month") || "");
  const route = String(url.searchParams.get("route") || "");
  if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "대상 월을 확인해 주세요." }, 400);
  if (route && !VALID_ROUTES.has(route)) return json({ error: "경로를 확인해 주세요." }, 400);

  const [workforceResult, annualResult, overrideResult] = await Promise.all([
    context.env.DB.prepare(`
      SELECT route, regional_manager, manager, region1, region2, store_code, store_name,
             employee_id, employee_name, hire_date, group_hire_date, note
      FROM attendance_workforce_members WHERE month = ?
      ORDER BY route, employee_id, store_code
    `).bind(month).all(),
    context.env.DB.prepare(`
      SELECT route, employee_id, employee_name, regional_manager, manager, region1, region2,
             store_code, store_name, hire_date, basis_hire_date, termination_date, note
      FROM annual_leave_employees
      ORDER BY route, employee_id
    `).all(),
    context.env.DB.prepare(`
      SELECT month, route, employee_id, employee_name, issue_type, personnel_status,
             effective_from, effective_to, destination_route, note, source_type, updated_at
      FROM attendance_personnel_status_overrides WHERE month <= ?
      ORDER BY month DESC, updated_at DESC, route, employee_id
    `).bind(month).all(),
  ]);

  const items = buildChecks(month, workforceResult.results || [], annualResult.results || [], overrideResult.results || []);
  const filtered = route ? items.filter((item) => item.route === route) : items;
  return json({
    month,
    route,
    items: filtered,
    summary: {
      total: filtered.length,
      unresolved: filtered.filter((item) => !item.resolved).length,
      workforceMissing: filtered.filter((item) => item.issueType === "인력현황 미등록").length,
      annualMissing: filtered.filter((item) => item.issueType === "연차대장 미등록").length,
      routeMismatch: filtered.filter((item) => item.issueType === "경로 불일치").length,
      resigned: filtered.filter((item) => item.personnelStatus === "퇴사").length,
      transferred: filtered.filter((item) => item.personnelStatus === "경로이동").length,
      parentalLeave: filtered.filter((item) => item.personnelStatus === "육아휴직").length,
    },
  });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);
  const body = await context.request.json().catch(() => ({}));
  const month = String(body.month || "");
  if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "대상 월을 확인해 주세요." }, 400);
  const items = normalizeItems(body.items, month, body.sourceType || "manual");
  if (!items.length) return json({ error: "저장할 인력 변동 내역이 없습니다." }, 400);

  const statements = items.map((item) => {
    if (item.personnelStatus === "확인 요청") {
      return context.env.DB.prepare(`
        DELETE FROM attendance_personnel_status_overrides
        WHERE month = ? AND route = ? AND employee_id = ?
      `).bind(month, item.route, item.employeeId);
    }
    return context.env.DB.prepare(`
      INSERT INTO attendance_personnel_status_overrides
      (month, route, employee_id, employee_name, issue_type, personnel_status,
       effective_from, effective_to, destination_route, note, source_type, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
      ON CONFLICT(month, route, employee_id) DO UPDATE SET
        employee_name = excluded.employee_name,
        issue_type = excluded.issue_type,
        personnel_status = excluded.personnel_status,
        effective_from = excluded.effective_from,
        effective_to = excluded.effective_to,
        destination_route = excluded.destination_route,
        note = excluded.note,
        source_type = excluded.source_type,
        updated_at = datetime('now')
    `).bind(
      month, item.route, item.employeeId, item.employeeName, item.issueType, item.personnelStatus,
      item.effectiveFrom || null, item.effectiveTo || null, item.destinationRoute || null,
      item.note, item.sourceType,
    );
  });
  await runBatch(context.env.DB, statements);
  return json({ ok: true, month, saved: items.length });
}

function buildChecks(month, workforceRows, annualRows, overrideRows) {
  const monthStart = `${month}-01`;
  const [year, monthNumber] = month.split("-").map(Number);
  const monthEnd = `${month}-${String(new Date(year, monthNumber, 0).getDate()).padStart(2, "0")}`;
  const workforce = new Map();
  for (const row of workforceRows) {
    const id = normalizeId(row.employee_id);
    if (!id) continue;
    const key = `${row.route}|${id}`;
    if (!workforce.has(key)) workforce.set(key, row);
  }
  const annual = new Map();
  for (const row of annualRows) {
    const id = normalizeId(row.employee_id);
    if (!id) continue;
    annual.set(`${row.route}|${id}`, row);
  }
  const overrides = new Map();
  for (const row of overrideRows) {
    const id = normalizeId(row.employee_id);
    const key = `${row.route}|${id}`;
    if (!id || overrides.has(key)) continue;
    const from = clean(row.effective_from);
    const to = clean(row.effective_to);
    const hasRange = Boolean(from || to);
    const applies = hasRange
      ? (!from || from <= monthEnd) && (!to || to >= monthStart)
      : String(row.month || "") === month;
    if (applies) overrides.set(key, row);
  }
  const employeeRoutes = new Map();
  const addRoute = (id, type, route) => {
    if (!employeeRoutes.has(id)) employeeRoutes.set(id, { workforce: new Set(), annual: new Set() });
    employeeRoutes.get(id)[type].add(route);
  };
  for (const row of workforce.values()) addRoute(normalizeId(row.employee_id), "workforce", row.route);
  for (const row of annual.values()) addRoute(normalizeId(row.employee_id), "annual", row.route);

  const results = new Map();
  const add = (route, id, issueType, wf, al) => {
    const key = `${route}|${id}`;
    const override = overrides.get(key) || {};
    const employeeName = clean(override.employee_name || wf?.employee_name || al?.employee_name);
    const row = {
      month,
      route,
      employeeId: id,
      employeeName,
      regionalManager: clean(wf?.regional_manager || al?.regional_manager),
      manager: clean(wf?.manager || al?.manager),
      region: clean(wf?.region2 || wf?.region1 || al?.region2 || al?.region1),
      storeCode: clean(wf?.store_code || al?.store_code),
      storeName: clean(wf?.store_name || al?.store_name),
      issueType: clean(override.issue_type || issueType || "인력 변동 직접 입력"),
      workforceRoute: wf?.route || "",
      annualRoute: al?.route || "",
      personnelStatus: clean(override.personnel_status || "확인 요청"),
      effectiveFrom: clean(override.effective_from),
      effectiveTo: clean(override.effective_to),
      destinationRoute: clean(override.destination_route),
      note: clean(override.note || wf?.note || al?.note),
      sourceType: clean(override.source_type || "auto"),
      updatedAt: clean(override.updated_at),
    };
    row.resolved = row.personnelStatus !== "확인 요청";
    results.set(key, row);
  };

  for (const [key, wf] of workforce.entries()) {
    const [route, id] = key.split("|");
    const al = annual.get(key);
    if (al) {
      if (al.termination_date && String(al.termination_date) < monthStart) add(route, id, "퇴사자 인력현황 잔존", wf, al);
      continue;
    }
    const otherAnnualRoute = [...(employeeRoutes.get(id)?.annual || [])].find((value) => value !== route);
    const otherAnnual = otherAnnualRoute ? annual.get(`${otherAnnualRoute}|${id}`) : null;
    add(route, id, otherAnnual ? "경로 불일치" : "연차대장 미등록", wf, otherAnnual);
  }
  for (const [key, al] of annual.entries()) {
    const [route, id] = key.split("|");
    const wf = workforce.get(key);
    if (wf) continue;
    if (al.termination_date && String(al.termination_date) < monthStart) continue;
    const otherWorkforceRoute = [...(employeeRoutes.get(id)?.workforce || [])].find((value) => value !== route);
    const otherWorkforce = otherWorkforceRoute ? workforce.get(`${otherWorkforceRoute}|${id}`) : null;
    add(route, id, otherWorkforce ? "경로 불일치" : "인력현황 미등록", otherWorkforce, al);
  }
  for (const [key, override] of overrides.entries()) {
    if (results.has(key)) continue;
    const [route, id] = key.split("|");
    add(route, id, override.issue_type || "인력 변동 직접 입력", workforce.get(key), annual.get(key));
  }
  return [...results.values()].sort((a, b) => a.route.localeCompare(b.route) || a.issueType.localeCompare(b.issueType, "ko") || a.employeeName.localeCompare(b.employeeName, "ko"));
}

function normalizeItems(items, month, defaultSourceType) {
  const map = new Map();
  for (const raw of Array.isArray(items) ? items : []) {
    const route = String(raw.route || "");
    const employeeId = normalizeId(raw.employeeId || raw.employee_id);
    if (!VALID_ROUTES.has(route) || !employeeId) continue;
    const personnelStatus = VALID_STATUSES.has(String(raw.personnelStatus || "")) ? String(raw.personnelStatus) : "확인 요청";
    const sourceType = ["manual", "workforce", "evidence"].includes(String(raw.sourceType || defaultSourceType))
      ? String(raw.sourceType || defaultSourceType) : "manual";
    map.set(`${route}|${employeeId}`, {
      month,
      route,
      employeeId,
      employeeName: clean(raw.employeeName || raw.employee_name),
      issueType: clean(raw.issueType || raw.issue_type),
      personnelStatus,
      effectiveFrom: normalizeDate(raw.effectiveFrom || raw.effective_from),
      effectiveTo: normalizeDate(raw.effectiveTo || raw.effective_to),
      destinationRoute: VALID_ROUTES.has(String(raw.destinationRoute || raw.destination_route)) ? String(raw.destinationRoute || raw.destination_route) : "",
      note: clean(raw.note),
      sourceType,
    });
  }
  return [...map.values()];
}

function normalizeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}
function normalizeDate(value) {
  const raw = clean(value);
  const match = raw.match(/^(20\d{2})[-./]?(\d{1,2})[-./]?(\d{1,2})$/);
  if (!match) return "";
  return `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
}
function clean(value) { return String(value ?? "").trim().slice(0, 1000); }
async function runBatch(db, statements) {
  for (let index = 0; index < statements.length; index += 40) {
    const chunk = statements.slice(index, index + 40);
    if (chunk.length) await db.batch(chunk);
  }
}
