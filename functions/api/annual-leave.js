import { json, requireAuth } from "../_lib/auth.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const APPROVED_STATUS = new Set(["승인", "승인완료", "완료"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const route = String(url.searchParams.get("route") || "");
  const month = String(url.searchParams.get("month") || "");
  const asOf = normalizeDate(url.searchParams.get("asOf")) || (month ? endOfMonth(month) : todayISO());
  if (!VALID_ROUTES.has(route)) return json({ error: "경로를 확인해 주세요." }, 400);
  if (month && !/^\d{4}-\d{2}$/.test(month)) return json({ error: "대상 월을 확인해 주세요." }, 400);

  const baseline = await context.env.DB.prepare(`
    SELECT route, baseline_date, file_name, employee_count, created_at, updated_at
    FROM annual_leave_baseline_uploads WHERE route = ? LIMIT 1
  `).bind(route).first();

  const [employeesResult, applicationsResult, monthlyUploadsResult] = await Promise.all([
    context.env.DB.prepare(`
      SELECT * FROM annual_leave_employees
      WHERE route = ?
      ORDER BY under_one_year, regional_manager, manager, region2, store_name, employee_name
    `).bind(route).all(),
    context.env.DB.prepare(`
      SELECT route, month, employee_id, employee_name, leave_date, days, status, leave_type,
             application_date, note, source_index
      FROM annual_leave_applications
      WHERE route = ? AND leave_date <= ?
      ORDER BY leave_date, employee_id, source_index
    `).bind(route, asOf).all(),
    context.env.DB.prepare(`
      SELECT route, month, file_name, row_count, approved_days, rejected_days, pending_days, created_at, updated_at
      FROM annual_leave_monthly_uploads
      WHERE route = ? ORDER BY month DESC LIMIT 36
    `).bind(route).all(),
  ]);

  const applications = (applicationsResult.results || []).map(toClientApplication);
  const byEmployee = new Map();
  for (const row of applications) {
    if (!byEmployee.has(row.employeeId)) byEmployee.set(row.employeeId, []);
    byEmployee.get(row.employeeId).push(row);
  }

  const monthStart = month ? `${month}-01` : startOfMonth(asOf);
  const openingDate = addDays(monthStart, -1);
  const employees = [];
  const reminders = [];
  for (const raw of employeesResult.results || []) {
    const employee = toClientEmployee(raw);
    const employeeApps = byEmployee.get(employee.employeeId) || [];
    const current = calculateBalance(employee, employeeApps, asOf);
    const opening = calculateBalance(employee, employeeApps, openingDate);
    const currentMonthApps = employeeApps.filter((item) => item.leaveDate >= monthStart && item.leaveDate <= asOf);
    const approvedCurrentMonth = roundHalf(currentMonthApps.filter(isApprovedAnnual).reduce((sum, item) => sum + item.days, 0));
    const requestedCurrentMonth = roundHalf(currentMonthApps.filter(isAnnualType).reduce((sum, item) => sum + item.days, 0));
    const record = {
      ...employee,
      ...current,
      openingRemaining: opening.remaining,
      approvedCurrentMonth,
      requestedCurrentMonth,
    };
    employees.push(record);

    if (!employee.underOneYear && current.remaining > 0) {
      for (const [promotionType, dueDate] of [["1차(6개월)", current.firstPromotionDate], ["2차(9개월)", current.secondPromotionDate]]) {
        if (!dueDate) continue;
        const visibleFrom = addDays(dueDate, -31);
        const visibleTo = addMonths(dueDate, 1);
        if (asOf >= visibleFrom && asOf <= visibleTo) {
          reminders.push({
            route,
            employeeId: employee.employeeId,
            employeeName: employee.employeeName,
            regionalManager: employee.regionalManager,
            manager: employee.manager,
            storeName: employee.storeName,
            promotionType,
            dueDate,
            remaining: current.remaining,
            cycleStart: current.cycleStart,
            cycleEnd: current.cycleEnd,
            status: asOf < dueDate ? "예정" : "촉진 진행",
          });
        }
      }
    }
  }

  const currentMonthApplications = applications.filter((item) => !month || item.month === month);
  return json({
    route,
    month,
    asOf,
    baseline: baseline || null,
    employees,
    reminders: reminders.sort((a, b) => a.dueDate.localeCompare(b.dueDate) || a.employeeName.localeCompare(b.employeeName, "ko")),
    applications: currentMonthApplications,
    monthlyUploads: monthlyUploadsResult.results || [],
    summary: {
      employeeCount: employees.length,
      underOneYearCount: employees.filter((item) => item.underOneYear).length,
      totalRemaining: roundHalf(employees.reduce((sum, item) => sum + Number(item.remaining || 0), 0)),
      approvedCurrentMonth: roundHalf(employees.reduce((sum, item) => sum + Number(item.approvedCurrentMonth || 0), 0)),
      reminderCount: reminders.length,
    },
  });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const body = await context.request.json().catch(() => ({}));
  const action = String(body.action || "");
  if (action === "baseline") return saveBaseline(context, body);
  if (action === "monthly") return saveMonthly(context, body);
  return json({ error: "저장 구분이 올바르지 않습니다." }, 400);
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);
  const url = new URL(context.request.url);
  const route = String(url.searchParams.get("route") || "");
  const kind = String(url.searchParams.get("kind") || "");
  const month = String(url.searchParams.get("month") || "");
  if (!VALID_ROUTES.has(route)) return json({ error: "경로를 확인해 주세요." }, 400);

  if (kind === "monthly") {
    if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "대상 월을 확인해 주세요." }, 400);
    await context.env.DB.batch([
      context.env.DB.prepare(`DELETE FROM annual_leave_applications WHERE route = ? AND month = ?`).bind(route, month),
      context.env.DB.prepare(`DELETE FROM annual_leave_monthly_uploads WHERE route = ? AND month = ?`).bind(route, month),
    ]);
    return json({ ok: true, kind, route, month });
  }
  if (kind === "baseline") {
    await context.env.DB.batch([
      context.env.DB.prepare(`DELETE FROM annual_leave_employees WHERE route = ?`).bind(route),
      context.env.DB.prepare(`DELETE FROM annual_leave_baseline_uploads WHERE route = ?`).bind(route),
    ]);
    return json({ ok: true, kind, route });
  }
  return json({ error: "삭제 구분을 확인해 주세요." }, 400);
}

async function saveBaseline(context, body) {
  const route = String(body.route || "");
  const baselineDate = normalizeDate(body.baselineDate);
  const fileName = clean(body.fileName, 200);
  if (!VALID_ROUTES.has(route) || !baselineDate) return json({ error: "경로와 기준일을 확인해 주세요." }, 400);
  const employees = normalizeEmployees(body.employees, route, baselineDate);
  if (!employees.length) return json({ error: "연차 기준대장에서 저장할 직원을 찾지 못했습니다." }, 400);

  const existed = await context.env.DB.prepare(`SELECT route FROM annual_leave_baseline_uploads WHERE route = ?`).bind(route).first();
  await context.env.DB.batch([
    context.env.DB.prepare(`DELETE FROM annual_leave_employees WHERE route = ?`).bind(route),
    context.env.DB.prepare(`DELETE FROM annual_leave_baseline_uploads WHERE route = ?`).bind(route),
    context.env.DB.prepare(`
      INSERT INTO annual_leave_baseline_uploads
      (route, baseline_date, file_name, employee_count, created_at, updated_at)
      VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
    `).bind(route, baselineDate, fileName, employees.length),
  ]);

  const statements = employees.map((item) => context.env.DB.prepare(`
    INSERT INTO annual_leave_employees
    (route, employee_id, employee_name, regional_manager, manager, region1, region2,
     store_code, store_name, portal_id, hire_date, basis_hire_date, policy_type,
     under_one_year, baseline_date, baseline_granted, baseline_used, baseline_remaining,
     cycle_start, cycle_end, first_promotion_date, second_promotion_date, expiry_date,
     termination_date, note, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
  `).bind(
    route, item.employeeId, item.employeeName, item.regionalManager, item.manager, item.region1, item.region2,
    item.storeCode, item.storeName, item.portalId, item.hireDate || null, item.basisHireDate || null, item.policyType,
    item.underOneYear ? 1 : 0, baselineDate, item.baselineGranted, item.baselineUsed, item.baselineRemaining,
    item.cycleStart || null, item.cycleEnd || null, item.firstPromotionDate || null, item.secondPromotionDate || null, item.expiryDate || null,
    item.terminationDate || null, item.note
  ));
  await runBatch(context.env.DB, statements);
  return json({ ok: true, replaced: Boolean(existed), route, baselineDate, employeeCount: employees.length }, existed ? 200 : 201);
}

async function saveMonthly(context, body) {
  const route = String(body.route || "");
  const month = String(body.month || "");
  const fileName = clean(body.fileName, 200);
  if (!VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month)) return json({ error: "경로와 대상 월을 확인해 주세요." }, 400);
  const applications = normalizeApplications(body.applications, route, month);
  if (!applications.length) return json({ error: `${month} 연차 신청 자료를 찾지 못했습니다.` }, 400);

  const existed = await context.env.DB.prepare(`SELECT route FROM annual_leave_monthly_uploads WHERE route = ? AND month = ?`).bind(route, month).first();
  const approvedDays = roundHalf(applications.filter(isApprovedAnnual).reduce((sum, item) => sum + item.days, 0));
  const rejectedDays = roundHalf(applications.filter((item) => isAnnualType(item) && /반려|미사용|취소/.test(item.status)).reduce((sum, item) => sum + item.days, 0));
  const pendingDays = roundHalf(applications.filter((item) => isAnnualType(item) && !isApprovedAnnual(item) && !/반려|미사용|취소/.test(item.status)).reduce((sum, item) => sum + item.days, 0));

  await context.env.DB.batch([
    context.env.DB.prepare(`DELETE FROM annual_leave_applications WHERE route = ? AND month = ?`).bind(route, month),
    context.env.DB.prepare(`DELETE FROM annual_leave_monthly_uploads WHERE route = ? AND month = ?`).bind(route, month),
    context.env.DB.prepare(`
      INSERT INTO annual_leave_monthly_uploads
      (route, month, file_name, row_count, approved_days, rejected_days, pending_days, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
    `).bind(route, month, fileName, applications.length, approvedDays, rejectedDays, pendingDays),
  ]);

  const statements = applications.map((item) => context.env.DB.prepare(`
    INSERT INTO annual_leave_applications
    (route, month, employee_id, employee_name, leave_date, days, status, leave_type,
     application_date, note, source_index, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
  `).bind(
    route, month, item.employeeId, item.employeeName, item.leaveDate, item.days, item.status,
    item.leaveType, item.applicationDate || null, item.note, item.sourceIndex
  ));
  await runBatch(context.env.DB, statements);
  return json({ ok: true, replaced: Boolean(existed), route, month, rowCount: applications.length, approvedDays, rejectedDays, pendingDays }, existed ? 200 : 201);
}

function normalizeEmployees(items, route, baselineDate) {
  const map = new Map();
  for (const raw of Array.isArray(items) ? items : []) {
    const employeeId = normalizeId(raw.employeeId);
    const employeeName = clean(raw.employeeName, 100);
    if (!employeeId || !employeeName) continue;
    const cycleStart = normalizeDate(raw.cycleStart);
    const cycleEnd = normalizeDate(raw.cycleEnd);
    const underOneYear = Boolean(raw.underOneYear);
    const hireDate = normalizeDate(raw.hireDate);
    const basisHireDate = normalizeDate(raw.basisHireDate) || hireDate;
    const firstPromotionDate = normalizeDate(raw.firstPromotionDate) || (!underOneYear && cycleStart ? addMonths(cycleStart, 6) : "");
    const secondPromotionDate = normalizeDate(raw.secondPromotionDate) || (!underOneYear && cycleStart ? addMonths(cycleStart, 9) : "");
    const expiryDate = normalizeDate(raw.expiryDate) || cycleEnd;
    map.set(employeeId, {
      employeeId,
      employeeName,
      regionalManager: clean(raw.regionalManager, 100),
      manager: clean(raw.manager, 100),
      region1: clean(raw.region1, 100),
      region2: clean(raw.region2, 100),
      storeCode: clean(raw.storeCode, 50),
      storeName: clean(raw.storeName, 150),
      portalId: clean(raw.portalId, 80),
      hireDate,
      basisHireDate,
      policyType: route === "homeplus" ? "jan1" : "anniversary",
      underOneYear,
      baselineDate,
      baselineGranted: roundHalf(raw.baselineGranted),
      baselineUsed: roundHalf(raw.baselineUsed),
      baselineRemaining: roundHalf(raw.baselineRemaining),
      cycleStart,
      cycleEnd,
      firstPromotionDate,
      secondPromotionDate,
      expiryDate,
      terminationDate: normalizeDate(raw.terminationDate),
      note: clean(raw.note, 1000),
    });
  }
  return [...map.values()];
}

function normalizeApplications(items, route, month) {
  const rows = [];
  let index = 0;
  for (const raw of Array.isArray(items) ? items : []) {
    const employeeId = normalizeId(raw.employeeId);
    const leaveDate = normalizeDate(raw.leaveDate || raw.date);
    const days = roundHalf(raw.days ?? raw.requestedDays);
    if (!employeeId || !leaveDate || !leaveDate.startsWith(month) || !(days > 0)) continue;
    rows.push({
      route,
      month,
      employeeId,
      employeeName: clean(raw.employeeName || raw.name, 100),
      leaveDate,
      days,
      status: clean(raw.status || raw.applicationStatus, 50) || "대기",
      leaveType: clean(raw.leaveType || raw.requestedKind, 100) || "연차",
      applicationDate: normalizeDate(raw.applicationDate),
      note: clean(raw.note, 1000),
      sourceIndex: Number(raw.sourceIndex || ++index),
    });
  }
  return rows;
}

function calculateBalance(employee, applications, asOf) {
  if (!asOf) return emptyBalance(employee);
  const importedCycle = employee.cycleStart && employee.cycleEnd && asOf >= employee.cycleStart && asOf <= employee.cycleEnd;
  if (employee.underOneYear && asOf > employee.baselineDate && importedCycle) {
    const hireDate = employee.basisHireDate || employee.hireDate || employee.cycleStart;
    const accrued = Math.max(employee.baselineGranted, Math.min(11, countMonthlyAnniversaries(hireDate, asOf)));
    const afterBaseline = roundHalf(applications
      .filter((item) => item.leaveDate > employee.baselineDate && item.leaveDate <= asOf && isApprovedAnnual(item))
      .reduce((sum, item) => sum + item.days, 0));
    const extraAccrual = roundHalf(accrued - employee.baselineGranted);
    return {
      grantType: "월차",
      granted: accrued,
      approvedUsed: roundHalf(employee.baselineUsed + afterBaseline),
      remaining: roundHalf(employee.baselineRemaining + extraAccrual - afterBaseline),
      cycleStart: employee.cycleStart || hireDate,
      cycleEnd: employee.cycleEnd || addDays(addYears(hireDate, 1), -1),
      firstPromotionDate: "",
      secondPromotionDate: "",
      expiryDate: employee.cycleEnd || addDays(addYears(hireDate, 1), -1),
    };
  }
  if (importedCycle || asOf <= employee.baselineDate) {
    const afterBaseline = roundHalf(applications
      .filter((item) => item.leaveDate > employee.baselineDate && item.leaveDate <= asOf && isApprovedAnnual(item))
      .reduce((sum, item) => sum + item.days, 0));
    const used = roundHalf(employee.baselineUsed + afterBaseline);
    return {
      grantType: employee.underOneYear ? "월차" : "연차",
      granted: employee.baselineGranted,
      approvedUsed: used,
      remaining: roundHalf(employee.baselineRemaining - afterBaseline),
      cycleStart: employee.cycleStart,
      cycleEnd: employee.cycleEnd,
      firstPromotionDate: employee.firstPromotionDate,
      secondPromotionDate: employee.secondPromotionDate,
      expiryDate: employee.expiryDate || employee.cycleEnd,
    };
  }

  const cycle = resolveCycle(employee, asOf);
  if (!cycle.cycleStart) return emptyBalance(employee);
  const approvedUsed = roundHalf(applications
    .filter((item) => item.leaveDate >= cycle.cycleStart && item.leaveDate <= asOf && isApprovedAnnual(item))
    .reduce((sum, item) => sum + item.days, 0));
  return {
    ...cycle,
    approvedUsed,
    remaining: roundHalf(cycle.granted - approvedUsed),
  };
}

function resolveCycle(employee, asOf) {
  const hireDate = employee.basisHireDate || employee.hireDate;
  if (!hireDate) return emptyBalance(employee);
  const firstAnniversary = addYears(hireDate, 1);
  if (asOf < firstAnniversary) {
    const granted = Math.min(11, countMonthlyAnniversaries(hireDate, asOf));
    return {
      grantType: "월차",
      granted,
      cycleStart: hireDate,
      cycleEnd: addDays(firstAnniversary, -1),
      firstPromotionDate: "",
      secondPromotionDate: "",
      expiryDate: addDays(firstAnniversary, -1),
    };
  }

  if (employee.policyType === "jan1") {
    const nextJan = `${Number(firstAnniversary.slice(0, 4)) + 1}-01-01`;
    let cycleStart;
    let cycleEnd;
    if (asOf < nextJan) {
      cycleStart = firstAnniversary;
      cycleEnd = `${firstAnniversary.slice(0, 4)}-12-31`;
    } else {
      cycleStart = `${asOf.slice(0, 4)}-01-01`;
      cycleEnd = `${asOf.slice(0, 4)}-12-31`;
    }
    const years = completedYears(hireDate, cycleStart);
    return annualCycle(cycleStart, cycleEnd, annualGrant(years));
  }

  const years = Math.max(1, completedYears(hireDate, asOf));
  const cycleStart = addYears(hireDate, years);
  const cycleEnd = addDays(addYears(cycleStart, 1), -1);
  return annualCycle(cycleStart, cycleEnd, annualGrant(years));
}

function annualCycle(cycleStart, cycleEnd, granted) {
  return {
    grantType: "연차",
    granted,
    cycleStart,
    cycleEnd,
    firstPromotionDate: addMonths(cycleStart, 6),
    secondPromotionDate: addMonths(cycleStart, 9),
    expiryDate: cycleEnd,
  };
}

function emptyBalance(employee) {
  return {
    grantType: employee?.underOneYear ? "월차" : "연차",
    granted: 0,
    approvedUsed: 0,
    remaining: 0,
    cycleStart: employee?.cycleStart || "",
    cycleEnd: employee?.cycleEnd || "",
    firstPromotionDate: employee?.firstPromotionDate || "",
    secondPromotionDate: employee?.secondPromotionDate || "",
    expiryDate: employee?.expiryDate || "",
  };
}

function toClientEmployee(row) {
  return {
    route: row.route,
    employeeId: row.employee_id || "",
    employeeName: row.employee_name || "",
    regionalManager: row.regional_manager || "",
    manager: row.manager || "",
    region1: row.region1 || "",
    region2: row.region2 || "",
    storeCode: row.store_code || "",
    storeName: row.store_name || "",
    portalId: row.portal_id || "",
    hireDate: row.hire_date || "",
    basisHireDate: row.basis_hire_date || "",
    policyType: row.policy_type || (row.route === "homeplus" ? "jan1" : "anniversary"),
    underOneYear: Boolean(row.under_one_year),
    baselineDate: row.baseline_date || "",
    baselineGranted: roundHalf(row.baseline_granted),
    baselineUsed: roundHalf(row.baseline_used),
    baselineRemaining: roundHalf(row.baseline_remaining),
    cycleStart: row.cycle_start || "",
    cycleEnd: row.cycle_end || "",
    firstPromotionDate: row.first_promotion_date || "",
    secondPromotionDate: row.second_promotion_date || "",
    expiryDate: row.expiry_date || "",
    terminationDate: row.termination_date || "",
    note: row.note || "",
  };
}

function toClientApplication(row) {
  return {
    route: row.route,
    month: row.month,
    employeeId: row.employee_id || "",
    employeeName: row.employee_name || "",
    leaveDate: row.leave_date || "",
    days: roundHalf(row.days),
    status: row.status || "",
    leaveType: row.leave_type || "",
    applicationDate: row.application_date || "",
    note: row.note || "",
    sourceIndex: Number(row.source_index || 0),
  };
}

function isAnnualType(item) {
  const type = String(item?.leaveType || "").replace(/\s+/g, "");
  return type.includes("연차") || type.includes("반차");
}

function isApprovedAnnual(item) {
  const status = String(item?.status || "").replace(/\s+/g, "");
  return isAnnualType(item) && (APPROVED_STATUS.has(status) || status.startsWith("승인"));
}

function annualGrant(completedYearCount) {
  const years = Math.max(1, Number(completedYearCount) || 1);
  return Math.min(25, 15 + Math.floor((years - 1) / 2));
}

function completedYears(start, end) {
  const a = parseISO(start);
  const b = parseISO(end);
  if (!a || !b || b < a) return 0;
  let years = b.getUTCFullYear() - a.getUTCFullYear();
  const anniversary = new Date(Date.UTC(b.getUTCFullYear(), a.getUTCMonth(), a.getUTCDate()));
  if (b < anniversary) years -= 1;
  return Math.max(0, years);
}

function countMonthlyAnniversaries(start, end) {
  let cursor = addMonths(start, 1);
  let count = 0;
  while (cursor && cursor <= end && count < 11) {
    count += 1;
    cursor = addMonths(cursor, 1);
  }
  return count;
}

function addYears(dateText, years) {
  const date = parseISO(dateText);
  if (!date) return "";
  const month = date.getUTCMonth();
  date.setUTCFullYear(date.getUTCFullYear() + Number(years || 0));
  if (date.getUTCMonth() !== month) date.setUTCDate(0);
  return formatISO(date);
}

function addMonths(dateText, months) {
  const date = parseISO(dateText);
  if (!date) return "";
  const originalDay = date.getUTCDate();
  date.setUTCDate(1);
  date.setUTCMonth(date.getUTCMonth() + Number(months || 0));
  const lastDay = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate();
  date.setUTCDate(Math.min(originalDay, lastDay));
  return formatISO(date);
}

function addDays(dateText, days) {
  const date = parseISO(dateText);
  if (!date) return "";
  date.setUTCDate(date.getUTCDate() + Number(days || 0));
  return formatISO(date);
}

function startOfMonth(dateText) {
  return `${dateText.slice(0, 7)}-01`;
}

function endOfMonth(monthText) {
  const [year, month] = monthText.split("-").map(Number);
  return `${monthText}-${String(new Date(Date.UTC(year, month, 0)).getUTCDate()).padStart(2, "0")}`;
}

function parseISO(value) {
  const normalized = normalizeDate(value);
  if (!normalized) return null;
  return new Date(`${normalized}T00:00:00Z`);
}

function formatISO(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function normalizeDate(value) {
  const text = String(value || "").trim();
  const match = text.match(/^(20\d{2})[-./]?(\d{1,2})[-./]?(\d{1,2})/);
  if (!match) return "";
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return "";
  return formatISO(date);
}

function normalizeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}

function clean(value, max = 500) {
  return String(value ?? "").trim().slice(0, max);
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

async function runBatch(db, statements, size = 50) {
  for (let index = 0; index < statements.length; index += size) {
    const chunk = statements.slice(index, index + size);
    if (chunk.length) await db.batch(chunk);
  }
}
