const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const ROUTE_LABELS = { homeplus: "홈플러스", electroland: "전자랜드" };
const WEEKDAY_LABELS = ["일", "월", "화", "수", "목", "금", "토"];
const REQUIRED_CLOCK_PLANS = new Set(["근무", "근무A", "근무B", "근무C", "교육", "오전반차", "오후반차", "공백"]);
const UNEXPECTED_CLOCK_PLANS = new Set([
  "휴무", "무급휴가", "연차", "공가", "휴가", "경조",
  "대체휴일(1일)", "대체휴일(0.5일)", "보상휴가(1일)", "보상휴가(0.5일)",
]);
const VALID_PLAN_CODES = new Set([
  "무급휴가", "오전반차", "오후반차", "교육", "연차", "공가", "휴가", "기타",
  "근무", "경조", "근무A", "근무B", "근무C", "대체휴일(1일)", "대체휴일(0.5일)",
  "보상휴가(0.5일)", "보상휴가(1일)", "휴무", "공백",
]);

const state = {
  planFile: null,
  attendanceFile: null,
  result: null,
  priorLedger: emptyLedger(),
  backend: { available: false, configured: false, loggedIn: false },
  activeResultTab: "missing",
  currentPage: 1,
  pageSize: 50,
  grants: [],
};

init();

async function init() {
  setDefaultDates();
  bindEvents();
  setupDropzone("planDropzone", "planFile", setPlanFile);
  setupDropzone("attendanceDropzone", "attendanceFile", setAttendanceFile);
  await checkBackend();
  syncRouteRuleHelp();
  if (state.backend.loggedIn) await Promise.all([loadHistory(), loadGrants()]);
}

function bindEvents() {
  $("#analyzeButton").addEventListener("click", analyzeFiles);
  $("#resetButton").addEventListener("click", resetAll);
  $("#searchInput").addEventListener("input", () => { state.currentPage = 1; renderActiveTable(); });
  $("#storeFilter").addEventListener("change", () => { state.currentPage = 1; renderActiveTable(); });
  $("#exportButton").addEventListener("click", exportResults);
  $("#saveClosureButton").addEventListener("click", saveClosure);
  $("#refreshHistoryButton").addEventListener("click", loadHistory);
  $("#refreshGrantButton").addEventListener("click", loadGrants);
  $("#grantRouteFilter").addEventListener("change", loadGrants);
  $("#grantForm").addEventListener("submit", saveGrant);
  $("#grantMonth").addEventListener("change", syncGrantDates);
  $("#grantCancelEdit").addEventListener("click", resetGrantForm);
  $("#loginButton").addEventListener("click", openLogin);
  $("#logoutButton").addEventListener("click", logout);
  $("#loginCancel").addEventListener("click", () => $("#loginDialog").close());
  $("#loginForm").addEventListener("submit", login);
  $("#targetMonth").addEventListener("change", () => { syncCutoffWithMonth(); syncRouteRuleHelp(); });
  $$('input[name="route"]').forEach((input) => input.addEventListener("change", syncRouteRuleHelp));
  $$(".tab[data-view]").forEach((tab) => tab.addEventListener("click", () => switchView(tab.dataset.view)));
  $$(".inner-tab").forEach((tab) => tab.addEventListener("click", () => switchResultTab(tab.dataset.resultTab)));
}

function setDefaultDates() {
  const now = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  $("#targetMonth").value = month;
  $("#cutoffDate").value = toISODate(now);
  $("#grantMonth").value = month;
  syncGrantDates();
}

function syncCutoffWithMonth() {
  const month = $("#targetMonth").value;
  if (!month) return;
  const [year, monthNumber] = month.split("-").map(Number);
  const now = new Date();
  const isCurrent = now.getFullYear() === year && now.getMonth() + 1 === monthNumber;
  $("#cutoffDate").value = toISODate(isCurrent ? now : new Date(year, monthNumber, 0));
}

function syncGrantDates() {
  const month = $("#grantMonth").value;
  if (!month) return;
  const [year, monthNumber] = month.split("-").map(Number);
  $("#grantValidFrom").value = `${month}-01`;
  $("#grantValidTo").value = toISODate(new Date(year, monthNumber + 1, 0));
}

function syncRouteRuleHelp() {
  const route = selectedRoute();
  const month = $("#targetMonth").value;
  if (route === "homeplus") {
    $("#routeRuleHelp").textContent = "홈플러스 기본 휴무는 월 6일입니다.";
  } else {
    const count = month ? countWeekendDays(month) : 0;
    $("#routeRuleHelp").textContent = `전자랜드 기본 휴무는 대상 월 토요일+일요일 ${count}일입니다.`;
  }
}

function setupDropzone(zoneId, inputId, setter) {
  const zone = $(`#${zoneId}`);
  const input = $(`#${inputId}`);
  input.addEventListener("change", () => setter(input.files?.[0] || null));
  ["dragenter", "dragover"].forEach((name) => zone.addEventListener(name, (event) => {
    event.preventDefault();
    zone.classList.add("dragover");
  }));
  ["dragleave", "drop"].forEach((name) => zone.addEventListener(name, (event) => {
    event.preventDefault();
    zone.classList.remove("dragover");
  }));
  zone.addEventListener("drop", (event) => setter(event.dataTransfer.files?.[0] || null));
}

function setPlanFile(file) {
  state.planFile = file;
  $("#planFileName").textContent = file ? file.name : "계획표를 선택하거나 끌어놓기";
}

function setAttendanceFile(file) {
  state.attendanceFile = file;
  $("#attendanceFileName").textContent = file ? file.name : "근태표를 선택하거나 끌어놓기";
}

async function analyzeFiles() {
  const button = $("#analyzeButton");
  try {
    if (!window.XLSX) throw new Error("엑셀 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.");
    if (!state.planFile || !state.attendanceFile) throw new Error("계획표와 실제 근태표를 모두 선택해 주세요.");

    const targetMonth = $("#targetMonth").value;
    const cutoffDate = $("#cutoffDate").value;
    const route = selectedRoute();
    if (!targetMonth || !cutoffDate) throw new Error("대상 월과 비교 기준일을 입력해 주세요.");
    if (!cutoffDate.startsWith(targetMonth)) throw new Error("비교 기준일은 대상 월 안의 날짜여야 합니다.");

    button.disabled = true;
    button.textContent = "파일 분석 중...";

    const [planMatrix, attendanceMatrix] = await Promise.all([
      fileToMatrix(state.planFile),
      fileToMatrix(state.attendanceFile),
    ]);
    const plan = parsePlan(planMatrix);
    const attendance = parseAttendance(attendanceMatrix, targetMonth);
    state.priorLedger = await loadPriorLedger(route, targetMonth);

    const result = compareAttendance({ plan, attendance, route, targetMonth, cutoffDate, ledger: state.priorLedger });
    state.result = {
      ...result,
      route,
      routeLabel: ROUTE_LABELS[route],
      targetMonth,
      cutoffDate,
      planFileName: state.planFile.name,
      attendanceFileName: state.attendanceFile.name,
      analyzedAt: new Date().toISOString(),
    };

    renderResult();
    showToast(`분석 완료: 출근기록 없음 ${result.missingRows.length}건 · 휴무·휴가 출근 ${result.unexpectedRows.length}건 · 불일치 ${result.mismatchRows.length}건`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "분석 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "근태 종합 분석";
  }
}

async function loadPriorLedger(route, month) {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) return emptyLedger();
  try {
    const response = await fetch(`/api/substitute-balances?route=${encodeURIComponent(route)}&month=${encodeURIComponent(month)}`, { cache: "no-store" });
    if (response.status === 401) {
      await checkBackend();
      return emptyLedger();
    }
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "누적 자료 조회 실패");
    return {
      lotsByEmployee: data.lotsByEmployee || {},
      annualLeaveBefore: data.annualLeaveBefore || {},
      currentGrant: data.currentGrant || null,
    };
  } catch (error) {
    showToast(`${error.message}. 이번 분석은 저장된 이전 월 누적을 제외하고 계산합니다.`);
    return emptyLedger();
  }
}

function emptyLedger() {
  return { lotsByEmployee: {}, annualLeaveBefore: {}, currentGrant: null };
}

async function fileToMatrix(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, raw: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false, blankrows: false });
}

function parsePlan(matrix) {
  const headerIndex = findHeaderRow(matrix, ["사번", "매장명", "이름"]);
  if (headerIndex < 0) throw new Error("계획표에서 ‘매장명·사번·이름’ 머리글을 찾지 못했습니다.");

  const headers = matrix[headerIndex].map(normalizeHeader);
  const columns = {
    store: findHeaderIndex(headers, ["매장명", "매장", "지점명"]),
    employeeId: findHeaderIndex(headers, ["사번", "사원번호"]),
    name: findHeaderIndex(headers, ["이름", "성명"]),
    employment: findHeaderIndex(headers, ["재직상태", "근무상태"]),
  };
  const dayColumns = new Map();
  headers.forEach((header, index) => {
    const match = header.match(/^(\d{1,2})일$/);
    if (match) dayColumns.set(Number(match[1]), index);
  });
  if (!dayColumns.size) throw new Error("계획표에서 ‘01일~31일’ 날짜 열을 찾지 못했습니다.");

  const rows = matrix.slice(headerIndex + 1).map((row, sourceIndex) => ({
    store: text(row[columns.store]),
    employeeId: normalizeEmployeeId(row[columns.employeeId]),
    name: text(row[columns.name]),
    employment: columns.employment >= 0 ? text(row[columns.employment]) : "",
    plans: Object.fromEntries([...dayColumns.entries()].map(([day, col]) => [day, text(row[col])])),
    sourceIndex: sourceIndex + headerIndex + 2,
  })).filter((row) => row.employeeId && row.name && (!row.employment || !row.employment.includes("퇴사")));

  if (!rows.length) throw new Error("계획표에서 유효한 사번 데이터를 찾지 못했습니다.");
  return { rows, headerIndex, dayColumns, detectedRoute: detectRouteFromPlan(rows) };
}

function parseAttendance(matrix, targetMonth) {
  const headerIndex = findHeaderRow(matrix, ["사번", "근무일자"]);
  if (headerIndex < 0) throw new Error("근태표에서 ‘사번·근무일자’ 머리글을 찾지 못했습니다.");

  const headers = matrix[headerIndex].map(normalizeHeader);
  const columns = {
    employeeId: findHeaderIndex(headers, ["사번", "사원번호"]),
    name: findHeaderIndex(headers, ["이름", "성명"]),
    date: findHeaderIndex(headers, ["근무일자", "근무일", "일자"]),
    actualIn: findHeaderIndex(headers, ["(실제)출근시간", "실제출근시간", "출근시간"]),
    changedIn: findHeaderIndex(headers, ["(변경)출근시간", "변경출근시간", "수정출근시간"]),
    location: findHeaderIndex(headers, ["출근지점", "출근매장", "근무지점", "매장명"]),
    actualStatus: findHeaderIndex(headers, [
      "실제근태", "근태", "근태구분", "근태항목", "근태명", "근태코드",
      "근무구분", "실제근무구분", "(실제)근태", "처리근태",
    ]),
  };
  if (columns.date < 0 || columns.employeeId < 0) throw new Error("근태표의 날짜 또는 사번 열을 확인해 주세요.");
  if (columns.actualIn < 0 && columns.changedIn < 0 && columns.actualStatus < 0) {
    throw new Error("근태표에서 실제/변경 출근시간 또는 실제 근태 열을 찾지 못했습니다.");
  }

  const rows = matrix.slice(headerIndex + 1).map((row, sourceIndex) => {
    const date = parseDateCell(row[columns.date]);
    return {
      employeeId: normalizeEmployeeId(row[columns.employeeId]),
      name: columns.name >= 0 ? text(row[columns.name]) : "",
      date,
      actualIn: columns.actualIn >= 0 ? text(row[columns.actualIn]) : "",
      changedIn: columns.changedIn >= 0 ? text(row[columns.changedIn]) : "",
      location: columns.location >= 0 ? text(row[columns.location]) : "",
      actualStatus: columns.actualStatus >= 0 ? text(row[columns.actualStatus]) : "",
      sourceIndex: sourceIndex + headerIndex + 2,
    };
  }).filter((row) => row.employeeId && row.date && row.date.startsWith(targetMonth));

  if (!rows.length) throw new Error(`${targetMonth} 근태 기록을 찾지 못했습니다. 대상 월을 확인해 주세요.`);
  return {
    rows,
    headerIndex,
    detectedRoute: detectRouteFromAttendance(rows),
    hasActualStatusColumn: columns.actualStatus >= 0,
  };
}

function compareAttendance({ plan, attendance, route, targetMonth, cutoffDate, ledger }) {
  const attendanceIds = new Set(attendance.rows.map((row) => row.employeeId));
  const planIds = new Set(plan.rows.map((row) => row.employeeId));
  const matchedIds = [...planIds].filter((id) => attendanceIds.has(id));
  const matchRate = planIds.size ? Math.round((matchedIds.length / planIds.size) * 1000) / 10 : 0;
  const attendanceById = groupBy(attendance.rows, (row) => row.employeeId);
  const selectedPlans = choosePlanRows(plan.rows, attendanceById);
  const attendanceMap = buildAttendanceMap(attendance.rows);
  const [year, monthNumber] = targetMonth.split("-").map(Number);
  const daysInMonth = new Date(year, monthNumber, 0).getDate();
  const lastDay = Math.min(daysInMonth, Number(cutoffDate.slice(-2)));
  const baseAllowance = route === "homeplus" ? 6 : countWeekendDays(targetMonth);

  const missingRows = [];
  const unexpectedRows = [];
  const mismatchRows = [];
  const employeeSummaries = [];
  const employeeFacts = [];

  for (const person of selectedPlans) {
    const basicDayoffDates = [];
    const explicitSubstituteEvents = [];
    const annualLeaveEvents = [];

    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const planCode = normalizePlanCode(person.plans[day]);
      if (planCode === "휴무") basicDayoffDates.push(date);
      const substituteDays = substitutePlanValue(planCode);
      if (substituteDays > 0) {
        explicitSubstituteEvents.push({ date, days: substituteDays, source: "표기 대체휴무", planStatus: planCode });
      }
      const annualDays = annualLeaveValue(planCode);
      if (annualDays > 0) annualLeaveEvents.push({ date, days: annualDays, planStatus: planCode });
    }

    const baseExcessEvents = basicDayoffDates.slice(baseAllowance).map((date) => ({
      date,
      days: 1,
      source: "기본 휴무 초과",
      planStatus: "휴무",
    }));
    const substituteEvents = [...explicitSubstituteEvents, ...baseExcessEvents]
      .sort((a, b) => a.date.localeCompare(b.date));

    for (let day = 1; day <= lastDay; day += 1) {
      const rawPlanStatus = person.plans[day];
      const planStatus = normalizePlanCode(rawPlanStatus);
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const dateObject = new Date(`${date}T00:00:00`);
      const attendanceValue = attendanceMap.get(`${person.employeeId}|${date}`) || emptyAttendanceValue();

      if (REQUIRED_CLOCK_PLANS.has(planStatus) && !attendanceValue.hasClockIn) {
        missingRows.push(makeIssueRow({
          issueType: "missing_clock_in",
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: "근무인데 출근기록 없음",
          reason: `${planStatus} 계획이나 실제 출근시간과 변경 출근시간이 모두 없음`,
        }));
      }

      if (UNEXPECTED_CLOCK_PLANS.has(planStatus) && attendanceValue.hasClockIn) {
        unexpectedRows.push(makeIssueRow({
          issueType: "unexpected_clock_in",
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: "휴무·휴가인데 출근기록 있음",
          reason: `${planStatus} 계획이나 출근기록이 있음`,
        }));
      }

      const mismatch = evaluateMismatch(planStatus, attendanceValue.actualStatus, attendance.hasActualStatusColumn);
      if (mismatch) {
        mismatchRows.push(makeMismatchRow({
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: mismatch.result,
          reason: mismatch.reason,
        }));
      }
    }

    const basicDayoffUsed = roundHalf(basicDayoffDates.length);
    const explicitSubDayoffUsed = roundHalf(explicitSubstituteEvents.reduce((sum, event) => sum + event.days, 0));
    const baseExcess = roundHalf(Math.max(0, basicDayoffUsed - baseAllowance));
    const substituteNeeded = roundHalf(baseExcess + explicitSubDayoffUsed);
    const annualLeaveUsed = roundHalf(annualLeaveEvents.reduce((sum, event) => sum + event.days, 0));
    const preview = calculatePreviewLedger({
      employeeId: person.employeeId,
      substituteEvents,
      targetMonth,
      ledger,
    });
    const cumulativeAnnualLeave = roundHalf(Number(ledger.annualLeaveBefore[person.employeeId] || 0) + annualLeaveUsed);
    const judgment = buildDayoffJudgment({
      baseAllowance,
      basicDayoffUsed,
      explicitSubDayoffUsed,
      baseExcess,
      substituteNeeded,
      ...preview,
    });

    const fact = {
      route,
      routeLabel: ROUTE_LABELS[route],
      store: person.store,
      employeeId: person.employeeId,
      name: person.name,
      baseAllowance,
      basicDayoffUsed,
      explicitSubDayoffUsed,
      baseExcess,
      substituteNeeded,
      annualLeaveUsed,
      substituteEvents,
      annualLeaveEvents,
      duplicatePlanNote: person.duplicatePlanNote || "",
    };
    employeeFacts.push(fact);
    employeeSummaries.push({
      ...fact,
      availableSubstitute: preview.availableSubstitute,
      substituteApplied: preview.substituteApplied,
      remainingSubstitute: preview.remainingSubstitute,
      expiredSubstitute: preview.expiredSubstitute,
      shortage: preview.shortage,
      currentAnnualLeave: annualLeaveUsed,
      cumulativeAnnualLeave,
      judgment,
    });
  }

  missingRows.sort(issueSort);
  unexpectedRows.sort(issueSort);
  mismatchRows.sort(issueSort);
  sortSummaries(employeeSummaries);

  const diagnostics = buildDiagnostics({
    route,
    plan,
    attendance,
    matchRate,
    planIds,
    attendanceIds,
    baseAllowance,
  });

  return assembleResultCollections({
    missingRows,
    unexpectedRows,
    mismatchRows,
    employeeSummaries,
    employeeFacts,
    planPeople: planIds.size,
    attendancePeople: attendanceIds.size,
    matchedPeople: matchedIds.length,
    matchRate,
    baseAllowance,
    diagnostics,
  });
}

function assembleResultCollections(base) {
  return {
    ...base,
    missingPeople: uniquePeople(base.missingRows),
    unexpectedPeople: uniquePeople(base.unexpectedRows),
    mismatchPeople: uniquePeople(base.mismatchRows),
    dayoffExcessRows: base.employeeSummaries.filter((row) => row.baseExcess > 0),
    balanceRows: [...base.employeeSummaries],
    shortageRows: base.employeeSummaries.filter((row) => row.shortage > 0),
    annualRows: base.employeeSummaries.filter((row) => row.currentAnnualLeave > 0 || row.cumulativeAnnualLeave > 0),
    dayoffExcessPeople: base.employeeSummaries.filter((row) => row.baseExcess > 0).length,
    substituteShortagePeople: base.employeeSummaries.filter((row) => row.shortage > 0).length,
    annualLeavePeople: base.employeeSummaries.filter((row) => row.currentAnnualLeave > 0).length,
  };
}

function calculatePreviewLedger({ employeeId, substituteEvents, targetMonth, ledger }) {
  const lots = (ledger.lotsByEmployee[employeeId] || []).map((lot) => ({
    grantId: lot.grantId,
    grantMonth: lot.grantMonth,
    validFrom: lot.validFrom,
    validTo: lot.validTo,
    remaining: roundHalf(lot.remaining),
  }));

  const currentGrant = ledger.currentGrant;
  if (currentGrant && currentGrant.grant_month === targetMonth) {
    lots.push({
      grantId: currentGrant.id,
      grantMonth: currentGrant.grant_month,
      validFrom: currentGrant.valid_from,
      validTo: currentGrant.valid_to,
      remaining: roundHalf(currentGrant.granted_days),
    });
  }

  const monthStart = `${targetMonth}-01`;
  const monthEnd = endOfMonth(targetMonth);
  const nextMonthStart = startOfNextMonth(targetMonth);
  const availableSubstitute = roundHalf(lots
    .filter((lot) => lot.remaining > 0 && lot.validFrom <= monthEnd && lot.validTo >= monthStart)
    .reduce((sum, lot) => sum + lot.remaining, 0));

  let substituteApplied = 0;
  let shortage = 0;
  for (const event of [...substituteEvents].sort((a, b) => a.date.localeCompare(b.date))) {
    let need = roundHalf(event.days);
    const candidates = lots
      .filter((lot) => lot.remaining > 0 && lot.validFrom <= event.date && lot.validTo >= event.date)
      .sort((a, b) => a.validTo.localeCompare(b.validTo) || a.validFrom.localeCompare(b.validFrom) || a.grantMonth.localeCompare(b.grantMonth));
    for (const lot of candidates) {
      if (need <= 0) break;
      const used = roundHalf(Math.min(need, lot.remaining));
      lot.remaining = roundHalf(lot.remaining - used);
      need = roundHalf(need - used);
      substituteApplied = roundHalf(substituteApplied + used);
    }
    shortage = roundHalf(shortage + Math.max(0, need));
  }

  const remainingSubstitute = roundHalf(lots
    .filter((lot) => lot.remaining > 0 && lot.validTo >= nextMonthStart)
    .reduce((sum, lot) => sum + lot.remaining, 0));
  const expiredSubstitute = roundHalf(lots
    .filter((lot) => lot.remaining > 0 && lot.validTo >= monthStart && lot.validTo < nextMonthStart)
    .reduce((sum, lot) => sum + lot.remaining, 0));

  return { availableSubstitute, substituteApplied, remainingSubstitute, expiredSubstitute, shortage };
}

function buildAttendanceMap(rows) {
  const map = new Map();
  for (const row of rows) {
    const key = `${row.employeeId}|${row.date}`;
    const existing = map.get(key) || emptyAttendanceValue();
    map.set(key, {
      hasClockIn: existing.hasClockIn || Boolean(row.actualIn || row.changedIn),
      actualIn: row.actualIn || existing.actualIn,
      changedIn: row.changedIn || existing.changedIn,
      location: row.location || existing.location,
      actualStatus: row.actualStatus || existing.actualStatus,
    });
  }
  return map;
}

function emptyAttendanceValue() {
  return { hasClockIn: false, actualIn: "", changedIn: "", location: "", actualStatus: "" };
}

function makeIssueRow({ issueType, route, person, date, dateObject, planStatus, attendanceValue, result, reason }) {
  return {
    issueType,
    route,
    routeLabel: ROUTE_LABELS[route],
    store: person.store,
    employeeId: person.employeeId,
    name: person.name,
    date,
    weekday: WEEKDAY_LABELS[dateObject.getDay()],
    planStatus,
    actualStatus: attendanceValue.actualStatus || "",
    actualIn: attendanceValue.actualIn || "",
    changedIn: attendanceValue.changedIn || "",
    clockStatus: attendanceValue.hasClockIn ? (attendanceValue.changedIn || attendanceValue.actualIn) : "미기록",
    result,
    reason,
    duplicatePlanNote: person.duplicatePlanNote || "",
  };
}

function makeMismatchRow({ route, person, date, dateObject, planStatus, attendanceValue, result, reason }) {
  return {
    route,
    routeLabel: ROUTE_LABELS[route],
    store: person.store,
    employeeId: person.employeeId,
    name: person.name,
    date,
    weekday: WEEKDAY_LABELS[dateObject.getDay()],
    planStatus,
    actualStatus: attendanceValue.actualStatus || "미기재",
    actualIn: attendanceValue.actualIn || "",
    changedIn: attendanceValue.changedIn || "",
    clockStatus: attendanceValue.hasClockIn ? (attendanceValue.changedIn || attendanceValue.actualIn) : "미기록",
    result,
    reason,
    duplicatePlanNote: person.duplicatePlanNote || "",
  };
}

function evaluateMismatch(planStatus, actualStatusRaw, hasActualStatusColumn) {
  if (planStatus === "기타") {
    return { result: "검토 필요", reason: "기타는 자동 단정하지 않고 계획·실제 근태 확인이 필요함" };
  }
  if (!VALID_PLAN_CODES.has(planStatus)) {
    return { result: "검토 필요", reason: `등록되지 않은 계획 코드 ‘${planStatus}’ 확인 필요` };
  }
  if (!hasActualStatusColumn || !text(actualStatusRaw)) return null;

  const actualStatus = normalizeActualCode(actualStatusRaw);
  if (!actualStatus) return null;
  const planComparable = comparableCode(planStatus);
  const actualComparable = comparableCode(actualStatus);
  if (planComparable === actualComparable) return null;

  return {
    result: "계획·실제 근태 불일치",
    reason: `계획 ${planStatus} / 실제 ${actualStatusRaw}`,
  };
}

function choosePlanRows(planRows, attendanceById) {
  const plansById = groupBy(planRows, (row) => row.employeeId);
  return [...plansById.entries()].map(([employeeId, rows]) => {
    if (rows.length === 1) return rows[0];
    const locations = (attendanceById.get(employeeId) || []).map((row) => normalizeStore(row.location));
    const scored = rows.map((row) => ({
      row,
      score: locations.filter((location) => {
        const store = normalizeStore(row.store);
        return store && location && (location.includes(store) || store.includes(location));
      }).length,
    })).sort((a, b) => b.score - a.score);
    return { ...scored[0].row, duplicatePlanNote: `중복 계획 ${rows.length}건 중 출근지점 매칭으로 선택` };
  });
}

function normalizePlanCode(value) {
  const status = normalizeStatus(value);
  return status || "공백";
}

function normalizeActualCode(value) {
  const raw = normalizeStatus(value);
  if (!raw || raw === "공백") return "";
  if (raw.includes("오전반차") || raw.includes("반차(오전)")) return "오전반차";
  if (raw.includes("오후반차") || raw.includes("반차(오후)")) return "오후반차";
  if (raw.includes("대체") && (raw.includes("0.5") || raw.includes("반일"))) return "대체휴일(0.5일)";
  if (raw.includes("대체")) return "대체휴일(1일)";
  if (raw.includes("보상") && (raw.includes("0.5") || raw.includes("반일"))) return "보상휴가(0.5일)";
  if (raw.includes("보상")) return "보상휴가(1일)";
  if (raw.includes("무급")) return "무급휴가";
  if (raw.includes("연차")) return "연차";
  if (raw.includes("공가")) return "공가";
  if (raw.includes("경조")) return "경조";
  if (raw.includes("교육")) return "교육";
  if (raw === "휴무" || raw.includes("정기휴무") || raw.includes("월휴무")) return "휴무";
  if (raw.includes("휴가")) return "휴가";
  if (raw.includes("근무") || raw.includes("출근") || raw === "정상") return "근무";
  return raw;
}

function comparableCode(code) {
  if (["공백", "근무", "근무A", "근무B", "근무C"].includes(code)) return "근무";
  return code;
}

function substitutePlanValue(planCode) {
  if (["대체휴일(0.5일)", "보상휴가(0.5일)"].includes(planCode)) return 0.5;
  if (["대체휴일(1일)", "보상휴가(1일)"].includes(planCode)) return 1;
  return 0;
}

function annualLeaveValue(planCode) {
  if (planCode === "연차") return 1;
  if (["오전반차", "오후반차"].includes(planCode)) return 0.5;
  return 0;
}

function buildDayoffJudgment({ baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess, substituteNeeded, remainingSubstitute, expiredSubstitute, shortage }) {
  if (shortage > 0) return `대체휴무 ${formatDays(shortage)} 초과 사용`;
  if (baseExcess > 0) {
    const explicit = explicitSubDayoffUsed > 0 ? ` · 표기 대체휴무 ${formatDays(explicitSubDayoffUsed)} 사용` : "";
    const expired = expiredSubstitute > 0 ? ` · ${formatDays(expiredSubstitute)} 만료` : "";
    return `휴무 개수 초과 ${formatDays(baseExcess)} · 대체휴무 여분 활용${explicit} · 잔여 ${formatDays(remainingSubstitute)}${expired}`;
  }
  if (substituteNeeded > 0) {
    const expired = expiredSubstitute > 0 ? ` · ${formatDays(expiredSubstitute)} 만료` : "";
    return `대체휴무 ${formatDays(substituteNeeded)} 사용 · 잔여 ${formatDays(remainingSubstitute)}${expired}`;
  }
  if (expiredSubstitute > 0) return `기본 휴무 정상 · 미사용 대체휴무 ${formatDays(expiredSubstitute)} 만료`;
  return `정상 · 기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 잔여 ${formatDays(remainingSubstitute)}`;
}

function buildDiagnostics({ route, plan, attendance, matchRate, planIds, attendanceIds, baseAllowance }) {
  const messages = [`${ROUTE_LABELS[route]} 경로 기본 휴무 기준: ${baseAllowance}일`];
  if (plan.detectedRoute && plan.detectedRoute !== route) messages.push(`계획표 내용은 ‘${ROUTE_LABELS[plan.detectedRoute]} 경로’로 감지되었습니다.`);
  if (attendance.detectedRoute && attendance.detectedRoute !== route) messages.push(`근태표 내용은 ‘${ROUTE_LABELS[attendance.detectedRoute]} 경로’로 감지되었습니다.`);
  if (matchRate < 50) messages.push(`사번 매칭률이 ${matchRate}%로 낮습니다. 서로 다른 경로 또는 월 파일인지 확인해 주세요.`);
  if (planIds.size && attendanceIds.size && matchRate >= 50) messages.push(`사번 ${[...planIds].filter((id) => attendanceIds.has(id)).length}명이 정상 매칭되었습니다.`);
  if (!attendance.hasActualStatusColumn) messages.push("실제 근태 열을 찾지 못해 계획·실제 불일치는 ‘기타/미등록 계획 코드’만 표시합니다.");
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) messages.push("관리자 로그인 전에는 이전 월 대체휴무 잔여와 누적 연차를 제외하고 계산합니다.");
  return messages;
}

function renderResult() {
  const result = state.result;
  $("#emptyState").classList.add("hidden");
  $("#resultArea").classList.remove("hidden");
  refreshResultMetrics();
  $("#resultDescription").textContent = `${result.routeLabel} 경로 · ${result.targetMonth} · ${result.cutoffDate}까지 출근 판정 · 기본 휴무 ${result.baseAllowance}일`;

  const banner = $("#diagnosticBanner");
  const warning = result.matchRate < 50 || result.diagnostics.some((message) => message.includes("감지") || message.includes("찾지 못해"));
  banner.className = `alert ${warning ? "warning" : "success"}`;
  banner.innerHTML = result.diagnostics.map(escapeHtml).join("<br>");
  banner.classList.remove("hidden");

  const stores = [...new Set(result.employeeSummaries.map((row) => row.store).filter(Boolean))].sort();
  $("#storeFilter").innerHTML = `<option value="">전체 매장</option>${stores.map((store) => `<option>${escapeHtml(store)}</option>`).join("")}`;
  $("#searchInput").value = "";
  switchResultTab("missing");
}

function refreshResultMetrics() {
  const result = state.result;
  $("#missingCount").textContent = number(result.missingRows.length);
  $("#unexpectedCount").textContent = number(result.unexpectedRows.length);
  $("#mismatchCount").textContent = number(result.mismatchRows.length);
  $("#dayoffExcessPeople").textContent = number(result.dayoffExcessPeople);
  $("#substituteShortagePeople").textContent = number(result.substituteShortagePeople);
  $("#annualLeavePeople").textContent = number(result.annualLeavePeople);
  $("#matchRate").textContent = `${result.matchRate}%`;
  $("#missingTabCount").textContent = result.missingRows.length;
  $("#unexpectedTabCount").textContent = result.unexpectedRows.length;
  $("#mismatchTabCount").textContent = result.mismatchRows.length;
  $("#excessTabCount").textContent = result.dayoffExcessRows.length;
  $("#balanceTabCount").textContent = result.balanceRows.length;
  $("#shortageTabCount").textContent = result.shortageRows.length;
  $("#annualTabCount").textContent = result.annualRows.length;
}

function switchResultTab(tabName) {
  state.activeResultTab = tabName;
  state.currentPage = 1;
  $$(".inner-tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.resultTab === tabName));
  renderActiveTable();
}

function renderActiveTable() {
  if (!state.result) return;
  const config = tableConfig(state.activeResultTab);
  const rows = filterRows(config.rows);
  const totalPages = Math.max(1, Math.ceil(rows.length / state.pageSize));
  state.currentPage = Math.min(state.currentPage, totalPages);
  const pageRows = rows.slice((state.currentPage - 1) * state.pageSize, state.currentPage * state.pageSize);

  $("#resultTableHead").innerHTML = `<tr>${config.columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}</tr>`;
  $("#resultTableBody").innerHTML = pageRows.length
    ? pageRows.map((row) => `<tr>${config.columns.map((column) => `<td class="${column.className || ""}">${column.render ? column.render(row) : escapeHtml(row[column.key] ?? "")}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${config.columns.length}" class="empty-cell">조건에 맞는 항목이 없습니다.</td></tr>`;
  renderPagination(rows.length);
}

function tableConfig(tab) {
  const issueColumns = [
    { label: "경로", render: (row) => escapeHtml(row.routeLabel) },
    { label: "매장", key: "store" },
    { label: "사번", key: "employeeId" },
    { label: "이름", render: (row) => `<strong>${escapeHtml(row.name)}</strong>` },
    { label: "일자", key: "date" },
    { label: "요일", key: "weekday" },
    { label: "계획", render: (row) => `<span class="plan-pill">${escapeHtml(row.planStatus)}</span>` },
    { label: "실제 근태", render: (row) => escapeHtml(row.actualStatus || "-") },
    { label: "출근기록", key: "clockStatus" },
    { label: "판정", className: "message-cell", render: (row) => `<span class="${row.issueType === "missing_clock_in" ? "status-pill" : "warning-pill"}">${escapeHtml(row.result)}</span>` },
    { label: "사유", className: "message-cell", key: "reason" },
  ];
  const commonSummary = [
    { label: "경로", render: (row) => escapeHtml(row.routeLabel) },
    { label: "매장", key: "store" },
    { label: "사번", key: "employeeId" },
    { label: "이름", render: (row) => `<strong>${escapeHtml(row.name)}</strong>` },
  ];

  const configs = {
    missing: { rows: state.result.missingRows, columns: issueColumns },
    unexpected: { rows: state.result.unexpectedRows, columns: issueColumns },
    mismatch: {
      rows: state.result.mismatchRows,
      columns: [
        ...commonSummary,
        { label: "일자", key: "date" },
        { label: "요일", key: "weekday" },
        { label: "계획 근태", render: (row) => `<span class="plan-pill">${escapeHtml(row.planStatus)}</span>` },
        { label: "실제 근태", key: "actualStatus" },
        { label: "출근기록", key: "clockStatus" },
        { label: "판정", render: (row) => `<span class="warning-pill">${escapeHtml(row.result)}</span>` },
        { label: "검토 사유", className: "message-cell", key: "reason" },
      ],
    },
    excess: {
      rows: state.result.dayoffExcessRows,
      columns: [
        ...commonSummary,
        { label: "기본 휴무 기준", render: (row) => formatDays(row.baseAllowance) },
        { label: "휴무 사용", render: (row) => formatDays(row.basicDayoffUsed) },
        { label: "초과 휴무", render: (row) => formatDays(row.baseExcess) },
        { label: "표기 대체휴무", render: (row) => formatDays(row.explicitSubDayoffUsed) },
        { label: "총 대체휴무 필요", render: (row) => formatDays(row.substituteNeeded) },
        { label: "판정", className: "message-cell", render: renderJudgment },
      ],
    },
    balance: {
      rows: state.result.balanceRows,
      columns: [
        ...commonSummary,
        { label: "당월 가용", render: (row) => formatDays(row.availableSubstitute) },
        { label: "당월 필요", render: (row) => formatDays(row.substituteNeeded) },
        { label: "당월 적용", render: (row) => formatDays(row.substituteApplied) },
        { label: "다음 달 이월 잔여", render: (row) => formatDays(row.remainingSubstitute) },
        { label: "당월 만료", render: (row) => formatDays(row.expiredSubstitute) },
        { label: "초과 사용", render: (row) => formatDays(row.shortage) },
        { label: "판정", className: "message-cell", render: renderJudgment },
      ],
    },
    shortage: {
      rows: state.result.shortageRows,
      columns: [
        ...commonSummary,
        { label: "대체휴무 필요", render: (row) => formatDays(row.substituteNeeded) },
        { label: "적용 가능", render: (row) => formatDays(row.substituteApplied) },
        { label: "초과 사용", render: (row) => `<span class="status-pill">${formatDays(row.shortage)}</span>` },
        { label: "판정", className: "message-cell", render: renderJudgment },
      ],
    },
    annual: {
      rows: state.result.annualRows,
      columns: [
        ...commonSummary,
        { label: "당월 연차", render: (row) => formatDays(row.currentAnnualLeave) },
        { label: "누적 연차", render: (row) => `<strong>${formatDays(row.cumulativeAnnualLeave)}</strong>` },
        { label: "당월 등록 내역", className: "message-cell", render: (row) => escapeHtml(formatAnnualEvents(row.annualLeaveEvents)) },
      ],
    },
  };
  return configs[tab] || configs.missing;
}

function renderJudgment(row) {
  const statusClass = row.shortage > 0 ? "status-pill" : row.baseExcess > 0 || row.substituteNeeded > 0 ? "warning-pill" : "success-pill";
  return `<span class="${statusClass}">${escapeHtml(row.judgment)}</span>`;
}

function filterRows(rows) {
  const query = $("#searchInput").value.trim().toLowerCase();
  const store = $("#storeFilter").value;
  return rows.filter((row) => {
    const haystack = Object.values(row).map((value) => typeof value === "object" ? JSON.stringify(value) : String(value ?? "")).join(" ").toLowerCase();
    return (!query || haystack.includes(query)) && (!store || row.store === store);
  });
}

function renderPagination(rowCount) {
  const totalPages = Math.max(1, Math.ceil(rowCount / state.pageSize));
  const current = state.currentPage;
  const visible = [];
  for (let page = Math.max(1, current - 2); page <= Math.min(totalPages, current + 2); page += 1) visible.push(page);
  $("#resultPagination").innerHTML = `
    <button class="page-button" data-page="${Math.max(1, current - 1)}">‹</button>
    ${visible.map((page) => `<button class="page-button ${page === current ? "active" : ""}" data-page="${page}">${page}</button>`).join("")}
    <button class="page-button" data-page="${Math.min(totalPages, current + 1)}">›</button>`;
  $$("#resultPagination .page-button").forEach((button) => button.addEventListener("click", () => {
    state.currentPage = Number(button.dataset.page);
    renderActiveTable();
  }));
}

function exportResults() {
  if (!state.result) return;
  const result = state.result;
  const workbook = XLSX.utils.book_new();

  appendSheet(workbook, "1_근무인데 출근없음", result.missingRows, issueExportColumns());
  appendSheet(workbook, "2_휴무휴가인데 출근", result.unexpectedRows, issueExportColumns());
  appendSheet(workbook, "3_계획실제 불일치", result.mismatchRows, [
    ["경로", (row) => row.routeLabel], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
    ["일자", (row) => row.date], ["요일", (row) => row.weekday], ["계획 근태", (row) => row.planStatus], ["실제 근태", (row) => row.actualStatus],
    ["실제출근시간", (row) => row.actualIn], ["변경출근시간", (row) => row.changedIn], ["판정", (row) => row.result], ["사유", (row) => row.reason],
  ]);
  appendSheet(workbook, "4_기본휴무 초과", result.dayoffExcessRows, [
    ...summaryIdentityExportColumns(), ["기본 휴무 기준", (row) => row.baseAllowance], ["휴무 사용", (row) => row.basicDayoffUsed],
    ["초과 휴무", (row) => row.baseExcess], ["표기 대체휴무", (row) => row.explicitSubDayoffUsed],
    ["총 대체휴무 필요", (row) => row.substituteNeeded], ["판정", (row) => row.judgment],
  ]);
  appendSheet(workbook, "5_대체휴무 잔여", result.balanceRows, [
    ...summaryIdentityExportColumns(), ["당월 가용", (row) => row.availableSubstitute], ["당월 필요", (row) => row.substituteNeeded],
    ["당월 적용", (row) => row.substituteApplied], ["다음 달 이월 잔여", (row) => row.remainingSubstitute],
    ["당월 만료", (row) => row.expiredSubstitute], ["초과 사용", (row) => row.shortage], ["판정", (row) => row.judgment],
  ]);
  appendSheet(workbook, "6_대체휴무 초과", result.shortageRows, [
    ...summaryIdentityExportColumns(), ["대체휴무 필요", (row) => row.substituteNeeded], ["적용 가능", (row) => row.substituteApplied],
    ["초과 사용", (row) => row.shortage], ["판정", (row) => row.judgment],
  ]);
  appendSheet(workbook, "7_연차 등록누적", result.annualRows, [
    ...summaryIdentityExportColumns(), ["당월 연차", (row) => row.currentAnnualLeave], ["누적 연차", (row) => row.cumulativeAnnualLeave],
    ["당월 등록 내역", (row) => formatAnnualEvents(row.annualLeaveEvents)],
  ]);

  XLSX.writeFile(workbook, `${result.routeLabel}_${result.targetMonth}_근태종합결과.xlsx`);
}

function issueExportColumns() {
  return [
    ["경로", (row) => row.routeLabel], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
    ["일자", (row) => row.date], ["요일", (row) => row.weekday], ["계획 근태", (row) => row.planStatus], ["실제 근태", (row) => row.actualStatus],
    ["실제출근시간", (row) => row.actualIn], ["변경출근시간", (row) => row.changedIn], ["판정", (row) => row.result], ["사유", (row) => row.reason],
  ];
}

function summaryIdentityExportColumns() {
  return [
    ["경로", (row) => row.routeLabel], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
  ];
}

function appendSheet(workbook, sheetName, rows, columns) {
  const values = [columns.map(([header]) => header)];
  for (const row of rows) values.push(columns.map(([, getter]) => getter(row)));
  const sheet = XLSX.utils.aoa_to_sheet(values);
  const widths = columns.map(([header]) => ({ wch: Math.max(12, Math.min(32, String(header).length * 2 + 4)) }));
  sheet["!cols"] = widths;
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
}

async function saveClosure() {
  if (!state.result) return;
  if (!(state.backend.available && state.backend.configured)) {
    showToast("D1 연결이 필요합니다. 월 마감은 서버에만 저장합니다.");
    return;
  }
  if (!state.backend.loggedIn) {
    openLogin();
    return;
  }

  const button = $("#saveClosureButton");
  const result = state.result;
  const payload = {
    route: result.route,
    month: result.targetMonth,
    cutoffDate: result.cutoffDate,
    planFileName: result.planFileName,
    attendanceFileName: result.attendanceFileName,
    planPeople: result.planPeople,
    attendancePeople: result.attendancePeople,
    matchedPeople: result.matchedPeople,
    matchRate: result.matchRate,
    issueRows: [...result.missingRows, ...result.unexpectedRows],
    mismatchRows: result.mismatchRows,
    employeeFacts: result.employeeFacts,
  };

  try {
    button.disabled = true;
    button.textContent = "월 마감 저장·재계산 중...";
    const response = await fetch("/api/closures", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (response.status === 401) {
      await checkBackend();
      openLogin();
      return;
    }
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "월 마감 저장 실패");

    applyServerSummaries(data.summaries || []);
    await Promise.all([loadHistory(), loadGrants()]);
    const action = data.replaced ? "기존 월 마감을 완전히 교체했습니다." : "새 월 마감을 저장했습니다.";
    showToast(`${action} ${data.affectedMonths || 0}개 월의 연차·대체휴무 누적을 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message || "월 마감 저장 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "월 마감 교체 저장";
  }
}

function applyServerSummaries(serverRows) {
  if (!state.result || !serverRows.length) return;
  const factsByEmployee = new Map(state.result.employeeFacts.map((row) => [row.employeeId, row]));
  const employeeSummaries = serverRows.map((row) => {
    const fact = factsByEmployee.get(row.employee_id) || {};
    return {
      ...fact,
      route: row.route,
      routeLabel: ROUTE_LABELS[row.route] || state.result.routeLabel,
      store: row.store,
      employeeId: row.employee_id,
      name: row.employee_name,
      baseAllowance: roundHalf(row.base_allowance),
      basicDayoffUsed: roundHalf(row.basic_dayoff_used),
      explicitSubDayoffUsed: roundHalf(row.explicit_sub_dayoff_used),
      baseExcess: roundHalf(row.base_excess),
      substituteNeeded: roundHalf(row.substitute_needed),
      availableSubstitute: roundHalf(row.available_substitute),
      substituteApplied: roundHalf(row.substitute_applied),
      remainingSubstitute: roundHalf(row.remaining_substitute),
      expiredSubstitute: roundHalf(row.expired_substitute),
      shortage: roundHalf(row.shortage),
      currentAnnualLeave: roundHalf(row.current_annual_leave),
      cumulativeAnnualLeave: roundHalf(row.cumulative_annual_leave),
      judgment: row.judgment || "",
      annualLeaveEvents: fact.annualLeaveEvents || [],
      substituteEvents: fact.substituteEvents || [],
    };
  });
  sortSummaries(employeeSummaries);
  const updated = assembleResultCollections({
    ...state.result,
    employeeSummaries,
    employeeFacts: state.result.employeeFacts,
  });
  state.result = updated;
  refreshResultMetrics();
  renderActiveTable();
}

async function loadHistory() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    renderHistory([]);
    return;
  }
  try {
    const response = await fetch("/api/closures", { cache: "no-store" });
    if (!response.ok) throw new Error("월 마감 기록 조회 실패");
    renderHistory((await response.json()).items || []);
  } catch (error) {
    showToast(error.message);
  }
}

function renderHistory(items) {
  $("#historyEmpty").classList.toggle("hidden", items.length > 0);
  $("#historyList").innerHTML = items.map((item) => {
    const routeLabel = ROUTE_LABELS[item.route] || item.route;
    return `<article class="history-card">
      <div><div class="month">${escapeHtml(item.month)}</div><div class="history-meta">${escapeHtml(routeLabel)} 경로 · 기준 ${escapeHtml(item.cutoff_date || "-")}</div></div>
      <div><strong>${escapeHtml(item.plan_file_name || "계획표")}</strong><div class="history-meta">${escapeHtml(item.attendance_file_name || "근태표")} · ${item.created_at ? new Date(`${item.created_at}Z`).toLocaleString("ko-KR") : ""}</div></div>
      <div class="history-kpis">
        <div><span>출근기록 없음</span><strong>${number(item.missing_count || 0)}</strong></div>
        <div><span>휴무·휴가 출근</span><strong>${number(item.unexpected_count || 0)}</strong></div>
        <div><span>불일치</span><strong>${number(item.mismatch_count || 0)}</strong></div>
        <div><span>대체휴무 초과</span><strong>${number(item.substitute_shortage_people || 0)}</strong></div>
        <div><span>연차 등록자</span><strong>${number(item.annual_leave_people || 0)}</strong></div>
      </div>
    </article>`;
  }).join("");
}

async function saveGrant(event) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    openLogin();
    return;
  }

  const payload = {
    route: $("#grantRoute").value,
    grantMonth: $("#grantMonth").value,
    grantedDays: Number($("#grantDays").value),
    validFrom: $("#grantValidFrom").value,
    validTo: $("#grantValidTo").value,
    reason: text($("#grantReason").value),
    note: text($("#grantNote").value),
  };
  if (!payload.grantMonth || !(payload.grantedDays > 0) || !payload.validFrom || !payload.validTo) {
    showToast("발생 월·공통 부여 일수·사용기간을 확인해 주세요.");
    return;
  }
  if (payload.validFrom > payload.validTo) {
    showToast("사용 종료일은 시작일보다 빠를 수 없습니다.");
    return;
  }

  const button = $("#grantSubmitButton");
  try {
    button.disabled = true;
    button.textContent = "저장·재계산 중...";
    const response = await fetch("/api/substitute-grants", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "공통 대체휴무 저장 실패");
    resetGrantForm();
    await Promise.all([loadGrants(), loadHistory()]);
    showToast(`${data.replaced ? "기존 경로·월 설정을 교체" : "공통 부여 설정을 저장"}하고 ${data.affectedMonths || 0}개 월을 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message);
  } finally {
    button.disabled = false;
    button.textContent = "경로 공통 부여 저장";
  }
}

async function loadGrants() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    state.grants = [];
    $("#grantTableBody").innerHTML = `<tr><td colspan="11" class="empty-cell">관리자 로그인 후 조회할 수 있습니다.</td></tr>`;
    renderGrantSummary([]);
    return;
  }
  try {
    const route = $("#grantRouteFilter").value;
    const query = route ? `?route=${encodeURIComponent(route)}` : "";
    const response = await fetch(`/api/substitute-grants${query}`, { cache: "no-store" });
    if (!response.ok) throw new Error("공통 대체휴무 조회 실패");
    state.grants = (await response.json()).items || [];
    renderGrants(state.grants);
  } catch (error) {
    showToast(error.message);
  }
}

function renderGrants(items) {
  renderGrantSummary(items);
  $("#grantTableBody").innerHTML = items.length ? items.map((item) => {
    const status = grantStatus(item);
    const statusClass = status === "사용 가능" ? "success-pill" : status === "만료" ? "status-pill" : "neutral-pill";
    return `<tr>
      <td>${escapeHtml(ROUTE_LABELS[item.route] || item.route)}</td>
      <td><strong>${escapeHtml(item.grant_month)}</strong></td>
      <td>${formatDays(item.granted_days)}</td>
      <td>${number(item.eligible_people || 0)}명</td>
      <td>${formatDays(item.assigned_days)}</td>
      <td>${formatDays(item.used_days)}</td>
      <td>${formatDays(item.unused_days)}</td>
      <td>${escapeHtml(item.valid_from)} ~ ${escapeHtml(item.valid_to)}</td>
      <td><span class="${statusClass}">${escapeHtml(status)}</span></td>
      <td class="message-cell">${escapeHtml([item.reason, item.note].filter(Boolean).join(" · "))}</td>
      <td class="action-cell"><button class="btn secondary small grant-edit" data-id="${escapeHtml(item.id)}" type="button">수정</button><button class="btn danger small grant-delete" data-id="${escapeHtml(item.id)}" type="button">삭제</button></td>
    </tr>`;
  }).join("") : `<tr><td colspan="11" class="empty-cell">등록된 경로 공통 대체휴무가 없습니다.</td></tr>`;

  $$(".grant-edit").forEach((button) => button.addEventListener("click", () => editGrant(button.dataset.id)));
  $$(".grant-delete").forEach((button) => button.addEventListener("click", () => deleteGrant(button.dataset.id)));
}

function renderGrantSummary(items) {
  const eligiblePeople = items.reduce((sum, item) => sum + Number(item.eligible_people || 0), 0);
  const assigned = items.reduce((sum, item) => sum + Number(item.assigned_days || 0), 0);
  const used = items.reduce((sum, item) => sum + Number(item.used_days || 0), 0);
  const unused = items.reduce((sum, item) => sum + Number(item.unused_days || 0), 0);
  $("#grantSummary").innerHTML = [
    ["공통 설정", `${items.length}건`],
    ["적용 인원 합계", `${number(eligiblePeople)}명`],
    ["총 부여", formatDays(assigned)],
    ["총 사용", formatDays(used)],
    ["총 미사용", formatDays(unused)],
  ].map(([label, value]) => `<div class="summary-chip"><span>${label}</span><strong>${value}</strong></div>`).join("");
}

function editGrant(id) {
  const item = state.grants.find((grant) => grant.id === id);
  if (!item) return;
  $("#grantEditingId").value = item.id;
  $("#grantRoute").value = item.route;
  $("#grantMonth").value = item.grant_month;
  $("#grantDays").value = Number(item.granted_days || 0);
  $("#grantValidFrom").value = item.valid_from;
  $("#grantValidTo").value = item.valid_to;
  $("#grantReason").value = item.reason || "";
  $("#grantNote").value = item.note || "";
  $("#grantSubmitButton").textContent = "경로·월 공통 설정 교체";
  $("#grantCancelEdit").classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetGrantForm() {
  $("#grantEditingId").value = "";
  $("#grantDays").value = "1";
  $("#grantReason").value = "";
  $("#grantNote").value = "";
  $("#grantSubmitButton").textContent = "경로 공통 부여 저장";
  $("#grantCancelEdit").classList.add("hidden");
}

async function deleteGrant(id) {
  if (!confirm("이 경로 공통 대체휴무 설정을 삭제하시겠습니까? 관련 월의 잔여·초과·이월이 모두 다시 계산됩니다.")) return;
  try {
    const response = await fetch(`/api/substitute-grants/${encodeURIComponent(id)}`, { method: "DELETE" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "삭제 실패");
    resetGrantForm();
    await Promise.all([loadGrants(), loadHistory()]);
    showToast(`공통 부여 설정을 삭제하고 ${data.affectedMonths || 0}개 월을 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message);
  }
}

function grantStatus(item) {
  if (Number(item.eligible_people || 0) === 0) return "월 마감 대기";
  const today = toISODate(new Date());
  if (item.valid_to < today) return "만료";
  if (item.valid_from > today) return "사용 전";
  if (Number(item.unused_days || 0) <= 0) return "소진";
  return "사용 가능";
}

async function checkBackend() {
  try {
    const response = await fetch("/api/auth", { cache: "no-store" });
    if (!response.ok) throw new Error();
    const data = await response.json();
    state.backend = { available: true, configured: Boolean(data.configured), loggedIn: Boolean(data.loggedIn) };
  } catch {
    state.backend = { available: false, configured: false, loggedIn: false };
  }
  updateStorageUI();
}

function updateStorageUI() {
  const badge = $("#storageBadge");
  const loginButton = $("#loginButton");
  const logoutButton = $("#logoutButton");
  if (state.backend.available && state.backend.configured) {
    badge.className = "badge cloud";
    badge.textContent = state.backend.loggedIn ? "D1 영구 저장 연결" : "D1 로그인 필요";
    loginButton.classList.toggle("hidden", state.backend.loggedIn);
    logoutButton.classList.toggle("hidden", !state.backend.loggedIn);
  } else {
    badge.className = "badge local";
    badge.textContent = "서버 설정 확인 필요";
    loginButton.classList.add("hidden");
    logoutButton.classList.add("hidden");
  }
}

function openLogin() {
  $("#loginError").textContent = "";
  $("#passwordInput").value = "";
  $("#loginDialog").showModal();
}

async function login(event) {
  event.preventDefault();
  try {
    const response = await fetch("/api/auth", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ password: $("#passwordInput").value }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "로그인 실패");
    $("#loginDialog").close();
    await checkBackend();
    await Promise.all([loadHistory(), loadGrants()]);
    showToast("관리자 로그인되었습니다.");
  } catch (error) {
    $("#loginError").textContent = error.message;
  }
}

async function logout() {
  await fetch("/api/auth", { method: "DELETE" });
  await checkBackend();
  renderHistory([]);
  renderGrants([]);
  showToast("로그아웃되었습니다.");
}

function switchView(view) {
  $$(".tab[data-view]").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === view));
  $$(".view").forEach((section) => section.classList.remove("active"));
  $(`#${view}View`).classList.add("active");
  if (view === "history") loadHistory();
  if (view === "substitute") loadGrants();
}

function resetAll() {
  state.planFile = null;
  state.attendanceFile = null;
  state.result = null;
  state.priorLedger = emptyLedger();
  $("#planFile").value = "";
  $("#attendanceFile").value = "";
  setPlanFile(null);
  setAttendanceFile(null);
  setDefaultDates();
  syncRouteRuleHelp();
  $("#resultArea").classList.add("hidden");
  $("#emptyState").classList.remove("hidden");
}

function selectedRoute() {
  return document.querySelector('input[name="route"]:checked').value;
}

function countWeekendDays(monthText) {
  if (!/^\d{4}-\d{2}$/.test(monthText || "")) return 0;
  const [year, month] = monthText.split("-").map(Number);
  const days = new Date(year, month, 0).getDate();
  let count = 0;
  for (let day = 1; day <= days; day += 1) {
    const weekday = new Date(year, month - 1, day).getDay();
    if (weekday === 0 || weekday === 6) count += 1;
  }
  return count;
}

function findHeaderRow(matrix, requiredHeaders) {
  return matrix.slice(0, 15).findIndex((row) => {
    const headers = row.map(normalizeHeader);
    return requiredHeaders.every((required) => headers.includes(normalizeHeader(required)));
  });
}

function findHeaderIndex(headers, candidates) {
  for (const candidate of candidates) {
    const index = headers.indexOf(normalizeHeader(candidate));
    if (index >= 0) return index;
  }
  return -1;
}

function normalizeHeader(value) {
  return text(value).replace(/\s+/g, "").replace(/[\[\]]/g, "");
}

function normalizeStatus(value) {
  return text(value).replace(/\s+/g, "");
}

function normalizeEmployeeId(value) {
  return text(value).toUpperCase().replace(/\s+/g, "");
}

function normalizeStore(value) {
  return text(value)
    .replace(/^\d+_/, "")
    .replace(/홈플러스/g, "")
    .replace(/전자랜드/g, "")
    .replace(/점$/g, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function text(value) {
  return value == null ? "" : String(value).trim();
}

function parseDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toISODate(value);
  const digits = text(value).replace(/\D/g, "");
  if (digits.length >= 8) return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? "" : toISODate(parsed);
}

function detectRouteFromPlan(rows) {
  const homeplusCount = rows.filter((row) => row.store.includes("홈플러스")).length;
  const electrolandCount = rows.filter((row) => row.store.includes("전자랜드")).length;
  if (!homeplusCount && !electrolandCount) return null;
  return homeplusCount > electrolandCount ? "homeplus" : electrolandCount > homeplusCount ? "electroland" : null;
}

function detectRouteFromAttendance(rows) {
  const homeplusCount = rows.filter((row) => row.location.includes("홈플러스")).length;
  const electrolandCount = rows.filter((row) => row.location.includes("전자랜드")).length;
  if (!homeplusCount && !electrolandCount) return null;
  return homeplusCount > electrolandCount ? "homeplus" : electrolandCount > homeplusCount ? "electroland" : null;
}

function groupBy(items, selector) {
  const map = new Map();
  items.forEach((item) => {
    const key = selector(item);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  });
  return map;
}

function uniquePeople(rows) {
  return new Set(rows.map((row) => row.employeeId).filter(Boolean)).size;
}

function sortSummaries(rows) {
  rows.sort((a, b) => b.shortage - a.shortage || b.baseExcess - a.baseExcess || a.store.localeCompare(b.store) || a.name.localeCompare(b.name));
}

function issueSort(a, b) {
  return a.date.localeCompare(b.date) || a.store.localeCompare(b.store) || a.name.localeCompare(b.name);
}

function formatAnnualEvents(events) {
  if (!Array.isArray(events) || !events.length) return "당월 등록 없음";
  return events.map((event) => `${event.date} ${event.planStatus}(${formatDays(event.days)})`).join(" · ");
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

function toISODate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function number(value) {
  return new Intl.NumberFormat("ko-KR").format(Number(value) || 0);
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}

function formatDays(value) {
  const numberValue = roundHalf(value);
  return `${Number.isInteger(numberValue) ? numberValue : numberValue.toFixed(1)}일`;
}

function escapeHtml(value) {
  return text(value).replace(/[&<>'"]/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
  }[char]));
}

let toastTimer;
function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 5200);
}
