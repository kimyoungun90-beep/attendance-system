import { buildFinalTemplateWorkbook, buildFinalTemplateFile } from "./final-template.js?v=39";

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
  annualFile: null,
  referenceFile: null,
  closureBaseFile: null,
  closureTargetFile: null,
  closureComparison: null,
  result: null,
  priorLedger: emptyLedger(),
  backend: { available: false, configured: false, loggedIn: false, fileStorageConfigured: false, fileStorageMode: "d1" },
  activeResultTab: "missing",
  currentPage: 1,
  pageSize: 50,
  grants: [],
  archiveFiles: [],
  workforceUploads: [],
  workforce: null,
  personnelChecks: null,
  annualLeaveDashboard: null,
  annualBaselinePreviews: { homeplus: null, electroland: null },
  annualMonthlyPreviews: { homeplus: null, electroland: null },
  archiveUploadPreviews: { homeplus: [], electroland: [] },
};

init();

async function init() {
  setDefaultDates();
  bindEvents();
  setupDropzone("planDropzone", "planFile", setPlanFile);
  setupDropzone("attendanceDropzone", "attendanceFile", setAttendanceFile);
  setupDropzone("annualDropzone", "annualFile", setAnnualFile);
  setupDropzone("referenceDropzone", "referenceFile", setReferenceFile);
  setupDropzone("closureBaseDropzone", "closureBaseFile", setClosureBaseFile);
  setupDropzone("closureTargetDropzone", "closureTargetFile", setClosureTargetFile);
  await checkBackend();
  syncRouteRuleHelp();
  if (state.backend.loggedIn) await Promise.all([loadHistory(), loadGrants(), loadArchiveFiles(), loadWorkforceUploads(), loadAnnualLeaveDashboard(), loadPersonnelChecks()]);
}

function bindEvents() {
  $("#analyzeButton").addEventListener("click", analyzeFiles);
  $("#resetButton").addEventListener("click", resetAll);
  $("#searchInput").addEventListener("input", () => { state.currentPage = 1; renderActiveTable(); });
  $("#storeFilter").addEventListener("change", () => { state.currentPage = 1; renderActiveTable(); });
  $("#managerFilter").addEventListener("change", () => { state.currentPage = 1; renderActiveTable(); });
  $("#exportButton").addEventListener("click", exportResults);
  $("#saveClosureButton").addEventListener("click", saveClosure);
  $("#refreshHistoryButton").addEventListener("click", loadHistory);
  $("#refreshGrantButton").addEventListener("click", loadGrants);
  $("#grantRouteFilter").addEventListener("change", loadGrants);
  $("#grantForm").addEventListener("submit", saveGrant);
  $("#grantMonth").addEventListener("change", () => { syncGrantDates(); syncGrantFormVisibility(); });
  $("#grantScope").addEventListener("change", syncGrantFormVisibility);
  $("#grantType").addEventListener("change", syncGrantFormVisibility);
  $("#grantEligibility").addEventListener("change", syncGrantFormVisibility);
  $("#grantCancelEdit").addEventListener("click", resetGrantForm);
  $("#archiveForm").addEventListener("submit", saveArchiveFiles);
  $("#refreshArchiveButton").addEventListener("click", loadArchiveFiles);
  $("#archiveRouteFilter").addEventListener("change", loadArchiveFiles);
  $("#archiveMonthFilter").addEventListener("change", loadArchiveFiles);
  $("#archiveMonth").addEventListener("change", previewAllArchiveUploads);
  $("#archiveHomeplusFiles").addEventListener("change", () => previewArchiveUploads("homeplus"));
  $("#archiveElectrolandFiles").addEventListener("change", () => previewArchiveUploads("electroland"));
  $("#compareClosuresButton").addEventListener("click", analyzeClosureComparison);
  $("#resetClosureCompare").addEventListener("click", resetClosureComparison);
  $("#exportClosureCompare").addEventListener("click", exportClosureComparison);
  $("#workforceForm").addEventListener("submit", saveWorkforceFile);
  $("#refreshWorkforceButton").addEventListener("click", loadWorkforceUploads);
  for (const route of ["homeplus", "electroland"]) {
    $(`#${route}AnnualBaselineForm`).addEventListener("submit", (event) => saveAnnualBaselineForRoute(event, route));
    $(`#${route}AnnualBaselineFile`).addEventListener("change", () => previewAnnualBaselineForRoute(route));
    $(`#${route}AnnualBaselineDate`).addEventListener("change", () => previewAnnualBaselineForRoute(route));
    $(`#${route}AnnualMonthlyForm`).addEventListener("submit", (event) => saveAnnualMonthlyForRoute(event, route));
    $(`#${route}AnnualMonthlyFile`).addEventListener("change", () => previewAnnualMonthlyForRoute(route));
    $(`#${route}AnnualMonthlyMonth`).addEventListener("change", () => previewAnnualMonthlyForRoute(route));
  }
  $("#annualLedgerRoute").addEventListener("change", loadAnnualLeaveDashboard);
  $("#annualLedgerMonth").addEventListener("change", loadAnnualLeaveDashboard);
  $("#refreshAnnualLedger").addEventListener("click", loadAnnualLeaveDashboard);
  $("#exportAnnualLedger").addEventListener("click", exportAnnualLeaveDashboard);
  $("#workforceMonth").addEventListener("change", previewWorkforceFile);
  $("#workforceFile").addEventListener("change", previewWorkforceFile);
  $("#portalReferenceFile").addEventListener("change", previewWorkforceFile);
  $("#refreshPersonnelChecks").addEventListener("click", loadPersonnelChecks);
  $("#personnelCheckMonth").addEventListener("change", loadPersonnelChecks);
  $("#personnelCheckRoute").addEventListener("change", loadPersonnelChecks);
  $("#savePersonnelChecks").addEventListener("click", savePersonnelChecks);
  $("#loginButton").addEventListener("click", openLogin);
  $("#logoutButton").addEventListener("click", logout);
  $("#loginCancel").addEventListener("click", () => $("#loginDialog").close());
  $("#loginForm").addEventListener("submit", login);
  $("#targetMonth").addEventListener("change", () => { syncCutoffWithMonth(); syncRouteRuleHelp(); updateWorkforceStatus(); });
  $$('input[name="route"]').forEach((input) => input.addEventListener("change", () => { syncRouteRuleHelp(); updateWorkforceStatus(); }));
  $$(".tab[data-view]").forEach((tab) => tab.addEventListener("click", () => switchView(tab.dataset.view)));
  $$(".inner-tab").forEach((tab) => tab.addEventListener("click", () => switchResultTab(tab.dataset.resultTab)));
}

function setDefaultDates() {
  const now = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  $("#targetMonth").value = month;
  $("#cutoffDate").value = toISODate(now);
  $("#grantMonth").value = month;
  $("#archiveMonth").value = month;
  $("#workforceMonth").value = month;
  $("#personnelCheckMonth").value = month;
  $("#homeplusAnnualMonthlyMonth").value = month;
  $("#electrolandAnnualMonthlyMonth").value = month;
  $("#annualLedgerMonth").value = month;
  syncGrantDates();
  syncGrantFormVisibility();
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
  if (!text($("#grantEditingId")?.value)) $("#grantOccurrenceDates").value = `${month}-01`;
}

function syncGrantFormVisibility() {
  const scope = $("#grantScope")?.value || "route";
  const grantType = $("#grantType")?.value || "substitute";
  const routeSettlement = scope === "route" && ["substitute", "compensation"].includes(grantType);
  $("#grantEmployeeField")?.style.setProperty("display", scope === "employee" ? "block" : "none");
  $("#grantAmountRow")?.classList.toggle("single", scope !== "employee");
  $("#grantEligibilityField")?.classList.toggle("hidden", scope !== "route");
  $("#grantExclusionField")?.classList.toggle("hidden", scope !== "route");
  $("#grantCriterionField")?.classList.add("hidden");
  if (routeSettlement) {
    $("#grantEligibility").value = "all";
    $("#grantEligibility").disabled = true;
    $("#grantCriterionDate").value = "";
  } else {
    $("#grantEligibility").disabled = false;
  }
  const help = $("#grantRuleHelp");
  if (help) help.textContent = grantType === "compensation" && scope === "route"
    ? "발생일 당일까지 입사한 직원 중 실제 출근자는 입력한 보상휴가 일수를 받고, 미출근자는 보상휴가 없이 해당 월 기본 휴무가 1일 늘어납니다. 홈플러스·전자랜드에 동일 적용됩니다."
    : grantType === "substitute" && scope === "route"
      ? "발생일 당일까지 입사한 직원에게 출근·휴무 여부와 관계없이 대체휴무를 동일 부여합니다. 기본 휴무는 가감하지 않습니다."
      : "사번별 개별 부여는 지정 사번에게 입력한 휴가 일수를 부여합니다.";
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

function setAnnualFile(file) {
  state.annualFile = file;
  $("#annualFileName").textContent = file ? file.name : "연차신청현황 파일 선택";
}

function setReferenceFile(file) {
  state.referenceFile = file;
  $("#referenceFileName").textContent = file ? file.name : "증빙 O 입력 최종본 선택";
}

function setClosureBaseFile(file) {
  state.closureBaseFile = file;
  state.closureComparison = null;
  $("#closureBaseFileName").textContent = file ? file.name : "기준 최종본을 선택하거나 끌어놓기";
  resetClosureComparisonOutput();
}

function setClosureTargetFile(file) {
  state.closureTargetFile = file;
  state.closureComparison = null;
  $("#closureTargetFileName").textContent = file ? file.name : "비교 최종본을 선택하거나 끌어놓기";
  resetClosureComparisonOutput();
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

    if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
      openLogin();
      throw new Error("월별 인력·매장매칭 자료를 불러오려면 관리자 로그인이 필요합니다.");
    }
    const [planWorkbook, attendanceWorkbook, workforce, annualWorkbook, referenceWorkbook, annualLedger, personnelCheckData] = await Promise.all([
      fileToWorkbookSheets(state.planFile),
      fileToWorkbookSheets(state.attendanceFile),
      loadWorkforceMonth(targetMonth, route),
      state.annualFile ? fileToWorkbookSheets(state.annualFile) : Promise.resolve(null),
      state.referenceFile ? fileToWorkbookSheets(state.referenceFile) : Promise.resolve(null),
      fetchAnnualLeaveDashboard(route, targetMonth, cutoffDate),
      fetchPersonnelChecks(targetMonth, route),
    ]);
    if (!workforce?.members?.length) throw new Error(`${targetMonth} ${ROUTE_LABELS[route]} 경로의 인력·매장매칭 파일이 없습니다. 인력 매칭에서 먼저 등록해 주세요.`);
    const plan = parsePlan(planWorkbook, targetMonth);
    const attendance = parseAttendance(attendanceWorkbook, targetMonth);
    state.priorLedger = await loadPriorLedger(route, targetMonth);

    const parsedReference = referenceWorkbook ? parseReferenceFinalWorkbook(referenceWorkbook, targetMonth, route) : null;
    const evidenceOverrides = parsedReference?.evidenceKeys || [];
    const workflowOverrides = parsedReference?.workflowOverrides || emptyWorkflowOverrides();
    const personnelOverrides = mergePersonnelOverrides(personnelCheckData?.items || [], parsedReference?.personnelOverrides || [], targetMonth, route);
    const excludedPersonnelIds = personnelExcludedIds(personnelOverrides, targetMonth, route);
    const result = compareAttendance({ plan, attendance, route, targetMonth, cutoffDate, ledger: state.priorLedger, evidenceKeys: evidenceOverrides });
    appendWorkforceMatchingIssues(result, workforce, plan, route, targetMonth, excludedPersonnelIds);
    applyEvidenceOverrides(result, evidenceOverrides);
    applyWorkflowOverrides(result, workflowOverrides);
    const parsedAnnual = annualWorkbook
      ? parseAnnualApplications(annualWorkbook, targetMonth)
      : annualApplicationsToParsed(annualLedger?.applications || [], targetMonth);
    const annualComparison = parsedAnnual?.rows?.length
      ? compareAnnualApplications(parsedAnnual, plan, attendance, targetMonth, cutoffDate, workforce, excludedPersonnelIds, evidenceOverrides)
      : emptyAnnualComparison();
    mergeAnnualLeaveLedger(result, annualLedger, annualComparison);
    const referenceComparison = parsedReference
      ? compareReferenceFinal(parsedReference, result, targetMonth, cutoffDate, workforce)
      : emptyReferenceComparison();
    const managerRequests = buildManagerRequests(result, annualComparison, workforce, targetMonth, workflowOverrides);
    state.result = {
      ...result,
      route,
      routeLabel: ROUTE_LABELS[route],
      targetMonth,
      cutoffDate,
      planFileName: state.planFile.name,
      attendanceFileName: state.attendanceFile.name,
      annualFileName: state.annualFile?.name || "",
      referenceFileName: state.referenceFile?.name || "",
      annualComparison,
      annualSourceSheets: annualWorkbook || [],
      annualLedger: annualLedger || null,
      annualApplications: parsedAnnual?.rows || [],
      referenceComparison,
      evidenceOverrides,
      workflowOverrides,
      personnelChecks: personnelOverrides,
      personnelOverrides: (parsedReference?.personnelOverrides || []).map((item) => ({ ...item, sourceType: "evidence" })),
      managerRequests,
      analyzedAt: new Date().toISOString(),
      workforce,
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
      currentGrants: data.currentGrants || [],
      settlementGrants: data.settlementGrants || [],
      autoUseDates: data.autoUseDates || [],
      previousMonth: data.previousMonth || "",
      previousMonthFacts: data.previousMonthFacts || {},
    };
  } catch (error) {
    showToast(`${error.message}. 이번 분석은 저장된 이전 월 누적을 제외하고 계산합니다.`);
    return emptyLedger();
  }
}

function emptyLedger() {
  return { lotsByEmployee: {}, annualLeaveBefore: {}, currentGrants: [], settlementGrants: [], autoUseDates: [], previousMonth: "", previousMonthFacts: {} };
}

async function fileToWorkbookSheets(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, raw: false });
  return workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    repairWorksheetRange(sheet);
    return {
      sheetName,
      matrix: XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false,
      }),
    };
  }).filter((sheet) => sheet.matrix.some((row) => row.some((cell) => text(cell))));
}

function repairWorksheetRange(sheet) {
  if (!sheet || !window.XLSX?.utils) return;
  const cells = Object.keys(sheet).filter((key) => !key.startsWith("!") && /^[A-Z]+\d+$/.test(key));
  if (!cells.length) return;
  let minRow = Infinity, minCol = Infinity, maxRow = 0, maxCol = 0;
  for (const address of cells) {
    const cell = XLSX.utils.decode_cell(address);
    minRow = Math.min(minRow, cell.r); minCol = Math.min(minCol, cell.c);
    maxRow = Math.max(maxRow, cell.r); maxCol = Math.max(maxCol, cell.c);
  }
  // 시트가 B열부터 시작해도 실제 엑셀 열 위치를 유지하도록 A1부터 읽습니다.
  // 이전 방식은 열이 한 칸 밀리면서 사원명 대신 연락처가 이름으로 저장될 수 있었습니다.
  const actualRef = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
  const current = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
  if (!current || current.s.r > 0 || current.s.c > 0 || current.e.r < maxRow || current.e.c < maxCol) sheet["!ref"] = actualRef;
}

function parsePlan(sheets, targetMonth) {
  const candidates = [];
  const errors = [];
  for (const sheet of normalizeWorkbookInput(sheets)) {
    try {
      candidates.push(tryParsePlanSheet(sheet, targetMonth));
    } catch (error) {
      errors.push(`${sheet.sheetName}: ${error.message}`);
    }
  }
  if (!candidates.length) {
    throw new Error(`계획표에서 사번·이름·날짜 열을 찾지 못했습니다. 확인한 시트: ${normalizeWorkbookInput(sheets).map((sheet) => sheet.sheetName).join(", ") || "없음"}`);
  }
  candidates.sort((a, b) => b.rows.length - a.rows.length || b.dayColumns.size - a.dayColumns.size);
  return candidates[0];
}

function tryParsePlanSheet(sheet, targetMonth) {
  const matrix = sheet.matrix;
  const headerIndex = findPlanHeaderRow(matrix, targetMonth);
  if (headerIndex < 0) throw new Error("사번·이름·날짜 머리글 없음");

  const rawHeaders = matrix[headerIndex] || [];
  const headers = rawHeaders.map(normalizeHeader);
  const columns = {
    store: findHeaderIndex(headers, ["매장명", "매장", "점포", "점포명", "지점명", "근무지", "근무매장"]),
    employeeId: findHeaderIndex(headers, ["사번", "사원번호", "직원번호", "사원ID", "사번ID", "EMPLOYEEID", "EMPLOYEENO"]),
    name: findHeaderIndex(headers, ["이름", "성명", "사원명", "직원명", "매니저명"]),
    employment: findHeaderIndex(headers, ["재직상태", "근무상태", "재직여부"]),
  };

  if (columns.store < 0 && headers.length > 0) columns.store = 0;
  if (columns.employeeId < 0 && headers.length > 1) columns.employeeId = 1;
  if (columns.name < 0 && headers.length > 3) columns.name = 3;

  const dayColumns = new Map();
  rawHeaders.forEach((header, index) => {
    const day = parsePlanHeaderDay(header, targetMonth);
    if (day && !dayColumns.has(day)) dayColumns.set(day, index);
  });
  for (const nearbyIndex of [headerIndex - 1, headerIndex + 1]) {
    if (nearbyIndex < 0 || nearbyIndex >= matrix.length) continue;
    (matrix[nearbyIndex] || []).forEach((header, index) => {
      if (index < 6) return;
      const day = parsePlanHeaderDay(header, targetMonth);
      if (day && !dayColumns.has(day)) dayColumns.set(day, index);
    });
  }
  if (!dayColumns.size) throw new Error("01일~31일 날짜 열 없음");

  const rows = matrix.slice(headerIndex + 1).map((row, sourceIndex) => ({
    store: columns.store >= 0 ? text(row[columns.store]) : "",
    employeeId: normalizeEmployeeId(row[columns.employeeId]),
    name: columns.name >= 0 ? text(row[columns.name]) : "",
    employment: columns.employment >= 0 ? text(row[columns.employment]) : "",
    plans: Object.fromEntries([...dayColumns.entries()].map(([day, col]) => [day, text(row[col])])),
    rawRow: [...row],
    sourceIndex: sourceIndex + headerIndex + 2,
  })).filter((row) => looksLikeEmployeeId(row.employeeId) && row.name && (!row.employment || !row.employment.includes("퇴사")));

  if (!rows.length) throw new Error("유효한 사번 데이터 없음");
  return {
    rows,
    headerIndex,
    rawHeaders: [...rawHeaders],
    matrix,
    columns,
    dayColumns,
    detectedRoute: detectRouteFromPlan(rows),
    sheetName: sheet.sheetName,
  };
}

function parseAttendance(sheets, targetMonth) {
  const candidates = [];
  for (const sheet of normalizeWorkbookInput(sheets)) {
    try {
      const parsed = tryParseAttendanceSheet(sheet, targetMonth);
      if (parsed.rows.length) candidates.push(parsed);
    } catch {
      // 다른 시트 후보를 계속 확인합니다.
    }
  }
  if (!candidates.length) {
    throw new Error(`${targetMonth} 근태 기록 또는 사번·근무일자 머리글을 찾지 못했습니다. 통합문서의 모든 시트를 확인했습니다.`);
  }
  candidates.sort((a, b) => b.rows.length - a.rows.length);
  return candidates[0];
}

function tryParseAttendanceSheet(sheet, targetMonth) {
  const matrix = sheet.matrix;
  const headerIndex = findAttendanceHeaderRow(matrix, targetMonth);
  if (headerIndex < 0) throw new Error("사번·근무일자 머리글 없음");

  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const columns = {
    employeeId: findHeaderIndex(headers, ["사번", "사원번호", "직원번호", "사원ID", "사번ID", "EMPLOYEEID", "EMPLOYEENO"]),
    name: findHeaderIndex(headers, ["이름", "성명", "사원명", "직원명", "매니저명"]),
    date: findHeaderIndex(headers, ["근무일자", "근무일", "근태일자", "일자", "날짜"]),
    actualIn: findHeaderIndex(headers, ["(실제)출근시간", "실제출근시간", "출근시간", "출근시각", "출근"]),
    changedIn: findHeaderIndex(headers, ["(변경)출근시간", "변경출근시간", "수정출근시간", "인정출근시간"]),
    location: findHeaderIndex(headers, ["출근지점", "출근매장", "근무지점", "매장명", "점포명", "근무지"]),
    actualStatus: findHeaderIndex(headers, [
      "실제근태", "근태", "근태구분", "근태항목", "근태명", "근태코드",
      "근무구분", "실제근무구분", "(실제)근태", "처리근태",
    ]),
  };

  // 사용 중인 근태표의 고정 구조(A 이름, B 근무일자, C 출근시간, D 사번, E 퇴근시간) 보조 인식.
  if (columns.name < 0 && headers.length > 0) columns.name = 0;
  if (columns.date < 0 && headers.length > 1) columns.date = 1;
  if (columns.actualIn < 0 && headers.length > 2) columns.actualIn = 2;
  if (columns.employeeId < 0 && headers.length > 3) columns.employeeId = 3;

  if (columns.date < 0 || columns.employeeId < 0) throw new Error("날짜 또는 사번 열 없음");
  if (columns.actualIn < 0 && columns.changedIn < 0 && columns.actualStatus < 0) throw new Error("출근시간 또는 실제 근태 열 없음");

  const rows = matrix.slice(headerIndex + 1).map((row, sourceIndex) => {
    const date = parseDateCell(row[columns.date]);
    return {
      employeeId: normalizeEmployeeId(row[columns.employeeId]),
      name: columns.name >= 0 ? text(row[columns.name]) : "",
      date,
      actualIn: columns.actualIn >= 0 ? cleanClockValue(row[columns.actualIn]) : "",
      changedIn: columns.changedIn >= 0 ? cleanClockValue(row[columns.changedIn]) : "",
      location: columns.location >= 0 ? cleanPlaceholderValue(row[columns.location]) : "",
      actualStatus: columns.actualStatus >= 0 ? cleanPlaceholderValue(row[columns.actualStatus]) : "",
      sourceIndex: sourceIndex + headerIndex + 2,
    };
  }).filter((row) => looksLikeEmployeeId(row.employeeId) && row.date && row.date.startsWith(targetMonth));

  if (!rows.length) throw new Error(`${targetMonth} 유효 행 없음`);
  return {
    rows,
    headerIndex,
    matrix,
    columns,
    detectedRoute: detectRouteFromAttendance(rows),
    hasActualStatusColumn: columns.actualStatus >= 0,
    sheetName: sheet.sheetName,
  };
}


function emptyAnnualComparison() {
  return { rows: [], matchCount: 0, mismatchCount: 0, missingApplicationCount: 0, supplied: false };
}

function emptyReferenceComparison() {
  return { rows: [], mismatchCount: 0, matchCount: 0, supplied: false, sameMonth: false, summary: "비교 파일 미선택" };
}


async function fetchPersonnelChecks(month, route = "") {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) return { items: [], summary: {} };
  try {
    const response = await fetch(`/api/personnel-checks?month=${encodeURIComponent(month)}&route=${encodeURIComponent(route || "")}`, { cache: "no-store" });
    if (response.status === 401) return { items: [], summary: {} };
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "인력 확인 요청 조회 실패");
    return data;
  } catch (error) {
    console.warn(error);
    return { items: [], summary: {} };
  }
}

function normalizePersonnelStatus(value) {
  const raw = normalizeHeader(value);
  if (!raw || raw.includes("확인요청") || raw === "미처리") return "확인 요청";
  if (raw.includes("육아휴직")) return "육아휴직";
  if (raw.includes("휴직")) return "기타휴직";
  if (raw.includes("퇴사") || raw.includes("퇴직")) return "퇴사";
  if (raw.includes("경로이동") || raw.includes("이동") || raw.includes("전환")) return "경로이동";
  if (raw.includes("제외")) return "제외";
  if (raw.includes("재직") || raw.includes("포함") || raw.includes("유지") || raw.includes("정상")) return "재직·포함";
  return "확인 요청";
}

function mergePersonnelOverrides(serverItems, evidenceItems, month, route) {
  const map = new Map();
  const add = (raw, sourceType) => {
    const employeeId = normalizeEmployeeId(raw.employeeId || raw.employee_id);
    const itemRoute = raw.route || route;
    if (!employeeId || !["homeplus", "electroland"].includes(itemRoute)) return;
    const key = `${itemRoute}|${employeeId}`;
    const previous = map.get(key) || {};
    const status = normalizePersonnelStatus(raw.personnelStatus || raw.personnel_status);
    map.set(key, {
      ...previous,
      ...raw,
      month,
      route: itemRoute,
      routeLabel: ROUTE_LABELS[itemRoute],
      employeeId,
      employeeName: text(raw.employeeName || raw.employee_name || previous.employeeName),
      issueType: text(raw.issueType || raw.issue_type || previous.issueType || "인력 변동 직접 입력"),
      personnelStatus: status !== "확인 요청" || !previous.personnelStatus ? status : previous.personnelStatus,
      effectiveFrom: parseDateCell(raw.effectiveFrom || raw.effective_from) || previous.effectiveFrom || "",
      effectiveTo: parseDateCell(raw.effectiveTo || raw.effective_to) || previous.effectiveTo || "",
      destinationRoute: routeValue(raw.destinationRoute || raw.destination_route) || raw.destinationRoute || previous.destinationRoute || "",
      note: text(raw.note || previous.note),
      sourceType,
    });
  };
  for (const item of serverItems || []) add(item, item.sourceType || "server");
  for (const item of evidenceItems || []) add(item, "evidence");
  return [...map.values()].map((item) => ({ ...item, resolved: item.personnelStatus !== "확인 요청" }));
}

function personnelStatusApplies(item, month) {
  const monthStart = `${month}-01`;
  const [year, monthNumber] = month.split("-").map(Number);
  const monthEnd = `${month}-${String(new Date(year, monthNumber, 0).getDate()).padStart(2, "0")}`;
  if (item.effectiveFrom && item.effectiveFrom > monthEnd) return false;
  if (item.effectiveTo && item.effectiveTo < monthStart) return false;
  return true;
}

function personnelExcludedIds(items, month, route) {
  const excludedStatuses = new Set(["퇴사", "경로이동", "육아휴직", "기타휴직", "제외"]);
  return new Set((items || [])
    .filter((item) => {
      if (item.route !== route || !excludedStatuses.has(item.personnelStatus) || !personnelStatusApplies(item, month)) return false;
      if (item.personnelStatus === "경로이동" && item.destinationRoute && item.destinationRoute === route) return false;
      return true;
    })
    .map((item) => normalizeEmployeeId(item.employeeId))
    .filter(Boolean));
}

function parsePersonnelOverridesFromWorkbook(sheets, targetMonth, defaultRoute) {
  const candidates = normalizeWorkbookInput(sheets).filter((sheet) => /인력.*(변동|확인)|인원.*확인|재직.*현황|퇴사|휴직|경로.*이동/.test(normalizeHeader(sheet.sheetName)));
  const rows = [];
  for (const sheet of candidates) {
    const matrix = sheet.matrix || [];
    const headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
      findHeaderIndex(headers, ["사번", "사원번호"]) >= 0
      && findHeaderIndex(headers, ["처리구분", "인력상태", "재직상태", "변동구분"]) >= 0
    ));
    if (headerIndex < 0) continue;
    const headers = (matrix[headerIndex] || []).map(normalizeHeader);
    const columns = {
      route: findHeaderIndex(headers, ["경로", "현재경로"]),
      employeeId: findHeaderIndex(headers, ["사번", "사원번호"]),
      employeeName: findHeaderIndex(headers, ["이름", "성명"]),
      issueType: findHeaderIndex(headers, ["확인유형", "확인요청", "사유"]),
      status: findHeaderIndex(headers, ["처리구분", "인력상태", "재직상태", "변동구분"]),
      effectiveFrom: findHeaderIndex(headers, ["적용일", "변동일", "퇴사일", "이동일", "휴직시작일"]),
      effectiveTo: findHeaderIndex(headers, ["종료일", "복직일", "휴직종료일"]),
      destinationRoute: findHeaderIndex(headers, ["이동경로", "변경경로", "이동처"]),
      note: findHeaderIndex(headers, ["비고", "관리메모", "상세내용"]),
      regionalManager: findHeaderIndex(headers, ["지역장"]),
      manager: findHeaderIndex(headers, ["매니저"]),
      region: findHeaderIndex(headers, ["지역", "권역"]),
      storeName: findHeaderIndex(headers, ["매장명", "점포명", "매장"]),
    };
    for (const row of matrix.slice(headerIndex + 1)) {
      const employeeId = normalizeEmployeeId(row[columns.employeeId]);
      if (!looksLikeEmployeeId(employeeId)) continue;
      const status = normalizePersonnelStatus(row[columns.status]);
      if (status === "확인 요청" && !text(row[columns.note])) continue;
      rows.push({
        month: targetMonth,
        route: routeValue(columns.route >= 0 ? row[columns.route] : "") || defaultRoute,
        employeeId,
        employeeName: columns.employeeName >= 0 ? text(row[columns.employeeName]) : "",
        issueType: columns.issueType >= 0 ? text(row[columns.issueType]) : "인력 변동 직접 입력",
        personnelStatus: status,
        effectiveFrom: columns.effectiveFrom >= 0 ? parseDateCell(row[columns.effectiveFrom]) : "",
        effectiveTo: columns.effectiveTo >= 0 ? parseDateCell(row[columns.effectiveTo]) : "",
        destinationRoute: columns.destinationRoute >= 0 ? routeValue(row[columns.destinationRoute]) : "",
        note: columns.note >= 0 ? text(row[columns.note]) : "",
        regionalManager: columns.regionalManager >= 0 ? text(row[columns.regionalManager]) : "",
        manager: columns.manager >= 0 ? text(row[columns.manager]) : "",
        region: columns.region >= 0 ? text(row[columns.region]) : "",
        storeName: columns.storeName >= 0 ? text(row[columns.storeName]) : "",
      });
    }
  }
  return rows;
}

function parseAnnualApplications(sheets, targetMonth) {
  const candidates = [];
  for (const sheet of normalizeWorkbookInput(sheets)) {
    const matrix = sheet.matrix || [];
    const headerIndex = findFlexibleHeaderRow(matrix, (headers) => {
      const employee = findHeaderIndex(headers, ["사번", "사원번호", "직원번호", "사원ID"]);
      const start = findHeaderIndex(headers, ["휴가기간(시작)", "휴가시작일", "시작일", "휴가기간시작"]);
      const end = findHeaderIndex(headers, ["휴가기간(종료)", "휴가종료일", "종료일", "휴가기간종료"]);
      return employee >= 0 && start >= 0 && end >= 0;
    });
    if (headerIndex < 0) continue;
    const headers = (matrix[headerIndex] || []).map(normalizeHeader);
    const columns = {
      name: findHeaderIndex(headers, ["이름", "성명", "사원명"]),
      employeeId: findHeaderIndex(headers, ["사번", "사원번호", "직원번호", "사원ID"]),
      status: findHeaderIndex(headers, ["상태", "신청상태"]),
      applicationDate: findHeaderIndex(headers, ["신청일", "신청일자", "등록일"]),
      start: findHeaderIndex(headers, ["휴가기간(시작)", "휴가시작일", "시작일", "휴가기간시작"]),
      end: findHeaderIndex(headers, ["휴가기간(종료)", "휴가종료일", "종료일", "휴가기간종료"]),
      amount: findHeaderIndex(headers, ["사용갯수", "사용개수", "사용일수", "휴가일수", "일수"]),
      leaveType: findHeaderIndex(headers, ["휴가구분", "휴가종류", "구분"]),
      note: findHeaderIndex(headers, ["비고", "사유", "메모"]),
    };
    const rows = [];
    let sourceIndex = 0;
    for (const source of matrix.slice(headerIndex + 1)) {
      sourceIndex += 1;
      const employeeId = normalizeEmployeeId(source[columns.employeeId]);
      const startDateText = parseDateCell(source[columns.start]);
      const endDateText = parseDateCell(source[columns.end]);
      if (!looksLikeEmployeeId(employeeId) || !startDateText || !endDateText) continue;
      const startDate = new Date(`${startDateText}T00:00:00`);
      const endDate = new Date(`${endDateText}T00:00:00`);
      if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime()) || startDate > endDate) continue;
      const leaveType = columns.leaveType >= 0 ? text(source[columns.leaveType]) : "연차";
      const status = columns.status >= 0 ? text(source[columns.status]) : "";
      const rawAmountText = columns.amount >= 0 ? String(source[columns.amount] ?? "").trim() : "";
      const amountRaw = rawAmountText === "" ? null : Number(rawAmountText.replace(/[^0-9.\-]/g, ""));
      if (amountRaw !== null && (!Number.isFinite(amountRaw) || amountRaw <= 0)) continue;
      const dates = [];
      for (let cursor = new Date(startDate); cursor <= endDate; cursor.setDate(cursor.getDate() + 1)) dates.push(toISODate(cursor));
      const totalDays = amountRaw === null ? dates.length : amountRaw;
      if (!(totalDays > 0)) continue;
      const perDay = roundHalf(totalDays / Math.max(1, dates.length));
      for (const date of dates) {
        if (targetMonth && !date.startsWith(targetMonth)) continue;
        const normalizedType = normalizeHeader(leaveType);
        const requestedKind = normalizedType.includes("반차") || perDay === 0.5 ? "반차"
          : normalizedType.includes("연차") ? "연차" : leaveType || "기타휴가";
        rows.push({
          employeeId,
          name: columns.name >= 0 ? text(source[columns.name]) : "",
          date,
          leaveDate: date,
          requestedDays: perDay,
          days: perDay,
          requestedKind,
          applicationStatus: status,
          status,
          leaveType,
          applicationDate: columns.applicationDate >= 0 ? parseDateCell(source[columns.applicationDate]) : "",
          note: columns.note >= 0 ? text(source[columns.note]) : "",
          sourceIndex,
        });
      }
    }
    if (rows.length) candidates.push({ rows, sheetName: sheet.sheetName, supplied: true });
  }
  if (!candidates.length) throw new Error("연차신청현황에서 사번·휴가기간 시작/종료 열을 찾지 못했습니다.");
  candidates.sort((a, b) => b.rows.length - a.rows.length);
  return candidates[0];
}

function annualApplicationsToParsed(applications, targetMonth) {
  const rows = (applications || []).filter((item) => !targetMonth || String(item.leaveDate || item.date || "").startsWith(targetMonth)).map((item) => ({
    employeeId: normalizeEmployeeId(item.employeeId),
    name: text(item.employeeName || item.name),
    date: item.leaveDate || item.date || "",
    leaveDate: item.leaveDate || item.date || "",
    requestedDays: roundHalf(item.days || item.requestedDays),
    days: roundHalf(item.days || item.requestedDays),
    requestedKind: normalizeHeader(item.leaveType).includes("반차") || Number(item.days) === 0.5 ? "반차" : normalizeHeader(item.leaveType).includes("연차") ? "연차" : item.leaveType || "기타휴가",
    applicationStatus: item.status || item.applicationStatus || "",
    status: item.status || item.applicationStatus || "",
    leaveType: item.leaveType || "연차",
    applicationDate: item.applicationDate || "",
    note: item.note || "",
    sourceIndex: Number(item.sourceIndex || 0),
  }));
  return { rows, sheetName: "저장된 월별 승인·반려", supplied: rows.length > 0 };
}

function isApprovedAnnualApplication(row) {
  const status = normalizeHeader(row.applicationStatus || row.status);
  const type = normalizeHeader(row.leaveType || row.requestedKind);
  return (type.includes("연차") || type.includes("반차")) && (status.startsWith("승인") || status === "완료");
}

function mergeAnnualLeaveLedger(result, dashboard, annualComparison) {
  const ledgerById = new Map((dashboard?.employees || []).map((row) => [normalizeEmployeeId(row.employeeId), row]));
  const compareById = new Map();
  for (const row of annualComparison?.rows || []) {
    const id = normalizeEmployeeId(row.employeeId);
    if (!id) continue;
    if (!compareById.has(id)) compareById.set(id, []);
    compareById.get(id).push(row);
  }
  for (const summary of result.employeeSummaries || []) {
    const id = normalizeEmployeeId(summary.employeeId);
    const ledger = ledgerById.get(id);
    const rows = compareById.get(id) || [];
    const planned = roundHalf((summary.annualLeaveEvents || []).reduce((sum, item) => sum + Number(item.days || 0), 0));
    const approved = roundHalf(rows.filter(isApprovedAnnualApplication).reduce((sum, item) => sum + Number(item.requestedDays || 0), 0));
    const missing = roundHalf(rows.filter((item) => String(item.result || "").includes("신청내역 없음")).reduce((sum, item) => sum + annualLeaveValue(item.planStatus), 0));
    summary.currentAnnualLeave = approved;
    summary.annualLeaveUsed = approved;
    const savedCurrentApproved = ledger ? roundHalf(ledger.approvedCurrentMonth || 0) : 0;
    summary.cumulativeAnnualLeave = ledger ? roundHalf(Number(ledger.approvedUsed || 0) - savedCurrentApproved + approved) : summary.cumulativeAnnualLeave;
    summary.annualPlanned = planned;
    summary.annualApproved = approved;
    summary.annualMissingApplication = missing;
    summary.annualOpeningRemaining = ledger ? roundHalf(ledger.openingRemaining || 0) : "";
    summary.annualRemaining = ledger ? roundHalf(Number(ledger.remaining || 0) + savedCurrentApproved - approved) : "";
    summary.annualGranted = ledger ? roundHalf(ledger.granted || 0) : "";
    summary.annualCycleStart = ledger?.cycleStart || "";
    summary.annualCycleEnd = ledger?.cycleEnd || "";
    summary.annualGrantType = ledger?.grantType || "";
    summary.annualUnderOneYear = Boolean(ledger?.underOneYear);
    summary.annualFirstPromotionDate = ledger?.firstPromotionDate || "";
    summary.annualSecondPromotionDate = ledger?.secondPromotionDate || "";
  }
  result.annualRows = (result.employeeSummaries || []).filter((row) => Number(row.annualPlanned || 0) > 0 || Number(row.annualApproved || 0) > 0 || row.annualRemaining !== "");
  result.annualLeavePeople = (result.employeeSummaries || []).filter((row) => Number(row.annualApproved || 0) > 0).length;
}

function workforceMetaMap(workforce) {
  const map = new Map();
  for (const member of workforce?.members || []) {
    const employeeId = normalizeEmployeeId(member.employeeId ?? member.employee_id);
    if (!employeeId || map.has(employeeId)) continue;
    map.set(employeeId, {
      regionalManager: text(member.regionalManager ?? member.regional_manager),
      manager: text(member.manager),
      region: text(member.region2 ?? member.region1),
      store: text(member.storeName ?? member.store_name),
      name: text(member.employeeName ?? member.employee_name),
    });
  }
  return map;
}

function normalizeDashboardRegion(regionValue, storeValue) {
  const region = String(regionValue || "").replace(/\s+/g, "");
  const store = String(storeValue || "").replace(/\s+/g, "");
  if (region.includes("서울") || region.includes("경원") || region.includes("강원")) return "서울";
  if (region.includes("경인") || region.includes("경기") || region.includes("인천")) return "경인";
  if (region.includes("충청") || region.includes("대전") || region.includes("세종")) return "충청";
  if (region.includes("경북") || region.includes("대구")) return "경북";
  if (region.includes("경남") || region.includes("부산") || region.includes("울산")) return "경남";
  if (region.includes("전라") || region.includes("광주") || region.includes("전북") || region.includes("전남")) return "전라";
  if (/대전|세종|천안|청주|서산|아산|충주/.test(store)) return "충청";
  if (/대구|경산|경주|포항|구미|안동|영주/.test(store)) return "경북";
  if (/부산|울산|창원|김해|거제|양산|진주/.test(store)) return "경남";
  if (/광주|전주|순천|광양|목포|익산|군산|여수/.test(store)) return "전라";
  if (/인천|일산|운정|양주|남양주|경기|수원|용인|평택|안산|부천|김포/.test(store)) return "경인";
  return "서울";
}

function compareAnnualApplications(parsed, plan, attendance, targetMonth, cutoffDate, workforce, excludedPersonnelIds = new Set(), evidenceKeys = []) {
  const metaMap = workforceMetaMap(workforce);
  const planRows = choosePlanRows(plan.rows || [], new Map());
  const planById = new Map(planRows.map((row) => [normalizeEmployeeId(row.employeeId), row]));
  const attendanceMap = buildAttendanceMap(attendance?.rows || []);
  const evidenceSet = new Set((evidenceKeys || []).map(String));
  const requestKeys = new Set();
  const rows = [];
  const planActual = (employeeId, date) => {
    const raw = attendanceMap.get(`${employeeId}|${date}`) || attendanceMap.get(`${normalizeEmployeeId(employeeId)}|${date}`) || emptyAttendanceValue();
    return withEvidenceAttendance(raw, evidenceSet.has(`${normalizeEmployeeId(employeeId)}|${date}`));
  };
  const classify = ({ planStatus, requestedKind, applicationStatus, hasClockIn, hasApplication = true }) => {
    const plannedDays = annualLeaveValue(planStatus);
    const plannedAnnual = plannedDays > 0;
    const planKind = plannedDays === 0.5 ? "반차" : planStatus === "연차" ? "연차" : "";
    const samePlanAndRequest = plannedAnnual && hasApplication && requestedKind === planKind;

    // 승인 여부(대기/승인)는 계획·신청 일치 판정과 분리합니다.
    // 대기 상태라도 계획과 신청 종류가 같고 실제근태 조건이 맞으면 '동일'입니다.
    if (samePlanAndRequest && planKind === "연차" && !hasClockIn) {
      return { category: "동일", result: "계획 연차·신청 동일", needsReview: false, sortOrder: 1 };
    }
    // 오전·오후 반차는 실제 출근기록이 있어야 정상입니다.
    if (samePlanAndRequest && planKind === "반차" && hasClockIn) {
      return { category: "동일", result: "계획 반차·신청 동일", needsReview: false, sortOrder: 1 };
    }

    // 계획과 신청 종류가 다르면 출근 여부와 관계없이 주황색 '계획·신청 다름'입니다.
    // 예: 계획 오후반차 / 신청 연차 / 실제 출근.
    if (plannedAnnual && hasApplication && requestedKind !== planKind) {
      return { category: "계획·신청 다름", result: `계획 ${planStatus} / 신청서 ${requestedKind}`, needsReview: true, sortOrder: 2 };
    }

    // 계획 연차·반차인데 신청서가 없으면 빨간색 확인 요청입니다.
    if (plannedAnnual && !hasApplication) {
      return { category: "계획 연차·신청 없음", result: "계획 연차·반차 / 신청내역 없음", needsReview: true, sortOrder: 3 };
    }

    // 하루 연차인데 출근했거나, 반차인데 출근기록이 없으면 실제근태 확인이 필요합니다.
    if (samePlanAndRequest && ((planKind === "연차" && hasClockIn) || (planKind === "반차" && !hasClockIn))) {
      return {
        category: "출근 기록 확인",
        result: planKind === "연차" ? "계획 연차·신청 동일이나 출근기록 있음" : "계획 반차·신청 동일이나 출근기록 없음",
        needsReview: true,
        sortOrder: 4,
      };
    }

    if (planStatus === "공백" && hasApplication) {
      return { category: "계획·신청 다름", result: `계획 공백 / 신청서 ${requestedKind}`, needsReview: true, sortOrder: 5 };
    }
    if (hasApplication) {
      return { category: "계획·신청 다름", result: `계획 ${planStatus} / 신청서 ${requestedKind}`, needsReview: true, sortOrder: 5 };
    }
    return { category: "확인 필요", result: "연차 등록 확인 필요", needsReview: true, sortOrder: 6 };
  };

  for (const application of parsed.rows || []) {
    if (!["연차", "반차"].includes(application.requestedKind)) continue;
    const employeeId = normalizeEmployeeId(application.employeeId);
    if (excludedPersonnelIds.has(employeeId)) continue;
    const day = Number(application.date.slice(-2));
    requestKeys.add(`${employeeId}|${day}`);
    const planRow = planById.get(employeeId);
    const planStatus = planRow ? normalizePlanCode(planRow.plans?.[day]) : "공백";
    const attendanceValue = planActual(employeeId, application.date);
    const classification = classify({ planStatus, requestedKind: application.requestedKind, applicationStatus: application.applicationStatus, hasClockIn: attendanceValue.hasClockIn, hasApplication: true });
    const meta = metaMap.get(employeeId) || {};
    rows.push({
      date: application.date, regionalManager: meta.regionalManager || "", manager: meta.manager || "", region: meta.region || "",
      store: meta.store || planRow?.store || "", employeeId, name: meta.name || application.name || planRow?.name || "",
      requestedKind: application.requestedKind, requestedDays: application.requestedDays, planStatus,
      actualStatus: attendanceValue.hasClockIn ? "출근" : "미출근", applicationStatus: application.applicationStatus,
      leaveType: application.leaveType || application.requestedKind, applicationDate: application.applicationDate || "",
      sourceIndex: application.sourceIndex || 0, note: application.note || "", ...classification,
    });
  }

  const cutoffDay = Number(String(cutoffDate || `${targetMonth}-31`).slice(-2)) || 31;
  for (const planRow of planRows) {
    const employeeId = normalizeEmployeeId(planRow.employeeId);
    if (excludedPersonnelIds.has(employeeId)) continue;
    const meta = metaMap.get(employeeId) || {};
    for (let day = 1; day <= cutoffDay; day += 1) {
      const planStatus = normalizePlanCode(planRow.plans?.[day]);
      if (annualLeaveValue(planStatus) <= 0 || requestKeys.has(`${employeeId}|${day}`)) continue;
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const attendanceValue = planActual(employeeId, date);
      const classification = classify({ planStatus, requestedKind: "-", applicationStatus: "-", hasClockIn: attendanceValue.hasClockIn, hasApplication: false });
      rows.push({
        date, regionalManager: meta.regionalManager || "", manager: meta.manager || "", region: meta.region || "", store: meta.store || planRow.store || "",
        employeeId, name: meta.name || planRow.name || "", requestedKind: "-", requestedDays: 0, planStatus,
        actualStatus: attendanceValue.hasClockIn ? "출근" : "미출근", applicationStatus: "-", note: "", ...classification,
      });
    }
  }
  rows.sort((a, b) => a.sortOrder - b.sortOrder || normalizeDashboardRegion(a.region, a.store).localeCompare(normalizeDashboardRegion(b.region, b.store), "ko") || a.date.localeCompare(b.date) || a.store.localeCompare(b.store, "ko") || a.name.localeCompare(b.name, "ko"));
  return {
    supplied: true, rows,
    matchCount: rows.filter((row) => !row.needsReview).length,
    mismatchCount: rows.filter((row) => row.needsReview).length,
    missingApplicationCount: rows.filter((row) => row.category === "계획 연차·신청 없음").length,
    reviewCount: rows.filter((row) => row.needsReview).length,
  };
}

function parseReferenceFinalDate(value) {
  const raw = text(value);
  const digits = raw.replace(/\D/g, "");
  if (digits.length === 8 && /^20\d{6}$/.test(digits)) return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  return parseDateCell(value);
}

function parseReferenceFinalWorkbook(sheets, targetMonth = "", defaultRoute = "") {
  const normalizedSheets = normalizeWorkbookInput(sheets);
  const main = normalizedSheets.find((sheet) => /상담사\s*근태/.test(sheet.sheetName)) || normalizedSheets[0];
  if (!main) throw new Error("비교용 최종본의 상담사근태 시트를 찾지 못했습니다.");
  const matrix = main.matrix || [];
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
    findHeaderIndex(headers, ["사번", "사원번호"]) >= 0
    && findHeaderIndex(headers, ["성명", "이름"]) >= 0
  ));
  if (headerIndex < 0) throw new Error("비교용 최종본에서 사번·성명 머리글을 찾지 못했습니다.");
  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const employeeIdCol = findHeaderIndex(headers, ["사번", "사원번호"]);
  const nameCol = findHeaderIndex(headers, ["성명", "이름"]);
  const storeCol = findHeaderIndex(headers, ["매장명", "점포명", "매장"]);
  let dateRowIndex = -1;
  let dateColumns = new Map();
  for (let candidate = Math.max(0, headerIndex - 3); candidate < headerIndex; candidate += 1) {
    const found = new Map();
    (matrix[candidate] || []).forEach((cell, col) => {
      const date = parseReferenceFinalDate(cell);
      if (date) found.set(col, date);
    });
    if (found.size > dateColumns.size) {
      dateColumns = found;
      dateRowIndex = candidate;
    }
  }
  const monthFromSheet = String(main.sheetName).match(/(\d{1,2})월/)?.[1];
  const firstDate = [...dateColumns.values()][0] || "";
  const month = firstDate ? firstDate.slice(0, 7) : monthFromSheet ? `2026-${String(monthFromSheet).padStart(2, "0")}` : "";
  const values = new Map();
  const employees = new Set();
  for (const row of matrix.slice(headerIndex + 1)) {
    const employeeId = normalizeEmployeeId(row[employeeIdCol]);
    if (!looksLikeEmployeeId(employeeId)) continue;
    employees.add(employeeId);
    const name = nameCol >= 0 ? text(row[nameCol]) : "";
    const store = storeCol >= 0 ? text(row[storeCol]) : "";
    for (const [col, date] of dateColumns.entries()) {
      values.set(`${employeeId}|${date}`, { employeeId, name, store, date, value: text(row[col]) });
    }
  }
  return {
    mainSheetName: main.sheetName,
    month,
    dateRowIndex,
    values,
    employeeCount: employees.size,
    evidenceKeys: parseEvidenceOverrides(normalizedSheets),
    workflowOverrides: parseWorkflowOverrides(normalizedSheets),
    personnelOverrides: parsePersonnelOverridesFromWorkbook(normalizedSheets, targetMonth || month, defaultRoute),
    sheetNames: normalizedSheets.map((sheet) => sheet.sheetName),
  };
}


function parseFinalLeaveImportWorkbook(sheets, targetMonth, route) {
  const normalizedSheets = normalizeWorkbookInput(sheets);
  const main = normalizedSheets.find((sheet) => /상담사\s*근태|상담사.*근태/.test(String(sheet.sheetName || "")));
  if (!main) throw new Error("이전 최종본에서 상담사근태 시트를 찾지 못했습니다.");
  const matrix = main.matrix || [];
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
    findHeaderIndex(headers, ["사번", "사원번호"]) >= 0
    && findHeaderIndex(headers, ["성명", "이름"]) >= 0
  ));
  if (headerIndex < 0) throw new Error("상담사근태 시트에서 사번·성명 머리글을 찾지 못했습니다.");

  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const employeeIdCol = findHeaderIndex(headers, ["사번", "사원번호"]);
  const nameCol = findHeaderIndex(headers, ["성명", "이름"]);
  const storeCol = findHeaderIndex(headers, ["매장명", "점포명", "매장"]);
  // 과거 최종본의 수기 이월 개수는 부여 설정과 충돌할 수 있으므로 잔여 계산에 사용하지 않습니다.
  // 상담사근태의 날짜별 대체휴무·보상휴가 사용내역만 가져오고,
  // 잔여는 관리 화면에 저장한 발생일·사용 시작일·종료일 기준으로 다시 계산합니다.
  const openingSubstituteCol = -1;
  const openingCompensationCol = -1;
  const openingBalancePresent = false;
  const metadataColumns = {
    regionalManager: findHeaderIndex(headers, ["지역장"]),
    manager: findHeaderIndex(headers, ["매니저"]),
    region1: findHeaderIndex(headers, ["지역1", "1차지역"]),
    region2: findHeaderIndex(headers, ["지역2", "2차지역", "권역"]),
    storeCode: findHeaderIndex(headers, ["매장코드", "점포코드"]),
    portalId: findHeaderIndex(headers, ["포탈사번", "포털사번", "포탈ID", "포털ID"]),
    hireDate: findHeaderIndex(headers, ["제니엘입사일", "입사일"]),
    groupHireDate: findHeaderIndex(headers, ["고용승계일", "그룹입사일"]),
    note: findHeaderIndex(headers, ["비고", "휴/퇴사일", "휴퇴사일"]),
  };
  let dateColumns = new Map();
  for (let candidate = Math.max(0, headerIndex - 6); candidate < headerIndex; candidate += 1) {
    const found = new Map();
    (matrix[candidate] || []).forEach((cell, col) => {
      const date = parseReferenceFinalDate(cell);
      if (date && (!targetMonth || date.startsWith(targetMonth))) found.set(col, date);
    });
    if (found.size > dateColumns.size) dateColumns = found;
  }
  if (!dateColumns.size) throw new Error(`${targetMonth || "대상 월"} 날짜 열을 상담사근태 시트에서 찾지 못했습니다.`);

  const employees = new Map();
  for (const row of matrix.slice(headerIndex + 1)) {
    const employeeId = normalizeEmployeeId(row?.[employeeIdCol]);
    if (!looksLikeEmployeeId(employeeId)) continue;
    const metadata = {
      regionalManager: metadataColumns.regionalManager >= 0 ? text(row?.[metadataColumns.regionalManager]) : "",
      manager: metadataColumns.manager >= 0 ? text(row?.[metadataColumns.manager]) : "",
      region1: metadataColumns.region1 >= 0 ? text(row?.[metadataColumns.region1]) : "",
      region2: metadataColumns.region2 >= 0 ? text(row?.[metadataColumns.region2]) : "",
      storeCode: metadataColumns.storeCode >= 0 ? text(row?.[metadataColumns.storeCode]).replace(/,/g, "").replace(/\.0+$/, "") : "",
      portalId: metadataColumns.portalId >= 0 ? text(row?.[metadataColumns.portalId]).replace(/,/g, "").replace(/\.0+$/, "") : "",
      hireDate: metadataColumns.hireDate >= 0 ? parseDateCell(row?.[metadataColumns.hireDate]) : "",
      groupHireDate: metadataColumns.groupHireDate >= 0 ? parseDateCell(row?.[metadataColumns.groupHireDate]) : "",
      note: metadataColumns.note >= 0 ? text(row?.[metadataColumns.note]) : "",
    };
    const current = employees.get(employeeId) || {
      employeeId,
      name: nameCol >= 0 ? text(row?.[nameCol]) : "",
      store: storeCol >= 0 ? text(row?.[storeCol]) : "",
      ...metadata,
      basicDayoffDates: [],
      substituteEvents: [],
      compensationEvents: [],
      workedDates: [],
      dailyStatuses: [],
      importedOpeningBalancePresent: false,
      importedOpeningSubstitute: 0,
      importedOpeningCompensation: 0,
    };
    if (!current.name && nameCol >= 0) current.name = text(row?.[nameCol]);
    if (!current.store && storeCol >= 0) current.store = text(row?.[storeCol]);
    for (const [key, value] of Object.entries(metadata)) {
      if (!current[key] && value) current[key] = value;
    }

    for (const [col, date] of dateColumns.entries()) {
      const rawValue = text(row?.[col]);
      const normalized = normalizeActualCode(rawValue);
      const hasClockIn = finalImportWorkedValue(rawValue, normalized);
      const substituteDays = substitutePlanValue(normalized);
      const compensationDays = compensationPlanValue(normalized);
      if (normalized === "휴무") current.basicDayoffDates.push(date);
      if (substituteDays > 0) current.substituteEvents.push({
        date, days: substituteDays, source: "이전 최종본 상담사근태", planStatus: normalized,
      });
      if (compensationDays > 0) current.compensationEvents.push({
        date, days: compensationDays, source: "이전 최종본 상담사근태", planStatus: normalized,
      });
      if (hasClockIn) current.workedDates.push(date);
      current.dailyStatuses.push({
        date,
        planStatus: normalized || (rawValue ? rawValue : "공백"),
        hasClockIn,
        actualStatus: rawValue,
        importedFinal: true,
      });
    }
    employees.set(employeeId, current);
  }

  const baseAllowance = route === "homeplus" ? 6 : countWeekendDays(targetMonth);
  const employeeFacts = [...employees.values()].map((row) => {
    const substituteNeeded = roundHalf(row.substituteEvents.reduce((sum, event) => sum + Number(event.days || 0), 0));
    const compensationNeeded = roundHalf(row.compensationEvents.reduce((sum, event) => sum + Number(event.days || 0), 0));
    const basicDayoffUsed = roundHalf(new Set(row.basicDayoffDates).size);
    return {
      employeeId: row.employeeId,
      name: row.name,
      store: row.store,
      regionalManager: row.regionalManager || "",
      manager: row.manager || "",
      region1: row.region1 || "",
      region2: row.region2 || "",
      storeCode: row.storeCode || "",
      portalId: row.portalId || "",
      hireDate: row.hireDate || "",
      groupHireDate: row.groupHireDate || "",
      note: row.note || "",
      baseAllowanceRaw: baseAllowance,
      baseAllowance,
      basicDayoffUsed,
      explicitSubDayoffUsed: substituteNeeded,
      baseExcess: roundHalf(Math.max(0, basicDayoffUsed - baseAllowance)),
      substituteNeeded,
      compensationLeaveUsed: compensationNeeded,
      compensationNeeded,
      annualLeaveUsed: 0,
      substituteEvents: dedupeUsageEvents(row.substituteEvents),
      compensationEvents: dedupeUsageEvents(row.compensationEvents),
      annualLeaveEvents: [],
      workedDates: [...new Set(row.workedDates)].sort(),
      occurrenceSubstituteDates: [],
      occurrenceRestDays: 0,
      occurrenceRestAllowances: [],
      dailyStatuses: row.dailyStatuses,
      evidenceDates: [],
      importedOpeningBalancePresent: false,
      importedOpeningSubstitute: 0,
      importedOpeningCompensation: 0,
    };
  }).sort((a, b) => a.store.localeCompare(b.store, "ko") || a.name.localeCompare(b.name, "ko"));

  return {
    mainSheetName: main.sheetName,
    month: targetMonth,
    employeeFacts,
    employeeCount: employeeFacts.length,
    substituteEvents: employeeFacts.reduce((sum, row) => sum + row.substituteEvents.length, 0),
    substituteDays: roundHalf(employeeFacts.reduce((sum, row) => sum + row.substituteNeeded, 0)),
    compensationEvents: employeeFacts.reduce((sum, row) => sum + row.compensationEvents.length, 0),
    compensationDays: roundHalf(employeeFacts.reduce((sum, row) => sum + row.compensationNeeded, 0)),
    workedOnFirstDay: employeeFacts.filter((row) => row.workedDates.includes(`${targetMonth}-01`)).length,
    restedOnFirstDay: employeeFacts.filter((row) => !row.workedDates.includes(`${targetMonth}-01`)).length,
    rosterReadyCount: employeeFacts.filter((row) => row.hireDate || row.groupHireDate).length,
  };
}

function buildStackedHeaders(matrix, headerIndex, lookbackRows = 3) {
  const start = Math.max(0, headerIndex - Math.max(0, Number(lookbackRows) || 0));
  const maxCols = Math.max(0, ...matrix.slice(start, headerIndex + 1).map((row) => row?.length || 0));
  return Array.from({ length: maxCols }, (_, col) => matrix
    .slice(start, headerIndex + 1)
    .map((row) => normalizeHeader(row?.[col]))
    .filter(Boolean)
    .join(""));
}

function parseLegacyBalance(value) {
  const raw = text(value).replace(/,/g, "");
  const matched = raw.match(/-?\d+(?:\.\d+)?/);
  return Math.max(0, roundHalf(matched ? Number(matched[0]) : 0));
}

function finalImportWorkedValue(rawValue, normalizedValue) {
  const raw = text(rawValue).replace(/\s+/g, "");
  if (!raw || raw === "#NAME?" || raw === "미입력" || raw.includes("미입력")) return false;
  if (substitutePlanValue(normalizedValue) > 0 || compensationPlanValue(normalizedValue) > 0) return false;
  if (["휴무", "연차", "공가", "휴가", "무급휴가", "경조"].includes(normalizedValue)) return false;
  if (/^\d{1,2}:\d{2}(?::\d{2})?$/.test(raw)) return true;
  return comparableCode(normalizedValue) === "근무";
}

function dedupeUsageEvents(events) {
  const map = new Map();
  for (const event of events || []) {
    const key = `${event.date}|${event.planStatus || ""}`;
    if (!map.has(key)) map.set(key, { ...event, days: roundHalf(event.days) });
  }
  return [...map.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function findEvidenceDateColumn(headers) {
  return findHeaderIndex(headers, [
    "발생일", "근무일자", "일자", "날짜", "근태수정발생일", "근태수정필요일", "근태수정필요일자",
    "근태수정일", "근태수정필요일", "수정필요일", "수정일자",
  ]);
}

function findDateLikeColumn(matrix, startRow = 0) {
  const sampleRows = (matrix || []).slice(startRow, startRow + 40);
  const maxCols = Math.max(0, ...sampleRows.map((row) => row?.length || 0));
  let best = { col: -1, count: 0 };
  for (let col = 0; col < maxCols; col += 1) {
    const count = sampleRows.reduce((sum, row) => sum + (parseReferenceFinalDate(row?.[col]) ? 1 : 0), 0);
    if (count > best.count) best = { col, count };
  }
  return best.count >= 2 ? best.col : -1;
}

function parseEvidenceOverrides(sheets) {
  const keys = new Set();
  const candidates = (sheets || []).filter((sheet) => /출근\s*미등록|증빙/.test(String(sheet.sheetName || "")));
  // 시트명이 수기로 바뀐 경우에도 머리글을 찾아 읽을 수 있도록 전체 시트를 보조 탐색합니다.
  const scanSheets = candidates.length ? candidates : (sheets || []);

  for (const evidenceSheet of scanSheets) {
    const matrix = evidenceSheet.matrix || [];
    let headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
      findHeaderIndex(headers, ["제니엘사번", "사번", "사원번호"]) >= 0
      && findHeaderIndex(headers, ["증빙여부(O입력)", "증빙여부(O 입력)", "증빙여부", "출근증빙", "증빙"]) >= 0
    ));

    // 기존 최종본의 고정 양식(B~K)을 수기로 일부 변경한 경우를 위한 안전한 보조 탐색입니다.
    if (headerIndex < 0) {
      headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
        findHeaderIndex(headers, ["제니엘사번", "사번", "사원번호"]) >= 0
        && findHeaderIndex(headers, ["근태수정발생일", "근태수정필요일", "근태수정필요일자", "발생일", "근무일자", "일자", "날짜"]) >= 0
      ));
    }
    if (headerIndex < 0) continue;

    const headers = (matrix[headerIndex] || []).map(normalizeHeader);
    let employeeIdCol = findHeaderIndex(headers, ["제니엘사번", "사번", "사원번호"]);
    let dateCol = findEvidenceDateColumn(headers);
    let evidenceCol = findHeaderIndex(headers, ["증빙여부(O입력)", "증빙여부(O 입력)", "증빙여부", "출근증빙", "증빙"]);

    if (dateCol < 0) dateCol = findDateLikeColumn(matrix, headerIndex + 1);
    // 최종본 고정양식에서 머리글이 지워진 경우: G=사번, H=발생일, K=증빙여부.
    if (employeeIdCol < 0 && (matrix[headerIndex] || []).length >= 11) employeeIdCol = 6;
    if (dateCol < 0 && (matrix[headerIndex] || []).length >= 11) dateCol = 7;
    if (evidenceCol < 0 && (matrix[headerIndex] || []).length >= 11) evidenceCol = 10;
    if (employeeIdCol < 0 || dateCol < 0 || evidenceCol < 0) continue;

    for (const row of matrix.slice(headerIndex + 1)) {
      const employeeId = normalizeEmployeeId(row[employeeIdCol]);
      const date = parseReferenceFinalDate(row[dateCol]);
      const mark = String(row[evidenceCol] ?? "").trim().toUpperCase().replace(/\s+/g, "");
      const approved = ["O", "○", "ㅇ", "Y", "YES", "TRUE", "완료", "증빙완료"].includes(mark);
      if (approved && looksLikeEmployeeId(employeeId) && date) keys.add(`${employeeId}|${date}`);
    }
  }
  return [...keys];
}

function emptyWorkflowOverrides() {
  return { planMismatchCompletedKeys: [], dayoffCompletedIds: [], managerDeliveredIds: [] };
}

function isWorkflowApproved(value) {
  const mark = String(value ?? "").trim().toUpperCase().replace(/\s+/g, "");
  return ["O", "○", "ㅇ", "Y", "YES", "TRUE", "완료", "처리완료", "전달완료"].includes(mark);
}

function parseWorkflowOverrides(sheets) {
  const result = emptyWorkflowOverrides();
  const planSet = new Set(); const dayoffSet = new Set(); const managerSet = new Set();
  const scan = (pattern, markNames, handler) => {
    for (const sheet of (sheets || []).filter((item) => pattern.test(String(item.sheetName || "")))) {
      const matrix = sheet.matrix || [];
      const headerIndex = findFlexibleHeaderRow(matrix, (headers) => findHeaderIndex(headers, ["사번", "사원번호", "제니엘사번"]) >= 0 && findHeaderIndex(headers, markNames) >= 0);
      if (headerIndex < 0) continue;
      const headers = (matrix[headerIndex] || []).map(normalizeHeader);
      const idCol = findHeaderIndex(headers, ["사번", "사원번호", "제니엘사번"]);
      const dateCol = findHeaderIndex(headers, ["발생일", "근무일자", "일자", "날짜"]);
      const markCol = findHeaderIndex(headers, markNames);
      for (const row of matrix.slice(headerIndex + 1)) {
        const employeeId = normalizeEmployeeId(row[idCol]);
        if (!looksLikeEmployeeId(employeeId) || !isWorkflowApproved(row[markCol])) continue;
        handler(employeeId, dateCol >= 0 ? parseReferenceFinalDate(row[dateCol]) : "");
      }
    }
  };
  scan(/계획.*근태.*상이/, ["처리여부(O입력)", "처리여부(O 입력)", "처리체크", "완료여부", "처리완료"], (id, date) => { if (date) planSet.add(`${id}|${date}`); });
  scan(/휴무\s*초과/, ["처리여부(O입력)", "처리여부(O 입력)", "처리체크", "완료여부", "처리완료"], (id) => dayoffSet.add(id));
  scan(/매니저별.*이상.*근태/, ["전달체크(O입력)", "전달체크(O 입력)", "전달체크", "전달여부", "전달완료"], (id) => managerSet.add(id));
  result.planMismatchCompletedKeys = [...planSet]; result.dayoffCompletedIds = [...dayoffSet]; result.managerDeliveredIds = [...managerSet];
  return result;
}

function applyWorkflowOverrides(result, overrides = emptyWorkflowOverrides()) {
  const planSet = new Set(overrides.planMismatchCompletedKeys || []);
  const dayoffSet = new Set((overrides.dayoffCompletedIds || []).map(normalizeEmployeeId));
  for (const row of result.mismatchRows || []) row.resolved = planSet.has(`${normalizeEmployeeId(row.employeeId)}|${row.date}`);
  for (const row of result.employeeSummaries || []) row.dayoffResolved = dayoffSet.has(normalizeEmployeeId(row.employeeId));
  for (const row of result.employeeFacts || []) row.dayoffResolved = dayoffSet.has(normalizeEmployeeId(row.employeeId));
  result.workflowOverrides = overrides;
  return result;
}

function applyEvidenceOverrides(result, evidenceKeys = []) {
  const evidenceSet = new Set((evidenceKeys || []).map(String));
  result.evidenceOverrides = [...evidenceSet];
  if (!evidenceSet.size) return result;

  // K열 O는 해당 사번·일자의 실제 출근을 확정합니다.
  // 다만 계획이 휴무·연차 등인데 출근한 건은 계획 상이이므로 계속 남겨야 합니다.
  const evidenceDatesByEmployee = new Map();
  for (const key of evidenceSet) {
    const separator = key.indexOf("|");
    if (separator < 0) continue;
    const employeeId = normalizeEmployeeId(key.slice(0, separator));
    const date = key.slice(separator + 1);
    if (!employeeId || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
    if (!evidenceDatesByEmployee.has(employeeId)) evidenceDatesByEmployee.set(employeeId, new Set());
    evidenceDatesByEmployee.get(employeeId).add(date);
  }

  for (const collection of [result.employeeFacts || [], result.employeeSummaries || []]) {
    for (const row of collection) {
      const dates = evidenceDatesByEmployee.get(normalizeEmployeeId(row.employeeId));
      if (!dates?.size) continue;
      row.workedDates = [...new Set([...(row.workedDates || []), ...dates])].sort();
      row.evidenceDates = [...new Set([...(row.evidenceDates || []), ...dates])].sort();
      row.dailyStatuses = (row.dailyStatuses || []).map((daily) => dates.has(String(daily.date || ""))
        ? { ...daily, hasClockIn: true, actualStatus: "출근", evidenced: true }
        : daily);
    }
  }

  const keyOf = (row) => `${normalizeEmployeeId(row.employeeId)}|${row.date}`;
  const hasEvidence = (row) => evidenceSet.has(keyOf(row));
  const resolvesAttendanceMissing = (row) => hasEvidence(row)
    && ["missing_clock_in", "missing_plan_and_clock"].includes(row.issueType);
  const markEvidence = (row) => hasEvidence(row) ? { ...row, evidenced: true } : row;

  // O 입력은 해당 날짜의 실제 출근을 확정합니다.
  // 계획까지 비어 있던 건도 출근 증빙이 확인되면 별도 계획 미입력 오류로 다시 만들지 않습니다.
  result.missingRows = (result.missingRows || []).filter((row) => !resolvesAttendanceMissing(row)).map(markEvidence);
  // 휴무·연차 등인데 실제 출근한 건은 계획 상이이므로 유지합니다.
  result.unexpectedRows = (result.unexpectedRows || []).map(markEvidence);
  result.mismatchRows = (result.mismatchRows || []).flatMap((row) => {
    // 계획이 비어 있어도 실제 출근 또는 증빙 출근이 확인되면 정상 출근으로 인정하여 상이 목록에서도 제외합니다.
    if (row.issueType === "missing_plan" && (row.evidenced || row.actualIn || row.changedIn || row.actualStatus)) return [];
    if (!hasEvidence(row)) return [row];
    // 근무계획이 있는 단순 출근 미입력은 증빙 O로 완전히 해소합니다.
    if (row.issueType === "missing_clock_in") return [];
    // 계획과 출근이 모두 없던 건도 증빙 O가 들어오면 실제 출근으로 확정하고 상이 목록에서 제외합니다.
    if (row.issueType === "missing_plan_and_clock") return [];
    return [{ ...row, evidenced: true }];
  });
  result.missingPeople = uniquePeople(result.missingRows);
  result.unexpectedPeople = uniquePeople(result.unexpectedRows);
  result.mismatchPeople = uniquePeople(result.mismatchRows);
  return result;
}

function finalDisplayValue(planStatus, attendanceValue) {
  const planMissing = planStatus === "공백";
  const clockMissing = !attendanceValue?.hasClockIn;
  if (planMissing && clockMissing) return "출ㆍ계 미입력";
  const actual = normalizeActualCode(attendanceValue?.actualStatus);
  if (actual) return comparableCode(actual) === "근무" ? "출근" : actual;
  // 근무계획이 비어 있어도 실제 출근시간이 있으면 정상 출근으로 인정합니다.
  if (attendanceValue?.hasClockIn) return "출근";
  if (["근무", "근무A", "근무B", "근무C", "교육", "오전반차", "오후반차"].includes(planStatus)) return "출근 미입력";
  return planStatus;
}

function normalizeFinalCompareValue(value) {
  const raw = normalizeActualCode(value);
  if (!raw) return text(value);
  return comparableCode(raw) === "근무" ? "출근" : raw;
}

function compareReferenceFinal(reference, result, targetMonth, cutoffDate, workforce) {
  const rows = [];
  if (reference.month && reference.month !== targetMonth) {
    const currentNames = ["상담사근태", "계획&근태 상이 인원", "출근 미등록", "휴무 초과자", "전체 요약본", "매니저별 이상 근태", "해당 월 연차 등록 현황 및 일자", "연차 누적 현황", "근무 계획", "근태 RAW"];
    const normalizeName = (name) => String(name).replace(/^\d+월/, "O월");
    const currentSet = new Set(currentNames.map(normalizeName));
    const referenceSet = new Set(reference.sheetNames.map(normalizeName));
    const names = [...new Set([...currentSet, ...referenceSet])].sort();
    for (const name of names) {
      const generated = currentSet.has(name) ? "있음" : "없음";
      const ref = referenceSet.has(name) ? "있음" : "없음";
      rows.push({
        comparisonType: "시트 구조",
        employeeId: "",
        name: "",
        store: "",
        date: name,
        generatedValue: generated,
        referenceValue: ref,
        match: generated === ref,
        informational: true,
        result: generated === ref ? "구조 일치" : "구조 차이",
        reason: generated === ref ? "동일 역할 시트 존재" : "한쪽 파일에만 시트가 존재",
      });
    }
    rows.unshift({
      comparisonType: "대상 월",
      employeeId: "",
      name: "",
      store: "",
      date: "월 비교",
      generatedValue: targetMonth,
      referenceValue: reference.month || "확인 불가",
      match: false,
      informational: true,
      result: "월 불일치",
      reason: "직원별 날짜·근태 값은 비교하지 않고 시트 구조만 비교함",
    });
    return { supplied: true, sameMonth: false, rows, mismatchCount: rows.filter((row) => !row.match && !row.informational).length, matchCount: rows.filter((row) => row.match).length, summary: "월 불일치 · 구조 비교" };
  }

  const metaMap = workforceMetaMap(workforce);
  const attendanceById = groupBy(result.attendance.rows || [], (row) => row.employeeId);
  const selectedPlans = choosePlanRows(result.plan.rows || [], attendanceById);
  const attendanceMap = buildAttendanceMap(result.attendance.rows || []);
  const cutoffDay = Number(String(cutoffDate || `${targetMonth}-31`).slice(-2)) || 31;
  const generated = new Map();
  const evidenceSet = new Set(result.evidenceOverrides || []);
  for (const person of selectedPlans) {
    for (let day = 1; day <= cutoffDay; day += 1) {
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const key = `${normalizeEmployeeId(person.employeeId)}|${date}`;
      const planStatus = normalizePlanCode(person.plans?.[day]);
      const attendanceValue = attendanceMap.get(`${person.employeeId}|${date}`) || emptyAttendanceValue();
      generated.set(key, {
        employeeId: person.employeeId,
        name: person.name,
        store: person.store,
        date,
        value: evidenceSet.has(key) ? "출근" : finalDisplayValue(planStatus, attendanceValue),
      });
    }
  }
  // 비교 기준일 이후의 비교 최종본 값은 결과·요약에서 제외합니다.
  // 예: 기준일이 6월 28일이면 비교 파일에 들어 있는 6월 29~30일은 차이로 만들지 않습니다.
  const referenceKeys = [...reference.values.keys()].filter((key) => {
    const date = String(key).split("|")[1] || "";
    return /^\d{4}-\d{2}-\d{2}$/.test(date)
      && date.startsWith(`${targetMonth}-`)
      && date <= cutoffDate;
  });
  const completedPlanKeys = new Set(result.workflowOverrides?.planMismatchCompletedKeys || []);
  const keys = new Set([...generated.keys(), ...referenceKeys]);
  for (const key of keys) {
    // 계획&근태 상이 시트에서 O 처리한 사번·일자는 비교 차이와 전체 요약에 다시 올리지 않습니다.
    if (completedPlanKeys.has(key)) continue;
    const current = generated.get(key);
    const ref = reference.values.get(key);
    const generatedValue = normalizeFinalCompareValue(current?.value || "직원/일자 없음");
    const referenceValue = normalizeFinalCompareValue(ref?.value || "직원/일자 없음");
    if (generatedValue === referenceValue) continue;
    const employeeId = current?.employeeId || ref?.employeeId || key.split("|")[0];
    const meta = metaMap.get(employeeId) || {};
    rows.push({
      comparisonType: "사번+날짜",
      employeeId,
      name: meta.name || current?.name || ref?.name || "",
      store: meta.store || current?.store || ref?.store || "",
      date: current?.date || ref?.date || key.split("|")[1],
      generatedValue,
      referenceValue,
      match: false,
      informational: false,
      result: "불일치",
      reason: !current ? "자동 생성본에 직원·일자 없음" : !ref ? "비교 파일에 직원·일자 없음" : "같은 사번·날짜의 근태 값이 다름",
    });
  }
  rows.sort((a, b) => a.date.localeCompare(b.date) || a.store.localeCompare(b.store) || a.name.localeCompare(b.name));
  return { supplied: true, sameMonth: true, rows, mismatchCount: rows.length, matchCount: Math.max(0, keys.size - rows.length), summary: rows.length ? `${rows.length}건 불일치` : "전체 일치" };
}

function buildManagerRequests(result, annualComparison, workforce, targetMonth, workflowOverrides = emptyWorkflowOverrides()) {
  const metaMap = workforceMetaMap(workforce);
  const deliveredSet = new Set((workflowOverrides.managerDeliveredIds || []).map(normalizeEmployeeId));
  const groups = new Map();
  const add = (employeeId, fallback, message) => {
    const id = normalizeEmployeeId(employeeId);
    if (!id || !message) return;
    if (!groups.has(id)) groups.set(id, { fallback, issues: [] });
    const group = groups.get(id); if (!group.issues.includes(message)) group.issues.push(message);
  };
  const monthNo = Number(targetMonth.slice(5, 7));
  for (const row of result.mismatchRows || []) {
    if (row.issueType === "missing_plan" || row.resolved) continue;
    const day = Number(String(row.date || "").slice(-2));
    let message;
    if (row.issueType === "missing_plan_and_clock") message = `${monthNo}월 ${day}일 근무계획·출근기록 모두 미입력`;
    else if (row.result === "근무인데 출근기록 없음") message = `${monthNo}월 ${day}일 출근기록 미입력`;
    else if (row.result === "휴무·휴가인데 출근기록 있음") message = `${monthNo}월 ${day}일 ${row.planStatus}이나 출근 기록 있음`;
    else if (String(row.reason || "").includes("인력·매장매칭") || String(row.reason || "").includes("사번 없음")) message = `${monthNo}월 인력·매장매칭 확인 필요(${row.reason})`;
    else message = `${monthNo}월 ${day}일 계획 ${row.planStatus} / 실제 ${row.actualStatus || row.clockStatus || "미기재"}`;
    add(row.employeeId, row, message);
  }
  for (const summary of result.employeeSummaries || []) {
    if (summary.dayoffResolved) continue;
    if (Number(summary.dayoffReplacementShortage || 0) > 0) add(summary.employeeId, summary, `${monthNo}월 휴무초과 ${formatDays(summary.dayoffReplacementShortage)}일 미대체`);
    if (Number(summary.shortage || 0) > 0) add(summary.employeeId, summary, `${monthNo}월 대체휴무 ${formatDays(summary.shortage)} 초과 사용`);
    if (Number(summary.compensationShortage || 0) > 0) add(summary.employeeId, summary, `${monthNo}월 보상휴가 ${formatDays(summary.compensationShortage)} 초과 사용`);
  }
  for (const row of annualComparison?.rows || []) {
    if (!row.needsReview) continue;
    const day = Number(String(row.date || "").slice(-2));
    add(row.employeeId, row, `${monthNo}월 ${day}일 연차 확인 요청(${row.result})`);
  }
  return [...groups.entries()].map(([employeeId, group]) => {
    const meta = metaMap.get(employeeId) || {}; const fallback = group.fallback || {};
    const manager = meta.manager || fallback.manager || ""; const store = meta.store || fallback.store || ""; const name = meta.name || fallback.name || "";
    const delivered = deliveredSet.has(employeeId);
    return {
      regionalManager: meta.regionalManager || "", manager, region: meta.region || "", store, employeeId, name,
      issueCount: group.issues.length, issueText: group.issues.join(" / "), delivered,
      status: delivered ? "전달 완료" : "미전달",
      message: `${manager} 매니저님, ${store} ${name} 상담사의 ${group.issues.join(", ")} 항목이 확인됩니다. 근무계획·근태·연차 신청내역 확인 후 수정 바랍니다.`,
    };
  }).sort((a, b) => a.regionalManager.localeCompare(b.regionalManager, "ko") || a.manager.localeCompare(b.manager, "ko") || a.store.localeCompare(b.store, "ko") || a.name.localeCompare(b.name, "ko"));
}

function normalizeWorkbookInput(value) {
  if (!Array.isArray(value)) return [];
  if (value.length && Array.isArray(value[0])) return [{ sheetName: "첫 번째 시트", matrix: value }];
  return value;
}

function findFlexibleHeaderRow(matrix, predicate) {
  const limit = Math.min(matrix.length, 60);
  for (let index = 0; index < limit; index += 1) {
    const headers = (matrix[index] || []).map(normalizeHeader);
    if (predicate(headers)) return index;
  }
  return -1;
}

function findPlanHeaderRow(matrix, targetMonth) {
  const limit = Math.min(matrix.length, 100);
  let best = { index: -1, score: -1 };
  for (let index = 0; index < limit; index += 1) {
    const row = matrix[index] || [];
    const headers = row.map(normalizeHeader);
    const employeeId = findHeaderIndex(headers, ["사번", "사원번호", "직원번호", "사원ID", "사번ID", "EMPLOYEEID", "EMPLOYEENO"]);
    const name = findHeaderIndex(headers, ["이름", "성명", "사원명", "직원명", "매니저명"]);
    const store = findHeaderIndex(headers, ["매장명", "매장", "점포", "점포명", "지점명", "근무지", "근무매장"]);
    let dayCount = row.filter((cell) => parsePlanHeaderDay(cell, targetMonth)).length;
    if (index > 0) dayCount = Math.max(dayCount, (matrix[index - 1] || []).filter((cell) => parsePlanHeaderDay(cell, targetMonth)).length);
    if (index + 1 < matrix.length) dayCount = Math.max(dayCount, (matrix[index + 1] || []).filter((cell) => parsePlanHeaderDay(cell, targetMonth)).length);

    const knownLayoutMatches = matrix.slice(index + 1, index + 12).filter((dataRow) => (
      looksLikeEmployeeId(dataRow?.[1]) && text(dataRow?.[3])
    )).length;
    const score = (employeeId >= 0 ? 20 : 0) + (name >= 0 ? 15 : 0) + (store >= 0 ? 5 : 0) + Math.min(dayCount, 31) + knownLayoutMatches * 2;
    const valid = (employeeId >= 0 && name >= 0 && dayCount >= 1) || (dayCount >= 5 && knownLayoutMatches >= 2);
    if (valid && score > best.score) best = { index, score };
  }
  return best.index;
}

function findAttendanceHeaderRow(matrix, targetMonth) {
  const limit = Math.min(matrix.length, 100);
  let best = { index: -1, score: -1 };
  for (let index = 0; index < limit; index += 1) {
    const headers = (matrix[index] || []).map(normalizeHeader);
    const employeeId = findHeaderIndex(headers, ["사번", "사원번호", "직원번호", "사원ID", "사번ID", "EMPLOYEEID", "EMPLOYEENO"]);
    const date = findHeaderIndex(headers, ["근무일자", "근무일", "근태일자", "일자", "날짜"]);
    const clock = findHeaderIndex(headers, ["(실제)출근시간", "실제출근시간", "출근시간", "출근시각", "출근"]);
    const knownLayoutMatches = matrix.slice(index + 1, index + 15).filter((dataRow) => {
      const parsedDate = parseDateCell(dataRow?.[1]);
      return looksLikeEmployeeId(dataRow?.[3]) && parsedDate && (!targetMonth || parsedDate.startsWith(targetMonth));
    }).length;
    const score = (employeeId >= 0 ? 20 : 0) + (date >= 0 ? 20 : 0) + (clock >= 0 ? 5 : 0) + knownLayoutMatches * 3;
    const valid = (employeeId >= 0 && date >= 0) || knownLayoutMatches >= 2;
    if (valid && score > best.score) best = { index, score };
  }
  return best.index;
}

function parsePlanHeaderDay(value, targetMonth) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    if (!targetMonth || `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}` === targetMonth) return value.getDate();
    return null;
  }
  const raw = text(value);
  if (!raw) return null;
  let match = raw.match(/^(?:0?)([1-9]|[12]\d|3[01])\s*일$/);
  if (match) return Number(match[1]);
  match = raw.match(/^(?:\d{2,4}[.\-/년]\s*)?(\d{1,2})[.\-/월]\s*(\d{1,2})(?:일)?$/);
  if (match) return Number(match[2]);
  match = raw.match(/^\d{4}-\d{1,2}-(\d{1,2})/);
  if (match) return Number(match[1]);
  if (/^\d{1,2}$/.test(raw)) {
    const numberValue = Number(raw);
    return numberValue >= 1 && numberValue <= 31 ? numberValue : null;
  }
  return null;
}

function compareAttendance({ plan, attendance, route, targetMonth, cutoffDate, ledger, evidenceKeys = [] }) {
  const evidenceSet = new Set((evidenceKeys || []).map(String));
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
  const mismatchMap = new Map();
  const employeeSummaries = [];
  const employeeFacts = [];

  const addMismatch = (row) => {
    const key = `${row.employeeId}|${row.date}`;
    if (!mismatchMap.has(key)) mismatchMap.set(key, row);
  };

  for (const person of selectedPlans) {
    const basicDayoffDates = [];
    const explicitSubstituteEvents = [];
    const occurrenceRestAllowances = [];
    const occurrenceSubstituteDates = [];
    const compensationEvents = [];
    const annualLeaveEvents = [];
    const workedDates = [];
    const dailyStatuses = [];
    const evidenceDates = [];

    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const planCode = normalizePlanCode(person.plans[day]);
      const evidenceKey = `${normalizeEmployeeId(person.employeeId)}|${date}`;
      const attendanceValue = withEvidenceAttendance(attendanceMap.get(`${person.employeeId}|${date}`) || emptyAttendanceValue(), evidenceSet.has(evidenceKey));
      const occurrenceDays = settlementOccurrenceDays(ledger, person.employeeId, date, cutoffDate);
      const compensationRestDays = settlementCompensationRestDays(ledger, person.employeeId, date, cutoffDate, attendanceValue.hasClockIn);
      const displayedStatus = finalDisplayValue(planCode, attendanceValue);
      dailyStatuses.push({
        date,
        planStatus: planCode,
        hasClockIn: Boolean(attendanceValue.hasClockIn),
        actualStatus: attendanceValue.actualStatus || "",
        evidenced: Boolean(attendanceValue.evidenced),
      });
      if (attendanceValue.evidenced) evidenceDates.push(date);
      // 발생일 대체휴무는 해당 날짜의 근무계획·실제 출근 여부와 무관하게 동일 부여합니다.
      // 발생일 당일 포함 이전 입사자만 대상이며, 발생일 때문에 기본 휴무 가능 수량을 추가하지 않습니다.
      if (occurrenceDays > 0) {
        occurrenceSubstituteDates.push(date);
      }
      if (compensationRestDays > 0) {
        occurrenceRestAllowances.push({
          date, days: compensationRestDays, source: "보상휴가 발생일 미출근 기본휴무 추가", planStatus: displayedStatus,
        });
      }
      // 휴무 사용량은 계획표 원본이 아니라 최종 일별 표시값을 기준으로 계산합니다.
      // 계획이 휴무여도 실제 출근기록이 있으면 출근으로 보며 휴무 사용에서 제외합니다.
      if (displayedStatus === "휴무") basicDayoffDates.push(date);
      const substituteDays = substitutePlanValue(displayedStatus);
      if (substituteDays > 0) explicitSubstituteEvents.push({ date, days: substituteDays, source: "표기 대체휴무", planStatus: displayedStatus });
      const compensationDays = compensationPlanValue(displayedStatus);
      if (compensationDays > 0) compensationEvents.push({ date, days: compensationDays, source: "표기 보상휴가", planStatus: displayedStatus });
      const annualDays = annualLeaveValue(planCode);
      if (annualDays > 0) annualLeaveEvents.push({ date, days: annualDays, planStatus: planCode });
      if (attendanceValue.hasClockIn) workedDates.push(date);
    }

    // 대체휴무 발생일은 기본휴무를 가감하지 않습니다.
    // 보상휴가 발생일에 미출근한 대상자만 발생일 1건당 기본휴무를 1일 추가합니다.
    const occurrenceRestDays = roundHalf(occurrenceRestAllowances.reduce((sum, row) => sum + Number(row.days || 0), 0));
    const personBaseAllowance = roundHalf(baseAllowance + occurrenceRestDays);
    const baseExcessEvents = basicDayoffDates.slice(Math.max(0, Math.floor(personBaseAllowance))).map((date) => ({
      date,
      days: 1,
      source: "기본 휴무 초과",
      planStatus: "휴무",
    }));
    // 기본 휴무 초과와 대체휴무 사용은 별도 관리합니다.
    // 발생일의 출근·휴무 여부는 기본 휴무 기준과 대체휴무 부여 여부에 영향을 주지 않습니다.
    const substituteEvents = [...explicitSubstituteEvents].sort((a, b) => a.date.localeCompare(b.date));

    for (let day = 1; day <= lastDay; day += 1) {
      const planStatus = normalizePlanCode(person.plans[day]);
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const dateObject = new Date(`${date}T00:00:00`);
      const evidenceKey = `${normalizeEmployeeId(person.employeeId)}|${date}`;
      const attendanceValue = withEvidenceAttendance(attendanceMap.get(`${person.employeeId}|${date}`) || emptyAttendanceValue(), evidenceSet.has(evidenceKey));
      let hasPrimaryIssue = false;

      if (planStatus === "공백" && attendanceValue.hasClockIn) {
        // 계획이 비어 있어도 실제 출근시간·변경 출근시간·실제근태 출근이 확인되면 정상 출근으로 인정합니다.
        // 출근 미등록, 계획&근태 상이, 전체 요약본, 매니저별 이상 근태에는 넣지 않습니다.
        // 일반 불일치 비교도 건너뛰도록 처리 완료 상태로 표시합니다.
        hasPrimaryIssue = true;
      } else if (planStatus === "공백" && !attendanceValue.hasClockIn) {
        const row = makeIssueRow({
          issueType: "missing_plan_and_clock",
          missingType: "출ㆍ계 미입력",
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: "근무계획·출근기록 없음",
          reason: "근무계획과 실제 출근시간·변경 출근시간이 모두 없음",
        });
        missingRows.push(row);
        addMismatch(makeMismatchRow({
          ...row,
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: "검토 필요",
          reason: row.reason,
        }));
        hasPrimaryIssue = true;
      } else if (REQUIRED_CLOCK_PLANS.has(planStatus) && !attendanceValue.hasClockIn) {
        const row = makeIssueRow({
          issueType: "missing_clock_in",
          missingType: "출근 미입력",
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: "근무인데 출근기록 없음",
          reason: `${planStatus} 계획이나 실제 출근시간과 변경 출근시간이 모두 없음`,
        });
        missingRows.push(row);
        addMismatch(makeMismatchRow({ ...row, route, person, date, dateObject, planStatus, attendanceValue, result: row.result, reason: row.reason }));
        hasPrimaryIssue = true;
      }

      if (UNEXPECTED_CLOCK_PLANS.has(planStatus) && attendanceValue.hasClockIn) {
        const row = makeIssueRow({
          issueType: "unexpected_clock_in",
          route,
          person,
          date,
          dateObject,
          planStatus,
          attendanceValue,
          result: "휴무·휴가인데 출근기록 있음",
          reason: `${planStatus} 계획이나 출근기록이 있음`,
        });
        unexpectedRows.push(row);
        addMismatch(makeMismatchRow({ ...row, route, person, date, dateObject, planStatus, attendanceValue, result: row.result, reason: row.reason }));
        hasPrimaryIssue = true;
      }

      if (!hasPrimaryIssue) {
        const mismatch = evaluateMismatch(planStatus, attendanceValue, attendance.hasActualStatusColumn);
        if (mismatch) {
          addMismatch(makeMismatchRow({
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
    }

    const basicDayoffUsed = roundHalf(basicDayoffDates.length);
    const explicitSubDayoffUsed = roundHalf(explicitSubstituteEvents.reduce((sum, event) => sum + event.days, 0));
    const autoSubstituteUsed = 0;
    const compensationLeaveUsed = roundHalf(compensationEvents.reduce((sum, event) => sum + event.days, 0));
    const baseExcess = roundHalf(Math.max(0, basicDayoffUsed - personBaseAllowance));
    const substituteNeeded = roundHalf(explicitSubDayoffUsed);
    const compensationNeeded = compensationLeaveUsed;
    const annualLeaveUsed = roundHalf(annualLeaveEvents.reduce((sum, event) => sum + event.days, 0));
    const openingCarryover = calculateOpeningCarryover(
      ledger.lotsByEmployee?.[person.employeeId] || [],
      targetMonth,
    );
    const preview = calculatePreviewLedger({
      employeeId: person.employeeId,
      substituteEvents,
      compensationEvents,
      dayoffReplacementEvents: baseExcessEvents,
      workedDates,
      occurrenceSubstituteDates,
      targetMonth,
      cutoffDate,
      ledger,
    });
    const previousFact = ledger.previousMonthFacts?.[normalizeEmployeeId(person.employeeId)] || {};
    const combinedAvailable = roundHalf(Number(preview.availableSubstitute || 0) + Number(preview.availableCompensation || 0));
    const combinedLeaveUsed = roundHalf(substituteNeeded + compensationLeaveUsed + Number(preview.dayoffReplacementUsed || 0));
    const combinedShortage = roundHalf(Number(preview.shortage || 0) + Number(preview.compensationShortage || 0) + Number(preview.dayoffReplacementShortage || 0));
    const combinedRemaining = roundHalf(Number(preview.remainingSubstitute || 0) + Number(preview.remainingCompensation || 0));
    const cumulativeAnnualLeave = roundHalf(Number(ledger.annualLeaveBefore[person.employeeId] || 0) + annualLeaveUsed);
    const judgment = buildDayoffJudgment({
      baseAllowance: personBaseAllowance,
      basicDayoffUsed,
      explicitSubDayoffUsed,
      baseExcess,
      baseExcessEvents,
      substituteNeeded,
      remainingSubstitute: preview.remainingSubstitute,
      expiredSubstitute: preview.expiredSubstitute,
      shortage: preview.shortage,
    });
    const compensationJudgment = buildCompensationJudgment({
      compensationNeeded,
      remainingCompensation: preview.remainingCompensation,
      expiredCompensation: preview.expiredCompensation,
      compensationShortage: preview.compensationShortage,
    });

    const fact = {
      route,
      routeLabel: ROUTE_LABELS[route],
      store: person.store,
      employeeId: person.employeeId,
      name: person.name,
      baseAllowanceRaw: baseAllowance,
      baseAllowance: personBaseAllowance,
      basicDayoffUsed,
      explicitSubDayoffUsed,
      autoSubstituteUsed,
      occurrenceRestDays,
      occurrenceRestAllowances,
      compensationLeaveUsed,
      baseExcess,
      baseExcessEvents,
      substituteNeeded,
      compensationNeeded,
      annualLeaveUsed,
      substituteEvents,
      compensationEvents,
      annualLeaveEvents,
      workedDates,
      evidenceDates: [...new Set(evidenceDates)].sort(),
      dailyStatuses,
      occurrenceSubstituteDates: [...new Set(occurrenceSubstituteDates)].sort(),
      duplicatePlanNote: person.duplicatePlanNote || "",
    };
    employeeFacts.push(fact);
    employeeSummaries.push({
      ...fact,
      openingCarryoverSubstitute: openingCarryover.substitute,
      openingCarryoverCompensation: openingCarryover.compensation,
      openingCarryoverTotal: openingCarryover.total,
      availableSubstitute: preview.availableSubstitute,
      substituteApplied: preview.substituteApplied,
      remainingSubstitute: preview.remainingSubstitute,
      expiredSubstitute: preview.expiredSubstitute,
      shortage: preview.shortage,
      substituteShortageDates: preview.substituteShortageDates || [],
      availableCompensation: preview.availableCompensation,
      compensationApplied: preview.compensationApplied,
      remainingCompensation: preview.remainingCompensation,
      expiredCompensation: preview.expiredCompensation,
      compensationShortage: preview.compensationShortage,
      compensationShortageDates: preview.compensationShortageDates || [],
      dayoffReplacementUsed: preview.dayoffReplacementUsed || 0,
      dayoffReplacementSubstituteUsed: preview.dayoffReplacementSubstituteUsed || 0,
      dayoffReplacementCompensationUsed: preview.dayoffReplacementCompensationUsed || 0,
      dayoffReplacementShortage: preview.dayoffReplacementShortage || 0,
      dayoffReplacementShortageDates: preview.dayoffReplacementShortageDates || [],
      previousMonth: ledger.previousMonth || "",
      priorDayoffExcess: roundHalf(previousFact.baseExcess || 0),
      priorDayoffReplacementUsed: roundHalf(previousFact.dayoffReplacementUsed || 0),
      priorDayoffReplacementShortage: roundHalf(previousFact.dayoffReplacementShortage || 0),
      combinedAvailable,
      combinedLeaveUsed,
      combinedRemaining,
      combinedShortage,
      currentAnnualLeave: annualLeaveUsed,
      cumulativeAnnualLeave,
      judgment,
      compensationJudgment,
    });
  }

  const mismatchRows = [...mismatchMap.values()];
  missingRows.sort(issueSort);
  unexpectedRows.sort(issueSort);
  mismatchRows.sort(issueSort);
  sortSummaries(employeeSummaries);

  const diagnostics = buildDiagnostics({ route, plan, attendance, matchRate, planIds, attendanceIds, baseAllowance });

  return assembleResultCollections({
    missingRows,
    unexpectedRows,
    mismatchRows,
    employeeSummaries,
    employeeFacts,
    plan,
    attendance,
    planPeople: planIds.size,
    attendancePeople: attendanceIds.size,
    matchedPeople: matchedIds.length,
    matchRate,
    baseAllowance,
    diagnostics,
  });
}

function appendWorkforceMatchingIssues(result, workforce, plan, route, targetMonth, excludedPersonnelIds = new Set()) {
  const members = (workforce?.members || []).filter((row) => row.route === route);
  const workforceById = new Map();
  for (const member of members) {
    const id = normalizeEmployeeId(member.employeeId || member.employee_id);
    if (id && !workforceById.has(id)) workforceById.set(id, member);
  }
  const planById = new Map();
  for (const row of plan?.rows || []) {
    const id = normalizeEmployeeId(row.employeeId);
    if (id && !planById.has(id)) planById.set(id, row);
  }
  const existing = new Set((result.mismatchRows || []).map((row) => `${normalizeEmployeeId(row.employeeId)}|${row.date}|${row.reason}`));
  const date = `${targetMonth}-01`;
  const dateObject = new Date(`${date}T00:00:00`);
  const weekday = WEEKDAY_LABELS[dateObject.getDay()];
  const add = (row) => {
    const key = `${normalizeEmployeeId(row.employeeId)}|${row.date}|${row.reason}`;
    if (existing.has(key)) return;
    existing.add(key);
    result.mismatchRows.push(row);
  };

  for (const [employeeId, member] of workforceById.entries()) {
    if (excludedPersonnelIds.has(employeeId)) continue;
    if (planById.has(employeeId)) continue;
    add({
      route, routeLabel: ROUTE_LABELS[route], store: member.storeName || member.store_name || "",
      employeeId, name: member.employeeName || member.employee_name || "", date, weekday,
      planStatus: "미등록", actualStatus: "", actualIn: "", changedIn: "", clockStatus: "미기록",
      result: "검토 필요", reason: "인력·매장매칭에는 있으나 근무계획에 사번 없음", duplicatePlanNote: "",
    });
  }
  for (const [employeeId, person] of planById.entries()) {
    if (excludedPersonnelIds.has(employeeId)) continue;
    if (workforceById.has(employeeId)) continue;
    add({
      route, routeLabel: ROUTE_LABELS[route], store: person.store || "",
      employeeId, name: person.name || "", date, weekday,
      planStatus: normalizePlanCode(person.plans?.[1]), actualStatus: "", actualIn: "", changedIn: "", clockStatus: "미기록",
      result: "검토 필요", reason: "근무계획에는 있으나 인력·매장매칭에 사번 없음", duplicatePlanNote: "",
    });
  }
  result.mismatchRows.sort(issueSort);
  result.mismatchPeople = uniquePeople(result.mismatchRows);
}

function assembleResultCollections(base) {
  const excessRows = base.employeeSummaries.filter((row) => Number(row.baseExcess || 0) > 0 || Number(row.priorDayoffExcess || 0) > 0);
  return {
    ...base,
    missingPeople: uniquePeople(base.missingRows), unexpectedPeople: uniquePeople(base.unexpectedRows), mismatchPeople: uniquePeople(base.mismatchRows),
    dayoffExcessRows: excessRows, balanceRows: [...base.employeeSummaries],
    shortageRows: base.employeeSummaries.filter((row) => row.shortage > 0 || row.compensationShortage > 0 || row.dayoffReplacementShortage > 0 || row.priorDayoffReplacementShortage > 0),
    annualRows: base.employeeSummaries.filter((row) => row.currentAnnualLeave > 0 || row.cumulativeAnnualLeave > 0),
    dayoffExcessPeople: excessRows.length,
    substituteShortagePeople: base.employeeSummaries.filter((row) => row.shortage > 0).length,
    compensationShortagePeople: base.employeeSummaries.filter((row) => row.compensationShortage > 0).length,
    annualLeavePeople: base.employeeSummaries.filter((row) => row.currentAnnualLeave > 0).length,
  };
}

function settlementOccurrenceDays(ledger, employeeId, date, cutoffDate) {
  if (!date || !cutoffDate || date > cutoffDate) return 0;
  const normalizedId = normalizeEmployeeId(employeeId);
  return roundHalf((ledger?.settlementGrants || [])
    .filter((grant) => (grant.grantType || "substitute") === "substitute")
    .filter((grant) => String(grant.occurrenceDate || "") === date)
    .filter((grant) => (grant.eligibleEmployeeIds || []).map(normalizeEmployeeId).includes(normalizedId))
    .reduce((sum, grant) => sum + Number(grant.grantedDays || 0), 0));
}

function settlementCompensationRestDays(ledger, employeeId, date, cutoffDate, hasClockIn) {
  if (!date || !cutoffDate || date > cutoffDate || hasClockIn) return 0;
  const normalizedId = normalizeEmployeeId(employeeId);
  return roundHalf((ledger?.settlementGrants || [])
    .filter((grant) => grant.grantType === "compensation")
    .filter((grant) => String(grant.occurrenceDate || "") === date)
    .filter((grant) => (grant.eligibleEmployeeIds || []).map(normalizeEmployeeId).includes(normalizedId))
    .reduce((sum) => sum + 1, 0));
}

function calculateOpeningCarryover(lots, targetMonth) {
  const monthStart = `${targetMonth}-01`;
  const [year, month] = String(targetMonth).split("-").map(Number);
  const previousDate = new Date(year, month - 1, 0);
  const previousMonthEnd = `${previousDate.getFullYear()}-${String(previousDate.getMonth() + 1).padStart(2, "0")}-${String(previousDate.getDate()).padStart(2, "0")}`;
  const eligible = (lots || []).filter((lot) => Number(lot?.remaining || 0) > 0
    && String(lot?.validFrom || "") <= previousMonthEnd
    && String(lot?.validTo || "") >= monthStart);
  const substitute = roundHalf(eligible
    .filter((lot) => (lot.grantType || "substitute") === "substitute")
    .reduce((sum, lot) => sum + Number(lot.remaining || 0), 0));
  const compensation = roundHalf(eligible
    .filter((lot) => lot.grantType === "compensation")
    .reduce((sum, lot) => sum + Number(lot.remaining || 0), 0));
  return { substitute, compensation, total: roundHalf(substitute + compensation) };
}

function calculatePreviewLedger({ employeeId, substituteEvents, compensationEvents, dayoffReplacementEvents = [], workedDates, occurrenceSubstituteDates, targetMonth, cutoffDate, ledger }) {
  const lots = (ledger.lotsByEmployee[employeeId] || []).map((lot) => ({
    grantId: lot.grantId, grantType: lot.grantType || "substitute", grantMonth: lot.grantMonth,
    occurrenceDate: lot.occurrenceDate || "", settlementMode: Boolean(lot.settlementMode),
    validFrom: lot.validFrom, validTo: lot.validTo, remaining: roundHalf(lot.remaining),
  }));
  const existingGrantIds = new Set(lots.map((lot) => String(lot.grantId || "")));
  for (const grant of ledger.currentGrants || []) {
    if (grant.grantMonth !== targetMonth || existingGrantIds.has(String(grant.id || "")) || !isCurrentGrantEligible(grant, employeeId, workedDates)) continue;
    lots.push({ grantId: grant.id, grantType: grant.grantType || "substitute", grantMonth: grant.grantMonth, occurrenceDate: grant.occurrenceDate || "", settlementMode: Boolean(grant.settlementMode), validFrom: grant.validFrom, validTo: grant.validTo, remaining: roundHalf(grant.grantedDays) });
    existingGrantIds.add(String(grant.id || ""));
  }

  // 1) 계획표에 직접 표기한 대체휴무·보상휴가를 먼저 차감합니다.
  const explicitPools = consumePreviewCombinedPools(lots, substituteEvents, compensationEvents, targetMonth);
  // 2) 기본 휴무 초과분은 남은 대체휴무를 우선 사용하고, 부족하면 보상휴가로 자동 대체합니다.
  const replacementPools = consumePreviewCombinedPools(lots, dayoffReplacementEvents, [], targetMonth);
  return {
    availableSubstitute: explicitPools.substitute.available,
    substituteApplied: explicitPools.substitute.applied,
    remainingSubstitute: replacementPools.substitute.remaining,
    expiredSubstitute: replacementPools.substitute.expired,
    shortage: explicitPools.substitute.shortage,
    substituteShortageDates: explicitPools.substitute.shortageDates,
    availableCompensation: explicitPools.compensation.available,
    compensationApplied: explicitPools.compensation.applied,
    remainingCompensation: replacementPools.compensation.remaining,
    expiredCompensation: replacementPools.compensation.expired,
    compensationShortage: explicitPools.compensation.shortage,
    compensationShortageDates: explicitPools.compensation.shortageDates,
    dayoffReplacementUsed: roundHalf(replacementPools.substitute.applied + replacementPools.compensation.applied),
    dayoffReplacementSubstituteUsed: roundHalf(replacementPools.substitute.applied),
    dayoffReplacementCompensationUsed: roundHalf(replacementPools.compensation.applied),
    dayoffReplacementShortage: roundHalf(replacementPools.substitute.shortage),
    dayoffReplacementShortageDates: replacementPools.substitute.shortageDates || [],
  };
}

function isCurrentGrantEligible(grant, employeeId, workedDates) {
  const normalizedId = normalizeEmployeeId(employeeId);
  if (grant.grantScope === "employee" && normalizedId !== normalizeEmployeeId(grant.employeeId)) return false;
  if (grant.grantScope !== "employee" && (grant.excludedEmployeeIds || []).map(normalizeEmployeeId).includes(normalizedId)) return false;
  if (grant.settlementMode && grant.grantType === "substitute") return true;
  if (grant.settlementMode && grant.grantType === "compensation") return (workedDates || []).includes(grant.occurrenceDate || grant.criterionDate);
  if (grant.eligibilityMode === "worked_on_date") return (workedDates || []).includes(grant.criterionDate);
  return true;
}

function consumePreviewCombinedPools(lots, substituteEvents, compensationEvents, targetMonth) {
  const monthStart = `${targetMonth}-01`;
  const monthEnd = endOfMonth(targetMonth);
  const nextMonthStart = startOfNextMonth(targetMonth);
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
  const shortageDatesByOrigin = { substitute: new Set(), compensation: new Set() };
  const events = [
    ...(substituteEvents || []).map((event) => ({ ...event, originType: "substitute" })),
    ...(compensationEvents || []).map((event) => ({ ...event, originType: "compensation" })),
  ].filter((event) => event?.date && Number(event.days) > 0)
    .sort((a, b) => a.date.localeCompare(b.date));

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
      lot.remaining = roundHalf(lot.remaining - used);
      need = roundHalf(need - used);
      appliedByType[lot.grantType] = roundHalf((appliedByType[lot.grantType] || 0) + used);
    }
    const eventShortage = roundHalf(Math.max(0, need));
    shortageByOrigin[event.originType] = roundHalf((shortageByOrigin[event.originType] || 0) + eventShortage);
    if (eventShortage > 0) shortageDatesByOrigin[event.originType].add(String(event.date));
  }

  const resultFor = (grantType, shortage) => ({
    available: availableByType[grantType],
    applied: appliedByType[grantType],
    remaining: roundHalf(lots
      .filter((lot) => lot.grantType === grantType && lot.remaining > 0 && lot.validTo >= nextMonthStart)
      .reduce((sum, lot) => sum + lot.remaining, 0)),
    expired: roundHalf(lots
      .filter((lot) => lot.grantType === grantType && lot.remaining > 0 && lot.validTo >= monthStart && lot.validTo < nextMonthStart)
      .reduce((sum, lot) => sum + lot.remaining, 0)),
    shortage: roundHalf(shortage),
    shortageDates: [...shortageDatesByOrigin[grantType]].sort(),
  });
  return {
    substitute: resultFor("substitute", shortageByOrigin.substitute),
    compensation: resultFor("compensation", shortageByOrigin.compensation),
  };
}

function buildAttendanceMap(rows) {
  const map = new Map();
  for (const row of rows) {
    const key = `${row.employeeId}|${row.date}`;
    const existing = map.get(key) || emptyAttendanceValue();
    const actualIn = cleanClockValue(row.actualIn);
    const changedIn = cleanClockValue(row.changedIn);
    const actualStatus = cleanPlaceholderValue(row.actualStatus);
    const actualStatusIsWork = comparableCode(normalizeActualCode(actualStatus)) === "근무";
    map.set(key, {
      // 출근시간뿐 아니라 실제근태가 출근·근무·정상으로 기록된 경우도 출근으로 인정합니다.
      hasClockIn: existing.hasClockIn || Boolean(actualIn || changedIn) || actualStatusIsWork,
      actualIn: actualIn || existing.actualIn,
      changedIn: changedIn || existing.changedIn,
      location: cleanPlaceholderValue(row.location) || existing.location,
      actualStatus: actualStatus || existing.actualStatus,
    });
  }
  return map;
}

function withEvidenceAttendance(attendanceValue, evidenced) {
  if (!evidenced) return attendanceValue;
  return {
    ...attendanceValue,
    hasClockIn: true,
    // 원본 실제근태가 '미입력'으로 남아 있어도 K열 O가 최종 확정값입니다.
    actualStatus: "출근",
    changedIn: attendanceValue.changedIn || attendanceValue.actualIn || "증빙",
    evidenced: true,
  };
}

function emptyAttendanceValue() {
  return { hasClockIn: false, actualIn: "", changedIn: "", location: "", actualStatus: "" };
}

function makeIssueRow({ issueType, missingType = "", route, person, date, dateObject, planStatus, attendanceValue, result, reason }) {
  return {
    issueType,
    missingType,
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
    evidenced: Boolean(attendanceValue.evidenced),
    result,
    reason,
    duplicatePlanNote: person.duplicatePlanNote || "",
  };
}

function makeMismatchRow({ issueType = "", missingType = "", route, person, date, dateObject, planStatus, attendanceValue, result, reason }) {
  return {
    issueType,
    missingType,
    route,
    routeLabel: ROUTE_LABELS[route],
    store: person.store,
    employeeId: person.employeeId,
    name: person.name,
    date,
    weekday: WEEKDAY_LABELS[dateObject.getDay()],
    planStatus,
    actualStatus: attendanceValue.actualStatus || (attendanceValue.hasClockIn ? "출근" : "미기재"),
    actualIn: attendanceValue.actualIn || "",
    changedIn: attendanceValue.changedIn || "",
    clockStatus: attendanceValue.hasClockIn ? (attendanceValue.changedIn || attendanceValue.actualIn || "출근") : "미기록",
    evidenced: Boolean(attendanceValue.evidenced),
    result,
    reason,
    duplicatePlanNote: person.duplicatePlanNote || "",
  };
}

function evaluateMismatch(planStatus, attendanceValue, hasActualStatusColumn) {
  if (planStatus === "기타") {
    return { result: "검토 필요", reason: "기타는 자동 단정하지 않고 계획·실제 근태 확인이 필요함" };
  }
  if (!VALID_PLAN_CODES.has(planStatus)) {
    return { result: "검토 필요", reason: `등록되지 않은 계획 코드 ‘${planStatus}’ 확인 필요` };
  }

  const actualStatusRaw = attendanceValue.actualStatus;
  if (["오전반차", "오후반차"].includes(planStatus)) {
    if (!attendanceValue.hasClockIn) return null;
    if (!hasActualStatusColumn || !text(actualStatusRaw)) return null;
    const actualStatus = normalizeActualCode(actualStatusRaw);
    if (["근무", "근무A", "근무B", "근무C", "오전반차", "오후반차", "반일근무"].includes(actualStatus)) return null;
    return { result: "계획·실제 근태 불일치", reason: `계획 ${planStatus} / 실제 ${actualStatusRaw}` };
  }

  if (!hasActualStatusColumn || !text(actualStatusRaw)) return null;
  const actualStatus = normalizeActualCode(actualStatusRaw);
  if (!actualStatus) return null;
  if (comparableCode(planStatus) === comparableCode(actualStatus)) return null;
  return { result: "계획·실제 근태 불일치", reason: `계획 ${planStatus} / 실제 ${actualStatusRaw}` };
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
  return !status || status === "546" ? "공백" : status;
}

function normalizeActualCode(value) {
  const raw = normalizeStatus(value);
  if (!raw || raw === "공백") return "";
  if (raw.includes("오전반차") || raw.includes("반차(오전)")) return "오전반차";
  if (raw.includes("오후반차") || raw.includes("반차(오후)")) return "오후반차";
  // 대체휴일·보상휴가의 0.5일 표기는 일반 반일근무보다 먼저 구분해야 합니다.
  if (raw.includes("대체") && (raw.includes("0.5") || raw.includes("반일"))) return "대체휴일(0.5일)";
  if (raw.includes("대체")) return "대체휴일(1일)";
  if (raw.includes("보상") && (raw.includes("0.5") || raw.includes("반일"))) return "보상휴가(0.5일)";
  if (raw.includes("보상")) return "보상휴가(1일)";
  if (raw.includes("0.5") || raw.includes("반일근무") || raw === "반차") return "반일근무";
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
  if (["오전반차", "오후반차", "반일근무"].includes(code)) return "반일근무";
  return code;
}

function substitutePlanValue(planCode) {
  if (planCode === "대체휴일(0.5일)") return 0.5;
  if (planCode === "대체휴일(1일)") return 1;
  return 0;
}

function compensationPlanValue(planCode) {
  if (planCode === "보상휴가(0.5일)") return 0.5;
  if (planCode === "보상휴가(1일)") return 1;
  return 0;
}

function annualLeaveValue(planCode) {
  if (planCode === "연차") return 1;
  if (["오전반차", "오후반차"].includes(planCode)) return 0.5;
  return 0;
}

function buildDayoffJudgment({ baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess, substituteNeeded, remainingSubstitute, expiredSubstitute, shortage }) {
  const parts = [];
  if (baseExcess > 0) parts.push(`휴무 개수 초과 ${formatDays(baseExcess)}`);
  else parts.push(`기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)}`);
  if (shortage > 0) parts.push(`대체휴무 ${formatDays(shortage)} 초과 사용`);
  else if (substituteNeeded > 0) parts.push(`대체휴무 ${formatDays(substituteNeeded)} 사용 · 잔여 ${formatDays(remainingSubstitute)}`);
  else parts.push(`대체휴무 사용 없음 · 잔여 ${formatDays(remainingSubstitute)}`);
  if (expiredSubstitute > 0) parts.push(`${formatDays(expiredSubstitute)} 만료`);
  return parts.join(" · ");
}

function buildCompensationJudgment({ compensationNeeded, remainingCompensation, expiredCompensation, compensationShortage }) {
  if (compensationShortage > 0) return `보상휴가 ${formatDays(compensationShortage)} 초과 사용`;
  if (compensationNeeded > 0) {
    const expired = expiredCompensation > 0 ? ` · ${formatDays(expiredCompensation)} 만료` : "";
    return `보상휴가 ${formatDays(compensationNeeded)} 사용 · 잔여 ${formatDays(remainingCompensation)}${expired}`;
  }
  if (expiredCompensation > 0) return `미사용 보상휴가 ${formatDays(expiredCompensation)} 만료`;
  return `보상휴가 사용 없음 · 잔여 ${formatDays(remainingCompensation)}`;
}

function buildDiagnostics({ route, plan, attendance, matchRate, planIds, attendanceIds, baseAllowance }) {
  const messages = [`${ROUTE_LABELS[route]} 경로 기본 휴무 기준: ${baseAllowance}일`];
  if (plan.detectedRoute && plan.detectedRoute !== route) messages.push(`계획표 내용은 ‘${ROUTE_LABELS[plan.detectedRoute]} 경로’로 감지되었습니다.`);
  if (attendance.detectedRoute && attendance.detectedRoute !== route) messages.push(`근태표 내용은 ‘${ROUTE_LABELS[attendance.detectedRoute]} 경로’로 감지되었습니다.`);
  if (matchRate < 50) messages.push(`사번 매칭률이 ${matchRate}%로 낮습니다. 서로 다른 경로 또는 월 파일인지 확인해 주세요.`);
  if (planIds.size && attendanceIds.size && matchRate >= 50) messages.push(`사번 ${[...planIds].filter((id) => attendanceIds.has(id)).length}명이 정상 매칭되었습니다.`);
  if (!attendance.hasActualStatusColumn) messages.push("실제 근태 열이 없어도 출근시간 기준으로 근무 미출근·휴무 출근을 계획 불일치에 함께 표시합니다. 세부 근태명 비교는 실제 근태 열이 있을 때만 적용합니다.");
  messages.push("오전반차·오후반차·0.5일 근무는 출근기록이 있으면 정상이며, 출근기록이 전혀 없을 때만 미출근으로 표시합니다.");
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) messages.push("관리자 로그인 전에는 이전 월 대체휴무·보상휴가 잔여와 누적 연차를 제외하고 계산합니다.");
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
  const managers = [...new Set((result.managerRequests || []).map((row) => row.manager).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
  $("#managerFilter").innerHTML = `<option value="">전체 매니저</option>${managers.map((manager) => `<option>${escapeHtml(manager)}</option>`).join("")}`;
  $("#searchInput").value = "";
  $("#managerFilter").value = "";
  switchResultTab("missing");
}

function refreshResultMetrics() {
  const result = state.result;
  $("#missingCount").textContent = number(result.missingRows.length);
  $("#unexpectedCount").textContent = number(result.unexpectedRows.length);
  $("#mismatchCount").textContent = number(result.mismatchRows.length);
  $("#dayoffExcessPeople").textContent = number(result.dayoffExcessPeople);
  $("#substituteShortagePeople").textContent = number(result.substituteShortagePeople);
  $("#compensationShortagePeople").textContent = number(result.compensationShortagePeople);
  $("#annualLeavePeople").textContent = number(result.annualLeavePeople);
  $("#matchRate").textContent = `${result.matchRate}%`;
  $("#annualCompareMismatch").textContent = number(result.annualComparison?.reviewCount ?? result.annualComparison?.mismatchCount ?? 0);
  $("#referenceCompareMismatch").textContent = number(result.referenceComparison?.mismatchCount || 0);
  $("#missingTabCount").textContent = result.missingRows.length;
  $("#unexpectedTabCount").textContent = result.unexpectedRows.length;
  $("#mismatchTabCount").textContent = result.mismatchRows.length;
  $("#excessTabCount").textContent = result.dayoffExcessRows.length;
  $("#balanceTabCount").textContent = result.balanceRows.length;
  $("#shortageTabCount").textContent = result.shortageRows.length;
  $("#annualTabCount").textContent = result.annualRows.length;
  $("#annualCompareTabCount").textContent = result.annualComparison?.rows?.length || 0;
  $("#managerRequestTabCount").textContent = result.managerRequests?.length || 0;
  $("#referenceCompareTabCount").textContent = result.referenceComparison?.rows?.length || 0;
  $("#personnelCheckTabCount").textContent = (result.personnelChecks || []).filter((item) => !item.resolved).length;
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
        { label: "당월 자동 대체", render: (row) => formatDays(row.dayoffReplacementUsed || 0) },
        { label: "당월 미대체", render: (row) => `<span class="${Number(row.dayoffReplacementShortage || 0) > 0 ? "status-pill" : "success-pill"}">${formatDays(row.dayoffReplacementShortage || 0)}</span>` },
        { label: "전월 초과", render: (row) => formatDays(row.priorDayoffExcess || 0) },
        { label: "전월 대체", render: (row) => formatDays(row.priorDayoffReplacementUsed || 0) },
        { label: "판정", className: "message-cell", render: (row) => Number(row.dayoffReplacementShortage || 0) > 0 || Number(row.priorDayoffReplacementShortage || 0) > 0 ? '<span class="status-pill">휴무초과 확인 요청</span>' : '<span class="success-pill">휴무초과 대체 완료</span>' },
      ],
    },
    balance: {
      rows: state.result.balanceRows,
      columns: [
        ...commonSummary,
        { label: "대체 가용", render: (row) => formatDays(row.availableSubstitute) },
        { label: "대체 사용", render: (row) => formatDays(row.substituteNeeded) },
        { label: "대체 잔여", render: (row) => formatDays(row.remainingSubstitute) },
        { label: "대체 초과", render: (row) => formatDays(row.shortage) },
        { label: "보상 가용", render: (row) => formatDays(row.availableCompensation) },
        { label: "보상 사용", render: (row) => formatDays(row.compensationNeeded) },
        { label: "보상 잔여", render: (row) => formatDays(row.remainingCompensation) },
        { label: "보상 초과", render: (row) => formatDays(row.compensationShortage) },
        { label: "판정", className: "message-cell", render: renderCombinedJudgment },
      ],
    },
    shortage: {
      rows: state.result.shortageRows,
      columns: [
        ...commonSummary,
        { label: "대체휴무 필요", render: (row) => formatDays(row.substituteNeeded) },
        { label: "대체휴무 적용", render: (row) => formatDays(row.substituteApplied) },
        { label: "대체휴무 초과", render: (row) => `<span class="${row.shortage > 0 ? "status-pill" : "neutral-pill"}">${formatDays(row.shortage)}</span>` },
        { label: "보상휴가 필요", render: (row) => formatDays(row.compensationNeeded) },
        { label: "보상휴가 적용", render: (row) => formatDays(row.compensationApplied) },
        { label: "보상휴가 초과", render: (row) => `<span class="${row.compensationShortage > 0 ? "status-pill" : "neutral-pill"}">${formatDays(row.compensationShortage)}</span>` },
        { label: "판정", className: "message-cell", render: renderCombinedJudgment },
      ],
    },
    annual: {
      rows: state.result.annualRows,
      columns: [
        ...commonSummary,
        { label: "구분", render: (row) => escapeHtml(row.annualUnderOneYear ? "1년 미만·월차" : row.annualGrantType || "연차") },
        { label: "사용기간", render: (row) => escapeHtml([row.annualCycleStart, row.annualCycleEnd].filter(Boolean).join(" ~ ") || "-") },
        { label: "발생", render: (row) => row.annualGranted === "" ? "-" : formatDays(row.annualGranted) },
        { label: "당월 계획", render: (row) => formatDays(row.annualPlanned || 0) },
        { label: "당월 승인", render: (row) => `<strong>${formatDays(row.annualApproved || 0)}</strong>` },
        { label: "연차 미신청", render: (row) => `<span class="${Number(row.annualMissingApplication || 0) > 0 ? "warning-pill" : "neutral-pill"}">${formatDays(row.annualMissingApplication || 0)}</span>` },
        { label: "월초 잔여", render: (row) => row.annualOpeningRemaining === "" ? "-" : formatDays(row.annualOpeningRemaining) },
        { label: "현재 잔여", render: (row) => row.annualRemaining === "" ? "-" : `<strong>${formatDays(row.annualRemaining)}</strong>` },
        { label: "당월 계획 내역", className: "message-cell", render: (row) => escapeHtml(formatAnnualEvents(row.annualLeaveEvents)) },
      ],
    },
    annualCompare: {
      rows: state.result.annualComparison?.rows || [],
      columns: [
        { label: "일자", key: "date" },
        { label: "매니저", key: "manager" },
        { label: "매장", key: "store" },
        { label: "사번", key: "employeeId" },
        { label: "이름", render: (row) => `<strong>${escapeHtml(row.name)}</strong>` },
        { label: "신청 구분", key: "requestedKind" },
        { label: "신청 일수", render: (row) => row.requestedDays ? formatDays(row.requestedDays) : "-" },
        { label: "근무계획", render: (row) => `<span class="plan-pill">${escapeHtml(row.planStatus || "공백")}</span>` },
        { label: "실제근태", render: (row) => escapeHtml(row.actualStatus || "미출근") },
        { label: "신청 상태", key: "applicationStatus" },
        { label: "대조 구분", className: "message-cell", render: (row) => `<span class="${row.needsReview ? "warning-pill" : "success-pill"}">${escapeHtml(row.category || row.result)}</span>` },
        { label: "확인상태", render: (row) => `<span class="${row.needsReview ? "status-pill" : "success-pill"}">${row.needsReview ? "확인 요청" : "동일"}</span>` },
        { label: "비고", className: "message-cell", key: "note" },
      ],
    },
    managerRequests: {
      rows: state.result.managerRequests || [],
      columns: [
        { label: "지역장", key: "regionalManager" },
        { label: "매니저", key: "manager" },
        { label: "매장", key: "store" },
        { label: "사번", key: "employeeId" },
        { label: "이름", render: (row) => `<strong>${escapeHtml(row.name)}</strong>` },
        { label: "문제 건수", key: "issueCount" },
        { label: "문제 요약", className: "message-cell", key: "issueText" },
        { label: "복사용 수정요청 멘트", className: "message-cell", render: (row) => `<span class="copy-message">${escapeHtml(row.message)}</span>` },
        { label: "전달상태", render: (row) => `<span class="${row.delivered ? "success-pill" : "status-pill"}">${escapeHtml(row.status || "미전달")}</span>` },
      ],
    },
    referenceCompare: {
      rows: state.result.referenceComparison?.rows || [],
      columns: [
        { label: "비교 구분", key: "comparisonType" },
        { label: "사번", key: "employeeId" },
        { label: "이름", key: "name" },
        { label: "매장", key: "store" },
        { label: "일자/항목", key: "date" },
        { label: "자동 생성값", key: "generatedValue" },
        { label: "비교 파일값", key: "referenceValue" },
        { label: "판정", render: (row) => `<span class="comparison-badge ${row.match ? "match" : row.informational ? "info" : "diff"}">${escapeHtml(row.result)}</span>` },
        { label: "차이 사유", className: "message-cell", key: "reason" },
      ],
    },
    personnelChecks: {
      rows: state.result.personnelChecks || [],
      columns: [
        { label: "경로", render: (row) => escapeHtml(ROUTE_LABELS[row.route] || row.route) },
        { label: "확인 유형", render: (row) => `<span class="${row.resolved ? "success-pill" : "warning-pill"}">${escapeHtml(row.issueType || "확인 요청")}</span>` },
        { label: "지역장", key: "regionalManager" },
        { label: "매니저", key: "manager" },
        { label: "매장", render: (row) => escapeHtml(row.storeName || row.store || "-") },
        { label: "사번", key: "employeeId" },
        { label: "이름", render: (row) => `<strong>${escapeHtml(row.employeeName || "-")}</strong>` },
        { label: "인력현황", render: (row) => escapeHtml(ROUTE_LABELS[row.workforceRoute] || row.workforceRoute || "없음") },
        { label: "연차대장", render: (row) => escapeHtml(ROUTE_LABELS[row.annualRoute] || row.annualRoute || "없음") },
        { label: "처리구분", render: (row) => `<span class="${row.resolved ? "success-pill" : "warning-pill"}">${escapeHtml(row.personnelStatus || "확인 요청")}</span>` },
        { label: "적용기간", render: (row) => escapeHtml([row.effectiveFrom, row.effectiveTo].filter(Boolean).join(" ~ ") || "-") },
        { label: "이동경로", render: (row) => escapeHtml(ROUTE_LABELS[row.destinationRoute] || row.destinationRoute || "-") },
        { label: "비고", className: "message-cell", key: "note" },
      ],
    },
  };
  return configs[tab] || configs.missing;
}

function renderJudgment(row) {
  const statusClass = row.shortage > 0 ? "status-pill" : row.baseExcess > 0 || row.substituteNeeded > 0 ? "warning-pill" : "success-pill";
  return `<span class="${statusClass}">${escapeHtml(row.judgment)}</span>`;
}

function renderCombinedJudgment(row) {
  const hasShortage = row.shortage > 0 || row.compensationShortage > 0;
  const hasUsage = row.substituteNeeded > 0 || row.compensationNeeded > 0 || row.baseExcess > 0;
  const statusClass = hasShortage ? "status-pill" : hasUsage ? "warning-pill" : "success-pill";
  return `<span class="${statusClass}">${escapeHtml(`${row.judgment} / ${row.compensationJudgment}`)}</span>`;
}

function filterRows(rows) {
  const query = $("#searchInput").value.trim().toLowerCase();
  const store = $("#storeFilter").value;
  const manager = $("#managerFilter")?.value || "";
  return rows.filter((row) => {
    const haystack = Object.values(row).map((value) => typeof value === "object" ? JSON.stringify(value) : String(value ?? "")).join(" ").toLowerCase();
    const managerMatch = !manager || state.activeResultTab !== "managerRequests" || row.manager === manager;
    return (!query || haystack.includes(query)) && (!store || row.store === store) && managerMatch;
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

async function exportResults() {
  if (!state.result) return;
  const button = $("#exportButton");
  try {
    button.disabled = true;
    button.textContent = "최종본 양식 생성 중...";
    const file = await buildFinalTemplateFile(state.result);
    downloadGeneratedFile(file);
    showToast("기존 최종본 양식으로 엑셀을 생성했습니다.");
  } catch (error) {
    console.error(error);
    showToast(error.message || "최종본 엑셀 생성 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "최종본 양식 엑셀 저장";
  }
}

function downloadGeneratedFile(file) {
  const url = URL.createObjectURL(file);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = file.name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function buildResultsWorkbook(result) {
  const workbook = XLSX.utils.book_new();
  appendCorrectedPlanSheet(workbook, result);
  appendAttendanceSheet(workbook, result);
  appendSheet(workbook, "3_근무인데 출근없음", result.missingRows, issueExportColumns());
  appendSheet(workbook, "4_휴무휴가인데 출근", result.unexpectedRows, issueExportColumns());
  appendSheet(workbook, "5_계획실제 불일치", result.mismatchRows, [
    ["경로", (row) => row.routeLabel], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
    ["일자", (row) => row.date], ["요일", (row) => row.weekday], ["계획 근태", (row) => row.planStatus], ["실제 근태", (row) => row.actualStatus],
    ["실제출근시간", (row) => row.actualIn], ["변경출근시간", (row) => row.changedIn], ["판정", (row) => row.result], ["사유", (row) => row.reason],
  ]);
  appendSheet(workbook, "6_기본휴무 초과", result.dayoffExcessRows, [
    ...summaryIdentityExportColumns(), ["기본 휴무 기준", (row) => row.baseAllowance], ["휴무 사용", (row) => row.basicDayoffUsed],
    ["초과 휴무", (row) => row.baseExcess], ["표기 대체휴무", (row) => row.explicitSubDayoffUsed],
    ["총 대체휴무 필요", (row) => row.substituteNeeded], ["판정", (row) => row.judgment],
  ]);
  appendSheet(workbook, "7_대체보상 잔여", result.balanceRows, [
    ...summaryIdentityExportColumns(),
    ["대체휴무 가용", (row) => row.availableSubstitute], ["대체휴무 필요", (row) => row.substituteNeeded],
    ["대체휴무 적용", (row) => row.substituteApplied], ["대체휴무 잔여", (row) => row.remainingSubstitute],
    ["대체휴무 만료", (row) => row.expiredSubstitute], ["대체휴무 초과", (row) => row.shortage],
    ["보상휴가 가용", (row) => row.availableCompensation], ["보상휴가 필요", (row) => row.compensationNeeded],
    ["보상휴가 적용", (row) => row.compensationApplied], ["보상휴가 잔여", (row) => row.remainingCompensation],
    ["보상휴가 만료", (row) => row.expiredCompensation], ["보상휴가 초과", (row) => row.compensationShortage],
    ["대체휴무 판정", (row) => row.judgment], ["보상휴가 판정", (row) => row.compensationJudgment],
  ]);
  appendSheet(workbook, "8_대체보상 초과", result.shortageRows, [
    ...summaryIdentityExportColumns(), ["대체휴무 필요", (row) => row.substituteNeeded], ["대체휴무 적용", (row) => row.substituteApplied],
    ["대체휴무 초과", (row) => row.shortage], ["보상휴가 필요", (row) => row.compensationNeeded],
    ["보상휴가 적용", (row) => row.compensationApplied], ["보상휴가 초과", (row) => row.compensationShortage],
    ["대체휴무 판정", (row) => row.judgment], ["보상휴가 판정", (row) => row.compensationJudgment],
  ]);
  appendSheet(workbook, "9_연차 등록누적", result.annualRows, [
    ...summaryIdentityExportColumns(), ["당월 연차", (row) => row.currentAnnualLeave], ["누적 연차", (row) => row.cumulativeAnnualLeave],
    ["당월 등록 내역", (row) => formatAnnualEvents(row.annualLeaveEvents)],
  ]);
  return workbook;
}

function appendCorrectedPlanSheet(workbook, result) {
  const plan = result.plan;
  const attendanceById = groupBy(result.attendance.rows, (row) => row.employeeId);
  const people = choosePlanRows(plan.rows, attendanceById);
  const maxDayColumn = Math.max(...plan.dayColumns.values());
  const maxCol = Math.max(plan.rawHeaders.length, ...people.map((person) => person.rawRow.length), maxDayColumn + 1);
  const header = [...plan.rawHeaders];
  while (header.length < maxCol) header.push("");
  const values = [header];
  const issueMap = new Map();
  for (const issue of result.mismatchRows) {
    addPlanExportIssue(issueMap, issue.employeeId, issue.date, issue.reason || issue.result);
  }
  for (const summary of result.employeeSummaries) {
    for (const event of summary.baseExcessEvents || []) {
      addPlanExportIssue(issueMap, summary.employeeId, event.date, `기본 휴무 기준 ${formatDays(summary.baseAllowance)} 초과분`);
    }
    if (Number(summary.shortage || 0) > 0) {
      for (const event of summary.substituteEvents || []) {
        addPlanExportIssue(issueMap, summary.employeeId, event.date, `대체휴무 잔여 부족 · 총 ${formatDays(summary.shortage)} 초과 사용`);
      }
    }
    if (Number(summary.compensationShortage || 0) > 0) {
      for (const event of summary.compensationEvents || []) {
        addPlanExportIssue(issueMap, summary.employeeId, event.date, `보상휴가 잔여 부족 · 총 ${formatDays(summary.compensationShortage)} 초과 사용`);
      }
    }
  }

  const redCells = [];
  const reasonRows = [];
  const firstDayCol = Math.min(...plan.dayColumns.values());
  for (const person of people) {
    const row = [...person.rawRow];
    while (row.length < maxCol) row.push("");
    values.push(row);
    const personRowIndex = values.length - 1;
    const reasons = new Array(maxCol).fill("");
    let hasIssue = false;
    for (const [day, col] of plan.dayColumns.entries()) {
      const messages = issueMap.get(`${person.employeeId}|${day}`) || [];
      if (!messages.length) continue;
      hasIssue = true;
      reasons[col] = [...new Set(messages)].join(" / ");
      redCells.push({ row: personRowIndex, col });
    }
    if (hasIssue) {
      const reasonLabel = person.duplicatePlanNote
        ? `빨간색 체크 사유(확인 필요) · ${person.duplicatePlanNote}`
        : "빨간색 체크 사유(확인 필요)";
      reasons[0] = reasonLabel;
      values.push(reasons);
      reasonRows.push(values.length - 1);
    }
  }

  const sheet = XLSX.utils.aoa_to_sheet(values);
  sheet["!freeze"] = { xSplit: 0, ySplit: 1 };
  sheet["!cols"] = Array.from({ length: maxCol }, (_, index) => ({ wch: index < firstDayCol ? (index === 3 ? 14 : 11) : 13 }));
  sheet["!rows"] = values.map((_, index) => ({ hpt: reasonRows.includes(index) ? 34 : 22 }));
  sheet["!merges"] = sheet["!merges"] || [];
  if (firstDayCol > 1) {
    for (const row of reasonRows) sheet["!merges"].push({ s: { r: row, c: 0 }, e: { r: row, c: firstDayCol - 1 } });
  }

  for (let col = 0; col < maxCol; col += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    if (cell) cell.s = excelHeaderStyle();
  }
  for (const { row, col } of redCells) {
    const ref = XLSX.utils.encode_cell({ r: row, c: col });
    if (!sheet[ref]) sheet[ref] = { t: "s", v: "" };
    sheet[ref].s = {
      fill: { patternType: "solid", fgColor: { rgb: "FFFF0000" } },
      font: { bold: true, color: { rgb: "FF000000" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: excelThinBorder(),
    };
  }
  for (const row of reasonRows) {
    for (let col = 0; col < maxCol; col += 1) {
      const ref = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[ref]) sheet[ref] = { t: "s", v: "" };
      sheet[ref].s = {
        fill: { patternType: "solid", fgColor: { rgb: col < firstDayCol ? "FFF4F4F4" : "FFFFFFFF" } },
        font: { bold: col === 0, color: { rgb: col === 0 ? "FFCC0000" : "FF333333" }, sz: 10 },
        alignment: { horizontal: col < firstDayCol ? "center" : "left", vertical: "center", wrapText: true },
        border: excelThinBorder(),
      };
    }
  }
  applyBodyBorders(sheet, values.length, maxCol);
  XLSX.utils.book_append_sheet(workbook, sheet, "1_근무 계획표");
}

function addPlanExportIssue(issueMap, employeeId, date, message) {
  const day = Number(String(date || "").slice(-2));
  if (!employeeId || !day || !message) return;
  const key = `${employeeId}|${day}`;
  if (!issueMap.has(key)) issueMap.set(key, []);
  issueMap.get(key).push(message);
}

function appendAttendanceSheet(workbook, result) {
  const rows = [...result.attendance.rows].sort((a, b) => a.date.localeCompare(b.date) || a.employeeId.localeCompare(b.employeeId));
  appendSheet(workbook, "2_근태 관리", rows, [
    ["이름", (row) => row.name], ["근무일자", (row) => row.date], ["실제출근시간", (row) => row.actualIn],
    ["변경출근시간", (row) => row.changedIn], ["사번", (row) => row.employeeId], ["출근지점", (row) => row.location],
    ["실제 근태", (row) => row.actualStatus], ["출근 판정", (row) => (row.actualIn || row.changedIn ? "출근 기록 있음" : "출근 기록 없음")],
  ]);
}

function excelHeaderStyle() {
  return {
    fill: { patternType: "solid", fgColor: { rgb: "FFD9E1F2" } },
    font: { bold: true, color: { rgb: "FF1F1F1F" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: excelThinBorder(),
  };
}

function excelThinBorder() {
  const side = { style: "thin", color: { rgb: "FFB7B7B7" } };
  return { top: side, bottom: side, left: side, right: side };
}

function applyBodyBorders(sheet, rowCount, colCount) {
  for (let row = 1; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      const ref = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[ref]) continue;
      sheet[ref].s = sheet[ref].s || {};
      sheet[ref].s.border = sheet[ref].s.border || excelThinBorder();
      sheet[ref].s.alignment = sheet[ref].s.alignment || { horizontal: "center", vertical: "center", wrapText: true };
    }
  }
}

async function buildResultFile(result) {
  return buildFinalTemplateFile(result);
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
  sheet["!cols"] = columns.map(([header]) => ({ wch: Math.max(12, Math.min(34, String(header).length * 2 + 4)) }));
  sheet["!rows"] = values.map((_, index) => ({ hpt: index === 0 ? 24 : 21 }));
  for (let col = 0; col < columns.length; col += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    if (cell) cell.s = excelHeaderStyle();
  }
  applyBodyBorders(sheet, values.length, columns.length);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
}

async function persistAnnualMonthlyFromAnalysis(result) {
  const applications = (result.annualApplications || []).filter((row) => ["연차", "반차"].includes(row.requestedKind)).map((row) => ({
    employeeId: row.employeeId, employeeName: row.name, leaveDate: row.date, days: row.requestedDays,
    status: row.applicationStatus, leaveType: row.leaveType || row.requestedKind,
    applicationDate: row.applicationDate, note: row.note, sourceIndex: row.sourceIndex,
  }));
  if (!applications.length) return null;
  const response = await fetch("/api/annual-leave", {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ action: "monthly", route: result.route, month: result.targetMonth, fileName: result.annualFileName || "연차신청현황.xlsx", applications }),
  });
  const data = await readJsonResponse(response);
  if (!response.ok) throw new Error(data.error || "월별 연차 승인·반려 저장 실패");
  return data;
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
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "월 마감 저장 실패");

    applyServerSummaries(data.summaries || []);
    mergeAnnualLeaveLedger(state.result, state.result.annualLedger, state.result.annualComparison);
    if (state.annualFile && (state.result.annualApplications || []).length) {
      await persistAnnualMonthlyFromAnalysis(state.result);
    }
    if ((state.result.personnelOverrides || []).length) {
      await persistPersonnelOverrides(state.result.personnelOverrides, state.result.targetMonth);
    }

    const correctedResultFile = await buildResultFile(state.result);
    const fileResults = await Promise.allSettled([
      uploadArchiveFile({
        file: state.planFile,
        route: result.route,
        month: result.targetMonth,
        fileKind: "plan",
        note: "월 마감 저장 시 자동 보관",
        sourceType: "closure",
        closureId: data.id,
        replace: true,
      }),
      uploadArchiveFile({
        file: state.attendanceFile,
        route: result.route,
        month: result.targetMonth,
        fileKind: "attendance",
        note: "월 마감 저장 시 자동 보관",
        sourceType: "closure",
        closureId: data.id,
        replace: true,
      }),
      uploadArchiveFile({
        file: correctedResultFile,
        route: result.route,
        month: result.targetMonth,
        fileKind: "result",
        note: "최종본 양식 자동생성 · 증빙 O 자동반영 · 매니저 공유용 · 연차/최종본 비교 포함",
        sourceType: "closure",
        closureId: data.id,
        replace: true,
      }),
      ...(state.annualFile ? [uploadArchiveFile({
        file: state.annualFile, route: result.route, month: result.targetMonth, fileKind: "other",
        note: "월 마감 연차신청현황 대조 원본", sourceType: "closure", closureId: data.id, replace: false,
      })] : []),
      ...(state.referenceFile ? [uploadArchiveFile({
        file: state.referenceFile, route: result.route, month: result.targetMonth, fileKind: "other",
        note: "증빙 O 반영·최종본 비교 원본", sourceType: "closure", closureId: data.id, replace: false,
      })] : []),
    ]);
    await Promise.all([loadHistory(), loadGrants(), loadArchiveFiles(), loadWorkforceUploads(), loadAnnualLeaveDashboard(), loadPersonnelChecks()]);
    const action = data.replaced ? "기존 월 마감을 완전히 교체했습니다." : "새 월 마감을 저장했습니다.";
    const recalculateMessage = data.recalculateWarning
      ? ` 월 마감 원본은 저장됐지만 누적 재계산 경고가 있습니다: ${data.recalculateWarning}`
      : ` ${data.affectedMonths || 0}개 월의 연차·대체휴무·보상휴가 누적을 다시 계산했습니다.`;
    const failedFiles = fileResults.filter((item) => item.status === "rejected");
    const fileMessage = failedFiles.length
      ? ` 월 마감은 정상 저장됐지만 원본 파일 ${failedFiles.length}개는 보관하지 못했습니다: ${failedFiles.map((item) => item.reason?.message || "파일 보관 오류").join(" / ")}`
      : " 계획표·근태표·보고용 결과본과 선택한 비교 원본을 파일 보관함에 저장했습니다.";
    showToast(`${action}${recalculateMessage}${fileMessage}`);
  } catch (error) {
    showToast(error.message || "월 마감 저장 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "월 마감 교체 저장";
  }
}


function resetClosureComparisonOutput() {
  state.closureComparison = null;
  const summary = $("#closureCompareSummary");
  const wrap = $("#closureCompareTableWrap");
  const body = $("#closureCompareTableBody");
  if (summary) { summary.innerHTML = ""; summary.classList.add("hidden"); }
  if (wrap) wrap.classList.add("hidden");
  if (body) body.innerHTML = "";
  if ($("#exportClosureCompare")) $("#exportClosureCompare").disabled = true;
}

function resetClosureComparison() {
  state.closureBaseFile = null;
  state.closureTargetFile = null;
  $("#closureBaseFile").value = "";
  $("#closureTargetFile").value = "";
  $("#closureBaseFileName").textContent = "기준 최종본을 선택하거나 끌어놓기";
  $("#closureTargetFileName").textContent = "비교 최종본을 선택하거나 끌어놓기";
  resetClosureComparisonOutput();
  $("#closureCompareNotice").textContent = "보관함의 결과·수정본에서 ‘기준본 A’ 또는 ‘비교본 B’를 눌러도 바로 선택할 수 있습니다.";
}

async function analyzeClosureComparison() {
  const button = $("#compareClosuresButton");
  try {
    if (!state.closureBaseFile || !state.closureTargetFile) throw new Error("기준 마감본 A와 비교 마감본 B를 모두 선택해 주세요.");
    button.disabled = true;
    button.textContent = "두 파일 비교 중...";
    const [baseSheets, targetSheets] = await Promise.all([
      fileToWorkbookSheets(state.closureBaseFile),
      fileToWorkbookSheets(state.closureTargetFile),
    ]);
    const base = parseReferenceFinalWorkbook(baseSheets);
    const target = parseReferenceFinalWorkbook(targetSheets);
    const comparison = compareTwoFinalWorkbooks(base, target);
    state.closureComparison = {
      ...comparison,
      baseFileName: state.closureBaseFile.name,
      targetFileName: state.closureTargetFile.name,
    };
    renderClosureComparison(state.closureComparison);
    showToast(comparison.rows.length ? `월 마감 최종본 차이 ${comparison.rows.length}건을 찾았습니다.` : "두 월 마감 최종본이 모두 일치합니다.");
  } catch (error) {
    console.error(error);
    showToast(error.message || "월 마감 최종본 비교 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "두 마감본 차이 찾기";
  }
}

function compareTwoFinalWorkbooks(base, target) {
  const rows = [];
  const sameMonth = Boolean(base.month && target.month && base.month === target.month);
  const keys = new Set([...base.values.keys(), ...target.values.keys()]);
  for (const key of keys) {
    const left = base.values.get(key);
    const right = target.values.get(key);
    const baseValue = normalizeFinalCompareValue(left?.value || "직원/일자 없음");
    const targetValue = normalizeFinalCompareValue(right?.value || "직원/일자 없음");
    if (baseValue === targetValue) continue;
    const [employeeId, date] = key.split("|");
    rows.push({
      comparisonType: !left || !right ? "직원·일자 누락" : "근태 값 차이",
      store: left?.store || right?.store || "",
      employeeId,
      name: left?.name || right?.name || "",
      date: left?.date || right?.date || date || "",
      baseValue,
      targetValue,
      result: !left ? "기준본 없음" : !right ? "비교본 없음" : "불일치",
    });
  }
  rows.sort((a, b) => a.date.localeCompare(b.date) || a.store.localeCompare(b.store, "ko") || a.name.localeCompare(b.name, "ko"));
  return {
    sameMonth,
    baseMonth: base.month || "확인 불가",
    targetMonth: target.month || "확인 불가",
    baseEmployeeCount: base.employeeCount || 0,
    targetEmployeeCount: target.employeeCount || 0,
    comparedCells: keys.size,
    rows,
  };
}

function renderClosureComparison(comparison) {
  const summary = $("#closureCompareSummary");
  const wrap = $("#closureCompareTableWrap");
  const body = $("#closureCompareTableBody");
  summary.classList.remove("hidden");
  wrap.classList.remove("hidden");
  summary.innerHTML = [
    ["기준 월", comparison.baseMonth],
    ["비교 월", comparison.targetMonth],
    ["기준 인원", `${number(comparison.baseEmployeeCount)}명`],
    ["비교 인원", `${number(comparison.targetEmployeeCount)}명`],
    ["차이", `${number(comparison.rows.length)}건`],
  ].map(([label, value]) => `<div class="summary-chip"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
  body.innerHTML = comparison.rows.length ? comparison.rows.map((row, index) => `<tr>
    <td>${index + 1}</td>
    <td>${escapeHtml(row.comparisonType)}</td>
    <td>${escapeHtml(row.store)}</td>
    <td>${escapeHtml(row.employeeId)}</td>
    <td>${escapeHtml(row.name)}</td>
    <td>${escapeHtml(row.date)}</td>
    <td><span class="comparison-badge diff">${escapeHtml(row.baseValue)}</span></td>
    <td><span class="comparison-badge diff">${escapeHtml(row.targetValue)}</span></td>
    <td>${escapeHtml(row.result)}</td>
  </tr>`).join("") : `<tr><td colspan="9" class="empty-cell">두 최종본의 사번·날짜별 근태 값이 모두 일치합니다.</td></tr>`;
  $("#closureCompareNotice").innerHTML = comparison.sameMonth
    ? `같은 월(${escapeHtml(comparison.baseMonth)}) 기준으로 ${number(comparison.comparedCells)}개 사번·날짜 셀을 비교했습니다.`
    : `<strong>대상 월이 다릅니다.</strong> 기준본 ${escapeHtml(comparison.baseMonth)} / 비교본 ${escapeHtml(comparison.targetMonth)} · 월이 다른 파일은 날짜별 차이가 많이 발생할 수 있습니다.`;
  $("#exportClosureCompare").disabled = false;
}

function exportClosureComparison() {
  const comparison = state.closureComparison;
  if (!comparison) return;
  const rows = [
    ["No", "구분", "점포", "사번", "이름", "날짜", "기준본 A", "비교본 B", "판정"],
    ...comparison.rows.map((row, index) => [index + 1, row.comparisonType, row.store, row.employeeId, row.name, row.date, row.baseValue, row.targetValue, row.result]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!views"] = [{ showGridLines: false }];
  ws["!cols"] = [{ wch: 7 }, { wch: 15 }, { wch: 18 }, { wch: 13 }, { wch: 12 }, { wch: 13 }, { wch: 16 }, { wch: 16 }, { wch: 12 }];
  for (let col = 0; col < rows[0].length; col += 1) {
    const address = XLSX.utils.encode_cell({ r: 0, c: col });
    if (ws[address]) ws[address].s = { fill: { patternType: "solid", fgColor: { rgb: "FF203764" } }, font: { name: "맑은 고딕", bold: true, color: { rgb: "FFFFFFFF" } }, alignment: { horizontal: "center", vertical: "center" } };
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "월 마감 차이");
  const month = comparison.baseMonth !== "확인 불가" ? comparison.baseMonth : "월미확인";
  XLSX.writeFile(wb, `${month}_월마감_최종본_차이.xlsx`, { bookType: "xlsx", cellStyles: true });
}

function archiveKindFromName(fileName) {
  const name = String(fileName || "").replace(/\s+/g, "");
  if (/최종본|출퇴근현황|결과|공유용|수정본/.test(name)) return "result";
  if (/계획|스케줄/.test(name)) return "plan";
  if (/근태|출근RAW|근태RAW|출퇴근RAW/.test(name)) return "attendance";
  return "other";
}

function archiveKindFromSheets(sheets, fallback = "other") {
  const names = (sheets || []).map((sheet) => String(sheet.sheetName || "").replace(/\s+/g, ""));
  if (names.some((name) => /상담사.*근태/.test(name))) return "result";
  if (names.some((name) => /근무계획|매장근무계획|스케줄/.test(name))) return "plan";
  if (names.some((name) => /근태관리|근태RAW|출퇴근RAW|출퇴근기록/.test(name))) return "attendance";
  return fallback;
}

async function inspectArchiveFile(file, route, month) {
  const fallbackKind = archiveKindFromName(file.name);
  let sheets = null;
  let fileKind = fallbackKind;
  let parsedFinal = null;
  let warning = "";
  try {
    sheets = await fileToWorkbookSheets(file);
    fileKind = archiveKindFromSheets(sheets, fallbackKind);
    if (fileKind === "result") {
      const hasFinalSheet = sheets.some((sheet) => /상담사\s*근태|상담사.*근태/.test(String(sheet.sheetName || "")));
      if (hasFinalSheet) parsedFinal = parseFinalLeaveImportWorkbook(sheets, month, route);
      else warning = "상담사근태 시트가 없어 휴가 사용내역은 자동 반영되지 않습니다.";
    }
  } catch (error) {
    warning = `파일 구조 확인 오류: ${error.message || "파일을 읽지 못했습니다."}`;
  }
  return { file, route, month, fileKind, parsedFinal, warning };
}

async function previewArchiveUploads(route) {
  const month = $("#archiveMonth")?.value || "";
  const input = route === "homeplus" ? $("#archiveHomeplusFiles") : $("#archiveElectrolandFiles");
  const target = route === "homeplus" ? $("#archiveHomeplusPreview") : $("#archiveElectrolandPreview");
  const files = [...(input?.files || [])];
  state.archiveUploadPreviews[route] = [];
  if (!target) return;
  if (!month) {
    target.textContent = "먼저 대상 월을 선택해 주세요.";
    return;
  }
  if (!files.length) {
    target.textContent = `선택된 ${ROUTE_LABELS[route]} 파일이 없습니다.`;
    return;
  }
  target.innerHTML = `<div class="preview-loading">${number(files.length)}개 파일을 자동 분류하는 중입니다...</div>`;
  const previews = await Promise.all(files.map((file) => inspectArchiveFile(file, route, month)));
  state.archiveUploadPreviews[route] = previews;
  renderArchiveUploadPreview(route, previews);
}

function previewAllArchiveUploads() {
  previewArchiveUploads("homeplus");
  previewArchiveUploads("electroland");
}

function renderArchiveUploadPreview(route, previews) {
  const target = route === "homeplus" ? $("#archiveHomeplusPreview") : $("#archiveElectrolandPreview");
  if (!target) return;
  if (!previews.length) {
    target.textContent = `선택된 ${ROUTE_LABELS[route]} 파일이 없습니다.`;
    return;
  }
  target.innerHTML = previews.map((item) => {
    const finalInfo = item.parsedFinal
      ? `<span class="preview-detail">직원 ${number(item.parsedFinal.employeeCount)}명 · 사용 대체 ${formatDays(item.parsedFinal.substituteDays)} · 사용 보상 ${formatDays(item.parsedFinal.compensationDays)} 자동 반영 · 잔여는 부여 설정 기준 재계산</span>`
      : item.warning
        ? `<span class="preview-warning">${escapeHtml(item.warning)}</span>`
        : `<span class="preview-detail">파일 저장 대상</span>`;
    return `<div class="file-preview-item">
      <div class="file-preview-main"><strong>${escapeHtml(item.file.name)}</strong><span class="classification-pill kind-${escapeHtml(item.fileKind)}">${escapeHtml(archiveKindLabel(item.fileKind))}</span></div>
      ${finalInfo}
    </div>`;
  }).join("");
}

function selectedArchiveInputs() {
  return [
    { route: "homeplus", files: [...($("#archiveHomeplusFiles")?.files || [])] },
    { route: "electroland", files: [...($("#archiveElectrolandFiles")?.files || [])] },
  ];
}

async function saveArchiveFiles(event) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    openLogin();
    return;
  }
  const month = $("#archiveMonth")?.value || "";
  if (!month) {
    showToast("대상 월을 선택해 주세요.");
    return;
  }
  const inputs = selectedArchiveInputs();
  const selectedCount = inputs.reduce((sum, group) => sum + group.files.length, 0);
  if (!selectedCount) {
    showToast("홈플러스 또는 전자랜드 파일을 하나 이상 선택해 주세요.");
    return;
  }

  const button = $("#archiveSubmitButton");
  const commonNote = text($("#archiveNote")?.value);
  const saved = [];
  const imported = [];
  const failed = [];
  const importFailed = [];
  try {
    button.disabled = true;
    for (const group of inputs) {
      let previews = state.archiveUploadPreviews[group.route] || [];
      const previewFiles = new Set(previews.map((item) => item.file));
      if (previews.length !== group.files.length || group.files.some((file) => !previewFiles.has(file)) || previews.some((item) => item.month !== month)) {
        button.textContent = `${ROUTE_LABELS[group.route]} 파일 확인 중...`;
        previews = await Promise.all(group.files.map((file) => inspectArchiveFile(file, group.route, month)));
        state.archiveUploadPreviews[group.route] = previews;
        renderArchiveUploadPreview(group.route, previews);
      }
      for (let index = 0; index < previews.length; index += 1) {
        const item = previews[index];
        button.textContent = `파일 저장 중 ${saved.length + failed.length + 1}/${selectedCount}`;
        try {
          await uploadArchiveFile({
            file: item.file,
            route: item.route,
            month,
            fileKind: item.fileKind,
            note: commonNote,
            sourceType: "manual",
            replace: true,
          });
          saved.push(item);
          if (item.fileKind === "result" && item.parsedFinal) {
            button.textContent = `${ROUTE_LABELS[item.route]} 최종본 휴가내역 반영 중...`;
            try {
              const result = await importFinalLeaveData(item.file, item.route, month, item.parsedFinal);
              imported.push({ item, result });
            } catch (error) {
              importFailed.push({ item, error });
            }
          }
        } catch (error) {
          failed.push({ item, error });
        }
      }
    }
    $("#archiveHomeplusFiles").value = "";
    $("#archiveElectrolandFiles").value = "";
    $("#archiveNote").value = "";
    state.archiveUploadPreviews = { homeplus: [], electroland: [] };
    renderArchiveUploadPreview("homeplus", []);
    renderArchiveUploadPreview("electroland", []);
    await Promise.all([loadArchiveFiles(), loadHistory(), loadGrants()]);
    const summary = [`파일 ${saved.length}개 저장`];
    if (imported.length) summary.push(`최종본 휴가내역 ${imported.length}건 반영`);
    if (importFailed.length) summary.push(`휴가 반영 대기 ${importFailed.length}건`);
    if (failed.length) summary.push(`파일 저장 실패 ${failed.length}건`);
    showToast(summary.join(" · "));
  } catch (error) {
    showToast(error.message || "월별 파일 저장 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "선택한 파일 저장·교체";
  }
}

async function importFinalLeaveData(file, route, month, parsed = null) {
  let finalParsed = parsed;
  if (!finalParsed) {
    const sheets = await fileToWorkbookSheets(file);
    finalParsed = parseFinalLeaveImportWorkbook(sheets, month, route);
  }
  const response = await fetch("/api/final-leave-import", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ route, month, fileName: file.name, employeeFacts: finalParsed.employeeFacts }),
  });
  if (response.status === 401) {
    await checkBackend();
    throw new Error("관리자 로그인이 만료되었습니다. 다시 로그인해 주세요.");
  }
  const data = await readJsonResponse(response);
  if (!response.ok) throw new Error(data.error || "최종본 휴가 사용내역 반영 실패");
  return data;
}

async function uploadArchiveFile({ file, route, month, fileKind, note = "", sourceType = "manual", closureId = "", replace = false }) {
  if (!(file instanceof File)) throw new Error("보관할 파일이 없습니다.");
  const form = new FormData();
  form.append("file", file, file.name);
  form.append("route", route);
  form.append("month", month);
  form.append("fileKind", fileKind);
  form.append("note", note);
  form.append("sourceType", sourceType);
  form.append("closureId", closureId || "");
  form.append("replace", replace ? "true" : "false");

  const response = await fetch("/api/archive-files", { method: "POST", body: form });
  if (response.status === 401) {
    await checkBackend();
    throw new Error("관리자 로그인이 만료되었습니다. 다시 로그인해 주세요.");
  }
  const data = await readJsonResponse(response);
  if (!response.ok) throw new Error(data.error || "파일 보관 실패");
  return data;
}

async function loadArchiveFiles() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    state.archiveFiles = [];
    renderArchiveFiles([]);
    return;
  }
  try {
    const params = new URLSearchParams();
    const route = $("#archiveRouteFilter")?.value || "";
    const month = $("#archiveMonthFilter")?.value || "";
    if (route) params.set("route", route);
    if (month) params.set("month", month);
    const response = await fetch(`/api/archive-files${params.toString() ? `?${params}` : ""}`, { cache: "no-store" });
    if (response.status === 401) {
      await checkBackend();
      renderArchiveFiles([]);
      return;
    }
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "월 마감 파일 조회 실패");
    state.archiveFiles = data.items || [];
    state.backend.fileStorageConfigured = Boolean(data.r2Configured);
    state.backend.fileStorageMode = data.r2Configured ? "r2" : "d1";
    renderArchiveFiles(state.archiveFiles);
    updateStorageUI();
  } catch (error) {
    state.archiveFiles = [];
    renderArchiveFiles([]);
    showToast(error.message || "월 마감 파일 조회 실패");
  }
}

function renderArchiveFiles(items) {
  const body = $("#archiveTableBody");
  if (!body) return;
  renderArchiveSummary(items);
  body.innerHTML = items.length ? items.map((item) => {
    const canAnalyze = item.file_kind === "plan" || item.file_kind === "attendance";
    const looksFinal = item.file_kind === "result" || /최종본|출퇴근현황|결과|공유용|수정본/.test(String(item.file_name || ""));
    const leaveStatus = looksFinal
      ? Number(item.leave_imported || 0) > 0
        ? '<span class="success-pill">반영 완료</span>'
        : '<span class="warning-pill">미반영</span>'
      : '<span class="neutral-pill">해당 없음</span>';
    return `<tr>
      <td>${escapeHtml(ROUTE_LABELS[item.route] || item.route)}</td>
      <td><strong>${escapeHtml(item.month)}</strong></td>
      <td><span class="classification-pill kind-${escapeHtml(item.file_kind)}">${escapeHtml(archiveKindLabel(item.file_kind))}</span>${item.source_type === "closure" ? '<br><span class="neutral-pill">마감 자동보관</span>' : ""}</td>
      <td class="message-cell"><strong>${escapeHtml(item.file_name)}</strong></td>
      <td>${leaveStatus}</td>
      <td>${escapeHtml(formatFileSize(item.size_bytes))}</td>
      <td>${item.storage_type === "r2" ? "R2" : "D1"}</td>
      <td class="message-cell">${escapeHtml(item.note || "-")}</td>
      <td>${escapeHtml(formatStoredDate(item.created_at))}</td>
      <td class="action-cell archive-actions">
        <a class="btn secondary small" href="/api/archive-files/${encodeURIComponent(item.id)}">다운로드</a>
        ${canAnalyze && !looksFinal ? `<button class="btn secondary small archive-use" data-id="${escapeHtml(item.id)}" type="button">분석에 불러오기</button>` : ""}
        ${looksFinal ? `<button class="btn primary small archive-leave-import" data-id="${escapeHtml(item.id)}" type="button">${Number(item.leave_imported || 0) > 0 ? "휴가 재반영" : "휴가내역 반영"}</button><button class="btn secondary small archive-evidence" data-id="${escapeHtml(item.id)}" type="button">증빙 반영</button><button class="btn secondary small archive-compare" data-side="base" data-id="${escapeHtml(item.id)}" type="button">기준본 A</button><button class="btn secondary small archive-compare" data-side="target" data-id="${escapeHtml(item.id)}" type="button">비교본 B</button>` : ""}
        <button class="btn danger small archive-delete" data-id="${escapeHtml(item.id)}" type="button">삭제</button>
      </td>
    </tr>`;
  }).join("") : `<tr><td colspan="10" class="empty-cell">저장된 월별 파일이 없습니다.</td></tr>`;

  $$(".archive-use").forEach((button) => button.addEventListener("click", () => useArchiveFile(button.dataset.id, "analyze")));
  $$(".archive-leave-import").forEach((button) => button.addEventListener("click", () => importLeaveFromArchive(button.dataset.id, button)));
  $$(".archive-evidence").forEach((button) => button.addEventListener("click", () => useArchiveFile(button.dataset.id, "evidence")));
  $$(".archive-compare").forEach((button) => button.addEventListener("click", () => useArchiveFile(button.dataset.id, button.dataset.side)));
  $$(".archive-delete").forEach((button) => button.addEventListener("click", () => deleteArchiveFile(button.dataset.id)));
}

function renderArchiveSummary(items) {
  const target = $("#archiveSummary");
  if (!target) return;
  const totalBytes = items.reduce((sum, item) => sum + Number(item.size_bytes || 0), 0);
  const homeplusCount = items.filter((item) => item.route === "homeplus").length;
  const electrolandCount = items.filter((item) => item.route === "electroland").length;
  const finalFiles = items.filter((item) => item.file_kind === "result" || /최종본|출퇴근현황|결과|공유용|수정본/.test(String(item.file_name || "")));
  const importedCount = finalFiles.filter((item) => Number(item.leave_imported || 0) > 0).length;
  target.innerHTML = [
    ["보관 파일", `${items.length}개`],
    ["홈플러스", `${homeplusCount}개`],
    ["전자랜드", `${electrolandCount}개`],
    ["최종본 휴가 반영", `${importedCount}/${finalFiles.length}개`],
    ["전체 용량", formatFileSize(totalBytes)],
  ].map(([label, value]) => `<div class="summary-chip"><span>${label}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
}

async function fetchArchiveFile(item) {
  const response = await fetch(`/api/archive-files/${encodeURIComponent(item.id)}`, { cache: "no-store" });
  if (!response.ok) {
    const data = await readJsonResponse(response);
    throw new Error(data.error || "보관 파일 불러오기 실패");
  }
  const blob = await response.blob();
  return new File([blob], item.file_name, { type: item.content_type || blob.type || "application/octet-stream" });
}

async function useArchiveFile(id, mode = "analyze") {
  const item = state.archiveFiles.find((row) => row.id === id);
  if (!item) return;
  try {
    const file = await fetchArchiveFile(item);
    if (mode === "base" || mode === "target") {
      if (mode === "base") setClosureBaseFile(file);
      else setClosureTargetFile(file);
      switchView("compare");
      showToast(`${mode === "base" ? "기준 마감본 A" : "비교 마감본 B"}에 파일을 넣었습니다.`);
      return;
    }

    const sheets = await fileToWorkbookSheets(file);
    const isFinalWorkbook = sheets.some((sheet) => /상담사\s*근태/.test(sheet.sheetName))
      && sheets.some((sheet) => /출근\s*미등록|전체\s*요약본|매니저별\s*이상/.test(sheet.sheetName));
    if (mode === "evidence" || item.file_kind === "result" || isFinalWorkbook) {
      setReferenceFile(file);
      const routeInput = document.querySelector(`input[name="route"][value="${item.route}"]`);
      if (routeInput) routeInput.checked = true;
      $("#targetMonth").value = item.month;
      syncCutoffWithMonth();
      syncRouteRuleHelp();
      switchView("checker");
      showToast("결과·수정본을 증빙 반영 칸에 불러왔습니다. 계획표와 실제 근태표를 선택한 뒤 다시 분석해 주세요.");
      return;
    }

    const routeInput = document.querySelector(`input[name="route"][value="${item.route}"]`);
    if (routeInput) routeInput.checked = true;
    $("#targetMonth").value = item.month;
    syncCutoffWithMonth();
    syncRouteRuleHelp();
    if (item.file_kind === "plan") setPlanFile(file);
    else if (item.file_kind === "attendance") setAttendanceFile(file);
    else throw new Error("이 파일은 계획표나 실제 근태표로 불러올 수 없습니다.");
    switchView("checker");
    showToast(`${archiveKindLabel(item.file_kind)}를 분석 화면에 불러왔습니다.`);
  } catch (error) {
    showToast(error.message || "보관 파일 불러오기 실패");
  }
}

async function importLeaveFromArchive(id, button = null) {
  const item = state.archiveFiles.find((row) => row.id === id);
  if (!item) return;
  const originalText = button?.textContent || "휴가내역 반영";
  try {
    if (button) {
      button.disabled = true;
      button.textContent = "반영 중...";
    }
    const file = await fetchArchiveFile(item);
    const result = await importFinalLeaveData(file, item.route, item.month);
    await Promise.all([loadArchiveFiles(), loadHistory(), loadGrants()]);
    showToast(`${item.month} ${ROUTE_LABELS[item.route]} 최종본을 반영했습니다. 사용 대체 ${formatDays(result.substituteDays)} · 사용 보상 ${formatDays(result.compensationDays)}을 저장했고, 잔여는 부여 설정의 발생일·사용기간 기준으로 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message || "최종본 휴가내역 반영 실패");
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = originalText;
    }
  }
}

async function deleteArchiveFile(id) {
  const item = state.archiveFiles.find((row) => row.id === id);
  if (!item || !confirm(`보관된 파일 “${item.file_name}”을 삭제하시겠습니까?`)) return;
  try {
    const response = await fetch(`/api/archive-files/${encodeURIComponent(id)}`, { method: "DELETE" });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "파일 삭제 실패");
    await loadArchiveFiles();
    showToast("보관 파일을 삭제했습니다.");
  } catch (error) {
    showToast(error.message || "파일 삭제 실패");
  }
}

function archiveKindLabel(kind) {
  return ({ plan: "계획표", attendance: "실제 근태표", result: "최종본·수정본", other: "기타" })[kind] || kind;
}

function formatFileSize(value) {
  const bytes = Number(value || 0);
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(bytes >= 10240 ? 0 : 1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function formatStoredDate(value) {
  if (!value) return "-";
  const normalized = /Z$|[+-]\d{2}:?\d{2}$/.test(value) ? value : `${value}Z`;
  const parsed = new Date(normalized);
  return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleString("ko-KR");
}

async function readJsonResponse(response) {
  const textValue = await response.text();
  if (!textValue) return {};
  try { return JSON.parse(textValue); } catch { return { error: textValue }; }
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
      compensationNeeded: roundHalf(row.compensation_needed),
      availableCompensation: roundHalf(row.available_compensation),
      compensationApplied: roundHalf(row.compensation_applied),
      remainingCompensation: roundHalf(row.remaining_compensation),
      expiredCompensation: roundHalf(row.expired_compensation),
      compensationShortage: roundHalf(row.compensation_shortage),
      dayoffReplacementUsed: roundHalf(row.dayoff_replacement_used),
      dayoffReplacementShortage: roundHalf(row.dayoff_replacement_shortage),
      dayoffReplacementSubstituteUsed: roundHalf(row.dayoff_replacement_substitute_used),
      dayoffReplacementCompensationUsed: roundHalf(row.dayoff_replacement_compensation_used),
      previousMonth: state.priorLedger.previousMonth || "",
      priorDayoffExcess: roundHalf(state.priorLedger.previousMonthFacts?.[normalizeEmployeeId(row.employee_id)]?.baseExcess || 0),
      priorDayoffReplacementUsed: roundHalf(state.priorLedger.previousMonthFacts?.[normalizeEmployeeId(row.employee_id)]?.dayoffReplacementUsed || 0),
      priorDayoffReplacementShortage: roundHalf(state.priorLedger.previousMonthFacts?.[normalizeEmployeeId(row.employee_id)]?.dayoffReplacementShortage || 0),
      combinedAvailable: roundHalf(Number(row.available_substitute || 0) + Number(row.available_compensation || 0)),
      combinedLeaveUsed: roundHalf(Number(row.substitute_needed || 0) + Number(row.compensation_needed || 0) + Number(row.dayoff_replacement_used || 0)),
      combinedRemaining: roundHalf(Number(row.remaining_substitute || 0) + Number(row.remaining_compensation || 0)),
      combinedShortage: roundHalf(Number(row.shortage || 0) + Number(row.compensation_shortage || 0) + Number(row.dayoff_replacement_shortage || 0)),
      currentAnnualLeave: roundHalf(row.current_annual_leave),
      cumulativeAnnualLeave: roundHalf(row.cumulative_annual_leave),
      judgment: row.judgment || "",
      compensationJudgment: row.compensation_judgment || "",
      annualLeaveEvents: fact.annualLeaveEvents || [],
      substituteEvents: fact.substituteEvents || [],
      compensationEvents: fact.compensationEvents || [],
    };
  });
  sortSummaries(employeeSummaries);
  state.result = assembleResultCollections({
    ...state.result,
    employeeSummaries,
    employeeFacts: state.result.employeeFacts,
  });
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
    if (response.status === 401) {
      await checkBackend();
      renderHistory([]);
      return;
    }
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "월 마감 기록 조회 실패");
    renderHistory(data.items || []);
  } catch (error) {
    renderHistory([]);
    showToast(error.message || "월 마감 기록 조회 실패");
  }
}

function renderHistory(items) {
  $("#historyEmpty").classList.toggle("hidden", items.length > 0);
  $("#historyList").innerHTML = items.map((item) => {
    const routeLabel = ROUTE_LABELS[item.route] || item.route;
    return `<article class="history-card">
      <div><div class="month">${escapeHtml(item.month)}</div><div class="history-meta">${escapeHtml(routeLabel)} 경로 · 기준 ${escapeHtml(item.cutoff_date || "-")}</div></div>
      <div><strong>${escapeHtml(item.plan_file_name || "계획표")}</strong><div class="history-meta">${escapeHtml(item.attendance_file_name || "근태표")} · ${escapeHtml(formatStoredDate(item.created_at))}</div></div>
      <div class="history-kpis">
        <div><span>출근기록 없음</span><strong>${number(item.missing_count || 0)}</strong></div>
        <div><span>휴무·휴가 출근</span><strong>${number(item.unexpected_count || 0)}</strong></div>
        <div><span>불일치</span><strong>${number(item.mismatch_count || 0)}</strong></div>
        <div><span>대체 초과</span><strong>${number(item.substitute_shortage_people || 0)}</strong></div>
        <div><span>보상 초과</span><strong>${number(item.compensation_shortage_people || 0)}</strong></div>
        <div><span>연차 등록자</span><strong>${number(item.annual_leave_people || 0)}</strong></div>
        <div class="action-cell"><button class="btn secondary small history-files" data-route="${escapeHtml(item.route)}" data-month="${escapeHtml(item.month)}" type="button">파일 보기</button><button class="btn danger small history-delete" data-id="${escapeHtml(item.id)}" data-route-label="${escapeHtml(routeLabel)}" data-month="${escapeHtml(item.month)}" type="button">마감 삭제</button></div>
      </div>
    </article>`;
  }).join("");
  $$(".history-files").forEach((button) => button.addEventListener("click", () => openHistoryFiles(button.dataset.route, button.dataset.month)));
  $$(".history-delete").forEach((button) => button.addEventListener("click", () => deleteClosure(button.dataset.id, button.dataset.routeLabel, button.dataset.month)));
}

function openHistoryFiles(route, month) {
  $("#archiveRouteFilter").value = route || "";
  $("#archiveMonthFilter").value = month || "";
  switchView("files");
}

async function deleteClosure(id, routeLabel, month) {
  if (!confirm(`${routeLabel} 경로 ${month} 월 마감을 삭제하시겠습니까? 이후 월의 연차·대체휴무·보상휴가 누적도 다시 계산됩니다.`)) return;
  try {
    const response = await fetch(`/api/closures/${encodeURIComponent(id)}`, { method: "DELETE" });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "월 마감 삭제 실패");
    await Promise.all([loadHistory(), loadGrants()]);
    showToast(`월 마감을 삭제하고 ${data.affectedMonths || 0}개 이후 월을 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message || "월 마감 삭제 실패");
  }
}


async function saveGrant(event) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    openLogin();
    return;
  }

  const payload = {
    id: $("#grantEditingId").value,
    route: $("#grantRoute").value,
    grantType: $("#grantType").value,
    grantScope: $("#grantScope").value,
    grantMonth: $("#grantMonth").value,
    occurrenceDates: parseDateInput($("#grantOccurrenceDates").value),
    grantedDays: Number($("#grantDays").value),
    validFrom: $("#grantValidFrom").value,
    validTo: $("#grantValidTo").value,
    eligibilityMode: $("#grantScope").value === "route" ? $("#grantEligibility").value : "all",
    criterionDate: $("#grantCriterionDate").value,
    employeeId: normalizeEmployeeId($("#grantEmployeeId").value),
    excludedEmployeeIds: parseEmployeeIdInput($("#grantExcludedIds").value),
    reason: text($("#grantReason").value),
    note: text($("#grantNote").value),
  };
  if (!payload.grantMonth || !payload.occurrenceDates.length || !(payload.grantedDays > 0) || !payload.validFrom || !payload.validTo) {
    showToast("발생 월·발생일·부여 일수·사용기간을 확인해 주세요.");
    return;
  }
  if (payload.occurrenceDates.some((date) => !/^\d{4}-\d{2}-\d{2}$/.test(date) || !date.startsWith(payload.grantMonth))) {
    showToast("발생일은 발생 월 안의 날짜로 입력해 주세요.");
    return;
  }
  if (payload.id && payload.occurrenceDates.length !== 1) {
    showToast("기존 부여 기록 수정 시 발생일은 하나만 입력해 주세요.");
    return;
  }
  if (payload.validFrom > payload.validTo) {
    showToast("사용 종료일은 시작일보다 빠를 수 없습니다.");
    return;
  }
  if (payload.grantScope === "employee" && !payload.employeeId) {
    showToast("사번별 개별 부여는 대상 사번을 입력해 주세요.");
    return;
  }
  if (payload.eligibilityMode === "worked_on_date" && !payload.criterionDate) {
    showToast("실제 출근 기준일을 입력해 주세요.");
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
    if (response.status === 401) {
      await checkBackend();
      openLogin();
      return;
    }
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "휴가 부여 저장 실패");
    resetGrantForm();
    await Promise.all([loadGrants(), loadHistory()]);
    const savedMessage = `${data.replaced ? "기존 부여 설정을 수정" : `발생일 ${data.createdCount || 1}건을 저장`}했습니다.`;
    showToast(data.recalculateWarning
      ? `${savedMessage} 기존 기록은 유지됐으며 누적 재계산은 월 마감 저장 시 다시 실행됩니다.`
      : `${savedMessage} ${data.affectedMonths || 0}개 월의 대체휴무·보상휴가를 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message);
  } finally {
    button.disabled = false;
    button.textContent = "휴가 부여 저장";
  }
}

async function loadGrants() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    state.grants = [];
    $("#grantTableBody").innerHTML = `<tr><td colspan="15" class="empty-cell">관리자 로그인 후 조회할 수 있습니다.</td></tr>`;
    renderGrantSummary([]);
    return;
  }
  try {
    const route = $("#grantRouteFilter").value;
    const query = route ? `?route=${encodeURIComponent(route)}` : "";
    const response = await fetch(`/api/substitute-grants${query}`, { cache: "no-store" });
    if (response.status === 401) {
      await checkBackend();
      state.grants = [];
      renderGrants([]);
      return;
    }
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "휴가 부여 내역 조회 실패");
    state.grants = data.items || [];
    renderGrants(state.grants);
  } catch (error) {
    state.grants = [];
    renderGrants([]);
    const body = $("#grantTableBody");
    if (body) body.innerHTML = `<tr><td colspan="15" class="empty-cell">조회 오류: ${escapeHtml(error.message || "휴가 부여 내역 조회 실패")}</td></tr>`;
    showToast(error.message || "휴가 부여 내역 조회 실패");
  }
}

function renderGrants(items) {
  renderGrantSummary(items);
  $("#grantTableBody").innerHTML = items.length ? items.map((item) => {
    const status = grantStatus(item);
    const statusClass = status === "사용 가능" ? "success-pill" : status === "만료" ? "status-pill" : "neutral-pill";
    const typeLabel = item.grant_type === "compensation" ? "보상휴가" : "대체휴무";
    const scopeLabel = item.grant_scope === "employee" ? `사번별 · ${item.employee_id || "-"}` : "경로 일괄";
    const criterion = item.settlement_mode && item.grant_type === "compensation"
      ? `${item.occurrence_date || item.criterion_date} 실제 출근 ${number(item.eligible_people || 0)}명 보상휴가 부여 / 미출근 ${number(item.rest_people || 0)}명 기본휴무 +1`
      : item.settlement_mode
        ? `${item.occurrence_date || item.criterion_date} 당일 포함 이전 입사자 전원 동일 부여 / 계획·출근 여부 무관 / 기본휴무 추가 없음`
      : item.eligibility_mode === "worked_on_date"
        ? `${item.criterion_date || item.occurrence_date} 실제 출근자`
        : item.grant_scope === "employee" ? "지정 사번" : `${item.eligibility_cutoff || item.occurrence_date || `${item.grant_month}-01`} 당일 포함 이전 입사자`;
    return `<tr>
      <td>${escapeHtml(ROUTE_LABELS[item.route] || item.route)}</td>
      <td><span class="${item.grant_type === "compensation" ? "warning-pill" : "plan-pill"}">${escapeHtml(typeLabel)}</span></td>
      <td>${escapeHtml(scopeLabel)}</td>
      <td><strong>${escapeHtml(item.grant_month)}</strong></td>
      <td><strong>${escapeHtml(item.occurrence_date || `${item.grant_month}-01`)}</strong></td>
      <td>${formatDays(item.granted_days)}</td>
      <td>${number(item.eligible_people || 0)}명</td>
      <td>${formatDays(item.assigned_days)}</td>
      <td>${formatDays(item.used_days)}</td>
      <td>${formatDays(item.unused_days)}</td>
      <td class="message-cell">${escapeHtml(criterion)}<br><span class="${statusClass}">${escapeHtml(status)}</span></td>
      <td>${number(item.excluded_count || 0)}명</td>
      <td>${escapeHtml(item.valid_from)} ~ ${escapeHtml(item.valid_to)}</td>
      <td class="message-cell">${escapeHtml([item.reason, item.note].filter(Boolean).join(" · "))}</td>
      <td class="action-cell"><button class="btn secondary small grant-edit" data-id="${escapeHtml(item.id)}" type="button">수정</button><button class="btn danger small grant-delete" data-id="${escapeHtml(item.id)}" type="button">삭제</button></td>
    </tr>`;
  }).join("") : `<tr><td colspan="15" class="empty-cell">등록된 대체휴무·보상휴가 부여 내역이 없습니다.</td></tr>`;

  $$(".grant-edit").forEach((button) => button.addEventListener("click", () => editGrant(button.dataset.id)));
  $$(".grant-delete").forEach((button) => button.addEventListener("click", () => deleteGrant(button.dataset.id)));
}

function renderGrantSummary(items) {
  const substituteAssigned = items.filter((item) => item.grant_type !== "compensation").reduce((sum, item) => sum + Number(item.assigned_days || 0), 0);
  const compensationAssigned = items.filter((item) => item.grant_type === "compensation").reduce((sum, item) => sum + Number(item.assigned_days || 0), 0);
  const used = items.reduce((sum, item) => sum + Number(item.used_days || 0), 0);
  const unused = items.reduce((sum, item) => sum + Number(item.unused_days || 0), 0);
  $("#grantSummary").innerHTML = [
    ["부여 설정", `${items.length}건`],
    ["대체휴무 총 부여", formatDays(substituteAssigned)],
    ["보상휴가 총 부여", formatDays(compensationAssigned)],
    ["총 사용", formatDays(used)],
    ["총 미사용", formatDays(unused)],
  ].map(([label, value]) => `<div class="summary-chip"><span>${label}</span><strong>${value}</strong></div>`).join("");
}

function editGrant(id) {
  const item = state.grants.find((grant) => grant.id === id);
  if (!item) return;
  $("#grantEditingId").value = item.id;
  $("#grantRoute").value = item.route;
  $("#grantType").value = item.grant_type || "substitute";
  $("#grantScope").value = item.grant_scope || "route";
  $("#grantMonth").value = item.grant_month;
  $("#grantOccurrenceDates").value = item.occurrence_date || `${item.grant_month}-01`;
  $("#grantDays").value = Number(item.granted_days || 0);
  $("#grantEmployeeId").value = item.employee_id || "";
  $("#grantValidFrom").value = item.valid_from;
  $("#grantValidTo").value = item.valid_to;
  $("#grantEligibility").value = item.eligibility_mode || "all";
  $("#grantCriterionDate").value = item.criterion_date || "";
  $("#grantExcludedIds").value = parseStoredEmployeeIds(item.excluded_employee_ids_json).join("\n");
  $("#grantReason").value = item.reason || "";
  $("#grantNote").value = item.note || "";
  $("#grantSubmitButton").textContent = "부여 설정 수정 저장";
  $("#grantCancelEdit").classList.remove("hidden");
  syncGrantFormVisibility();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetGrantForm() {
  $("#grantEditingId").value = "";
  $("#grantType").value = "substitute";
  $("#grantScope").value = "route";
  $("#grantOccurrenceDates").value = $("#grantMonth").value ? `${$("#grantMonth").value}-01` : "";
  $("#grantDays").value = "1";
  $("#grantEmployeeId").value = "";
  $("#grantEligibility").value = "all";
  $("#grantCriterionDate").value = "";
  $("#grantExcludedIds").value = "";
  $("#grantReason").value = "";
  $("#grantNote").value = "";
  $("#grantSubmitButton").textContent = "휴가 부여 저장";
  $("#grantCancelEdit").classList.add("hidden");
  syncGrantFormVisibility();
}

async function deleteGrant(id) {
  if (!confirm("이 대체휴무·보상휴가 부여 설정을 삭제하시겠습니까? 관련 월의 잔여·초과·이월이 다시 계산됩니다.")) return;
  try {
    const response = await fetch(`/api/substitute-grants/${encodeURIComponent(id)}`, { method: "DELETE" });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "삭제 실패");
    resetGrantForm();
    await Promise.all([loadGrants(), loadHistory()]);
    showToast(`부여 설정을 삭제하고 ${data.affectedMonths || 0}개 월을 다시 계산했습니다.`);
  } catch (error) {
    showToast(error.message);
  }
}

function parseDateInput(value) {
  return [...new Set(String(value || "").split(/[\s,;]+/).map((item) => item.trim()).filter(Boolean))];
}

function parseEmployeeIdInput(value) {
  return [...new Set(String(value || "").split(/[\s,;]+/).map(normalizeEmployeeId).filter(Boolean))];
}

function parseStoredEmployeeIds(value) {
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed.map(normalizeEmployeeId).filter(Boolean) : [];
  } catch {
    return [];
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



async function loadPersonnelChecks() {
  const body = $("#personnelCheckTableBody");
  const summaryNode = $("#personnelCheckSummary");
  if (!body || !summaryNode) return;
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    body.innerHTML = '<tr><td colspan="15" class="empty-cell">관리자 로그인 후 조회할 수 있습니다.</td></tr>';
    summaryNode.innerHTML = "";
    return;
  }
  const month = $("#personnelCheckMonth").value || $("#targetMonth").value;
  const route = $("#personnelCheckRoute").value || "";
  if (!month) return;
  body.innerHTML = '<tr><td colspan="15" class="empty-cell">인력현황과 연차대장을 비교 중입니다.</td></tr>';
  const data = await fetchPersonnelChecks(month, route);
  state.personnelChecks = data;
  const summary = data.summary || {};
  summaryNode.innerHTML = [
    ["확인 요청", summary.unresolved || 0],
    ["인력현황 미등록", summary.workforceMissing || 0],
    ["연차대장 미등록", summary.annualMissing || 0],
    ["경로 불일치", summary.routeMismatch || 0],
    ["처리 완료", Math.max(0, Number(summary.total || 0) - Number(summary.unresolved || 0))],
  ].map(([label, value]) => `<div><span>${escapeHtml(label)}</span><strong>${number(value)}명</strong></div>`).join("");
  renderPersonnelCheckRows(data.items || []);
}

function personnelStatusOptions(selected) {
  return ["확인 요청", "재직·포함", "퇴사", "경로이동", "육아휴직", "기타휴직", "제외"]
    .map((item) => `<option value="${item}" ${item === selected ? "selected" : ""}>${item}</option>`).join("");
}

function destinationRouteOptions(selected) {
  return [
    ["", "-"] , ["homeplus", "홈플러스"], ["electroland", "전자랜드"],
  ].map(([value, label]) => `<option value="${value}" ${value === selected ? "selected" : ""}>${label}</option>`).join("");
}

function renderPersonnelCheckRows(items) {
  const body = $("#personnelCheckTableBody");
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="15" class="empty-cell">인력현황과 연차대장이 일치하며 별도 확인 대상이 없습니다.</td></tr>';
    return;
  }
  body.innerHTML = items.map((item, index) => `
    <tr class="${item.resolved ? "resolved-row" : "request-row"}" data-index="${index}" data-route="${escapeHtml(item.route)}" data-employee-id="${escapeHtml(item.employeeId)}">
      <td class="route-text">${escapeHtml(ROUTE_LABELS[item.route] || item.route)}</td>
      <td><span class="${item.resolved ? "success-pill" : "warning-pill"}">${escapeHtml(item.issueType || "확인 요청")}</span></td>
      <td>${escapeHtml(item.regionalManager || "-")}</td>
      <td>${escapeHtml(item.manager || "-")}</td>
      <td>${escapeHtml(item.storeName || "-")}</td>
      <td>${escapeHtml(item.employeeId)}</td>
      <td><strong>${escapeHtml(item.employeeName || "-")}</strong></td>
      <td>${escapeHtml(ROUTE_LABELS[item.workforceRoute] || item.workforceRoute || "없음")}</td>
      <td>${escapeHtml(ROUTE_LABELS[item.annualRoute] || item.annualRoute || "없음")}</td>
      <td><select class="input personnel-status">${personnelStatusOptions(item.personnelStatus || "확인 요청")}</select></td>
      <td><input class="input personnel-from" type="date" value="${escapeHtml(item.effectiveFrom || "")}" /></td>
      <td><input class="input personnel-to" type="date" value="${escapeHtml(item.effectiveTo || "")}" /></td>
      <td><select class="input personnel-destination">${destinationRouteOptions(item.destinationRoute || "")}</select></td>
      <td><input class="input personnel-note" value="${escapeHtml(item.note || "")}" placeholder="퇴사일·이동·휴직 사유" /></td>
      <td>${item.resolved ? '<span class="success-pill">처리 완료</span>' : '<span class="warning-pill">확인 요청</span>'}</td>
    </tr>
  `).join("");
}

async function savePersonnelChecks() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    openLogin();
    return;
  }
  const month = $("#personnelCheckMonth").value || $("#targetMonth").value;
  const sourceItems = state.personnelChecks?.items || [];
  const items = [...$("#personnelCheckTableBody").querySelectorAll("tr[data-index]")].map((row) => {
    const source = sourceItems[Number(row.dataset.index)] || {};
    return {
      month,
      route: row.dataset.route,
      employeeId: row.dataset.employeeId,
      employeeName: source.employeeName || "",
      issueType: source.issueType || "인력 변동 직접 입력",
      personnelStatus: row.querySelector(".personnel-status")?.value || "확인 요청",
      effectiveFrom: row.querySelector(".personnel-from")?.value || "",
      effectiveTo: row.querySelector(".personnel-to")?.value || "",
      destinationRoute: row.querySelector(".personnel-destination")?.value || "",
      note: row.querySelector(".personnel-note")?.value || "",
      sourceType: "manual",
    };
  });
  if (!items.length) {
    showToast("저장할 인력 확인 대상이 없습니다.");
    return;
  }
  const button = $("#savePersonnelChecks");
  try {
    button.disabled = true;
    button.textContent = "인력 변동 저장 중...";
    const response = await fetch("/api/personnel-checks", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ month, items, sourceType: "manual" }),
    });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "인력 변동 저장 실패");
    await loadPersonnelChecks();
    showToast(`${month} 인력 변동 ${number(data.saved)}명을 저장했습니다.`);
  } catch (error) {
    showToast(error.message || "인력 변동 저장 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "수기 변동 저장";
  }
}

async function persistPersonnelOverrides(items, month) {
  const normalized = (items || []).filter((item) => item.employeeId && item.route && item.personnelStatus && item.personnelStatus !== "확인 요청");
  if (!normalized.length) return null;
  const response = await fetch("/api/personnel-checks", {
    method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ month, items: normalized, sourceType: "evidence" }),
  });
  const data = await readJsonResponse(response);
  if (!response.ok) throw new Error(data.error || "증빙 파일의 인력 변동 저장 실패");
  return data;
}

async function previewWorkforceFile() {
  const file = $("#workforceFile")?.files?.[0];
  const reference = $("#portalReferenceFile")?.files?.[0];
  const textNode = $("#workforcePreview");
  if (!textNode) return;
  if (!file) {
    textNode.textContent = "인력 및 매장매칭 엑셀을 선택해 주세요.";
    return;
  }
  textNode.textContent = `선택 파일: ${file.name}${reference ? ` · 포탈사번 참고: ${reference.name}` : ""}`;
}

async function saveWorkforceFile(event) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    openLogin();
    return;
  }
  const file = $("#workforceFile").files?.[0];
  const month = $("#workforceMonth").value;
  const referenceFile = $("#portalReferenceFile").files?.[0] || null;
  if (!file || !month) {
    showToast("적용 월과 인력 및 매장매칭 엑셀을 선택해 주세요.");
    return;
  }

  const button = $("#workforceSubmitButton");
  try {
    button.disabled = true;
    button.textContent = "파일 분석·저장 중...";
    const sheets = await fileToWorkbookSheets(file);
    const members = parseWorkforceWorkbook(sheets);
    if (!members.length) throw new Error("인력DB와 매장매칭 시트에서 유효한 사번을 찾지 못했습니다.");
    let portalMappings = [];
    if (referenceFile) {
      const referenceSheets = await fileToWorkbookSheets(referenceFile);
      portalMappings = parsePortalMappings(referenceSheets);
    }

    const memberRoutes = new Map(members.map((item) => [normalizeEmployeeId(item.employeeId), item.route]));
    const memberStatusRows = members
      .filter((item) => item.personnelStatus && item.personnelStatus !== "재직·포함" && item.personnelStatus !== "확인 요청")
      .map((item) => ({
        route: item.route, employeeId: item.employeeId, employeeName: item.employeeName,
        issueType: "인력현황 변동 입력", personnelStatus: item.personnelStatus,
        effectiveFrom: item.effectiveFrom || "", effectiveTo: item.effectiveTo || "",
        destinationRoute: item.destinationRoute || "", note: item.statusNote || item.note || "",
      }));
    const statusSheetRows = parsePersonnelOverridesFromWorkbook(sheets, month, "").map((item) => ({
      ...item, route: item.route || memberRoutes.get(normalizeEmployeeId(item.employeeId)) || "",
    }));
    const personnelOverrides = [...new Map([...memberStatusRows, ...statusSheetRows]
      .filter((item) => item.route && item.employeeId)
      .map((item) => [`${item.route}|${normalizeEmployeeId(item.employeeId)}`, item])).values()];

    const form = new FormData();
    form.append("file", file, file.name);
    form.append("month", month);
    form.append("members", JSON.stringify(members));
    form.append("portalMappings", JSON.stringify(portalMappings));
    form.append("personnelOverrides", JSON.stringify(personnelOverrides));
    const response = await fetch("/api/workforce", { method: "POST", body: form });
    if (response.status === 401) {
      await checkBackend();
      openLogin();
      return;
    }
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "인력·매장매칭 저장 실패");
    $("#workforceFile").value = "";
    $("#portalReferenceFile").value = "";
    $("#workforcePreview").textContent = "저장할 새 파일을 선택해 주세요.";
    state.workforce = null;
    $("#personnelCheckMonth").value = month;
    await Promise.all([loadWorkforceUploads(), loadPersonnelChecks()]);
    showToast(`${month} 인력·매장매칭을 ${data.replaced ? "교체" : "저장"}했습니다. 전자랜드 ${number(data.electrolandCount)}명 · 홈플러스 ${number(data.homeplusCount)}명 · 포탈사번 ${number(data.portalCount)}명 · 인력변동 ${number(data.personnelStatusCount || 0)}명`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "인력·매장매칭 저장 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "월별 인력·매장매칭 저장";
  }
}

function parseWorkforceWorkbook(sheets) {
  const normalizedSheets = normalizeWorkbookInput(sheets);
  const matchingSheet = normalizedSheets.find((sheet) => normalizeHeader(sheet.sheetName).includes("매장매칭"));
  if (!matchingSheet) throw new Error("‘매장매칭’ 시트를 찾지 못했습니다.");
  const storeMap = parseStoreMatchingSheet(matchingSheet);
  const output = [];
  const definitions = [
    { route: "electroland", keywords: ["랜드인력DB", "전자랜드인력DB"] },
    { route: "homeplus", keywords: ["홈플인력DB", "홈플러스인력DB"] },
  ];
  for (const definition of definitions) {
    const sheet = normalizedSheets.find((candidate) => definition.keywords.some((keyword) => normalizeHeader(candidate.sheetName).includes(keyword)));
    if (!sheet) continue;
    output.push(...parseWorkforcePeopleSheet(sheet, definition.route, storeMap));
  }
  if (!output.length) throw new Error("‘랜드 인력DB’ 또는 ‘홈플 인력DB’에서 유효한 사번을 찾지 못했습니다.");
  return output;
}

function parseStoreMatchingSheet(sheet) {
  const matrix = sheet.matrix || [];
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
    findHeaderIndex(headers, ["매장코드"]) >= 0
    && findHeaderIndex(headers, ["경로"]) >= 0
    && findHeaderIndex(headers, ["지역장"]) >= 0
  ));
  if (headerIndex < 0) throw new Error("매장매칭 시트의 매장코드·경로·지역장 머리글을 찾지 못했습니다.");
  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const columns = {
    storeCode: findHeaderIndex(headers, ["매장코드"]),
    region2: findHeaderIndex(headers, ["지역", "지역2"]),
    route: findHeaderIndex(headers, ["경로"]),
    storeName: findHeaderIndex(headers, ["매장", "매장명"]),
    regionalManager: findHeaderIndex(headers, ["지역장"]),
    manager: findHeaderIndex(headers, ["매니저"]),
    closedDate: findHeaderIndex(headers, ["폐점날짜", "폐점일"]),
    note: findHeaderIndex(headers, ["비고"]),
  };
  const map = new Map();
  for (const row of matrix.slice(headerIndex + 1)) {
    const code = normalizeStoreCode(row[columns.storeCode]);
    if (!code) continue;
    map.set(code, {
      route: routeValue(row[columns.route]),
      region2: text(row[columns.region2]),
      storeName: text(row[columns.storeName]),
      regionalManager: text(row[columns.regionalManager]),
      manager: text(row[columns.manager]),
      closedDate: columns.closedDate >= 0 ? parseDateCell(row[columns.closedDate]) : "",
      note: columns.note >= 0 ? text(row[columns.note]) : "",
    });
  }
  return map;
}

function parseWorkforcePeopleSheet(sheet, route, storeMap) {
  const matrix = sheet.matrix || [];
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
    findHeaderIndex(headers, ["사번", "사원번호"]) >= 0
    && findHeaderIndex(headers, ["성명", "이름"]) >= 0
    && findHeaderIndex(headers, ["매장코드"]) >= 0
  ));
  if (headerIndex < 0) throw new Error(`${sheet.sheetName}의 사번·성명·매장코드 머리글을 찾지 못했습니다.`);
  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const columns = {
    manager: findHeaderIndex(headers, ["매니저"]),
    region1: findHeaderIndex(headers, ["지역", "지역1"]),
    storeCode: findHeaderIndex(headers, ["매장코드"]),
    storeName: findHeaderIndex(headers, ["매장명", "매장"]),
    employeeId: findHeaderIndex(headers, ["사번", "사원번호"]),
    employeeName: findHeaderIndex(headers, ["성명", "이름"]),
    hireDate: findHeaderIndex(headers, ["제니엘입사일", "입사일"]),
    groupHireDate: findHeaderIndex(headers, ["그룹입사일"]),
    portalId: findHeaderIndex(headers, ["포탈사번", "포털사번"]),
    employmentStatus: findHeaderIndex(headers, ["인력상태", "재직상태", "재직여부", "상태", "변동구분"]),
    terminationDate: findHeaderIndex(headers, ["퇴사일", "퇴직일"]),
    statusStartDate: findHeaderIndex(headers, ["적용일", "변동일", "이동일", "휴직시작일"]),
    statusEndDate: findHeaderIndex(headers, ["복직일", "휴직종료일", "종료일"]),
    destinationRoute: findHeaderIndex(headers, ["이동경로", "변경경로", "전환경로", "이동처"]),
    note: findHeaderIndex(headers, ["비고", "휴퇴사일", "휴/퇴사일"]),
  };
  const rows = [];
  for (const row of matrix.slice(headerIndex + 1)) {
    const employeeId = normalizeEmployeeId(row[columns.employeeId]);
    const employeeName = text(row[columns.employeeName]);
    if (!looksLikeEmployeeId(employeeId) || !employeeName) continue;
    const storeCode = normalizeStoreCode(row[columns.storeCode]);
    const matched = storeMap.get(storeCode) || {};
    const matchedRoute = matched.route || route;
    if (matchedRoute && matchedRoute !== route) continue;
    const rawNote = columns.note >= 0 ? text(row[columns.note]) : "";
    const rawStatus = columns.employmentStatus >= 0 ? text(row[columns.employmentStatus]) : "";
    const terminationDate = columns.terminationDate >= 0 ? parseDateCell(row[columns.terminationDate]) : "";
    const statusStartDate = columns.statusStartDate >= 0 ? parseDateCell(row[columns.statusStartDate]) : "";
    const statusEndDate = columns.statusEndDate >= 0 ? parseDateCell(row[columns.statusEndDate]) : "";
    const statusText = [rawStatus, rawNote, terminationDate ? "퇴사" : ""].filter(Boolean).join(" ");
    const personnelStatus = normalizePersonnelStatus(statusText);
    const notes = [rawNote, matched.closedDate ? `${matched.closedDate} 폐점` : "", matched.note || ""].filter(Boolean);
    rows.push({
      route,
      regionalManager: matched.regionalManager || "",
      manager: matched.manager || (columns.manager >= 0 ? text(row[columns.manager]) : ""),
      region1: columns.region1 >= 0 ? text(row[columns.region1]) : "",
      region2: matched.region2 || "",
      storeCode,
      storeName: text(row[columns.storeName]) || matched.storeName || "",
      portalId: columns.portalId >= 0 ? text(row[columns.portalId]) : "",
      employeeId,
      employeeName,
      hireDate: columns.hireDate >= 0 ? parseDateCell(row[columns.hireDate]) : "",
      groupHireDate: columns.groupHireDate >= 0 ? parseDateCell(row[columns.groupHireDate]) : "",
      personnelStatus,
      effectiveFrom: terminationDate || statusStartDate,
      effectiveTo: statusEndDate,
      destinationRoute: columns.destinationRoute >= 0 ? routeValue(row[columns.destinationRoute]) : "",
      statusNote: rawNote,
      note: [...new Set(notes)].join(" · "),
    });
  }
  return rows;
}

function parsePortalMappings(sheets) {
  const mappings = [];
  for (const sheet of normalizeWorkbookInput(sheets)) {
    const matrix = sheet.matrix || [];
    const headerIndex = findFlexibleHeaderRow(matrix, (headers) => (
      findHeaderIndex(headers, ["포탈사번", "포털사번"]) >= 0
      && findHeaderIndex(headers, ["사번", "사원번호"]) >= 0
    ));
    if (headerIndex < 0) continue;
    const headers = (matrix[headerIndex] || []).map(normalizeHeader);
    const portalCol = findHeaderIndex(headers, ["포탈사번", "포털사번"]);
    const employeeCol = findHeaderIndex(headers, ["사번", "사원번호"]);
    const nameCol = findHeaderIndex(headers, ["성명", "이름"]);
    for (const row of matrix.slice(headerIndex + 1)) {
      const employeeId = normalizeEmployeeId(row[employeeCol]);
      const portalId = text(row[portalCol]);
      if (!looksLikeEmployeeId(employeeId) || !portalId) continue;
      mappings.push({ employeeId, portalId, employeeName: nameCol >= 0 ? text(row[nameCol]) : "" });
    }
  }
  return [...new Map(mappings.map((item) => [item.employeeId, item])).values()];
}

function normalizeStoreCode(value) {
  const raw = text(value).replace(/\.0+$/, "").replace(/\s+/g, "");
  return raw;
}

function routeValue(value) {
  const raw = normalizeHeader(value);
  if (raw.includes("홈플")) return "homeplus";
  if (raw.includes("전자랜드") || raw === "랜드") return "electroland";
  return "";
}

async function loadWorkforceMonth(month, route) {
  if (state.workforce?.month === month && state.workforce?.route === route) return state.workforce;
  const response = await fetch(`/api/workforce?month=${encodeURIComponent(month)}&route=${encodeURIComponent(route)}`, { cache: "no-store" });
  if (response.status === 401) {
    await checkBackend();
    openLogin();
    throw new Error("관리자 로그인이 필요합니다.");
  }
  const data = await readJsonResponse(response);
  if (!response.ok) throw new Error(data.error || "인력·매장매칭 조회 실패");
  state.workforce = { ...data, month, route };
  return state.workforce;
}

async function loadWorkforceUploads() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    state.workforceUploads = [];
    renderWorkforceUploads([]);
    updateWorkforceStatus();
    return;
  }
  try {
    const response = await fetch("/api/workforce", { cache: "no-store" });
    if (response.status === 401) {
      await checkBackend();
      renderWorkforceUploads([]);
      return;
    }
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "인력·매장매칭 목록 조회 실패");
    state.workforceUploads = data.items || [];
    renderWorkforceUploads(state.workforceUploads);
    updateWorkforceStatus();
  } catch (error) {
    renderWorkforceUploads([]);
    showToast(error.message || "인력·매장매칭 목록 조회 실패");
  }
}

function renderWorkforceUploads(items) {
  const body = $("#workforceTableBody");
  if (!body) return;
  body.innerHTML = items.length ? items.map((item) => `<tr>
    <td><strong>${escapeHtml(item.month)}</strong></td>
    <td>${escapeHtml(item.file_name || "-")}</td>
    <td>${number(item.electroland_count || 0)}명</td>
    <td>${number(item.homeplus_count || 0)}명</td>
    <td>${number(item.portal_count || 0)}명</td>
    <td>${formatFileSize(item.size_bytes || 0)}</td>
    <td>${escapeHtml(formatStoredDate(item.updated_at || item.created_at))}</td>
    <td class="action-cell"><a class="btn secondary small" href="/api/workforce/${encodeURIComponent(item.month)}" download>원본 다운로드</a><button class="btn danger small workforce-delete" data-month="${escapeHtml(item.month)}" type="button">삭제</button></td>
  </tr>`).join("") : `<tr><td colspan="8" class="empty-cell">등록된 월별 인력·매장매칭 파일이 없습니다.</td></tr>`;
  $$(".workforce-delete").forEach((button) => button.addEventListener("click", () => deleteWorkforceMonth(button.dataset.month)));
  const summary = $("#workforceSummary");
  if (summary) {
    const latest = items[0];
    summary.innerHTML = [
      ["등록 월", `${items.length}개`],
      ["최근 적용 월", latest?.month || "-"],
      ["전자랜드 인원", latest ? `${number(latest.electroland_count || 0)}명` : "-"],
      ["홈플러스 인원", latest ? `${number(latest.homeplus_count || 0)}명` : "-"],
      ["포탈사번 보유", latest ? `${number(latest.portal_count || 0)}명` : "-"],
    ].map(([label, value]) => `<div class="summary-chip"><span>${label}</span><strong>${escapeHtml(value)}</strong></div>`).join("");
  }
}

async function deleteWorkforceMonth(month) {
  if (!confirm(`${month} 인력·매장매칭 파일과 직원 매칭 정보를 삭제하시겠습니까?`)) return;
  try {
    const response = await fetch(`/api/workforce/${encodeURIComponent(month)}`, { method: "DELETE" });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "삭제 실패");
    state.workforce = null;
    await loadWorkforceUploads();
    showToast(`${month} 인력·매장매칭을 삭제했습니다.`);
  } catch (error) {
    showToast(error.message || "인력·매장매칭 삭제 실패");
  }
}

function updateWorkforceStatus() {
  const target = $("#workforceStatus");
  if (!target) return;
  const month = $("#targetMonth")?.value;
  const route = selectedRoute();
  const item = state.workforceUploads.find((row) => row.month === month);
  if (!item) {
    target.className = "alert warning matching-status";
    target.textContent = `${month || "대상 월"} 인력·매장매칭이 등록되지 않았습니다. 인력 매칭에서 먼저 저장해 주세요.`;
    return;
  }
  const count = route === "homeplus" ? item.homeplus_count : item.electroland_count;
  target.className = "alert success matching-status";
  target.textContent = `${month} ${ROUTE_LABELS[route]} 인력·매장매칭 ${number(count || 0)}명 등록됨 · 상담사근태 B~J 자동 반영`;
}

async function checkBackend() {
  try {
    const response = await fetch("/api/auth", { cache: "no-store" });
    if (!response.ok) throw new Error();
    const data = await response.json();
    state.backend = {
      available: true,
      configured: Boolean(data.configured),
      loggedIn: Boolean(data.loggedIn),
      fileStorageConfigured: Boolean(data.fileStorageConfigured),
      fileStorageMode: data.fileStorageMode === "r2" ? "r2" : "d1",
    };
  } catch {
    state.backend = { available: false, configured: false, loggedIn: false, fileStorageConfigured: false, fileStorageMode: "d1" };
  }
  updateStorageUI();
}

function updateStorageUI() {
  const badge = $("#storageBadge");
  const loginButton = $("#loginButton");
  const logoutButton = $("#logoutButton");
  const fileNotice = $("#fileStorageNotice");
  if (state.backend.available && state.backend.configured) {
    badge.className = "badge cloud";
    badge.textContent = state.backend.loggedIn ? "D1 영구 저장 연결" : "D1 로그인 필요";
    loginButton.classList.toggle("hidden", state.backend.loggedIn);
    logoutButton.classList.toggle("hidden", !state.backend.loggedIn);
    if (fileNotice) {
      fileNotice.className = "alert success file-storage-notice";
      fileNotice.textContent = state.backend.fileStorageConfigured
        ? "대용량 파일 보관용 R2가 연결되어 있습니다. 월 마감 원본과 수정본을 용량 제한 없이 보관할 수 있습니다."
        : "현재는 기존 D1에 파일을 여러 조각으로 나누어 보관합니다. 파일당 최대 20MB까지 별도 설정 없이 저장할 수 있습니다.";
    }
  } else {
    badge.className = "badge local";
    badge.textContent = "서버 설정 확인 필요";
    loginButton.classList.add("hidden");
    logoutButton.classList.add("hidden");
    if (fileNotice) {
      fileNotice.className = "alert warning file-storage-notice";
      fileNotice.textContent = "D1 연결과 관리자 환경 변수를 확인해야 파일을 저장할 수 있습니다.";
    }
  }
}


function syncAnnualBaselineGuide() {
  // v34부터 홈플러스·전자랜드 전용 업로드칸에 안내문이 고정 표시됩니다.
}

async function previewAnnualBaselineForRoute(route) {
  const file = $(`#${route}AnnualBaselineFile`)?.files?.[0];
  const previewBox = $(`#${route}AnnualBaselinePreview`);
  if (!file) {
    state.annualBaselinePreviews[route] = null;
    previewBox.textContent = `${ROUTE_LABELS[route]} 연차대장을 선택해 주세요.`;
    return;
  }
  try {
    previewBox.textContent = "연차 기준대장 확인 중...";
    const sheets = await fileToWorkbookSheets(file);
    const baselineDate = $(`#${route}AnnualBaselineDate`).value;
    const parsed = parseInitialAnnualLedger(sheets, route, baselineDate);
    const employees = parsed.employees;
    state.annualBaselinePreviews[route] = { employees, file, route, baselineDate, meta: parsed.meta };
    const under = employees.filter((item) => item.underOneYear).length;
    const remaining = roundHalf(employees.reduce((sum, item) => sum + Number(item.baselineRemaining || 0), 0));
    const excludedText = parsed.meta.excludedCount ? ` · 제외 ${number(parsed.meta.excludedCount)}명` : "";
    previewBox.innerHTML = `<strong>${escapeHtml(file.name)}</strong><p><span class="comparison-badge match">${escapeHtml(parsed.meta.formatLabel)}</span> · 읽은 시트 ${escapeHtml(parsed.meta.sheetNames.join(", "))}</p><p>반영 ${number(employees.length)}명 · 1년 이상 ${number(employees.length - under)}명 · 1년 미만 ${number(under)}명 · 기준 잔여 ${number(remaining)}일${excludedText}</p><p>${escapeHtml(parsed.meta.valueRule)}</p>`;
  } catch (error) {
    state.annualBaselinePreviews[route] = null;
    previewBox.textContent = error.message || `${ROUTE_LABELS[route]} 연차대장을 읽지 못했습니다.`;
  }
}

async function previewAnnualMonthlyForRoute(route) {
  const file = $(`#${route}AnnualMonthlyFile`)?.files?.[0];
  const monthInput = $(`#${route}AnnualMonthlyMonth`);
  const previewBox = $(`#${route}AnnualMonthlyPreview`);
  const month = monthInput?.value || "";
  if (!file || !month) {
    state.annualMonthlyPreviews[route] = null;
    previewBox.textContent = `${ROUTE_LABELS[route]} 대상 월과 승인·반려 파일을 선택해 주세요.`;
    return;
  }
  try {
    previewBox.textContent = "승인·반려 현황 확인 중...";
    const sheets = await fileToWorkbookSheets(file);
    const parsed = parseAnnualApplications(sheets, "");
    const annualRows = parsed.rows.filter((row) => ["연차", "반차"].includes(row.requestedKind));
    const monthCounts = new Map();
    for (const row of annualRows) {
      const key = String(row.date || "").slice(0, 7);
      if (/^\d{4}-\d{2}$/.test(key)) monthCounts.set(key, (monthCounts.get(key) || 0) + 1);
    }
    let resolvedMonth = month;
    let applications = annualRows.filter((row) => String(row.date || "").startsWith(resolvedMonth));
    if (!applications.length && monthCounts.size) {
      resolvedMonth = [...monthCounts.entries()].sort((a, b) => b[1] - a[1])[0][0];
      applications = annualRows.filter((row) => String(row.date || "").startsWith(resolvedMonth));
      monthInput.value = resolvedMonth;
    }
    if (!applications.length) throw new Error(`${month} 연차·반차 신청내역을 찾지 못했습니다.`);
    state.annualMonthlyPreviews[route] = { applications, file, route, month: resolvedMonth };
    const approved = roundHalf(applications.filter(isApprovedAnnualApplication).reduce((sum, row) => sum + row.requestedDays, 0));
    const rejected = roundHalf(applications.filter((row) => /반려|미사용|취소/.test(normalizeHeader(row.applicationStatus))).reduce((sum, row) => sum + row.requestedDays, 0));
    const requestCount = new Set(applications.map((row) => row.sourceIndex)).size;
    previewBox.innerHTML = `<strong>${escapeHtml(file.name)}</strong><p>${ROUTE_LABELS[route]} · 대상 월 ${escapeHtml(resolvedMonth)} · 신청 ${number(requestCount)}건 · 사용일 ${number(applications.length)}건 · 승인 ${number(approved)}일 · 반려·미사용 ${number(rejected)}일</p>`;
  } catch (error) {
    state.annualMonthlyPreviews[route] = null;
    previewBox.textContent = error.message || `${ROUTE_LABELS[route]} 월별 연차 파일을 읽지 못했습니다.`;
  }
}

async function saveAnnualBaselineForRoute(event, route) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) return openLogin();
  await previewAnnualBaselineForRoute(route);
  const preview = state.annualBaselinePreviews[route];
  if (!preview?.employees?.length) return showToast(`${ROUTE_LABELS[route]} 저장 자료를 확인해 주세요.`);
  const button = $(`#${route}AnnualBaselineSubmit`);
  try {
    button.disabled = true; button.textContent = "기준값 저장 중...";
    const response = await fetch("/api/annual-leave", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ action: "baseline", route, baselineDate: preview.baselineDate, fileName: preview.file.name, employees: preview.employees }) });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "연차 기준값 저장 실패");
    await uploadArchiveFile({ file: preview.file, route, month: preview.baselineDate.slice(0, 7), fileKind: "other", note: `${ROUTE_LABELS[route]} 전용 기존 연차대장 · 수기 수정값 반영`, sourceType: "manual", replace: false }).catch(() => null);
    $("#annualLedgerRoute").value = route; $("#annualLedgerMonth").value = preview.baselineDate.slice(0, 7); $("#personnelCheckMonth").value = preview.baselineDate.slice(0, 7);
    await Promise.all([loadAnnualLeaveDashboard(), loadPersonnelChecks()]);
    showToast(data.replaced ? `${ROUTE_LABELS[route]} 기준대장을 수정 파일 기준으로 교체했습니다.` : `${ROUTE_LABELS[route]} 기준대장을 등록했습니다.`);
  } catch (error) { showToast(error.message || "연차 기준값 저장 중 오류가 발생했습니다."); }
  finally { button.disabled = false; button.textContent = `${ROUTE_LABELS[route]} 기준대장 저장·교체`; }
}

async function saveAnnualMonthlyForRoute(event, route) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) return openLogin();
  await previewAnnualMonthlyForRoute(route);
  const preview = state.annualMonthlyPreviews[route];
  if (!preview?.applications?.length) return showToast(`${ROUTE_LABELS[route]} 월별 자료를 확인해 주세요.`);
  const button = $(`#${route}AnnualMonthlySubmit`);
  try {
    button.disabled = true; button.textContent = "월별 현황 저장 중...";
    const applications = preview.applications.map((row) => ({ employeeId: row.employeeId, employeeName: row.name, leaveDate: row.date, days: row.requestedDays, status: row.applicationStatus, leaveType: row.leaveType || row.requestedKind, applicationDate: row.applicationDate, note: row.note, sourceIndex: row.sourceIndex }));
    const response = await fetch("/api/annual-leave", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ action: "monthly", route, month: preview.month, fileName: preview.file.name, applications }) });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "월별 연차 현황 저장 실패");
    await uploadArchiveFile({ file: preview.file, route, month: preview.month, fileKind: "other", note: `${ROUTE_LABELS[route]} 월별 연차 승인·반려 원본`, sourceType: "manual", replace: false }).catch(() => null);
    $("#annualLedgerRoute").value = route; $("#annualLedgerMonth").value = preview.month; await loadAnnualLeaveDashboard();
    showToast(data.replaced ? `${ROUTE_LABELS[route]} ${preview.month} 자료를 교체했습니다.` : `${ROUTE_LABELS[route]} ${preview.month} 자료를 반영했습니다.`);
  } catch (error) { showToast(error.message || "월별 연차 현황 저장 중 오류가 발생했습니다."); }
  finally { button.disabled = false; button.textContent = `${ROUTE_LABELS[route]} 월별 현황 저장·교체`; }
}

async function fetchAnnualLeaveDashboard(route, month, asOf = "") {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) return null;
  const params = new URLSearchParams({ route });
  if (month) params.set("month", month);
  if (asOf) params.set("asOf", asOf);
  const response = await fetch(`/api/annual-leave?${params}`, { cache: "no-store" });
  if (response.status === 401) return null;
  const data = await readJsonResponse(response);
  if (!response.ok) throw new Error(data.error || "연차 누적현황 조회 실패");
  return data;
}

async function loadAnnualLeaveDashboard() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    state.annualLeaveDashboard = null;
    renderAnnualLeaveDashboard(null);
    return;
  }
  try {
    const route = $("#annualLedgerRoute")?.value || "electroland";
    const month = $("#annualLedgerMonth")?.value || $("#targetMonth")?.value || "";
    const data = await fetchAnnualLeaveDashboard(route, month);
    state.annualLeaveDashboard = data;
    renderAnnualLeaveDashboard(data);
  } catch (error) {
    state.annualLeaveDashboard = null;
    renderAnnualLeaveDashboard(null, error.message || "연차 누적현황 조회 실패");
    showToast(error.message || "연차 누적현황 조회 실패");
  }
}

function renderAnnualLeaveDashboard(data, errorMessage = "") {
  const status = $("#annualLedgerStatus");
  const summary = $("#annualLedgerSummary");
  const employeeBody = $("#annualLedgerTableBody");
  const reminderBody = $("#annualReminderTableBody");
  const historyBody = $("#annualUploadHistoryBody");
  if (!status || !summary || !employeeBody || !reminderBody || !historyBody) return;
  if (!data) {
    status.className = "alert warning";
    status.textContent = errorMessage || "관리자 로그인 후 연차 누적현황을 조회할 수 있습니다.";
    summary.innerHTML = "";
    employeeBody.innerHTML = '<tr><td colspan="15" class="empty-cell">조회할 연차 자료가 없습니다.</td></tr>';
    reminderBody.innerHTML = '<tr><td colspan="10" class="empty-cell">표시할 촉진 대상이 없습니다.</td></tr>';
    historyBody.innerHTML = '<tr><td colspan="8" class="empty-cell">등록 이력이 없습니다.</td></tr>';
    return;
  }
  if (data.baseline) {
    status.className = "alert success";
    status.textContent = `${ROUTE_LABELS[data.route]} 최초 기준: ${data.baseline.baseline_date} · ${data.baseline.file_name || "파일명 없음"} · ${number(data.baseline.employee_count)}명`;
  } else {
    status.className = "alert warning";
    status.textContent = `${ROUTE_LABELS[data.route]} 연차 누적현황 초본이 아직 등록되지 않았습니다.`;
  }
  const s = data.summary || {};
  summary.innerHTML = `
    <div><span>관리 인원</span><strong>${number(s.employeeCount || 0)}</strong></div>
    <div><span>1년 미만</span><strong>${number(s.underOneYearCount || 0)}</strong></div>
    <div><span>총 잔여 연차</span><strong>${number(s.totalRemaining || 0)}</strong></div>
    <div><span>당월 승인 사용</span><strong>${number(s.approvedCurrentMonth || 0)}</strong></div>
    <div><span>촉진 대상</span><strong>${number(s.reminderCount || 0)}</strong></div>`;
  const employees = data.employees || [];
  employeeBody.innerHTML = employees.length ? employees.map((row) => `<tr>
    <td>${row.underOneYear ? "1년 미만·월차" : "1년 이상·연차"}</td><td>${escapeHtml(row.regionalManager || "")}</td><td>${escapeHtml(row.manager || "")}</td>
    <td>${escapeHtml(row.storeName || "")}</td><td>${escapeHtml(row.employeeId || "")}</td><td>${escapeHtml(row.employeeName || "")}</td>
    <td>${escapeHtml(row.basisHireDate || row.hireDate || "")}</td><td>${escapeHtml([row.cycleStart, row.cycleEnd].filter(Boolean).join(" ~ "))}</td>
    <td>${number(row.granted || 0)}</td><td>${number(row.approvedUsed || 0)}</td><td>${number(row.approvedCurrentMonth || 0)}</td><td><strong>${number(row.remaining || 0)}</strong></td>
    <td>${escapeHtml(row.firstPromotionDate || "-")}</td><td>${escapeHtml(row.secondPromotionDate || "-")}</td><td>${escapeHtml(row.note || "")}</td>
  </tr>`).join("") : '<tr><td colspan="15" class="empty-cell">등록된 직원이 없습니다.</td></tr>';
  const reminders = data.reminders || [];
  reminderBody.innerHTML = reminders.length ? reminders.map((row) => `<tr>
    <td><span class="${row.status === "예정" ? "warning-pill" : "status-pill"}">${escapeHtml(row.status)}</span></td><td>${escapeHtml(row.promotionType)}</td><td>${escapeHtml(row.dueDate)}</td>
    <td>${escapeHtml(row.regionalManager || "")}</td><td>${escapeHtml(row.manager || "")}</td><td>${escapeHtml(row.storeName || "")}</td><td>${escapeHtml(row.employeeId)}</td><td>${escapeHtml(row.employeeName)}</td>
    <td>${number(row.remaining || 0)}</td><td>${escapeHtml([row.cycleStart, row.cycleEnd].filter(Boolean).join(" ~ "))}</td>
  </tr>`).join("") : '<tr><td colspan="10" class="empty-cell">표시할 촉진 대상이 없습니다.</td></tr>';
  const uploads = data.monthlyUploads || [];
  historyBody.innerHTML = uploads.length ? uploads.map((row) => `<tr>
    <td>${escapeHtml(row.month)}</td><td>${escapeHtml(row.file_name || "")}</td><td>${number(row.row_count || 0)}</td><td>${number(row.approved_days || 0)}</td>
    <td>${number(row.rejected_days || 0)}</td><td>${number(row.pending_days || 0)}</td><td>${escapeHtml(row.updated_at || row.created_at || "")}</td>
    <td><button class="text-button annual-month-delete" data-route="${escapeHtml(row.route)}" data-month="${escapeHtml(row.month)}" type="button">삭제</button></td>
  </tr>`).join("") : '<tr><td colspan="8" class="empty-cell">등록 이력이 없습니다.</td></tr>';
  $$(".annual-month-delete").forEach((button) => button.addEventListener("click", () => deleteAnnualMonthly(button.dataset.route, button.dataset.month)));
}


function exportAnnualLeaveDashboard() {
  const data = state.annualLeaveDashboard;
  const employees = data?.employees || [];
  if (!employees.length) return showToast("엑셀로 저장할 연차 누적현황이 없습니다.");
  if (employees.some((row) => isPhoneNumberText(row.employeeName))) {
    return showToast("기존 자료에 연락처가 이름으로 저장되어 있습니다. 수정 버전 배포 후 최초 연차대장을 다시 등록해 주세요.");
  }
  const routeLabel = ROUTE_LABELS[data.route] || "경로";
  const asOf = data.asOf || data.baseline?.baseline_date || toISODate(new Date());
  const wb = XLSX.utils.book_new();
  const border = { top:{style:"thin",color:{rgb:"D7DEE8"}}, bottom:{style:"thin",color:{rgb:"D7DEE8"}}, left:{style:"thin",color:{rgb:"D7DEE8"}}, right:{style:"thin",color:{rgb:"D7DEE8"}} };
  const headerStyle = { fill:{fgColor:{rgb:"1F4E78"}}, font:{name:"맑은 고딕",bold:true,color:{rgb:"FFFFFF"}}, alignment:{horizontal:"center",vertical:"center",wrapText:true}, border };
  const bodyStyle = { font:{name:"맑은 고딕",sz:10}, alignment:{vertical:"center"}, border };
  const centerStyle = { ...bodyStyle, alignment:{horizontal:"center",vertical:"center"} };
  const numberStyle = { ...bodyStyle, alignment:{horizontal:"right",vertical:"center"}, numFmt:"0.0" };
  const styleRange = (ws, range, style) => { const r=XLSX.utils.decode_range(range); for(let y=r.s.r;y<=r.e.r;y+=1) for(let x=r.s.c;x<=r.e.c;x+=1){ const a=XLSX.utils.encode_cell({r:y,c:x}); if(ws[a]) ws[a].s=style; } };

  const region = new Map();
  for (const row of employees) {
    const key=row.regionalManager||"미지정";
    if(!region.has(key)) region.set(key,{count:0,under:0,granted:0,used:0,remaining:0});
    const item=region.get(key); item.count+=1; item.under+=row.underOneYear?1:0; item.granted+=Number(row.granted||0); item.used+=Number(row.approvedUsed||0); item.remaining+=Number(row.remaining||0);
  }
  const summaryRows=[
    [`${routeLabel} 연차 누적 현황 보고`,"","","","",""],
    [`기준일: ${asOf}`,"","","","",""],
    ["관리 인원",employees.length,"1년 이상",employees.filter(r=>!r.underOneYear).length,"1년 미만",employees.filter(r=>r.underOneYear).length],
    ["총 발생",employees.reduce((s,r)=>s+Number(r.granted||0),0),"누적 사용",employees.reduce((s,r)=>s+Number(r.approvedUsed||0),0),"총 잔여",employees.reduce((s,r)=>s+Number(r.remaining||0),0)],
    [],["지역장","관리 인원","1년 미만","발생","누적 사용","잔여"],
    ...[...region.entries()].sort((a,b)=>a[0].localeCompare(b[0],"ko")).map(([name,v])=>[name,v.count,v.under,v.granted,v.used,v.remaining]),
  ];
  const wsSummary=XLSX.utils.aoa_to_sheet(summaryRows);
  wsSummary["!merges"]=[XLSX.utils.decode_range("A1:F1"),XLSX.utils.decode_range("A2:F2")];
  wsSummary["!cols"]=[{wch:18},{wch:13},{wch:13},{wch:13},{wch:13},{wch:13}];
  wsSummary.A1.s={fill:{fgColor:{rgb:"17365D"}},font:{name:"맑은 고딕",sz:18,bold:true,color:{rgb:"FFFFFF"}},alignment:{vertical:"center"}};
  styleRange(wsSummary,"A3:F4",{fill:{fgColor:{rgb:"D9EAF7"}},font:{name:"맑은 고딕",bold:true,color:{rgb:"17365D"}},alignment:{horizontal:"center",vertical:"center"},border});
  styleRange(wsSummary,`A6:F${summaryRows.length}`,bodyStyle); styleRange(wsSummary,"A6:F6",headerStyle); styleRange(wsSummary,`B7:F${summaryRows.length}`,numberStyle);
  XLSX.utils.book_append_sheet(wb,wsSummary,"전체 요약");

  const headers=["구분","지역장","매니저","지역","점포","사번","이름","연차 기준 입사일","사용기간 시작","사용기간 종료","발생","누적 사용","당월 승인","잔여","1차 촉진","2차 촉진","비고"];
  const rows=employees.map(row=>[row.underOneYear?"1년 미만·월차":"1년 이상·연차",row.regionalManager||"",row.manager||"",row.region2||row.region1||"",row.storeName||"",row.employeeId||"",row.employeeName||"",row.basisHireDate||row.hireDate||"",row.cycleStart||"",row.cycleEnd||"",Number(row.granted||0),Number(row.approvedUsed||0),Number(row.approvedCurrentMonth||0),Number(row.remaining||0),row.firstPromotionDate||"",row.secondPromotionDate||"",row.note||""]);
  const ws=XLSX.utils.aoa_to_sheet([headers,...rows]);
  ws["!cols"]=[{wch:15},{wch:11},{wch:11},{wch:9},{wch:22},{wch:13},{wch:10},{wch:15},{wch:13},{wch:13},{wch:9},{wch:10},{wch:10},{wch:9},{wch:13},{wch:13},{wch:30}];
  ws["!autofilter"]={ref:`A1:Q${rows.length+1}`};
  styleRange(ws,`A1:Q${rows.length+1}`,bodyStyle); styleRange(ws,"A1:Q1",headerStyle); styleRange(ws,`A2:J${rows.length+1}`,centerStyle); styleRange(ws,`K2:N${rows.length+1}`,numberStyle);
  for(let r=2;r<=rows.length+1;r+=1){ if(ws[`N${r}`]&&Number(ws[`N${r}`].v)<=3) ws[`N${r}`].s={...numberStyle,fill:{fgColor:{rgb:"FCE4D6"}},font:{name:"맑은 고딕",bold:true,color:{rgb:"C00000"}}}; }
  XLSX.utils.book_append_sheet(wb,ws,"연차 누적 현황");

  const reminders=(data.reminders||[]).map(r=>[r.status||"",r.promotionType||"",r.dueDate||"",r.regionalManager||"",r.manager||"",r.storeName||"",r.employeeId||"",r.employeeName||"",Number(r.remaining||0),[r.cycleStart,r.cycleEnd].filter(Boolean).join(" ~ ")]);
  const wsR=XLSX.utils.aoa_to_sheet([["상태","촉진 구분","촉진일","지역장","매니저","점포","사번","이름","잔여","사용기간"],...reminders]);
  wsR["!cols"]=[{wch:12},{wch:13},{wch:13},{wch:11},{wch:11},{wch:22},{wch:13},{wch:10},{wch:9},{wch:27}];
  styleRange(wsR,`A1:J${Math.max(1,reminders.length+1)}`,bodyStyle); styleRange(wsR,"A1:J1",headerStyle); styleRange(wsR,`A2:J${reminders.length+1}`,centerStyle);
  XLSX.utils.book_append_sheet(wb,wsR,"촉진 대상");
  XLSX.writeFile(wb,`${routeLabel}_연차누적현황_${asOf}.xlsx`,{bookType:"xlsx",cellStyles:true,compression:true});
  showToast("보고용 연차 누적현황 엑셀을 저장했습니다.");
}

async function deleteAnnualMonthly(route, month) {
  if (!confirm(`${month} 연차 승인·반려 자료를 삭제하시겠습니까?`)) return;
  try {
    const response = await fetch(`/api/annual-leave?kind=monthly&route=${encodeURIComponent(route)}&month=${encodeURIComponent(month)}`, { method: "DELETE" });
    const data = await readJsonResponse(response);
    if (!response.ok) throw new Error(data.error || "삭제 실패");
    await loadAnnualLeaveDashboard();
    showToast(`${month} 연차 승인·반려 자료를 삭제했습니다.`);
  } catch (error) {
    showToast(error.message || "연차 자료 삭제 실패");
  }
}

function parseInitialAnnualLedger(sheets, route, baselineDate) {
  if (!baselineDate) throw new Error("수기 연차대장 기준일을 입력해 주세요.");
  const normalized = normalizeWorkbookInput(sheets);
  const year = baselineDate.slice(0, 4);
  let employees = [];
  let formatLabel = "";
  let sourceSheets = [];
  let valueRule = "";

  const homeplusLegacy = route === "homeplus" ? findHomeplusLegacyAnnualSheet(normalized, year) : null;
  const standardSheets = normalized.filter((sheet) => looksLikeStandardAnnualLedgerSheet(sheet.matrix));
  const aboveSheet = normalized.find((sheet) => /1년\s*이상|재직\s*1년이상/.test(normalizeHeader(sheet.sheetName)));
  const underSheet = normalized.find((sheet) => /1년\s*미만|재직\s*1년미만/.test(normalizeHeader(sheet.sheetName)));

  if (homeplusLegacy) {
    employees = parseHomeplusLegacyAnnualLedger(homeplusLegacy.matrix, baselineDate);
    formatLabel = "홈플러스 기존 단일 연차대장";
    sourceSheets = [homeplusLegacy.sheetName];
    valueRule = `${year}년 사용가능·사용·잔여·적용기간의 저장된 셀 값을 반영합니다.`;
  } else if (standardSheets.length) {
    const preferred = standardSheets.find((sheet) => /연차\s*누적\s*현황/.test(normalizeHeader(sheet.sheetName))) || standardSheets[0];
    employees = parseStandardAnnualLedgerSheet(preferred.matrix, route, baselineDate);
    formatLabel = "시스템 표준 연차 누적현황";
    sourceSheets = [preferred.sheetName];
    valueRule = "발생·누적 사용·잔여·사용기간의 저장된 셀 값을 반영합니다.";
  } else if (aboveSheet || underSheet) {
    if (aboveSheet) employees.push(...parseAnnualBaselineSheet(aboveSheet.matrix, route, baselineDate, false));
    if (underSheet) employees.push(...parseAnnualBaselineSheet(underSheet.matrix, route, baselineDate, true));
    formatLabel = route === "electroland" ? "전자랜드 기존 1년 이상·미만 분리대장" : "1년 이상·미만 분리 연차대장";
    sourceSheets = [aboveSheet?.sheetName, underSheet?.sheetName].filter(Boolean);
    valueRule = "각 시트의 발생·사용·잔여·사용기간과 촉진일의 저장된 셀 값을 반영합니다.";
  } else {
    throw new Error("선택한 경로의 기존 연차대장 구조를 찾지 못했습니다. 홈플러스 단일 대장, 전자랜드 1년 이상·미만 대장 또는 시스템 표준 원장을 넣어 주세요.");
  }

  const unique = new Map();
  let excludedCount = 0;
  for (const employee of employees) {
    if (!employee.employeeId || !employee.employeeName || /필터|합계/.test(employee.employeeName)) continue;
    if (shouldExcludeAnnualBaselineEmployee(employee, route, baselineDate)) {
      excludedCount += 1;
      continue;
    }
    unique.set(employee.employeeId, employee);
  }
  if (!unique.size) throw new Error("연차대장에서 저장할 상담사 정보를 찾지 못했습니다.");
  return {
    employees: [...unique.values()],
    meta: {
      formatLabel,
      sheetNames: sourceSheets,
      excludedCount,
      valueRule,
    },
  };
}

function findHomeplusLegacyAnnualSheet(sheets, year) {
  for (const sheet of sheets) {
    const headerIndex = findFlexibleHeaderRow(sheet.matrix, (headers) => {
      const normalized = headers.map(normalizeHeader);
      const hasId = findHeaderIndex(normalized, ["제니엘사번", "사번"]) >= 0;
      const hasName = findHeaderIndex(normalized, ["사원명", "성명", "이름"]) >= 0;
      const hasYearBlock = normalized.some((value) => value === `${year}발생` || value === `${year}사용가능` || value === `${year}잔여`);
      return hasId && hasName && hasYearBlock;
    });
    if (headerIndex >= 0) return sheet;
  }
  return null;
}

function looksLikeStandardAnnualLedgerSheet(matrix) {
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => {
    const normalized = headers.map(normalizeHeader);
    return findHeaderIndex(normalized, ["사번", "제니엘사번"]) >= 0
      && findHeaderIndex(normalized, ["이름", "사원명", "성명"]) >= 0
      && findHeaderIndex(normalized, ["발생"]) >= 0
      && findHeaderIndex(normalized, ["잔여"]) >= 0
      && findHeaderIndex(normalized, ["누적사용", "누적 사용"]) >= 0;
  });
  return headerIndex >= 0;
}

function parseHomeplusLegacyAnnualLedger(matrix, baselineDate) {
  const year = baselineDate.slice(0, 4);
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => {
    const normalized = headers.map(normalizeHeader);
    return findHeaderIndex(normalized, ["제니엘사번", "사번"]) >= 0
      && findHeaderIndex(normalized, ["사원명", "이름", "성명"]) >= 0
      && normalized.some((value) => value === `${year}발생` || value === `${year}사용가능` || value === `${year}잔여`);
  });
  if (headerIndex < 0) return [];
  const top = (matrix[headerIndex] || []).map(normalizeHeader);
  const sub = (matrix[headerIndex + 1] || []).map(normalizeHeader);
  const combined = top.map((value, index) => normalizeHeader(`${value}${sub[index] || ""}`));
  const exact = (names, fallback = -1) => {
    const wanted = names.map(normalizeHeader);
    for (const name of wanted) {
      const index = top.findIndex((value) => value === name);
      if (index >= 0) return index;
    }
    for (const name of wanted) {
      const index = combined.findIndex((value) => value === name || value.includes(name));
      if (index >= 0) return index;
    }
    return fallback;
  };
  const after = (start, matcher) => {
    let found = -1;
    for (let index = Math.max(0, start); index < top.length; index += 1) if (matcher(top[index], combined[index], index)) found = index;
    return found;
  };
  const cols = {
    status: exact(["구분1"]), tenure: exact(["근속기간"]), category: exact(["구분2"]),
    region: exact(["지역"]), manager: exact(["매니저"]), portalId: exact(["포탈사번", "스핀사번"]),
    employeeId: exact(["제니엘사번", "사번"]), name: exact(["사원명", "이름", "성명"]),
    hireDate: exact(["제니엘입사일", "입사일"]), basisHireDate: exact(["연차기준일", "고용승계입사일"]),
    termination: exact(["퇴사예정일", "퇴사일"]), note: exact(["비고"]),
    grant: exact([`${year}발생`]), available: exact([`${year}사용가능`]),
    used: exact([`${year}사용입사월이후`, `${year}년사용입사월이후`, `${year}사용`]),
    grantDate: exact(["연차발생예정일", "연차발생일"]), remaining: exact([`${year}잔여`]),
  };
  const yearStart = Math.max(cols.grant, cols.available, 0);
  cols.period = after(yearStart, (topValue) => topValue === "적용기간");
  if (cols.note < 0) cols.note = after(Math.max(0, cols.remaining), (topValue) => topValue === "비고");

  const rows = [];
  for (const source of matrix.slice(headerIndex + 1)) {
    const employeeId = normalizeEmployeeId(source[cols.employeeId]);
    const employeeName = text(source[cols.name]);
    if (!looksLikeEmployeeId(employeeId) || !isPlausibleEmployeeName(employeeName)) continue;
    const category = text(source[cols.category]);
    const tenure = text(source[cols.tenure]);
    const underOneYear = /1년\s*미만/.test(`${category} ${tenure}`);
    const grantedCell = readNumericCell(source[cols.available]);
    const fallbackGrantCell = readNumericCell(source[cols.grant]);
    const usedCell = readNumericCell(source[cols.used]);
    const remainingCell = readNumericCell(source[cols.remaining]);
    const baselineGranted = grantedCell.present ? grantedCell.value : fallbackGrantCell.value;
    let baselineRemaining = remainingCell.present ? remainingCell.value : roundHalf(baselineGranted - usedCell.value);
    let baselineUsed = usedCell.present ? usedCell.value : roundHalf(baselineGranted - baselineRemaining);
    if (grantedCell.present && remainingCell.present) baselineUsed = roundHalf(baselineGranted - baselineRemaining);
    const period = parseDateRange(source[cols.period]);
    const grantDate = parseDateCell(source[cols.grantDate]);
    const hireDate = parseDateCell(source[cols.hireDate]);
    const basisHireDate = parseDateCell(source[cols.basisHireDate]) || hireDate;
    const sourceStatus = text(source[cols.status]);
    const note = [sourceStatus, text(source[cols.note])].filter(Boolean).join(" · ");
    rows.push({
      employeeId, employeeName, regionalManager: "", manager: text(source[cols.manager]),
      region1: text(source[cols.region]), region2: text(source[cols.region]), storeCode: "", storeName: "",
      portalId: text(source[cols.portalId]), hireDate, basisHireDate, policyType: "jan1", underOneYear,
      baselineDate, baselineGranted, baselineUsed, baselineRemaining,
      cycleStart: period.start || grantDate || `${year}-01-01`, cycleEnd: period.end || `${year}-12-31`,
      firstPromotionDate: "", secondPromotionDate: "", expiryDate: period.end || `${year}-12-31`,
      terminationDate: parseDateCell(source[cols.termination]), note, sourceStatus,
    });
  }
  return rows;
}

function parseStandardAnnualLedgerSheet(matrix, route, baselineDate) {
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => {
    const normalized = headers.map(normalizeHeader);
    return findHeaderIndex(normalized, ["사번", "제니엘사번"]) >= 0
      && findHeaderIndex(normalized, ["이름", "사원명", "성명"]) >= 0
      && findHeaderIndex(normalized, ["발생"]) >= 0
      && findHeaderIndex(normalized, ["잔여"]) >= 0;
  });
  if (headerIndex < 0) return [];
  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const find = (names) => {
    const exactIndex = findExactHeaderIndex(headers, names);
    return exactIndex >= 0 ? exactIndex : findHeaderIndex(headers, names);
  };
  const cols = {
    type: find(["구분"]), regionalManager: find(["지역장"]), manager: find(["매니저"]), region: find(["지역"]),
    storeCode: find(["점포코드", "근무처코드"]), storeName: find(["점포", "매장명", "근무처명"]),
    portalId: find(["스핀사번", "포탈사번"]), employeeId: find(["사번", "제니엘사번"]), name: find(["이름", "사원명", "성명"]),
    hireDate: find(["입사일", "제니엘입사일"]), basisHireDate: find(["연차기준입사일", "연차기준일", "고용승계입사일"]),
    cycleStart: find(["사용기간시작"]), cycleEnd: find(["사용기간종료"]), period: find(["사용기간"]),
    granted: find(["발생"]), used: find(["누적사용", "사용"]), remaining: find(["잔여"]),
    firstPromotion: find(["1차촉진", "1차촉진일"]), secondPromotion: find(["2차촉진", "2차촉진일"]),
    expiry: find(["소진일", "연차소진일"]), termination: find(["퇴사일", "퇴사예정일"]), note: find(["비고"]),
  };
  const rows = [];
  for (const source of matrix.slice(headerIndex + 1)) {
    const employeeId = normalizeEmployeeId(source[cols.employeeId]);
    const employeeName = text(source[cols.name]);
    if (!looksLikeEmployeeId(employeeId) || !isPlausibleEmployeeName(employeeName)) continue;
    const type = text(source[cols.type]);
    const period = parseDateRange(source[cols.period]);
    const cycleStart = parseDateCell(source[cols.cycleStart]) || period.start;
    const cycleEnd = parseDateCell(source[cols.cycleEnd]) || period.end || parseDateCell(source[cols.expiry]);
    const baselineGranted = numericCell(source[cols.granted]);
    const remainingCell = readNumericCell(source[cols.remaining]);
    let baselineUsed = numericCell(source[cols.used]);
    const baselineRemaining = remainingCell.present ? remainingCell.value : roundHalf(baselineGranted - baselineUsed);
    if (remainingCell.present && baselineGranted) baselineUsed = roundHalf(baselineGranted - baselineRemaining);
    rows.push({
      employeeId, employeeName, regionalManager: text(source[cols.regionalManager]), manager: text(source[cols.manager]),
      region1: text(source[cols.region]), region2: text(source[cols.region]), storeCode: text(source[cols.storeCode]), storeName: text(source[cols.storeName]),
      portalId: text(source[cols.portalId]), hireDate: parseDateCell(source[cols.hireDate]),
      basisHireDate: parseDateCell(source[cols.basisHireDate]) || parseDateCell(source[cols.hireDate]),
      policyType: route === "homeplus" ? "jan1" : "anniversary", underOneYear: /1년\s*미만|월차/.test(type),
      baselineDate, baselineGranted, baselineUsed, baselineRemaining, cycleStart, cycleEnd,
      firstPromotionDate: parseDateCell(source[cols.firstPromotion]), secondPromotionDate: parseDateCell(source[cols.secondPromotion]),
      expiryDate: parseDateCell(source[cols.expiry]) || cycleEnd, terminationDate: parseDateCell(source[cols.termination]),
      note: text(source[cols.note]), sourceStatus: type,
    });
  }
  return rows;
}

function readNumericCell(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return { present: false, value: 0 };
  const parsed = Number(raw.replace(/,/g, "").replace(/[^0-9.\-]/g, ""));
  return { present: Number.isFinite(parsed), value: Number.isFinite(parsed) ? roundHalf(parsed) : 0 };
}

function shouldExcludeAnnualBaselineEmployee(employee, route, baselineDate) {
  const terminationDate = employee.terminationDate || "";
  if (terminationDate && terminationDate <= baselineDate) return true;
  const status = normalizeHeader(employee.sourceStatus || "");
  if (/퇴사/.test(status) && (!terminationDate || terminationDate <= baselineDate)) return true;
  if (route === "homeplus" && /(홈플러스|홈플).*(전자랜드|랜드|코스트코)/.test(status)) return true;
  if (route === "electroland" && /(전자랜드|랜드).*(홈플러스|홈플|코스트코)/.test(status)) return true;
  return false;
}

function parseAnnualBaselineSheet(matrix, route, baselineDate, underOneYear) {
  const headerIndex = findFlexibleHeaderRow(matrix, (headers) => findHeaderIndex(headers, ["제니엘사번", "사번", "사원번호"]) >= 0 && findHeaderIndex(headers, ["이름", "성명", "사원명", "직원명"]) >= 0);
  if (headerIndex < 0) return [];
  const headerRows = matrix.slice(Math.max(0, headerIndex - 5), headerIndex + 1);
  const headers = (matrix[headerIndex] || []).map(normalizeHeader);
  const find = (names, fallback = -1) => {
    const exact = findExactHeaderIndex(headers, names);
    if (exact >= 0) return exact;
    const direct = findHeaderIndex(headers, names);
    if (direct >= 0) return direct;
    for (let r = headerRows.length - 1; r >= 0; r -= 1) {
      const rowHeaders = (headerRows[r] || []).map(normalizeHeader);
      const exactInRow = findExactHeaderIndex(rowHeaders, names);
      if (exactInRow >= 0) return exactInRow;
      const idx = findHeaderIndex(rowHeaders, names);
      if (idx >= 0) return idx;
    }
    return fallback;
  };
  const cols = underOneYear ? {
    regionalManager: find(["지역장"], 3), manager: find(["매니저"], 4), region1: find(["지역1"], 5), region2: find(["지역2"], 6),
    storeCode: find(["근무처코드", "점포코드"], 7), storeName: find(["근무처명", "점포", "매장명"], 8), portalId: find(["스핀사번", "포탈사번"], 10),
    employeeId: find(["제니엘사번", "사번"], 11), name: find(["사원명", "성명", "이름", "직원명"], 12), hireDate: find(["입사일", "제니엘입사일"], 15), basisHireDate: find(["고용승계입사일"], -1),
    granted: find(["발생", "월차발생"], 16), used: find([`${baselineDate.slice(0,4)}사용`, `${baselineDate.slice(0,4)}년사용`, "사용"], 18), remaining: find(["잔여"], 19), period: find(["사용기간"], 20),
    termination: find(["퇴사일"], 21), firstPromotion: find(["1차촉진", "1차촉진일"], 23), secondPromotion: find(["2차촉진", "2차촉진일"], 25), expiry: find(["연차소진일", "소진일"], 27), note: find(["비고"], 28),
  } : findAnnualAboveColumns(matrix, headerIndex, baselineDate);
  const rows = [];
  for (const source of matrix.slice(headerIndex + 1)) {
    const employeeId = normalizeEmployeeId(source[cols.employeeId]);
    const employeeName = resolveAnnualEmployeeName(source, cols);
    if (!looksLikeEmployeeId(employeeId) || !employeeName) continue;
    const activeCycle = underOneYear ? null : selectActiveAnnualCycle(source, cols.cycles || [], baselineDate);
    const period = activeCycle?.period || parseDateRange(source[cols.period]);
    const grantDate = activeCycle?.grantDate || parseDateCell(source[cols.grantDate]);
    const hireDate = parseDateCell(source[cols.hireDate]);
    const basisHireDate = parseDateCell(source[cols.basisHireDate]) || hireDate;
    let baselineRemaining = activeCycle ? activeCycle.remaining : numericCell(source[cols.remaining]);
    let baselineGranted = activeCycle ? activeCycle.granted : numericCell(source[cols.granted]);
    let baselineUsed = activeCycle ? activeCycle.used : numericCell(source[cols.used]);
    if (!baselineGranted && baselineUsed + baselineRemaining > 0) baselineGranted = roundHalf(baselineUsed + baselineRemaining);
    if (baselineGranted || baselineRemaining) baselineUsed = roundHalf(baselineGranted - baselineRemaining);
    const cycleStart = period.start || grantDate || (underOneYear ? hireDate : "");
    const cycleEnd = period.end || parseDateCell(source[cols.expiry]) || "";
    rows.push({
      employeeId, employeeName, regionalManager: text(source[cols.regionalManager]), manager: text(source[cols.manager]),
      region1: text(source[cols.region1]), region2: text(source[cols.region2]), storeCode: text(source[cols.storeCode]), storeName: text(source[cols.storeName]),
      portalId: text(source[cols.portalId]), hireDate, basisHireDate, policyType: route === "homeplus" ? "jan1" : "anniversary", underOneYear,
      baselineDate, baselineGranted, baselineUsed, baselineRemaining, cycleStart, cycleEnd,
      firstPromotionDate: parseDateCell(source[cols.firstPromotion]), secondPromotionDate: parseDateCell(source[cols.secondPromotion]),
      expiryDate: parseDateCell(source[cols.expiry]) || cycleEnd, terminationDate: parseDateCell(source[cols.termination]), note: text(source[cols.note]),
    });
  }
  return rows;
}


function findExactHeaderIndex(headers, candidates) {
  for (const candidate of candidates.map(normalizeHeader)) {
    const index = headers.findIndex((header) => header === candidate);
    if (index >= 0) return index;
  }
  return -1;
}

function isPhoneNumberText(value) {
  const normalized = text(value).replace(/\s+/g, "");
  return /^0\d{1,2}-?\d{3,4}-?\d{4}$/.test(normalized) || /^\d{9,11}$/.test(normalized.replace(/-/g, ""));
}

function isPlausibleEmployeeName(value) {
  const normalized = text(value);
  if (!normalized || normalized.length > 50) return false;
  if (/필터|합계|연락처|이메일/.test(normalized)) return false;
  if (isPhoneNumberText(normalized) || normalized.includes("@")) return false;
  // 동명이인 구분을 위해 사원명 뒤에 1·2가 붙는 경우(예: 김현수2)는 정상 이름으로 인정합니다.
  if (/^[A-Za-z]{1,5}\d{3,}$/.test(normalized) || /^\d+(?:\.\d+)?$/.test(normalized)) return false;
  return /[가-힣A-Za-z]/.test(normalized);
}

function resolveAnnualEmployeeName(source, cols) {
  const primary = text(source[cols.name]);
  if (isPlausibleEmployeeName(primary)) return primary;
  // 전자랜드 기준대장은 제니엘 사번 바로 오른쪽 열이 사원명입니다.
  for (const index of [...new Set([Number(cols.employeeId) + 1, Number(cols.name) - 1, Number(cols.name) + 1])]) {
    if (!Number.isInteger(index) || index < 0) continue;
    const candidate = text(source[index]);
    if (isPlausibleEmployeeName(candidate)) return candidate;
  }
  return "";
}

function findAnnualAboveColumns(matrix, headerIndex, baselineDate) {
  const yearText = baselineDate.slice(0, 4);
  const searchRows = matrix.slice(0, headerIndex + 1);
  const header = (matrix[headerIndex] || []).map(normalizeHeader);
  const maxCol = Math.max(...searchRows.map((row) => row.length), 0);
  const flattened = [];
  for (let col = 0; col < maxCol; col += 1) flattened[col] = normalizeHeader(searchRows.map((row) => row[col]).filter(Boolean).join(" "));
  const findAll = (patterns) => flattened.map((value, index) => patterns.some((p) => value.includes(normalizeHeader(p))) ? index : -1).filter((index) => index >= 0);
  const direct = (patterns, fallback) => {
    const normalizedPatterns = patterns.map(normalizeHeader);
    for (const pattern of normalizedPatterns) {
      const exact = header.findIndex((value) => value === pattern);
      if (exact >= 0) return exact;
    }
    for (const pattern of normalizedPatterns) {
      const close = header.findIndex((value) => value && pattern.length >= 2 && (value.startsWith(pattern) || value.endsWith(pattern)));
      if (close >= 0) return close;
    }
    const candidates = findAll(patterns);
    return candidates.length ? candidates[0] : fallback;
  };
  const termination = direct(["퇴사일"], 53);
  const grantColumns = [];
  for (let col = 0; col < Math.min(header.length, termination); col += 1) {
    const value = header[col];
    const match = value.match(/(20\d{2})년?.*발생/);
    if (match) grantColumns.push({ year: Number(match[1]), grantCol: col });
  }
  const cycles = grantColumns.map((item, index) => {
    const endCol = index + 1 < grantColumns.length ? grantColumns[index + 1].grantCol : termination;
    const columns = [];
    for (let col = item.grantCol; col < endCol; col += 1) columns.push(col);
    const grantDateCol = columns.find((col) => header[col].includes("연차발생일")) ?? -1;
    const remainingCol = [...columns].reverse().find((col) => header[col].includes("잔여")) ?? -1;
    const periodCol = [...columns].reverse().find((col) => header[col].includes("사용기간")) ?? -1;
    const usedCols = columns.filter((col) => header[col].includes("사용") && !header[col].includes("사용기간"));
    return { ...item, grantDateCol, remainingCol, periodCol, usedCols };
  }).filter((cycle) => cycle.periodCol >= 0 || cycle.grantDateCol >= 0);
  const preferred = cycles.find((cycle) => cycle.year === Number(yearText)) || cycles[cycles.length - 1] || {};
  return {
    regionalManager: direct(["지역장"], 3), manager: direct(["매니저"], 4), region1: direct(["지역1"], 5), region2: direct(["지역2"], 6),
    storeCode: direct(["근무처코드", "점포코드"], 7), storeName: direct(["근무처명", "매장명"], 8), portalId: direct(["스핀사번", "포탈사번"], 10),
    employeeId: direct(["제니엘사번", "사번"], 11), name: direct(["사원명", "성명", "이름"], 12), hireDate: direct(["제니엘입사일", "입사일"], 15), basisHireDate: direct(["고용승계입사일"], 16),
    granted: preferred.grantCol ?? 45, grantDate: preferred.grantDateCol ?? 46,
    used: preferred.usedCols?.[0] ?? 47, remaining: preferred.remainingCol ?? 48, period: preferred.periodCol ?? 49,
    termination, firstPromotion: -1, secondPromotion: -1, expiry: -1, note: direct(["비고"], 75), cycles,
  };
}

function selectActiveAnnualCycle(source, cycles, baselineDate) {
  const candidates = (cycles || []).map((cycle) => {
    const period = parseDateRange(source[cycle.periodCol]);
    const grantDate = parseDateCell(source[cycle.grantDateCol]);
    const granted = numericCell(source[cycle.grantCol]);
    const used = roundHalf((cycle.usedCols || []).reduce((sum, col) => sum + numericCell(source[col]), 0));
    const remaining = numericCell(source[cycle.remainingCol]);
    return { ...cycle, period, grantDate, granted, used, remaining };
  }).filter((cycle) => cycle.period.start || cycle.grantDate || cycle.granted || cycle.used || cycle.remaining);
  const containing = candidates.filter((cycle) => cycle.period.start && cycle.period.end && cycle.period.start <= baselineDate && baselineDate <= cycle.period.end);
  if (containing.length) return containing.sort((a, b) => b.period.start.localeCompare(a.period.start))[0];
  const started = candidates.filter((cycle) => (cycle.period.start || cycle.grantDate || "") <= baselineDate)
    .sort((a, b) => String(b.period.start || b.grantDate || "").localeCompare(String(a.period.start || a.grantDate || "")));
  return started[0] || candidates[candidates.length - 1] || null;
}

function parseDateRange(value) {
  const raw = text(value);
  const parts = raw.split(/\s*(?:~|～|부터|에서)\s*/).map((item) => parseDateCell(item)).filter(Boolean);
  if (parts.length >= 2) return { start: parts[0], end: parts[1] };
  const matches = raw.match(/20\d{2}[.\/-]\d{1,2}[.\/-]\d{1,2}/g) || [];
  return { start: parseDateCell(matches[0]) || "", end: parseDateCell(matches[1]) || "" };
}

function numericCell(value) {
  const parsed = Number(String(value ?? "").replace(/,/g, "").replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(parsed) ? roundHalf(parsed) : 0;
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
    await Promise.all([loadHistory(), loadGrants(), loadArchiveFiles(), loadWorkforceUploads(), loadAnnualLeaveDashboard()]);
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
  state.archiveFiles = [];
  state.workforceUploads = [];
  state.workforce = null;
  state.annualLeaveDashboard = null;
  renderArchiveFiles([]);
  renderWorkforceUploads([]);
  renderAnnualLeaveDashboard(null);
  showToast("로그아웃되었습니다.");
}

function switchView(view) {
  $$(".tab[data-view]").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === view));
  $$(".view").forEach((section) => section.classList.remove("active"));
  $(`#${view}View`).classList.add("active");
  if (view === "history") loadHistory();
  if (view === "substitute") loadGrants();
  if (view === "files") loadArchiveFiles();
  if (view === "system") Promise.all([loadWorkforceUploads(), loadPersonnelChecks()]);
  if (view === "annualLeave") loadAnnualLeaveDashboard();
}

function resetAll() {
  state.planFile = null;
  state.attendanceFile = null;
  state.annualFile = null;
  state.referenceFile = null;
  state.closureBaseFile = null;
  state.closureTargetFile = null;
  state.closureComparison = null;
  state.result = null;
  state.priorLedger = emptyLedger();
  $("#planFile").value = "";
  $("#attendanceFile").value = "";
  $("#annualFile").value = "";
  $("#referenceFile").value = "";
  setPlanFile(null);
  setAttendanceFile(null);
  setAnnualFile(null);
  setReferenceFile(null);
  setDefaultDates();
  syncRouteRuleHelp();
  $("#resultArea").classList.add("hidden");
  $("#emptyState").classList.remove("hidden");
  updateWorkforceStatus();
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
    const normalizedCandidate = normalizeHeader(candidate);
    let index = headers.indexOf(normalizedCandidate);
    if (index >= 0) return index;
    index = headers.findIndex((header) => header && normalizedCandidate.length >= 2 && (
      header.startsWith(normalizedCandidate) || header.endsWith(normalizedCandidate)
    ));
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
  return text(value)
    .toUpperCase()
    .replace(/\.0+$/, "")
    .replace(/[\s\u00A0-]+/g, "")
    .replace(/[^0-9A-Z가-힣]/g, "");
}

function looksLikeEmployeeId(value) {
  const normalized = normalizeEmployeeId(value);
  if (!normalized || normalized.length < 3 || normalized.length > 30) return false;
  if (["사번", "사원번호", "사원ID", "사번ID", "EMPLOYEEID", "ID"].includes(normalized)) return false;
  // 실제 운영 사번은 E248024처럼 영문+숫자 또는 숫자로 구성됩니다.
  // 이전 버전처럼 비어 있지 않은 사번은 폭넓게 인정하되 숫자가 하나 이상 있어야 합니다.
  return /\d/.test(normalized);
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

function cleanClockValue(value) {
  const raw = text(value);
  return !raw || raw === "546" || raw === "0" ? "" : raw;
}

function cleanPlaceholderValue(value) {
  const raw = text(value);
  return raw === "546" ? "" : raw;
}

function parseDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toISODate(value);
  if (typeof value === "number" && Number.isFinite(value) && window.XLSX?.SSF?.parse_date_code) {
    const parsedSerial = XLSX.SSF.parse_date_code(value);
    if (parsedSerial?.y && parsedSerial?.m && parsedSerial?.d) {
      return `${String(parsedSerial.y).padStart(4, "0")}-${String(parsedSerial.m).padStart(2, "0")}-${String(parsedSerial.d).padStart(2, "0")}`;
    }
  }

  const raw = text(value);
  if (!raw) return "";
  const digits = raw.replace(/\D/g, "");
  if (digits.length === 8) return validISODate(`${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`);
  if (digits.length === 6) {
    const year = Number(digits.slice(0, 2)) >= 70 ? `19${digits.slice(0, 2)}` : `20${digits.slice(0, 2)}`;
    return validISODate(`${year}-${digits.slice(2, 4)}-${digits.slice(4, 6)}`);
  }

  const match = raw.match(/(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/);
  if (match) return validISODate(`${match[1]}-${String(Number(match[2])).padStart(2, "0")}-${String(Number(match[3])).padStart(2, "0")}`);
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? "" : toISODate(parsed);
}

function validISODate(value) {
  const parsed = new Date(`${value}T00:00:00`);
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
  rows.sort((a, b) => Math.max(b.shortage || 0, b.compensationShortage || 0) - Math.max(a.shortage || 0, a.compensationShortage || 0)
    || (b.baseExcess || 0) - (a.baseExcess || 0)
    || a.store.localeCompare(b.store)
    || a.name.localeCompare(b.name));
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
