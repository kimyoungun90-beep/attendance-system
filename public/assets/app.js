const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const state = {
  planFile: null,
  attendanceFile: null,
  result: null,
  balances: {},
  backend: { available: false, configured: false, loggedIn: false },
  activeResultTab: "missing",
  pages: { missing: 1, unexpected: 1, dayoff: 1 },
  pageSize: 50,
  filtered: { missing: [], unexpected: [], dayoff: [] },
};

const companyLabels = { homeplus: "홈플러스", electroland: "전자랜드" };
const weekdayLabels = ["일", "월", "화", "수", "목", "금", "토"];
const FULL_LEAVE_KEYWORDS = ["연차", "공가", "월차", "경조", "병가", "휴직", "퇴사", "출산", "육아", "결근"];
const SUBSTITUTE_KEYWORDS = ["대체휴일", "대체휴무", "보상휴가", "보상휴일"];

init();

async function init() {
  setDefaultDates();
  bindEvents();
  setupDropzone("planDropzone", "planFile", setPlanFile);
  setupDropzone("attendanceDropzone", "attendanceFile", setAttendanceFile);
  await checkBackend();
  syncCompanyRuleHelp();
}

function bindEvents() {
  $("#analyzeButton").addEventListener("click", analyzeFiles);
  $("#resetButton").addEventListener("click", resetAll);
  $("#searchInput").addEventListener("input", applyFilters);
  $("#storeFilter").addEventListener("change", applyFilters);
  $("#exportButton").addEventListener("click", exportResults);
  $("#saveClosureButton").addEventListener("click", saveClosure);
  $("#refreshHistoryButton").addEventListener("click", loadHistory);
  $("#refreshGrantButton").addEventListener("click", loadGrants);
  $("#grantCompanyFilter").addEventListener("change", loadGrants);
  $("#grantForm").addEventListener("submit", saveGrant);
  $("#grantMonth").addEventListener("change", syncGrantDates);
  $("#loginButton").addEventListener("click", openLogin);
  $("#logoutButton").addEventListener("click", logout);
  $("#loginCancel").addEventListener("click", () => $("#loginDialog").close());
  $("#loginForm").addEventListener("submit", login);
  $("#targetMonth").addEventListener("change", () => { syncCutoffWithMonth(); syncCompanyRuleHelp(); });
  $$('input[name="company"]').forEach((input) => input.addEventListener("change", syncCompanyRuleHelp));
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
  const nextMonthEnd = new Date(year, monthNumber + 1, 0);
  $("#grantValidTo").value = toISODate(nextMonthEnd);
}

function syncCompanyRuleHelp() {
  const company = selectedCompany();
  const month = $("#targetMonth").value;
  if (company === "homeplus") {
    $("#companyRuleHelp").textContent = "홈플러스 기본 휴무는 월 6일입니다.";
  } else {
    const count = month ? countWeekendDays(month) : 0;
    $("#companyRuleHelp").textContent = `전자랜드 기본 휴무는 대상 월 토·일 합계 ${count}일입니다.`;
  }
}

function setupDropzone(zoneId, inputId, setter) {
  const zone = $(`#${zoneId}`);
  const input = $(`#${inputId}`);
  input.addEventListener("change", () => setter(input.files?.[0] || null));
  ["dragenter", "dragover"].forEach((name) => zone.addEventListener(name, (event) => {
    event.preventDefault(); zone.classList.add("dragover");
  }));
  ["dragleave", "drop"].forEach((name) => zone.addEventListener(name, (event) => {
    event.preventDefault(); zone.classList.remove("dragover");
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
    if (!window.XLSX) throw new Error("엑셀 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하세요.");
    if (!state.planFile || !state.attendanceFile) throw new Error("계획표와 실제 근태표를 모두 선택해 주세요.");
    const targetMonth = $("#targetMonth").value;
    const cutoffDate = $("#cutoffDate").value;
    if (!targetMonth || !cutoffDate) throw new Error("대상 월과 비교 기준일을 입력해 주세요.");
    if (!cutoffDate.startsWith(targetMonth)) throw new Error("비교 기준일은 대상 월 안의 날짜여야 합니다.");

    button.disabled = true;
    button.textContent = "파일 분석 중...";

    const [planMatrix, attendanceMatrix] = await Promise.all([
      fileToMatrix(state.planFile), fileToMatrix(state.attendanceFile),
    ]);
    const plan = parsePlan(planMatrix);
    const attendance = parseAttendance(attendanceMatrix, targetMonth);
    state.balances = await loadBalancesForAnalysis(selectedCompany(), targetMonth);

    const result = compareAttendance({
      plan,
      attendance,
      company: selectedCompany(),
      targetMonth,
      cutoffDate,
      checkMode: $("#checkMode").value,
      balances: state.balances,
    });

    state.result = {
      ...result,
      company: selectedCompany(),
      companyLabel: companyLabels[selectedCompany()],
      targetMonth,
      cutoffDate,
      checkMode: $("#checkMode").value,
      planFileName: state.planFile.name,
      attendanceFileName: state.attendanceFile.name,
      analyzedAt: new Date().toISOString(),
    };
    renderResult();
    showToast(`분석 완료: 미출근 ${result.missingRows.length}건 · 휴무 출근 ${result.unexpectedRows.length}건`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "분석 중 오류가 발생했습니다.");
  } finally {
    button.disabled = false;
    button.textContent = "근태 종합 분석";
  }
}

async function loadBalancesForAnalysis(company, month) {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) return {};
  try {
    const response = await fetch(`/api/substitute-balances?company=${encodeURIComponent(company)}&month=${encodeURIComponent(month)}`, { cache: "no-store" });
    if (!response.ok) throw new Error("대체휴무 잔여 조회 실패");
    return (await response.json()).balances || {};
  } catch (error) {
    showToast(`${error.message}. 대체휴무 가용 수량은 0으로 계산합니다.`);
    return {};
  }
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
  return { rows, headerIndex, dayColumns, detectedCompany: detectCompanyFromPlan(rows) };
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
    location: findHeaderIndex(headers, ["출근지점", "출근매장", "근무지점"]),
  };
  if (columns.date < 0 || columns.employeeId < 0) throw new Error("근태표의 날짜 또는 사번 열을 확인해 주세요.");
  if (columns.actualIn < 0 && columns.changedIn < 0) throw new Error("근태표에서 실제/변경 출근시간 열을 찾지 못했습니다.");

  const rows = matrix.slice(headerIndex + 1).map((row, sourceIndex) => {
    const date = parseDateCell(row[columns.date]);
    return {
      employeeId: normalizeEmployeeId(row[columns.employeeId]),
      name: columns.name >= 0 ? text(row[columns.name]) : "",
      date,
      actualIn: columns.actualIn >= 0 ? text(row[columns.actualIn]) : "",
      changedIn: columns.changedIn >= 0 ? text(row[columns.changedIn]) : "",
      location: columns.location >= 0 ? text(row[columns.location]) : "",
      sourceIndex: sourceIndex + headerIndex + 2,
    };
  }).filter((row) => row.employeeId && row.date && row.date.startsWith(targetMonth));

  if (!rows.length) throw new Error(`${targetMonth} 근태 기록을 찾지 못했습니다. 대상 월을 확인해 주세요.`);
  return { rows, headerIndex, detectedCompany: detectCompanyFromAttendance(rows) };
}

function compareAttendance({ plan, attendance, company, targetMonth, cutoffDate, checkMode, balances }) {
  const attendanceIds = new Set(attendance.rows.map((row) => row.employeeId));
  const planIds = new Set(plan.rows.map((row) => row.employeeId));
  const matchedIds = [...planIds].filter((id) => attendanceIds.has(id));
  const matchRate = planIds.size ? Math.round((matchedIds.length / planIds.size) * 1000) / 10 : 0;
  const attendanceById = groupBy(attendance.rows, (row) => row.employeeId);
  const selectedPlans = choosePlanRows(plan.rows, attendanceById);
  const clocked = buildClockMap(attendance.rows);
  const [year, month] = targetMonth.split("-").map(Number);
  const lastDay = Math.min(new Date(year, month, 0).getDate(), Number(cutoffDate.slice(-2)));
  const baseAllowance = company === "homeplus" ? 6 : countWeekendDays(targetMonth);
  const missingRows = [];
  const unexpectedRows = [];
  const employeeSummaries = [];

  selectedPlans.forEach((person) => {
    let basicDayoffUsed = 0;
    let explicitSubDayoffUsed = 0;

    const daysInMonth = new Date(year, month, 0).getDate();
    for (let day = 1; day <= daysInMonth; day += 1) {
      const monthlyPlanStatus = text(person.plans[day]);
      if (isBasicDayOff(monthlyPlanStatus)) basicDayoffUsed += 1;
      explicitSubDayoffUsed += substituteDayValue(monthlyPlanStatus);
    }

    for (let day = 1; day <= lastDay; day += 1) {
      const planStatus = text(person.plans[day]);
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const attendanceValue = clocked.get(`${person.employeeId}|${date}`) || { hasClockIn: false, actualIn: "", changedIn: "", location: "" };
      const dateObject = new Date(`${date}T00:00:00`);

      if (shouldRequireClockIn(planStatus, checkMode) && !attendanceValue.hasClockIn) {
        missingRows.push(makeIssueRow({
          type: "missing_clock_in", company, person, date, dateObject, planStatus: planStatus || "공백", attendanceValue,
          result: "근무인데 미출근", reason: planStatus ? `${planStatus} 계획이나 출근시간 없음` : "계획 공백이나 출근시간 없음",
        }));
      }

      if (isFullDayOff(planStatus) && attendanceValue.hasClockIn) {
        unexpectedRows.push(makeIssueRow({
          type: "unexpected_clock_in", company, person, date, dateObject, planStatus, attendanceValue,
          result: "휴무인데 출근", reason: `${planStatus} 계획이나 출근시간이 기록됨`,
        }));
      }
    }

    const availableSubstitute = roundHalf(Number(balances[person.employeeId] || 0));
    const baseExcess = roundHalf(Math.max(0, basicDayoffUsed - baseAllowance));
    const substituteNeeded = roundHalf(baseExcess + explicitSubDayoffUsed);
    const substituteApplied = roundHalf(Math.min(substituteNeeded, availableSubstitute));
    const shortage = roundHalf(Math.max(0, substituteNeeded - availableSubstitute));
    const remaining = roundHalf(Math.max(0, availableSubstitute - substituteApplied));
    const judgment = buildDayoffJudgment({ baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess, substituteNeeded, availableSubstitute, substituteApplied, shortage, remaining });

    employeeSummaries.push({
      company,
      companyLabel: companyLabels[company],
      store: person.store,
      employeeId: person.employeeId,
      name: person.name,
      baseAllowance,
      basicDayoffUsed: roundHalf(basicDayoffUsed),
      explicitSubDayoffUsed: roundHalf(explicitSubDayoffUsed),
      baseExcess,
      substituteNeeded,
      availableSubstitute,
      substituteApplied,
      remaining,
      shortage,
      judgment,
      issue: baseExcess > 0 || explicitSubDayoffUsed > 0 || shortage > 0,
      duplicatePlanNote: person.duplicatePlanNote || "",
    });
  });

  missingRows.sort(issueSort);
  unexpectedRows.sort(issueSort);
  employeeSummaries.sort((a, b) => b.shortage - a.shortage || b.baseExcess - a.baseExcess || a.store.localeCompare(b.store) || a.name.localeCompare(b.name));
  const diagnostics = buildDiagnostics({ company, plan, attendance, matchRate, planIds, attendanceIds, baseAllowance });
  return {
    missingRows,
    unexpectedRows,
    employeeSummaries,
    planPeople: planIds.size,
    attendancePeople: attendanceIds.size,
    matchedPeople: matchedIds.length,
    matchRate,
    missingPeople: new Set(missingRows.map((row) => row.employeeId)).size,
    unexpectedPeople: new Set(unexpectedRows.map((row) => row.employeeId)).size,
    dayoffExcessPeople: employeeSummaries.filter((row) => row.baseExcess > 0).length,
    substituteShortagePeople: employeeSummaries.filter((row) => row.shortage > 0).length,
    baseAllowance,
    diagnostics,
  };
}

function buildClockMap(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = `${row.employeeId}|${row.date}`;
    const existing = map.get(key) || { hasClockIn: false, actualIn: "", changedIn: "", location: "" };
    map.set(key, {
      hasClockIn: existing.hasClockIn || Boolean(row.actualIn || row.changedIn),
      actualIn: row.actualIn || existing.actualIn,
      changedIn: row.changedIn || existing.changedIn,
      location: row.location || existing.location,
    });
  });
  return map;
}

function makeIssueRow({ type, company, person, date, dateObject, planStatus, attendanceValue, result, reason }) {
  return {
    issueType: type,
    company,
    companyLabel: companyLabels[company],
    store: person.store,
    employeeId: person.employeeId,
    name: person.name,
    date,
    weekday: weekdayLabels[dateObject.getDay()],
    planStatus,
    actualIn: attendanceValue.actualIn || "",
    changedIn: attendanceValue.changedIn || "",
    clockStatus: attendanceValue.hasClockIn ? (attendanceValue.changedIn || attendanceValue.actualIn) : "미기록",
    result,
    reason,
    duplicatePlanNote: person.duplicatePlanNote || "",
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

function shouldRequireClockIn(planStatus, mode) {
  const status = text(planStatus);
  if (isFullDayOff(status)) return false;
  if (!status) return true;
  if (status.includes("0.5") || status.includes("반차") || status.includes("반반차")) return true;
  if (mode === "work-only") return status.includes("근무") || status.includes("출근");
  return status.includes("근무") || status.includes("출근") || status.includes("교육") || status.includes("출장");
}

function isBasicDayOff(value) {
  const status = normalizeStatus(value);
  return status === "휴무" || status === "월휴무" || status === "정기휴무";
}

function substituteDayValue(value) {
  const status = text(value);
  if (!SUBSTITUTE_KEYWORDS.some((keyword) => status.includes(keyword))) return 0;
  if (status.includes("0.5") || status.includes("반일")) return 0.5;
  return 1;
}

function isFullDayOff(value) {
  const status = text(value);
  if (!status) return false;
  if (isBasicDayOff(status)) return true;
  if (SUBSTITUTE_KEYWORDS.some((keyword) => status.includes(keyword))) return substituteDayValue(status) >= 1;
  if (FULL_LEAVE_KEYWORDS.some((keyword) => status.includes(keyword))) return !status.includes("0.5") && !status.includes("반차") && !status.includes("반반차");
  return status.includes("휴가") && !status.includes("0.5") && !status.includes("반차");
}

function buildDayoffJudgment({ baseAllowance, basicDayoffUsed, explicitSubDayoffUsed, baseExcess, substituteNeeded, availableSubstitute, shortage, remaining }) {
  if (shortage > 0) {
    return `휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 대체휴무 필요 ${formatDays(substituteNeeded)}, 가용 ${formatDays(availableSubstitute)} → ${formatDays(shortage)} 초과 사용`;
  }
  if (baseExcess > 0) {
    return `휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)} · 초과 ${formatDays(baseExcess)}은 대체휴무 여분 활용 · 잔여 ${formatDays(remaining)}`;
  }
  if (explicitSubDayoffUsed > 0) {
    return `기본 휴무 정상 · 대체휴무 ${formatDays(explicitSubDayoffUsed)} 사용 · 잔여 ${formatDays(remaining)}`;
  }
  if (basicDayoffUsed > baseAllowance) return `휴무 기준 초과 ${formatDays(basicDayoffUsed - baseAllowance)}`;
  return `정상 · 기본 휴무 ${formatDays(basicDayoffUsed)} / 기준 ${formatDays(baseAllowance)}`;
}

function buildDiagnostics({ company, plan, attendance, matchRate, planIds, attendanceIds, baseAllowance }) {
  const messages = [`${companyLabels[company]} 휴무 기준: ${baseAllowance}일`];
  if (plan.detectedCompany && plan.detectedCompany !== company) messages.push(`계획표 내용은 ‘${companyLabels[plan.detectedCompany]}’로 감지되었습니다.`);
  if (attendance.detectedCompany && attendance.detectedCompany !== company) messages.push(`근태표 내용은 ‘${companyLabels[attendance.detectedCompany]}’로 감지되었습니다.`);
  if (matchRate < 50) messages.push(`사번 매칭률이 ${matchRate}%로 낮습니다. 서로 다른 회사 또는 월 파일인지 확인해 주세요.`);
  if (planIds.size && attendanceIds.size && matchRate >= 50) messages.push(`사번 ${[...planIds].filter((id) => attendanceIds.has(id)).length}명이 정상 매칭되었습니다.`);
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) messages.push("관리자 로그인 전에는 대체휴무 가용 수량을 0으로 계산합니다.");
  return messages;
}

function renderResult() {
  const result = state.result;
  $("#emptyState").classList.add("hidden");
  $("#resultArea").classList.remove("hidden");
  $("#missingCount").textContent = number(result.missingRows.length);
  $("#unexpectedCount").textContent = number(result.unexpectedRows.length);
  $("#dayoffExcessPeople").textContent = number(result.dayoffExcessPeople);
  $("#substituteShortagePeople").textContent = number(result.substituteShortagePeople);
  $("#matchRate").textContent = `${result.matchRate}%`;
  $("#missingTabCount").textContent = result.missingRows.length;
  $("#unexpectedTabCount").textContent = result.unexpectedRows.length;
  $("#dayoffTabCount").textContent = result.employeeSummaries.length;
  $("#resultDescription").textContent = `${result.companyLabel} · ${result.targetMonth} · ${result.cutoffDate}까지 · 기본 휴무 ${result.baseAllowance}일`;

  const banner = $("#diagnosticBanner");
  banner.className = `alert ${result.matchRate < 50 || result.diagnostics.some((message) => message.includes("감지")) ? "warning" : "success"}`;
  banner.innerHTML = result.diagnostics.map(escapeHtml).join("<br>");
  banner.classList.remove("hidden");

  const stores = [...new Set(result.employeeSummaries.map((row) => row.store).filter(Boolean))].sort();
  $("#storeFilter").innerHTML = `<option value="">전체 매장</option>${stores.map((store) => `<option>${escapeHtml(store)}</option>`).join("")}`;
  $("#searchInput").value = "";
  state.pages = { missing: 1, unexpected: 1, dayoff: 1 };
  applyFilters();
  switchResultTab("missing");
}

function applyFilters() {
  if (!state.result) return;
  const query = $("#searchInput").value.trim().toLowerCase();
  const store = $("#storeFilter").value;
  const matches = (row) => {
    const haystack = `${row.store} ${row.employeeId} ${row.name} ${row.date || ""} ${row.planStatus || ""} ${row.judgment || ""}`.toLowerCase();
    return (!query || haystack.includes(query)) && (!store || row.store === store);
  };
  state.filtered.missing = state.result.missingRows.filter(matches);
  state.filtered.unexpected = state.result.unexpectedRows.filter(matches);
  state.filtered.dayoff = state.result.employeeSummaries.filter(matches);
  state.pages = { missing: 1, unexpected: 1, dayoff: 1 };
  renderAllTables();
}

function renderAllTables() {
  renderIssueTable("missing", "missingTableBody", "missingPagination");
  renderIssueTable("unexpected", "unexpectedTableBody", "unexpectedPagination");
  renderDayoffTable();
}

function renderIssueTable(kind, bodyId, paginationId) {
  const rows = state.filtered[kind];
  const page = state.pages[kind];
  const pageRows = rows.slice((page - 1) * state.pageSize, page * state.pageSize);
  $(`#${bodyId}`).innerHTML = pageRows.length ? pageRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.companyLabel)}</td><td>${escapeHtml(row.store)}</td><td>${escapeHtml(row.employeeId)}</td>
      <td><strong>${escapeHtml(row.name)}</strong></td><td>${escapeHtml(row.date)}</td><td>${escapeHtml(row.weekday)}</td>
      <td><span class="plan-pill">${escapeHtml(row.planStatus)}</span></td><td>${escapeHtml(row.clockStatus)}</td>
      <td><span class="${kind === "missing" ? "status-pill" : "warning-pill"}">${escapeHtml(row.result)}</span></td>
    </tr>`).join("") : `<tr><td colspan="9" style="padding:45px;color:#7b8598">조건에 맞는 항목이 없습니다.</td></tr>`;
  renderPagination(kind, paginationId, rows.length, () => renderIssueTable(kind, bodyId, paginationId));
}

function renderDayoffTable() {
  const rows = state.filtered.dayoff;
  const pageRows = rows.slice((state.pages.dayoff - 1) * state.pageSize, state.pages.dayoff * state.pageSize);
  $("#dayoffTableBody").innerHTML = pageRows.length ? pageRows.map((row) => {
    const statusClass = row.shortage > 0 ? "status-pill" : row.baseExcess > 0 || row.explicitSubDayoffUsed > 0 ? "warning-pill" : "success-pill";
    return `<tr>
      <td>${escapeHtml(row.store)}</td><td>${escapeHtml(row.employeeId)}</td><td><strong>${escapeHtml(row.name)}</strong></td>
      <td>${formatDays(row.baseAllowance)}</td><td>${formatDays(row.basicDayoffUsed)}</td><td>${formatDays(row.explicitSubDayoffUsed)}</td>
      <td>${formatDays(row.substituteNeeded)}</td><td>${formatDays(row.availableSubstitute)}</td><td>${formatDays(row.substituteApplied)}</td>
      <td>${formatDays(row.remaining)}</td><td>${formatDays(row.shortage)}</td><td class="message-cell"><span class="${statusClass}">${escapeHtml(row.judgment)}</span></td>
    </tr>`;
  }).join("") : `<tr><td colspan="12" style="padding:45px;color:#7b8598">계획표 인원이 없습니다.</td></tr>`;
  renderPagination("dayoff", "dayoffPagination", rows.length, renderDayoffTable);
}

function renderPagination(kind, elementId, rowCount, rerender) {
  const totalPages = Math.max(1, Math.ceil(rowCount / state.pageSize));
  const current = state.pages[kind];
  const visible = [];
  for (let page = Math.max(1, current - 2); page <= Math.min(totalPages, current + 2); page += 1) visible.push(page);
  $(`#${elementId}`).innerHTML = `
    <button class="page-button" data-page="${Math.max(1, current - 1)}">‹</button>
    ${visible.map((page) => `<button class="page-button ${page === current ? "active" : ""}" data-page="${page}">${page}</button>`).join("")}
    <button class="page-button" data-page="${Math.min(totalPages, current + 1)}">›</button>`;
  $$(`#${elementId} .page-button`).forEach((button) => button.addEventListener("click", () => {
    state.pages[kind] = Number(button.dataset.page); rerender();
  }));
}

function switchResultTab(tabName) {
  state.activeResultTab = tabName;
  $$(".inner-tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.resultTab === tabName));
  $$(".result-pane").forEach((pane) => pane.classList.remove("active"));
  $(`#${tabName}ResultPane`).classList.add("active");
}

function exportResults() {
  if (!state.result) return;
  const result = state.result;
  const issueToRow = (row) => ({
    구분: row.companyLabel, 매장명: row.store, 사번: row.employeeId, 이름: row.name,
    일자: row.date, 요일: row.weekday, 계획표: row.planStatus,
    실제출근시간: row.actualIn, 변경출근시간: row.changedIn, 판정: row.result, 사유: row.reason,
  });
  const dayoffRows = result.employeeSummaries.map((row) => ({
    구분: row.companyLabel, 매장명: row.store, 사번: row.employeeId, 이름: row.name,
    "기준 휴무": row.baseAllowance, "기본 휴무 사용": row.basicDayoffUsed,
    "표기 대체휴무": row.explicitSubDayoffUsed, "대체휴무 필요": row.substituteNeeded,
    "가용 대체휴무": row.availableSubstitute, "대체휴무 적용": row.substituteApplied,
    "대체휴무 잔여": row.remaining, "대체휴무 부족": row.shortage, 판정: row.judgment,
  }));
  const infoRows = [
    ["구분", result.companyLabel], ["대상 월", result.targetMonth], ["비교 기준일", result.cutoffDate],
    ["기본 휴무 기준", result.baseAllowance], ["계획표 파일", result.planFileName], ["근태표 파일", result.attendanceFileName],
    ["계획표 인원", result.planPeople], ["근태표 인원", result.attendancePeople], ["사번 매칭률", `${result.matchRate}%`],
    ["근무인데 미출근", result.missingRows.length], ["휴무인데 출근", result.unexpectedRows.length],
    ["휴무 기준 초과 인원", result.dayoffExcessPeople], ["대체휴무 부족 인원", result.substituteShortagePeople],
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.missingRows.map(issueToRow)), "근무인데 미출근");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.unexpectedRows.map(issueToRow)), "휴무인데 출근");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dayoffRows), "휴무_대체휴무 판정");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(infoRows), "분석정보");
  XLSX.writeFile(workbook, `${result.companyLabel}_${result.targetMonth}_근태종합점검.xlsx`);
}

async function saveClosure() {
  if (!state.result) return;
  if (!(state.backend.available && state.backend.configured)) {
    showToast("D1 연결이 필요합니다. 월 마감은 서버에만 저장합니다."); return;
  }
  if (!state.backend.loggedIn) { openLogin(); return; }
  const result = state.result;
  const payload = {
    company: result.company, month: result.targetMonth, cutoffDate: result.cutoffDate, checkMode: result.checkMode,
    planFileName: result.planFileName, attendanceFileName: result.attendanceFileName,
    planPeople: result.planPeople, attendancePeople: result.attendancePeople, matchedPeople: result.matchedPeople,
    matchRate: result.matchRate, missingCount: result.missingRows.length, missingPeople: result.missingPeople,
    unexpectedCount: result.unexpectedRows.length, unexpectedPeople: result.unexpectedPeople,
    dayoffExcessPeople: result.dayoffExcessPeople, substituteShortagePeople: result.substituteShortagePeople,
    issueRows: [...result.missingRows, ...result.unexpectedRows], employeeSummaries: result.employeeSummaries,
  };
  try {
    const response = await fetch("/api/closures", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(payload) });
    if (response.status === 401) { await checkBackend(); openLogin(); return; }
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "저장 실패");
    showToast("월 마감 저장 완료 · 대체휴무가 만료일 순서로 차감되었습니다.");
    await loadHistory();
    state.balances = await loadBalancesForAnalysis(result.company, result.targetMonth);
  } catch (error) { showToast(error.message); }
}

async function loadHistory() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    renderHistory([]); return;
  }
  try {
    const response = await fetch("/api/closures", { cache: "no-store" });
    if (!response.ok) throw new Error("기록 조회 실패");
    renderHistory((await response.json()).items || []);
  } catch (error) { showToast(error.message); }
}

function renderHistory(list) {
  $("#historyEmpty").classList.toggle("hidden", list.length > 0);
  $("#historyList").innerHTML = list.map((item) => {
    const label = companyLabels[item.company] || item.company;
    return `<article class="history-card">
      <div><div class="month">${escapeHtml(item.month)}</div><div class="history-meta">${escapeHtml(label)} · 기준 ${escapeHtml(item.cutoff_date || "-")}</div></div>
      <div><strong>${escapeHtml(item.plan_file_name || "계획표")}</strong><div class="history-meta">${escapeHtml(item.attendance_file_name || "근태표")} · ${item.created_at ? new Date(`${item.created_at}Z`).toLocaleString("ko-KR") : ""}</div></div>
      <div class="history-kpis">
        <div><span>미출근</span><strong>${number(item.missing_count || 0)}</strong></div>
        <div><span>휴무 출근</span><strong>${number(item.unexpected_count || 0)}</strong></div>
        <div><span>대체휴무 부족</span><strong>${number(item.substitute_shortage_people || 0)}</strong></div>
      </div>
    </article>`;
  }).join("");
}

async function saveGrant(event) {
  event.preventDefault();
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) { openLogin(); return; }
  const payload = {
    company: $("#grantCompany").value,
    employeeId: normalizeEmployeeId($("#grantEmployeeId").value),
    employeeName: text($("#grantEmployeeName").value),
    store: text($("#grantStore").value),
    grantMonth: $("#grantMonth").value,
    grantedDays: Number($("#grantDays").value),
    validFrom: $("#grantValidFrom").value,
    validTo: $("#grantValidTo").value,
    reason: text($("#grantReason").value),
    note: text($("#grantNote").value),
  };
  if (!payload.employeeId || !payload.employeeName || !payload.grantMonth || !payload.validFrom || !payload.validTo || !(payload.grantedDays > 0)) {
    showToast("사번·이름·발생월·부여일수·사용기간을 확인해 주세요."); return;
  }
  if (payload.validFrom > payload.validTo) { showToast("사용 종료일은 시작일보다 빠를 수 없습니다."); return; }
  try {
    const response = await fetch("/api/substitute-grants", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(payload) });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "저장 실패");
    showToast("대체휴무를 부여했습니다.");
    $("#grantEmployeeId").value = ""; $("#grantEmployeeName").value = ""; $("#grantStore").value = "";
    $("#grantDays").value = "1"; $("#grantReason").value = ""; $("#grantNote").value = "";
    await loadGrants();
  } catch (error) { showToast(error.message); }
}

async function loadGrants() {
  if (!(state.backend.available && state.backend.configured && state.backend.loggedIn)) {
    $("#grantTableBody").innerHTML = `<tr><td colspan="12" style="padding:45px;color:#7b8598">관리자 로그인 후 조회할 수 있습니다.</td></tr>`;
    renderGrantSummary([]); return;
  }
  try {
    const company = $("#grantCompanyFilter").value;
    const query = company ? `?company=${encodeURIComponent(company)}` : "";
    const response = await fetch(`/api/substitute-grants${query}`, { cache: "no-store" });
    if (!response.ok) throw new Error("대체휴무 조회 실패");
    renderGrants((await response.json()).items || []);
  } catch (error) { showToast(error.message); }
}

function renderGrants(items) {
  renderGrantSummary(items);
  $("#grantTableBody").innerHTML = items.length ? items.map((item) => {
    const status = grantStatus(item);
    const statusClass = status === "사용 가능" ? "success-pill" : status === "만료" ? "status-pill" : "neutral-pill";
    return `<tr>
      <td>${escapeHtml(companyLabels[item.company] || item.company)}</td><td>${escapeHtml(item.store || "")}</td>
      <td>${escapeHtml(item.employee_id)}</td><td><strong>${escapeHtml(item.employee_name)}</strong></td><td>${escapeHtml(item.grant_month)}</td>
      <td>${formatDays(item.granted_days)}</td><td>${formatDays(item.used_days)}</td><td>${formatDays(item.remaining_days)}</td>
      <td>${escapeHtml(item.valid_from)} ~ ${escapeHtml(item.valid_to)}</td><td><span class="${statusClass}">${status}</span></td>
      <td class="message-cell">${escapeHtml([item.reason, item.note].filter(Boolean).join(" · "))}</td>
      <td><button class="btn danger grant-delete" data-id="${escapeHtml(item.id)}" type="button">삭제</button></td>
    </tr>`;
  }).join("") : `<tr><td colspan="12" style="padding:45px;color:#7b8598">등록된 대체휴무가 없습니다.</td></tr>`;
  $$(".grant-delete").forEach((button) => button.addEventListener("click", () => deleteGrant(button.dataset.id)));
}

function renderGrantSummary(items) {
  const granted = items.reduce((sum, item) => sum + Number(item.granted_days || 0), 0);
  const used = items.reduce((sum, item) => sum + Number(item.used_days || 0), 0);
  const remaining = items.reduce((sum, item) => sum + Number(item.remaining_days || 0), 0);
  const expiring = items.filter((item) => Number(item.remaining_days) > 0 && daysUntil(item.valid_to) >= 0 && daysUntil(item.valid_to) <= 30).reduce((sum, item) => sum + Number(item.remaining_days), 0);
  $("#grantSummary").innerHTML = [
    ["총 부여", formatDays(granted)], ["총 사용", formatDays(used)], ["현재 잔여", formatDays(remaining)], ["30일 내 만료", formatDays(expiring)],
  ].map(([label, value]) => `<div class="summary-chip"><span>${label}</span><strong>${value}</strong></div>`).join("");
}

async function deleteGrant(id) {
  if (!confirm("이 대체휴무 부여 기록을 삭제하시겠습니까? 이미 월 마감에 사용된 기록은 삭제되지 않습니다.")) return;
  try {
    const response = await fetch(`/api/substitute-grants/${encodeURIComponent(id)}`, { method: "DELETE" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "삭제 실패");
    showToast("대체휴무 부여 기록을 삭제했습니다.");
    await loadGrants();
  } catch (error) { showToast(error.message); }
}

function grantStatus(item) {
  if (Number(item.remaining_days) <= 0) return "소진";
  const today = toISODate(new Date());
  if (item.valid_to < today) return "만료";
  if (item.valid_from > today) return "사용 전";
  return "사용 가능";
}

async function checkBackend() {
  try {
    const response = await fetch("/api/auth", { cache: "no-store" });
    if (!response.ok) throw new Error();
    const data = await response.json();
    state.backend = { available: true, configured: Boolean(data.configured), loggedIn: Boolean(data.loggedIn) };
  } catch { state.backend = { available: false, configured: false, loggedIn: false }; }
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
    badge.className = "badge local"; badge.textContent = "서버 설정 확인 필요";
    loginButton.classList.add("hidden"); logoutButton.classList.add("hidden");
  }
}

function openLogin() { $("#loginError").textContent = ""; $("#passwordInput").value = ""; $("#loginDialog").showModal(); }
async function login(event) {
  event.preventDefault();
  try {
    const response = await fetch("/api/auth", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ password: $("#passwordInput").value }) });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "로그인 실패");
    $("#loginDialog").close(); await checkBackend(); await Promise.all([loadHistory(), loadGrants()]); showToast("관리자 로그인되었습니다.");
  } catch (error) { $("#loginError").textContent = error.message; }
}
async function logout() { await fetch("/api/auth", { method: "DELETE" }); await checkBackend(); showToast("로그아웃되었습니다."); }

function switchView(view) {
  $$(".tab[data-view]").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === view));
  $$(".view").forEach((section) => section.classList.remove("active"));
  $(`#${view}View`).classList.add("active");
  if (view === "history") loadHistory();
  if (view === "substitute") loadGrants();
}

function resetAll() {
  state.planFile = null; state.attendanceFile = null; state.result = null; state.balances = {};
  $("#planFile").value = ""; $("#attendanceFile").value = "";
  setPlanFile(null); setAttendanceFile(null); setDefaultDates();
  $("#resultArea").classList.add("hidden"); $("#emptyState").classList.remove("hidden");
}

function selectedCompany() { return document.querySelector('input[name="company"]:checked').value; }
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
function findHeaderRow(matrix, requiredHeaders) { return matrix.slice(0, 15).findIndex((row) => { const headers = row.map(normalizeHeader); return requiredHeaders.every((required) => headers.includes(normalizeHeader(required))); }); }
function findHeaderIndex(headers, candidates) { for (const candidate of candidates) { const index = headers.indexOf(normalizeHeader(candidate)); if (index >= 0) return index; } return -1; }
function normalizeHeader(value) { return text(value).replace(/\s+/g, "").replace(/[\[\]]/g, ""); }
function normalizeStatus(value) { return text(value).replace(/\s+/g, ""); }
function normalizeEmployeeId(value) { return text(value).toUpperCase().replace(/\s+/g, ""); }
function normalizeStore(value) { return text(value).replace(/^\d+_/, "").replace(/홈플러스/g, "").replace(/전자랜드/g, "").replace(/점$/g, "").replace(/\s+/g, "").toLowerCase(); }
function text(value) { return value == null ? "" : String(value).trim(); }
function parseDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toISODate(value);
  const digits = text(value).replace(/\D/g, "");
  if (digits.length >= 8) return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  const parsed = new Date(value); return Number.isNaN(parsed.getTime()) ? "" : toISODate(parsed);
}
function detectCompanyFromPlan(rows) { const hp = rows.filter((row) => row.store.includes("홈플러스")).length; return hp > rows.length / 2 ? "homeplus" : hp === 0 ? "electroland" : null; }
function detectCompanyFromAttendance(rows) { const hp = rows.filter((row) => row.location.includes("홈플러스")).length; return hp > rows.length / 2 ? "homeplus" : hp === 0 ? "electroland" : null; }
function groupBy(items, selector) { const map = new Map(); items.forEach((item) => { const key = selector(item); if (!map.has(key)) map.set(key, []); map.get(key).push(item); }); return map; }
function toISODate(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`; }
function number(value) { return new Intl.NumberFormat("ko-KR").format(value || 0); }
function roundHalf(value) { return Math.round((Number(value) || 0) * 2) / 2; }
function formatDays(value) { const numberValue = roundHalf(value); return `${Number.isInteger(numberValue) ? numberValue : numberValue.toFixed(1)}일`; }
function issueSort(a, b) { return a.date.localeCompare(b.date) || a.store.localeCompare(b.store) || a.name.localeCompare(b.name); }
function daysUntil(dateText) { const today = new Date(`${toISODate(new Date())}T00:00:00`); const target = new Date(`${dateText}T00:00:00`); return Math.ceil((target - today) / 86400000); }
function escapeHtml(value) { return text(value).replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char])); }
let toastTimer;
function showToast(message) { const toast = $("#toast"); toast.textContent = message; toast.classList.add("show"); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove("show"), 3800); }
