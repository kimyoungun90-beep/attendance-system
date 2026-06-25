const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const state = {
  planFile: null,
  attendanceFile: null,
  result: null,
  filteredRows: [],
  page: 1,
  pageSize: 50,
  backend: { available: false, configured: false, loggedIn: false },
};

const companyLabels = { homeplus: "홈플러스", electroland: "전자랜드" };
const weekdayLabels = ["일", "월", "화", "수", "목", "금", "토"];

init();

async function init() {
  setDefaultDates();
  bindEvents();
  setupDropzone("planDropzone", "planFile", setPlanFile);
  setupDropzone("attendanceDropzone", "attendanceFile", setAttendanceFile);
  await checkBackend();
}

function bindEvents() {
  $("#analyzeButton").addEventListener("click", analyzeFiles);
  $("#resetButton").addEventListener("click", resetAll);
  $("#searchInput").addEventListener("input", applyFilters);
  $("#storeFilter").addEventListener("change", applyFilters);
  $("#exportButton").addEventListener("click", exportResults);
  $("#saveClosureButton").addEventListener("click", saveClosure);
  $("#refreshHistoryButton").addEventListener("click", loadHistory);
  $("#loginButton").addEventListener("click", openLogin);
  $("#logoutButton").addEventListener("click", logout);
  $("#loginCancel").addEventListener("click", () => $("#loginDialog").close());
  $("#loginForm").addEventListener("submit", login);
  $("#targetMonth").addEventListener("change", syncCutoffWithMonth);
  $$(".tab[data-view]").forEach((tab) => tab.addEventListener("click", () => switchView(tab.dataset.view)));
}

function setDefaultDates() {
  const now = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  $("#targetMonth").value = month;
  $("#cutoffDate").value = toISODate(now);
}

function syncCutoffWithMonth() {
  const month = $("#targetMonth").value;
  if (!month) return;
  const [year, monthNumber] = month.split("-").map(Number);
  const now = new Date();
  const isCurrent = now.getFullYear() === year && now.getMonth() + 1 === monthNumber;
  const lastDay = new Date(year, monthNumber, 0);
  $("#cutoffDate").value = toISODate(isCurrent ? now : lastDay);
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
  try {
    if (!window.XLSX) throw new Error("엑셀 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하세요.");
    if (!state.planFile || !state.attendanceFile) throw new Error("계획표와 실제 근태표를 모두 선택해 주세요.");
    const targetMonth = $("#targetMonth").value;
    const cutoffDate = $("#cutoffDate").value;
    if (!targetMonth || !cutoffDate) throw new Error("대상 월과 비교 기준일을 입력해 주세요.");
    if (!cutoffDate.startsWith(targetMonth)) throw new Error("비교 기준일은 대상 월 안의 날짜여야 합니다.");

    const button = $("#analyzeButton");
    button.disabled = true;
    button.textContent = "파일 비교 중...";

    const [planMatrix, attendanceMatrix] = await Promise.all([
      fileToMatrix(state.planFile), fileToMatrix(state.attendanceFile),
    ]);

    const plan = parsePlan(planMatrix);
    const attendance = parseAttendance(attendanceMatrix, targetMonth);
    const result = compareAttendance({
      plan,
      attendance,
      company: selectedCompany(),
      targetMonth,
      cutoffDate,
      checkMode: $("#checkMode").value,
      excludedKeywords: getExcludedKeywords(),
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
    showToast(`분석 완료: ${result.rows.length}건의 출근 누락을 찾았습니다.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "분석 중 오류가 발생했습니다.");
  } finally {
    const button = $("#analyzeButton");
    button.disabled = false;
    button.textContent = "출근 누락자 분석";
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

function compareAttendance({ plan, attendance, company, targetMonth, cutoffDate, checkMode, excludedKeywords }) {
  const attendanceIds = new Set(attendance.rows.map((row) => row.employeeId));
  const planIds = new Set(plan.rows.map((row) => row.employeeId));
  const matchedIds = [...planIds].filter((id) => attendanceIds.has(id));
  const matchRate = planIds.size ? Math.round((matchedIds.length / planIds.size) * 1000) / 10 : 0;

  const attendanceById = groupBy(attendance.rows, (row) => row.employeeId);
  const selectedPlans = choosePlanRows(plan.rows, attendanceById);
  const clocked = new Map();
  attendance.rows.forEach((row) => {
    const key = `${row.employeeId}|${row.date}`;
    const hasClockIn = Boolean(row.actualIn || row.changedIn);
    const existing = clocked.get(key) || { hasClockIn: false, actualIn: "", changedIn: "", location: "" };
    clocked.set(key, {
      hasClockIn: existing.hasClockIn || hasClockIn,
      actualIn: row.actualIn || existing.actualIn,
      changedIn: row.changedIn || existing.changedIn,
      location: row.location || existing.location,
    });
  });

  const cutoff = new Date(`${cutoffDate}T23:59:59`);
  const [year, month] = targetMonth.split("-").map(Number);
  const lastDay = Math.min(new Date(year, month, 0).getDate(), cutoff.getDate());
  const rows = [];

  selectedPlans.forEach((person) => {
    for (let day = 1; day <= lastDay; day += 1) {
      const planStatus = text(person.plans[day]);
      if (!shouldCheck(planStatus, checkMode, excludedKeywords)) continue;
      const date = `${targetMonth}-${String(day).padStart(2, "0")}`;
      const attendanceValue = clocked.get(`${person.employeeId}|${date}`);
      if (attendanceValue?.hasClockIn) continue;
      const dateObject = new Date(`${date}T00:00:00`);
      rows.push({
        company,
        companyLabel: companyLabels[company],
        store: person.store,
        employeeId: person.employeeId,
        name: person.name,
        date,
        weekday: weekdayLabels[dateObject.getDay()],
        planStatus: planStatus || "공백",
        actualIn: attendanceValue?.actualIn || "",
        changedIn: attendanceValue?.changedIn || "",
        clockStatus: "미기록",
        result: "출근 누락",
        reason: planStatus ? `${planStatus} 계획이나 출근시간 없음` : "계획 공백이나 출근시간 없음",
        duplicatePlanNote: person.duplicatePlanNote || "",
      });
    }
  });

  rows.sort((a, b) => a.date.localeCompare(b.date) || a.store.localeCompare(b.store) || a.name.localeCompare(b.name));
  const missingPeople = new Set(rows.map((row) => row.employeeId)).size;
  const diagnostics = buildDiagnostics({ company, plan, attendance, matchRate, planIds, attendanceIds });
  return {
    rows,
    planPeople: planIds.size,
    attendancePeople: attendanceIds.size,
    matchedPeople: matchedIds.length,
    matchRate,
    missingPeople,
    diagnostics,
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

function shouldCheck(planStatus, mode, excludedKeywords) {
  const status = text(planStatus);
  if (excludedKeywords.some((keyword) => keyword && status.includes(keyword))) return false;
  if (mode === "broad") return true;
  if (!status) return true;
  return status.includes("근무") || status.includes("출근");
}

function buildDiagnostics({ company, plan, attendance, matchRate, planIds, attendanceIds }) {
  const messages = [];
  if (plan.detectedCompany && plan.detectedCompany !== company) messages.push(`계획표 내용은 ‘${companyLabels[plan.detectedCompany]}’로 감지되었습니다.`);
  if (attendance.detectedCompany && attendance.detectedCompany !== company) messages.push(`근태표 내용은 ‘${companyLabels[attendance.detectedCompany]}’로 감지되었습니다.`);
  if (matchRate < 50) messages.push(`사번 매칭률이 ${matchRate}%로 낮습니다. 서로 다른 회사 또는 월 파일인지 확인해 주세요.`);
  if (planIds.size && attendanceIds.size && matchRate >= 50) messages.push(`사번 ${[...planIds].filter((id) => attendanceIds.has(id)).length}명이 정상 매칭되었습니다.`);
  return messages;
}

function renderResult() {
  const result = state.result;
  $("#emptyState").classList.add("hidden");
  $("#resultArea").classList.remove("hidden");
  $("#missingCount").textContent = number(result.rows.length);
  $("#missingPeopleCount").textContent = number(result.missingPeople);
  $("#planPeopleCount").textContent = number(result.planPeople);
  $("#matchRate").textContent = `${result.matchRate}%`;
  $("#resultDescription").textContent = `${result.companyLabel} · ${result.targetMonth} · ${result.cutoffDate}까지 · ${result.checkMode === "strict" ? "근무/출근/공백 기준" : "휴무·휴가 외 전체 기준"}`;

  const banner = $("#diagnosticBanner");
  if (result.diagnostics.length) {
    banner.className = `alert ${result.matchRate < 50 || result.diagnostics.some((m) => m.includes("감지")) ? "warning" : "success"}`;
    banner.innerHTML = result.diagnostics.map(escapeHtml).join("<br>");
    banner.classList.remove("hidden");
  } else banner.classList.add("hidden");

  const stores = [...new Set(result.rows.map((row) => row.store))].sort();
  $("#storeFilter").innerHTML = `<option value="">전체 매장</option>${stores.map((store) => `<option>${escapeHtml(store)}</option>`).join("")}`;
  $("#searchInput").value = "";
  state.page = 1;
  applyFilters();
}

function applyFilters() {
  if (!state.result) return;
  const query = $("#searchInput").value.trim().toLowerCase();
  const store = $("#storeFilter").value;
  state.filteredRows = state.result.rows.filter((row) => {
    const haystack = `${row.store} ${row.employeeId} ${row.name} ${row.date} ${row.planStatus}`.toLowerCase();
    return (!query || haystack.includes(query)) && (!store || row.store === store);
  });
  state.page = 1;
  renderTable();
}

function renderTable() {
  const start = (state.page - 1) * state.pageSize;
  const pageRows = state.filteredRows.slice(start, start + state.pageSize);
  $("#resultTableBody").innerHTML = pageRows.length ? pageRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.companyLabel)}</td>
      <td>${escapeHtml(row.store)}</td>
      <td>${escapeHtml(row.employeeId)}</td>
      <td><strong>${escapeHtml(row.name)}</strong></td>
      <td>${escapeHtml(row.date)}</td>
      <td>${escapeHtml(row.weekday)}</td>
      <td><span class="plan-pill">${escapeHtml(row.planStatus)}</span></td>
      <td>${escapeHtml(row.clockStatus)}</td>
      <td><span class="status-pill">${escapeHtml(row.result)}</span></td>
    </tr>`).join("") : `<tr><td colspan="9" style="padding:45px;color:#7b8598">조건에 맞는 누락자가 없습니다.</td></tr>`;
  renderPagination();
}

function renderPagination() {
  const totalPages = Math.max(1, Math.ceil(state.filteredRows.length / state.pageSize));
  const visible = [];
  for (let page = Math.max(1, state.page - 2); page <= Math.min(totalPages, state.page + 2); page += 1) visible.push(page);
  $("#pagination").innerHTML = `
    <button class="page-button" data-page="${Math.max(1, state.page - 1)}">‹</button>
    ${visible.map((page) => `<button class="page-button ${page === state.page ? "active" : ""}" data-page="${page}">${page}</button>`).join("")}
    <button class="page-button" data-page="${Math.min(totalPages, state.page + 1)}">›</button>`;
  $$("#pagination .page-button").forEach((button) => button.addEventListener("click", () => {
    state.page = Number(button.dataset.page); renderTable();
  }));
}

function exportResults() {
  if (!state.result) return;
  const result = state.result;
  const detailRows = result.rows.map((row) => ({
    구분: row.companyLabel,
    매장명: row.store,
    사번: row.employeeId,
    이름: row.name,
    누락일자: row.date,
    요일: row.weekday,
    계획표: row.planStatus,
    실제출근시간: row.actualIn,
    변경출근시간: row.changedIn,
    판정: row.result,
    사유: row.reason,
    중복계획처리: row.duplicatePlanNote,
  }));
  const summaryMap = new Map();
  result.rows.forEach((row) => {
    const key = `${row.employeeId}|${row.name}|${row.store}`;
    if (!summaryMap.has(key)) summaryMap.set(key, { 구분: row.companyLabel, 매장명: row.store, 사번: row.employeeId, 이름: row.name, "누락 횟수": 0, "누락 날짜": [] });
    const item = summaryMap.get(key); item["누락 횟수"] += 1; item["누락 날짜"].push(row.date);
  });
  const personRows = [...summaryMap.values()].map((item) => ({ ...item, "누락 날짜": item["누락 날짜"].join(", ") })).sort((a, b) => b["누락 횟수"] - a["누락 횟수"]);
  const infoRows = [
    ["구분", result.companyLabel], ["대상 월", result.targetMonth], ["비교 기준일", result.cutoffDate],
    ["판정 범위", result.checkMode === "strict" ? "근무·출근·공백만 검사" : "휴무·휴가 외 모든 계획 검사"],
    ["계획표 파일", result.planFileName], ["근태표 파일", result.attendanceFileName],
    ["계획표 인원", result.planPeople], ["근태표 인원", result.attendancePeople],
    ["사번 매칭률", `${result.matchRate}%`], ["출근 누락 건수", result.rows.length], ["출근 누락 인원", result.missingPeople],
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailRows), "누락자 상세");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(personRows), "사람별 요약");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(infoRows), "분석정보");
  XLSX.writeFile(workbook, `${result.companyLabel}_${result.targetMonth}_출근누락.xlsx`);
}

async function saveClosure() {
  if (!state.result) return;
  const result = state.result;
  const payload = {
    company: result.company,
    month: result.targetMonth,
    cutoffDate: result.cutoffDate,
    checkMode: result.checkMode,
    planFileName: result.planFileName,
    attendanceFileName: result.attendanceFileName,
    planPeople: result.planPeople,
    attendancePeople: result.attendancePeople,
    matchedPeople: result.matchedPeople,
    matchRate: result.matchRate,
    missingCount: result.rows.length,
    missingPeople: result.missingPeople,
    rows: result.rows,
  };

  if (state.backend.available && state.backend.configured) {
    if (!state.backend.loggedIn) { openLogin(); return; }
    try {
      const response = await fetch("/api/closures", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(payload) });
      if (response.status === 401) { await checkBackend(); openLogin(); return; }
      if (!response.ok) throw new Error((await response.json()).error || "저장 실패");
      showToast("D1 데이터베이스에 월 마감 결과를 저장했습니다.");
      await loadHistory();
    } catch (error) { showToast(error.message); }
  } else {
    saveLocalClosure(payload);
    showToast("현재 브라우저에 월 마감 결과를 저장했습니다.");
    await loadHistory();
  }
}

function saveLocalClosure(payload) {
  const list = JSON.parse(localStorage.getItem("attendanceClosures") || "[]");
  list.unshift({ ...payload, id: crypto.randomUUID(), createdAt: new Date().toISOString() });
  localStorage.setItem("attendanceClosures", JSON.stringify(list.slice(0, 120)));
}

async function loadHistory() {
  let list = [];
  try {
    if (state.backend.available && state.backend.configured && state.backend.loggedIn) {
      const response = await fetch("/api/closures");
      if (!response.ok) throw new Error("기록 조회 실패");
      list = (await response.json()).items || [];
    } else list = JSON.parse(localStorage.getItem("attendanceClosures") || "[]");
  } catch (error) { showToast(error.message); }
  renderHistory(list);
}

function renderHistory(list) {
  $("#historyEmpty").classList.toggle("hidden", list.length > 0);
  $("#historyList").innerHTML = list.map((item) => {
    const label = companyLabels[item.company] || item.companyLabel || item.company;
    const created = item.created_at || item.createdAt;
    return `<article class="history-card">
      <div><div class="month">${escapeHtml(item.month)}</div><div class="history-meta">${escapeHtml(label)} · 기준 ${escapeHtml(item.cutoff_date || item.cutoffDate || "-")}</div></div>
      <div><strong>${escapeHtml(item.plan_file_name || item.planFileName || "계획표")}</strong><div class="history-meta">${escapeHtml(item.attendance_file_name || item.attendanceFileName || "근태표")} · ${created ? new Date(created).toLocaleString("ko-KR") : ""}</div></div>
      <div class="history-kpis"><div><span>누락 건수</span><strong>${number(item.missing_count ?? item.missingCount ?? 0)}</strong></div><div><span>누락 인원</span><strong>${number(item.missing_people ?? item.missingPeople ?? 0)}</strong></div></div>
    </article>`;
  }).join("");
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
  const login = $("#loginButton");
  const logout = $("#logoutButton");
  if (state.backend.available && state.backend.configured) {
    badge.className = "badge cloud";
    badge.textContent = state.backend.loggedIn ? "D1 영구 저장 연결" : "D1 로그인 필요";
    login.classList.toggle("hidden", state.backend.loggedIn);
    logout.classList.toggle("hidden", !state.backend.loggedIn);
  } else {
    badge.className = "badge local";
    badge.textContent = "브라우저 임시 저장";
    login.classList.add("hidden"); logout.classList.add("hidden");
  }
}

function openLogin() { $("#loginError").textContent = ""; $("#passwordInput").value = ""; $("#loginDialog").showModal(); }
async function login(event) {
  event.preventDefault();
  try {
    const response = await fetch("/api/auth", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ password: $("#passwordInput").value }) });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "로그인 실패");
    $("#loginDialog").close(); await checkBackend(); await loadHistory(); showToast("관리자 로그인되었습니다.");
  } catch (error) { $("#loginError").textContent = error.message; }
}
async function logout() { await fetch("/api/auth", { method: "DELETE" }); await checkBackend(); showToast("로그아웃되었습니다."); }

function switchView(view) {
  $$(".tab[data-view]").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === view));
  $("#checkerView").classList.toggle("active", view === "checker");
  $("#historyView").classList.toggle("active", view === "history");
  if (view === "history") loadHistory();
}

function resetAll() {
  state.planFile = null; state.attendanceFile = null; state.result = null;
  $("#planFile").value = ""; $("#attendanceFile").value = "";
  setPlanFile(null); setAttendanceFile(null); setDefaultDates();
  $("#resultArea").classList.add("hidden"); $("#emptyState").classList.remove("hidden");
}

function selectedCompany() { return document.querySelector('input[name="company"]:checked').value; }
function getExcludedKeywords() { return $("#excludedKeywords").value.split(/[,\n]/).map((item) => item.trim()).filter(Boolean); }
function findHeaderRow(matrix, requiredHeaders) { return matrix.slice(0, 15).findIndex((row) => { const headers = row.map(normalizeHeader); return requiredHeaders.every((required) => headers.includes(normalizeHeader(required))); }); }
function findHeaderIndex(headers, candidates) { for (const candidate of candidates) { const index = headers.indexOf(normalizeHeader(candidate)); if (index >= 0) return index; } return -1; }
function normalizeHeader(value) { return text(value).replace(/\s+/g, "").replace(/[\[\]]/g, ""); }
function normalizeEmployeeId(value) { return text(value).toUpperCase().replace(/\s+/g, ""); }
function normalizeStore(value) { return text(value).replace(/^\d+_/, "").replace(/홈플러스/g, "").replace(/전자랜드/g, "").replace(/점$/g, "").replace(/\s+/g, "").toLowerCase(); }
function text(value) { return value == null ? "" : String(value).trim(); }
function parseDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toISODate(value);
  const digits = text(value).replace(/\D/g, "");
  if (digits.length >= 8) return `${digits.slice(0,4)}-${digits.slice(4,6)}-${digits.slice(6,8)}`;
  const parsed = new Date(value); return Number.isNaN(parsed.getTime()) ? "" : toISODate(parsed);
}
function detectCompanyFromPlan(rows) { const hp = rows.filter((row) => row.store.includes("홈플러스")).length; return hp > rows.length / 2 ? "homeplus" : hp === 0 ? "electroland" : null; }
function detectCompanyFromAttendance(rows) { const hp = rows.filter((row) => row.location.includes("홈플러스")).length; return hp > rows.length / 2 ? "homeplus" : hp === 0 ? "electroland" : null; }
function groupBy(items, selector) { const map = new Map(); items.forEach((item) => { const key = selector(item); if (!map.has(key)) map.set(key, []); map.get(key).push(item); }); return map; }
function toISODate(date) { return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`; }
function number(value) { return new Intl.NumberFormat("ko-KR").format(value || 0); }
function escapeHtml(value) { return text(value).replace(/[&<>'"]/g, (char) => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", "'":"&#39;", '"':"&quot;" }[char])); }
let toastTimer;
function showToast(message) { const toast = $("#toast"); toast.textContent = message; toast.classList.add("show"); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove("show"), 3300); }
