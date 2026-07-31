import {
  ELECTROLAND_MANAGERS,
  ROUTE_LABELS,
  buildAnalysis,
  buildAnnualWorkbookFile,
  buildAttendanceWorkbookFile,
  buildClosingWorkbookFile,
  getDayoffAllowance,
  parseClosingWorkbook,
  parseLedgerWorkbook,
  parseMasterWorkbook,
  parseTargetMonth,
  readWorkbook,
} from "./attendance-engine.js?v=clean4";

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const state = {
  backendAvailable: false,
  planFile: null,
  attendanceFile: null,
  annualFile: null,
  evidenceFile: null,
  closingFile: null,
  managerFiles: new Map(),
  workforceRows: [],
  annualLedgerRows: [],
  substituteGrants: [],
  webEvidenceRows: [],
  analysis: null,
};

init();

function init() {
  setDefaultDates();
  renderManagerUploadGrid();
  bindEvents();
  setupDropzone("planDropzone", "planFile", "planFileName", "planFile");
  setupDropzone("attendanceDropzone", "attendanceFile", "attendanceFileName", "attendanceFile");
  setupDropzone("annualDropzone", "annualFile", "annualFileName", "annualFile");
  setupDropzone("evidenceDropzone", "evidenceFile", "evidenceFileName", "evidenceFile");
  setupDropzone("closingDropzone", "closingFile", "closingFileName", "closingFile");
  checkBackend().then(loadScopedData);
  syncRouteHelp();
  syncSubstituteDefaults();
  renderSubstituteGrants();
  renderWebEvidenceQueue();
}

function bindEvents() {
  $$(".tab").forEach((tab) => tab.addEventListener("click", () => switchView(tab.dataset.view)));
  $$("[data-jump]").forEach((button) => button.addEventListener("click", () => switchView(button.dataset.jump)));
  $("#routeSelect").addEventListener("change", handleScopeChange);
  $("#targetMonth").addEventListener("change", handleScopeChange);
  $("#resetButton").addEventListener("click", resetAnalysisInputs);
  $("#clearManagerFiles").addEventListener("click", clearManagerFiles);
  $("#analyzeButton").addEventListener("click", () => runAnalysis());
  $("#searchInput").addEventListener("input", renderPreview);

  $("#exportAttendanceButton").addEventListener("click", () => exportSpecificWorkbook("attendance"));
  $("#exportAnnualButton").addEventListener("click", () => exportSpecificWorkbook("annual"));
  $("#exportClosingButton").addEventListener("click", () => exportSpecificWorkbook("closing"));
  $("#exportAttendanceButtonTop").addEventListener("click", () => exportSpecificWorkbook("attendance"));
  $("#exportAnnualButtonTop").addEventListener("click", () => exportSpecificWorkbook("annual"));
  $("#exportClosingButtonTop").addEventListener("click", () => exportSpecificWorkbook("closing"));

  $("#saveWorkforceButton").addEventListener("click", () => saveMasterFile("workforce", $("#workforceFile"), parseMasterWorkbook, "workforceRows", "workforceStatus", "인력 DB"));
  $("#loadWorkforceButton").addEventListener("click", () => loadMasterData("workforce", "workforceRows", "workforceStatus", "인력 DB"));
  $("#saveAnnualLedgerButton").addEventListener("click", () => saveMasterFile("annual-ledger", $("#annualLedgerFile"), parseLedgerWorkbook, "annualLedgerRows", "annualLedgerStatus", "연차 누적 DB"));
  $("#loadAnnualLedgerButton").addEventListener("click", () => loadMasterData("annual-ledger", "annualLedgerRows", "annualLedgerStatus", "연차 누적 DB"));

  $("#substituteGrantForm").addEventListener("submit", saveSubstituteGrant);
  $("#loadSubstituteGrantsButton").addEventListener("click", loadSubstituteGrants);
  $("#loadWebEvidenceButton").addEventListener("click", loadWebEvidenceData);
  $("#clearWebEvidenceButton").addEventListener("click", clearWebEvidenceData);
}

function setDefaultDates() {
  const now = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  $("#targetMonth").value = month;
  $("#cutoffDate").value = toISODate(now);
  $("#substituteGrantMonth").value = month;
  $("#substituteValidFrom").value = `${month}-01`;
  $("#substituteValidTo").value = `${month}-${String(new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate()).padStart(2, "0")}`;
}

async function handleScopeChange() {
  state.analysis = null;
  state.workforceRows = [];
  state.annualLedgerRows = [];
  state.substituteGrants = [];
  state.webEvidenceRows = [];
  syncRouteHelp();
  syncSubstituteDefaults();
  syncScopedStatuses();
  await loadScopedData();
  renderEmptyPreview();
  syncActionState();
}

function selectedRoute() {
  return $("#routeSelect").value || "electroland";
}

function syncRouteHelp() {
  const route = selectedRoute();
  const month = parseTargetMonth($("#targetMonth").value);
  if (route === "electroland" && month.valid) {
    $("#routeHelp").textContent = `${$("#targetMonth").value} · 전자랜드 토/일 기준 휴무 ${getDayoffAllowance(route, month)}회`;
  } else {
    $("#routeHelp").textContent = `${$("#targetMonth").value} · 홈플러스 기본 휴무 월 6회`;
  }
}

function syncSubstituteDefaults() {
  const month = parseTargetMonth($("#targetMonth").value);
  $("#substituteRoute").value = selectedRoute();
  $("#substituteGrantMonth").value = $("#targetMonth").value;
  if (month.valid) {
    $("#substituteValidFrom").value = `${month.key}-01`;
    $("#substituteValidTo").value = `${month.key}-${String(month.daysInMonth).padStart(2, "0")}`;
  }
}

function syncScopedStatuses() {
  const scope = `${ROUTE_LABELS[selectedRoute()]} ${$("#targetMonth").value}`;
  $("#workforceStatus").textContent = `${scope} 기준 인력 DB 불러오기 전`;
  $("#annualLedgerStatus").textContent = `${scope} 기준 연차 누적 DB 불러오기 전`;
  $("#grantCount").textContent = "0건";
}

async function loadScopedData() {
  const workforce = await loadJson("workforce");
  const annualLedger = await loadJson("annual-ledger");
  const substituteGrants = await loadJson("substitute-grants");
  const webEvidence = await loadJson("web-evidence");
  state.workforceRows = workforce?.rows || [];
  state.annualLedgerRows = annualLedger?.rows || [];
  state.substituteGrants = substituteGrants?.rows || [];
  state.webEvidenceRows = webEvidence?.rows || [];
  $("#workforceStatus").textContent = state.workforceRows.length ? `${scopeText()} 기준 ${state.workforceRows.length}건 불러옴 · ${workforce.fileName || "인력 DB"}` : `${scopeText()} 기준 등록된 인력 DB 없음`;
  $("#annualLedgerStatus").textContent = state.annualLedgerRows.length ? `${scopeText()} 기준 ${state.annualLedgerRows.length}건 불러옴 · ${annualLedger.fileName || "연차 누적 DB"}` : `${scopeText()} 기준 등록된 연차 누적 DB 없음`;
  renderSubstituteGrants();
  renderWebEvidenceQueue();
}

function setupDropzone(dropzoneId, inputId, labelId, stateKey) {
  const dropzone = $(`#${dropzoneId}`);
  const input = $(`#${inputId}`);
  const label = $(`#${labelId}`);
  if (!dropzone || !input || !label) return;
  input.addEventListener("change", () => setFileState(input.files?.[0] || null, input, label, stateKey));
  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove("dragging");
    });
  });
  dropzone.addEventListener("drop", (event) => {
    const file = [...(event.dataTransfer?.files || [])].find(isExcelFile);
    if (!file) return showToast("엑셀 파일만 등록할 수 있습니다.");
    setFileState(file, input, label, stateKey);
  });
}

function setFileState(file, input, label, stateKey) {
  state[stateKey] = file;
  label.textContent = file ? file.name : "파일 선택";
  if (input) input.value = "";
  state.analysis = null;
  syncActionState();
}

function renderManagerUploadGrid() {
  const grid = $("#managerUploadGrid");
  grid.innerHTML = "";
  for (const item of ELECTROLAND_MANAGERS) {
    const card = document.createElement("article");
    card.className = "manager-card";
    card.innerHTML = `
      <strong>${item.manager}</strong>
      <small>${item.regionalManager || "-"} · ${item.region || "-"}</small>
      <input type="file" accept=".xlsx,.xls,.xlsb" data-manager="${item.manager}" />
      <p class="field-help" data-manager-label="${item.manager}">파일 없음</p>
    `;
    const input = card.querySelector("input");
    input.addEventListener("change", () => {
      const file = input.files?.[0] || null;
      if (file) state.managerFiles.set(item.manager, file);
      else state.managerFiles.delete(item.manager);
      card.querySelector("[data-manager-label]").textContent = file ? file.name : "파일 없음";
      state.analysis = null;
      syncManagerSummary();
      syncActionState();
    });
    grid.appendChild(card);
  }
}

function switchView(viewName) {
  $$(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === viewName));
  $$(".view").forEach((view) => view.classList.remove("active"));
  $(`#${viewName}View`)?.classList.add("active");
}

async function checkBackend() {
  try {
    const response = await fetch("/api/health", { cache: "no-store" });
    const data = response.ok ? await response.json() : null;
    state.backendAvailable = Boolean(data?.ok);
    $("#backendBadge").textContent = state.backendAvailable ? "D1 연결" : "브라우저 임시 저장";
    $("#backendBadge").className = `badge ${state.backendAvailable ? "ok" : "warn"}`;
  } catch {
    state.backendAvailable = false;
    $("#backendBadge").textContent = "브라우저 임시 저장";
    $("#backendBadge").className = "badge warn";
  }
}

async function runAnalysis({ silent = false } = {}) {
  try {
    if (!state.planFile && !state.attendanceFile && !state.annualFile && !state.managerFiles.size && !state.workforceRows.length) {
      if (!silent) showToast("제모스 근태기록, 매니저 파일, 연차 기록, 인력 DB 중 하나 이상을 넣어주세요.");
      return;
    }
    $("#analyzeButton").disabled = true;
    $("#analyzeButton").textContent = "분석 중...";
    const managerWorkbooks = [];
    for (const [manager, file] of state.managerFiles.entries()) {
      managerWorkbooks.push({ manager, workbook: await readWorkbook(file) });
    }
    state.analysis = buildAnalysis({
      route: selectedRoute(),
      targetMonth: $("#targetMonth").value,
      cutoffDate: $("#cutoffDate").value,
      planWorkbook: await readWorkbook(state.planFile),
      attendanceWorkbook: await readWorkbook(state.attendanceFile),
      annualWorkbook: await readWorkbook(state.annualFile),
      evidenceWorkbook: await readWorkbook(state.evidenceFile),
      closingData: parseClosingWorkbook(await readWorkbook(state.closingFile), $("#targetMonth").value),
      managerWorkbooks,
      workforceRows: state.workforceRows,
      annualLedgerRows: state.annualLedgerRows,
      substituteGrants: state.substituteGrants,
      webEvidenceRows: state.webEvidenceRows,
    });
    renderAnalysis();
    if (!silent) showToast("분석이 완료됐습니다.");
  } catch (error) {
    console.error(error);
    showToast(error.message || "분석 중 오류가 발생했습니다.");
  } finally {
    $("#analyzeButton").disabled = false;
    $("#analyzeButton").textContent = "엑셀 생성 준비";
    syncActionState();
  }
}

function renderAnalysis() {
  const analysis = state.analysis;
  if (!analysis) return;
  $("#peopleCount").textContent = `${analysis.stats.people}명`;
  $("#evidenceCount").textContent = `${analysis.stats.evidence}건`;
  $("#dayoffExcessCount").textContent = `${analysis.stats.dayoffExcess}명`;
  $("#subcompUserCount").textContent = `${analysis.stats.subcompUsers}명`;
  $("#resultDescription").textContent = `${analysis.targetMonth} ${analysis.routeLabel} 기준 · 웹 증빙 확인값과 대체휴무 웹 부여값을 반영`;
  renderPreview();
  renderWebEvidenceQueue();
}

function renderPreview() {
  const analysis = state.analysis;
  if (!analysis) return renderEmptyPreview();
  const query = ($("#searchInput").value || "").trim().toLowerCase();
  const rows = analysis.evidenceRows.filter((row) => {
    if (!query) return true;
    return [row.manager, row.storeName, row.employeeName, row.employeeId, row.date, row.issue].join(" ").toLowerCase().includes(query);
  }).slice(0, 300);
  $("#previewBody").innerHTML = rows.map((row, index) => `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(row.manager)}</td>
      <td>${escapeHtml(row.storeName)}</td>
      <td>${escapeHtml(row.employeeName)}</td>
      <td>${escapeHtml(row.employeeId)}</td>
      <td>${escapeHtml(row.date)}</td>
      <td>${escapeHtml(row.planStatus)}</td>
      <td>${escapeHtml(row.managerStatus)}</td>
      <td>${escapeHtml(row.issue)}</td>
    </tr>
  `).join("") || `<tr><td colspan="9">확인 대상이 없습니다.</td></tr>`;
}

function renderEmptyPreview() {
  $("#previewBody").innerHTML = `<tr><td colspan="9">아직 분석 결과가 없습니다.</td></tr>`;
  $("#webEvidenceBody").innerHTML = `<tr><td colspan="7">분석 후 출근증빙 대상이 표시됩니다.</td></tr>`;
}

async function exportSpecificWorkbook(type) {
  if (!state.analysis) return showToast("먼저 분석을 실행해주세요.");
  const buttonMap = {
    attendance: [$("#exportAttendanceButton"), $("#exportAttendanceButtonTop")],
    annual: [$("#exportAnnualButton"), $("#exportAnnualButtonTop")],
    closing: [$("#exportClosingButton"), $("#exportClosingButtonTop")],
  };
  const labelMap = { attendance: "근태 엑셀", annual: "연차 엑셀", closing: "월 마감 누적 파일" };
  const buildMap = { attendance: buildAttendanceWorkbookFile, annual: buildAnnualWorkbookFile, closing: buildClosingWorkbookFile };
  const buttons = buttonMap[type] || [];
  try {
    buttons.forEach((button) => {
      if (button) {
        button.disabled = true;
        button.textContent = "생성 중...";
      }
    });
    const file = await buildMap[type](state.analysis);
    downloadFile(file);
    showToast(`${labelMap[type]}을 생성했습니다.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || `${labelMap[type]} 생성 중 오류가 발생했습니다.`);
  } finally {
    buttons.forEach((button) => {
      if (button) button.textContent = labelMap[type];
    });
    syncActionState();
  }
}

async function saveMasterFile(type, input, parser, stateKey, statusId, label) {
  const file = input.files?.[0];
  if (!file) return showToast("저장할 엑셀 파일을 먼저 선택해주세요.");
  try {
    const rows = parser(await readWorkbook(file));
    state[stateKey] = rows;
    await saveJson(type, { rows, fileName: file.name, savedAt: new Date().toISOString() });
    $(`#${statusId}`).textContent = `${scopeText()} 기준 ${rows.length}건 저장됨 · ${file.name}`;
    showToast(`${label}을 월별로 저장했습니다.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "저장 중 오류가 발생했습니다.");
  }
}

async function loadMasterData(type, stateKey, statusId, label) {
  try {
    const data = await loadJson(type);
    const rows = data?.rows || [];
    state[stateKey] = rows;
    $(`#${statusId}`).textContent = rows.length ? `${scopeText()} 기준 ${rows.length}건 불러옴 · ${data.fileName || label}` : `${scopeText()} 기준 등록된 ${label} 없음`;
    showToast(`${label}을 불러왔습니다.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "불러오기 중 오류가 발생했습니다.");
  }
}

async function saveSubstituteGrant(event) {
  event.preventDefault();
  const grant = {
    id: crypto.randomUUID(),
    route: $("#substituteRoute").value,
    grantMonth: $("#substituteGrantMonth").value,
    grantedDays: Number($("#substituteGrantDays").value || 0),
    validFrom: $("#substituteValidFrom").value,
    validTo: $("#substituteValidTo").value,
    note: $("#substituteGrantNote").value.trim(),
    createdAt: new Date().toISOString(),
  };
  if (!grant.grantMonth || !grant.grantedDays) return showToast("발생 월과 부여 일수를 입력해주세요.");
  state.substituteGrants = [grant, ...state.substituteGrants];
  await saveJson("substitute-grants", { rows: state.substituteGrants, savedAt: new Date().toISOString() });
  $("#substituteGrantNote").value = "";
  renderSubstituteGrants();
  state.analysis = null;
  syncActionState();
  showToast("대체휴무 부여 내역을 저장했습니다.");
}

async function loadSubstituteGrants() {
  const data = await loadJson("substitute-grants");
  state.substituteGrants = data?.rows || [];
  renderSubstituteGrants();
  showToast("대체휴무 부여 내역을 불러왔습니다.");
}

function renderSubstituteGrants() {
  $("#grantCount").textContent = `${state.substituteGrants.length}건`;
  const rows = state.substituteGrants;
  $("#substituteGrantBody").innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(ROUTE_LABELS[row.route] || row.route)}</td>
      <td>${escapeHtml(row.grantMonth)}</td>
      <td>${escapeHtml(daysText(row.grantedDays))}</td>
      <td>${escapeHtml([row.validFrom, row.validTo].filter(Boolean).join(" ~ "))}</td>
      <td>${escapeHtml(row.note)}</td>
      <td><button class="text-button danger-text" data-delete-grant="${row.id}" type="button">삭제</button></td>
    </tr>
  `).join("") || `<tr><td colspan="6">등록된 대체휴무 부여 내역이 없습니다.</td></tr>`;
  $$("[data-delete-grant]").forEach((button) => {
    button.addEventListener("click", async () => {
      state.substituteGrants = state.substituteGrants.filter((row) => row.id !== button.dataset.deleteGrant);
      await saveJson("substitute-grants", { rows: state.substituteGrants, savedAt: new Date().toISOString() });
      renderSubstituteGrants();
      state.analysis = null;
      syncActionState();
    });
  });
}

function renderWebEvidenceQueue() {
  const rows = state.analysis?.evidenceRows || [];
  if (!rows.length) {
    $("#webEvidenceBody").innerHTML = `<tr><td colspan="7">분석 후 출근증빙 대상이 표시됩니다.</td></tr>`;
    return;
  }
  $("#webEvidenceBody").innerHTML = rows.slice(0, 300).map((row) => {
    const saved = findWebEvidence(row.employeeId, row.date);
    return `
      <tr data-evidence-key="${escapeHtml(`${row.employeeId}|${row.date}`)}">
        <td>${escapeHtml(row.employeeName)}</td>
        <td>${escapeHtml(row.employeeId)}</td>
        <td>${escapeHtml(row.date)}</td>
        <td>${escapeHtml(row.issue)}</td>
        <td>
          <select class="input compact-select" data-evidence-status>
            ${["", "출근확인", "휴무", "연차", "오전반차", "오후반차", "공가", "휴가", "경조", "대체휴무", "보상휴가", "제외"].map((value) => `<option value="${value}" ${saved?.status === value ? "selected" : ""}>${value || "선택"}</option>`).join("")}
          </select>
        </td>
        <td><input class="input compact-input" data-evidence-note value="${escapeHtml(saved?.note || "")}" /></td>
        <td><button class="btn secondary" data-save-evidence type="button">저장</button></td>
      </tr>
    `;
  }).join("");
  $$("[data-save-evidence]").forEach((button) => button.addEventListener("click", () => saveWebEvidenceRow(button.closest("tr"))));
}

async function saveWebEvidenceRow(rowNode) {
  const [employeeId, date] = rowNode.dataset.evidenceKey.split("|");
  const source = state.analysis?.evidenceRows.find((row) => row.employeeId === employeeId && row.date === date);
  const status = rowNode.querySelector("[data-evidence-status]").value;
  const note = rowNode.querySelector("[data-evidence-note]").value.trim();
  if (!status) return showToast("처리값을 선택해주세요.");
  const item = {
    employeeId,
    employeeName: source?.employeeName || "",
    date,
    status,
    note,
    savedAt: new Date().toISOString(),
  };
  state.webEvidenceRows = state.webEvidenceRows.filter((row) => !(row.employeeId === employeeId && row.date === date));
  state.webEvidenceRows.push(item);
  await saveJson("web-evidence", { rows: state.webEvidenceRows, savedAt: new Date().toISOString() });
  showToast("웹 출근증빙을 저장했습니다. 재분석에 바로 반영합니다.");
  await runAnalysis({ silent: true });
}

async function loadWebEvidenceData() {
  const data = await loadJson("web-evidence");
  state.webEvidenceRows = data?.rows || [];
  renderWebEvidenceQueue();
  showToast("웹 출근증빙 저장값을 불러왔습니다.");
}

async function clearWebEvidenceData() {
  state.webEvidenceRows = [];
  await saveJson("web-evidence", { rows: [], savedAt: new Date().toISOString() });
  renderWebEvidenceQueue();
  showToast("현재 경로·월의 웹 출근증빙을 초기화했습니다.");
}

function findWebEvidence(employeeId, date) {
  return state.webEvidenceRows.find((row) => row.employeeId === employeeId && row.date === date);
}

async function saveJson(type, payload) {
  const key = storeKey(type);
  if (state.backendAvailable) {
    const response = await fetch("/api/store", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ key, type, route: selectedRoute(), month: $("#targetMonth").value, payload }),
    });
    if (!response.ok) throw new Error("D1 저장에 실패했습니다.");
    return;
  }
  localStorage.setItem(key, JSON.stringify(payload));
}

async function loadJson(type) {
  const key = storeKey(type);
  if (state.backendAvailable) {
    const response = await fetch(`/api/store?key=${encodeURIComponent(key)}`, { cache: "no-store" });
    if (!response.ok) return null;
    const data = await response.json();
    return data?.payload || null;
  }
  const raw = localStorage.getItem(key);
  return raw ? JSON.parse(raw) : null;
}

function storeKey(type) {
  return `attendance-clean-v4|${type}|${selectedRoute()}|${$("#targetMonth").value}`;
}

function clearManagerFiles() {
  state.managerFiles.clear();
  $$("#managerUploadGrid input").forEach((input) => { input.value = ""; });
  $$("[data-manager-label]").forEach((label) => { label.textContent = "파일 없음"; });
  state.analysis = null;
  syncManagerSummary();
  syncActionState();
}

function resetAnalysisInputs() {
  for (const key of ["planFile", "attendanceFile", "annualFile", "evidenceFile", "closingFile"]) state[key] = null;
  for (const inputId of ["planFile", "attendanceFile", "annualFile", "evidenceFile", "closingFile"]) {
    const input = $(`#${inputId}`);
    if (input) input.value = "";
  }
  $("#planFileName").textContent = "근무계획 파일 선택";
  $("#attendanceFileName").textContent = "근태기록 파일 선택";
  $("#annualFileName").textContent = "연차 승인·반려 양식 선택";
  $("#evidenceFileName").textContent = "출근증빙·휴무확인 O 파일 선택";
  $("#closingFileName").textContent = "전월 연차 마감 파일 선택";
  state.analysis = null;
  renderEmptyPreview();
  syncActionState();
}

function syncManagerSummary() {
  $("#managerSummary").textContent = state.managerFiles.size ? `등록된 매니저 증빙 파일 ${state.managerFiles.size}개` : "등록된 매니저 수정본 없음";
}

function syncActionState() {
  const hasAnalysis = Boolean(state.analysis);
  for (const id of ["exportAttendanceButton", "exportAnnualButton", "exportClosingButton", "exportAttendanceButtonTop", "exportAnnualButtonTop", "exportClosingButtonTop"]) {
    const button = $(`#${id}`);
    if (button) button.disabled = !hasAnalysis;
  }
}

function downloadFile(file) {
  const url = URL.createObjectURL(file);
  const link = document.createElement("a");
  link.href = url;
  link.download = file.name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 3000);
}

function scopeText() {
  return `${ROUTE_LABELS[selectedRoute()] || selectedRoute()} ${$("#targetMonth").value}`;
}

function daysText(value) {
  const num = Number(value || 0);
  return Number.isInteger(num) ? `${num}일` : `${num}일`;
}

function isExcelFile(file) {
  return /\.(xlsx|xls|xlsb)$/i.test(file?.name || "");
}

function toISODate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[char]));
}
