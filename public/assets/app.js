import {
  ELECTROLAND_MANAGERS,
  ROUTE_LABELS,
  buildAnalysis,
  buildAnnualWorkbookFile,
  buildAttendanceWorkbookFile,
  buildClosingWorkbookFile,
  buildWorkbookFile,
  getDayoffAllowance,
  parseClosingWorkbook,
  parseLedgerWorkbook,
  parseMasterWorkbook,
  parseTargetMonth,
  readWorkbook,
} from "./attendance-engine.js?v=clean3";

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
  subcompLedgerRows: [],
  analysis: null,
  snapshots: [],
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
  checkBackend();
  syncRouteHelp();
}

function setDefaultDates() {
  const now = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  $("#targetMonth").value = month;
  $("#cutoffDate").value = toISODate(now);
}

function bindEvents() {
  $$(".tab").forEach((tab) => tab.addEventListener("click", () => switchView(tab.dataset.view)));
  $$('input[name="route"]').forEach((input) => input.addEventListener("change", syncRouteHelp));
  $("#targetMonth").addEventListener("change", () => {
    syncRouteHelp();
    state.analysis = null;
    syncActionState();
  });
  $("#resetButton").addEventListener("click", resetAnalysisInputs);
  $("#openManagerButton").addEventListener("click", () => switchView("manager"));
  $("#clearManagerFiles").addEventListener("click", clearManagerFiles);
  $("#analyzeButton").addEventListener("click", runAnalysis);
  $("#exportButton").addEventListener("click", exportWorkbook);
  $("#exportAttendanceButton")?.addEventListener("click", () => exportSpecificWorkbook("attendance"));
  $("#exportAnnualButton")?.addEventListener("click", () => exportSpecificWorkbook("annual"));
  $("#exportClosingButton")?.addEventListener("click", () => exportSpecificWorkbook("closing"));
  $("#exportAttendanceButtonTop")?.addEventListener("click", () => exportSpecificWorkbook("attendance"));
  $("#exportAnnualButtonTop")?.addEventListener("click", () => exportSpecificWorkbook("annual"));
  $("#exportClosingButtonTop")?.addEventListener("click", () => exportSpecificWorkbook("closing"));
  $("#searchInput").addEventListener("input", renderPreview);
  $("#saveWorkforceButton").addEventListener("click", () => saveMasterFile("workforce", $("#workforceFile"), parseMasterWorkbook, "workforceRows", "workforceStatus"));
  $("#loadWorkforceButton").addEventListener("click", () => loadMasterData("workforce", "workforceRows", "workforceStatus", "인력 DB"));
  $("#saveAnnualLedgerButton").addEventListener("click", () => saveMasterFile("annual-ledger", $("#annualLedgerFile"), parseLedgerWorkbook, "annualLedgerRows", "annualLedgerStatus"));
  $("#loadAnnualLedgerButton").addEventListener("click", () => loadMasterData("annual-ledger", "annualLedgerRows", "annualLedgerStatus", "연차 누적 DB"));
  $("#saveSubcompLedgerButton").addEventListener("click", () => saveMasterFile("subcomp-ledger", $("#subcompLedgerFile"), parseLedgerWorkbook, "subcompLedgerRows", "subcompLedgerStatus"));
  $("#loadSubcompLedgerButton").addEventListener("click", () => loadMasterData("subcomp-ledger", "subcompLedgerRows", "subcompLedgerStatus", "대체·보상 DB"));
  $("#saveSnapshotButton").addEventListener("click", saveSnapshot);
  $("#loadSnapshotsButton").addEventListener("click", loadSnapshots);
}

function setupDropzone(dropzoneId, inputId, labelId, stateKey) {
  const dropzone = $(`#${dropzoneId}`);
  const input = $(`#${inputId}`);
  const label = $(`#${labelId}`);
  input.addEventListener("change", () => {
    const file = input.files?.[0] || null;
    state[stateKey] = file;
    label.textContent = file ? file.name : "파일 선택";
    state.analysis = null;
    syncActionState();
  });
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
    state[stateKey] = file;
    label.textContent = file.name;
    input.value = "";
    state.analysis = null;
    syncActionState();
  });
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

function selectedRoute() {
  return $('input[name="route"]:checked')?.value || "homeplus";
}

function syncRouteHelp() {
  const route = selectedRoute();
  const month = parseTargetMonth($("#targetMonth").value);
  if (route === "electroland" && month.valid) {
    $("#routeHelp").textContent = `전자랜드 기본 휴무는 대상 월 토요일+일요일 ${getDayoffAllowance(route, month)}회입니다.`;
  } else {
    $("#routeHelp").textContent = "홈플러스 기본 휴무는 월 6회입니다.";
  }
}

async function checkBackend() {
  try {
    const response = await fetch("/api/health", { cache: "no-store" });
    const data = response.ok ? await response.json() : null;
    state.backendAvailable = Boolean(data?.ok);
    $("#backendBadge").textContent = state.backendAvailable ? "D1 clean 저장 연결" : "브라우저 임시 저장";
    $("#backendBadge").className = `badge ${state.backendAvailable ? "ok" : "warn"}`;
  } catch {
    state.backendAvailable = false;
    $("#backendBadge").textContent = "브라우저 임시 저장";
    $("#backendBadge").className = "badge warn";
  }
}

async function runAnalysis() {
  try {
    if (!state.planFile && !state.attendanceFile) {
      showToast("근무계획 또는 근태기록 파일을 먼저 넣어주세요.");
      return;
    }
    $("#analyzeButton").disabled = true;
    $("#analyzeButton").textContent = "분석 중...";
    const managerWorkbooks = [];
    for (const [manager, file] of state.managerFiles.entries()) {
      managerWorkbooks.push({ manager, workbook: await readWorkbook(file) });
    }
    const input = {
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
      subcompLedgerRows: state.subcompLedgerRows,
    };
    state.analysis = buildAnalysis(input);
    renderAnalysis();
    showToast("새 기준 분석이 완료됐습니다.");
  } catch (error) {
    console.error(error);
    showToast(error.message || "분석 중 오류가 발생했습니다.");
  } finally {
    $("#analyzeButton").disabled = false;
    $("#analyzeButton").textContent = "근태 분석 실행";
    syncActionState();
  }
}

function renderAnalysis() {
  const analysis = state.analysis;
  if (!analysis) return;
  $("#emptyState").classList.add("hidden");
  $("#resultArea").classList.remove("hidden");
  $("#peopleCount").textContent = analysis.stats.people;
  $("#evidenceCount").textContent = analysis.stats.evidence;
  $("#dayoffExcessCount").textContent = analysis.stats.dayoffExcess;
  $("#subcompUserCount").textContent = analysis.stats.subcompUsers;
  $("#annualIssueCount").textContent = analysis.stats.annualIssues;
  $("#resultDescription").textContent = `${analysis.targetMonth} ${analysis.routeLabel} 기준 · 출근시간 있는 날짜는 경고 제외 · 휴무 초과자는 기본 휴무 초과자만 별도 표시`;
  renderPreview();
}

function renderPreview() {
  const analysis = state.analysis;
  if (!analysis) return;
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

async function exportWorkbook() {
  if (!state.analysis) {
    showToast("먼저 분석을 실행해주세요.");
    return;
  }
  try {
    $("#exportButton").disabled = true;
    $("#exportButton").textContent = "엑셀 생성 중...";
    const file = await buildWorkbookFile(state.analysis);
    downloadFile(file);
    showToast("엑셀 파일을 생성했습니다.");
  } catch (error) {
    console.error(error);
    showToast(error.message || "엑셀 생성 중 오류가 발생했습니다.");
  } finally {
    $("#exportButton").textContent = "엑셀 생성";
    syncActionState();
  }
}

async function exportSpecificWorkbook(type) {
  if (!state.analysis) {
    showToast("먼저 분석을 실행해주세요.");
    return;
  }
  const buttonMap = {
    attendance: $("#exportAttendanceButton"),
    annual: $("#exportAnnualButton"),
    closing: $("#exportClosingButton"),
  };
  const labelMap = {
    attendance: "근태 엑셀",
    annual: "연차 엑셀",
    closing: "월 마감 누적 파일",
  };
  const buildMap = {
    attendance: buildAttendanceWorkbookFile,
    annual: buildAnnualWorkbookFile,
    closing: buildClosingWorkbookFile,
  };
  const button = buttonMap[type];
  try {
    if (button) {
      button.disabled = true;
      button.textContent = "생성 중...";
    }
    const file = await buildMap[type](state.analysis);
    downloadFile(file);
    showToast(`${labelMap[type]}을 생성했습니다.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || `${labelMap[type]} 생성 중 오류가 발생했습니다.`);
  } finally {
    if (button) {
      button.textContent = labelMap[type];
      syncActionState();
    }
  }
}

async function saveMasterFile(type, input, parser, stateKey, statusId) {
  const file = input.files?.[0];
  if (!file) return showToast("저장할 엑셀 파일을 먼저 선택해주세요.");
  try {
    const workbook = await readWorkbook(file);
    const rows = parser(workbook);
    state[stateKey] = rows;
    await saveJson(type, { rows, fileName: file.name, savedAt: new Date().toISOString() });
    $(`#${statusId}`).textContent = `${rows.length}건 저장됨 · ${file.name}`;
    showToast("저장했습니다.");
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
    $(`#${statusId}`).textContent = rows.length ? `${rows.length}건 불러옴 · ${data.fileName || label}` : `등록된 ${label} 없음`;
    showToast(`${label}을 불러왔습니다.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "불러오기 중 오류가 발생했습니다.");
  }
}

async function saveSnapshot() {
  if (!state.analysis) return;
  const name = $("#snapshotName").value.trim() || `${state.analysis.targetMonth} ${ROUTE_LABELS[state.analysis.route]} 중간 저장`;
  const snapshot = {
    id: crypto.randomUUID(),
    name,
    targetMonth: state.analysis.targetMonth,
    route: state.analysis.route,
    cutoffDate: $("#cutoffDate").value,
    savedAt: new Date().toISOString(),
    analysis: state.analysis,
  };
  const current = (await loadJson("snapshots"))?.rows || [];
  current.unshift(snapshot);
  await saveJson("snapshots", { rows: current.slice(0, 20), savedAt: new Date().toISOString() });
  $("#snapshotName").value = "";
  await loadSnapshots();
  showToast("중간 저장했습니다.");
}

async function loadSnapshots() {
  const data = await loadJson("snapshots");
  state.snapshots = data?.rows || [];
  const list = $("#snapshotList");
  list.innerHTML = "";
  if (!state.snapshots.length) {
    list.innerHTML = `<div class="snapshot-item"><span>저장된 중간 결과가 없습니다.</span></div>`;
    return;
  }
  for (const item of state.snapshots) {
    const node = document.createElement("div");
    node.className = "snapshot-item";
    node.innerHTML = `
      <div>
        <strong>${escapeHtml(item.name)}</strong>
        <small>${escapeHtml(item.targetMonth)} · ${escapeHtml(ROUTE_LABELS[item.route] || item.route)} · ${formatDateTime(item.savedAt)}</small>
      </div>
      <div class="button-row">
        <button class="btn secondary" type="button" data-load="${item.id}">불러오기</button>
        <button class="btn danger" type="button" data-delete="${item.id}">삭제</button>
      </div>
    `;
    node.querySelector("[data-load]").addEventListener("click", () => {
      state.analysis = item.analysis;
      $("#targetMonth").value = item.targetMonth;
      $$(`input[name="route"]`).forEach((input) => { input.checked = input.value === item.route; });
      renderAnalysis();
      syncActionState();
      switchView("analysis");
    });
    node.querySelector("[data-delete]").addEventListener("click", async () => {
      const rows = state.snapshots.filter((snapshot) => snapshot.id !== item.id);
      await saveJson("snapshots", { rows, savedAt: new Date().toISOString() });
      await loadSnapshots();
    });
    list.appendChild(node);
  }
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
  const scoped = ["snapshots"].includes(type) ? "global" : `${selectedRoute()}|${$("#targetMonth").value}`;
  return `attendance-clean-v1|${type}|${scoped}`;
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
  for (const inputId of ["planFile", "attendanceFile", "annualFile", "evidenceFile", "closingFile"]) $(`#${inputId}`).value = "";
  $("#planFileName").textContent = "근무계획 파일 선택";
  $("#attendanceFileName").textContent = "근태기록 파일 선택";
  $("#annualFileName").textContent = "연차 승인·반려 양식 선택";
  $("#evidenceFileName").textContent = "출근증빙·휴무확인 O 파일 선택";
  $("#closingFileName").textContent = "전월 연차 마감 파일 선택";
  state.analysis = null;
  $("#resultArea").classList.add("hidden");
  $("#emptyState").classList.remove("hidden");
  syncActionState();
}

function syncManagerSummary() {
  const count = state.managerFiles.size;
  $("#managerSummary").textContent = count ? `등록된 매니저 수정본 ${count}개` : "등록된 매니저 수정본 없음";
}

function syncActionState() {
  $("#exportButton").disabled = !state.analysis;
  $("#exportAttendanceButton").disabled = !state.analysis;
  $("#exportAnnualButton").disabled = !state.analysis;
  $("#exportClosingButton").disabled = !state.analysis;
  $("#exportAttendanceButtonTop").disabled = !state.analysis;
  $("#exportAnnualButtonTop").disabled = !state.analysis;
  $("#exportClosingButtonTop").disabled = !state.analysis;
  $("#saveSnapshotButton").disabled = !state.analysis;
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

function isExcelFile(file) {
  return /\.(xlsx|xls|xlsb)$/i.test(file?.name || "");
}

function toISODate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDateTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")} ${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
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
