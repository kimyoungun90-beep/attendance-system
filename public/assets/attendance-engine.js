import { buildAnnualTemplateWorkbookFile } from "./annual-template-export.js";

export const ROUTE_LABELS = {
  homeplus: "홈플러스",
  electroland: "전자랜드",
};

export const ELECTROLAND_MANAGERS = [
  { regionalManager: "김태권", region: "서울", manager: "서지원" },
  { regionalManager: "신종철", region: "경인", manager: "윤시원" },
  { regionalManager: "신종철", region: "경인", manager: "최강욱" },
  { regionalManager: "오상섭", region: "전라", manager: "강윤민" },
  { regionalManager: "오상섭", region: "전라", manager: "강지훈" },
  { regionalManager: "오상섭", region: "전라", manager: "서재민" },
  { regionalManager: "한건수", region: "충청", manager: "이창우" },
  { regionalManager: "한건수", region: "충청", manager: "임익현" },
  { regionalManager: "강기림", region: "경남", manager: "김윤나" },
  { regionalManager: "강기림", region: "경남", manager: "김준희" },
  { regionalManager: "강기림", region: "경남", manager: "박성건" },
  { regionalManager: "강기림", region: "경남", manager: "원재식" },
  { regionalManager: "정준호", region: "경북", manager: "유지웅" },
  { regionalManager: "정준호", region: "경북", manager: "정해건" },
];

const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];
const ATTENDANCE_REQUIRED = new Set(["근무", "근무A", "근무B", "근무C", "교육", "공백", "미등록", "출근 미입력", "출ㆍ계 미입력", "출계 미입력", "근태 미입력"]);
const HALF_DAY_STATUSES = new Set(["오전반차", "오후반차"]);
const APPROVAL_REQUIRED = new Set(["연차", "오전반차", "오후반차", "출산휴가", "육아휴직"]);
const BASIC_DAYOFF = new Set(["휴무", "휴무(공백)"]);
const MANAGER_ALLOWED = new Set(["휴무", "연차", "오전반차", "오후반차", "출산휴가", "육아휴직", "공가", "휴가", "경조", "대체휴무", "보상휴가"]);
const SUBCOMP_STATUSES = new Set(["대체휴무", "보상휴가"]);
const STATUS_ORDER = ["휴무", "연차", "오전반차", "오후반차", "출산휴가", "육아휴직", "공가", "휴가", "경조", "대체휴무", "보상휴가"];

const STYLE = {
  title: { font: { bold: true, color: { rgb: "FFFFFF" }, sz: 15 }, fill: { fgColor: { rgb: "123B72" } }, alignment: { vertical: "center" } },
  note: { font: { color: { rgb: "3F5874" } }, alignment: { vertical: "center" } },
  header: { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "0F4C88" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("9FB4CE") },
  metric: { font: { bold: true, color: { rgb: "123B72" } }, fill: { fgColor: { rgb: "EAF2FF" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("B8C9E0") },
  normal: { alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("D8E1EF") },
  meta: { alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("D8E1EF") },
  dayoff: { fill: { fgColor: { rgb: "DDEEFF" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("D8E1EF") },
  leave: { fill: { fgColor: { rgb: "E2F4D7" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("D8E1EF") },
  subcomp: { font: { color: { rgb: "FFFFFF" }, bold: true }, fill: { fgColor: { rgb: "1D4F91" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("D8E1EF") },
  issue: { font: { color: { rgb: "D60000" }, bold: true }, fill: { fgColor: { rgb: "FFE4E4" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("F4B4B4") },
  warning: { fill: { fgColor: { rgb: "FFE6B8" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("E8C78E") },
  success: { font: { color: { rgb: "087A2B" }, bold: true }, fill: { fgColor: { rgb: "E8F7EC" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("B9DCC4") },
  input: { fill: { fgColor: { rgb: "FFF2CC" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("E0C879") },
};

export async function readWorkbook(file) {
  if (!file) return null;
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true, cellStyles: true, raw: false });
}

export function parseMasterWorkbook(workbook, options = {}) {
  if (!workbook) return [];

  // 인력 및 매장매칭 통합 파일은 사람 시트와 매장매칭 시트를 반드시 합쳐야 합니다.
  const structured = parseStructuredWorkforceWorkbook(workbook, options);
  if (structured.length) return structured;

  const rows = firstTableRows(workbook);
  const parsed = parseGenericRows(rows).map((row) => normalizeMember({
    route: row.route || row["경로"] || "",
    regionalManager: row.regionalManager || row["지역장"] || "",
    manager: row.manager || row["매니저"] || row["담당매니저"] || "",
    region: row.region || row["지역1"] || row["지역"] || row["권역"] || "",
    subRegion: row.subRegion || row["지역2"] || row["세부지역"] || "",
    storeCode: row.storeCode || row["매장코드"] || row["점포코드"] || "",
    storeName: row.storeName || row["매장명"] || row["점포명"] || row["매장"] || row["점포"] || "",
    portalId: row.portalId || row["포탈사번"] || row["포털사번"] || row["스핀사번"] || row["스핀사원번호"] || "",
    employeeId: row.employeeId || row["제니엘사번"] || row["사번"] || row["직원번호"] || row["사원번호"] || "",
    employeeName: row.employeeName || row["이름"] || row["성명"] || row["직원명"] || "",
    hireDate: row.hireDate || row["제니엘입사일"] || row["입사일"] || "",
    groupHireDate: row.groupHireDate || row["그룹입사일"] || row["고용승계입사일"] || "",
    note: row.note || row["비고"] || "",
  })).filter((row) => row.employeeId || row.employeeName);

  const routeFilter = options?.route || "";
  return routeFilter
    ? parsed.filter((row) => !row.route || normalizeRouteValue(row.route) === routeFilter)
    : parsed;
}

function parseStructuredWorkforceWorkbook(workbook, options = {}) {
  const names = workbook?.SheetNames || [];
  const matchingName = names.find((name) => cleanHeader(name).includes("매장매칭"));
  const definitions = [
    { route: "electroland", keywords: ["랜드인력db", "전자랜드인력db"] },
    { route: "homeplus", keywords: ["홈플인력db", "홈플러스인력db"] },
  ];
  const peopleSheets = definitions.map((definition) => ({
    ...definition,
    sheetName: names.find((name) => definition.keywords.some((keyword) => cleanHeader(name).includes(keyword))) || "",
  })).filter((item) => item.sheetName);

  if (!matchingName || !peopleSheets.length) return [];
  const storeMap = parseWorkforceStoreMap(workbook.Sheets[matchingName]);
  const routeFilter = options?.route || "";
  const rows = [];
  for (const definition of peopleSheets) {
    if (routeFilter && definition.route !== routeFilter) continue;
    rows.push(...parseWorkforcePeopleRows(workbook.Sheets[definition.sheetName], definition.route, storeMap));
  }
  return rows;
}

function parseWorkforceStoreMap(sheet) {
  const rows = sheetRows(sheet);
  const headerIndex = findHeaderRow(rows, [["매장코드"], ["경로"], ["지역장"]]);
  if (headerIndex < 0) return new Map();
  const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
  const columns = {
    storeCode: firstHeaderIndex(headers, ["매장코드", "점포코드"]),
    region1: firstHeaderIndex(headers, ["권역", "지역1"]),
    subRegion: firstHeaderIndex(headers, ["지역2", "지역"]),
    route: firstHeaderIndex(headers, ["경로", "회사"]),
    storeName: firstHeaderIndex(headers, ["매장명", "점포명", "매장"]),
    regionalManager: firstHeaderIndex(headers, ["지역장", "총괄"]),
    manager: firstHeaderIndex(headers, ["매니저", "담당매니저"]),
    closedDate: firstHeaderIndex(headers, ["폐점날짜", "폐점일"]),
    note: firstHeaderIndex(headers, ["비고"]),
  };
  const map = new Map();
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const storeCode = normalizeStoreCodeValue(pick(row, columns.storeCode));
    if (!storeCode) continue;
    map.set(storeCode, {
      route: normalizeRouteValue(pick(row, columns.route)),
      region1: text(pick(row, columns.region1)),
      subRegion: text(pick(row, columns.subRegion)),
      storeName: text(pick(row, columns.storeName)),
      regionalManager: text(pick(row, columns.regionalManager)),
      manager: text(pick(row, columns.manager)),
      closedDate: normalizeDateText(pick(row, columns.closedDate)),
      note: text(pick(row, columns.note)),
    });
  }
  return map;
}

function parseWorkforcePeopleRows(sheet, route, storeMap) {
  const rows = sheetRows(sheet);
  const headerIndex = findHeaderRow(rows, [["사번", "제니엘사번", "사원번호"], ["성명", "이름"], ["매장코드"]]);
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
  const columns = {
    manager: firstHeaderIndex(headers, ["매니저", "담당매니저"]),
    region: firstHeaderIndex(headers, ["지역1", "지역"]),
    storeCode: firstHeaderIndex(headers, ["매장코드", "점포코드"]),
    storeName: firstHeaderIndex(headers, ["매장명", "점포명", "매장"]),
    portalId: firstHeaderIndex(headers, ["포탈사번", "포털사번", "스핀사번", "스핀사원번호", "포탈id", "포털id"]),
    employeeId: firstHeaderIndex(headers, ["제니엘사번", "사번", "사원번호"]),
    employeeName: firstHeaderIndex(headers, ["성명", "이름", "사원명"]),
    hireDate: firstHeaderIndex(headers, ["제니엘입사일", "입사일"]),
    groupHireDate: firstHeaderIndex(headers, ["그룹입사일", "고용승계입사일"]),
    note: firstHeaderIndex(headers, ["비고", "휴퇴사일", "휴/퇴사일"]),
  };
  const output = [];
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const employeeId = normalizeId(pick(row, columns.employeeId));
    const employeeName = text(pick(row, columns.employeeName));
    if (!employeeId || !employeeName) continue;
    const storeCode = normalizeStoreCodeValue(pick(row, columns.storeCode));
    const matched = storeMap.get(storeCode) || {};
    if (matched.route && matched.route !== route) continue;
    output.push(normalizeMember({
      route,
      regionalManager: matched.regionalManager || "",
      manager: matched.manager || text(pick(row, columns.manager)),
      region: text(pick(row, columns.region)) || matched.region1 || "",
      subRegion: matched.subRegion || "",
      storeCode,
      storeName: text(pick(row, columns.storeName)) || matched.storeName || "",
      portalId: text(pick(row, columns.portalId)),
      employeeId,
      employeeName,
      hireDate: normalizeDateText(pick(row, columns.hireDate)),
      groupHireDate: normalizeDateText(pick(row, columns.groupHireDate)),
      note: [text(pick(row, columns.note)), matched.closedDate ? `${matched.closedDate} 폐점` : "", matched.note || ""].filter(Boolean).join(" · "),
    }));
  }
  return output;
}

function normalizeRouteValue(value) {
  const raw = cleanHeader(value);
  if (raw.includes("홈플")) return "homeplus";
  if (raw.includes("전자랜드") || raw === "랜드" || raw.includes("electroland")) return "electroland";
  return text(value);
}

function normalizeStoreCodeValue(value) {
  return text(value).replace(/\.0+$/, "").replace(/\s+/g, "");
}

export function parseLedgerWorkbook(workbook) {
  if (!workbook) return [];
  return parseGenericRows(firstTableRows(workbook));
}

export function parseClosingWorkbook(workbook, targetMonth = "") {
  if (!workbook) return emptyClosingData();
  const month = parseTargetMonth(targetMonth);
  const people = new Map();
  const annualUsage = new Map();
  const latestClosingMonth = latestAnnualClosingSheetMonth(workbook.SheetNames || []);

  for (const sheetName of workbook.SheetNames || []) {
    const compact = sheetName.replace(/\s+/g, "");
    const sheetMonth = extractKoreanMonthFromSheetName(sheetName);
    if (compact.includes("상담사") && compact.includes("재직1년") && compact.includes("월") && !compact.includes("미사용") && !compact.includes("숨기기") && sheetMonth === latestClosingMonth && parseGeneratedClosingPeopleSheet(workbook.Sheets[sheetName], people, compact.includes("미만") ? "재직 1년 미만" : "재직 1년 이상")) {
      continue;
    } else if (compact.includes("재직1년이상") && compact.includes("월") && !compact.includes("미사용") && sheetMonth === latestClosingMonth) {
      parseClosingPeopleSheet(workbook.Sheets[sheetName], {
        people,
        group: "재직 1년 이상",
        headerRow: 7,
        dataStartRow: 10,
        columns: {
          regionalManager: 4,
          manager: 5,
          region: 6,
          subRegion: 7,
          storeCode: 8,
          storeName: 9,
          portalId: 11,
          employeeId: 12,
          employeeName: 13,
          hireDate: 16,
          groupHireDate: 17,
          job: 18,
          currentYearGranted: 46,
          currentYearUsed: 47,
          remainingAnnual: 48,
          annualPeriod: 49,
          firstPromotionDate: 70,
          firstPromotionDone: 71,
          secondPromotionDate: 72,
          secondPromotionDone: 73,
          annualExhaustionDate: 74,
          note: 75,
        },
      });
    } else if (compact.includes("재직1년미만") && compact.includes("월") && !compact.includes("숨기기") && !compact.includes("미사용") && sheetMonth === latestClosingMonth) {
      parseClosingPeopleSheet(workbook.Sheets[sheetName], {
        people,
        group: "재직 1년 미만",
        headerRow: 10,
        dataStartRow: 13,
        columns: {
          regionalManager: 4,
          manager: 5,
          region: 6,
          subRegion: 7,
          storeCode: 8,
          storeName: 9,
          portalId: 11,
          employeeId: 12,
          employeeName: 13,
          hireDate: 16,
          currentYearGranted: 17,
          priorYearUsed: 18,
          currentYearUsed: 19,
          remainingAnnual: 20,
          annualPeriod: 21,
          firstPromotionDate: 24,
          firstPromotionDone: 25,
          secondPromotionDate: 26,
          secondPromotionDone: 27,
          annualExhaustionDate: 28,
          note: 29,
        },
      });
    } else if (/^\d{2,4}년?연차사용/.test(compact) || compact.includes("연차사용")) {
      parseAnnualUsageSheet(workbook.Sheets[sheetName], { annualUsage, sheetName, month });
    }
  }

  return {
    people: [...people.values()].sort(personSort),
    annualUsage: [...annualUsage.values()].sort((a, b) => String(a.year).localeCompare(String(b.year)) || String(a.employeeId).localeCompare(String(b.employeeId))),
    sourceSheetNames: workbook.SheetNames || [],
  };
}

export function buildAnalysis(input) {
  const month = parseTargetMonth(input.targetMonth);
  if (!month.valid) throw new Error("대상 월 형식이 올바르지 않습니다. 예: 2026-06");

  const route = input.route || "homeplus";
  const routeLabel = ROUTE_LABELS[route] || route;
  const planRows = parsePlanWorkbook(input.planWorkbook, month);
  const attendance = parseAttendanceWorkbook(input.attendanceWorkbook, month);
  const annual = parseAnnualWorkbook(input.annualWorkbook, month);
  const evidence = mergeEvidenceData(
    parseEvidenceWorkbook(input.evidenceWorkbook, month),
    parseWebEvidenceRows(input.webEvidenceRows || [], month),
  );
  const manager = parseManagerWorkbooks(input.managerWorkbooks || [], month);
  const closingData = input.closingData || emptyClosingData();
  const workforce = (input.workforceRows || []).map(normalizeMember).filter((row) => row.employeeId || row.employeeName);
  const annualLedger = input.annualLedgerRows || [];
  const substituteGrants = input.substituteGrants || [];
  const people = buildPeople({
    route,
    planRows,
    attendanceRows: attendance.rows,
    annualRows: annual.rows,
    managerRows: manager.rows,
    workforce,
  });
  const dayoffAllowance = getDayoffAllowance(route, month);

  const rows = [];
  const evidenceRows = [];
  const annualRows = [];
  const managerCompareRows = [];

  for (const person of people) {
    const plan = person.plan || {};
    const daily = [];
    const subcompEvents = [];
    const annualEvents = [];
    let basicDayoffCount = 0;

    for (let day = 1; day <= month.daysInMonth; day += 1) {
      const date = `${input.targetMonth}-${String(day).padStart(2, "0")}`;
      const key = `${person.employeeId}|${date}`;
      const rawPlan = text(plan.plans?.[day]);
      const planStatus = normalizeStatus(rawPlan) || "공백";
      const attendanceDay = attendance.map.get(key) || emptyAttendance(date);
      const annualInfo = annual.map.get(key) || null;
      const evidenceInfo = evidence.map.get(key) || null;
      const managerInfo = manager.map.get(key) || null;

      if (managerInfo) {
        managerCompareRows.push(buildManagerCompareRow(person, managerInfo, attendanceDay, date));
      }

      const resolved = resolveDailyStatus({
        date,
        day,
        rawPlan,
        planStatus,
        attendance: attendanceDay,
        annualInfo,
        managerInfo,
        evidenceInfo,
      });
      daily.push(resolved);

      if (BASIC_DAYOFF.has(resolved.managerStatus)) basicDayoffCount += statusAmount(resolved.managerRaw || resolved.managerStatus);
      if (SUBCOMP_STATUSES.has(resolved.managerStatus)) subcompEvents.push({ date, status: resolved.managerStatus, amount: statusAmount(resolved.managerRaw || resolved.managerStatus) });
      if (isAnnualStatus(resolved.managerStatus)) annualEvents.push({ date, status: resolved.managerStatus, source: resolved.managerSource, approved: Boolean(annualInfo?.approved) });

      if (resolved.issue) {
        evidenceRows.push(buildEvidenceRow(person, resolved));
      }
      if (isAnnualStatus(resolved.managerStatus) || annualInfo || (resolved.issueType === "annual_approval")) {
        annualRows.push(buildAnnualRow(person, resolved, annualInfo));
      }
    }

    const dayoffExcess = Math.max(0, roundHalf(basicDayoffCount - dayoffAllowance));
    markDayoffExcessDates(daily, dayoffExcess);
    const grantSummary = resolveSubstituteGrants(person, substituteGrants, route, month);
    const subcompSummary = {
      available: grantSummary.available,
      carryover: 0,
      grants: grantSummary.grants,
    };
    const closingPerson = resolveClosingPerson(person.employeeId, closingData.people);
    const annualLedgerRow = resolveAnnualLedger(person.employeeId, annualLedger, closingPerson);

    rows.push({
      ...person,
      daily,
      dayoffAllowance,
      basicDayoffCount: roundHalf(basicDayoffCount),
      dayoffExcess,
      subcompEvents,
      subcompUsed: roundHalf(subcompEvents.reduce((sum, row) => sum + row.amount, 0)),
      subcompAvailable: subcompSummary.available,
      subcompCarryover: subcompSummary.carryover,
      subcompShortage: Math.max(0, roundHalf(subcompEvents.reduce((sum, row) => sum + row.amount, 0) - subcompSummary.available)),
      substituteGrantNotes: subcompSummary.grants.map((grant) => grant.note).filter(Boolean).join(", "),
      substituteGrantPeriod: subcompSummary.grants.map((grant) => [grant.validFrom, grant.validTo].filter(Boolean).join("~")).filter(Boolean).join(", "),
      annualEvents,
      annualLedger: annualLedgerRow,
      closingPerson,
      monthlyAnnualUsed: roundHalf(annualEvents.reduce((sum, event) => sum + statusAmount(event.status), 0)),
    });
  }

  const dayoffExcessRows = rows.filter((row) => row.dayoffExcess > 0);
  const subcompUserRows = rows.filter((row) => row.subcompUsed > 0 || row.subcompAvailable > 0 || row.dayoffExcess > 0);
  const annualIssueRows = annualRows.filter((row) => row.result !== "정상");

  return {
    route,
    routeLabel,
    targetMonth: input.targetMonth,
    cutoffDate: input.cutoffDate,
    month,
    generatedAt: new Date().toISOString(),
    planRows,
    attendanceRows: attendance.rows,
    annualRows,
    rawAnnualApplications: annual.rows,
    evidenceOverrides: evidence.rows,
    managerCompareRows,
    closingData,
    people: rows.sort(personSort),
    evidenceRows,
    dayoffExcessRows,
    subcompUserRows,
    annualIssueRows,
    stats: {
      people: rows.length,
      evidence: evidenceRows.length,
      dayoffExcess: dayoffExcessRows.length,
      subcompUsers: subcompUserRows.length,
      annualIssues: annualIssueRows.length,
    },
  };
}

export async function buildWorkbookFile(analysis) {
  const workbook = buildWorkbook(analysis);
  const raw = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true });
  const buffer = await applyWorkbookViewSettings(raw, workbook.SheetNames);
  const fileName = `${analysis.targetMonth.replace("-", "년 ")}월_${analysis.routeLabel}_근태관리_clean_v1.xlsx`;
  return new File([buffer], fileName, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

export async function buildAttendanceWorkbookFile(analysis) {
  const workbook = XLSX.utils.book_new();
  appendMainSheet(workbook, "상담사근태_관리자반영", analysis, "manager");
  appendDayoffExcessSheet(workbook, analysis);
  appendSubcompSheet(workbook, analysis);
  appendAnnualUserSheet(workbook, analysis);
  appendAnnualCumulativeSheet(workbook, analysis);
  appendAnnualPromotionSheet(workbook, analysis);
  appendOtherStatusSheet(workbook, analysis);
  appendAttendanceCheckSheet(workbook, analysis);
  appendRawSheet(workbook, "출근 기록", attendanceRawRows(analysis));
  workbook.Props = workbookProps(analysis, "근태 엑셀");
  const raw = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true });
  const buffer = await applyWorkbookViewSettings(raw, workbook.SheetNames);
  return namedWorkbookFile(buffer, `${analysis.targetMonth.replace("-", "년 ")}월_${analysis.routeLabel}_근태_엑셀.xlsx`);
}

export async function buildAnnualWorkbookFile(analysis) {
  return buildAnnualTemplateWorkbookFile(analysis);
}

export async function buildClosingWorkbookFile(analysis) {
  const workbook = XLSX.utils.book_new();
  appendClosingPeopleSheet(workbook, analysis);
  appendClosingAnnualUsageSheet(workbook, analysis);
  appendClosingSubcompSheet(workbook, analysis);
  workbook.Props = workbookProps(analysis, "월 마감 누적 파일");
  const raw = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true });
  const buffer = await applyWorkbookViewSettings(raw, workbook.SheetNames);
  return namedWorkbookFile(buffer, `${analysis.targetMonth.replace("-", "년 ")}월_${analysis.routeLabel}_월마감_누적파일.xlsx`);
}

export function buildWorkbook(analysis) {
  const workbook = XLSX.utils.book_new();
  appendMainSheet(workbook, "상담사근태_관리자반영", analysis, "manager");
  appendManagerCompareSheet(workbook, analysis);
  appendAttendanceCheckSheet(workbook, analysis);
  appendDayoffExcessSheet(workbook, analysis);
  appendSubcompSheet(workbook, analysis);
  appendAnnualUserSheet(workbook, analysis);
  appendAnnualCumulativeSheet(workbook, analysis);
  appendAnnualPromotionSheet(workbook, analysis);
  appendOtherStatusSheet(workbook, analysis);
  appendRawSheet(workbook, "근무계획 원본", planRawRows(analysis));
  appendRawSheet(workbook, "출근 기록", attendanceRawRows(analysis));
  workbook.Props = workbookProps(analysis, "전체 통합 엑셀");
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.CalcPr = { calcMode: "auto", fullCalcOnLoad: true, forceFullCalc: true };
  return workbook;
}

function parsePlanWorkbook(workbook, month) {
  if (!workbook) return [];
  const rows = firstTableRows(workbook, ["근무", "계획", "상담사근태"]);
  const headerIndex = findHeaderRow(rows, [["사번", "직원번호", "사원번호"], ["이름", "성명", "직원명"]]);
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
  const dateHeaderRow = rows[headerIndex - 1] || [];
  const backupDateHeaderRow = rows[headerIndex - 2] || [];
  const dayColumns = headers
    .map((header, index) => ({
      index,
      day: parseDayHeader(dateHeaderRow[index], month) || parseDayHeader(backupDateHeaderRow[index], month) || parseDayHeader(header, month),
    }))
    .filter((row) => row.day);
  const indexes = buildHeaderIndexes(headers);
  const parsed = [];

  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const employeeId = normalizeId(pick(row, indexes.employeeId));
    const employeeName = text(pick(row, indexes.employeeName));
    if (!employeeId && !employeeName) continue;
    const plans = {};
    for (const { index, day } of dayColumns) plans[day] = normalizeStatus(row[index]);
    parsed.push({
      rowNumber: r + 1,
      route: "",
      regionalManager: text(pick(row, indexes.regionalManager)),
      manager: text(pick(row, indexes.manager)),
      region: text(pick(row, indexes.region)),
      subRegion: text(pick(row, indexes.subRegion)),
      storeCode: text(pick(row, indexes.storeCode)),
      storeName: text(pick(row, indexes.storeName)),
      employeeId,
      employeeName,
      hireDate: normalizeDateText(pick(row, indexes.hireDate)),
      groupHireDate: normalizeDateText(pick(row, indexes.groupHireDate)),
      plans,
      raw: row,
    });
  }
  return parsed;
}

function parseAttendanceWorkbook(workbook, month) {
  const map = new Map();
  const parsedRows = [];
  if (!workbook) return { map, rows: parsedRows };
  const sheetNames = attendanceSheetNames(workbook);
  for (const sheetName of sheetNames) {
    const rows = sheetRows(workbook.Sheets[sheetName]);
    const headerIndex = findHeaderRow(rows, [["사번", "직원번호", "사원번호"], ["근무일자", "일자", "날짜"]]);
    if (headerIndex < 0) continue;
    const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
    const indexes = buildHeaderIndexes(headers);
    const dateIndex = firstHeaderIndex(headers, ["근무일자", "일자", "날짜"]);
    if (indexes.employeeId < 0 || dateIndex < 0) continue;
    const clockIndexes = allHeaderIndexes(headers, ["출근시간", "실제출근", "변경출근", "입실", "출근"]).filter((index) => {
      const header = cleanHeader(headers[index]);
      return !header.includes("퇴근") && !header.includes("지점");
    });
    const statusIndex = firstHeaderIndex(headers, ["근태", "근태상태", "근무상태", "실제근태", "상태"]);
    for (let r = headerIndex + 1; r < rows.length; r += 1) {
      const row = rows[r];
      const employeeId = normalizeId(pick(row, indexes.employeeId));
      const employeeName = text(pick(row, indexes.employeeName));
      const date = normalizeDateText(pick(row, dateIndex), month);
      if (!employeeId || !date || !date.startsWith(month.key)) continue;
      const clockIn = firstText(clockIndexes.map((index) => row[index]));
      const status = normalizeStatus(pick(row, statusIndex));
      const item = {
        employeeId,
        employeeName,
        date,
        clockIn,
        hasClockIn: isClockValue(clockIn),
        status,
        raw: row,
        sheetName,
        rowNumber: r + 1,
      };
      parsedRows.push(item);
      const key = `${employeeId}|${date}`;
      const prev = map.get(key);
      if (!prev || (!prev.hasClockIn && item.hasClockIn)) map.set(key, item);
    }
  }
  return { map, rows: parsedRows };
}

function parseAnnualWorkbook(workbook, month) {
  const map = new Map();
  const rowsOut = [];
  if (!workbook) return { map, rows: rowsOut };
  for (const sheetName of workbook.SheetNames) {
    const rows = sheetRows(workbook.Sheets[sheetName]);
    const headerIndex = findHeaderRow(rows, [["사번", "직원번호", "사원번호"], ["휴가", "연차", "일자", "시작일"]]);
    if (headerIndex < 0) continue;
    const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
    const indexes = buildHeaderIndexes(headers);
    const startIndex = firstValidIndex(indexes.date, indexes.startDate);
    const endIndex = firstValidIndex(indexes.endDate, startIndex);
    if (indexes.employeeId < 0 || startIndex < 0) continue;
    for (let r = headerIndex + 1; r < rows.length; r += 1) {
      const row = rows[r];
      const employeeId = normalizeId(pick(row, indexes.employeeId));
      if (!employeeId) continue;
      const start = normalizeDateText(pick(row, startIndex), month);
      const end = normalizeDateText(pick(row, endIndex), month) || start;
      if (!start) continue;
      const typeRaw = text(pick(row, firstValidIndex(indexes.leaveType, indexes.status)));
      const statusRaw = text(pick(row, indexes.approvalStatus));
      const leaveStatus = normalizeLeaveType(typeRaw);
      const approved = isApprovedStatus(statusRaw);
      const rejected = isRejectedStatus(statusRaw);
      const dates = expandDateRange(start, end).filter((date) => date.startsWith(month.key));
      for (const date of dates) {
        const item = {
          employeeId,
          employeeName: text(pick(row, indexes.employeeName)),
          date,
          leaveStatus,
          typeRaw,
          statusRaw,
          approved,
          rejected,
          sheetName,
          rowNumber: r + 1,
        };
        rowsOut.push(item);
        const key = `${employeeId}|${date}`;
        if (!map.has(key) || approved) map.set(key, item);
      }
    }
  }
  return { map, rows: rowsOut };
}

function parseEvidenceWorkbook(workbook, month) {
  const map = new Map();
  const rowsOut = [];
  if (!workbook) return { map, rows: rowsOut };
  for (const sheetName of workbook.SheetNames) {
    if (!sheetName.includes("증빙") && !sheetName.includes("휴무확인") && !sheetName.includes("출근")) continue;
    const rows = sheetRows(workbook.Sheets[sheetName]);
    const headerIndex = findHeaderRow(rows, [["사번", "직원번호", "사원번호"], ["발생일", "일자", "날짜"]]);
    if (headerIndex < 0) continue;
    const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
    const indexes = buildHeaderIndexes(headers);
    const dateIndex = firstValidIndex(indexes.date, indexes.issueDate);
    if (indexes.employeeId < 0 || dateIndex < 0) continue;
    for (let r = headerIndex + 1; r < rows.length; r += 1) {
      const row = rows[r];
      const employeeId = normalizeId(pick(row, indexes.employeeId));
      const date = normalizeDateText(pick(row, dateIndex), month);
      if (!employeeId || !date || !date.startsWith(month.key)) continue;
      const resolved = resolveEvidenceStatus(row, headers);
      if (!resolved.confirmed) continue;
      const item = {
        employeeId,
        date,
        status: resolved.status,
        confirmed: true,
        rawStatus: resolved.rawStatus,
        sheetName,
        rowNumber: r + 1,
      };
      rowsOut.push(item);
      map.set(`${employeeId}|${date}`, item);
    }
  }
  return { map, rows: rowsOut };
}

function parseWebEvidenceRows(rows, month) {
  const map = new Map();
  const rowsOut = [];
  for (const row of rows || []) {
    const employeeId = normalizeId(row.employeeId || row["사번"]);
    const date = normalizeDateText(row.date || row["발생일"] || row["일자"], month);
    const status = normalizeEvidenceOverrideStatus(row.status || row["처리값"] || row["확인값"]);
    if (!employeeId || !date || !date.startsWith(month.key) || !status) continue;
    const item = {
      employeeId,
      employeeName: text(row.employeeName || row["이름"]),
      date,
      status: status === "제외" ? "" : status,
      confirmed: true,
      rawStatus: status,
      note: text(row.note || row["메모"]),
      source: "web",
    };
    rowsOut.push(item);
    map.set(`${employeeId}|${date}`, item);
  }
  return { map, rows: rowsOut };
}

function mergeEvidenceData(fileEvidence, webEvidence) {
  const map = new Map(fileEvidence.map);
  const rows = [...fileEvidence.rows];
  for (const row of webEvidence.rows) {
    map.set(`${row.employeeId}|${row.date}`, row);
    rows.push(row);
  }
  return { map, rows };
}

function parseManagerWorkbooks(managerWorkbooks, month) {
  const map = new Map();
  const rows = [];
  for (const entry of managerWorkbooks) {
    if (!entry?.workbook) continue;
    const parsed = parseManagerWorkbook(entry.manager, entry.workbook, month);
    for (const row of parsed) {
      rows.push(row);
      const key = `${row.employeeId}|${row.date}`;
      if (!map.has(key) || row.applied) map.set(key, row);
    }
  }
  return { map, rows };
}

function parseManagerWorkbook(managerName, workbook, month) {
  const sheetName = chooseSheetName(workbook, ["상담사근태_관리자반영", "상담사근태", "관리자"]);
  if (!sheetName) return [];
  const rows = sheetRows(workbook.Sheets[sheetName]);
  const headerIndex = findHeaderRow(rows, [["사번", "직원번호", "사원번호"], ["이름", "성명", "직원명"]]);
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
  const indexes = buildHeaderIndexes(headers);
  const dateHeaderRow = rows[headerIndex - 1] || [];
  const backupDateHeaderRow = rows[headerIndex - 2] || [];
  const dayColumns = headers
    .map((header, index) => ({
      index,
      day: parseDayHeader(dateHeaderRow[index], month) || parseDayHeader(backupDateHeaderRow[index], month) || parseDayHeader(header, month),
    }))
    .filter((row) => row.day);
  const parsed = [];
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const employeeId = normalizeId(pick(row, indexes.employeeId));
    if (!employeeId) continue;
    for (const { index, day } of dayColumns) {
      const raw = text(row[index]);
      if (!raw) continue;
      const normalized = normalizeStatus(raw);
      const date = `${month.key}-${String(day).padStart(2, "0")}`;
      const attendanceLike = isAttendanceLike(raw);
      const applied = MANAGER_ALLOWED.has(normalized) && !attendanceLike;
      if (!applied && !attendanceLike) continue;
      parsed.push({
        regionalManager: text(pick(row, indexes.regionalManager)),
        manager: managerName || text(pick(row, indexes.manager)),
        region: text(pick(row, indexes.region)),
        subRegion: text(pick(row, indexes.subRegion)),
        storeCode: text(pick(row, indexes.storeCode)),
        storeName: text(pick(row, indexes.storeName)),
        employeeId,
        employeeName: text(pick(row, indexes.employeeName)),
        date,
        raw,
        status: applied ? normalized : "",
        applied,
        skippedReason: applied ? "" : "출근·근무·시간값은 증빙 확인 전 자동 반영 제외",
        sheetName,
        rowNumber: r + 1,
      });
    }
  }
  return parsed;
}

function parseGeneratedClosingPeopleSheet(sheet, people, group) {
  const rows = sheetRows(sheet);
  const headerIndex = findHeaderRow(rows, [["제니엘사번", "제니엘 사번", "사번"], ["사원명", "이름"], ["잔여연차"]]);
  if (headerIndex < 0) return false;
  const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
  const index = (aliases) => firstHeaderIndex(headers, aliases);
  const indexes = {
    regionalManager: index(["지역장"]), manager: index(["매니저"]), region: index(["지역1", "지역"]), subRegion: index(["지역2"]),
    storeCode: index(["근무처코드", "매장코드"]), storeName: index(["근무처명", "매장명"]), portalId: index(["스핀사원번호", "포탈사번"]),
    employeeId: index(["제니엘사번", "사번"]), employeeName: index(["사원명", "이름"]), hireDate: index(["제니엘입사일"]), groupHireDate: index(["고용승계입사일"]),
    job: index(["직무"]), granted: index(["연차", "발생"]), currentUsed: index(["년사용", "현재연도사용"]), priorUsed: index(["전년사용"]), remaining: index(["잔여연차"]),
    period: index(["사용기간"]), firstDate: index(["1차촉진"]), firstDone: index(["1차촉진유무", "촉진유무"]), secondDate: index(["2차촉진"]),
    exhaustion: index(["연차소진일"]), note: index(["비고"]), employmentGroup: index(["입사구분", "구분"]),
  };
  // 같은 이름의 '촉진 유무' 열이 두 개인 경우 위치로 1차/2차를 구분합니다.
  const promotionDoneIndexes = allHeaderIndexes(headers, ["촉진유무"]);
  if (promotionDoneIndexes.length) indexes.firstDone = promotionDoneIndexes[0];
  indexes.secondDone = promotionDoneIndexes.length > 1 ? promotionDoneIndexes[1] : -1;
  if (indexes.employeeId < 0 || indexes.employeeName < 0 || indexes.remaining < 0) return false;
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const employeeId = normalizeId(pick(row, indexes.employeeId));
    const employeeName = text(pick(row, indexes.employeeName));
    if (!employeeId || !employeeName || employeeName === "필터용") continue;
    people.set(employeeId, {
      employmentGroup: text(pick(row, indexes.employmentGroup)) || group,
      regionalManager: text(pick(row, indexes.regionalManager)), manager: text(pick(row, indexes.manager)), region: text(pick(row, indexes.region)), subRegion: text(pick(row, indexes.subRegion)),
      storeCode: text(pick(row, indexes.storeCode)), storeName: text(pick(row, indexes.storeName)), portalId: text(pick(row, indexes.portalId)), employeeId, employeeName,
      hireDate: normalizeDateText(pick(row, indexes.hireDate)), groupHireDate: normalizeDateText(pick(row, indexes.groupHireDate)), job: text(pick(row, indexes.job)),
      currentYearGranted: numberOrBlank(pick(row, indexes.granted)), currentYearUsed: numberOrBlank(pick(row, indexes.currentUsed)), priorYearUsed: numberOrBlank(pick(row, indexes.priorUsed)),
      remainingAnnual: numberOrBlank(pick(row, indexes.remaining)), annualPeriod: text(pick(row, indexes.period)), firstPromotionDate: normalizeDateText(pick(row, indexes.firstDate)),
      firstPromotionDone: text(pick(row, indexes.firstDone)), secondPromotionDate: normalizeDateText(pick(row, indexes.secondDate)), secondPromotionDone: text(pick(row, indexes.secondDone)),
      annualExhaustionDate: normalizeDateText(pick(row, indexes.exhaustion)), note: text(pick(row, indexes.note)),
    });
  }
  return true;
}

function parseClosingPeopleSheet(sheet, options) {
  const rows = sheetRows(sheet);
  const cols = options.columns;
  for (let r = options.dataStartRow - 1; r < rows.length; r += 1) {
    const row = rows[r];
    const employeeId = normalizeId(row[(cols.employeeId || 0) - 1]);
    const employeeName = text(row[(cols.employeeName || 0) - 1]);
    if (!employeeId || !employeeName || employeeName === "필터용") continue;
    options.people.set(employeeId, {
      employmentGroup: options.group,
      regionalManager: text(row[(cols.regionalManager || 0) - 1]),
      manager: text(row[(cols.manager || 0) - 1]),
      region: text(row[(cols.region || 0) - 1]),
      subRegion: text(row[(cols.subRegion || 0) - 1]),
      storeCode: text(row[(cols.storeCode || 0) - 1]),
      storeName: text(row[(cols.storeName || 0) - 1]),
      portalId: text(row[(cols.portalId || 0) - 1]),
      employeeId,
      employeeName,
      hireDate: normalizeDateText(row[(cols.hireDate || 0) - 1]),
      groupHireDate: normalizeDateText(row[(cols.groupHireDate || 0) - 1]),
      job: text(row[(cols.job || 0) - 1]),
      currentYearGranted: numberOrBlank(row[(cols.currentYearGranted || 0) - 1]),
      currentYearUsed: numberOrBlank(row[(cols.currentYearUsed || 0) - 1]),
      priorYearUsed: numberOrBlank(row[(cols.priorYearUsed || 0) - 1]),
      remainingAnnual: numberOrBlank(row[(cols.remainingAnnual || 0) - 1]),
      annualPeriod: text(row[(cols.annualPeriod || 0) - 1]),
      firstPromotionDate: normalizeDateText(row[(cols.firstPromotionDate || 0) - 1]),
      firstPromotionDone: text(row[(cols.firstPromotionDone || 0) - 1]),
      secondPromotionDate: normalizeDateText(row[(cols.secondPromotionDate || 0) - 1]),
      secondPromotionDone: text(row[(cols.secondPromotionDone || 0) - 1]),
      annualExhaustionDate: normalizeDateText(row[(cols.annualExhaustionDate || 0) - 1]),
      resignDate: normalizeDateText(row[(cols.resignDate || 0) - 1]),
      note: text(row[(cols.note || 0) - 1]),
    });
  }
}

function parseAnnualUsageSheet(sheet, { annualUsage, sheetName }) {
  const rows = sheetRows(sheet);
  const headerIndex = findHeaderRow(rows, [["제니엘사번", "사번"], ["이름"], ["1월"]]);
  if (headerIndex < 0) return;
  const headers = rows[headerIndex].map((cell) => cleanHeader(cell));
  const portalIndex = firstHeaderIndex(headers, ["포탈사번", "스핀사원번호", "사번"]);
  const employeeIndex = firstHeaderIndex(headers, ["제니엘사번", "제니엘 사번"]);
  const nameIndex = firstHeaderIndex(headers, ["이름", "사원명"]);
  const totalIndex = firstHeaderIndex(headers, ["총합계"]);
  const hireIndex = firstHeaderIndex(headers, ["제니엘입사일"]);
  const groupHireIndex = firstHeaderIndex(headers, ["고용승계입사일"]);
  const resignIndex = firstHeaderIndex(headers, ["퇴사일"]);
  const noteIndex = firstHeaderIndex(headers, ["비고"]);
  const yearMatch = String(sheetName || "").match(/(\d{2,4})년/);
  const parsedYear = yearMatch ? Number(yearMatch[1].length === 2 ? `20${yearMatch[1]}` : yearMatch[1]) : "";
  const monthIndexes = {};
  for (let m = 1; m <= 12; m += 1) {
    const index = firstHeaderIndex(headers, [`${m}월`]);
    if (index >= 0) monthIndexes[`${m}월`] = index;
  }
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const employeeId = normalizeId(row[employeeIndex]);
    const employeeName = text(row[nameIndex]);
    if (!employeeId || !employeeName) continue;
    const months = {};
    for (let m = 1; m <= 12; m += 1) months[`${m}월`] = numberValue(row[monthIndexes[`${m}월`]]);
    annualUsage.set(`${parsedYear}|${employeeId}`, {
      year: parsedYear,
      portalId: text(row[portalIndex]),
      employeeId,
      employeeName,
      months,
      total: numberValue(row[totalIndex]),
      hireDate: normalizeDateText(row[hireIndex]),
      groupHireDate: normalizeDateText(row[groupHireIndex]),
      resignDate: normalizeDateText(row[resignIndex]),
      note: text(row[noteIndex]),
    });
  }
}

function latestAnnualClosingSheetMonth(sheetNames) {
  const months = sheetNames
    .filter((name) => {
      const compact = String(name || "").replace(/\s+/g, "");
      return compact.includes("상담사") && compact.includes("재직1년") && compact.includes("월") && !compact.includes("미사용") && !compact.includes("숨기기");
    })
    .map(extractKoreanMonthFromSheetName)
    .filter(Boolean);
  return months.length ? Math.max(...months) : 0;
}

function extractKoreanMonthFromSheetName(sheetName) {
  const matches = [...String(sheetName || "").matchAll(/(\d{1,2})월/g)].map((match) => Number(match[1])).filter((value) => value >= 1 && value <= 12);
  return matches.length ? matches[matches.length - 1] : 0;
}

function buildPeople({ route, planRows, attendanceRows, annualRows, managerRows, workforce }) {
  const people = new Map();
  const workforceById = new Map();
  for (const member of workforce) {
    const id = normalizeId(member.employeeId);
    if (id && !workforceById.has(id)) workforceById.set(id, normalizeMember(member));
  }

  for (const plan of planRows) {
    const id = normalizeId(plan.employeeId);
    if (!id) continue;
    const member = workforceById.get(id) || normalizeMember(plan);
    people.set(id, normalizePerson(route, { ...member, ...blankToFallback(plan, member), plan }));
  }
  for (const member of workforce) {
    const id = normalizeId(member.employeeId);
    if (!id || people.has(id)) continue;
    people.set(id, normalizePerson(route, { ...member, plan: null }));
  }
  for (const row of managerRows || []) addPersonFromSource(people, route, workforceById, row);
  for (const row of attendanceRows) {
    addPersonFromSource(people, route, workforceById, row);
  }
  for (const row of annualRows || []) addPersonFromSource(people, route, workforceById, row);
  return [...people.values()];
}

function addPersonFromSource(people, route, workforceById, row) {
  const id = normalizeId(row.employeeId);
  if (!id || people.has(id)) return;
  const member = workforceById.get(id) || normalizeMember(row);
  people.set(id, normalizePerson(route, {
    ...member,
    ...blankToFallback(normalizeMember(row), member),
    employeeId: id,
    employeeName: text(row.employeeName) || member.employeeName,
    plan: null,
  }));
}

function resolveDailyStatus({ date, day, rawPlan, planStatus, attendance, annualInfo, managerInfo, evidenceInfo }) {
  const hasClockIn = Boolean(attendance?.hasClockIn);
  const clockText = hasClockIn ? normalizeClock(attendance.clockIn) : "";
  const approvedLeave = annualInfo?.approved ? annualInfo.leaveStatus : "";
  let baseStatus = hasClockIn ? clockText : (approvedLeave || planStatus || "공백");
  let baseRaw = hasClockIn ? clockText : (approvedLeave || rawPlan || planStatus || "공백");
  let baseSource = hasClockIn ? "clock" : approvedLeave ? "annual_approved" : "plan";

  let managerStatus = baseStatus;
  let managerRaw = baseRaw;
  let managerSource = baseSource;
  if (!hasClockIn && managerInfo?.applied) {
    managerStatus = managerInfo.status;
    managerRaw = managerInfo.raw;
    managerSource = "manager";
  }

  const evidenceConfirmed = Boolean(evidenceInfo?.confirmed);
  if (!hasClockIn && evidenceConfirmed) {
    if (evidenceInfo.status) {
      managerStatus = evidenceInfo.status === "출근확인" ? "출근확인" : evidenceInfo.status;
      managerRaw = evidenceInfo.rawStatus || evidenceInfo.status;
    }
    managerSource = "evidence";
  }

  const issue = classifyIssue({
    hasClockIn,
    planStatus,
    managerStatus,
    managerSource,
    managerInfo,
    annualInfo,
    evidenceConfirmed,
  });

  return {
    date,
    day,
    weekday: WEEKDAYS[new Date(date).getDay()],
    rawPlan,
    planStatus,
    clockIn: clockText,
    hasClockIn,
    attendanceStatus: attendance?.status || "",
    approvedLeave,
    baseStatus,
    baseRaw,
    baseSource,
    managerStatus,
    managerRaw,
    managerSource,
    managerApplied: Boolean(managerInfo?.applied),
    managerSkippedReason: managerInfo?.skippedReason || "",
    evidenceConfirmed,
    issue: issue?.label || "",
    issueType: issue?.type || "",
    dayoffExcessDate: false,
  };
}

function classifyIssue({ hasClockIn, planStatus, managerStatus, managerSource, managerInfo, annualInfo, evidenceConfirmed }) {
  if (hasClockIn) return null;
  if (evidenceConfirmed) return null;
  if (managerInfo && !managerInfo.applied && managerInfo.skippedReason) return { type: "manual_clock", label: "수기 출근 확인 필요" };
  if (HALF_DAY_STATUSES.has(managerStatus)) return { type: "halfday_clock", label: "반차 출근증빙 필요" };
  if (APPROVAL_REQUIRED.has(managerStatus) && !annualInfo?.approved && managerSource !== "evidence") {
    return { type: "annual_approval", label: "승인 확인 필요한 휴가/연차" };
  }
  if (isMissingStatus(managerStatus) || ATTENDANCE_REQUIRED.has(managerStatus)) {
    return { type: "missing_clock", label: missingLabel(managerStatus || planStatus) };
  }
  return null;
}

function buildEvidenceRow(person, daily) {
  return {
    regionalManager: person.regionalManager,
    manager: person.manager,
    region: person.region,
    storeName: person.storeName,
    employeeName: person.employeeName,
    employeeId: person.employeeId,
    date: daily.date,
    weekday: daily.weekday,
    planStatus: daily.planStatus,
    managerStatus: daily.managerStatus,
    issue: daily.issue,
    note: "",
  };
}

function buildAnnualRow(person, daily, annualInfo) {
  const normal = Boolean(annualInfo?.approved) && !daily.issue;
  return {
    regionalManager: person.regionalManager,
    manager: person.manager,
    region: person.region,
    storeName: person.storeName,
    employeeName: person.employeeName,
    employeeId: person.employeeId,
    date: daily.date,
    planStatus: daily.planStatus,
    finalStatus: daily.managerStatus,
    applicationStatus: annualInfo?.statusRaw || "신청 없음",
    applicationType: annualInfo?.typeRaw || "",
    result: normal ? "정상" : annualInfo?.approved ? "출근증빙 필요" : "승인 양식 없음",
  };
}

function buildManagerCompareRow(person, managerInfo, attendance, date) {
  return {
    manager: managerInfo.manager,
    regionalManager: person.regionalManager,
    region: person.region,
    storeName: person.storeName,
    employeeName: person.employeeName || managerInfo.employeeName,
    employeeId: person.employeeId,
    date,
    raw: managerInfo.raw,
    appliedValue: managerInfo.applied ? managerInfo.status : "",
    result: managerInfo.applied ? "반영" : "제외",
    reason: managerInfo.applied ? "관리자반영 시트에 반영" : managerInfo.skippedReason,
    hasClockIn: Boolean(attendance?.hasClockIn),
  };
}

function markDayoffExcessDates(daily, dayoffExcess) {
  let remaining = Number(dayoffExcess || 0);
  if (remaining <= 0) return;
  const dayoffDays = daily.filter((row) => BASIC_DAYOFF.has(row.managerStatus)).reverse();
  for (const row of dayoffDays) {
    if (remaining <= 0) break;
    row.dayoffExcessDate = true;
    remaining = roundHalf(remaining - statusAmount(row.managerRaw || row.managerStatus));
  }
}

function appendMainSheet(workbook, name, analysis, mode) {
  const year = analysis.month.year;
  const monthNo = analysis.month.monthNo;
  const daysInMonth = analysis.month.daysInMonth;
  const firstDayCol0 = 14; // O열
  const summaryHeaders = [
    "총 등록 현황", "출근 등록 횟수", "휴무 가능 개수", "휴무 등록 개수", "휴무 초과 개수",
    "대체+보상 가능 개수", "대체+보상 등록 개수", "대체+보상 초과 개수", "출근 수정", "교육",
    "반차", "연차", "공가", "무급휴무", "경조사", "총 일수", "연차 미신청",
    "연차 신청(승인)", "출근 증빙", "수정 완료", "비고",
  ];
  const summaryStartCol0 = firstDayCol0 + daysInMonth;
  const summaryEndCol0 = summaryStartCol0 + summaryHeaders.length - 1;
  const totalCols = summaryEndCol0 + 1;
  const rowCount = 6 + analysis.people.length;
  const matrix = Array.from({ length: Math.max(7, rowCount) }, () => Array(totalCols).fill(""));

  matrix[1][1] = `■${year}년 ${monthNo}월 출퇴근 현황`;
  matrix[2][1] = '- 휴무 → "출근" or "연차" 수정 시, 출근 증빙자료 및 제모스 해당일자 연차신청 필수';
  matrix[3][1] = '- "연차미신청"으로 수정된 날짜는 모두 제모스에 연차신청 가이드 바랍니다.';
  matrix[4][13] = `${monthNo}월1일 \n근무계획\n일치확인`;

  const metaHeaders = ["지역장", "매니저", "지역1", "지역2", "매장코드", "매장명", "포탈사번", "사번", "성명", "제니엘입사일", "그룹입사일", "휴/퇴사일"];
  metaHeaders.forEach((value, index) => { matrix[5][index + 1] = value; });

  for (let day = 1; day <= daysInMonth; day += 1) {
    const col = firstDayCol0 + day - 1;
    const date = new Date(year, monthNo - 1, day);
    matrix[4][col] = date;
    matrix[5][col] = WEEKDAYS[date.getDay()];
  }
  summaryHeaders.forEach((value, index) => { matrix[4][summaryStartCol0 + index] = value; });
  matrix[5][summaryStartCol0 + 2] = "(기본 휴무)";
  matrix[5][summaryStartCol0 + 3] = "(최종 표시)";
  matrix[5][summaryStartCol0 + 4] = "(초과)";
  matrix[5][summaryStartCol0 + 5] = "(이월+당월)";
  matrix[5][summaryStartCol0 + 6] = "(최종 사용)";
  matrix[5][summaryStartCol0 + 7] = "(초과)";

  analysis.people.forEach((person, index) => {
    const r = 6 + index;
    const closing = person.closingPerson || {};
    const daily = person.daily || [];
    const displayed = daily.map((item) => mode === "manager" ? item.managerStatus : item.baseStatus);
    const planValues = daily.map((item) => item.planStatus);
    const registeredCount = displayed.filter((value) => value && !String(value).includes("미입력") && value !== "미등록").length;
    const clockCount = daily.filter((item) => item.hasClockIn).length;
    const dayoffCount = displayed.reduce((sum, value) => sum + (BASIC_DAYOFF.has(value) ? statusAmount(value) : 0), 0);
    const combinedAvailable = roundHalf(Number(person.subcompAvailable || 0) + Number(person.subcompCarryover || 0));
    const combinedUsed = roundHalf(Number(person.subcompUsed || 0));
    const combinedShortage = roundHalf(Math.max(0, combinedUsed - combinedAvailable));
    const clockCorrection = daily.filter((item) => item.hasClockIn && ["휴무", "휴무(공백)", "연차", "오전반차", "오후반차", "공가", "휴가", "경조", "출산휴가", "육아휴직", "대체휴무", "보상휴가"].includes(item.planStatus)).length;
    const annualRows = (analysis.annualRows || []).filter((item) => normalizeId(item.employeeId) === normalizeId(person.employeeId));
    const approvedAnnual = annualRows.filter((item) => item.applicationStatus && !String(item.applicationStatus).includes("반려") && !String(item.applicationStatus).includes("취소") && !String(item.applicationStatus).includes("철회") && item.result === "정상").reduce((sum, item) => sum + statusAmount(item.finalStatus), 0);
    const missingAnnual = annualRows.filter((item) => item.result === "승인 양식 없음").reduce((sum, item) => sum + statusAmount(item.finalStatus || item.planStatus), 0);
    const evidenceNeeded = daily.some((item) => Boolean(item.issue));
    const evidenceCompleted = daily.some((item) => Boolean(item.evidenceConfirmed));
    const noteParts = [];
    if (person.note) noteParts.push(person.note);
    if (person.dayoffExcess > 0) noteParts.push(`기본 휴무 ${compactNumber(person.dayoffExcess)}일 초과`);
    if (person.subcompShortage > 0) noteParts.push(`대체+보상 ${compactNumber(person.subcompShortage)}일 부족`);
    for (const item of daily) if (item.issue) noteParts.push(`${monthNo}/${item.day} ${item.issue}`);

    const values = [
      person.regionalManager || "", person.manager || "", person.region || "", person.subRegion || "",
      person.storeCode || "", person.storeName || "", closing.portalId || person.portalId || "", person.employeeId || "",
      person.employeeName || "", outputDateValue(person.hireDate || closing.hireDate || ""), outputDateValue(person.groupHireDate || closing.groupHireDate || ""),
      outputDateValue(closing.resignDate || ""),
    ];
    values.forEach((value, c) => { matrix[r][c + 1] = value; });
    matrix[r][13] = daily[0]?.planStatus === "공백" ? "" : daily[0]?.planStatus || "";
    displayed.forEach((value, dayIndex) => { matrix[r][firstDayCol0 + dayIndex] = value || ""; });

    const summaryValues = [
      registeredCount, clockCount, person.dayoffAllowance || 0, roundHalf(dayoffCount), roundHalf(person.dayoffExcess || 0),
      combinedAvailable, combinedUsed, combinedShortage, clockCorrection,
      planValues.filter((value) => value === "교육").length,
      planValues.filter((value) => HALF_DAY_STATUSES.has(value)).length,
      planValues.filter((value) => value === "연차").length,
      planValues.filter((value) => value === "공가").length,
      planValues.filter((value) => value === "무급휴가").length,
      planValues.filter((value) => value === "경조").length,
      daysInMonth, roundHalf(missingAnnual), roundHalf(approvedAnnual), evidenceCompleted ? "완료" : evidenceNeeded ? "필요" : "", "",
      [...new Set(noteParts)].join(" · "),
    ];
    summaryValues.forEach((value, c) => { matrix[r][summaryStartCol0 + c] = value; });
  });

  const sheet = XLSX.utils.aoa_to_sheet(matrix, { cellDates: true });
  sheet["!merges"] = [
    { s: { r: 1, c: 1 }, e: { r: 1, c: 9 } },
    { s: { r: 2, c: 1 }, e: { r: 2, c: 9 } },
    { s: { r: 3, c: 1 }, e: { r: 3, c: 9 } },
    { s: { r: 4, c: 1 }, e: { r: 4, c: 12 } },
    { s: { r: 4, c: 13 }, e: { r: 5, c: 13 } },
  ];
  sheet["!freeze"] = { xSplit: 9, ySplit: 6, topLeftCell: "J7", activePane: "bottomRight", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];
  sheet["!autofilter"] = { ref: `B6:M${Math.max(7, rowCount)}` };
  applyLegacyAttendanceDesign(sheet, analysis, summaryStartCol0, summaryEndCol0, totalCols, rowCount);
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

function applyLegacyAttendanceDesign(sheet, analysis, summaryStartCol0, summaryEndCol0, totalCols, rowCount) {
  const year = analysis.month.year;
  const monthNo = analysis.month.monthNo;
  const daysInMonth = analysis.month.daysInMonth;
  const firstDayCol0 = 14;
  const lastDayCol0 = firstDayCol0 + daysInMonth - 1;
  const lastRow0 = Math.max(6, rowCount - 1);
  const cols = Array.from({ length: totalCols }, () => ({ wch: 10 }));
  cols[0] = { wch: 2 };
  [1, 2, 3, 4].forEach((i) => { cols[i] = { wch: 11 }; });
  cols[5] = { wch: 12 }; cols[6] = { wch: 16 }; cols[7] = { wch: 15 }; cols[8] = { wch: 12 }; cols[9] = { wch: 11 };
  [10, 11, 12].forEach((i) => { cols[i] = { wch: 12 }; });
  cols[13] = { wch: 13 };
  for (let c = firstDayCol0; c <= lastDayCol0; c += 1) cols[c] = { wch: 10.5 };
  for (let c = summaryStartCol0; c <= summaryEndCol0; c += 1) cols[c] = { wch: 12 };
  cols[summaryStartCol0 + 2] = { wch: 13.8 }; cols[summaryStartCol0 + 3] = { wch: 13.8 }; cols[summaryStartCol0 + 4] = { wch: 13.8 };
  cols[summaryStartCol0 + 5] = { wch: 10.9 }; cols[summaryStartCol0 + 6] = { wch: 10.9 }; cols[summaryStartCol0 + 7] = { wch: 10.9 };
  cols[summaryEndCol0] = { wch: 46 };
  sheet["!cols"] = cols;
  sheet["!rows"] = Array.from({ length: Math.max(7, rowCount) }, (_, r) => ({ hpt: r === 1 || r === 2 || r === 3 ? 27 : r === 4 ? 33 : r === 5 ? 18 : 20 }));

  const white = { fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } }, font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } }, alignment: { horizontal: "left", vertical: "center", wrapText: true } };
  const title = { ...white, font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FF000000" } } };
  const note = { ...white, font: { name: "맑은 고딕", sz: 15, bold: true, color: { rgb: "FF0000FF" } } };
  const navy = { fill: { patternType: "solid", fgColor: { rgb: "FF203764" } }, font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFFFFFFF") };
  const dateHeader = { fill: { patternType: "solid", fgColor: { rgb: "FFD9E1F2" } }, font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFB4C7DC") };
  const meta = { font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFD9E1E8") };
  const planCheck = { fill: { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } }, font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF000000" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFD9E1E8") };

  styleRange(sheet, 1, 1, 1, 9, title);
  styleRange(sheet, 2, 1, 3, 9, note);
  styleRange(sheet, 4, 1, 5, 12, navy);
  styleRange(sheet, 4, 13, lastRow0, 13, planCheck);
  styleRange(sheet, 4, firstDayCol0, 4, lastDayCol0, dateHeader);
  styleRange(sheet, 5, firstDayCol0, 5, lastDayCol0, navy);
  styleRange(sheet, 4, summaryStartCol0, 5, summaryEndCol0, navy);
  styleRange(sheet, 6, 1, lastRow0, 12, meta);
  styleRange(sheet, 6, summaryStartCol0, lastRow0, summaryEndCol0, meta);

  const dayoffHead = { ...navy, fill: { patternType: "solid", fgColor: { rgb: "FF5B9BD5" } } };
  const extraHead = { ...navy, fill: { patternType: "solid", fgColor: { rgb: "FF8064A2" } } };
  styleRange(sheet, 4, summaryStartCol0 + 2, 4, summaryStartCol0 + 4, dayoffHead);
  styleRange(sheet, 4, summaryStartCol0 + 5, 4, summaryStartCol0 + 7, extraHead);

  for (let day = 1; day <= daysInMonth; day += 1) {
    const c = firstDayCol0 + day - 1;
    const ref = XLSX.utils.encode_cell({ r: 4, c });
    if (sheet[ref]) sheet[ref].z = "yyyymmdd";
    const date = new Date(year, monthNo - 1, day);
    if (date.getDay() === 0 || date.getDay() === 6) {
      for (const r of [4, 5]) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (sheet[addr]) sheet[addr].s = { ...(sheet[addr].s || {}), font: { ...((sheet[addr].s || {}).font || {}), color: { rgb: "FFFF0000" }, bold: r === 5 } };
      }
    }
  }

  analysis.people.forEach((person, index) => {
    const r = 6 + index;
    const altFill = index % 2 ? "FFF9FBFD" : "FFFFFFFF";
    for (let c = 1; c <= 12; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (!sheet[addr]) sheet[addr] = { t: "s", v: "" };
      sheet[addr].s = { ...(sheet[addr].s || {}), fill: { patternType: "solid", fgColor: { rgb: altFill } } };
    }
    const nameAddr = XLSX.utils.encode_cell({ r, c: 9 });
    if (sheet[nameAddr]) sheet[nameAddr].s = { ...(sheet[nameAddr].s || {}), fill: { patternType: "solid", fgColor: { rgb: "FFFFFF00" } } };
    for (const dateCol of [10, 11, 12]) {
      const dateAddr = XLSX.utils.encode_cell({ r, c: dateCol });
      if (sheet[dateAddr] && sheet[dateAddr].v !== "") sheet[dateAddr].z = "yyyy-mm-dd";
    }
    for (let dayIndex = 0; dayIndex < person.daily.length; dayIndex += 1) {
      const daily = person.daily[dayIndex];
      const c = firstDayCol0 + dayIndex;
      const status = daily.managerStatus;
      const addr = XLSX.utils.encode_cell({ r, c });
      setCellStyle(sheet, r, c, legacyDailyStyle(status, daily));
    }
    const tones = ["normal", "normal", "dayoffAvailable", "dayoffUsed", person.dayoffExcess > 0 ? "danger" : "safe", "extraAvailable", person.subcompUsed > 0 ? "extraUsed" : "extraAvailable", person.subcompShortage > 0 ? "danger" : "safe"];
    for (let offset = 0; offset <= summaryEndCol0 - summaryStartCol0; offset += 1) {
      const addr = XLSX.utils.encode_cell({ r, c: summaryStartCol0 + offset });
      const tone = tones[offset] || (offset === 20 ? "note" : "normal");
      setCellStyle(sheet, r, summaryStartCol0 + offset, legacySummaryStyle(tone, offset === 20));
    }
  });
}

function legacyDailyStyle(status, daily) {
  const base = { font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } }, fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFD9E1E8") };
  const text = String(status || "");
  if (daily?.hasClockIn) return base;
  if (daily?.dayoffExcessDate) return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFF4B183" } }, font: { ...base.font, bold: true, color: { rgb: "FF9C5700" } } };
  if (daily?.issue || text === "출근 미입력" || text === "미등록" || text === "출ㆍ계 미입력") return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFF33E0D" } }, font: { ...base.font, sz: 8, bold: true, color: { rgb: "FFFFFFFF" } } };
  if (text === "계획 미입력") return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFFFC000" } }, font: { ...base.font, bold: true, color: { rgb: "FF7F4100" } } };
  if (BASIC_DAYOFF.has(text)) return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFDDEBF7" } } };
  if (SUBCOMP_STATUSES.has(text) || text.includes("대체") || text.includes("보상")) return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FF2F5597" } }, font: { ...base.font, bold: true, color: { rgb: "FFFFFFFF" } } };
  if (isAnnualStatus(text) || ["오전반차", "오후반차"].includes(text)) return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFC6E0B4" } } };
  if (["공가", "휴가", "경조", "무급휴가", "교육", "출산휴가", "육아휴직"].includes(text)) return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFDDEBF7" } } };
  if (text === "출근확인") return { ...base, fill: { patternType: "solid", fgColor: { rgb: "FFE2F0D9" } }, font: { ...base.font, bold: true, color: { rgb: "FF107C41" } } };
  return base;
}

function legacySummaryStyle(tone = "normal", leftAlign = false) {
  const palette = {
    normal: ["FFF7FAFC", "FF1F2937"], safe: ["FFE2F0D9", "FF375623"], dayoffAvailable: ["FFDDEBF7", "FF1F4E78"],
    dayoffUsed: ["FFE7E6E6", "FF404040"], extraAvailable: ["FFE4DFEC", "FF5F497A"], extraUsed: ["FFFCE4D6", "FF9E480E"],
    danger: ["FFF33E0D", "FFFFFFFF"], note: ["FFFFF2CC", "FF7F6000"],
  };
  const [fill, font] = palette[tone] || palette.normal;
  return { fill: { patternType: "solid", fgColor: { rgb: fill } }, font: { name: "맑은 고딕", sz: 10, bold: tone === "danger", color: { rgb: font } }, alignment: { horizontal: leftAlign ? "left" : "center", vertical: "center", wrapText: true }, border: thinBorder("FFD9E1E8") };
}

function compactNumber(value) {
  const n = Number(value || 0);
  return Number.isInteger(n) ? String(n) : n.toFixed(1).replace(/\.0$/, "");
}

function appendManagerCompareSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 관리자수정 비교`],
    ["출근·근무·시간값은 자동 반영 제외, 휴무·연차·반차·공가·경조·대체·보상만 반영합니다."],
    [],
    ["No", "업로드 매니저", "지역장", "지역", "매장명", "이름", "사번", "발생일", "관리자 입력값", "반영값", "결과", "사유"],
  ];
  analysis.managerCompareRows.forEach((row, index) => rows.push([
    index + 1, row.manager, row.regionalManager, row.region, row.storeName, row.employeeName, row.employeeId, row.date, row.raw, row.appliedValue, row.result, row.reason,
  ]));
  const sheet = makeReportSheet(rows, 3, { titleCols: 11 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 11 }, { wch: 10 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 10 }, { wch: 34 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "관리자수정 비교");
}

function appendEvidenceSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 출근증빙·휴무확인`],
    ["상담사근태_관리자반영 기준으로 출근기록이 없거나 승인 확인이 필요한 항목만 표시합니다. O 처리 후 재생성하면 해당 항목은 빨간색 대상에서 제외됩니다."],
    ["총 확인 건수", analysis.stats.evidence, "대상 인원", uniqueCount(analysis.evidenceRows.map((row) => row.employeeId)), "휴무 초과자는 다음 시트에서 별도 확인"],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "발생일", "요일", "근무계획", "관리자반영값", "확인 구분", "출근확인(O 입력)", "휴무확인(O 입력)", "연차확인(O 입력)", "오전반차(O 입력)", "오후반차(O 입력)", "기타확인", "처리메모"],
  ];
  analysis.evidenceRows.forEach((row, index) => rows.push([
    index + 1, row.regionalManager, row.manager, row.region, row.storeName, row.employeeName, row.employeeId, row.date, row.weekday, row.planStatus, row.managerStatus, row.issue, "", "", "", "", "", "", "",
  ]));
  const sheet = makeReportSheet(rows, 4, { titleCols: 18 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 6 }, { wch: 13 }, { wch: 14 }, { wch: 22 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 26 }];
  for (let r = 5; r < rows.length; r += 1) {
    setCellStyle(sheet, r, 11, STYLE.issue);
    for (let c = 12; c <= 18; c += 1) setCellStyle(sheet, r, c, STYLE.input);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, "출근증빙·휴무확인");
}

function appendAttendanceCheckSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 근태 확인 인원`],
    ["출근 기록, 매니저 파일의 휴무·공가·휴가·경조, 승인 연차가 모두 없는 날짜만 표시합니다. 웹에서 출근 확인 처리하면 근태 시트에 반영하는 기준입니다."],
    ["총 확인 건수", analysis.stats.evidence, "대상 인원", uniqueCount(analysis.evidenceRows.map((row) => row.employeeId))],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "발생일", "요일", "근무계획", "최종값", "확인 구분", "출근확인", "메모"],
  ];
  analysis.evidenceRows.forEach((row, index) => rows.push([
    index + 1, row.regionalManager, row.manager, row.region, row.storeName, row.employeeName, row.employeeId, row.date, row.weekday, row.planStatus, row.managerStatus, row.issue, "", "",
  ]));
  const sheet = makeReportSheet(rows, 4, { titleCols: 13 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 6 }, { wch: 13 }, { wch: 13 }, { wch: 20 }, { wch: 12 }, { wch: 28 }];
  for (let r = 5; r < rows.length; r += 1) {
    setCellStyle(sheet, r, 11, STYLE.issue);
    setCellStyle(sheet, r, 12, STYLE.input);
    setCellStyle(sheet, r, 13, STYLE.input);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, "근태 확인 인원");
}

function appendDayoffExcessSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 휴무 초과자`],
    ["상담사근태_관리자반영 기준으로 기본 휴무 가능 개수를 초과한 사람만 표시합니다. 대체휴무 잔여 개수는 다음 시트에서 확인합니다."],
    ["휴무 초과자", analysis.dayoffExcessRows.length],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "휴무 가능 개수", "당월 휴무 사용 개수", "휴무 초과 개수", "초과 휴무일", "비고"],
  ];
  analysis.dayoffExcessRows.forEach((row, index) => rows.push([
    index + 1,
    row.regionalManager,
    row.manager,
    row.region,
    row.storeName,
    row.employeeName,
    row.employeeId,
    row.dayoffAllowance,
    row.basicDayoffCount,
    row.dayoffExcess,
    row.daily.filter((daily) => daily.dayoffExcessDate).map((daily) => daily.date).join(", "),
    "기본 휴무 초과",
  ]));
  const sheet = makeReportSheet(rows, 4, { titleCols: 11 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 17 }, { wch: 14 }, { wch: 34 }, { wch: 18 }];
  for (let r = 5; r < rows.length; r += 1) {
    setCellStyle(sheet, r, 9, STYLE.issue);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, "휴무 초과자");
}

function appendSubcompSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 대체 사용자`],
    ["전 인원 기준입니다. 좌측에 당월 휴무 초과 여부를 표시하고, 웹에서 부여한 대체휴무 사용 가능·사용·이월·초과 여부를 분리해 확인합니다."],
    ["대체 사용자", analysis.subcompUserRows.length],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "당월 휴무 초과자", "대체 사용 가능 개수", "당월 사용 개수", "이월 개수", "초과 여부", "사용 가능 기간", "부여 비고", "사용일자"],
  ];
  analysis.people.forEach((row, index) => rows.push([
    index + 1,
    row.regionalManager,
    row.manager,
    row.region,
    row.storeName,
    row.employeeName,
    row.employeeId,
    row.dayoffExcess > 0 ? `Y (${daysText(row.dayoffExcess)} 초과)` : "",
    row.subcompAvailable,
    row.subcompUsed,
    row.subcompCarryover,
    row.subcompShortage > 0 ? `초과 ${daysText(row.subcompShortage)}` : "정상",
    row.substituteGrantPeriod,
    row.substituteGrantNotes,
    row.subcompEvents.map((event) => `${event.date} ${event.status}`).join(", "),
  ]));
  const sheet = makeReportSheet(rows, 4, { titleCols: 14 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 24 }, { wch: 24 }, { wch: 44 }];
  for (let r = 5; r < rows.length; r += 1) {
    if (text(rows[r][7])) setCellStyle(sheet, r, 7, STYLE.issue);
    if (String(rows[r][11]).startsWith("초과")) setCellStyle(sheet, r, 11, STYLE.issue);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, "대체 사용자");
}

function appendOtherStatusSheet(workbook, analysis) {
  const events = [];
  for (const person of analysis.people) {
    for (const daily of person.daily) {
      if (["공가", "휴가", "경조"].includes(daily.managerStatus)) {
        events.push({
          person,
          daily,
        });
      }
    }
  }
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 그 외 사용자`],
    ["매니저별 엑셀 파일 기준으로 공가, 휴가, 경조를 사용한 사람만 기록합니다."],
    ["총 사용 건수", events.length, "대상 인원", uniqueCount(events.map((row) => row.person.employeeId))],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "사용일", "구분", "반영 기준"],
  ];
  events.forEach(({ person, daily }, index) => rows.push([
    index + 1, person.regionalManager, person.manager, person.region, person.storeName, person.employeeName, person.employeeId, daily.date, daily.managerStatus, daily.managerSource === "manager" ? "매니저 파일" : daily.managerSource,
  ]));
  const sheet = makeReportSheet(rows, 4, { titleCols: 9 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "그 외 사용자");
}

function appendAnnualUserSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 연차 사용자`],
    ["연차 승인·반려 양식 기준으로 인정합니다. 근무계획에는 연차/반차인데 승인 양식이 없는 사람은 확인 대상으로 표시합니다."],
    ["확인 대상", analysis.annualIssueRows.length],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "발생일", "근무계획", "최종반영값", "신청구분", "신청상태", "판정"],
  ];
  analysis.annualRows.forEach((row, index) => rows.push([
    index + 1, row.regionalManager, row.manager, row.region, row.storeName, row.employeeName, row.employeeId, row.date, row.planStatus, row.finalStatus, row.applicationType, row.applicationStatus, row.result,
  ]));
  const sheet = makeReportSheet(rows, 4, { titleCols: 12 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 13 }, { wch: 13 }, { wch: 16 }, { wch: 18 }, { wch: 18 }];
  for (let r = 5; r < rows.length; r += 1) {
    setCellStyle(sheet, r, 12, rows[r][12] === "정상" ? STYLE.success : STYLE.issue);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, "연차 사용자");
}

function appendAnnualClosingDesignSheet(workbook, analysis, group) {
  const monthLabel = `${analysis.month.monthNo}월`;
  const isUnderOneYear = group.includes("미만");
  const sourceRows = annualDesignPeople(analysis).filter((row) => {
    const employeeGroup = row.closingPerson?.employmentGroup || row.employmentGroup || "";
    return isUnderOneYear ? employeeGroup.includes("미만") : !employeeGroup.includes("미만");
  });
  const sheetName = `상담사(${group})_${monthLabel}`;
  const title = `■ 삼성전자 ${analysis.routeLabel} 상담사 연차(${group})`;
  const headers = isUnderOneYear
    ? ["순번", "구분", "지역장", "매니저", "지역1", "지역2", "근무처코드", "근무처명", "사원구분", "스핀 사원번호", "제니엘 사번", "사원명", "연락처", "이메일", "제니엘 입사일", `${analysis.month.year}년 연차`, "전년 사용", `${analysis.month.year}년 사용`, "잔여연차", "사용기간", "1차촉진", "촉진 유무", "2차 촉진", "촉진 유무", "연차 소진일", "비고"]
    : ["순번", "입사구분", "지역장", "매니저", "지역1", "지역2", "근무처코드", "근무처명", "사원구분", "스핀 사원번호", "제니엘 사번", "사원명", "연락처", "이메일", "제니엘 입사일", "고용승계 입사일", "직무", `${analysis.month.year}년 발생`, `${analysis.month.year}년 사용`, "잔여연차", "사용기간", "1차촉진", "촉진 유무", "2차 촉진", "촉진 유무", "연차 소진일", "비고"];
  const rows = [
    [title],
    ["연차는 소진이 원칙이므로 사용기간 내 소진될 수 있도록 안내 부탁드립니다."],
    ["발생연차는 개별 입사일 기준 사용기간 내 전체 소진 기준으로 관리합니다."],
    [],
    ["※ 잔여연차 노란음영 셀 확인"],
    [],
    headers,
    headers.map(() => "필터용"),
  ];

  sourceRows.forEach((row, index) => {
    const closing = row.closingPerson || row;
    const used = row.monthlyAnnualUsed || 0;
    if (isUnderOneYear) {
      rows.push([
        index + 1,
        closing.employmentGroup || "재직 1년 미만",
        closing.regionalManager || row.regionalManager || "",
        closing.manager || row.manager || "",
        closing.region || row.region || "",
        closing.subRegion || row.subRegion || "",
        closing.storeCode || row.storeCode || "",
        closing.storeName || row.storeName || "",
        "판매상담사",
        closing.portalId || person.portalId || "",
        closing.employeeId || row.employeeId || "",
        closing.employeeName || row.employeeName || "",
        "",
        "",
        closing.hireDate || row.hireDate || "",
        numberOrBlank(closing.currentYearGranted),
        numberOrBlank(closing.priorYearUsed),
        roundHalf(numberValue(closing.currentYearUsed) + used),
        closingRemainingAnnual(row, closing),
        closing.annualPeriod || "",
        closing.firstPromotionDate || "",
        closing.firstPromotionDone || "",
        closing.secondPromotionDate || "",
        closing.secondPromotionDone || "",
        closing.annualExhaustionDate || "",
        closing.note || "",
      ]);
    } else {
      rows.push([
        index + 1,
        closing.employmentGroup || "재직 1년 이상",
        closing.regionalManager || row.regionalManager || "",
        closing.manager || row.manager || "",
        closing.region || row.region || "",
        closing.subRegion || row.subRegion || "",
        closing.storeCode || row.storeCode || "",
        closing.storeName || row.storeName || "",
        "판매상담사",
        closing.portalId || person.portalId || "",
        closing.employeeId || row.employeeId || "",
        closing.employeeName || row.employeeName || "",
        "",
        "",
        closing.hireDate || row.hireDate || "",
        closing.groupHireDate || row.groupHireDate || "",
        closing.job || "",
        numberOrBlank(closing.currentYearGranted),
        roundHalf(numberValue(closing.currentYearUsed) + used),
        closingRemainingAnnual(row, closing),
        closing.annualPeriod || "",
        closing.firstPromotionDate || "",
        closing.firstPromotionDone || "",
        closing.secondPromotionDate || "",
        closing.secondPromotionDone || "",
        closing.annualExhaustionDate || "",
        closing.note || "",
      ]);
    }
  });
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  const lastCol = headers.length - 1;
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: lastCol } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: lastCol } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: Math.min(5, lastCol) } },
  ];
  sheet["!cols"] = headers.map((header) => ({ wch: annualDesignWidth(header) }));
  sheet["!rows"] = rows.map((_, index) => ({ hpt: index === 0 ? 30 : index === 6 ? 40 : 24 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 8, topLeftCell: "A9", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 80, zoomScaleNormal: 80 }];
  styleRange(sheet, 0, 0, 0, lastCol, STYLE.title);
  styleRange(sheet, 1, 0, 2, lastCol, STYLE.note);
  styleRange(sheet, 4, 0, 4, Math.min(5, lastCol), STYLE.warning);
  styleRange(sheet, 6, 0, 6, lastCol, STYLE.header);
  styleRange(sheet, 7, 0, 7, lastCol, STYLE.metric);
  for (let r = 8; r < rows.length; r += 1) {
    styleRange(sheet, r, 0, r, lastCol, STYLE.normal);
    const remainingCol = headers.indexOf("잔여연차");
    if (remainingCol >= 0) setCellStyle(sheet, r, remainingCol, STYLE.warning);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName.slice(0, 31));
}

function appendAnnualCumulativeSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 연차 누적관리`],
    ["업로드한 전월 마감 누적 파일을 기준으로 당월 승인 연차 사용량을 더해 다음 달 누적 기준으로 관리합니다."],
    ["기준월", analysis.targetMonth, "전월 누적 파일", analysis.closingData?.people?.length ? `${analysis.closingData.people.length}명 인식` : "없음"],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "입사구분", "기존 잔여", "당월 사용", "마감 후 잔여", "1차 촉진", "2차 촉진", "비고"],
  ];
  analysis.people.forEach((row, index) => {
    const previousRemaining = numberOrBlank(row.annualLedger.remaining);
    const used = row.monthlyAnnualUsed || 0;
    const closingRemaining = previousRemaining === "" ? "" : roundHalf(previousRemaining - used);
    rows.push([
      index + 1,
      row.regionalManager || row.closingPerson?.regionalManager || "",
      row.manager || row.closingPerson?.manager || "",
      row.region || row.closingPerson?.region || "",
      row.storeName || row.closingPerson?.storeName || "",
      row.employeeName || row.closingPerson?.employeeName || "",
      row.employeeId,
      row.closingPerson?.employmentGroup || "",
      previousRemaining,
      used,
      closingRemaining,
      row.closingPerson?.firstPromotionDone || "",
      row.closingPerson?.secondPromotionDone || "",
      row.closingPerson?.note || "",
    ]);
  });
  const sheet = makeReportSheet(rows, 4, { titleCols: 13 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "연차 누적관리");
}

function appendAnnualPromotionSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 연차 촉진 관리`],
    ["누적 잔여와 당월 승인 사용량을 반영하고, 전월 1·2차 촉진 상태를 별도로 이어받습니다."],
    ["※ 촉진일·촉진 유무·소진일은 다음 달 마감본에 그대로 승계됩니다."],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "입사구분", "입사일", "기존 잔여", "당월 사용", "마감 후 잔여", "1차 촉진일", "1차 촉진 유무", "2차 촉진일", "2차 촉진 유무", "연차 소진일", "비고"],
  ];
  analysis.people.forEach((row, index) => {
    const closing = row.closingPerson || {};
    const previous = numberOrBlank(row.annualLedger.remaining);
    const after = previous === "" ? "" : roundHalf(previous - (row.monthlyAnnualUsed || 0));
    rows.push([
      index + 1, row.regionalManager || closing.regionalManager || "", row.manager || closing.manager || "", row.region || closing.region || "", row.storeName || closing.storeName || "",
      row.employeeName || closing.employeeName || "", row.employeeId, closing.employmentGroup || "", row.hireDate || closing.hireDate || "", previous, row.monthlyAnnualUsed || 0, after,
      closing.firstPromotionDate || "", closing.firstPromotionDone || "", closing.secondPromotionDate || "", closing.secondPromotionDone || "", closing.annualExhaustionDate || "", closing.note || row.annualLedger.note || "",
    ]);
  });
  const sheet = makeReportSheet(rows, 4, { titleCols: 17 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 10 }, { wch: 13 }, { wch: 14 }, { wch: 13 }, { wch: 11 }, { wch: 11 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 30 }];
  for (let r = 5; r < rows.length; r += 1) {
    setCellStyle(sheet, r, 11, STYLE.warning);
    for (const c of [12, 13, 14, 15, 16, 17]) setCellStyle(sheet, r, c, STYLE.input);
  }
  XLSX.utils.book_append_sheet(workbook, sheet, "연차 촉진 관리");
}

function appendClosingPeopleSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 월 마감 누적 파일`],
    ["다음 달 마감 때 다시 업로드할 기준 파일입니다. 연차 잔여, 촉진 상태, 기본 인력 정보를 보관합니다."],
    ["대상 인원", analysis.people.length, "생성 기준", new Date().toLocaleString("ko-KR")],
    [],
    ["No", "지역장", "매니저", "지역", "지역2", "매장코드", "매장명", "포탈사번", "제니엘사번", "이름", "입사구분", "제니엘입사일", "고용승계입사일", "연차 잔여", "당월 연차 사용", "1차 촉진", "2차 촉진", "연차 소진일", "비고"],
  ];
  analysis.people.forEach((row, index) => {
    const previousRemaining = numberOrBlank(row.annualLedger.remaining);
    const closingRemaining = previousRemaining === "" ? "" : roundHalf(previousRemaining - (row.monthlyAnnualUsed || 0));
    rows.push([
      index + 1,
      row.regionalManager || row.closingPerson?.regionalManager || "",
      row.manager || row.closingPerson?.manager || "",
      row.region || row.closingPerson?.region || "",
      row.subRegion || row.closingPerson?.subRegion || "",
      row.storeCode || row.closingPerson?.storeCode || "",
      row.storeName || row.closingPerson?.storeName || "",
      row.closingPerson?.portalId || row.portalId || "",
      row.employeeId,
      row.employeeName || row.closingPerson?.employeeName || "",
      row.closingPerson?.employmentGroup || "",
      row.hireDate || row.closingPerson?.hireDate || "",
      row.groupHireDate || row.closingPerson?.groupHireDate || "",
      closingRemaining,
      row.monthlyAnnualUsed || 0,
      row.closingPerson?.firstPromotionDone || "",
      row.closingPerson?.secondPromotionDone || "",
      row.closingPerson?.annualExhaustionDate || "",
      row.closingPerson?.note || "",
    ]);
  });
  const sheet = makeReportSheet(rows, 4, { titleCols: 18 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 9 }, { wch: 11 }, { wch: 18 }, { wch: 14 }, { wch: 13 }, { wch: 10 }, { wch: 14 }, { wch: 13 }, { wch: 13 }, { wch: 12 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 13 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "월마감 누적관리");
}

function appendClosingAnnualUsageSheet(workbook, analysis) {
  const usageById = new Map((analysis.closingData?.annualUsage || []).map((row) => [`${row.year}|${row.employeeId}`, row]));
  const year = analysis.month.year;
  const monthCol = `${analysis.month.monthNo}월`;
  const rows = [
    [`${year}년 연차 사용 누적`],
    ["기존 월별 연차사용 누적에 당월 승인 연차 사용량을 반영합니다."],
    [],
    ["포탈사번", "제니엘사번", "이름", "1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "총합계", "제니엘입사일", "고용승계입사일", "퇴사일", "비고"],
  ];
  for (const person of analysis.people) {
    const existing = usageById.get(`${year}|${person.employeeId}`) || {};
    const months = {};
    for (let m = 1; m <= 12; m += 1) months[`${m}월`] = numberValue(existing.months?.[`${m}월`]);
    months[monthCol] = roundHalf(months[monthCol] + (person.monthlyAnnualUsed || 0));
    const total = roundHalf(Object.values(months).reduce((sum, value) => sum + numberValue(value), 0));
    rows.push([
      existing.portalId || person.closingPerson?.portalId || person.portalId || "",
      person.employeeId,
      person.employeeName || person.closingPerson?.employeeName || "",
      ...Array.from({ length: 12 }, (_, index) => months[`${index + 1}월`] || 0),
      total,
      person.hireDate || person.closingPerson?.hireDate || "",
      person.groupHireDate || person.closingPerson?.groupHireDate || "",
      person.closingPerson?.resignDate || "",
      person.closingPerson?.note || "",
    ]);
  }
  const sheet = makeReportSheet(rows, 3, { titleCols: 19 });
  sheet["!cols"] = [{ wch: 14 }, { wch: 13 }, { wch: 10 }, ...Array.from({ length: 13 }, () => ({ wch: 9 })), { wch: 13 }, { wch: 13 }, { wch: 13 }, { wch: 26 }];
  XLSX.utils.book_append_sheet(workbook, sheet, `${year}년 연차사용`);
}

function appendClosingSubcompSheet(workbook, analysis) {
  const rows = [
    [`${analysis.month.year}년 ${analysis.month.monthNo}월 대체 사용 누적`],
    ["대체휴무 월 부여, 사용 가능 기간, 당월 사용량, 마감 후 잔여를 다음 달 기준으로 보관합니다."],
    [],
    ["No", "지역장", "매니저", "매장명", "이름", "사번", "사용 가능 개수", "당월 사용", "마감 후 잔여", "초과", "사용일자"],
  ];
  analysis.people.forEach((row, index) => rows.push([
    index + 1,
    row.regionalManager,
    row.manager,
    row.storeName,
    row.employeeName,
    row.employeeId,
    row.subcompAvailable,
    row.subcompUsed,
    roundHalf(row.subcompAvailable - row.subcompUsed),
    row.subcompShortage > 0 ? row.subcompShortage : "",
    row.subcompEvents.map((event) => `${event.date} ${event.status}`).join(", "),
  ]));
  const sheet = makeReportSheet(rows, 3, { titleCols: 10 });
  sheet["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "대체 사용 누적");
}

function appendRawSheet(workbook, name, rows) {
  const sheet = XLSX.utils.aoa_to_sheet(rows.length ? rows : [["데이터 없음"]]);
  sheet["!cols"] = rows[0]?.map(() => ({ wch: 16 })) || [{ wch: 18 }];
  if (rows.length) styleRange(sheet, 0, 0, 0, rows[0].length - 1, STYLE.header);
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

function annualDesignPeople(analysis) {
  const byId = new Map();
  for (const closing of analysis.closingData?.people || []) {
    byId.set(normalizeId(closing.employeeId), {
      ...closing,
      employeeId: normalizeId(closing.employeeId),
      employeeName: closing.employeeName,
      closingPerson: closing,
      monthlyAnnualUsed: 0,
    });
  }
  for (const person of analysis.people || []) {
    const id = normalizeId(person.employeeId);
    const previous = byId.get(id);
    byId.set(id, {
      ...(previous || {}),
      ...person,
      closingPerson: person.closingPerson || previous?.closingPerson || null,
    });
  }
  return [...byId.values()].filter((row) => row.employeeId).sort(personSort);
}

function closingRemainingAnnual(row, closing = {}) {
  const remaining = numberOrBlank(row?.annualLedger?.remaining ?? closing.remainingAnnual);
  if (remaining === "") return "";
  return roundHalf(remaining - (row.monthlyAnnualUsed || 0));
}

function annualDesignWidth(header) {
  if (["순번", "구분", "지역1", "지역2", "직무"].includes(header)) return 9;
  if (["지역장", "매니저", "사원명"].includes(header)) return 11;
  if (["근무처명", "사용기간", "비고"].includes(header)) return 20;
  if (header.includes("사번") || header.includes("입사일") || header.includes("소진일") || header.includes("촉진")) return 14;
  if (header.includes("연차") || header.includes("사용") || header.includes("발생")) return 13;
  if (header.includes("이메일")) return 22;
  if (header.includes("연락처")) return 15;
  return 12;
}

function makeReportSheet(rows, headerRowIndex, options = {}) {
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  const titleCols = options.titleCols || Math.max(...rows.map((row) => row.length)) - 1;
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: titleCols } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: titleCols } },
  ];
  sheet["!rows"] = rows.map((_, index) => ({ hpt: index === 0 ? 28 : index === headerRowIndex ? 32 : 23 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: headerRowIndex + 1, topLeftCell: `A${headerRowIndex + 2}`, activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 80, zoomScaleNormal: 80 }];
  styleRange(sheet, 0, 0, 0, titleCols, STYLE.title);
  styleRange(sheet, 1, 0, 1, titleCols, STYLE.note);
  if (headerRowIndex > 1) styleRange(sheet, 2, 0, Math.max(2, headerRowIndex - 1), titleCols, STYLE.metric);
  styleRange(sheet, headerRowIndex, 0, headerRowIndex, rows[headerRowIndex].length - 1, STYLE.header);
  for (let r = headerRowIndex + 1; r < rows.length; r += 1) styleRange(sheet, r, 0, r, rows[headerRowIndex].length - 1, STYLE.normal);
  return sheet;
}

function dailyCellStyle(status, daily) {
  if (daily.hasClockIn) return STYLE.normal;
  if (daily.dayoffExcessDate) return STYLE.issue;
  if (daily.issue) return STYLE.issue;
  if (BASIC_DAYOFF.has(status)) return STYLE.dayoff;
  if (isAnnualStatus(status) || ["공가", "휴가", "경조", "출산휴가", "육아휴직"].includes(status)) return STYLE.leave;
  if (SUBCOMP_STATUSES.has(status)) return STYLE.subcomp;
  if (status === "출근확인") return STYLE.success;
  if (isMissingStatus(status)) return STYLE.issue;
  return STYLE.normal;
}

function styleRange(sheet, r1, c1, r2, c2, style) {
  for (let r = r1; r <= r2; r += 1) {
    for (let c = c1; c <= c2; c += 1) setCellStyle(sheet, r, c, style);
  }
}

function setCellStyle(sheet, row, col, style) {
  const ref = XLSX.utils.encode_cell({ r: row, c: col });
  if (!sheet[ref]) sheet[ref] = { t: "s", v: "" };
  sheet[ref].s = style;
}

function thinBorder(color) {
  return {
    top: { style: "thin", color: { rgb: color } },
    right: { style: "thin", color: { rgb: color } },
    bottom: { style: "thin", color: { rgb: color } },
    left: { style: "thin", color: { rgb: color } },
  };
}

async function applyWorkbookViewSettings(raw, sheetNames) {
  if (typeof JSZip === "undefined") return raw;
  try {
    const zip = await JSZip.loadAsync(raw);
    for (let index = 0; index < sheetNames.length; index += 1) {
      const path = `xl/worksheets/sheet${index + 1}.xml`;
      const file = zip.file(path);
      if (!file) continue;
      const name = sheetNames[index];
      const freeze = name === "상담사근태" || name === "상담사근태_관리자반영" || name === "근태 시트"
        ? { xSplit: 9, ySplit: 6, topLeftCell: "J7", activePane: "bottomRight" }
        : { xSplit: 0, ySplit: 5, topLeftCell: "A6", activePane: "bottomLeft" };
      const xml = await file.async("string");
      zip.file(path, injectFreezePane(xml, freeze));
    }
    return await zip.generateAsync({ type: "arraybuffer" });
  } catch (error) {
    console.warn("freeze pane setting skipped", error);
    return raw;
  }
}

function injectFreezePane(xml, freeze) {
  const paneAttrs = [
    freeze.xSplit ? `xSplit="${freeze.xSplit}"` : "",
    `ySplit="${freeze.ySplit}"`,
    `topLeftCell="${freeze.topLeftCell}"`,
    `activePane="${freeze.activePane}"`,
    `state="frozen"`,
  ].filter(Boolean).join(" ");
  const paneXml = `<pane ${paneAttrs}/>`;
  let next = xml.replace(/<pane\b[^>]*\/>/g, "").replace(/<pane\b[^>]*>[\s\S]*?<\/pane>/g, "");
  if (/<sheetViews>[\s\S]*?<sheetView\b[^>]*>/.test(next)) {
    return next.replace(/(<sheetViews>[\s\S]*?<sheetView\b[^>]*>)/, `$1${paneXml}`);
  }
  const sheetViews = `<sheetViews><sheetView workbookViewId="0">${paneXml}</sheetView></sheetViews>`;
  return next.replace(/(<sheetFormatPr\b[^>]*\/>)/, `${sheetViews}$1`);
}

function firstTableRows(workbook, preferred = []) {
  const sheetName = chooseSheetName(workbook, preferred) || workbook.SheetNames[0];
  return sheetRows(workbook.Sheets[sheetName]);
}

function chooseSheetName(workbook, preferred = []) {
  if (!workbook?.SheetNames?.length) return "";
  for (const keyword of preferred) {
    const found = workbook.SheetNames.find((name) => name.includes(keyword));
    if (found) return found;
  }
  return workbook.SheetNames[0];
}

function attendanceSheetNames(workbook) {
  const names = workbook?.SheetNames || [];
  const preferred = names.filter((name) => {
    const compact = cleanHeader(name);
    return compact.includes("근태raw") || compact.includes("출퇴근") || compact.includes("출근기록");
  });
  if (preferred.length) return preferred;
  return names.filter((name) => {
    const compact = cleanHeader(name);
    return !compact.includes("상담사근태")
      && !compact.includes("관리자")
      && !compact.includes("증빙")
      && !compact.includes("휴무초과")
      && !compact.includes("대체")
      && !compact.includes("연차")
      && !compact.includes("요약")
      && !compact.includes("확인자")
      && !compact.includes("정산");
  });
}

function sheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false, blankrows: false });
}

function parseGenericRows(rows) {
  const headerIndex = findHeaderRow(rows, [["사번", "직원번호", "사원번호", "이름", "성명", "매장명", "점포명"]]);
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((cell) => text(cell) || "");
  const parsed = [];
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const source = rows[r];
    if (!source.some((cell) => text(cell))) continue;
    const item = {};
    headers.forEach((header, index) => {
      if (header) item[header] = source[index];
    });
    parsed.push(item);
  }
  return parsed;
}

function findHeaderRow(rows, groups) {
  for (let i = 0; i < Math.min(rows.length, 80); i += 1) {
    const rowText = rows[i].map((cell) => cleanHeader(cell));
    const matched = groups.every((group) => group.some((word) => rowText.some((cell) => cell.includes(cleanHeader(word)))));
    if (matched) return i;
  }
  return -1;
}

function buildHeaderIndexes(headers) {
  return {
    route: firstHeaderIndex(headers, ["경로", "회사"]),
    regionalManager: firstHeaderIndex(headers, ["지역장", "총괄"]),
    manager: firstHeaderIndex(headers, ["매니저", "담당매니저", "관리자"]),
    region: firstHeaderIndex(headers, ["지역1", "지역", "권역"]),
    subRegion: firstHeaderIndex(headers, ["지역2", "세부지역", "상세지역"]),
    storeCode: firstHeaderIndex(headers, ["매장코드", "점포코드", "코드"]),
    storeName: firstHeaderIndex(headers, ["매장명", "점포명", "매장", "점포"]),
    employeeId: firstHeaderIndex(headers, ["사번", "직원번호", "사원번호", "employeeid"]),
    employeeName: firstHeaderIndex(headers, ["이름", "성명", "직원명", "상담사"]),
    hireDate: firstHeaderIndex(headers, ["입사일"]),
    groupHireDate: firstHeaderIndex(headers, ["그룹입사일", "제니엘입사일"]),
    date: firstHeaderIndex(headers, ["근무일자", "발생일", "일자", "날짜", "휴가일자"]),
    issueDate: firstHeaderIndex(headers, ["발생일"]),
    startDate: firstHeaderIndex(headers, ["시작일", "휴가시작일", "시작"]),
    endDate: firstHeaderIndex(headers, ["종료일", "휴가종료일", "종료"]),
    leaveType: firstHeaderIndex(headers, ["휴가구분", "신청구분", "구분", "휴가종류", "연차구분"]),
    status: firstHeaderIndex(headers, ["상태", "근태", "근무계획"]),
    approvalStatus: firstHeaderIndex(headers, ["승인", "결재", "처리상태", "신청상태", "상태"]),
  };
}

function firstHeaderIndex(headers, aliases) {
  const normalizedAliases = aliases.map(cleanHeader);
  const exactIndex = headers.findIndex((header) => {
    const clean = cleanHeader(header);
    return normalizedAliases.some((alias) => clean === alias);
  });
  if (exactIndex >= 0) return exactIndex;
  return headers.findIndex((header) => {
    const clean = cleanHeader(header);
    return normalizedAliases.some((alias) => clean.includes(alias));
  });
}

function allHeaderIndexes(headers, aliases) {
  const normalizedAliases = aliases.map(cleanHeader);
  const indexes = [];
  headers.forEach((header, index) => {
    const clean = cleanHeader(header);
    if (normalizedAliases.some((alias) => clean === alias || clean.includes(alias))) indexes.push(index);
  });
  return indexes;
}

function parseDayHeader(value, month) {
  const raw = text(value);
  if (!raw) return 0;
  const cleaned = raw.replace(/\s/g, "");
  const iso = normalizeDateText(raw, month);
  if (iso?.startsWith(month.key)) return Number(iso.slice(-2));
  let match = cleaned.match(/^(\d{1,2})일?$/);
  if (match) {
    const day = Number(match[1]);
    return day >= 1 && day <= month.daysInMonth ? day : 0;
  }
  match = cleaned.match(/^(\d{1,2})[./-](\d{1,2})$/);
  if (match) {
    const day = Number(match[2]);
    return day >= 1 && day <= month.daysInMonth ? day : 0;
  }
  return 0;
}

export function parseTargetMonth(targetMonth = "") {
  const match = String(targetMonth || "").match(/^(\d{4})-(\d{2})$/);
  if (!match) return { valid: false, key: "", year: 0, monthNo: 0, daysInMonth: 0 };
  const year = Number(match[1]);
  const monthNo = Number(match[2]);
  const daysInMonth = new Date(year, monthNo, 0).getDate();
  return { valid: Boolean(year && monthNo && daysInMonth), key: `${year}-${String(monthNo).padStart(2, "0")}`, year, monthNo, daysInMonth };
}

export function getDayoffAllowance(route, month) {
  if (route === "electroland") {
    let count = 0;
    for (let day = 1; day <= month.daysInMonth; day += 1) {
      const weekday = new Date(month.year, month.monthNo - 1, day).getDay();
      if (weekday === 0 || weekday === 6) count += 1;
    }
    return count;
  }
  return 6;
}

function normalizeMember(row = {}) {
  return {
    route: text(row.route || row["경로"] || ""),
    regionalManager: text(row.regionalManager || row["지역장"] || ""),
    manager: text(row.manager || row["매니저"] || row["담당매니저"] || ""),
    region: text(row.region || row["지역"] || ""),
    subRegion: text(row.subRegion || row["지역2"] || ""),
    storeCode: text(row.storeCode || row["매장코드"] || ""),
    storeName: text(row.storeName || row["매장명"] || row.store || row["점포명"] || ""),
    portalId: text(row.portalId || row["포탈사번"] || row["포털사번"] || row["스핀사번"] || row["스핀사원번호"] || ""),
    employeeId: normalizeId(row.employeeId || row["제니엘사번"] || row["사번"] || row["직원번호"] || ""),
    employeeName: text(row.employeeName || row.name || row["이름"] || row["성명"] || ""),
    hireDate: normalizeDateText(row.hireDate || row["제니엘입사일"] || row["입사일"] || ""),
    groupHireDate: normalizeDateText(row.groupHireDate || row["그룹입사일"] || row["고용승계입사일"] || ""),
    note: text(row.note || row["비고"] || ""),
  };
}

function normalizePerson(route, row) {
  const member = normalizeMember(row);
  return {
    ...member,
    route,
    plan: row.plan || null,
  };
}

function blankToFallback(primary, fallback) {
  const result = {};
  for (const key of ["regionalManager", "manager", "region", "subRegion", "storeCode", "storeName", "portalId", "employeeName", "hireDate", "groupHireDate"]) {
    result[key] = text(fallback?.[key]) ? fallback[key] : primary[key];
  }
  result.employeeId = primary.employeeId || fallback?.employeeId || "";
  return result;
}

function normalizeStatus(value) {
  const raw = text(value).replace(/\s+/g, "");
  if (!raw) return "";
  if (isClockValue(raw)) return raw;
  if (/출[ㆍ·]?계미입력|출계미입력/.test(raw)) return "출ㆍ계 미입력";
  if (raw.includes("출근미입력")) return "출근 미입력";
  if (raw.includes("근태미입력")) return "근태 미입력";
  if (raw.includes("미등록") || raw === "미입력") return "미등록";
  if (raw.includes("오전반차")) return "오전반차";
  if (raw.includes("오후반차")) return "오후반차";
  if (raw.includes("반차") && raw.includes("오전")) return "오전반차";
  if (raw.includes("반차") && raw.includes("오후")) return "오후반차";
  if (raw.includes("출산")) return "출산휴가";
  if (raw.includes("육아")) return "육아휴직";
  if (raw.includes("대체")) return "대체휴무";
  if (raw.includes("보상")) return "보상휴가";
  if (raw.includes("연차")) return "연차";
  if (raw.includes("공가")) return "공가";
  if (raw.includes("휴가")) return "휴가";
  if (raw.includes("경조")) return "경조";
  if (raw.includes("휴무")) return "휴무";
  if (raw.includes("교육")) return "교육";
  if (raw.includes("근무A")) return "근무A";
  if (raw.includes("근무B")) return "근무B";
  if (raw.includes("근무C")) return "근무C";
  if (raw.includes("근무")) return "근무";
  if (raw.includes("출근확인")) return "출근확인";
  if (raw === "공백" || raw === "-") return "공백";
  return text(value);
}

function normalizeLeaveType(value) {
  const status = normalizeStatus(value);
  if (["연차", "오전반차", "오후반차", "출산휴가", "육아휴직", "공가", "경조"].includes(status)) return status;
  return status || "연차";
}

function resolveEvidenceStatus(row, headers) {
  const yesColumn = (aliases) => {
    const indexes = allHeaderIndexes(headers, aliases);
    return indexes.some((index) => isYes(row[index]));
  };
  if (yesColumn(["출근확인"])) return { confirmed: true, status: "출근확인", rawStatus: "출근확인" };
  if (yesColumn(["휴무확인"])) return { confirmed: true, status: "휴무", rawStatus: "휴무확인" };
  if (yesColumn(["연차확인"])) return { confirmed: true, status: "연차", rawStatus: "연차확인" };
  if (yesColumn(["오전반차"])) return { confirmed: true, status: "오전반차", rawStatus: "오전반차" };
  if (yesColumn(["오후반차"])) return { confirmed: true, status: "오후반차", rawStatus: "오후반차" };

  const statusIndexes = allHeaderIndexes(headers, ["처리결과", "확정값", "확인결과", "처리상태", "구분", "기타확인"]);
  for (const index of statusIndexes) {
    const raw = text(row[index]);
    const status = normalizeStatus(raw);
    if (MANAGER_ALLOWED.has(status) || status === "출근확인") return { confirmed: true, status, rawStatus: raw };
    if (isYes(raw)) return { confirmed: true, status: "", rawStatus: raw };
  }
  return { confirmed: false, status: "", rawStatus: "" };
}

function isAttendanceLike(value) {
  const raw = text(value).replace(/\s+/g, "");
  if (!raw) return false;
  if (isClockValue(raw)) return true;
  return raw === "출근" || raw === "근무" || raw.includes("근무완료") || raw.includes("정상출근");
}

function isClockValue(value) {
  const raw = text(value).trim();
  if (!raw) return false;
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(raw)) return true;
  if (/^\d{1,2}시\s*\d{0,2}분?$/.test(raw)) return true;
  return false;
}

function normalizeClock(value) {
  const raw = text(value).trim();
  const match = raw.match(/^(\d{1,2}):(\d{2})/);
  if (match) return `${String(Number(match[1])).padStart(2, "0")}:${match[2]}`;
  return raw;
}

function isMissingStatus(status) {
  return ["공백", "미등록", "출근 미입력", "출ㆍ계 미입력", "출계 미입력", "근태 미입력", ""].includes(status);
}

function missingLabel(status) {
  if (String(status).includes("출ㆍ계") || String(status).includes("출계")) return "출ㆍ계 미입력";
  if (String(status).includes("근태")) return "근태 미입력";
  if (String(status).includes("미등록")) return "미등록";
  return "출근 미입력";
}

function isAnnualStatus(status) {
  return status === "연차" || HALF_DAY_STATUSES.has(status);
}

function isApprovedStatus(value) {
  const raw = text(value).replace(/\s+/g, "");
  if (!raw) return false;
  if (isRejectedStatus(raw)) return false;
  return raw.includes("승인") || raw.includes("완료") || raw.includes("확정") || raw.includes("결재완료");
}

function isRejectedStatus(value) {
  const raw = text(value).replace(/\s+/g, "");
  return raw.includes("반려") || raw.includes("취소") || raw.includes("철회") || raw.includes("삭제") || raw.includes("불가");
}

function isYes(value) {
  const raw = text(value).replace(/\s+/g, "").toUpperCase();
  return ["O", "OK", "Y", "YES", "1", "완료", "확인"].includes(raw) || raw.includes("확인완료");
}

function statusAmount(rawValue) {
  const raw = text(rawValue);
  const status = normalizeStatus(raw);
  if (HALF_DAY_STATUSES.has(status) || raw.includes("0.5") || raw.includes("0.5일") || raw.includes("반일")) return 0.5;
  return 1;
}

function resolveSubstituteGrants(person, grants, route, month) {
  const matched = (grants || []).filter((grant) => {
    if (grant.route && grant.route !== route) return false;
    if (grant.grantMonth && grant.grantMonth !== month.key) return false;
    if (grant.employeeId && normalizeId(grant.employeeId) !== person.employeeId) return false;
    if (grant.manager && text(grant.manager) !== text(person.manager)) return false;
    if (grant.storeName && text(grant.storeName) !== text(person.storeName)) return false;
    return true;
  });
  return {
    available: roundHalf(matched.reduce((sum, grant) => sum + numberValue(grant.grantedDays ?? grant.days ?? grant.available), 0)),
    grants: matched,
  };
}

function normalizeEvidenceOverrideStatus(value) {
  const raw = text(value).replace(/\s+/g, "");
  if (!raw) return "";
  if (["제외", "무시", "삭제"].includes(raw)) return "제외";
  if (raw.includes("출근")) return "출근확인";
  if (raw.includes("휴무")) return "휴무";
  if (raw.includes("연차")) return "연차";
  if (raw.includes("오전반차")) return "오전반차";
  if (raw.includes("오후반차")) return "오후반차";
  if (raw.includes("공가")) return "공가";
  if (raw.includes("휴가")) return "휴가";
  if (raw.includes("경조")) return "경조";
  if (raw.includes("대체")) return "대체휴무";
  if (raw.includes("보상")) return "보상휴가";
  const normalized = normalizeStatus(value);
  return MANAGER_ALLOWED.has(normalized) || normalized === "출근확인" ? normalized : "";
}

function resolveAnnualLedger(employeeId, rows, closingPerson = null) {
  // 사용자가 올린 전월 마감본을 최우선 기준으로 사용합니다. 저장된 웹 DB가 오래된 경우 중복 차감을 막습니다.
  if (closingPerson && numberOrBlank(closingPerson.remainingAnnual) !== "") {
    return { remaining: numberOrBlank(closingPerson.remainingAnnual), note: closingPerson.note || "" };
  }
  const row = findByEmployeeId(rows, employeeId);
  if (!row) return { remaining: "", note: closingPerson?.note || "" };
  return {
    remaining: numberOrBlank(row["연차 잔여"] ?? row["잔여연차"] ?? row["잔여"] ?? row.remaining ?? row.remainingDays),
    note: text(row["비고"] ?? row.note ?? ""),
  };
}

function resolveClosingPerson(employeeId, people = []) {
  const id = normalizeId(employeeId);
  return people.find((row) => normalizeId(row.employeeId) === id) || null;
}

function findByEmployeeId(rows, employeeId) {
  const id = normalizeId(employeeId);
  return rows.find((row) => normalizeId(row["사번"] ?? row["직원번호"] ?? row.employeeId ?? row.id) === id);
}

function emptyClosingData() {
  return { people: [], annualUsage: [], sourceSheetNames: [] };
}

function workbookProps(analysis, subject) {
  return {
    Title: `${analysis.targetMonth} ${analysis.routeLabel} ${subject}`,
    Subject: `근태 관리 시스템 ${subject}`,
    Author: "근태 관리 시스템",
    Comments: "신규 clean 기준으로 생성",
  };
}

function namedWorkbookFile(buffer, fileName) {
  return new File([buffer], fileName, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function planRawRows(analysis) {
  const days = [...Array(analysis.month.daysInMonth)].map((_, index) => index + 1);
  const rows = [["지역장", "매니저", "지역", "지역2", "매장코드", "매장명", "사번", "이름", "입사일", "그룹입사일", ...days.map((day) => `${String(day).padStart(2, "0")}일`)]];
  for (const row of analysis.planRows) {
    rows.push([row.regionalManager, row.manager, row.region, row.subRegion, row.storeCode, row.storeName, row.employeeId, row.employeeName, row.hireDate, row.groupHireDate, ...days.map((day) => row.plans?.[day] || "")]);
  }
  return rows;
}

function attendanceRawRows(analysis) {
  const rows = [["사번", "이름", "근무일자", "출근시간", "근태상태", "원본시트", "원본행"]];
  for (const row of analysis.attendanceRows) rows.push([row.employeeId, row.employeeName, row.date, row.clockIn, row.status, row.sheetName, row.rowNumber]);
  return rows;
}

function parseDateObject(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const raw = text(value);
  if (!raw) return null;
  const cleaned = raw.replace(/[.]/g, "-").replace(/[년월]/g, "-").replace(/[일]/g, "").replace(/\s+/g, "");
  let match = cleaned.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  match = cleaned.match(/^(\d{1,2})-(\d{1,2})$/);
  if (match) return { month: Number(match[1]), day: Number(match[2]) };
  return null;
}

function outputDateValue(value) {
  const normalized = normalizeDateText(value);
  const match = String(normalized || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return normalized || "";
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0);
}

function normalizeDateText(value, month = null) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const date = excelSerialDate(value);
    return toISODate(date);
  }
  const parsed = parseDateObject(value);
  if (parsed instanceof Date) return toISODate(parsed);
  if (parsed && month) {
    return `${month.year}-${String(parsed.month).padStart(2, "0")}-${String(parsed.day).padStart(2, "0")}`;
  }
  const raw = text(value);
  const compactDate = raw.replace(/\D/g, "");
  if (/^\d{8}$/.test(compactDate)) {
    return `${compactDate.slice(0, 4)}-${compactDate.slice(4, 6)}-${compactDate.slice(6, 8)}`;
  }
  const dayOnly = raw.match(/^(\d{1,2})일?$/);
  if (dayOnly && month) return `${month.key}-${String(Number(dayOnly[1])).padStart(2, "0")}`;
  return raw;
}

function expandDateRange(start, end) {
  const startDate = parseDateObject(start);
  const endDate = parseDateObject(end) || startDate;
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) return [start].filter(Boolean);
  const dates = [];
  for (let cursor = new Date(startDate); cursor <= endDate; cursor.setDate(cursor.getDate() + 1)) dates.push(toISODate(cursor));
  return dates;
}

function excelSerialDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function toISODate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function cleanHeader(value) {
  return text(value).replace(/\s+/g, "").replace(/[()[\]{}·ㆍ_\-]/g, "").toLowerCase();
}

function normalizeId(value) {
  const raw = text(value).trim();
  if (!raw) return "";
  if (/^\d+(\.0+)?$/.test(raw)) return String(Number(raw));
  return raw.replace(/\s+/g, "");
}

function text(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function pick(row, index) {
  return index >= 0 ? row[index] : "";
}

function firstText(values) {
  return text(values.find((value) => text(value)));
}

function firstValidIndex(...indexes) {
  return indexes.find((index) => Number.isInteger(index) && index >= 0) ?? -1;
}

function emptyAttendance(date) {
  return { date, clockIn: "", hasClockIn: false, status: "" };
}

function numberValue(value) {
  const n = Number(String(value ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? roundHalf(n) : 0;
}

function numberOrBlank(value) {
  const n = numberValue(value);
  return n || n === 0 ? n : "";
}

function roundHalf(value) {
  return Math.round(Number(value || 0) * 2) / 2;
}

function daysText(value) {
  return `${roundHalf(value)}일`;
}

function uniqueCount(values) {
  return new Set(values.filter(Boolean)).size;
}

function personSort(a, b) {
  return [a.regionalManager, a.manager, a.region, a.storeName, a.employeeName].join("|").localeCompare([b.regionalManager, b.manager, b.region, b.storeName, b.employeeName].join("|"), "ko");
}
