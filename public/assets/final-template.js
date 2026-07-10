const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];
const NON_WORK_CODES = new Set([
  "휴무", "무급휴가", "연차", "공가", "휴가", "경조", "출산휴가", "육아휴직",
  "대체휴일(1일)", "대체휴일(0.5일)", "보상휴가(1일)", "보상휴가(0.5일)",
]);
const WORK_CLOCK_CODES = new Set([
  "공백", "근무", "근무A", "근무B", "근무C", "교육", "오전반차", "오후반차",
]);

const KOREAN_PUBLIC_HOLIDAYS = {
  2026: new Set([
    "2026-01-01",
    "2026-02-16", "2026-02-17", "2026-02-18",
    "2026-03-01", "2026-03-02",
    "2026-05-05", "2026-05-24", "2026-05-25",
    "2026-06-03", "2026-06-06",
    "2026-08-15", "2026-08-17",
    "2026-09-24", "2026-09-25", "2026-09-26",
    "2026-10-03", "2026-10-05", "2026-10-09",
    "2026-12-25",
  ]),
};

export async function buildFinalTemplateWorkbook(result) {
  const response = await fetch("./assets/attendance-final-template.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("최종본 엑셀 양식 파일을 불러오지 못했습니다.");
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellStyles: true, cellDates: true, bookVBA: true });

  const monthNo = Number(String(result.targetMonth || "").slice(5, 7));
  const year = Number(String(result.targetMonth || "").slice(0, 4));
  const daysInMonth = new Date(year, monthNo, 0).getDate();

  renameSheet(workbook, "5월 상담사근태", "상담사근태");
  renameSheet(workbook, "상담사 근태", "상담사근태");
  renameSheet(workbook, "증빙(필수기입)", "출근 미등록");
  renameSheet(workbook, "증빙", "출근 미등록");
  renameSheet(workbook, "5월 근무계획", "근무 계획");
  renameSheet(workbook, "근무계획", "근무 계획");
  renameSheet(workbook, "근태RAW", "근태 RAW");

  const context = buildContext(result, daysInMonth);
  fillMainSheet(workbook.Sheets["상담사근태"], result, context, year, monthNo, daysInMonth);
  buildEvidenceDashboardSheet(workbook, result, context, year, monthNo);
  buildPlanAttendanceMatchSheet(workbook, result, context, year, monthNo);
  buildDayoffSubstituteSheet(workbook, result, context, year, monthNo);
  buildAnnualComparisonSheet(workbook, result, year, monthNo);
  buildIssueSummarySheet(workbook, result, context, year, monthNo);
  buildManagerRequestSheet(workbook, result, year, monthNo);
  buildWeeklyAttendanceCheckSheet(workbook, result, year, monthNo);
  buildAnnualLedgerSheet(workbook, result, context, year, monthNo);
  buildPersonnelStatusSheet(workbook, result, year, monthNo);
  buildHrPayrollAuditSheets(workbook, result, year, monthNo);
  fillPlanSheet(workbook.Sheets["근무 계획"], result, context, daysInMonth);
  fillAttendanceRawSheet(workbook.Sheets["근태 RAW"], result);

  workbook.SheetNames = [
    "상담사근태",
    "출근증빙·휴무확인",
    "계획&근태 상이 인원",
    "휴무 초과자",
    "인사팀 급여 확정표",
    "출근 미달자 정산",
    "연차 사용 필요자",
    "최종 문제자",
    "전체 요약본",
    "매니저별 이상 근태",
    "주 근태 확인자",
    "해당 월 연차 등록 현황 및 일자",
    "연차 누적 현황",
    "인력 변동 확인",
    "근무 계획",
    "근태 RAW",
  ];
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.CalcPr = { calcMode: "auto", fullCalcOnLoad: true, forceFullCalc: true };
  workbook.Props = {
    ...(workbook.Props || {}),
    Title: `${year}년 ${monthNo}월 ${result.routeLabel} 상담사 출퇴근현황`,
    Subject: "근태·연차 누적 관리 시스템 자동 생성 최종본",
    Author: "근태 관리 시스템",
    Comments: "최초 연차 초본 + 월별 승인·반려 + 근무계획 + 근태 원본 자동 매칭",
  };
  sanitizeWorkbookForExcel(workbook);
  return workbook;
}

export async function buildFinalTemplateFile(result) {
  const workbook = await buildFinalTemplateWorkbook(result);
  const rawBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true, bookVBA: true });
  const evidenceBuffer = await applyLiveEvidenceConditionalFormatting(rawBuffer, result);
  const buffer = await applyWorkbookOpenViewSettings(evidenceBuffer);
  const monthText = String(result.targetMonth || "").replace("-", "년 ") + "월";
  return new File([buffer], `${monthText}_${result.routeLabel}_출퇴근현황_최종본.xlsx`, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function buildContext(result, daysInMonth) {
  const workforce = [...(result.workforce?.members || [])]
    .filter((row) => row.route === result.route)
    .sort(workforceSort);
  const workforceById = new Map();
  for (const row of workforce) {
    const id = normalizeId(row.employeeId || row.employee_id);
    if (!id) continue;
    if (!workforceById.has(id)) workforceById.set(id, []);
    workforceById.get(id).push(normalizeMember(row));
  }

  const planById = groupBy(result.plan?.rows || [], (row) => normalizeId(row.employeeId));
  const attendanceByKey = buildAttendanceMap(result.attendance?.rows || []);
  const summaryById = new Map((result.employeeSummaries || []).map((row) => [normalizeId(row.employeeId), row]));
  const approvedLeaveStatusByKey = buildApprovedLeaveStatusMap(result.annualApplications || [], result.targetMonth);
  const leaveApplicationStatusByKey = buildLeaveApplicationStatusMap(result.annualApplications || [], result.targetMonth);
  const evidenceSet = new Set((result.evidenceOverrides || []).map(String));
  const issueMap = new Map();
  for (const issue of result.mismatchRows || []) {
    // 계획 미입력 + 실제 출근은 상담사근태의 경고색/비고에는 넣지 않고 계획&근태 상이 시트에서만 표시합니다.
    if (issue.issueType === "missing_plan") continue;
    addIssue(issueMap, issue.employeeId, issue.date, issue.reason || issue.result);
  }
  for (const summary of result.employeeSummaries || []) {
    // 기본 휴무 초과분은 대체휴무·보상휴가로 모두 대체된 경우 오류로 표시하지 않습니다.
    // 대체하지 못한 날짜만 빨간색 확인 대상으로 표시합니다.
    const dayoffShortageDates = resolveLeaveShortageDates(
      summary.baseExcessEvents || [],
      Number(summary.dayoffReplacementShortage || 0),
      summary.dayoffReplacementShortageDates || [],
    );
    for (const date of dayoffShortageDates) {
      addIssue(issueMap, summary.employeeId, date, `휴무초과 확인 요청 · ${daysText(summary.dayoffReplacementShortage)} 미대체`);
    }
    // 대체휴무와 보상휴가는 각각의 초과 발생일만 빨간색으로 표시합니다.
    const substituteShortageDates = resolveLeaveShortageDates(
      summary.substituteEvents || [],
      Number(summary.shortage || 0),
      summary.substituteShortageDates || [],
    );
    for (const date of substituteShortageDates) {
      addIssue(issueMap, summary.employeeId, date, `대체휴무 잔여 부족 · 총 ${daysText(summary.shortage)} 초과 사용`);
    }
    const compensationShortageDates = resolveLeaveShortageDates(
      summary.compensationEvents || [],
      Number(summary.compensationShortage || 0),
      summary.compensationShortageDates || [],
    );
    for (const date of compensationShortageDates) {
      addIssue(issueMap, summary.employeeId, date, `보상휴가 잔여 부족 · 총 ${daysText(summary.compensationShortage)} 초과 사용`);
    }
  }

  const people = [];
  const seenMemberKeys = new Set();
  const seenEmployeeIds = new Set();
  for (const member of workforce) {
    const id = normalizeId(member.employeeId || member.employee_id);
    if (!id) continue;
    const normalizedMember = normalizeMember(member);
    const memberKey = `${id}|${String(normalizedMember.storeCode || normalizedMember.storeName || "미지정")}`;
    if (seenMemberKeys.has(memberKey)) continue;
    seenMemberKeys.add(memberKey);
    seenEmployeeIds.add(id);
    const plan = choosePlan(planById.get(id) || [], normalizedMember);
    if (!plan) addIssue(issueMap, id, `${result.targetMonth}-01`, "인력·매장매칭에는 있으나 근무계획에 사번 없음");
    people.push({ key: memberKey, member: normalizedMember, plan, employeeId: id, name: normalizedMember.employeeName || plan?.name || "" });
  }
  for (const [id, plans] of planById.entries()) {
    if (seenEmployeeIds.has(id)) continue;
    const plan = plans[0];
    const member = normalizeMember({
      route: result.route,
      regionalManager: "",
      manager: "",
      region1: "",
      region2: "",
      storeCode: "",
      storeName: plan.store || "",
      portalId: "",
      employeeId: id,
      employeeName: plan.name || "",
      hireDate: "",
      groupHireDate: "",
      note: "인력·매장매칭 미등록",
    });
    const memberKey = `${id}|PLAN_ONLY`;
    addIssue(issueMap, id, `${result.targetMonth}-01`, "근무계획에는 있으나 인력·매장매칭에 사번 없음");
    people.push({ key: memberKey, member, plan, employeeId: id, name: plan.name || "" });
  }

  const dailyByKey = new Map();
  for (const person of people) {
    const daily = {};
    const summary = summaryById.get(person.employeeId) || {};
    const substituteShortageDateSet = new Set(resolveLeaveShortageDates(
      summary.substituteEvents || [], Number(summary.shortage || 0), summary.substituteShortageDates || [],
    ));
    const compensationShortageDateSet = new Set(resolveLeaveShortageDates(
      summary.compensationEvents || [], Number(summary.compensationShortage || 0), summary.compensationShortageDates || [],
    ));
    const dayoffExcessDateSet = new Set((summary.baseExcessEvents || []).map((event) => String(event?.date || "")).filter(Boolean));
    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = `${result.targetMonth}-${String(day).padStart(2, "0")}`;
      const rawPlanStatus = normalizePlan(person.plan?.plans?.[day]);
      const attendance = attendanceByKey.get(`${person.employeeId}|${date}`) || emptyAttendance();
      const evidenceKey = `${person.employeeId}|${date}`;
      const evidence = evidenceSet.has(evidenceKey);
      const finalAttendance = evidence
        ? { ...attendance, hasClockIn: true, actualStatus: "출근", evidenced: true }
        : attendance;
      // 연차신청현황에서 승인된 휴가/연차는 상담사근태에 먼저 반영합니다.
      // 단, 실제 출근기록이나 출근증빙이 있으면 기존 원칙대로 출근시간이 우선입니다.
      const approvedLeaveStatus = approvedLeaveStatusByKey.get(evidenceKey);
      const leaveApplicationStatus = leaveApplicationStatusByKey.get(evidenceKey);
      const planStatus = (!finalAttendance.hasClockIn && approvedLeaveStatus) ? approvedLeaveStatus.displayStatus : rawPlanStatus;
      const approvedLeaveResolvesMissing = Boolean(approvedLeaveStatus && !finalAttendance.hasClockIn);
      const plannedLeaveApprovalIssue = Boolean(!finalAttendance.hasClockIn && !approvedLeaveStatus && requiresLeaveApproval(rawPlanStatus));
      if (plannedLeaveApprovalIssue) {
        addIssue(issueMap, person.employeeId, date, plannedLeaveApprovalReason(rawPlanStatus, leaveApplicationStatus));
      }
      const issues = approvedLeaveResolvesMissing ? [] : (issueMap.get(`${person.employeeId}|${day}`) || []);
      const evidenceResolvesMissing = evidence && WORK_CLOCK_CODES.has(planStatus);
      daily[day] = {
        date,
        planStatus,
        rawPlanStatus,
        approvedLeaveStatus,
        leaveApplicationStatus,
        plannedLeaveApprovalIssue,
        attendance: finalAttendance,
        evidence,
        display: dailyDisplay(planStatus, finalAttendance),
        dayoffExcess: dayoffExcessDateSet.has(date),
        substituteShortage: substituteShortageDateSet.has(date),
        compensationShortage: compensationShortageDateSet.has(date),
        // 근무 계획의 단순 미입력은 O로 해소하지만, 휴무·연차 등 계획 상이는 유지합니다.
        issues: evidenceResolvesMissing ? [] : [...new Set(issues)],
      };
    }
    dailyByKey.set(person.key, daily);
  }

  return { workforce, workforceById, planById, attendanceByKey, summaryById, issueMap, approvedLeaveStatusByKey, leaveApplicationStatusByKey, people, dailyByKey };
}


function buildLeaveApplicationStatusMap(applications = [], targetMonth = "") {
  const map = new Map();
  for (const item of applications || []) {
    const employeeId = normalizeId(item.employeeId);
    const date = String(item.leaveDate || item.date || "");
    if (!employeeId || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
    if (targetMonth && !date.startsWith(targetMonth)) continue;
    const displayStatus = approvedLeaveDisplayStatus(item);
    const approvalStatus = leaveApprovalStatus(item);
    const key = `${employeeId}|${date}`;
    const current = map.get(key);
    const candidate = { displayStatus, approvalStatus, source: item };
    if (!current || leaveApplicationStatusPriority(candidate) < leaveApplicationStatusPriority(current)) map.set(key, candidate);
  }
  return map;
}

function leaveApprovalStatus(item = {}) {
  const status = normalizeTextKey(item.applicationStatus || item.status);
  if (!status) return "신청없음";
  if (status.includes("반려") || status.includes("취소") || status.includes("철회")) return "반려";
  if (status.includes("승인") || status.includes("완료")) return "승인";
  return "미승인";
}

function leaveApplicationStatusPriority(item = {}) {
  if (item.approvalStatus === "승인") return 1;
  if (item.approvalStatus === "미승인") return 2;
  if (item.approvalStatus === "반려") return 3;
  return 9;
}

function requiresLeaveApproval(status = "") {
  const normalized = normalizePlan(status);
  return [
    "연차", "오전반차", "오후반차", "출산휴가", "육아휴직", "공가", "경조",
    "대체휴일(1일)", "대체휴일(0.5일)", "보상휴가(1일)", "보상휴가(0.5일)", "휴가",
  ].includes(normalized);
}

function plannedLeaveApprovalReason(planStatus = "", applicationStatus = null) {
  const label = normalizePlan(planStatus) || "휴가";
  const state = applicationStatus?.approvalStatus || "신청없음";
  if (state === "미승인") return `${label} 계획이나 연차신청현황 승인 전입니다`;
  if (state === "반려") return `${label} 계획이나 연차신청현황 반려/취소/철회 상태입니다`;
  return `${label} 계획이나 연차신청현황 승인 건이 없습니다`;
}

function buildUnapprovedPlannedLeaveEvidenceRows(ctx = {}, result = {}) {
  const rows = [];
  const targetMonth = result.targetMonth || "";
  const daysInMonth = targetMonth ? new Date(Number(targetMonth.slice(0, 4)), Number(targetMonth.slice(5, 7)), 0).getDate() : 31;
  const seen = new Set((result.missingRows || []).map((row) => `${normalizeId(row.employeeId)}|${row.date}`));
  for (const person of ctx.people || []) {
    const daily = ctx.dailyByKey?.get(person.key) || {};
    for (let day = 1; day <= daysInMonth; day += 1) {
      const item = daily[day];
      if (!item?.plannedLeaveApprovalIssue) continue;
      const key = `${normalizeId(person.employeeId)}|${item.date}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const member = person.member || {};
      rows.push({
        issueType: "planned_leave_not_approved",
        missingType: "휴가 승인확인",
        route: result.route || member.route || "",
        employeeId: normalizeId(person.employeeId),
        name: person.name || member.employeeName || "",
        store: member.storeName || person.plan?.store || "",
        date: item.date,
        dateObject: item.date ? new Date(`${item.date}T00:00:00`) : null,
        planStatus: item.rawPlanStatus || item.planStatus,
        actualStatus: "",
        actualIn: "",
        changedIn: "",
        result: "휴가 승인확인 필요",
        reason: plannedLeaveApprovalReason(item.rawPlanStatus || item.planStatus, item.leaveApplicationStatus),
      });
    }
  }
  return rows;
}

function buildApprovedLeaveStatusMap(applications = [], targetMonth = "") {
  const map = new Map();
  for (const item of applications || []) {
    const employeeId = normalizeId(item.employeeId);
    const date = String(item.leaveDate || item.date || "");
    if (!employeeId || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
    if (targetMonth && !date.startsWith(targetMonth)) continue;
    if (!isApprovedLeaveApplication(item)) continue;
    const displayStatus = approvedLeaveDisplayStatus(item);
    if (!displayStatus) continue;
    const key = `${employeeId}|${date}`;
    const previous = map.get(key);
    // 같은 날짜에 여러 신청이 있으면 1일 휴가성 항목을 우선 표시합니다.
    if (!previous || leaveDisplayPriority(displayStatus) < leaveDisplayPriority(previous.displayStatus)) {
      map.set(key, { displayStatus, source: item });
    }
  }
  return map;
}

function isApprovedLeaveApplication(item = {}) {
  const status = normalizeTextKey(item.applicationStatus || item.status);
  if (!status) return false;
  if (status.includes("반려") || status.includes("취소") || status.includes("철회")) return false;
  return status.includes("승인") || status.includes("완료");
}

function approvedLeaveDisplayStatus(item = {}) {
  const rawType = normalizeTextKey(`${item.leaveType || ""} ${item.requestedKind || ""}`);
  const days = Number(item.days ?? item.requestedDays ?? 0);
  if (rawType.includes("출산")) return "출산휴가";
  if (rawType.includes("육아")) return "육아휴직";
  if (rawType.includes("공가")) return "공가";
  if (rawType.includes("경조")) return "경조";
  if (rawType.includes("대체") && (rawType.includes("0.5") || rawType.includes("반"))) return "대체휴일(0.5일)";
  if (rawType.includes("대체")) return "대체휴일(1일)";
  if (rawType.includes("보상") && (rawType.includes("0.5") || rawType.includes("반"))) return "보상휴가(0.5일)";
  if (rawType.includes("보상")) return "보상휴가(1일)";
  if (rawType.includes("오전") && rawType.includes("반차")) return "오전반차";
  if (rawType.includes("오후") && rawType.includes("반차")) return "오후반차";
  if (rawType.includes("반차") || days === 0.5) return "오전반차";
  if (rawType.includes("연차") || !rawType || rawType.includes("휴가")) return "연차";
  return item.leaveType || item.requestedKind || "연차";
}

function leaveDisplayPriority(value) {
  if (["연차", "출산휴가", "육아휴직", "공가", "경조"].includes(value)) return 1;
  if (["대체휴일(1일)", "보상휴가(1일)"].includes(value)) return 2;
  if (["오전반차", "오후반차", "대체휴일(0.5일)", "보상휴가(0.5일)"].includes(value)) return 3;
  return 9;
}

function normalizeTextKey(value) {
  return String(value ?? "").trim().replace(/\s+/g, "");
}

function resolveLeaveShortageDates(events = [], shortage = 0, recordedDates = []) {
  const explicit = [...new Set((recordedDates || []).map(String).filter(Boolean))].sort();
  if (explicit.length || Number(shortage || 0) <= 0) return explicit;
  // 구버전 결과에는 초과 일자 배열이 없으므로, 날짜순 사용분의 마지막 초과 수량부터 역산합니다.
  let remaining = roundHalf(Number(shortage || 0));
  const dates = new Set();
  const ordered = [...(events || [])]
    .filter((event) => event?.date && Number(event.days || 0) > 0)
    .sort((a, b) => String(b.date).localeCompare(String(a.date)));
  for (const event of ordered) {
    if (remaining <= 0) break;
    dates.add(String(event.date));
    remaining = roundHalf(remaining - Number(event.days || 0));
  }
  return [...dates].sort();
}



function buildHrPayrollAuditSheets(workbook, result, year, monthNo) {
  const audit = result.hrPayrollAudit || { summary: {}, settlementRows: [], dailyRows: [], annualNeedRows: [], problemRows: [] };
  const subtitle = `확인 순서: 실제 출근기록 → 출근 증빙 → 휴무 → 경조·공가 → 대체·보상 → 연차 → 최종 문제자 · 기준일 ${result.cutoffDate || ""}`;

  buildAuditTableSheet(workbook, "인사팀 급여 확정표", `${year}년 ${monthNo}월 인사팀 급여 확정표`, subtitle, audit.settlementRows || [], [
    ["지역장", (row) => row.regionalManager], ["매니저", (row) => row.manager], ["지역", (row) => row.region], ["매장명", (row) => row.store],
    ["사번", (row) => row.employeeId], ["이름", (row) => row.name], ["기준일수", (row) => row.requiredDays], ["설명완료", (row) => row.explainedDays], ["미해결", (row) => row.remainingUnexplainedDays],
    ["실제 출근일수", (row) => row.workedDays], ["출근 미달일", (row) => row.missingDays], ["휴무기준", (row) => row.baseAllowance], ["휴무사용", (row) => row.dayoffUsed], ["휴무초과", (row) => row.dayoffExcess],
    ["대체차감", (row) => row.substituteReplacement], ["보상대체", (row) => row.compensationReplacement], ["대체·보상 부족", (row) => row.replacementShortage],
    ["경조·공가", (row) => row.officialPaidDays], ["연차 승인", (row) => row.annualApprovedDays], ["연차 신청필요", (row) => row.annualNeededTotal],
    ["연차잔여", (row) => row.annualRemaining], ["급여상태", (row) => row.payrollStatus], ["정산메모", (row) => row.note],
  ], {
    summaryCards: [
      ["전체 인원", `${audit.summary?.totalPeople || 0}명`], ["급여 확정", `${audit.summary?.confirmedPeople || 0}명`], ["확인 필요", `${audit.summary?.problemPeople || 0}명`],
      ["휴무 초과", `${audit.summary?.dayoffExcessPeople || 0}명`], ["연차 필요", `${audit.summary?.annualNeedPeople || 0}명`],
    ],
    statusColumn: 21,
  });

  buildAuditTableSheet(workbook, "출근 미달자 정산", `${year}년 ${monthNo}월 출근 미달자 정산`, "출근기록이 없는 날짜만 휴무 → 경조·공가 → 대체·보상 → 연차 순서로 확인합니다.", audit.dailyRows || [], [
    ["지역장", (row) => row.regionalManager], ["매니저", (row) => row.manager], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
    ["일자", (row) => row.date], ["확인순서", (row) => row.checkOrder], ["출근상태", (row) => row.attendanceStatus], ["판정값", (row) => row.sourceStatus],
    ["연차상태", (row) => row.applicationStatus], ["필요일수", (row) => row.requestedDays], ["급여판정", (row) => row.payrollDecision],
    ["휴가판정", (row) => row.leaveDecision], ["최종상태", (row) => row.finalStatus], ["비고", (row) => row.note],
  ], { statusColumn: 13 });

  buildAuditTableSheet(workbook, "연차 사용 필요자", `${year}년 ${monthNo}월 연차 사용 필요자`, "휴무 초과 후 대체·보상으로 해결되지 않거나, 출근 미달일이 승인 휴가로 설명되지 않는 인원입니다.", audit.annualNeedRows || [], [
    ["지역장", (row) => row.regionalManager], ["매니저", (row) => row.manager], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
    ["휴무초과", (row) => row.dayoffExcess], ["대체·보상 대체", (row) => row.replacementCovered], ["대체·보상 부족", (row) => row.replacementShortage],
    ["초과분 연차필요", (row) => row.annualNeededFromDayoff], ["신청누락", (row) => row.annualMissingApplicationDays], ["승인필요", (row) => row.annualPendingDays],
    ["반려확인", (row) => row.annualRejectedDays], ["미확정", (row) => row.unresolvedDays], ["총 연차확인", (row) => row.annualNeededTotal],
    ["연차잔여", (row) => row.annualRemaining], ["급여상태", (row) => row.payrollStatus], ["비고", (row) => row.note],
  ], { statusColumn: 15 });

  buildAuditTableSheet(workbook, "최종 문제자", `${year}년 ${monthNo}월 최종 문제자`, "출근·증빙·휴무·경조/공가·대체/보상·연차 반영 후에도 급여 지급 전 확인이 필요한 인원만 남깁니다.", audit.problemRows || [], [
    ["지역장", (row) => row.regionalManager], ["매니저", (row) => row.manager], ["매장명", (row) => row.store], ["사번", (row) => row.employeeId], ["이름", (row) => row.name],
    ["실제 출근일수", (row) => row.workedDays], ["출근 미달일", (row) => row.missingDays], ["휴무초과", (row) => row.dayoffExcess],
    ["대체·보상 부족", (row) => row.replacementShortage], ["경조·공가", (row) => row.officialPaidDays], ["연차확인필요", (row) => row.annualNeededTotal],
    ["급여상태", (row) => row.payrollStatus], ["정산메모", (row) => row.note],
  ], { statusColumn: 11 });
}

function buildAuditTableSheet(workbook, sheetName, title, subtitle, rows, columns, options = {}) {
  const matrix = Array.from({ length: 7 }, () => Array(columns.length).fill(""));
  matrix[0][0] = title;
  matrix[1][0] = subtitle;
  const cards = options.summaryCards || [
    ["대상 건수", `${rows.length}건`],
    ["확인 필요", `${rows.filter((row) => String(row.payrollStatus || row.finalStatus || "").includes("필요") || String(row.payrollStatus || row.finalStatus || "").includes("확인")).length}건`],
    ["급여 확정", `${rows.filter((row) => (row.payrollStatus || row.finalStatus) === "급여확정").length}건`],
  ];
  cards.forEach(([label, value], index) => {
    const col = Math.min(index * 2, Math.max(0, columns.length - 2));
    matrix[2][col] = label;
    matrix[3][col] = value;
  });
  matrix[6] = columns.map(([header]) => header);
  for (const row of rows) matrix.push(columns.map(([, getter]) => getter(row)));

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: columns.length - 1 } },
  ];
  cards.forEach((_, index) => {
    const col = Math.min(index * 2, Math.max(0, columns.length - 2));
    if (col + 1 < columns.length) {
      sheet["!merges"].push({ s: { r: 2, c: col }, e: { r: 2, c: col + 1 } });
      sheet["!merges"].push({ s: { r: 3, c: col }, e: { r: 4, c: col + 1 } });
    }
  });
  sheet["!cols"] = columns.map(([header], index) => ({ wch: Math.max(10, Math.min(index === columns.length - 1 ? 46 : 22, String(header).length * 2 + 8)) }));
  sheet["!rows"] = matrix.map((_, index) => ({ hpt: index === 0 ? 34 : index === 1 ? 24 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 30 : 28 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  styleCellRange(sheet, 0, 0, 0, columns.length - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, columns.length - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FFF4F7FB" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF40516B" } },
    alignment: { horizontal: "right", vertical: "center" },
  });
  const cardPalettes = [
    { fill: "FFF7FAFF", border: "FF8FB7E8", font: "FF0B5CCB" },
    { fill: "FFFFF5F5", border: "FFF1A2A7", font: "FFC00000" },
    { fill: "FFFFF8EF", border: "FFF4C27A", font: "FFC55A11" },
    { fill: "FFF3F8FF", border: "FF9CC2EF", font: "FF2F75B5" },
    { fill: "FFF3FBF6", border: "FF9FD3B2", font: "FF107C41" },
  ];
  cards.forEach((_, index) => {
    const col = Math.min(index * 2, Math.max(0, columns.length - 2));
    const palette = cardPalettes[index] || cardPalettes[0];
    styleCellRange(sheet, 2, col, 4, Math.min(col + 1, columns.length - 1), {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF26364D" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder(palette.border),
    });
    styleCellRange(sheet, 3, col, 4, Math.min(col + 1, columns.length - 1), {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 19, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder(palette.border),
    });
  });
  styleCellRange(sheet, 6, 0, 6, columns.length - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FF8EA6C3"),
  });
  for (let r = 7; r < matrix.length; r += 1) {
    styleCellRange(sheet, r, 0, r, columns.length - 1, {
      fill: { patternType: "solid", fgColor: { rgb: r % 2 ? "FFFFFFFF" : "FFF9FBFD" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    const statusCol = Number.isInteger(options.statusColumn) ? options.statusColumn : -1;
    if (statusCol >= 0 && statusCol < columns.length) applyAuditStatusStyle(sheet, XLSX.utils.encode_cell({ r, c: statusCol }), matrix[r][statusCol]);
    const noteCol = columns.length - 1;
    const noteAddress = XLSX.utils.encode_cell({ r, c: noteCol });
    if (sheet[noteAddress]) sheet[noteAddress].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
  }
  setRef(sheet, matrix.length, columns.length);
  workbook.Sheets[sheetName] = sheet;
  if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);
}

function applyAuditStatusStyle(sheet, address, value) {
  if (!sheet[address]) return;
  const textValue = String(value || "");
  const palette = textValue === "급여확정"
    ? { fill: "FFD9EAD3", font: "FF107C41" }
    : textValue.includes("반려") || textValue.includes("최종")
      ? { fill: "FFF4CCCC", font: "FF9C0006" }
      : textValue.includes("승인") || textValue.includes("연차") || textValue.includes("확인")
        ? { fill: "FFFFE699", font: "FF7F6000" }
        : { fill: "FFDDEBF7", font: "FF1F4E78" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  };
}



function buildIssueSummarySheet(workbook, result, ctx, year, monthNo) {
  const sheetName = "전체 요약본";
  const regionOrder = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const groups = new Map();
  const ensureGroup = (employeeId, fallback = {}) => {
    const id = normalizeId(employeeId);
    if (!id) return null;
    if (!groups.has(id)) {
      const person = ctx.people.find((item) => item.employeeId === id);
      const member = person?.member || findMember(ctx, id, fallback.store) || {};
      groups.set(id, {
        employeeId: id,
        name: fallback.name || person?.name || member.employeeName || "",
        regionalManager: member.regionalManager || fallback.regionalManager || "",
        manager: member.manager || fallback.manager || "",
        region: member.region2 || member.region1 || fallback.region || "",
        stores: new Set([fallback.store || member.storeName || ""].filter(Boolean)),
        issues: [],
        evidence: [],
        priorities: new Set(),
      });
    }
    const group = groups.get(id);
    if (fallback.store) group.stores.add(fallback.store);
    if (!group.name && fallback.name) group.name = fallback.name;
    return group;
  };

  for (const row of result.mismatchRows || []) {
    if (row.issueType === "missing_plan" || row.resolved) continue;
    const group = ensureGroup(row.employeeId, row);
    if (!group) continue;
    const text = summarizeIssueRow(row);
    if (text && !group.issues.includes(text)) group.issues.push(text);
    const needsEvidence = row.result === "근무인데 출근기록 없음"
      || row.result === "휴무·휴가인데 출근기록 있음"
      || String(row.reason || "").includes("출근기록");
    if (needsEvidence) {
      const evidenceText = `${dayLabel(row.date)} ${row.result === "근무인데 출근기록 없음" ? "근태 미입력" : "근태 수정 증빙"}`;
      if (!group.evidence.includes(evidenceText)) group.evidence.push(evidenceText);
      group.priorities.add("긴급");
    } else if (String(row.reason || "").includes("인력·매장매칭") || String(row.reason || "").includes("사번 없음")) {
      group.priorities.add("등록 확인");
    } else {
      group.priorities.add("확인 필요");
    }
  }

  for (const summary of result.employeeSummaries || []) {
    if (summary.dayoffResolved) continue;
    const group = ensureGroup(summary.employeeId, summary);
    if (!group) continue;
    const dayoffShortage = Number(summary.dayoffReplacementShortage || 0);
    const priorShortage = Number(summary.priorDayoffReplacementShortage || 0);
    if (dayoffShortage > 0 || priorShortage > 0) {
      group.issues.push(`휴무초과 미대체 ${daysText(dayoffShortage + priorShortage)}`);
      group.priorities.add("긴급");
    }
    if (Number(summary.shortage || 0) > 0) {
      group.issues.push(`대체휴무 ${daysText(summary.shortage)} 초과 사용`);
      group.priorities.add("긴급");
    }
    if (Number(summary.compensationShortage || 0) > 0) {
      group.issues.push(`보상휴가 ${daysText(summary.compensationShortage)} 초과 사용`);
      group.priorities.add("긴급");
    }
  }

  const completedPlanKeys = new Set([
    ...(result.workflowOverrides?.dailyCompletedKeys || []),
    ...(result.workflowOverrides?.planMismatchCompletedKeys || []),
  ]);
  const resolvedMismatchKeys = new Set((result.mismatchRows || [])
    .filter((item) => item.resolved)
    .map((item) => `${normalizeId(item.employeeId)}|${item.date}`));
  for (const row of (result.referenceComparison?.rows || []).filter((item) => !item.match)) {
    const date = String(row.date || "");
    const key = `${normalizeId(row.employeeId)}|${date}`;
    // 비교 기준일 이후 값과 계획&근태 상이에서 이미 처리 완료한 일자는 전체 요약에 다시 표시하지 않습니다.
    if (/^\d{4}-\d{2}-\d{2}$/.test(date) && result.cutoffDate && date > result.cutoffDate) continue;
    if (completedPlanKeys.has(key) || resolvedMismatchKeys.has(key)) continue;
    const group = ensureGroup(row.employeeId, row);
    if (!group) continue;
    const text = `${row.date || row.comparisonType || "항목"} · 자동 ${row.generatedValue || "-"} / 비교 ${row.referenceValue || "-"}${row.reason ? ` · ${row.reason}` : ""}`;
    if (!group.issues.includes(text)) group.issues.push(text);
    group.priorities.add("확인 필요");
  }

  const rows = [...groups.values()]
    .filter((group) => group.issues.length)
    .map((group) => {
      const stores = [...group.stores].filter(Boolean);
      const priority = group.priorities.has("긴급") ? "긴급" : group.priorities.has("확인 필요") ? "확인 필요" : "등록 확인";
      const category = group.evidence.length ? "증빙 필요" : priority === "긴급" ? "초과·긴급" : priority;
      return {
        ...group,
        stores,
        issues: [...new Set(group.issues)],
        evidence: [...new Set(group.evidence)],
        priority,
        category,
        status: priority === "긴급" ? "미처리" : "확인중",
        dashboardRegion: normalizeDashboardRegion(group.region, stores[0] || ""),
      };
    })
    .sort((a, b) => regionOrder.indexOf(a.dashboardRegion) - regionOrder.indexOf(b.dashboardRegion)
      || String(a.regionalManager).localeCompare(String(b.regionalManager), "ko")
      || String(a.manager).localeCompare(String(b.manager), "ko")
      || String(a.stores[0] || "").localeCompare(String(b.stores[0] || ""), "ko")
      || String(a.name).localeCompare(String(b.name), "ko"));

  const totalIssues = rows.reduce((sum, row) => sum + row.issues.length, 0);
  const evidencePeople = rows.filter((row) => row.evidence.length).length;
  const urgentPeople = rows.filter((row) => row.priority === "긴급").length;
  const normalPeople = Math.max(0, rows.length - urgentPeople);

  const matrix = Array.from({ length: 7 }, () => Array(13).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 근태관리 전체 요약`;
  matrix[1][0] = `직원별 문제를 지역별로 묶어 확인 · 기준일 ${result.cutoffDate || `${year}-${String(monthNo).padStart(2, "0")}-${String(new Date(year, monthNo, 0).getDate()).padStart(2, "0")}`}`;
  const cards = [
    [0, "검토 대상 인원", `${rows.length}명`],
    [2, "총 문제 건수", `${totalIssues}건`],
    [4, "증빙 필요 인원", `${evidencePeople}명`],
    [6, "긴급 확인 인원", `${urgentPeople}명`],
    [8, "일반 확인 인원", `${normalPeople}명`],
  ];
  for (const [col, label, value] of cards) {
    matrix[2][col] = label;
    matrix[3][col] = value;
  }
  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 긴급·증빙 필요";
  matrix[3][11] = "● 확인 필요";
  matrix[4][10] = "● 등록 확인";
  matrix[4][11] = "● 처리 완료";
  matrix[6] = ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "문제 건수", "구분", "문제 요약", "증빙 필요 내역", "우선순위", "처리상태"];

  const regionRows = [];
  const dataRows = [];
  let number = 1;
  for (const region of regionOrder) {
    const group = rows.filter((row) => row.dashboardRegion === region);
    const counts = {
      evidence: group.filter((row) => row.category === "증빙 필요").length,
      urgent: group.filter((row) => row.priority === "긴급").length,
      review: group.filter((row) => row.priority === "확인 필요").length,
      register: group.filter((row) => row.priority === "등록 확인").length,
    };
    const regionRowIndex = matrix.length;
    matrix.push([
      `▼  ${region} (총 ${group.length}명)`, "", "", "", "", "", "", "",
      `증빙 ${counts.evidence}명  |  긴급 ${counts.urgent}명  |  검토 ${counts.review}명  |  등록 ${counts.register}명`, "", "", "", "",
    ]);
    regionRows.push({ row: regionRowIndex, region });
    for (const row of group) {
      matrix.push([
        number++, row.regionalManager, row.manager, region, row.stores.join(", "), row.name, row.employeeId,
        row.issues.length, row.category, row.issues.join(" / "), row.evidence.join(" / "), row.priority, row.status,
      ]);
      dataRows.push({ row: matrix.length - 1, category: row.category, priority: row.priority, status: row.status });
    }
  }

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 13 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 13 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 13 } },
    { s: { r: 3, c: 12 }, e: { r: 3, c: 13 } },
    { s: { r: 4, c: 12 }, e: { r: 4, c: 13 } },
  ];
  for (const item of regionRows) {
    sheet["!merges"].push(
      { s: { r: item.row, c: 0 }, e: { r: item.row, c: 7 } },
      { s: { r: item.row, c: 8 }, e: { r: item.row, c: 13 } },
    );
  }
  sheet["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 10 }, { wch: 16 }, { wch: 48 }, { wch: 30 }, { wch: 12 }, { wch: 12 },
  ];
  sheet["!rows"] = matrix.map((_, index) => ({ hpt: index === 0 ? 34 : index === 1 ? 24 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 30 : regionRows.some((item) => item.row === index) ? 26 : 36 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  styleCellRange(sheet, 0, 0, 0, 13, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, 13, {
    fill: { patternType: "solid", fgColor: { rgb: "FFF4F7FB" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF40516B" } },
    alignment: { horizontal: "right", vertical: "center" },
  });
  const cardPalettes = [
    { fill: "FFF7FAFF", border: "FF8FB7E8", font: "FF0B5CCB" },
    { fill: "FFFFF8EF", border: "FFF4C27A", font: "FFC55A11" },
    { fill: "FFFFF5F5", border: "FFF1A2A7", font: "FFC00000" },
    { fill: "FFFCE4D6", border: "FFF4B183", font: "FF9C5700" },
    { fill: "FFF3F8FF", border: "FF9CC2EF", font: "FF2F75B5" },
  ];
  cards.forEach(([col], index) => {
    const palette = cardPalettes[index];
    styleCellRange(sheet, 2, col, 4, col + 1, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF26364D" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder(palette.border),
    });
    styleCellRange(sheet, 3, col, 4, col + 1, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 19, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder(palette.border),
    });
  });
  applyDashboardLegendStyle(sheet);
  styleCellRange(sheet, 6, 0, 6, 13, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FF8EA6C3"),
  });
  const regionPalette = {
    서울: { fill: "FFEAF2FB", font: "FF0B5CCB" }, 경인: { fill: "FFF0F6FC", font: "FF2F75B5" },
    충청: { fill: "FFEFF7F1", font: "FF107C41" }, 경북: { fill: "FFF3EFF9", font: "FF7030A0" },
    경남: { fill: "FFFFF3EA", font: "FFC55A11" }, 전라: { fill: "FFEDF8FA", font: "FF008C95" },
  };
  for (const item of regionRows) {
    const palette = regionPalette[item.region];
    styleCellRange(sheet, item.row, 0, item.row, 13, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "left", vertical: "center" }, border: thinBorder("FFD6DFEA"),
    });
    const summaryAddr = XLSX.utils.encode_cell({ r: item.row, c: 8 });
    if (sheet[summaryAddr]) sheet[summaryAddr].s.alignment = { horizontal: "right", vertical: "center" };
  }
  dataRows.forEach((item, index) => {
    styleCellRange(sheet, item.row, 0, item.row, 13, {
      fill: { patternType: "solid", fgColor: { rgb: index % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFDCE3EC"),
    });
    for (const col of [9, 10]) {
      const addr = XLSX.utils.encode_cell({ r: item.row, c: col });
      if (sheet[addr]) sheet[addr].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
    }
    const categoryAddr = XLSX.utils.encode_cell({ r: item.row, c: 8 });
    const categoryPalette = item.category === "증빙 필요" || item.category === "초과·긴급"
      ? { fill: "FFFFE4E6", font: "FFC00000" }
      : item.category === "등록 확인"
        ? { fill: "FFE4F0FF", font: "FF2F75B5" }
        : { fill: "FFFFEFD9", font: "FFC55A11" };
    sheet[categoryAddr].s = { ...(sheet[categoryAddr].s || {}), fill: { patternType: "solid", fgColor: { rgb: categoryPalette.fill } }, font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: categoryPalette.font } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFFFFFFF") };
    const priorityAddr = XLSX.utils.encode_cell({ r: item.row, c: 11 });
    const priorityPalette = item.priority === "긴급" ? { fill: "FFF4CCCC", font: "FF9C0006" } : item.priority === "등록 확인" ? { fill: "FFDDEBF7", font: "FF1F4E78" } : { fill: "FFFFE699", font: "FF7F6000" };
    sheet[priorityAddr].s = { ...(sheet[priorityAddr].s || {}), fill: { patternType: "solid", fgColor: { rgb: priorityPalette.fill } }, font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: priorityPalette.font } }, alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFFFFFFF") };
    applyProcessStatusStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: 12 }), item.status);
  });

  setRef(sheet, matrix.length, 13);
  workbook.Sheets[sheetName] = sheet;
  if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);
}


function summarizeIssueRow(row) {
  const day = dayLabel(row.date);
  const plan = row.planStatus || "-";
  const actual = row.actualStatus || (row.clockStatus === "출근" ? "출근" : "-");
  if (row.result === "근무인데 출근기록 없음") {
    return plan === "공백" ? `${day} 계획 공백·근태 미입력` : `${day} 근태 미입력`;
  }
  if (row.result === "휴무·휴가인데 출근기록 있음") return `${day} ${plan}인데 출근`;
  if (String(row.reason || "").includes("인력·매장매칭") || String(row.reason || "").includes("사번 없음")) return `${day} ${row.reason}`;
  if (row.result === "계획·실제 불일치" || String(row.reason || "").includes("불일치")) return `${day} 계획 ${plan} / 실제 ${actual}`;
  return `${day} ${row.reason || row.result || "검토 필요"}`;
}

function dayLabel(date) {
  const day = Number(String(date || "").slice(-2));
  return day ? `${day}일` : "해당 월";
}

function styleSummarySheet(sheet, lastRow) {
  styleCellRange(sheet, 0, 0, 0, 11, {
    fill: { patternType: "solid", fgColor: { rgb: "FF123B72" } },
    font: { name: "맑은 고딕", sz: 16, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, 11, {
    fill: { patternType: "solid", fgColor: { rgb: "FFEAF2F8" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF365B7D" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  for (const [start, color] of [[0, "FFD9EAF7"], [2, "FFE2F0D9"], [4, "FFFFF2CC"], [6, "FFFCE4D6"]]) {
    styleCellRange(sheet, 3, start, 3, start + 1, {
      fill: { patternType: "solid", fgColor: { rgb: color } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder("FFCBD5E1"),
    });
    styleCellRange(sheet, 4, start, 4, start + 1, {
      fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 16, bold: true, color: { rgb: "FF123B72" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder("FFCBD5E1"),
    });
  }
  styleCellRange(sheet, 6, 0, 6, 11, {
    fill: { patternType: "solid", fgColor: { rgb: "FF1F4E78" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  });
  for (let r = 7; r < lastRow; r += 1) {
    styleCellRange(sheet, r, 0, r, 11, {
      fill: { patternType: "solid", fgColor: { rgb: r % 2 ? "FFFFFFFF" : "FFF7FAFC" } },
      font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFD8E0E8"),
    });
    for (const c of [8, 9]) {
      const addr = XLSX.utils.encode_cell({ r, c });
      sheet[addr].s.alignment.horizontal = "left";
    }
    const priority = String(sheet[XLSX.utils.encode_cell({ r, c: 10 })]?.v || "");
    const pAddr = XLSX.utils.encode_cell({ r, c: 10 });
    if (priority === "긴급") {
      sheet[pAddr].s.fill = { patternType: "solid", fgColor: { rgb: "FFF4CCCC" } };
      sheet[pAddr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF9C0006" } };
    } else if (priority === "확인 필요") {
      sheet[pAddr].s.fill = { patternType: "solid", fgColor: { rgb: "FFFFE699" } };
      sheet[pAddr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF7F6000" } };
    } else {
      sheet[pAddr].s.fill = { patternType: "solid", fgColor: { rgb: "FFDDEBF7" } };
      sheet[pAddr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF1F4E78" } };
    }
    sheet["!rows"][r] = { hpt: 34 };
  }
}


function buildManagerRequestSheet(workbook, result, year, monthNo) {
  const regions = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const rows = [...(result.managerRequests || [])]
    .map((r) => ({ ...r, regionGroup: normalizeDashboardRegion(r.region || "", r.store || "") }))
    .sort((a, b) => regions.indexOf(a.regionGroup) - regions.indexOf(b.regionGroup)
      || String(a.manager).localeCompare(String(b.manager), "ko")
      || String(a.name).localeCompare(String(b.name), "ko"));

  const matrix = Array.from({ length: 7 }, () => Array(13).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 매니저별 상담사 이상 근태`;
  matrix[1][0] = "복사용 수정요청 멘트를 전달한 뒤 L열에 O를 입력하면 전달 완료로 표시됩니다.";
  const cards = [
    [0, "총 요청 인원", `${rows.length}명`],
    [2, "총 문제 건수", `${rows.reduce((a, r) => a + Number(r.issueCount || 0), 0)}건`],
    [4, "대상 매니저", `${new Set(rows.map((r) => r.manager).filter(Boolean)).size}명`],
    [6, "미전달", `${rows.filter((r) => !r.delivered).length}건`],
    [8, "전달 완료", `${rows.filter((r) => r.delivered).length}건`],
  ];
  for (const [c, l, v] of cards) { matrix[2][c] = l; matrix[3][c] = v; }
  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 긴급 전달";
  matrix[3][12] = "● 확인 필요";
  matrix[4][10] = "● 등록 확인";
  matrix[4][12] = "● 전달 완료";
  matrix[6] = ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "문제 건수", "구분", "문제 요약", "복사용 수정요청 멘트", "전달체크(O 입력)", "전달상태"];

  const regionRows = [], dataRows = [];
  let no = 1;
  for (const region of regions) {
    const group = rows.filter((r) => r.regionGroup === region);
    const rr = matrix.length;
    matrix.push([`▼  ${region} (총 ${group.length}명)`, "", "", "", "", "", "", "", `미전달 ${group.filter((r) => !r.delivered).length}명  |  완료 ${group.filter((r) => r.delivered).length}명`, "", "", "", ""]);
    regionRows.push({ row: rr, region });
    for (const row of group) {
      const category = /출근기록|초과|미대체/.test(row.issueText || "") ? "긴급 전달" : /인력|사번/.test(row.issueText || "") ? "등록 확인" : "확인 필요";
      const r = matrix.length;
      matrix.push([
        no++, row.regionalManager || "", row.manager || "", region, row.store || "", row.name || "", normalizeId(row.employeeId),
        row.issueCount || 0, category, row.issueText || "", row.message || "", row.delivered ? "O" : "", row.delivered ? "전달 완료" : "미전달",
      ]);
      dataRows.push({ row: r, category, delivered: row.delivered });
    }
  }

  const sh = XLSX.utils.aoa_to_sheet(matrix);
  sh["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 12 } }, { s: { r: 3, c: 10 }, e: { r: 3, c: 11 } },
    { s: { r: 3, c: 12 }, e: { r: 3, c: 12 } }, { s: { r: 4, c: 10 }, e: { r: 4, c: 11 } }, { s: { r: 4, c: 12 }, e: { r: 4, c: 12 } },
  ];
  for (const x of regionRows) sh["!merges"].push({ s: { r: x.row, c: 0 }, e: { r: x.row, c: 7 } }, { s: { r: x.row, c: 8 }, e: { r: x.row, c: 12 } });
  sh["!cols"] = [{ wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 }, { wch: 10 }, { wch: 14 }, { wch: 42 }, { wch: 52 }, { wch: 16 }, { wch: 12 }];
  sh["!rows"] = matrix.map((_, i) => ({ hpt: i === 0 ? 34 : i === 1 ? 24 : i >= 2 && i <= 4 ? 28 : i === 5 ? 8 : i === 6 ? 32 : regionRows.some((x) => x.row === i) ? 26 : 34 }));
  sh["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sh["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  styleDashboardShell(sh, matrix.length, 13, cards, regionRows);
  applyLegendBlocks(sh, [
    { row: 3, startCol: 10, endCol: 11, fill: "FFFFF5F5", font: "FFC00000" },
    { row: 3, startCol: 12, endCol: 12, fill: "FFFFF8EF", font: "FFC55A11" },
    { row: 4, startCol: 10, endCol: 11, fill: "FFE4F0FF", font: "FF2F75B5" },
    { row: 4, startCol: 12, endCol: 12, fill: "FFF3FBF6", font: "FF107C41" },
  ]);

  dataRows.forEach((x, i) => {
    styleCellRange(sh, x.row, 0, x.row, 12, {
      fill: { patternType: "solid", fgColor: { rgb: i % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    for (const c of [9, 10]) {
      const a = XLSX.utils.encode_cell({ r: x.row, c });
      if (sh[a]) sh[a].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
    }
    applyManagerCategoryStyle(sh, XLSX.utils.encode_cell({ r: x.row, c: 8 }), x.category);
    const mark = XLSX.utils.encode_cell({ r: x.row, c: 11 });
    const status = XLSX.utils.encode_cell({ r: x.row, c: 12 });
    const excelRow = x.row + 1;
    sh[mark] = sh[mark] || { t: "s", v: x.delivered ? "O" : "" };
    sh[mark].s = {
      ...(sh[mark].s || {}),
      fill: { patternType: "solid", fgColor: { rgb: x.delivered ? "FFD9EAD3" : "FFFFF2CC" } },
      font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: x.delivered ? "FF107C41" : "FFC55A11" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder("FFDCE3EC"),
    };
    sh[status] = { t: "s", v: x.delivered ? "전달 완료" : "미전달", f: `IF(OR(UPPER(TRIM(L${excelRow}))="O",L${excelRow}="○",L${excelRow}="ㅇ"),"전달 완료","미전달")`, s: sh[status]?.s || {} };
    applyManagerDeliveryStatusStyle(sh, status, x.delivered ? "전달 완료" : "미전달");
  });
  setRef(sh, matrix.length, 13);
  workbook.Sheets["매니저별 이상 근태"] = sh;
  if (!workbook.SheetNames.includes("매니저별 이상 근태")) workbook.SheetNames.push("매니저별 이상 근태");
}


function buildWeeklyAttendanceCheckSheet(workbook, result, year, monthNo) {
  const weekly = result.weeklyAttendanceChecks || { weeks: [], rows: [], highCount: 0, lowCount: 0, partialWeeks: [] };
  const weeks = weekly.weeks || [];
  const rows = weekly.rows || [];
  const regions = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const weekStartCol = 7;
  const countCol = weekStartCol + weeks.length;
  const noteCol = countCol + 1;
  const colCount = Math.max(14, noteCol + 1);
  const matrix = Array.from({ length: 7 }, () => Array(colCount).fill(""));

  matrix[0][0] = `${year}년 ${monthNo}월 주 근태 확인자`;
  matrix[1][0] = "월~일 실제 출근기록 기준으로 주 6회 이상, 주 6회가 될 수 있는 공백 후보, 주 3회 이하 중 공백 미확인 건을 확인합니다. 월 경계 주차는 전월 월 마감 일별자료와 합산하며, 끝나지 않은 주차는 누적 중으로 표시합니다.";
  const completedWeekCount = weeks.filter((week) => !(weekly.partialWeeks || []).includes(week.label)).length;
  const cards = [
    [0, "총 확인 인원", `${rows.length}명`],
    [2, "주 6회 이상", `${Number(weekly.highCount || 0)}명`],
    [4, "주 6회 후보", `${Number(weekly.highCandidateCount || 0)}명`],
    [6, "3회 이하 공백", `${Number(weekly.lowCount || 0)}명`],
    [8, "완료/누적", `${completedWeekCount}개 / ${(weekly.partialWeeks || []).length}개`],
  ];
  for (const [col, label, value] of cards) { matrix[2][col] = label; matrix[3][col] = value; }
  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 주 6회 이상";
  matrix[3][12] = "● 주 6회 후보";
  matrix[4][10] = "● 3회 이하 공백";
  matrix[4][12] = "● 누적 중/월 경계";

  matrix[6] = ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번"];
  weeks.forEach((week, index) => {
    matrix[6][weekStartCol + index] = `${week.label}\n${shortWeekRange(week.startDate, week.endDate)}`;
  });
  matrix[6][countCol] = "이상 주차 수";
  matrix[6][noteCol] = "주차별 계획·근태 비고";

  const regionRows = [];
  const dataRows = [];
  let no = 1;
  for (const region of regions) {
    const group = rows.filter((row) => row.regionGroup === region);
    const regionRow = matrix.length;
    const high = group.filter((row) => row.weekResults.some((week) => week.category === "6회 이상")).length;
    const candidate = group.filter((row) => row.weekResults.some((week) => week.category === "6회 후보")).length;
    const low = group.filter((row) => row.weekResults.some((week) => week.category === "3회 이하 공백")).length;
    const regionData = Array(colCount).fill("");
    regionData[0] = `▼  ${region} (총 ${group.length}명)`;
    regionData[Math.min(weekStartCol, colCount - 1)] = `6회 이상 ${high}명  |  6회 후보 ${candidate}명  |  3회 이하 공백 ${low}명`;
    matrix.push(regionData);
    regionRows.push({ row: regionRow, region });

    for (const row of group) {
      const values = Array(colCount).fill("");
      values[0] = no++;
      values[1] = row.regionalManager || "";
      values[2] = row.manager || "";
      values[3] = region;
      values[4] = row.store || "";
      values[5] = row.name || "";
      values[6] = normalizeId(row.employeeId);
      weeks.forEach((week, index) => {
        const resultWeek = row.weekResults.find((item) => item.label === week.label);
        if (resultWeek?.category) values[weekStartCol + index] = `✓ ${resultWeek.category}\n(${resultWeek.attendanceCount}회)`;
      });
      values[countCol] = row.flaggedCount || 0;
      values[noteCol] = row.note || "";
      const rowIndex = matrix.length;
      matrix.push(values);
      dataRows.push({ row: rowIndex, weekResults: row.weekResults || [] });
    }
  }

  const sh = XLSX.utils.aoa_to_sheet(matrix);
  sh["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: colCount - 1 } },
  ];
  if (colCount >= 14) {
    sh["!merges"].push(
      { s: { r: 3, c: 10 }, e: { r: 3, c: 11 } },
      { s: { r: 3, c: 12 }, e: { r: 3, c: colCount - 1 } },
      { s: { r: 4, c: 10 }, e: { r: 4, c: 11 } },
      { s: { r: 4, c: 12 }, e: { r: 4, c: colCount - 1 } },
    );
  }
  for (const item of regionRows) {
    sh["!merges"].push(
      { s: { r: item.row, c: 0 }, e: { r: item.row, c: Math.min(6, colCount - 1) } },
      { s: { r: item.row, c: Math.min(7, colCount - 1) }, e: { r: item.row, c: colCount - 1 } },
    );
  }

  sh["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    ...weeks.map(() => ({ wch: 14 })), { wch: 12 }, { wch: 70 },
  ];
  while (sh["!cols"].length < colCount) sh["!cols"].push({ wch: 10 });
  sh["!rows"] = matrix.map((_, index) => ({
    hpt: index === 0 ? 34 : index === 1 ? 30 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 38 : regionRows.some((item) => item.row === index) ? 26 : 48,
  }));
  sh["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sh["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  styleDashboardShell(sh, matrix.length, colCount, cards, regionRows);
  applyLegendBlocks(sh, [
    { row: 3, startCol: 10, endCol: 11, fill: "FFF4CCCC", font: "FF9C0006" },
    { row: 3, startCol: 12, endCol: colCount - 1, fill: "FFFFE699", font: "FF9C6500" },
    { row: 4, startCol: 10, endCol: 11, fill: "FFDDEBF7", font: "FF1F4E78" },
    { row: 4, startCol: 12, endCol: colCount - 1, fill: "FFE2F0D9", font: "FF375623" },
  ]);

  // 끝나지 않은 주차 머리글은 누적 중임을 알 수 있도록 하늘색으로 표시합니다.
  weeks.forEach((week, index) => {
    if (!(weekly.partialWeeks || []).includes(week.label)) return;
    const address = XLSX.utils.encode_cell({ r: 6, c: weekStartCol + index });
    if (sh[address]) sh[address].s = {
      ...(sh[address].s || {}),
      fill: { patternType: "solid", fgColor: { rgb: "FF2F75B5" } },
      font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FF8EA6C3"),
    };
  });

  dataRows.forEach((item, index) => {
    styleCellRange(sh, item.row, 0, item.row, colCount - 1, {
      fill: { patternType: "solid", fgColor: { rgb: index % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    weeks.forEach((week, weekIndex) => {
      const resultWeek = item.weekResults.find((entry) => entry.label === week.label);
      const address = XLSX.utils.encode_cell({ r: item.row, c: weekStartCol + weekIndex });
      if (!sh[address] || !resultWeek?.category) return;
      const high = resultWeek.category === "6회 이상";
      const candidate = resultWeek.category === "6회 후보";
      sh[address].s = {
        ...(sh[address].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: high ? "FFF4CCCC" : candidate ? "FFFFE699" : "FFE2F0D9" } },
        font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: high ? "FF9C0006" : candidate ? "FF9C6500" : "FF375623" } },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: thinBorder("FFDCE3EC"),
      };
    });
    const noteAddress = XLSX.utils.encode_cell({ r: item.row, c: noteCol });
    if (sh[noteAddress]) sh[noteAddress].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
  });

  setRef(sh, matrix.length, colCount);
  workbook.Sheets["주 근태 확인자"] = sh;
  if (!workbook.SheetNames.includes("주 근태 확인자")) workbook.SheetNames.push("주 근태 확인자");
}

function shortWeekRange(startDate, endDate) {
  const parse = (value) => {
    const match = String(value || "").match(/^\d{4}-(\d{2})-(\d{2})$/);
    return match ? `${Number(match[1])}/${Number(match[2])}` : value;
  };
  return `${parse(startDate)}~${parse(endDate)}`;
}

function buildPersonnelStatusSheet(workbook, result, year, monthNo) {
  const rows = [...(result.personnelChecks || [])].sort((a, b) =>
    String(a.route || "").localeCompare(String(b.route || ""))
    || String(a.issueType || "").localeCompare(String(b.issueType || ""), "ko")
    || String(a.employeeName || "").localeCompare(String(b.employeeName || ""), "ko")
  );
  const unresolved = rows.filter((row) => !row.resolved).length;
  const completed = rows.length - unresolved;
  const resigned = rows.filter((row) => row.personnelStatus === "퇴사").length;
  const transferred = rows.filter((row) => row.personnelStatus === "경로이동").length;
  const leave = rows.filter((row) => ["육아휴직", "기타휴직"].includes(row.personnelStatus)).length;
  const matrix = Array.from({ length: 7 }, () => Array(16).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 인력현황 · 연차대장 확인 요청`;
  matrix[1][0] = "처리구분은 확인 요청 / 재직·포함 / 퇴사 / 경로이동 / 육아휴직 / 기타휴직 / 제외 중 입력 · 수정 후 출근 증빙 최종본으로 재업로드 가능";
  const cards = [
    [0, "총 확인 건수", rows.length], [3, "확인 요청", unresolved], [6, "처리 완료", completed],
    [9, "퇴사·이동", resigned + transferred], [12, "휴직", leave],
  ];
  for (const [col, label, value] of cards) {
    matrix[2][col] = label;
    matrix[3][col] = `${value}건`;
  }
  matrix[6] = [
    "No", "경로", "지역장", "매니저", "지역", "매장명", "이름", "사번",
    "확인유형", "인력현황경로", "연차대장경로", "처리구분", "적용일", "종료일·복직일", "이동경로", "비고",
  ];
  rows.forEach((row, index) => matrix.push([
    index + 1, row.routeLabel || routeLabel(row.route), row.regionalManager || "", row.manager || "", row.region || "",
    row.storeName || row.store || "", row.employeeName || row.name || "", normalizeId(row.employeeId), row.issueType || "확인 요청",
    routeLabel(row.workforceRoute), routeLabel(row.annualRoute), row.personnelStatus || "확인 요청",
    row.effectiveFrom || "", row.effectiveTo || "", routeLabel(row.destinationRoute), row.note || "",
  ]));
  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 15 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 15 } },
    ...cards.flatMap(([col]) => [
      { s: { r: 2, c: col }, e: { r: 2, c: Math.min(col + 2, 15) } },
      { s: { r: 3, c: col }, e: { r: 4, c: Math.min(col + 2, 15) } },
    ]),
  ];
  sheet["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 15 }, { wch: 13 }, { wch: 15 }, { wch: 14 }, { wch: 34 },
  ];
  sheet["!rows"] = matrix.map((_, index) => ({ hpt: index === 0 ? 34 : index === 1 ? 28 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 30 : 24 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];
  styleCellRange(sheet, 0, 0, 0, 15, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, 15, {
    fill: { patternType: "solid", fgColor: { rgb: "FFF4F7FB" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF40516B" } },
    alignment: { horizontal: "right", vertical: "center", wrapText: true },
  });
  const palettes = [
    ["FFF7FAFF", "FF0B5CCB"], ["FFFFF8EF", "FFC55A11"], ["FFF3FBF6", "FF107C41"],
    ["FFFFF5F5", "FFC00000"], ["FFF3EFF9", "FF7030A0"],
  ];
  cards.forEach(([col], index) => {
    styleCellRange(sheet, 2, col, 4, Math.min(col + 2, 15), {
      fill: { patternType: "solid", fgColor: { rgb: palettes[index][0] } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF26364D" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFC8D5E5"),
    });
    styleCellRange(sheet, 3, col, 4, Math.min(col + 2, 15), {
      fill: { patternType: "solid", fgColor: { rgb: palettes[index][0] } },
      font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: palettes[index][1] } },
      alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFC8D5E5"),
    });
  });
  styleCellRange(sheet, 6, 0, 6, 15, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FF8EA6C3"),
  });
  rows.forEach((row, index) => {
    const r = 7 + index;
    styleCellRange(sheet, r, 0, r, 15, {
      fill: { patternType: "solid", fgColor: { rgb: row.resolved ? "FFF0F9F4" : index % 2 ? "FFFFFBF2" : "FFFFF8E8" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFDCE3EC"),
    });
    const statusCell = sheet[XLSX.utils.encode_cell({ r, c: 11 })];
    if (statusCell) statusCell.s = {
      ...(statusCell.s || {}),
      fill: { patternType: "solid", fgColor: { rgb: row.resolved ? "FFD9EAD3" : "FFFFE699" } },
      font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: row.resolved ? "FF107C41" : "FFC55A11" } },
      alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFDCE3EC"),
    };
    const noteCell = sheet[XLSX.utils.encode_cell({ r, c: 15 })];
    if (noteCell) noteCell.s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
  });
  setRef(sheet, matrix.length, 16);
  workbook.Sheets["인력 변동 확인"] = sheet;
  if (!workbook.SheetNames.includes("인력 변동 확인")) workbook.SheetNames.push("인력 변동 확인");
}

function routeLabel(route) {
  if (route === "homeplus") return "홈플러스";
  if (route === "electroland") return "전자랜드";
  return route || "";
}

function buildAnnualComparisonSheet(workbook, result, year, monthNo) {
  const comparison = result.annualComparison || { rows: [], matchCount: 0, reviewCount: 0, supplied: false };
  const regions = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const rows = [...(comparison.rows || [])]
    .map((r) => ({ ...r, regionGroup: normalizeDashboardRegion(r.region || "", r.store || "") }))
    .sort((a, b) => Number(a.sortOrder || 9) - Number(b.sortOrder || 9)
      || regions.indexOf(a.regionGroup) - regions.indexOf(b.regionGroup)
      || String(a.date).localeCompare(String(b.date))
      || String(a.name || "").localeCompare(String(b.name || ""), "ko"));

  const approvalDisplay = (row) => {
    if (row.needsReview) return "확인 요청";
    const status = String(row.applicationStatus || "").replace(/\s+/g, "");
    if (status.startsWith("승인") || status.includes("완료")) return "승인 완료";
    return "승인 대기";
  };
  const isAttendanceReview = (row) => row.category === "출근 기록 확인";
  const isPlanRequestReview = (row) => ["계획·신청 다름", "계획 연차·신청 없음"].includes(row.category);

  const matrix = Array.from({ length: 7 }, () => Array(15).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 연차 등록 현황 및 일자`;
  matrix[1][0] = "계획·신청이 같으면 하늘색입니다. 하루 연차는 미출근, 오전·오후 반차는 출근기록이 있어야 정상입니다. 계획·신청 불일치는 주황색, 신청 누락·출근기록 오류는 빨간색입니다.";
  const cards = [
    [0, "전체 대조", `${rows.length}건`],
    [2, "동일", `${rows.filter((r) => !r.needsReview).length}건`],
    [4, "계획 연차·신청 없음", `${rows.filter((r) => r.category === "계획 연차·신청 없음").length}건`],
    [6, "출근 기록 확인", `${rows.filter(isAttendanceReview).length}건`],
    [8, "계획·신청 확인", `${rows.filter(isPlanRequestReview).length}건`],
  ];
  for (const [c, l, v] of cards) { matrix[2][c] = l; matrix[3][c] = v; }
  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 계획·신청 동일";
  matrix[3][13] = "● 계획·신청 확인";
  matrix[4][10] = "● 출근·신청 누락 확인";
  matrix[4][13] = "● 확인 완료";
  matrix[6] = ["No", "휴가일자", "지역장", "매니저", "지역", "매장명", "이름", "사번", "신청구분", "신청일수", "근무계획", "실제근태", "대조구분", "신청상태", "확인상태"];

  const regionRows = [], dataRows = [];
  let no = 1;
  for (const region of regions) {
    const group = rows.filter((r) => r.regionGroup === region);
    const rr = matrix.length;
    matrix.push([
      `▼  ${region} (총 ${group.length}건)`, "", "", "", "", "", "", "",
      `동일 ${group.filter((r) => !r.needsReview).length}건  |  확인요청 ${group.filter((r) => r.needsReview).length}건`, "", "", "", "", "", "",
    ]);
    regionRows.push({ row: rr, region });
    for (const row of group) {
      const r = matrix.length;
      const confirmationStatus = approvalDisplay(row);
      matrix.push([
        no++, row.date || "", row.regionalManager || "", row.manager || "", region, row.store || "", row.name || "", normalizeId(row.employeeId),
        row.requestedKind || "-", row.requestedDays || 0, row.planStatus || "공백", row.actualStatus || "미출근",
        row.category || row.result || "", row.applicationStatus || "-", confirmationStatus,
      ]);
      dataRows.push({ row: r, needsReview: row.needsReview, category: row.category, confirmationStatus });
    }
  }

  const sh = XLSX.utils.aoa_to_sheet(matrix);
  sh["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 14 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 14 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 14 } },
    { s: { r: 3, c: 10 }, e: { r: 3, c: 12 } }, { s: { r: 3, c: 13 }, e: { r: 3, c: 14 } },
    { s: { r: 4, c: 10 }, e: { r: 4, c: 12 } }, { s: { r: 4, c: 13 }, e: { r: 4, c: 14 } },
  ];
  for (const x of regionRows) sh["!merges"].push({ s: { r: x.row, c: 0 }, e: { r: x.row, c: 7 } }, { s: { r: x.row, c: 8 }, e: { r: x.row, c: 14 } });
  sh["!cols"] = [{ wch: 6 }, { wch: 13 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 }, { wch: 12 }, { wch: 10 }, { wch: 17 }, { wch: 12 }, { wch: 28 }, { wch: 12 }, { wch: 13 }];
  sh["!rows"] = matrix.map((_, i) => ({ hpt: i === 0 ? 34 : i === 1 ? 26 : i >= 2 && i <= 4 ? 28 : i === 5 ? 8 : i === 6 ? 32 : regionRows.some((x) => x.row === i) ? 26 : 25 }));
  sh["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sh["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  styleDashboardShell(sh, matrix.length, 15, cards, regionRows);
  applyLegendBlocks(sh, [
    { row: 3, startCol: 10, endCol: 12, fill: "FFDDEBF7", font: "FF1F4E78" },
    { row: 3, startCol: 13, endCol: 14, fill: "FFFFE699", font: "FF9C6500" },
    { row: 4, startCol: 10, endCol: 12, fill: "FFF4CCCC", font: "FF9C0006" },
    { row: 4, startCol: 13, endCol: 14, fill: "FFE2F0D9", font: "FF107C41" },
  ]);

  dataRows.forEach((x, i) => {
    styleCellRange(sh, x.row, 0, x.row, 14, {
      fill: { patternType: "solid", fgColor: { rgb: i % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    // 예시 시안처럼 이름(G열)과 대조구분(M열)을 같은 색 계열로 표시합니다.
    for (const col of [6, 12]) applyAnnualCategoryStyle(sh, XLSX.utils.encode_cell({ r: x.row, c: col }), x.category, x.needsReview);
    applyAnnualApprovalStatusStyle(sh, XLSX.utils.encode_cell({ r: x.row, c: 14 }), x.confirmationStatus);
  });

  setRef(sh, matrix.length, 15);
  workbook.Sheets["해당 월 연차 등록 현황 및 일자"] = sh;
  if (!workbook.SheetNames.includes("해당 월 연차 등록 현황 및 일자")) workbook.SheetNames.push("해당 월 연차 등록 현황 및 일자");
}

function isResolvedByApprovedLeave(row = {}, ctx = {}) {
  const employeeId = normalizeId(row.employeeId);
  const date = String(row.date || row.leaveDate || "");
  if (!employeeId || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
  const approved = ctx.approvedLeaveStatusByKey?.get(`${employeeId}|${date}`);
  if (!approved) return false;
  const attendance = ctx.attendanceByKey?.get(`${employeeId}|${date}`) || emptyAttendance();
  // 실제 출근기록이 있으면 어차피 증빙 대상이 아니며, 승인 휴가로 미출근 사유가 설명되는 건은 증빙 시트에서 제외합니다.
  return !attendance.hasClockIn;
}

function buildEvidenceDashboardSheet(workbook, result, ctx, year, monthNo) {
  const regionOrder = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const actionSpecs = [
    { title: "출근확인(O 입력)", short: "출근", status: "09:00" },
    { title: "휴무확인(O 입력)", short: "휴무", status: "휴무" },
    { title: "연차확인(O 입력)", short: "연차", status: "연차" },
    { title: "오전반차(O 입력)", short: "오전반차", status: "오전반차" },
    { title: "오후반차(O 입력)", short: "오후반차", status: "오후반차" },
    { title: "출산휴가(O 입력)", short: "출산", status: "출산휴가" },
    { title: "육아휴직(O 입력)", short: "육아", status: "육아휴직" },
    { title: "공가(O 입력)", short: "공가", status: "공가" },
    { title: "경조(O 입력)", short: "경조", status: "경조" },
    { title: "대체휴무 1일(O 입력)", short: "대체1", status: "대체휴일(1일)" },
    { title: "대체휴무 0.5일(O 입력)", short: "대체0.5", status: "대체휴일(0.5일)" },
    { title: "보상휴가 1일(O 입력)", short: "보상1", status: "보상휴가(1일)" },
    { title: "보상휴가 0.5일(O 입력)", short: "보상0.5", status: "보상휴가(0.5일)" },
  ];
  const ACTION_START_COL = 10; // K
  const ACTION_END_COL = ACTION_START_COL + actionSpecs.length - 1; // W
  const REASON_COL = ACTION_END_COL + 1; // X
  const STATUS_COL = ACTION_END_COL + 2; // Y
  const COL_COUNT = STATUS_COL + 1;
  const actionLetters = actionSpecs.map((_, index) => XLSX.utils.encode_col(ACTION_START_COL + index));
  const doneFormulaForRange = (range) => `(COUNTIF(${range},"O")+COUNTIF(${range},"○")+COUNTIF(${range},"ㅇ"))`;
  const doneFormulaForRow = (excelRow) => doneFormulaForRange(`$K${excelRow}:$W${excelRow}`);

  const rows = [
    ...(result.missingRows || []),
    ...buildUnapprovedPlannedLeaveEvidenceRows(ctx, result),
  ]
    .filter((row) => !isResolvedByApprovedLeave(row, ctx))
    .map((row) => {
      const member = findMember(ctx, row.employeeId, row.store) || {};
      return {
        ...row,
        member,
        region: normalizeDashboardRegion(member.region2 || member.region1 || row.region || "", member.storeName || row.store || ""),
      };
    });

  rows.sort((a, b) => regionOrder.indexOf(a.region) - regionOrder.indexOf(b.region)
    || String(a.name || a.member.employeeName || "").localeCompare(String(b.name || b.member.employeeName || ""), "ko")
    || String(a.member.storeName || a.store || "").localeCompare(String(b.member.storeName || b.store || ""), "ko")
    || String(a.date || "").localeCompare(String(b.date || "")));

  const uniquePeople = new Set(rows.map((row) => normalizeId(row.employeeId)).filter(Boolean)).size;
  const uniqueStores = new Set(rows.map((row) => row.member.storeName || row.store || "").filter(Boolean)).size;
  const matrix = Array.from({ length: 7 }, () => Array(COL_COUNT).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 출근증빙·휴무확인`;
  matrix[1][0] = `출근 기록이 없거나 승인되지 않은 휴가 계획, 수기 보정이 필요한 건을 처리합니다. K~W열 중 해당 처리 항목에 O 입력 후 재업로드하면 상담사근태 해당 날짜가 출근/휴무/연차/반차/출산휴가/육아휴직/공가/경조/대체/보상으로 반영됩니다. 수기 입력 영역도 동일한 목록으로 처리됩니다. · 기준일 ${result.cutoffDate || `${year}-${String(monthNo).padStart(2, "0")}-${String(new Date(year, monthNo, 0).getDate()).padStart(2, "0")}`}`;

  const cards = [
    [0, "총 확인 건수", `${rows.length}건`],
    [2, "대상 인원", `${uniquePeople}명`],
    [4, "대상 매장", `${uniqueStores}개`],
    [6, "미처리", `${rows.length}건`],
    [8, "처리 완료", "0건"],
  ];
  for (const [col, label, value] of cards) {
    matrix[2][col] = label;
    matrix[3][col] = value;
  }
  matrix[2][10] = "처리 목록";
  matrix[3][10] = "출근 / 휴무 / 연차 / 반차";
  matrix[3][14] = "출산휴가 / 육아휴직";
  matrix[4][10] = "공가 / 경조 / 대체 / 보상";
  matrix[4][14] = "0.5일은 별도 칸 사용";
  matrix[6] = ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "발생일", "근무계획", "구분", ...actionSpecs.map((item) => item.title), "상세사유", "처리상태"];

  const regionRows = [];
  const dataRows = [];
  let number = 1;
  for (const region of regionOrder) {
    const group = rows.filter((row) => row.region === region);
    const regionRowIndex = matrix.length;
    const typeCounts = {
      clock: group.filter((row) => evidenceMissingType(row) === "출근 미입력").length,
      plan: group.filter((row) => evidenceMissingType(row) === "계획 미입력").length,
      both: group.filter((row) => evidenceMissingType(row) === "출ㆍ계 미입력").length,
    };
    matrix.push([
      `▼  ${region} (총 ${group.length}건)`, "", "", "", "", "", "", "",
      `출근 미입력 ${typeCounts.clock}건  |  계획 미입력 ${typeCounts.plan}건  |  출ㆍ계 미입력 ${typeCounts.both}건  |  완료 0건`, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
    ]);
    regionRows.push({ row: regionRowIndex, region, count: group.length, typeCounts });
    for (const row of group) {
      const missingType = evidenceMissingType(row);
      matrix.push([
        number++,
        row.member.regionalManager || "",
        row.member.manager || "",
        region,
        row.member.storeName || row.store || "",
        row.name || row.member.employeeName || "",
        normalizeId(row.employeeId),
        row.date ? new Date(`${row.date}T00:00:00`) : "",
        row.planStatus && row.planStatus !== "공백" ? row.planStatus : "미입력",
        missingType,
        ...actionSpecs.map(() => ""),
        row.reason || row.result || "출근 기록 없음 · 휴무/출근/연차/휴직 등 처리 항목 확인 필요",
        "미처리",
      ]);
      dataRows.push({ row: matrix.length - 1, missingType });
    }
  }

  const manualTitleRow = matrix.length;
  matrix.push(["▼ 수기 입력 영역", "", "", "", "", "", "", "", "", "", "이름·발생일 입력 후 K~W 중 처리 항목에 O 입력", "", "", "", "", "", "", "", "", "", "", "", "", "휴무/연차/공백인데 출근했거나 연차↔휴무 변경 등 모든 수기 보정 가능", ""]);
  const manualRows = [];
  for (let i = 0; i < 40; i += 1) {
    const rowIndex = matrix.length;
    matrix.push([`수기${String(i + 1).padStart(2, "0")}`, "", "", "", "", "", "", "", "수기입력", "수기 보정", ...actionSpecs.map(() => ""), "", "미처리"]);
    manualRows.push(rowIndex);
  }

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: STATUS_COL } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: STATUS_COL } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: STATUS_COL } },
    { s: { r: 3, c: 10 }, e: { r: 3, c: 13 } },
    { s: { r: 3, c: 14 }, e: { r: 3, c: STATUS_COL } },
    { s: { r: 4, c: 10 }, e: { r: 4, c: 13 } },
    { s: { r: 4, c: 14 }, e: { r: 4, c: STATUS_COL } },
  ];
  sheet["!merges"].push(
    { s: { r: manualTitleRow, c: 0 }, e: { r: manualTitleRow, c: 9 } },
    { s: { r: manualTitleRow, c: 10 }, e: { r: manualTitleRow, c: 22 } },
    { s: { r: manualTitleRow, c: 23 }, e: { r: manualTitleRow, c: 24 } },
  );
  for (const item of regionRows) {
    sheet["!merges"].push(
      { s: { r: item.row, c: 0 }, e: { r: item.row, c: 7 } },
      { s: { r: item.row, c: 8 }, e: { r: item.row, c: STATUS_COL } },
    );
  }

  sheet["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 13 }, { wch: 15 }, { wch: 17 },
    ...actionSpecs.map(() => ({ wch: 13 })),
    { wch: 46 }, { wch: 12 },
  ];
  sheet["!rows"] = matrix.map((_, index) => ({ hpt: index === 0 ? 34 : index === 1 ? 34 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 34 : regionRows.some((item) => item.row === index) ? 26 : 24 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 65, zoomScaleNormal: 65 }];

  const rangeStart = 8;
  const rangeEnd = Math.max(rangeStart, matrix.length);
  const doneAllFormula = doneFormulaForRange(`$K$${rangeStart}:$W$${rangeEnd}`);
  setFormula(sheet, "A4", `COUNTIF($G$${rangeStart}:$G$${rangeEnd},"?*")&"건"`, `${rows.length}건`);
  setValue(sheet, "C4", `${uniquePeople}명`);
  setValue(sheet, "E4", `${uniqueStores}개`);
  setFormula(sheet, "G4", `(COUNTIF($G$${rangeStart}:$G$${rangeEnd},"?*")-${doneAllFormula})&"건"`, `${rows.length}건`);
  setFormula(sheet, "I4", `${doneAllFormula}&"건"`, "0건");

  for (const item of regionRows) {
    const excelRow = item.row + 1;
    const totalFormula = `COUNTIF($D$${rangeStart}:$D$${rangeEnd},"${item.region}")`;
    const doneFormula = actionLetters.map((letter) => `(COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$${letter}$${rangeStart}:$${letter}$${rangeEnd},"O")+COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$${letter}$${rangeStart}:$${letter}$${rangeEnd},"○")+COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$${letter}$${rangeStart}:$${letter}$${rangeEnd},"ㅇ"))`).join("+");
    const clockFormula = `COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$J$${rangeStart}:$J$${rangeEnd},"출근 미입력")`;
    const planFormula = `COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$J$${rangeStart}:$J$${rangeEnd},"계획 미입력")`;
    const bothFormula = `COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$J$${rangeStart}:$J$${rangeEnd},"출ㆍ계 미입력")`;
    setFormula(sheet, `A${excelRow}`, `"▼  ${item.region} (총 "&${totalFormula}&"건)"`, `▼  ${item.region} (총 ${item.count}건)`);
    setFormula(sheet, `I${excelRow}`, `"출근 미입력 "&${clockFormula}&"건  |  계획 미입력 "&${planFormula}&"건  |  출ㆍ계 미입력 "&${bothFormula}&"건  |  완료 "&(${doneFormula || "0"})&"건"`, `출근 미입력 ${item.typeCounts.clock}건  |  계획 미입력 ${item.typeCounts.plan}건  |  출ㆍ계 미입력 ${item.typeCounts.both}건  |  완료 0건`);
  }

  const applyRowFormulas = (rowIndex) => {
    const excelRow = rowIndex + 1;
    setFormula(sheet, `${XLSX.utils.encode_col(STATUS_COL)}${excelRow}`, `IF(${doneFormulaForRow(excelRow)}>0,"처리 완료","미처리")`, "미처리");
    setNumberFormat(sheet, `H${excelRow}`, "yyyy-mm-dd");
  };
  dataRows.forEach((item) => applyRowFormulas(item.row));
  manualRows.forEach((rowIndex) => {
    const excelRow = rowIndex + 1;
    setFormula(sheet, `G${excelRow}`, `IFERROR(IF($F${excelRow}="","",INDEX('상담사근태'!$I:$I,MATCH($F${excelRow},'상담사근태'!$J:$J,0))),"")`, "");
    applyRowFormulas(rowIndex);
  });

  styleCellRange(sheet, 0, 0, 0, STATUS_COL, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, STATUS_COL, {
    fill: { patternType: "solid", fgColor: { rgb: "FFF4F7FB" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF40516B" } },
    alignment: { horizontal: "right", vertical: "center", wrapText: true },
  });

  const cardPalettes = [
    { fill: "FFF7FAFF", border: "FF8FB7E8", font: "FF0B5CCB" },
    { fill: "FFF3F8FF", border: "FF9CC2EF", font: "FF2F75B5" },
    { fill: "FFFFF8EF", border: "FFF4C27A", font: "FFC55A11" },
    { fill: "FFFFF5F5", border: "FFF1A2A7", font: "FFC00000" },
    { fill: "FFF3FBF6", border: "FF9FD3B2", font: "FF107C41" },
  ];
  cards.forEach(([col], index) => {
    const palette = cardPalettes[index];
    styleCellRange(sheet, 2, col, 4, col + 1, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF26364D" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder(palette.border),
    });
    styleCellRange(sheet, 3, col, 4, col + 1, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 19, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder(palette.border),
    });
  });
  applyDashboardLegendStyle(sheet);

  styleCellRange(sheet, 6, 0, 6, STATUS_COL, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FF8EA6C3"),
  });

  const regionPalette = {
    서울: { fill: "FFEAF2FB", font: "FF0B5CCB" },
    경인: { fill: "FFF0F6FC", font: "FF2F75B5" },
    충청: { fill: "FFEFF7F1", font: "FF107C41" },
    경북: { fill: "FFF3EFF9", font: "FF7030A0" },
    경남: { fill: "FFFFF3EA", font: "FFC55A11" },
    전라: { fill: "FFEDF8FA", font: "FF008C95" },
  };
  for (const item of regionRows) {
    const palette = regionPalette[item.region];
    styleCellRange(sheet, item.row, 0, item.row, STATUS_COL, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "left", vertical: "center" },
      border: thinBorder("FFD6DFEA"),
    });
    const summaryAddr = XLSX.utils.encode_cell({ r: item.row, c: 8 });
    if (sheet[summaryAddr]) sheet[summaryAddr].s.alignment = { horizontal: "right", vertical: "center" };
  }

  const inputColumns = Array.from({ length: actionSpecs.length }, (_, i) => ACTION_START_COL + i);
  const highlightInputCells = (rowIndex) => {
    for (const c of inputColumns) {
      const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
      if (sheet[addr]) sheet[addr].s = {
        ...(sheet[addr].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } },
        font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFC55A11" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder("FFF4C27A"),
      };
    }
  };
  dataRows.forEach((item, index) => {
    styleCellRange(sheet, item.row, 0, item.row, STATUS_COL, {
      fill: { patternType: "solid", fgColor: { rgb: index % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    const reasonAddr = XLSX.utils.encode_cell({ r: item.row, c: REASON_COL });
    if (sheet[reasonAddr]) sheet[reasonAddr].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
    applyMissingTypeStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: 9 }), item.missingType);
    highlightInputCells(item.row);
    applyProcessStatusStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: STATUS_COL }), "미처리");
  });

  styleCellRange(sheet, manualTitleRow, 0, manualTitleRow, STATUS_COL, {
    fill: { patternType: "solid", fgColor: { rgb: "FFE2F0D9" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF107C41" } },
    alignment: { horizontal: "left", vertical: "center", wrapText: true },
    border: thinBorder("FFB7D7A8"),
  });
  manualRows.forEach((rowIndex, manualIndex) => {
    styleCellRange(sheet, rowIndex, 0, rowIndex, STATUS_COL, {
      fill: { patternType: "solid", fgColor: { rgb: manualIndex % 2 ? "FFF7FBF5" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    for (const addr of [XLSX.utils.encode_cell({ r: rowIndex, c: 5 }), XLSX.utils.encode_cell({ r: rowIndex, c: 7 }), ...inputColumns.map((c) => XLSX.utils.encode_cell({ r: rowIndex, c }))]) {
      if (sheet[addr]) sheet[addr].s = {
        ...(sheet[addr].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } },
        font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFC55A11" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder("FFF4C27A"),
      };
    }
    applyProcessStatusStyle(sheet, XLSX.utils.encode_cell({ r: rowIndex, c: STATUS_COL }), "미처리");
  });

  setRef(sheet, matrix.length, COL_COUNT);
  workbook.Sheets["출근증빙·휴무확인"] = sheet;
  if (!workbook.SheetNames.includes("출근증빙·휴무확인")) workbook.SheetNames.push("출근증빙·휴무확인");
  delete workbook.Sheets["출근 미등록"];
  workbook.SheetNames = workbook.SheetNames.filter((name) => name !== "출근 미등록");
}

function buildPlanAttendanceMatchSheet(workbook, result, ctx, year, monthNo) {
  const regions = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const rows = (result.mismatchRows || []).filter((row) => row.issueType !== "missing_plan").map((row) => {
    const member = findMember(ctx, row.employeeId, row.store) || {};
    const category = mismatchDashboardCategory(row);
    return { ...row, member, category, regionGroup: normalizeDashboardRegion(member.region2 || member.region1 || row.region || "", member.storeName || row.store || ""), resolved: Boolean(row.resolved) };
  }).sort((a,b)=>regions.indexOf(a.regionGroup)-regions.indexOf(b.regionGroup)||String(a.date).localeCompare(String(b.date))||String(a.name).localeCompare(String(b.name),"ko"));
  const completed = rows.filter(r=>r.resolved).length;
  const matrix=Array.from({length:7},()=>Array(14).fill(""));
  matrix[0][0]=`${year}년 ${monthNo}월 계획 & 근태 상이 인원`;
  matrix[1][0]="M열에 O 입력 후 재업로드하면 같은 사번·발생일의 계획&근태 상이와 출근 미등록이 함께 제거되며, 휴무 관련 건은 휴무 초과자에도 연동됩니다.";
  const cards=[[0,"총 상이 건수",rows.length],[2,"근무인데 출근기록 없음",rows.filter(r=>r.category==="근무인데 출근기록 없음"&&!r.resolved).length],[4,"휴무·휴가인데 출근기록 있음",rows.filter(r=>r.category==="휴무·휴가인데 출근기록 있음"&&!r.resolved).length],[6,"검토 필요",rows.filter(r=>r.category==="검토 필요"&&!r.resolved).length],[8,"처리 완료",completed]];
  for(const [c,l,v] of cards){matrix[2][c]=l;matrix[3][c]=`${v}건`;}
  matrix[2][10]="구분 색상 안내";matrix[3][10]="● 근무인데 출근기록 없음";matrix[3][12]="● 휴무·휴가인데 출근기록 있음";matrix[4][10]="● 검토 필요";matrix[4][12]="● 처리 완료";
  matrix[6]=["No","지역장","매니저","지역","매장명","이름","사번","발생일","근무계획","실제근태","구분","상세사유","처리여부(O 입력)","처리상태"];
  const regionRows=[],dataRows=[];let no=1;
  for(const region of regions){const group=rows.filter(r=>r.regionGroup===region);const rr=matrix.length;matrix.push([`▼  ${region} (총 ${group.length}건)`,"","","","","","","",`미처리 ${group.filter(r=>!r.resolved).length}건  |  완료 ${group.filter(r=>r.resolved).length}건`,"","","","",""]);regionRows.push({row:rr,region});for(const row of group){const r=matrix.length;matrix.push([no++,row.member.regionalManager||"",row.member.manager||"",region,row.member.storeName||row.store||"",row.name||row.member.employeeName||"",normalizeId(row.employeeId),row.date||"",row.planStatus||"공백",row.actualStatus||(row.clockStatus&&row.clockStatus!=="미기록"?"출근":"미입력"),row.category,row.reason||"",row.resolved?"O":"",row.resolved?"처리 완료":(row.category==="근무인데 출근기록 없음"?"미처리":"확인중")]);dataRows.push({row:r,category:row.category,resolved:row.resolved});}}
  const sh=XLSX.utils.aoa_to_sheet(matrix);sh["!merges"]=[{s:{r:0,c:0},e:{r:0,c:13}},{s:{r:1,c:0},e:{r:1,c:13}},{s:{r:2,c:0},e:{r:2,c:1}},{s:{r:3,c:0},e:{r:4,c:1}},{s:{r:2,c:2},e:{r:2,c:3}},{s:{r:3,c:2},e:{r:4,c:3}},{s:{r:2,c:4},e:{r:2,c:5}},{s:{r:3,c:4},e:{r:4,c:5}},{s:{r:2,c:6},e:{r:2,c:7}},{s:{r:3,c:6},e:{r:4,c:7}},{s:{r:2,c:8},e:{r:2,c:9}},{s:{r:3,c:8},e:{r:4,c:9}},{s:{r:2,c:10},e:{r:2,c:13}},{s:{r:3,c:10},e:{r:3,c:11}},{s:{r:3,c:12},e:{r:3,c:13}},{s:{r:4,c:10},e:{r:4,c:11}},{s:{r:4,c:12},e:{r:4,c:13}}];for(const x of regionRows){sh["!merges"].push({s:{r:x.row,c:0},e:{r:x.row,c:7}},{s:{r:x.row,c:8},e:{r:x.row,c:13}})}
  sh["!cols"]=[{wch:6},{wch:11},{wch:11},{wch:9},{wch:18},{wch:11},{wch:13},{wch:13},{wch:17},{wch:15},{wch:26},{wch:44},{wch:16},{wch:12}];sh["!rows"]=matrix.map((_,i)=>({hpt:i===0?34:i===1?24:i>=2&&i<=4?28:i===5?8:i===6?30:regionRows.some(x=>x.row===i)?26:25}));sh["!freeze"]={xSplit:0,ySplit:7,topLeftCell:"A8",activePane:"bottomLeft",state:"frozen"};sh["!views"]=[{showGridLines:false,zoomScale:70,zoomScaleNormal:70}];
  styleDashboardShell(sh,matrix.length,14,cards,regionRows);
  applyLegendBlocks(sh, [
    { row: 3, startCol: 10, endCol: 11, fill: "FFFFF5F5", font: "FFC00000" },
    { row: 3, startCol: 12, endCol: 13, fill: "FFFFF8EF", font: "FFC55A11" },
    { row: 4, startCol: 10, endCol: 11, fill: "FFF3F8FF", font: "FF2F75B5" },
    { row: 4, startCol: 12, endCol: 13, fill: "FFF3FBF6", font: "FF107C41" },
  ]);
  dataRows.forEach((x,i)=>{styleCellRange(sh,x.row,0,x.row,13,{fill:{patternType:"solid",fgColor:{rgb:i%2?"FFF9FBFD":"FFFFFFFF"}},font:{name:"맑은 고딕",sz:9,color:{rgb:"FF1F2937"}},alignment:{horizontal:"center",vertical:"center",wrapText:true},border:thinBorder("FFDCE3EC")});applyMismatchBadgeStyle(sh,XLSX.utils.encode_cell({r:x.row,c:10}),x.category);const mark=XLSX.utils.encode_cell({r:x.row,c:12}),status=XLSX.utils.encode_cell({r:x.row,c:13}),excelRow=x.row+1;sh[mark]=sh[mark]||{t:"s",v:x.resolved?"O":""};sh[status]={t:"s",v:x.resolved?"처리 완료":"미처리",f:`IF(OR(UPPER(TRIM(M${excelRow}))="O",M${excelRow}="○",M${excelRow}="ㅇ"),"처리 완료","미처리")`,s:sh[status]?.s||{}};applyProcessStatusStyle(sh,status,x.resolved?"처리 완료":"미처리");sh[mark].s={...(sh[mark].s||{}),fill:{patternType:"solid",fgColor:{rgb:x.resolved?"FFD9EAD3":"FFFFF2CC"}},font:{name:"맑은 고딕",sz:9,bold:true,color:{rgb:x.resolved?"FF107C41":"FFC55A11"}},alignment:{horizontal:"center",vertical:"center"},border:thinBorder("FFDCE3EC")};});
  setRef(sh,matrix.length,14);workbook.Sheets["계획&근태 상이 인원"]=sh;if(!workbook.SheetNames.includes("계획&근태 상이 인원"))workbook.SheetNames.push("계획&근태 상이 인원");
}

function mismatchDashboardCategory(row) {
  if (row?.result === "근무인데 출근기록 없음") return "근무인데 출근기록 없음";
  if (row?.result === "휴무·휴가인데 출근기록 있음") return "휴무·휴가인데 출근기록 있음";
  return "검토 필요";
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


function applyLegendBlocks(sheet, blocks = []) {
  for (const block of blocks) {
    styleCellRange(sheet, block.row, block.startCol, block.row, block.endCol, {
      fill: { patternType: "solid", fgColor: { rgb: block.fill } },
      font: { name: "맑은 고딕", sz: 8.5, bold: true, color: { rgb: block.font } },
      alignment: { horizontal: "left", vertical: "center", wrapText: true },
      border: thinBorder("FFD6DFEA"),
    });
  }
}

function applyDashboardLegendStyle(sheet) {
  styleCellRange(sheet, 2, 10, 2, 12, {
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: "FF2E3B52" } },
    alignment: { horizontal: "left", vertical: "center" },
    border: thinBorder("FFB8C5D6"),
  });
  const items = [
    { row: 3, startCol: 10, endCol: 10, fill: "FFFFF5F5", font: "FFC00000" },
    { row: 3, startCol: 11, endCol: 12, fill: "FFFFF8EF", font: "FFC55A11" },
    { row: 4, startCol: 10, endCol: 10, fill: "FFF3F8FF", font: "FF2F75B5" },
    { row: 4, startCol: 11, endCol: 12, fill: "FFF3FBF6", font: "FF107C41" },
  ];
  for (const item of items) {
    styleCellRange(sheet, item.row, item.startCol, item.row, item.endCol, {
      fill: { patternType: "solid", fgColor: { rgb: item.fill } },
      font: { name: "맑은 고딕", sz: 8.5, bold: true, color: { rgb: item.font } },
      alignment: { horizontal: "left", vertical: "center", wrapText: true },
      border: thinBorder("FFD6DFEA"),
    });
  }
}

function applyMismatchBadgeStyle(sheet, address, category) {
  if (!sheet[address]) return;
  const palette = category === "근무인데 출근기록 없음"
    ? { fill: "FFFFE4E6", font: "FFC00000" }
    : category === "휴무·휴가인데 출근기록 있음"
      ? { fill: "FFFFEFD9", font: "FFC55A11" }
      : { fill: "FFE4F0FF", font: "FF2F75B5" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  };
}

function applyMissingTypeStyle(sheet, address, missingType) {
  if (!sheet[address]) return;
  const palette = missingType === "계획 미입력"
    ? { fill: "FFFFF2CC", font: "FFC55A11", border: "FFF4C27A" }
    : missingType === "출ㆍ계 미입력"
      ? { fill: "FFF4EAF7", font: "FF7030A0", border: "FFD4B3E2" }
      : { fill: "FFFFE4E6", font: "FFC00000", border: "FFF1A2A7" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 8, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder(palette.border),
  };
}

function applyProcessStatusStyle(sheet, address, status) {
  if (!sheet[address]) return;
  const palette = status === "처리 완료"
    ? { fill: "FFE2F0D9", font: "FF107C41" }
    : status === "미처리"
      ? { fill: "FFFFE4E6", font: "FFC00000" }
      : { fill: "FFFFF2CC", font: "FFC55A11" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center" },
    border: thinBorder("FFDCE3EC"),
  };
}

function applyManagerCategoryStyle(sheet, address, category) {
  if (!sheet[address]) return;
  const palette = category === "긴급 전달"
    ? { fill: "FFF4CCCC", font: "FF9C0006" }
    : category === "등록 확인"
      ? { fill: "FFDDEBF7", font: "FF1F4E78" }
      : { fill: "FFFFE699", font: "FF9C6500" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFDCE3EC"),
  };
}

function applyManagerDeliveryStatusStyle(sheet, address, status) {
  if (!sheet[address]) return;
  const palette = status === "전달 완료"
    ? { fill: "FFE2F0D9", font: "FF107C41" }
    : { fill: "FFFFE4E6", font: "FFC00000" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center" },
    border: thinBorder("FFDCE3EC"),
  };
}

function applyAnnualCategoryStyle(sheet, address, category, needsReview) {
  if (!sheet[address]) return;
  let palette;
  if (!needsReview || category === "동일") {
    palette = { fill: "FFDDEBF7", font: "FF1F4E78" }; // 계획·신청 동일 · 하늘색
  } else if (["계획 연차·신청 없음", "출근 기록 확인"].includes(category)) {
    palette = { fill: "FFF4CCCC", font: "FF9C0006" }; // 신청 누락 또는 출근기록 오류 · 빨간색
  } else {
    palette = { fill: "FFFFE699", font: "FF9C6500" }; // 계획·신청 불일치 · 주황색
  }
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFDCE3EC"),
  };
}

function applyAnnualApprovalStatusStyle(sheet, address, status) {
  if (!sheet[address]) return;
  const palette = status === "승인 완료"
    ? { fill: "FFDDEBF7", font: "FF1F4E78" }
    : status === "승인 대기"
      ? { fill: "FFFFE699", font: "FF9C6500" }
      : status === "확인 완료"
        ? { fill: "FFE2F0D9", font: "FF107C41" }
        : { fill: "FFF4CCCC", font: "FF9C0006" };
  sheet[address].s = {
    ...(sheet[address].s || {}),
    fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
    font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFDCE3EC"),
  };
}

function buildDayoffSubstituteSheet(workbook, result, ctx, year, monthNo) {
  const regions = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const rows = [...(result.dayoffExcessRows || [])].map((row) => {
    const member = findMember(ctx, row.employeeId, row.store) || {};
    const openingBalance = roundHalf(row.openingCarryoverTotal || 0);
    const totalGrant = roundHalf(row.combinedAvailable || 0);
    const currentGrant = roundHalf(Math.max(0, totalGrant - openingBalance));
    const basicAllowance = roundHalf(row.baseAllowance || 0);
    const dayoffUsed = roundHalf(row.basicDayoffUsed || 0);
    const explicitLeaveUsed = roundHalf(Number(row.explicitSubDayoffUsed || 0) + Number(row.compensationLeaveUsed || 0));
    const totalOffUsed = roundHalf(dayoffUsed + explicitLeaveUsed);
    const dayoffExcess = roundHalf(Math.max(0, dayoffUsed - basicAllowance));
    const substituteExcess = roundHalf(Math.max(0, explicitLeaveUsed - totalGrant));
    const remainingAfterUse = roundHalf(totalGrant - explicitLeaveUsed - dayoffExcess);
    const combinedRemaining = roundHalf(basicAllowance + totalGrant - totalOffUsed);

    let judgmentType = "normal";
    let judgment = "이상 없음";
    if (combinedRemaining < 0) {
      judgmentType = "danger";
      judgment = `휴무 초과 확인 요청 · ${daysText(Math.abs(combinedRemaining))} 부족`;
    } else if (dayoffExcess > 0) {
      judgmentType = "adjust";
      judgment = `휴무 ${daysText(dayoffExcess)} 초과 · 대체(보상) 조율 필요`;
    } else if (substituteExcess > 0) {
      judgmentType = "adjust";
      judgment = `대체(보상) ${daysText(substituteExcess)} 초과 · 휴무 조율 필요`;
    }

    return {
      ...row,
      member,
      regionGroup: normalizeDashboardRegion(member.region2 || member.region1 || row.region || "", member.storeName || row.store || ""),
      resolved: Boolean(row.dayoffResolved),
      openingBalance,
      currentGrant,
      totalGrant,
      basicAllowance,
      dayoffUsed,
      explicitLeaveUsed,
      totalOffUsed,
      dayoffExcess,
      substituteExcess,
      remainingAfterUse,
      combinedRemaining,
      judgment,
      judgmentType,
    };
  }).sort((a, b) => regions.indexOf(a.regionGroup) - regions.indexOf(b.regionGroup)
    || String(a.name).localeCompare(String(b.name), "ko"));

  const dangerCount = rows.filter((row) => row.judgmentType === "danger" && !row.resolved).length;
  const adjustCount = rows.filter((row) => row.judgmentType === "adjust" && !row.resolved).length;
  const matrix = Array.from({ length: 7 }, () => Array(19).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 기본 휴무 초과자`;
  matrix[1][0] = "이월 잔여와 당월 발생을 합산해 비교합니다. R열에 O 입력 후 재업로드하면 같은 사번·회사·월의 휴무 초과와 휴무 관련 계획&근태 상이가 함께 제거됩니다.";

  const cards = [
    [0, "총 대상 인원", `${rows.length}명`],
    [2, "휴무 초과", `${rows.filter((row) => row.dayoffExcess > 0).length}명`],
    [4, "대체 초과", `${rows.filter((row) => row.substituteExcess > 0).length}명`],
    [6, "조율 필요", `${adjustCount}명`],
    [8, "확인 요청", `${dangerCount}명`],
  ];
  for (const [col, label, value] of cards) {
    matrix[2][col] = label;
    matrix[3][col] = value;
  }

  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 이월·당월 발생/합계";
  matrix[3][14] = "● 휴무초과 확인 요청";
  matrix[4][10] = "● 조율 필요";
  matrix[4][14] = "● 이상 없음";

  matrix[6] = [
    "No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "기본휴무",
    "이월 대체(보상)잔여", "당월 대체 발생", "이월+당월 대체(보상) 합계",
    "당월 휴무 사용 개수", "휴무+대체(보상)사용 개수", "휴무 초과 개수",
    "대체 초과 개수", "당월 사용 후 잔여 대체(보상)", "판정", "처리여부(O 입력)", "처리상태",
  ];

  const regionRows = [];
  const dataRows = [];
  let no = 1;
  for (const region of regions) {
    const group = rows.filter((row) => row.regionGroup === region);
    const regionRow = matrix.length;
    matrix.push([
      `▼  ${region} (총 ${group.length}명)`, "", "", "", "", "", "", "",
      `이월 ${daysText(group.reduce((sum, row) => sum + Number(row.openingBalance || 0), 0))}  |  당월 발생 ${daysText(group.reduce((sum, row) => sum + Number(row.currentGrant || 0), 0))}  |  조율 ${group.filter((row) => row.judgmentType === "adjust" && !row.resolved).length}명  |  확인 ${group.filter((row) => row.judgmentType === "danger" && !row.resolved).length}명`,
      "", "", "", "", "", "", "", "", "", "",
    ]);
    regionRows.push({ row: regionRow, region });

    for (const row of group) {
      const r = matrix.length;
      matrix.push([
        no++, row.member.regionalManager || "", row.member.manager || "", region,
        row.member.storeName || row.store || "", row.name || row.member.employeeName || "", normalizeId(row.employeeId),
        row.basicAllowance, row.openingBalance, row.currentGrant, row.totalGrant,
        row.dayoffUsed, row.totalOffUsed, row.dayoffExcess, row.substituteExcess,
        row.remainingAfterUse, row.judgment, row.resolved ? "O" : "", row.resolved ? "처리 완료" : "미처리",
      ]);
      dataRows.push({ row: r, ...row });
    }
  }

  const sh = XLSX.utils.aoa_to_sheet(matrix);
  sh["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 18 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 18 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 18 } },
    { s: { r: 3, c: 10 }, e: { r: 3, c: 13 } }, { s: { r: 3, c: 14 }, e: { r: 3, c: 18 } },
    { s: { r: 4, c: 10 }, e: { r: 4, c: 13 } }, { s: { r: 4, c: 14 }, e: { r: 4, c: 18 } },
  ];
  for (const item of regionRows) {
    sh["!merges"].push(
      { s: { r: item.row, c: 0 }, e: { r: item.row, c: 7 } },
      { s: { r: item.row, c: 8 }, e: { r: item.row, c: 18 } },
    );
  }

  sh["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 10 }, { wch: 15 }, { wch: 13 }, { wch: 18 }, { wch: 14 }, { wch: 18 },
    { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 32 }, { wch: 16 }, { wch: 12 },
  ];
  sh["!rows"] = matrix.map((_, index) => ({
    hpt: index === 0 ? 34 : index === 1 ? 30 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 42 : regionRows.some((item) => item.row === index) ? 26 : 27,
  }));
  sh["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sh["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  styleDashboardShell(sh, matrix.length, 19, cards, regionRows);
  applyLegendBlocks(sh, [
    { row: 3, startCol: 10, endCol: 13, fill: "FFDDEBF7", font: "FF1F4E78" },
    { row: 3, startCol: 14, endCol: 18, fill: "FFF4CCCC", font: "FF9C0006" },
    { row: 4, startCol: 10, endCol: 13, fill: "FFFFE699", font: "FF9C6500" },
    { row: 4, startCol: 14, endCol: 18, fill: "FFE2F0D9", font: "FF107C41" },
  ]);

  dataRows.forEach((item, index) => {
    styleCellRange(sh, item.row, 0, item.row, 18, {
      fill: { patternType: "solid", fgColor: { rgb: index % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });

    // 이월·당월 발생·합계
    for (const col of [8, 9, 10]) {
      const address = XLSX.utils.encode_cell({ r: item.row, c: col });
      if (sh[address]) sh[address].s = {
        ...(sh[address].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: "FFDDEBF7" } },
        font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: "FF1F4E78" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder("FFDCE3EC"),
      };
    }

    // 휴무·대체 초과값
    for (const col of [13, 14]) {
      const address = XLSX.utils.encode_cell({ r: item.row, c: col });
      if (!sh[address]) continue;
      const hasExcess = Number(sh[address].v || 0) > 0;
      sh[address].s = {
        ...(sh[address].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: hasExcess ? "FFFFE699" : "FFF3FBF6" } },
        font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: hasExcess ? "FF9C6500" : "FF107C41" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder("FFDCE3EC"),
      };
    }

    const remainingAddress = XLSX.utils.encode_cell({ r: item.row, c: 15 });
    if (sh[remainingAddress]) {
      const palette = item.combinedRemaining < 0
        ? { fill: "FFF4CCCC", font: "FF9C0006" }
        : item.remainingAfterUse < 0
          ? { fill: "FFFFE699", font: "FF9C6500" }
          : { fill: "FFE2F0D9", font: "FF107C41" };
      sh[remainingAddress].s = {
        ...(sh[remainingAddress].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
        font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: palette.font } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder("FFDCE3EC"),
      };
    }

    const judge = XLSX.utils.encode_cell({ r: item.row, c: 16 });
    const judgePalette = item.judgmentType === "danger"
      ? { fill: "FFF4CCCC", font: "FF9C0006" }
      : item.judgmentType === "adjust"
        ? { fill: "FFFFE699", font: "FF9C6500" }
        : { fill: "FFE2F0D9", font: "FF107C41" };
    sh[judge].s = {
      ...(sh[judge].s || {}),
      fill: { patternType: "solid", fgColor: { rgb: judgePalette.fill } },
      font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: judgePalette.font } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    };

    const mark = XLSX.utils.encode_cell({ r: item.row, c: 17 });
    const status = XLSX.utils.encode_cell({ r: item.row, c: 18 });
    const excelRow = item.row + 1;
    sh[mark] = sh[mark] || { t: "s", v: item.resolved ? "O" : "" };
    sh[mark].s = {
      ...(sh[mark].s || {}),
      fill: { patternType: "solid", fgColor: { rgb: item.resolved ? "FFD9EAD3" : "FFFFF2CC" } },
      font: { name: "맑은 고딕", sz: 9, bold: true, color: { rgb: item.resolved ? "FF107C41" : "FFC55A11" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder("FFDCE3EC"),
    };
    sh[status] = {
      t: "s", v: item.resolved ? "처리 완료" : "미처리",
      f: `IF(OR(UPPER(TRIM(R${excelRow}))="O",R${excelRow}="○",R${excelRow}="ㅇ"),"처리 완료","미처리")`,
      s: sh[status]?.s || {},
    };
    applyProcessStatusStyle(sh, status, item.resolved ? "처리 완료" : "미처리");
  });

  setRef(sh, matrix.length, 19);
  workbook.Sheets["휴무 초과자"] = sh;
  if (!workbook.SheetNames.includes("휴무 초과자")) workbook.SheetNames.push("휴무 초과자");
}

function buildReferenceComparisonSheet(workbook, result, year, monthNo) {
  const comparison = result.referenceComparison || { rows: [], supplied: false, sameMonth: false, mismatchCount: 0, summary: "비교 파일 미선택" };
  const matrix = [
    [`${year}년 ${monthNo}월 자동 생성본 ↔ 비교 최종본`],
    [comparison.supplied ? `비교 결과: ${comparison.summary || "확인"}` : "다른 작성자의 최종본을 선택하지 않아 비교하지 않았습니다."],
    [],
    ["비교 구분", "사번", "이름", "매장명", "일자/항목", "자동 생성값", "비교 파일값", "판정", "차이 사유"],
  ];
  for (const row of comparison.rows || []) matrix.push([
    row.comparisonType || "", row.employeeId || "", row.name || "", row.store || "", row.date || "",
    row.generatedValue || "", row.referenceValue || "", row.result || "", row.reason || "",
  ]);
  if (!(comparison.rows || []).length) matrix.push(["", "", "", "", "", "", "", comparison.supplied ? "전체 일치" : "비교 파일 없음", ""]);
  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 8 } },
  ];
  sheet["!cols"] = [
    { wch: 14 }, { wch: 13 }, { wch: 11 }, { wch: 18 }, { wch: 16 },
    { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 46 },
  ];
  sheet["!rows"] = [{ hpt: 30 }, { hpt: 24 }, { hpt: 8 }, { hpt: 30 }];
  styleSimpleReportSheet(sheet, matrix.length, 9, [8]);
  for (let r = 4; r < matrix.length; r += 1) {
    const addr = XLSX.utils.encode_cell({ r, c: 7 });
    const value = String(sheet[addr]?.v || "");
    const match = value.includes("일치") && !value.includes("불일치");
    sheet[addr].s.fill = { patternType: "solid", fgColor: { rgb: match ? "FFE2F0D9" : value.includes("월") || value.includes("구조") ? "FFFFF2CC" : "FFF4CCCC" } };
    sheet[addr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: match ? "FF375623" : value.includes("월") || value.includes("구조") ? "FF7F6000" : "FF9C0006" } };
  }
  workbook.Sheets["최종본 비교"] = sheet;
  if (!workbook.SheetNames.includes("최종본 비교")) workbook.SheetNames.push("최종본 비교");
}

function styleSimpleReportSheet(sheet, lastRow, colCount, leftColumns = []) {
  styleCellRange(sheet, 0, 0, 0, colCount - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FF173F73" } },
    font: { name: "맑은 고딕", sz: 16, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, colCount - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FFEAF2F8" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF234E70" } },
    alignment: { horizontal: "left", vertical: "center", wrapText: true },
  });
  const headerRow = sheet["!merges"]?.some((merge) => merge.s.r === 3) && colCount === 13 ? 6 : 3;
  styleCellRange(sheet, headerRow, 0, headerRow, colCount - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FF1F4E78" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFFFFFFF"),
  });
  for (let r = headerRow + 1; r < lastRow; r += 1) {
    styleCellRange(sheet, r, 0, r, colCount - 1, {
      fill: { patternType: "solid", fgColor: { rgb: r % 2 ? "FFFFFFFF" : "FFF7FAFC" } },
      font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: thinBorder("FFD9E1E8"),
    });
    for (const c of leftColumns) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (sheet[addr]) sheet[addr].s.alignment.horizontal = "left";
    }
    sheet["!rows"][r] = sheet["!rows"][r] || { hpt: 28 };
  }
  sheet["!autofilter"] = { ref: `${XLSX.utils.encode_cell({ r: headerRow, c: 0 })}:${XLSX.utils.encode_cell({ r: Math.max(headerRow + 1, lastRow - 1), c: colCount - 1 })}` };
}

function fillMainSheet(sheet, result, ctx, year, monthNo, daysInMonth) {
  if (!sheet) throw new Error("상담사근태 최종본 시트를 찾지 못했습니다.");

  // 상담사근태는 상단 6행과 A:M열을 항상 고정합니다.
  // 스크롤 시작 위치는 N7이며, 최종 저장 단계에서도 같은 설정을 XML로 재확인합니다.
  sheet["!freeze"] = { xSplit: 13, ySplit: 6, topLeftCell: "N7", activePane: "bottomRight", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

  const firstDayCol0 = 14; // O열
  const lastDayCol0 = firstDayCol0 + daysInMonth - 1;
  const summaryStartCol0 = firstDayCol0 + daysInMonth;
  const summaryHeaders = [
    "총 등록 현황", "출근 등록 횟수", "휴무 가능 개수", "휴무 등록 개수", "휴무 초과 개수",
    "대체+보상 가능 개수", "대체+보상 등록 개수", "대체+보상 초과 개수", "출근 수정", "교육", "반차", "연차", "공가",
    "무급휴무", "경조사", "총 일수", "연차 미신청", "연차 신청(승인)", "출근 증빙", "수정 완료", "비고",
  ];
  const summaryEndCol0 = summaryStartCol0 + summaryHeaders.length - 1;
  const maxColCount = Math.max(64, summaryEndCol0 + 1);
  const lastCol0 = maxColCount - 1;

  removeMergesIntersecting(sheet, 1, 1, 5, lastCol0);
  clearValues(sheet, 2, 2, 4, maxColCount);
  addMerge(sheet, 1, 1, 1, 9); // B2:J2
  addMerge(sheet, 2, 1, 2, 9); // B3:J3
  addMerge(sheet, 3, 1, 3, 9); // B4:J4
  addMerge(sheet, 4, 13, 5, 13); // N5:N6
  setValue(sheet, "B2", `■${year}년 ${monthNo}월 출퇴근 현황`);
  setValue(sheet, "B3", '- 휴무 → "출근" or "연차" 수정 시, 출근 증빙자료 및 제모스 해당일자 연차신청 필수');
  setValue(sheet, "B4", '- "연차미신청"으로 수정된 날짜는 모두 제모스에 연차신청 가이드 바랍니다.');
  setValue(sheet, "N5", `${monthNo}월1일 \n근무계획\n일치확인`);

  for (let day = 1; day <= 31; day += 1) {
    const col0 = firstDayCol0 + day - 1;
    const address5 = XLSX.utils.encode_cell({ r: 4, c: col0 });
    const address6 = XLSX.utils.encode_cell({ r: 5, c: col0 });
    if (day <= daysInMonth) {
      const date = new Date(year, monthNo - 1, day);
      setValue(sheet, address5, date);
      setNumberFormat(sheet, address5, "yyyymmdd");
      setValue(sheet, address6, WEEKDAYS[date.getDay()]);
    } else {
      clearCell(sheet, address5);
      clearCell(sheet, address6);
    }
  }

  // 월의 마지막 날짜 바로 다음 열부터 요약 열을 배치합니다.
  // 30일인 6월은 AS열부터 시작하고, 31일인 달은 AT열부터 시작합니다.
  clearValues(sheet, 5, summaryStartCol0 + 1, 6, maxColCount);
  summaryHeaders.forEach((value, index) => {
    const address = XLSX.utils.encode_cell({ r: 4, c: summaryStartCol0 + index });
    setValue(sheet, address, value);
  });

  const startRow = 7;
  const lastTemplateRow = 168;
  clearValues(sheet, startRow, 2, lastTemplateRow, maxColCount);
  const seenEmployeeRows = new Map();
  const storeMoveRows = [];
  let row = startRow;
  for (const person of ctx.people) {
    copyRowStyle(sheet, 7, row, maxColCount);
    const member = person.member;
    const summary = ctx.summaryById.get(person.employeeId) || {};
    const daily = ctx.dailyByKey.get(person.key) || {};
    const values = [
      member.regionalManager, member.manager, member.region1, member.region2,
      member.storeCode, member.storeName, member.portalId, person.employeeId,
      person.name, member.hireDate, member.groupHireDate, member.note,
    ];
    values.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: index + 1 }), value));
    const priorEmployeeRowCount = seenEmployeeRows.get(person.employeeId) || 0;
    if (priorEmployeeRowCount > 0) storeMoveRows.push(row);
    seenEmployeeRows.set(person.employeeId, priorEmployeeRowCount + 1);

    setValue(sheet, `N${row}`, daily[1]?.planStatus === "공백" ? "" : daily[1]?.planStatus || "");
    let clockCorrection = 0;
    let issueCount = 0;
    for (let day = 1; day <= 31; day += 1) {
      const col0 = firstDayCol0 + day - 1;
      const address = XLSX.utils.encode_cell({ r: row - 1, c: col0 });
      if (day > daysInMonth) {
        clearCell(sheet, address);
        continue;
      }
      const item = daily[day];
      applyNeutralDayStyle(sheet, address);
      const colLetter = XLSX.utils.encode_col(col0);
      const evidenceOverrides = [
        ["K", "09:00"],
        ["L", "휴무"],
        ["M", "연차"],
        ["N", "오전반차"],
        ["O", "오후반차"],
        ["P", "출산휴가"],
        ["Q", "육아휴직"],
        ["R", "공가"],
        ["S", "경조"],
        ["T", "대체휴일(1일)"],
        ["U", "대체휴일(0.5일)"],
        ["V", "보상휴가(1일)"],
        ["W", "보상휴가(0.5일)"],
      ];
      const checkedCount = (actionCol) => `(COUNTIFS('출근증빙·휴무확인'!$G$8:$G$1000,$I${row},'출근증빙·휴무확인'!$H$8:$H$1000,${colLetter}$5,'출근증빙·휴무확인'!$${actionCol}$8:$${actionCol}$1000,"O")+COUNTIFS('출근증빙·휴무확인'!$G$8:$G$1000,$I${row},'출근증빙·휴무확인'!$H$8:$H$1000,${colLetter}$5,'출근증빙·휴무확인'!$${actionCol}$8:$${actionCol}$1000,"○")+COUNTIFS('출근증빙·휴무확인'!$G$8:$G$1000,$I${row},'출근증빙·휴무확인'!$H$8:$H$1000,${colLetter}$5,'출근증빙·휴무확인'!$${actionCol}$8:$${actionCol}$1000,"ㅇ"))`;
      const fallbackDisplay = String(item?.display || "").replace(/"/g, '""');
      const overrideFormula = evidenceOverrides.reduceRight((formula, [actionCol, display]) => `IF(${checkedCount(actionCol)}>0,"${display}",${formula})`, `"${fallbackDisplay}"`);
      setFormula(sheet, address, overrideFormula, item?.display || "");
      applyStatusStyle(sheet, address, item?.display || "");
      if (item?.issues?.length) {
        // 미입력 유형은 각각의 구분색과 8pt 글씨를 유지하고, 그 외 오류만 공통 경고색을 적용합니다.
        if (!String(item?.display || "").includes("미입력")) applyIssueStyle(sheet, address, item.issues);
        issueCount += 1;
      }
      // 반차는 실제 출근시간을 표시하되 연두색으로 구분합니다.
      if (["오전반차", "오후반차"].includes(item?.planStatus) && item?.attendance?.hasClockIn) applyHalfDayClockStyle(sheet, address);
      // 근무 외 계획인데 실제 출근한 경우는 노란색으로 표시합니다. 교육·반차는 정상 출근 범주라 제외됩니다.
      if (NON_WORK_CODES.has(item?.planStatus) && item?.attendance?.hasClockIn) applyUnexpectedClockStyle(sheet, address);
      // 휴무·대체휴무·보상휴가 초과 사용일은 다른 색상보다 주황색을 우선 적용합니다.
      // 단, 연차신청현황에서 승인된 연차/반차가 반영된 날은 일반 연차 색상을 유지합니다.
      const approvedAnnualLike = Boolean(item?.approvedLeaveStatus && !item?.attendance?.hasClockIn
        && ["연차", "오전반차", "오후반차"].includes(item?.planStatus));
      if (!approvedAnnualLike && (item?.dayoffExcess || item?.substituteShortage || item?.compensationShortage)) applyLeaveShortageStyle(sheet, address);
      if (NON_WORK_CODES.has(item?.planStatus) && item?.attendance?.hasClockIn) clockCorrection += 1;
    }

    const planValues = Object.values(daily).map((item) => item.planStatus);
    const displayValues = Object.values(daily).map((item) => item?.display || "");
    const registeredCount = displayValues.filter((value) => value && !String(value).includes("미입력")).length;
    const displayedWorkCount = Object.values(daily).filter((item) => Boolean(item?.attendance?.hasClockIn)).length;
    const plannedDayoffCount = planValues.filter((value) => value === "휴무").length;
    const displayedDayoffCount = displayValues.filter((value) => value === "휴무").length;
    const educationCount = planValues.filter((value) => value === "교육").length;
    const halfCount = planValues.filter((value) => value === "오전반차" || value === "오후반차").length;
    const annualCount = planValues.filter((value) => value === "연차").length;
    const publicCount = planValues.filter((value) => value === "공가").length;
    const unpaidCount = planValues.filter((value) => value === "무급휴가").length;
    const familyCount = planValues.filter((value) => value === "경조").length;
    const evidenceNeeded = (result.missingRows || []).some((item) => normalizeId(item.employeeId) === person.employeeId);
    const evidenceCompleted = Object.values(daily).some((item) => Boolean(item?.evidence));

    const baseAllowance = roundHalf(Number(summary.baseAllowance || 0));
    const additionalAvailable = roundHalf(Number(summary.availableSubstitute || 0) + Number(summary.availableCompensation || 0));
    const displayedDayoffExcess = roundHalf(Math.max(0, displayedDayoffCount - baseAllowance));
    const explicitSubstituteUsed = roundHalf(Number(summary.explicitSubDayoffUsed || 0));
    const autoSubstituteUsed = roundHalf(Number(summary.autoSubstituteUsed || 0));
    const totalSubstituteUsed = roundHalf(Number(summary.substituteNeeded ?? (explicitSubstituteUsed + autoSubstituteUsed)));
    const compensationUsed = roundHalf(Number(summary.compensationLeaveUsed ?? summary.compensationNeeded ?? 0));
    const combinedLeaveUsed = roundHalf(totalSubstituteUsed + compensationUsed);
    // 기본 휴무 초과와 대체·보상 초과는 서로 합산하지 않습니다.
    const combinedShortage = roundHalf(Math.max(0, combinedLeaveUsed - additionalAvailable));

    const noteParts = [];
    if (member.note) noteParts.push(member.note);
    if (plannedDayoffCount !== displayedDayoffCount) noteParts.push(`계획 휴무 ${plannedDayoffCount}일 / 최종 표시 휴무 ${displayedDayoffCount}일`);
    if (displayedDayoffExcess > 0) noteParts.push(`기본 휴무 ${compactNumber(displayedDayoffExcess)}일 초과`);
    for (let day = 1; day <= daysInMonth; day += 1) {
      if (daily[day]?.issues?.length) noteParts.push(`${monthNo}/${day} ${daily[day].issues.join("/")}`);
    }

    const firstDayCol = XLSX.utils.encode_col(firstDayCol0);
    const lastDayCol = XLSX.utils.encode_col(lastDayCol0);
    const dailyRange = `$${firstDayCol}${row}:$${lastDayCol}${row}`;
    const col = (offset) => XLSX.utils.encode_col(summaryStartCol0 + offset);
    const addr = (offset) => `${col(offset)}${row}`;

    // '출근 미입력'·'계획 미입력'·'출ㆍ계 미입력'처럼 미입력이 포함된 모든 셀은 총 등록 현황에서 제외합니다.
    setFormula(sheet, addr(0), `COUNTIFS(${dailyRange},"<>",${dailyRange},"<>*미입력*")`, registeredCount);
    setFormula(sheet, addr(1), `COUNTIF(${dailyRange},"출근")+COUNTIF(${dailyRange},"??:??")`, displayedWorkCount);
    setValue(sheet, addr(2), baseAllowance);
    setFormula(sheet, addr(3), `COUNTIF(${dailyRange},"휴무")`, displayedDayoffCount);
    setFormula(sheet, addr(4), `MAX(0,${addr(3)}-${addr(2)})`, displayedDayoffExcess);
    setValue(sheet, addr(5), additionalAvailable);
    setFormula(sheet, addr(6), `COUNTIF(${dailyRange},"대체휴일(1일)")+COUNTIF(${dailyRange},"대체휴무")+COUNTIF(${dailyRange},"보상휴가(1일)")+COUNTIF(${dailyRange},"보상휴가")+0.5*COUNTIF(${dailyRange},"대체휴일(0.5일)")+0.5*COUNTIF(${dailyRange},"보상휴가(0.5일)")+${compactNumber(autoSubstituteUsed)}`, combinedLeaveUsed);
    setFormula(sheet, addr(7), `MAX(0,${addr(6)}-${addr(5)})`, combinedShortage);
    setValue(sheet, addr(8), clockCorrection);
    setValue(sheet, addr(9), educationCount);
    setValue(sheet, addr(10), halfCount);
    setValue(sheet, addr(11), annualCount);
    setValue(sheet, addr(12), publicCount);
    setValue(sheet, addr(13), unpaidCount);
    setValue(sheet, addr(14), familyCount);
    setValue(sheet, addr(15), daysInMonth);
    setValue(sheet, addr(16), roundHalf(Number(summary.annualMissingApplication || 0)));
    setValue(sheet, addr(17), roundHalf(Number(summary.annualApproved ?? summary.currentAnnualLeave ?? 0)));
    setValue(sheet, addr(18), evidenceCompleted ? "완료" : evidenceNeeded ? "필요" : "");
    setValue(sheet, addr(19), "");
    setValue(sheet, addr(20), [...new Set(noteParts)].join(" · "));

    for (let offset = 0; offset < summaryHeaders.length; offset += 1) {
      let tone = "normal";
      if (offset === 2) tone = "dayoffAvailable";
      if (offset === 3) tone = "dayoffUsed";
      if (offset === 4) tone = displayedDayoffExcess > 0 ? "danger" : "safe";
      if (offset === 5) tone = "extraAvailable";
      if (offset === 6) tone = combinedLeaveUsed > 0 ? "extraUsed" : "extraAvailable";
      if (offset === 7) tone = combinedShortage > 0 ? "danger" : "safe";
      if (offset === 20) tone = noteParts.length ? "note" : "normal";
      applySummaryMetricStyle(sheet, addr(offset), tone, offset === 20);
    }
    row += 1;
  }
  setRef(sheet, Math.max(lastTemplateRow, row - 1), maxColCount);
  styleMainSheet(sheet, Math.max(startRow, row - 1), daysInMonth, summaryStartCol0, summaryEndCol0, maxColCount, year, monthNo);
  for (const moveRow of storeMoveRows) {
    styleCellRange(sheet, moveRow - 1, 8, moveRow - 1, 9, {
      fill: { patternType: "solid", fgColor: { rgb: "FFF4CCCC" } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFF0000" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFD9E1E8"),
    });
  }
}

function buildAnnualLedgerSheet(workbook, result, ctx, year, monthNo) {
  const dashboard = result.annualLedger || {};
  const summaries = new Map((result.employeeSummaries || []).map((row) => [normalizeId(row.employeeId), row]));
  const sourceEmployees = dashboard.employees?.length ? dashboard.employees : (result.annualRows || []).map((row) => ({
    employeeId: row.employeeId, employeeName: row.name, storeName: row.store, manager: "", regionalManager: "",
    underOneYear: row.annualUnderOneYear, grantType: row.annualGrantType, basisHireDate: "", cycleStart: row.annualCycleStart,
    cycleEnd: row.annualCycleEnd, granted: row.annualGranted, openingRemaining: row.annualOpeningRemaining,
    approvedUsed: row.cumulativeAnnualLeave, remaining: row.annualRemaining, firstPromotionDate: row.annualFirstPromotionDate,
    secondPromotionDate: row.annualSecondPromotionDate, note: "",
  }));
  const sortRows = (rows) => [...rows].sort((a, b) => String(a.regionalManager || "").localeCompare(String(b.regionalManager || ""), "ko")
    || String(a.manager || "").localeCompare(String(b.manager || ""), "ko")
    || String(a.storeName || "").localeCompare(String(b.storeName || ""), "ko")
    || String(a.employeeName || "").localeCompare(String(b.employeeName || ""), "ko"));
  const above = sortRows(sourceEmployees.filter((row) => !row.underOneYear));
  const under = sortRows(sourceEmployees.filter((row) => row.underOneYear));
  const headers = ["구분", "지역장", "매니저", "점포", "사번", "이름", "기준 입사일", "연차 사용기간", "발생", "월초 잔여", "당월 계획", "당월 승인", "연차 미신청", "누적 승인사용", "현재 잔여", "1차 촉진(6개월)", "2차 촉진(9개월)", "비고"];
  const matrix = [
    [`${year}년 ${monthNo}월 연차 누적 현황`],
    [dashboard.baseline ? `최초 기준 ${dashboard.baseline.baseline_date || ""} · ${dashboard.baseline.file_name || ""} · 월별 승인건만 누적 차감 · 같은 월 재등록 시 마지막 파일 기준` : "연차 누적 초본 미등록 · 근무계획 기준 임시 표시"],
    [], headers,
  ];
  const appendEmployee = (row, category) => {
    const id = normalizeId(row.employeeId);
    const summary = summaries.get(id) || {};
    const planned = roundHalf(Number(summary.annualPlanned || 0));
    const approved = roundHalf(Number(summary.annualApproved ?? row.approvedCurrentMonth ?? 0));
    const missing = roundHalf(Number(summary.annualMissingApplication || 0));
    const currentRemaining = summary.annualRemaining !== "" && summary.annualRemaining !== undefined
      ? roundHalf(Number(summary.annualRemaining || 0)) : roundHalf(Number(row.remaining || 0));
    matrix.push([
      category, row.regionalManager || "", row.manager || "", row.storeName || summary.store || "", id,
      row.employeeName || summary.name || "", row.basisHireDate || row.hireDate || "",
      [row.cycleStart, row.cycleEnd].filter(Boolean).join(" ~ "), roundHalf(Number(row.granted || 0)),
      roundHalf(Number(row.openingRemaining ?? row.baselineRemaining ?? 0)), planned, approved, missing,
      roundHalf(Number(summary.cumulativeAnnualLeave ?? row.approvedUsed ?? 0)), currentRemaining,
      row.firstPromotionDate || "", row.secondPromotionDate || "", row.note || "",
    ]);
  };
  above.forEach((row) => appendEmployee(row, "1년 이상·연차"));
  if (!above.length) matrix.push(["1년 이상·연차", "", "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, "", "", "대상 없음"]);
  matrix.push([]);
  const underTitleRow = matrix.length;
  matrix.push(["1년 미만자 별도 관리(경로별 월차)"]);
  matrix.push(headers);
  under.forEach((row) => appendEmployee(row, "1년 미만·월차"));
  if (!under.length) matrix.push(["1년 미만·월차", "", "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, "", "", "대상 없음"]);

  matrix.push([], []);
  const reminderTitleRow = matrix.length;
  matrix.push(["촉진 대상 관리 · 연차 생성일 기준 6개월 / 9개월 · 촉진일 1개월 경과 후 자동 제외"]);
  matrix.push(["상태", "촉진 구분", "촉진일", "지역장", "매니저", "점포", "사번", "이름", "잔여 연차", "연차 사용기간"]);
  for (const row of dashboard.reminders || []) matrix.push([
    row.status || "", row.promotionType || "", row.dueDate || "", row.regionalManager || "", row.manager || "",
    row.storeName || "", row.employeeId || "", row.employeeName || "", roundHalf(Number(row.remaining || 0)),
    [row.cycleStart, row.cycleEnd].filter(Boolean).join(" ~ "),
  ]);
  if (!(dashboard.reminders || []).length) matrix.push(["대상 없음", "", "", "", "", "", "", "", 0, ""]);

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 17 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 17 } },
    { s: { r: underTitleRow, c: 0 }, e: { r: underTitleRow, c: 17 } },
    { s: { r: reminderTitleRow, c: 0 }, e: { r: reminderTitleRow, c: 17 } },
  ];
  sheet["!cols"] = [
    {wch:16},{wch:11},{wch:11},{wch:18},{wch:13},{wch:11},{wch:13},{wch:25},{wch:9},
    {wch:11},{wch:11},{wch:11},{wch:12},{wch:13},{wch:11},{wch:15},{wch:15},{wch:38},
  ];
  sheet["!rows"] = [{hpt:30},{hpt:24},{hpt:8},{hpt:30}];
  styleCellRange(sheet, 0, 0, 0, 17, { fill:{patternType:"solid",fgColor:{rgb:"FF173F73"}}, font:{name:"맑은 고딕",sz:16,bold:true,color:{rgb:"FFFFFFFF"}}, alignment:{horizontal:"left",vertical:"center"} });
  styleCellRange(sheet, 1, 0, 1, 17, { fill:{patternType:"solid",fgColor:{rgb:"FFEAF2F8"}}, font:{name:"맑은 고딕",sz:10,color:{rgb:"FF234E70"}}, alignment:{horizontal:"left",vertical:"center",wrapText:true} });
  for (const headerRow of [3, underTitleRow + 1, reminderTitleRow + 1]) styleCellRange(sheet, headerRow, 0, headerRow, headerRow === reminderTitleRow + 1 ? 9 : 17, {
    fill:{patternType:"solid",fgColor:{rgb:"FF1F4E78"}}, font:{name:"맑은 고딕",sz:10,bold:true,color:{rgb:"FFFFFFFF"}}, alignment:{horizontal:"center",vertical:"center",wrapText:true}, border:thinBorder("FFFFFFFF")
  });
  for (const titleRow of [underTitleRow, reminderTitleRow]) styleCellRange(sheet, titleRow, 0, titleRow, 17, {
    fill:{patternType:"solid",fgColor:{rgb:"FFD9EAF7"}}, font:{name:"맑은 고딕",sz:12,bold:true,color:{rgb:"FF173F73"}}, alignment:{horizontal:"left",vertical:"center"}, border:thinBorder("FFB4C7DC")
  });
  for (let r = 4; r < matrix.length; r += 1) {
    if ([underTitleRow, underTitleRow + 1, reminderTitleRow, reminderTitleRow + 1].includes(r) || !matrix[r]?.length) continue;
    const colCount = r > reminderTitleRow + 1 ? 10 : 18;
    styleCellRange(sheet, r, 0, r, colCount - 1, { fill:{patternType:"solid",fgColor:{rgb:r%2?"FFFFFFFF":"FFF7FAFC"}}, font:{name:"맑은 고딕",sz:10,color:{rgb:"FF1F2937"}}, alignment:{horizontal:"center",vertical:"center",wrapText:true}, border:thinBorder("FFD9E1E8") });
    sheet["!rows"][r] = { hpt: 28 };
  }
  sheet["!autofilter"] = { ref: `A4:R${Math.max(5, 3 + above.length + 1)}` };
  workbook.Sheets["연차 누적 현황"] = sheet;
  if (!workbook.SheetNames.includes("연차 누적 현황")) workbook.SheetNames.push("연차 누적 현황");
}

function fillPlanSheet(sheet, result, ctx, daysInMonth) {
  if (!sheet) return;
  const header = [...(result.plan?.rawHeaders || [])];
  const maxCol = Math.max(header.length, 47);
  clearValues(sheet, 1, 1, Math.max(350, 2 + ctx.people.length * 2), maxCol);
  copyRowStyle(sheet, 1, 1, maxCol);
  for (let col = 0; col < maxCol; col += 1) setValue(sheet, XLSX.utils.encode_cell({ r: 0, c: col }), header[col] ?? "");

  const dayColumns = result.plan?.dayColumns || new Map();
  const issueMap = ctx.issueMap;
  let row = 2;
  for (const person of ctx.people) {
    copyRowStyle(sheet, 2, row, maxCol);
    const raw = person.plan ? [...person.plan.rawRow] : makeMissingPlanRow(result.plan, person);
    while (raw.length < maxCol) raw.push("");
    for (const [, dayCol] of dayColumns.entries()) {
      if (normalizePlan(raw[dayCol]) === "공백") raw[dayCol] = "";
    }
    for (let col = 0; col < maxCol; col += 1) setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: col }), raw[col] ?? "");
    for (const [, col] of dayColumns.entries()) {
      const address = XLSX.utils.encode_cell({ r: row - 1, c: col });
      applyNeutralDayStyle(sheet, address);
      applyStatusStyle(sheet, address, normalizePlan(raw[col]) === "공백" ? "" : normalizePlan(raw[col]));
    }

    const reasonValues = new Array(maxCol).fill("");
    let hasIssue = !person.plan;
    if (!person.plan) reasonValues[0] = "빨간색 체크 사유(확인 필요) · 인력·매장매칭에는 있으나 근무계획에 사번 없음";
    for (const [day, col] of dayColumns.entries()) {
      const messages = issueMap.get(`${person.employeeId}|${day}`) || [];
      if (!messages.length) continue;
      hasIssue = true;
      reasonValues[col] = [...new Set(messages)].join(" / ");
      applyIssueStyle(sheet, XLSX.utils.encode_cell({ r: row - 1, c: col }), messages);
    }
    if (hasIssue) {
      row += 1;
      copyRowStyle(sheet, 3, row, maxCol);
      reasonValues[0] = reasonValues[0] || `빨간색 체크 사유(확인 필요)${person.plan?.duplicatePlanNote ? ` · ${person.plan.duplicatePlanNote}` : ""}`;
      for (let col = 0; col < maxCol; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row - 1, c: col });
        setValue(sheet, address, reasonValues[col]);
        applyReasonStyle(sheet, address, col === 0);
      }
      if (dayColumns.size) {
        const firstDayCol = Math.min(...dayColumns.values());
        if (firstDayCol > 1) addMerge(sheet, row - 1, 0, row - 1, firstDayCol - 1);
      }
    }
    row += 1;
  }
  setRef(sheet, Math.max(150, row - 1), maxCol);
  stylePlanSheet(sheet, Math.max(2, row - 1), maxCol, dayColumns);
}

function fillAnnualSheet(sheet, result, ctx) {
  if (!sheet) return;
  clearValues(sheet, 3, 2, Math.max(156, ctx.people.length + 3), 8);
  let row = 3;
  for (const person of ctx.people) {
    copyRowStyle(sheet, 3, row, 8);
    const daily = ctx.dailyByKey.get(person.key) || {};
    const half = Object.values(daily).filter((item) => ["오전반차", "오후반차"].includes(item.planStatus)).length * 0.5;
    const annual = Object.values(daily).filter((item) => item.planStatus === "연차").length;
    const total = roundHalf(half + annual);
    const values = [person.employeeId, person.name, half, annual, total, total, 0];
    values.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: index + 1 }), value));
    row += 1;
  }
  setRef(sheet, Math.max(156, row - 1), 8);
}

function fillCompensationSheet(sheet, result, ctx) {
  if (!sheet) return;
  clearValues(sheet, 4, 2, Math.max(100, ctx.people.length * 3), 8);
  let row = 4;
  for (const person of ctx.people) {
    const daily = ctx.dailyByKey.get(person.key) || {};
    for (const item of Object.values(daily)) {
      if (!["보상휴가(0.5일)", "보상휴가(1일)"].includes(item.planStatus)) continue;
      copyRowStyle(sheet, 4, row, 8);
      const days = item.planStatus.includes("0.5") ? 0.5 : 1;
      const values = [person.employeeId, person.name, days, item.date, "O", "-", item.issues.join(" / ")];
      values.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: index + 1 }), value));
      row += 1;
    }
  }
  setRef(sheet, Math.max(31, row - 1), 8);
}

function fillEvidenceSheet(sheet, result, ctx) {
  if (!sheet) return;
  const issues = [...(result.missingRows || [])].sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.store).localeCompare(String(b.store)));
  clearValues(sheet, 5, 2, Math.max(300, issues.length + 5), 11);
  let row = 5;
  issues.forEach((issue, index) => {
    copyRowStyle(sheet, 5, row, 11);
    const member = findMember(ctx, issue.employeeId, issue.store);
    const values = [
      index + 1,
      member?.regionalManager || "",
      member?.manager || "",
      member?.region2 || member?.region1 || "",
      member?.storeName || issue.store || "",
      normalizeId(issue.employeeId),
      issue.name || member?.employeeName || "",
      issue.date ? new Date(`${issue.date}T00:00:00`) : "",
      issue.reason || issue.result || "",
      "",
    ];
    values.forEach((value, valueIndex) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: valueIndex + 1 }), value));
    row += 1;
  });
  setValue(sheet, "K3", "증빙여부(O 입력)");
  for (let evidenceRow = 5; evidenceRow < row; evidenceRow += 1) setNumberFormat(sheet, `I${evidenceRow}`, "yyyy-mm-dd");
  setRef(sheet, Math.max(25, row - 1), 11);
  styleEvidenceSheet(sheet, Math.max(5, row - 1));
}

function fillAttendanceRawSheet(sheet, result) {
  if (!sheet) return;
  const matrix = result.attendance?.matrix || [];
  const headerIndex = Number(result.attendance?.headerIndex || 0);
  const sourceRows = matrix.slice(headerIndex);
  const maxCol = Math.max(11, ...sourceRows.map((row) => row.length));
  clearValues(sheet, 1, 1, Math.max(5000, sourceRows.length + 10), maxCol);
  sourceRows.forEach((source, rowIndex) => {
    copyRowStyle(sheet, rowIndex === 0 ? 1 : 2, rowIndex + 1, maxCol);
    for (let col = 0; col < maxCol; col += 1) setValue(sheet, XLSX.utils.encode_cell({ r: rowIndex, c: col }), source[col] ?? "");
  });
  setRef(sheet, Math.max(1, sourceRows.length), maxCol);
}

function fillAnnualRawSheet(sheet, result, ctx) {
  if (!sheet) return;
  const sourceSheet = (result.annualSourceSheets || []).find((item) => Array.isArray(item.matrix) && item.matrix.length);
  if (sourceSheet) {
    const sourceRows = sourceSheet.matrix;
    const maxCol = Math.max(13, ...sourceRows.map((row) => row.length));
    clearValues(sheet, 1, 1, Math.max(1000, sourceRows.length + 10), maxCol);
    sourceRows.forEach((source, rowIndex) => {
      copyRowStyle(sheet, rowIndex === 0 ? 1 : 2, rowIndex + 1, maxCol);
      for (let col = 0; col < maxCol; col += 1) setValue(sheet, XLSX.utils.encode_cell({ r: rowIndex, c: col }), source[col] ?? "");
    });
    setRef(sheet, Math.max(1, sourceRows.length), maxCol);
    return;
  }
  clearValues(sheet, 1, 1, Math.max(800, ctx.people.length * 10), 22);
  const headers = ["사번", "이름", "일자", "구분", "환산일수", "상태", "비고"];
  headers.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: 0, c: index }), value));
  let row = 2;
  for (const person of ctx.people) {
    const daily = ctx.dailyByKey.get(person.key) || {};
    for (const item of Object.values(daily)) {
      if (!["연차", "오전반차", "오후반차"].includes(item.planStatus)) continue;
      copyRowStyle(sheet, 2, row, 22);
      const values = [person.employeeId, person.name, item.date, item.planStatus, item.planStatus === "연차" ? 1 : 0.5, "계획표 등록", item.issues.join(" / ")];
      values.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: index }), value));
      row += 1;
    }
  }
  setRef(sheet, Math.max(2, row - 1), 22);
}

function fillEducationRawSheet(sheet, result, ctx) {
  if (!sheet) return;
  clearValues(sheet, 1, 1, Math.max(800, ctx.people.length * 10), 14);
  const headers = ["사번", "이름", "일자", "계획", "출근시간", "변경출근시간", "판정", "비고"];
  headers.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: 0, c: index }), value));
  let row = 2;
  for (const person of ctx.people) {
    const daily = ctx.dailyByKey.get(person.key) || {};
    for (const item of Object.values(daily)) {
      if (item.planStatus !== "교육") continue;
      copyRowStyle(sheet, 2, row, 14);
      const values = [
        person.employeeId, person.name, item.date, item.planStatus,
        item.attendance.actualIn || "", item.attendance.changedIn || "",
        item.attendance.hasClockIn ? "출근 확인" : "출근 미입력", item.issues.join(" / "),
      ];
      values.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: index }), value));
      row += 1;
    }
  }
  setRef(sheet, Math.max(2, row - 1), 14);
}

function normalizeMember(row) {
  return {
    route: row.route || "",
    regionalManager: row.regionalManager ?? row.regional_manager ?? "",
    manager: row.manager ?? "",
    region1: row.region1 ?? "",
    region2: row.region2 ?? "",
    storeCode: row.storeCode ?? row.store_code ?? "",
    storeName: row.storeName ?? row.store_name ?? "",
    portalId: row.portalId ?? row.portal_id ?? "",
    employeeId: normalizeId(row.employeeId ?? row.employee_id),
    employeeName: row.employeeName ?? row.employee_name ?? "",
    hireDate: row.hireDate ?? row.hire_date ?? "",
    groupHireDate: row.groupHireDate ?? row.group_hire_date ?? "",
    note: row.note ?? "",
  };
}

function choosePlan(rows, member) {
  if (!rows.length) return null;
  if (rows.length === 1) return rows[0];
  const memberStore = normalizeStore(member.storeName);
  const exact = rows.find((row) => normalizeStore(row.store) === memberStore);
  if (exact) return { ...exact, duplicatePlanNote: `중복 계획 ${rows.length}건 중 매장매칭 기준으로 선택` };
  return { ...rows[0], duplicatePlanNote: `중복 계획 ${rows.length}건 중 첫 번째 행 선택` };
}

function makeMissingPlanRow(plan, person) {
  const maxCol = Math.max(plan?.rawHeaders?.length || 0, 47);
  const row = new Array(maxCol).fill("");
  const columns = plan?.columns || {};
  if (columns.store >= 0) row[columns.store] = person.member.storeName;
  if (columns.employeeId >= 0) row[columns.employeeId] = person.employeeId;
  if (columns.name >= 0) row[columns.name] = person.name;
  return row;
}

function dailyDisplay(planStatus, attendance) {
  if (attendance?.finalOverride && attendance?.finalDisplay) return String(attendance.finalDisplay);
  const planMissing = planStatus === "공백";
  const clockMissing = !attendance.hasClockIn;
  if (planMissing && clockMissing) return "출ㆍ계 미입력";
  const actual = normalizeActual(attendance.actualStatus);
  if (attendance.hasClockIn) return effectiveClockDisplay(attendance) || (actual && actual !== "근무" ? actual : "출근");
  if (actual) return actual === "근무" ? "출근" : actual;
  if (["근무", "근무A", "근무B", "근무C", "교육", "오전반차", "오후반차"].includes(planStatus)) return "출근 미입력";
  return planStatus;
}

function effectiveClockDisplay(attendance) {
  const raw = cleanClock(attendance?.actualIn) || cleanClock(attendance?.changedIn);
  const match = String(raw || "").match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (!match) return "";
  return `${String(Number(match[1])).padStart(2, "0")}:${match[2]}`;
}

function evidenceMissingType(row) {
  if (row?.missingType) return row.missingType;
  const planMissing = !row?.planStatus || row.planStatus === "공백";
  const clockMissing = !row?.actualIn && !row?.changedIn && !row?.actualStatus;
  if (planMissing && clockMissing) return "출ㆍ계 미입력";
  if (planMissing) return "계획 미입력";
  return "출근 미입력";
}

function normalizeActual(value) {
  const raw = String(value || "").replace(/\s+/g, "");
  if (!raw || raw === "546") return "";
  if (raw.includes("오전반차")) return "오전반차";
  if (raw.includes("오후반차")) return "오후반차";
  if (raw.includes("대체") && raw.includes("0.5")) return "대체휴일(0.5일)";
  if (raw.includes("대체")) return "대체휴일(1일)";
  if (raw.includes("보상") && raw.includes("0.5")) return "보상휴가(0.5일)";
  if (raw.includes("보상")) return "보상휴가(1일)";
  if (raw.includes("무급")) return "무급휴가";
  if (raw.includes("연차")) return "연차";
  if (raw.includes("공가")) return "공가";
  if (raw.includes("경조")) return "경조";
  if (raw.includes("교육")) return "교육";
  if (raw.includes("휴무")) return "휴무";
  if (raw.includes("휴가")) return "휴가";
  if (raw.includes("근무") || raw.includes("출근") || raw === "정상") return "근무";
  return raw;
}

function normalizePlan(value) {
  const raw = String(value ?? "").trim().replace(/\s+/g, "");
  if (!raw || raw === "546") return "공백";
  return raw;
}

function buildAttendanceMap(rows) {
  const map = new Map();
  for (const row of rows || []) {
    const id = normalizeId(row.employeeId);
    if (!id || !row.date) continue;
    const key = `${id}|${row.date}`;
    const existing = map.get(key) || emptyAttendance();
    const actualIn = cleanClock(row.actualIn);
    const changedIn = cleanClock(row.changedIn);
    const actualStatus = cleanPlaceholder(row.actualStatus);
    const finalOverride = Boolean(row.finalOverride);
    const forceClockIn = Boolean(row.forceClockIn);
    const finalDisplay = cleanPlaceholder(row.finalDisplay);
    const actualStatusIsWork = normalizeActual(actualStatus) === "근무";
    map.set(key, {
      // 출근시간 또는 실제근태의 출근·근무·정상 값 중 하나라도 있으면 실제 출근으로 표시합니다.
      hasClockIn: finalOverride ? forceClockIn : (existing.hasClockIn || Boolean(actualIn || changedIn) || actualStatusIsWork),
      actualIn: actualIn || existing.actualIn,
      changedIn: changedIn || existing.changedIn,
      actualStatus: actualStatus || existing.actualStatus,
      location: cleanPlaceholder(row.location) || existing.location,
      finalOverride: finalOverride || existing.finalOverride,
      forceClockIn: finalOverride ? forceClockIn : existing.forceClockIn,
      finalDisplay: finalDisplay || existing.finalDisplay,
    });
  }
  return map;
}

function emptyAttendance() {
  return { hasClockIn: false, actualIn: "", changedIn: "", actualStatus: "", location: "", finalOverride: false, forceClockIn: false, finalDisplay: "" };
}

function cleanClock(value) {
  const raw = String(value ?? "").trim();
  return !raw || raw === "546" || raw === "0" ? "" : raw;
}

function cleanPlaceholder(value) {
  const raw = String(value ?? "").trim();
  return raw === "546" ? "" : raw;
}

function findMember(ctx, employeeId, store) {
  const rows = ctx.workforceById.get(normalizeId(employeeId)) || [];
  if (rows.length <= 1) return rows[0] || null;
  const normalizedStore = normalizeStore(store);
  return rows.find((row) => normalizeStore(row.storeName) === normalizedStore) || rows[0];
}

function addIssue(map, employeeId, date, message) {
  const id = normalizeId(employeeId);
  const day = Number(String(date || "").slice(-2));
  if (!id || !day || !message) return;
  const key = `${id}|${day}`;
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(String(message));
}

function workforceSort(a, b) {
  const aa = normalizeMember(a);
  const bb = normalizeMember(b);
  return String(aa.regionalManager).localeCompare(String(bb.regionalManager), "ko")
    || String(aa.manager).localeCompare(String(bb.manager), "ko")
    || String(aa.region2).localeCompare(String(bb.region2), "ko")
    || String(aa.storeCode).localeCompare(String(bb.storeCode), "ko")
    || String(aa.employeeName).localeCompare(String(bb.employeeName), "ko");
}

function groupBy(items, selector) {
  const map = new Map();
  for (const item of items || []) {
    const key = selector(item);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  }
  return map;
}

function normalizeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}

function normalizeStore(value) {
  return String(value || "").replace(/^\d+_/, "").replace(/홈플러스|전자랜드/g, "").replace(/점$/g, "").replace(/\s+/g, "").toLowerCase();
}

function roundHalf(value) {
  return Math.round((Number(value) || 0) * 2) / 2;
}

function compactNumber(value) {
  const number = roundHalf(value);
  return Number.isInteger(number) ? String(number) : number.toFixed(1);
}

function daysText(value) {
  const number = roundHalf(value);
  return `${Number.isInteger(number) ? number : number.toFixed(1)}일`;
}

function countWeekendDays(year, monthNo) {
  const days = new Date(year, monthNo, 0).getDate();
  let count = 0;
  for (let day = 1; day <= days; day += 1) {
    const weekday = new Date(year, monthNo - 1, day).getDay();
    if (weekday === 0 || weekday === 6) count += 1;
  }
  return count;
}

function countWeekdays(year, monthNo) {
  return new Date(year, monthNo, 0).getDate() - countWeekendDays(year, monthNo);
}

function renameSheet(workbook, oldName, newName) {
  if (!workbook.Sheets[oldName] || oldName === newName) return;
  const index = workbook.SheetNames.indexOf(oldName);
  workbook.Sheets[newName] = workbook.Sheets[oldName];
  delete workbook.Sheets[oldName];
  if (index >= 0) workbook.SheetNames[index] = newName;
}

function setValue(sheet, address, value) {
  const old = sheet[address] || {};
  const style = old.s ? clone(old.s) : undefined;
  const z = old.z;
  if (value === null || value === undefined || value === "") {
    sheet[address] = { t: "s", v: "" };
  } else if (typeof value === "number") {
    sheet[address] = { t: "n", v: value };
  } else if (value instanceof Date) {
    sheet[address] = { t: "d", v: value };
  } else {
    sheet[address] = { t: "s", v: String(value) };
  }
  if (style) sheet[address].s = style;
  if (z) sheet[address].z = z;
  extendRefForAddress(sheet, address);
}

function setFormula(sheet, address, formula, cachedValue = "") {
  const old = sheet[address] || {};
  const style = old.s ? clone(old.s) : undefined;
  const z = old.z;
  sheet[address] = { t: "s", v: cachedValue, f: String(formula || "").replace(/^=/, "") };
  if (style) sheet[address].s = style;
  if (z) sheet[address].z = z;
  extendRefForAddress(sheet, address);
}

function removeMergesIntersecting(sheet, sr, sc, er, ec) {
  sheet["!merges"] = (sheet["!merges"] || []).filter((merge) => (
    merge.e.r < sr || merge.s.r > er || merge.e.c < sc || merge.s.c > ec
  ));
}

function clearCell(sheet, address) {
  const old = sheet[address] || {};
  const style = old.s ? clone(old.s) : undefined;
  sheet[address] = { t: "s", v: "" };
  if (style) sheet[address].s = style;
}

function setNumberFormat(sheet, address, format) {
  if (!sheet[address]) sheet[address] = { t: "n", v: 0 };
  sheet[address].z = format;
}

function clearValues(sheet, startRow, startCol, endRow, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      clearCell(sheet, XLSX.utils.encode_cell({ r: row - 1, c: col - 1 }));
    }
  }
}

function copyRowStyle(sheet, sourceRow, targetRow, maxCol) {
  if (sourceRow === targetRow) return;
  for (let col = 0; col < maxCol; col += 1) {
    const source = sheet[XLSX.utils.encode_cell({ r: sourceRow - 1, c: col })];
    const targetAddress = XLSX.utils.encode_cell({ r: targetRow - 1, c: col });
    if (!sheet[targetAddress]) sheet[targetAddress] = { t: "s", v: "" };
    if (source?.s) sheet[targetAddress].s = clone(source.s);
    if (source?.z) sheet[targetAddress].z = source.z;
  }
  if (sheet["!rows"]?.[sourceRow - 1]) {
    sheet["!rows"] = sheet["!rows"] || [];
    sheet["!rows"][targetRow - 1] = clone(sheet["!rows"][sourceRow - 1]);
  }
}

function applyIssueStyle(sheet, address, messages = []) {
  if (!sheet[address]) sheet[address] = { t: "s", v: "" };
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  // 계획·근태 상이, 미입력, 잔여 부족 등 확인이 필요한 일자는 한 가지 색으로 통일합니다.
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFF33E0D" } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { ...(base.alignment || {}), horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFB7C3D0"),
  };
}

function applyNeutralDayStyle(sheet, address) {
  if (!sheet[address]) sheet[address] = { t: "s", v: "" };
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: false, color: { rgb: "FF000000" } },
    alignment: { ...(base.alignment || {}), horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  };
}

function applyStatusStyle(sheet, address, value) {
  if (!sheet[address]) return;
  const text = String(value || "");
  let fill = null;
  let fontColor = "FF000000";
  let bold = false;
  if (text === "출근 미입력") { fill = "FFF33E0D"; fontColor = "FFFFFFFF"; bold = true; }
  else if (text === "계획 미입력") { fill = "FFFFC000"; fontColor = "FF7F4100"; bold = true; }
  else if (text === "출ㆍ계 미입력") { fill = "FFFF0000"; fontColor = "FFFFFFFF"; bold = true; }
  else if (text === "미입력") { fill = "FFF33E0D"; fontColor = "FFFFFFFF"; bold = true; }
  else if (text === "휴무") fill = "FFDDEBF7";
  else if (text.includes("대체휴일") || text === "대체휴무" || text.includes("보상휴가")) { fill = "FF2F5597"; fontColor = "FFFFFFFF"; bold = true; }
  else if (["연차", "오전반차", "오후반차", "반일근무"].includes(text)) fill = "FFC6E0B4";
  else if (["공가", "휴가", "경조", "무급휴가", "교육", "출산휴가", "육아휴직"].includes(text)) fill = "FFDDEBF7";
  // 정상 출근시간(예: 09:03)과 구버전 출근 표시는 기본 흰색을 유지합니다.
  if (!fill) return;
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: fill } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: text.includes("미입력") ? 8 : 10, bold, color: { rgb: fontColor } },
  };
}

function applyHalfDayClockStyle(sheet, address) {
  if (!sheet[address]) return;
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFC6E0B4" } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: false, color: { rgb: "FF000000" } },
    alignment: { ...(base.alignment || {}), horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFB7C3D0"),
  };
}

function applyUnexpectedClockStyle(sheet, address) {
  if (!sheet[address]) return;
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFFFE699" } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF7F6000" } },
    alignment: { ...(base.alignment || {}), horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFB7C3D0"),
  };
}

function applyLeaveShortageStyle(sheet, address) {
  if (!sheet[address]) return;
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFF4B183" } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF9C5700" } },
    alignment: { ...(base.alignment || {}), horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFB7C3D0"),
  };
}

function applySummaryMetricStyle(sheet, address, tone = "normal", leftAlign = false) {
  if (!sheet[address]) sheet[address] = { t: "s", v: "" };
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  const palette = {
    normal: { fill: "FFF7FAFC", font: "FF1F2937" },
    safe: { fill: "FFE2F0D9", font: "FF375623" },
    covered: { fill: "FFE2F0D9", font: "FF375623" },
    used: { fill: "FFE4DFEC", font: "FF5F497A" },
    dayoffAvailable: { fill: "FFDDEBF7", font: "FF1F4E78" },
    dayoffUsed: { fill: "FFE7E6E6", font: "FF404040" },
    extraAvailable: { fill: "FFE4DFEC", font: "FF5F497A" },
    extraUsed: { fill: "FFFCE4D6", font: "FF9E480E" },
    danger: { fill: "FFF33E0D", font: "FFFFFFFF" },
    note: { fill: "FFFFF2CC", font: "FF7F6000" },
  };
  const selected = palette[tone] || palette.normal;
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: selected.fill } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: tone === "danger", color: { rgb: selected.font } },
    alignment: { ...(base.alignment || {}), horizontal: leftAlign ? "left" : "center", vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  };
}

function styleMainSheet(sheet, lastRow, daysInMonth, summaryStartCol0, summaryEndCol0, maxColCount, year, monthNo) {
  // N5:N6 병합 셀이 필터 머리글 행(6행)과 겹치면 Excel이 파일을
  // 손상된 통합문서로 판단할 수 있습니다. 고정 인적사항 B:M에만 필터를 둡니다.
  sheet["!autofilter"] = { ref: `B6:M${lastRow}` };
  const cols = Array.from({ length: maxColCount }, () => ({ wch: 10 }));
  cols[0] = { wch: 2 };
  [1, 2, 3, 4].forEach((i) => { cols[i] = { wch: 11 }; });
  cols[5] = { wch: 12 };
  cols[6] = { wch: 16 };
  cols[7] = { wch: 15 };
  cols[8] = { wch: 12 };
  cols[9] = { wch: 11 };
  [10, 11, 12].forEach((i) => { cols[i] = { wch: 12 }; });
  cols[13] = { wch: 13 };
  for (let i = 14; i < 45 && i < maxColCount; i += 1) cols[i] = { wch: i - 13 <= daysInMonth ? 10.5 : 3 };
  for (let i = summaryStartCol0; i <= summaryEndCol0; i += 1) cols[i] = { wch: 12 };
  cols[summaryStartCol0 + 2] = { wch: 13.78 };
  cols[summaryStartCol0 + 3] = { wch: 13.78 };
  cols[summaryStartCol0 + 4] = { wch: 13.78 };
  cols[summaryStartCol0 + 5] = { wch: 10.89 };
  cols[summaryStartCol0 + 6] = { wch: 10.89 };
  cols[summaryStartCol0 + 7] = { wch: 10.89 };
  cols[summaryEndCol0] = { wch: 46 };
  sheet["!cols"] = cols;

  sheet["!rows"] = sheet["!rows"] || [];
  sheet["!rows"][1] = { hpt: 27 };
  sheet["!rows"][2] = { hpt: 27 };
  sheet["!rows"][3] = { hpt: 27 };
  sheet["!rows"][4] = { hpt: 33 };
  sheet["!rows"][5] = { hpt: 14.3 };
  for (let r = 6; r < lastRow; r += 1) sheet["!rows"][r] = { hpt: 15.3 };

  styleCellRange(sheet, 1, 1, 3, maxColCount - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 1, 1, 9, {
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FF000000" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 2, 1, 3, 9, {
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
    font: { name: "맑은 고딕", sz: 15, bold: true, color: { rgb: "FF0000FF" } },
    alignment: { horizontal: "left", vertical: "center", wrapText: true },
  });

  const navyHeader = {
    fill: { patternType: "solid", fgColor: { rgb: "FF203764" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  };
  styleCellRange(sheet, 4, 1, 4, 12, navyHeader); // B5:M5
  styleCellRange(sheet, 5, 1, 5, 12, navyHeader); // B6:M6
  styleCellRange(sheet, 4, summaryStartCol0, 5, maxColCount - 1, navyHeader);

  styleCellRange(sheet, 4, 13, lastRow - 1, 13, {
    fill: { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF000000" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  });

  const firstDayCol0 = 14;
  const lastDayCol0 = firstDayCol0 + daysInMonth - 1;
  styleCellRange(sheet, 4, firstDayCol0, 4, lastDayCol0, {
    fill: { patternType: "solid", fgColor: { rgb: "FFD9E1F2" } },
    font: { name: "맑은 고딕", sz: 10, bold: false, color: { rgb: "FF000000" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFB4C7DC"),
  });
  styleCellRange(sheet, 5, firstDayCol0, 5, lastDayCol0, navyHeader);

  for (let day = 1; day <= daysInMonth; day += 1) {
    const date = new Date(year, monthNo - 1, day);
    if (!isKoreanWeekendOrHoliday(date)) continue;
    for (const row0 of [4, 5]) {
      const address = XLSX.utils.encode_cell({ r: row0, c: firstDayCol0 + day - 1 });
      if (!sheet[address]) continue;
      const base = sheet[address].s ? clone(sheet[address].s) : {};
      sheet[address].s = {
        ...base,
        font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: Boolean(base.font?.bold), color: { rgb: "FFFF0000" } },
      };
    }
  }

  // 휴무 집계와 대체·보상 집계의 제목 색을 구분합니다.
  styleCellRange(sheet, 4, summaryStartCol0 + 2, 4, summaryStartCol0 + 4, {
    fill: { patternType: "solid", fgColor: { rgb: "FF5B9BD5" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  });
  styleCellRange(sheet, 4, summaryStartCol0 + 5, 4, summaryStartCol0 + 7, {
    fill: { patternType: "solid", fgColor: { rgb: "FF8064A2" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  });

  styleCellRange(sheet, 6, 1, lastRow - 1, 12, {
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } },
    alignment: { vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  }, true);
  styleCellRange(sheet, 6, 9, lastRow - 1, 9, {
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFF00" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  });
  styleCellRange(sheet, 6, summaryStartCol0, lastRow - 1, summaryEndCol0, {
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  }, true);
  for (let r = 6; r < lastRow; r += 1) {
    const address = XLSX.utils.encode_cell({ r, c: summaryEndCol0 });
    if (sheet[address]) sheet[address].s = {
      ...(sheet[address].s || {}),
      font: { ...((sheet[address].s || {}).font || {}), name: "맑은 고딕", sz: 10, color: { rgb: "FF000000" } },
      alignment: { horizontal: "left", vertical: "center", wrapText: true },
      border: thinBorder("FFD9E1E8"),
    };
  }
}

function isKoreanWeekendOrHoliday(date) {
  const day = date.getDay();
  if (day === 0 || day === 6) return true;
  const year = date.getFullYear();
  const key = `${year}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  const fixedHoliday = new Set(["01-01", "03-01", "05-05", "06-06", "08-15", "10-03", "10-09", "12-25"]);
  return Boolean(KOREAN_PUBLIC_HOLIDAYS[year]?.has(key) || fixedHoliday.has(key.slice(5)));
}

function stylePlanSheet(sheet, lastRow, maxCol, dayColumns) {
  const endCol = XLSX.utils.encode_col(maxCol - 1);
  sheet["!autofilter"] = { ref: `A1:${endCol}${lastRow}` };
  const cols = Array.from({ length: maxCol }, () => ({ wch: 10 }));
  cols[0] = { wch: 16 };
  cols[1] = { wch: 13 };
  cols[2] = { wch: 13 };
  cols[3] = { wch: 11 };
  [4, 5, 6, 7].forEach((i) => { cols[i] = { wch: 11 }; });
  for (const col of dayColumns.values()) cols[col] = { wch: 9 };
  sheet["!cols"] = cols;
  sheet["!rows"] = sheet["!rows"] || [];
  sheet["!rows"][0] = { hpt: 28 };
  styleCellRange(sheet, 0, 0, 0, maxCol - 1, {
    fill: { patternType: "solid", fgColor: { rgb: "FF1F4E78" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  });
  styleCellRange(sheet, 1, 0, lastRow - 1, maxCol - 1, {
    font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF222222" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFD9E1E8"),
  }, true);
}

function styleEvidenceSheet(sheet, lastRow) {
  setValue(sheet, "B1", "근태 오류·증빙 제출 대상");
  setValue(sheet, "B2", "수정이 필요한 직원과 발생일을 확인한 뒤 K열에 O를 입력하세요. 근태 미입력 건은 상담사근태 시트에서 자동으로 ‘출근’으로 반영됩니다.");
  addMerge(sheet, 0, 1, 0, 10);
  addMerge(sheet, 1, 1, 1, 10);
  sheet["!autofilter"] = { ref: `B3:K${lastRow}` };
  sheet["!cols"] = [{ wch: 2 }, { wch: 7 }, { wch: 11 }, { wch: 11 }, { wch: 10 }, { wch: 16 }, { wch: 13 }, { wch: 11 }, { wch: 13 }, { wch: 44 }, { wch: 15 }];
  styleCellRange(sheet, 0, 1, 0, 10, {
    fill: { patternType: "solid", fgColor: { rgb: "FF123B72" } },
    font: { name: "맑은 고딕", sz: 15, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 1, 1, 10, {
    fill: { patternType: "solid", fgColor: { rgb: "FFEAF2F8" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF365B7D" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 2, 1, 2, 10, {
    fill: { patternType: "solid", fgColor: { rgb: "FF1F4E78" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: thinBorder("FFFFFFFF"),
  });
  for (let r = 4; r < lastRow; r += 1) {
    styleCellRange(sheet, r, 1, r, 10, {
      fill: { patternType: "solid", fgColor: { rgb: r % 2 ? "FFFFFFFF" : "FFF7FAFC" } },
      font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF222222" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFD9E1E8"),
    });
    const reasonAddr = XLSX.utils.encode_cell({ r, c: 9 });
    if (sheet[reasonAddr]) sheet[reasonAddr].s.alignment.horizontal = "left";
    const evidenceAddr = XLSX.utils.encode_cell({ r, c: 10 });
    if (sheet[evidenceAddr]) {
      sheet[evidenceAddr].s.fill = { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } };
      sheet[evidenceAddr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF7F6000" } };
    }
  }
}

function styleDashboardShell(sheet, rowCount, colCount, cards, regionRows) {
  styleCellRange(sheet,0,0,0,colCount-1,{fill:{patternType:"solid",fgColor:{rgb:"FF0B3B76"}},font:{name:"맑은 고딕",sz:18,bold:true,color:{rgb:"FFFFFFFF"}},alignment:{horizontal:"left",vertical:"center"}});
  styleCellRange(sheet,1,0,1,colCount-1,{fill:{patternType:"solid",fgColor:{rgb:"FFF4F7FB"}},font:{name:"맑은 고딕",sz:10,color:{rgb:"FF40516B"}},alignment:{horizontal:"right",vertical:"center"}});
  const palettes=[{fill:"FFF7FAFF",border:"FF8FB7E8",font:"FF0B5CCB"},{fill:"FFFFF5F5",border:"FFF1A2A7",font:"FFC00000"},{fill:"FFFFF8EF",border:"FFF4C27A",font:"FFC55A11"},{fill:"FFF3F8FF",border:"FF9CC2EF",font:"FF2F75B5"},{fill:"FFF3FBF6",border:"FF9FD3B2",font:"FF107C41"}];
  cards.forEach(([col],i)=>{const p=palettes[i]||palettes[0];styleCellRange(sheet,2,col,4,Math.min(col+1,colCount-1),{fill:{patternType:"solid",fgColor:{rgb:p.fill}},font:{name:"맑은 고딕",sz:10,bold:true,color:{rgb:"FF26364D"}},alignment:{horizontal:"center",vertical:"center",wrapText:true},border:thinBorder(p.border)});styleCellRange(sheet,3,col,4,Math.min(col+1,colCount-1),{fill:{patternType:"solid",fgColor:{rgb:p.fill}},font:{name:"맑은 고딕",sz:19,bold:true,color:{rgb:p.font}},alignment:{horizontal:"center",vertical:"center"},border:thinBorder(p.border)});});
  styleCellRange(sheet,2,10,4,colCount-1,{fill:{patternType:"solid",fgColor:{rgb:"FFFFFFFF"}},font:{name:"맑은 고딕",sz:8.5,bold:true,color:{rgb:"FF2E3B52"}},alignment:{horizontal:"left",vertical:"center",wrapText:true},border:thinBorder("FFD6DFEA")});
  styleCellRange(sheet,6,0,6,colCount-1,{fill:{patternType:"solid",fgColor:{rgb:"FF0B3B76"}},font:{name:"맑은 고딕",sz:10,bold:true,color:{rgb:"FFFFFFFF"}},alignment:{horizontal:"center",vertical:"center",wrapText:true},border:thinBorder("FF8EA6C3")});
  const rp={서울:{fill:"FFEAF2FB",font:"FF0B5CCB"},경인:{fill:"FFF0F6FC",font:"FF2F75B5"},충청:{fill:"FFEFF7F1",font:"FF107C41"},경북:{fill:"FFF3EFF9",font:"FF7030A0"},경남:{fill:"FFFFF3EA",font:"FFC55A11"},전라:{fill:"FFEDF8FA",font:"FF008C95"}};
  for(const x of regionRows){const p=rp[x.region]||rp.서울;styleCellRange(sheet,x.row,0,x.row,colCount-1,{fill:{patternType:"solid",fgColor:{rgb:p.fill}},font:{name:"맑은 고딕",sz:10,bold:true,color:{rgb:p.font}},alignment:{horizontal:"left",vertical:"center"},border:thinBorder("FFD6DFEA")});}
}

function styleCellRange(sheet, sr, sc, er, ec, style, preserveFill = false) {
  for (let r = sr; r <= er; r += 1) {
    for (let c = sc; c <= ec; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      if (!sheet[address]) sheet[address] = { t: "s", v: "" };
      const base = sheet[address].s ? clone(sheet[address].s) : {};
      const merged = { ...base, ...clone(style) };
      if (preserveFill && base.fill) merged.fill = base.fill;
      sheet[address].s = merged;
    }
  }
}

function thinBorder(color) {
  return {
    top: { style: "thin", color: { rgb: color } },
    bottom: { style: "thin", color: { rgb: color } },
    left: { style: "thin", color: { rgb: color } },
    right: { style: "thin", color: { rgb: color } },
  };
}

function applyReasonStyle(sheet, address, isLabel) {
  if (!sheet[address]) sheet[address] = { t: "s", v: "" };
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: isLabel ? "FFF2F2F2" : "FFFFFFFF" } },
    font: { ...(base.font || {}), bold: isLabel, color: { rgb: isLabel ? "FFCC0000" : "FF333333" }, sz: 10 },
    alignment: { ...(base.alignment || {}), horizontal: isLabel ? "center" : "left", vertical: "center", wrapText: true },
  };
}

function addMerge(sheet, sr, sc, er, ec) {
  sheet["!merges"] = sheet["!merges"] || [];
  const exists = sheet["!merges"].some((merge) => merge.s.r === sr && merge.s.c === sc && merge.e.r === er && merge.e.c === ec);
  if (!exists) sheet["!merges"].push({ s: { r: sr, c: sc }, e: { r: er, c: ec } });
}

function setRef(sheet, rows, cols) {
  sheet["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, rows - 1), c: Math.max(0, cols - 1) } });
}

function extendRefForAddress(sheet, address) {
  const cell = XLSX.utils.decode_cell(address);
  const current = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  current.e.r = Math.max(current.e.r, cell.r);
  current.e.c = Math.max(current.e.c, cell.c);
  sheet["!ref"] = XLSX.utils.encode_range(current);
}

async function applyLiveEvidenceConditionalFormatting(buffer, result) {
  if (typeof JSZip === "undefined" || typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") return buffer;
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookPath="xl/workbook.xml",relsPath="xl/_rels/workbook.xml.rels",stylesPath="xl/styles.xml";
    const workbookXml=await zip.file(workbookPath)?.async("string"),relsXml=await zip.file(relsPath)?.async("string"),stylesXml=await zip.file(stylesPath)?.async("string");
    if(!workbookXml||!relsXml||!stylesXml)return buffer;
    const dxf=appendConditionalDxfs(stylesXml);
    const configs=[
      ["출근증빙·휴무확인",[{sqref:"K8:W1000",rules:[{dxfId:dxf.evidence,formula:'OR(UPPER(TRIM(K8))="O",K8="○",K8="ㅇ")'}]},{sqref:"Y8:Y1000",rules:[{dxfId:dxf.completed,formula:'$Y8="처리 완료"'},{dxfId:dxf.pending,formula:'$Y8="미처리"'}]}]],
      ["계획&근태 상이 인원",[{sqref:"M8:M1000",rules:[{dxfId:dxf.evidence,formula:'OR(UPPER(TRIM(M8))="O",M8="○",M8="ㅇ")'}]},{sqref:"N8:N1000",rules:[{dxfId:dxf.completed,formula:'$N8="처리 완료"'},{dxfId:dxf.pending,formula:'$N8="미처리"'}]}]],
      ["휴무 초과자",[{sqref:"R8:R1000",rules:[{dxfId:dxf.evidence,formula:'OR(UPPER(TRIM(R8))="O",R8="○",R8="ㅇ")'}]},{sqref:"S8:S1000",rules:[{dxfId:dxf.completed,formula:'$S8="처리 완료"'},{dxfId:dxf.pending,formula:'$S8="미처리"'}]}]],
      ["매니저별 이상 근태",[{sqref:"L8:L1000",rules:[{dxfId:dxf.evidence,formula:'OR(UPPER(TRIM(L8))="O",L8="○",L8="ㅇ")'}]},{sqref:"M8:M1000",rules:[{dxfId:dxf.completed,formula:'$M8="전달 완료"'},{dxfId:dxf.pending,formula:'$M8="미전달"'}]}]],
      ["해당 월 연차 등록 현황 및 일자",[{sqref:"O8:O1000",rules:[
        {dxfId:dxf.annualApproved,formula:'$O8="승인 완료"'},
        {dxfId:dxf.annualWaiting,formula:'$O8="승인 대기"'},
        {dxfId:dxf.annualReview,formula:'$O8="확인 요청"'},
        {dxfId:dxf.completed,formula:'$O8="확인 완료"'}
      ]}]],
    ];
    const changed=[];
    for(const [sheetName,rules] of configs){const path=findWorksheetPath(workbookXml,relsXml,sheetName);if(!path)continue;const xml=await zip.file(path)?.async("string");if(!xml)continue;const updated=addWorksheetConditionalFormatting(xml,rules);parseXmlOrThrow(updated,`${sheetName} XML`);zip.file(path,updated);changed.push([path,sheetName]);}
    parseXmlOrThrow(dxf.xml,"styles.xml");zip.file(stylesPath,dxf.xml);const candidate=await zip.generateAsync({type:"arraybuffer",compression:"DEFLATE"});const verify=await JSZip.loadAsync(candidate);parseXmlOrThrow(await verify.file(stylesPath)?.async("string")||"","저장 후 styles.xml");for(const[path,name]of changed)parseXmlOrThrow(await verify.file(path)?.async("string")||"",`저장 후 ${name} XML`);return candidate;
  } catch(error){console.warn("실시간 처리 색상 규칙 적용 실패 · 기본 파일로 저장",error);return buffer;}
}

async function applyWorkbookOpenViewSettings(buffer) {
  if (typeof JSZip === "undefined" || typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") return buffer;
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!workbookXml || !relsXml) return buffer;

    const frozenDashboardSheets = new Set([
      "계획&근태 상이 인원", "출근증빙·휴무확인", "휴무 초과자", "인사팀 급여 확정표", "출근 미달자 정산", "연차 사용 필요자", "최종 문제자", "전체 요약본", "매니저별 이상 근태", "주 근태 확인자", "해당 월 연차 등록 현황 및 일자", "인력 변동 확인",
    ]);
    const attendanceFreezeSheet = "상담사근태";
    const sheetTags = workbookXml.match(/<sheet\b[^>]*\/?\s*>/g) || [];
    const changedPaths = [];
    for (const tag of sheetTags) {
      const encodedSheetName = tag.match(/\bname="([^"]+)"/)?.[1] || "";
      const sheetName = encodedSheetName
        .replace(/&amp;/g, "&").replace(/&quot;/g, "\"").replace(/&apos;/g, "'")
        .replace(/&lt;/g, "<").replace(/&gt;/g, ">");
      const relationId = tag.match(/r:id="([^"]+)"/)?.[1] || "";
      if (!relationId) continue;
      const relTags = relsXml.match(/<Relationship\b[^>]*\/?\s*>/g) || [];
      const rel = relTags.find((item) => item.includes(`Id="${relationId}"`));
      let target = rel?.match(/Target="([^"]+)"/)?.[1] || "";
      if (!target) continue;
      target = target.replace(/^\//, "");
      const worksheetPath = target.startsWith("xl/") ? target : `xl/${target.replace(/^\.\//, "")}`;
      const originalXml = await zip.file(worksheetPath)?.async("string");
      if (!originalXml) continue;

      const doc = parseXmlOrThrow(originalXml, `${sheetName || worksheetPath} 보기 설정 XML`);
      const root = doc.documentElement;
      const namespace = root.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
      let sheetViews = directXmlChild(root, "sheetViews");
      if (!sheetViews) {
        sheetViews = doc.createElementNS(namespace, "sheetViews");
        const anchorNames = new Set(["sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState"]);
        const anchor = xmlElementChildren(root).find((node) => anchorNames.has(node.localName)) || null;
        root.insertBefore(sheetViews, anchor);
      }
      let sheetView = xmlElementChildren(sheetViews).find((node) => node.localName === "sheetView") || null;
      if (!sheetView) {
        sheetView = doc.createElementNS(namespace, "sheetView");
        sheetViews.appendChild(sheetView);
      }
      sheetView.setAttribute("workbookViewId", sheetView.getAttribute("workbookViewId") || "0");
      sheetView.setAttribute("showGridLines", "0");
      sheetView.setAttribute("zoomScale", "70");
      sheetView.setAttribute("zoomScaleNormal", "70");

      if (sheetName === attendanceFreezeSheet || frozenDashboardSheets.has(sheetName)) {
        for (const pane of xmlElementChildren(sheetView).filter((node) => node.localName === "pane")) sheetView.removeChild(pane);
        const pane = doc.createElementNS(namespace, "pane");
        if (sheetName === attendanceFreezeSheet) {
          // 상담사근태: 1~6행과 A~M열 고정, N7부터 스크롤합니다.
          pane.setAttribute("xSplit", "13");
          pane.setAttribute("ySplit", "6");
          pane.setAttribute("topLeftCell", "N7");
          pane.setAttribute("activePane", "bottomRight");
        } else {
          pane.setAttribute("ySplit", "7");
          pane.setAttribute("topLeftCell", "A8");
          pane.setAttribute("activePane", "bottomLeft");
        }
        pane.setAttribute("state", "frozen");
        const selection = xmlElementChildren(sheetView).find((node) => node.localName === "selection") || null;
        sheetView.insertBefore(pane, selection);
      }

      const updatedXml = serializeXml(doc, originalXml);
      parseXmlOrThrow(updatedXml, `${sheetName || worksheetPath} 저장 전 보기 설정 XML`);
      zip.file(worksheetPath, updatedXml);
      changedPaths.push({ path: worksheetPath, name: sheetName });
    }

    const candidate = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
    const verifyZip = await JSZip.loadAsync(candidate);
    for (const item of changedPaths) {
      const xml = await verifyZip.file(item.path)?.async("string");
      const doc = parseXmlOrThrow(xml || "", `${item.name || item.path} 저장 후 보기 설정 XML`);
      const view = doc.getElementsByTagNameNS("*", "sheetView")[0];
      if (!view || view.getAttribute("zoomScale") !== "70") throw new Error(`${item.name} 확대 비율 저장 실패`);
      if (item.name === attendanceFreezeSheet) {
        const pane = doc.getElementsByTagNameNS("*", "pane")[0];
        if (!pane
          || pane.getAttribute("xSplit") !== "13"
          || pane.getAttribute("ySplit") !== "6"
          || pane.getAttribute("topLeftCell") !== "N7") {
          throw new Error("상담사근태 6행·M열 고정 저장 실패");
        }
      } else if (frozenDashboardSheets.has(item.name)) {
        const pane = doc.getElementsByTagNameNS("*", "pane")[0];
        if (!pane || pane.getAttribute("ySplit") !== "7") throw new Error(`${item.name} 7행 고정 저장 실패`);
      }
    }
    return candidate;
  } catch (error) {
    console.warn("엑셀 보기 설정 적용 실패 · 손상 방지를 위해 직전 정상 파일로 저장", error);
    return buffer;
  }
}

function findWorksheetPath(workbookXml, relsXml, sheetName) {
  // workbook.xml의 시트명은 &가 &amp;처럼 XML 이스케이프되어 저장됩니다.
  // 계획&근태 상이 인원도 정확히 찾아 조건부서식을 적용하도록 인코딩 후 검색합니다.
  const xmlSheetName = String(sheetName || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  const escapedName = escapeRegExp(xmlSheetName);
  const sheetTags = workbookXml.match(/<sheet\b[^>]*\/?\s*>/g) || [];
  const tag = sheetTags.find((item) => new RegExp(`\\bname="${escapedName}"`).test(item));
  const relationId = tag?.match(/r:id="([^"]+)"/)?.[1];
  if (!relationId) return "";
  const relTags = relsXml.match(/<Relationship\b[^>]*\/?\s*>/g) || [];
  const rel = relTags.find((item) => item.includes(`Id="${relationId}"`));
  let target = rel?.match(/Target="([^"]+)"/)?.[1] || "";
  if (!target) return "";
  target = target.replace(/^\//, "");
  if (target.startsWith("xl/")) return target;
  return `xl/${target.replace(/^\.\//, "")}`;
}

function appendConditionalDxfs(stylesXml) {
  const doc = parseXmlOrThrow(stylesXml, "styles.xml");
  const root = doc.documentElement;
  const namespace = root.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  let dxfs = directXmlChild(root, "dxfs");
  if (!dxfs) {
    dxfs = doc.createElementNS(namespace, "dxfs");
    const tableStyles = directXmlChild(root, "tableStyles");
    root.insertBefore(dxfs, tableStyles || null);
  }

  const existingDxfs = xmlElementChildren(dxfs).filter((node) => node.localName === "dxf");
  const start = existingDxfs.length;
  const formats = [
    { font: "FF107C41", fill: "FFC6E0B4" }, // 처리 완료 · 연녹색
    { font: "FFC00000", fill: "FFF4CCCC" }, // 미처리 · 연분홍색
    { font: "FF107C41", fill: "FFD9EAD3" }, // O 입력 · 연한 민트색
    { font: "FF000000", fill: "FFA9D08E" }, // 상담사근태 확정 출근
    { font: "FF1F4E78", fill: "FFDDEBF7" }, // 연차 승인 완료 · 하늘색
    { font: "FF9C6500", fill: "FFFFE699" }, // 연차 승인 대기 · 노란색
    { font: "FF9C0006", fill: "FFF4CCCC" }, // 연차 확인 요청 · 빨간색
  ];
  for (const format of formats) dxfs.appendChild(createDxfNode(doc, namespace, format));
  dxfs.setAttribute("count", String(start + formats.length));

  return {
    xml: serializeXml(doc, stylesXml),
    completed: start,
    pending: start + 1,
    evidence: start + 2,
    attendance: start + 3,
    annualApproved: start + 4,
    annualWaiting: start + 5,
    annualReview: start + 6,
  };
}

function createDxfNode(doc, namespace, format) {
  const dxf = doc.createElementNS(namespace, "dxf");
  const font = doc.createElementNS(namespace, "font");
  font.appendChild(doc.createElementNS(namespace, "b"));
  const fontColor = doc.createElementNS(namespace, "color");
  fontColor.setAttribute("rgb", format.font);
  font.appendChild(fontColor);
  dxf.appendChild(font);

  const fill = doc.createElementNS(namespace, "fill");
  const patternFill = doc.createElementNS(namespace, "patternFill");
  // Excel이 직접 만든 조건부서식 DXF와 동일하게 bgColor만 기록합니다.
  // 기존 solid+fgColor 방식은 일부 Excel 환경에서 검은색 채우기로 표시됐습니다.
  const background = doc.createElementNS(namespace, "bgColor");
  background.setAttribute("rgb", format.fill);
  patternFill.appendChild(background);
  fill.appendChild(patternFill);
  dxf.appendChild(fill);
  return dxf;
}

function addWorksheetConditionalFormatting(worksheetXml, configurations) {
  const doc = parseXmlOrThrow(worksheetXml, "worksheet.xml");
  const root = doc.documentElement;
  const namespace = root.namespaceURI || "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  // 동일 범위의 기존 규칙만 제거해 재생성·중복 실행에도 규칙이 누적되지 않게 합니다.
  const managedRanges = new Set(configurations.map((item) => item.sqref));
  for (const node of [...root.getElementsByTagNameNS("*", "conditionalFormatting")]) {
    if (node.parentNode === root && managedRanges.has(node.getAttribute("sqref") || "")) root.removeChild(node);
  }

  let maxPriority = 0;
  for (const rule of root.getElementsByTagNameNS("*", "cfRule")) {
    maxPriority = Math.max(maxPriority, Number(rule.getAttribute("priority") || 0));
  }

  const laterElements = new Set([
    "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
    "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags",
    "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
    "webPublishItems", "tableParts", "extLst",
  ]);
  const anchor = xmlElementChildren(root).find((node) => laterElements.has(node.localName)) || null;

  for (const configuration of configurations) {
    const conditional = doc.createElementNS(namespace, "conditionalFormatting");
    conditional.setAttribute("sqref", configuration.sqref);
    for (const configRule of configuration.rules || []) {
      const rule = doc.createElementNS(namespace, "cfRule");
      rule.setAttribute("type", "expression");
      rule.setAttribute("dxfId", String(configRule.dxfId));
      rule.setAttribute("priority", String(++maxPriority));
      const formula = doc.createElementNS(namespace, "formula");
      formula.textContent = configRule.formula;
      rule.appendChild(formula);
      conditional.appendChild(rule);
    }
    root.insertBefore(conditional, anchor);
  }
  return serializeXml(doc, worksheetXml);
}

function parseXmlOrThrow(xml, label) {
  if (!xml) throw new Error(`${label} 내용이 비어 있습니다.`);
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const errors = doc.getElementsByTagName("parsererror");
  if (errors.length) throw new Error(`${label} 파싱 실패: ${String(errors[0].textContent || "XML 오류").slice(0, 200)}`);
  return doc;
}

function serializeXml(doc, originalXml = "") {
  const body = new XMLSerializer().serializeToString(doc.documentElement);
  const declaration = String(originalXml).match(/^\s*(<\?xml[^?]*\?>)/)?.[1]
    || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const xml = `${declaration}${body}`;
  parseXmlOrThrow(xml, "직렬화 XML");
  return xml;
}

function directXmlChild(parent, localName) {
  return xmlElementChildren(parent).find((node) => node.localName === localName) || null;
}

function xmlElementChildren(parent) {
  return Array.from(parent.childNodes || []).filter((node) => node.nodeType === 1);
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sanitizeWorkbookForExcel(workbook) {
  for (const sheetName of workbook.SheetNames || []) {
    const sheet = workbook.Sheets?.[sheetName];
    if (!sheet) continue;

    // xlsx-js-style의 쓰기 옵션에 showGridLines를 직접 넘기지 않습니다.
    // 지원되는 시트 보기 속성으로만 저장해 파일 구조 손상을 피합니다.
    sheet["!views"] = [{ showGridLines: false, zoomScale: 70, zoomScaleNormal: 70 }];

    // 중복 또는 서로 겹치는 병합 범위는 Excel 복구 경고의 주요 원인입니다.
    const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
    const clean = [];
    for (const merge of merges) {
      if (!merge?.s || !merge?.e) continue;
      const normalized = {
        s: { r: Math.min(merge.s.r, merge.e.r), c: Math.min(merge.s.c, merge.e.c) },
        e: { r: Math.max(merge.s.r, merge.e.r), c: Math.max(merge.s.c, merge.e.c) },
      };
      const overlaps = clean.some((item) => !(
        normalized.e.r < item.s.r || normalized.s.r > item.e.r ||
        normalized.e.c < item.s.c || normalized.s.c > item.e.c
      ));
      if (!overlaps) clean.push(normalized);
    }
    sheet["!merges"] = clean;
  }
}

function clone(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}
