const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];
const NON_WORK_CODES = new Set([
  "휴무", "무급휴가", "연차", "공가", "휴가", "경조",
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
  buildAnnualLedgerSheet(workbook, result, context, year, monthNo);
  fillPlanSheet(workbook.Sheets["근무 계획"], result, context, daysInMonth);
  fillAttendanceRawSheet(workbook.Sheets["근태 RAW"], result);

  workbook.SheetNames = [
    "상담사근태",
    "계획&근태 상이 인원",
    "출근 미등록",
    "휴무 초과자",
    "전체 요약본",
    "매니저별 이상 근태",
    "해당 월 연차 등록 현황 및 일자",
    "연차 누적 현황",
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
  const buffer = await applyLiveEvidenceConditionalFormatting(rawBuffer, result);
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
  const evidenceSet = new Set((result.evidenceOverrides || []).map(String));
  const issueMap = new Map();
  for (const issue of result.mismatchRows || []) addIssue(issueMap, issue.employeeId, issue.date, issue.reason || issue.result);
  for (const summary of result.employeeSummaries || []) {
    for (const event of summary.baseExcessEvents || []) {
      addIssue(issueMap, summary.employeeId, event.date, `기본 휴무 기준 ${daysText(summary.baseAllowance)} 초과분`);
    }
    const combinedAvailable = roundHalf(Number(summary.availableSubstitute || 0) + Number(summary.availableCompensation || 0));
    const combinedUsed = roundHalf(Number(summary.substituteNeeded ?? summary.explicitSubDayoffUsed ?? 0) + Number(summary.compensationLeaveUsed ?? summary.compensationNeeded ?? 0));
    const combinedShortage = Number(summary.combinedShortage ?? Math.max(0, combinedUsed - combinedAvailable));
    if (combinedShortage > 0) {
      const combinedEvents = [...(summary.substituteEvents || []), ...(summary.compensationEvents || [])]
        .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
      let running = 0;
      for (const event of combinedEvents) {
        running = roundHalf(running + Number(event.days || 0));
        if (running > combinedAvailable) {
          addIssue(issueMap, summary.employeeId, event.date, `대체·보상 휴무 잔여 부족 · 총 ${daysText(combinedShortage)} 초과 사용`);
        }
      }
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
    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = `${result.targetMonth}-${String(day).padStart(2, "0")}`;
      const planStatus = normalizePlan(person.plan?.plans?.[day]);
      const attendance = attendanceByKey.get(`${person.employeeId}|${date}`) || emptyAttendance();
      const evidenceKey = `${person.employeeId}|${date}`;
      const evidence = evidenceSet.has(evidenceKey);
      const issues = issueMap.get(`${person.employeeId}|${day}`) || [];
      const finalAttendance = evidence
        ? { ...attendance, hasClockIn: true, actualStatus: "출근", evidenced: true }
        : attendance;
      const evidenceResolvesMissing = evidence && WORK_CLOCK_CODES.has(planStatus);
      daily[day] = {
        date,
        planStatus,
        attendance: finalAttendance,
        evidence,
        display: evidence ? "출근" : dailyDisplay(planStatus, finalAttendance),
        // 근무 계획의 단순 미입력은 O로 해소하지만, 휴무·연차 등 계획 상이는 유지합니다.
        issues: evidenceResolvesMissing ? [] : [...new Set(issues)],
      };
    }
    dailyByKey.set(person.key, daily);
  }

  return { workforce, workforceById, planById, attendanceByKey, summaryById, issueMap, people, dailyByKey };
}


function buildIssueSummarySheet(workbook, result, ctx, year, monthNo) {
  const sheetName = "전체 요약본";
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
        regionalManager: member.regionalManager || "",
        manager: member.manager || "",
        region: member.region2 || member.region1 || "",
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
    const group = ensureGroup(summary.employeeId, summary);
    if (!group) continue;
    const combinedShortage = Number(summary.combinedShortage ?? Math.max(0,
      Number(summary.substituteNeeded ?? summary.explicitSubDayoffUsed ?? 0) + Number(summary.compensationNeeded || 0)
      - Number(summary.availableSubstitute || 0) - Number(summary.availableCompensation || 0)
    ));
    if (combinedShortage > 0) {
      group.issues.push(`대체·보상 휴무 ${daysText(combinedShortage)} 초과 사용`);
      group.priorities.add("긴급");
    }
  }

  const rows = [...groups.values()]
    .filter((group) => group.issues.length)
    .map((group) => ({
      ...group,
      issues: [...new Set(group.issues)],
      evidence: [...new Set(group.evidence)],
      priority: group.priorities.has("긴급") ? "긴급" : group.priorities.has("확인 필요") ? "확인 필요" : "등록 확인",
    }))
    .sort((a, b) => String(a.regionalManager).localeCompare(String(b.regionalManager), "ko")
      || String(a.manager).localeCompare(String(b.manager), "ko")
      || String([...a.stores][0] || "").localeCompare(String([...b.stores][0] || ""), "ko")
      || String(a.name).localeCompare(String(b.name), "ko"));

  const totalIssues = rows.reduce((sum, row) => sum + row.issues.length, 0);
  const evidencePeople = rows.filter((row) => row.evidence.length).length;
  const urgentPeople = rows.filter((row) => row.priority === "긴급").length;
  const matrix = [
    [`${year}년 ${monthNo}월 근태관리 전체 요약`],
    [result.referenceComparison?.supplied ? `직원별 오류 요약 · 비교 최종본 결과: ${result.referenceComparison.summary || "확인 필요"}` : "직원별 이름·사번·날짜·문제 내용을 한 행에서 확인할 수 있도록 정리했습니다."],
    [],
    ["검토 대상 인원", "", "총 문제 건수", "", "증빙 필요 인원", "", "긴급 확인 인원", ""],
    [rows.length, "", totalIssues, "", evidencePeople, "", urgentPeople, ""],
    [],
    ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "문제 건수", "문제 요약", "증빙 필요 내역", "우선순위", "처리상태"],
  ];
  rows.forEach((row, index) => {
    matrix.push([
      index + 1,
      row.regionalManager,
      row.manager,
      row.region,
      [...row.stores].join(", "),
      row.name,
      row.employeeId,
      row.issues.length,
      row.issues.join(" / "),
      row.evidence.join(" / "),
      row.priority,
      "미처리",
    ]);
  });
  if (!rows.length) matrix.push([1, "", "", "", "", "", "", 0, "확인 필요한 근태 오류가 없습니다.", "", "정상", "완료"]);
  const referenceRows = (result.referenceComparison?.rows || []).filter((row) => !row.match);
  if (result.referenceComparison?.supplied) {
    matrix.push([]);
    matrix.push(["비교 최종본 불일치", "", "", "", "", "", "", referenceRows.length, result.referenceComparison.summary || "비교 결과 확인", "", "", ""]);
    for (const row of referenceRows) {
      matrix.push(["비교", row.regionalManager || "", row.manager || "", row.region || "", row.store || "", row.name || "", row.employeeId || "", 1,
        `${row.date || row.comparisonType || "항목"} · 자동 ${row.generatedValue || "-"} / 비교 ${row.referenceValue || "-"}`,
        row.reason || "", row.result || "불일치", "확인 필요"]);
    }
  }

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 11 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 3, c: 2 }, e: { r: 3, c: 3 } },
    { s: { r: 4, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 3, c: 4 }, e: { r: 3, c: 5 } },
    { s: { r: 4, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 3, c: 6 }, e: { r: 3, c: 7 } },
    { s: { r: 4, c: 6 }, e: { r: 4, c: 7 } },
  ];
  sheet["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 10 }, { wch: 16 }, { wch: 11 },
    { wch: 13 }, { wch: 9 }, { wch: 48 }, { wch: 28 }, { wch: 11 }, { wch: 11 },
  ];
  sheet["!rows"] = [{ hpt: 28 }, { hpt: 22 }, { hpt: 8 }, { hpt: 22 }, { hpt: 28 }, { hpt: 8 }, { hpt: 30 }];
  sheet["!autofilter"] = { ref: `A7:L${Math.max(8, matrix.length)}` };

  styleSummarySheet(sheet, matrix.length);
  workbook.Sheets[sheetName] = sheet;
  workbook.SheetNames = [sheetName, ...workbook.SheetNames.filter((name) => name !== sheetName)];
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
  const rows = result.managerRequests || [];
  const matrix = [
    [`${year}년 ${monthNo}월 매니저별 상담사 이상 근태`],
    ["복사용 수정요청 멘트를 그대로 전달할 수 있으며, 근태·계획·연차 신청 불일치가 함께 반영됩니다."],
    [],
    ["지역장", "매니저", "지역", "매장명", "이름", "사번", "문제 건수", "문제 요약", "복사용 수정요청 멘트", "전달상태"],
  ];
  for (const row of rows) matrix.push([
    row.regionalManager || "", row.manager || "", row.region || "", row.store || "", row.name || "",
    row.employeeId || "", row.issueCount || 0, row.issueText || "", row.message || "", "미전달",
  ]);
  if (!rows.length) matrix.push(["", "", "", "", "", "", 0, "수정 요청 대상이 없습니다.", "", "완료"]);
  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } },
  ];
  sheet["!cols"] = [
    { wch: 11 }, { wch: 11 }, { wch: 10 }, { wch: 17 }, { wch: 11 },
    { wch: 13 }, { wch: 9 }, { wch: 52 }, { wch: 90 }, { wch: 12 },
  ];
  sheet["!rows"] = [{ hpt: 30 }, { hpt: 24 }, { hpt: 8 }, { hpt: 30 }];
  styleSimpleReportSheet(sheet, matrix.length, 10, [7, 8]);
  for (let r = 4; r < matrix.length; r += 1) sheet["!rows"][r] = { hpt: 48 };
  workbook.Sheets["매니저별 이상 근태"] = sheet;
  if (!workbook.SheetNames.includes("매니저별 이상 근태")) workbook.SheetNames.push("매니저별 이상 근태");
}

function buildAnnualComparisonSheet(workbook, result, year, monthNo) {
  const comparison = result.annualComparison || { rows: [], matchCount: 0, mismatchCount: 0, missingApplicationCount: 0, supplied: false };
  const matrix = [
    [`${year}년 ${monthNo}월 연차 등록 현황 및 일자`],
    [comparison.supplied ? "계획 기준 연차·반차 신청과 월초 승인·반려 현황을 사번·날짜 기준으로 정리했습니다. 승인 건만 누적 연차에서 차감됩니다." : "저장된 승인·반려 파일이 없어 근무계획의 연차·반차만 표시합니다."],
    [],
    ["일치", "", "신청 있음·계획 불일치", "", "계획 있음·신청 없음", ""],
    [comparison.matchCount || 0, "", comparison.mismatchCount || 0, "", comparison.missingApplicationCount || 0, ""],
    [],
    ["휴가일자", "지역장", "매니저", "지역", "매장명", "이름", "사번", "신청구분", "신청일수", "근무계획", "대조결과", "신청상태", "비고"],
  ];
  for (const row of comparison.rows || []) matrix.push([
    row.date || "", row.regionalManager || "", row.manager || "", row.region || "", row.store || "",
    row.name || "", row.employeeId || "", row.requestedKind || "-", row.requestedDays || 0,
    row.planStatus || "공백", row.result || "", row.applicationStatus || "", row.note || "",
  ]);
  if (!(comparison.rows || []).length) matrix.push(["", "", "", "", "", "", "", "", 0, "", "대조 자료 없음", "", ""]);
  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 3, c: 2 }, e: { r: 3, c: 3 } },
    { s: { r: 4, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 3, c: 4 }, e: { r: 3, c: 5 } },
    { s: { r: 4, c: 4 }, e: { r: 4, c: 5 } },
  ];
  sheet["!cols"] = [
    { wch: 12 }, { wch: 11 }, { wch: 11 }, { wch: 10 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 11 }, { wch: 10 }, { wch: 16 }, { wch: 30 }, { wch: 11 }, { wch: 25 },
  ];
  sheet["!rows"] = [{ hpt: 30 }, { hpt: 24 }, { hpt: 8 }, { hpt: 22 }, { hpt: 28 }, { hpt: 8 }, { hpt: 30 }];
  styleSimpleReportSheet(sheet, matrix.length, 13, [10, 12]);
  for (const [start, color] of [[0, "FFE2F0D9"], [2, "FFFCE4D6"], [4, "FFFFF2CC"]]) {
    styleCellRange(sheet, 3, start, 3, start + 1, {
      fill: { patternType: "solid", fgColor: { rgb: color } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFCBD5E1"),
    });
    styleCellRange(sheet, 4, start, 4, start + 1, {
      fill: { patternType: "solid", fgColor: { rgb: "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 16, bold: true, color: { rgb: "FF173F73" } },
      alignment: { horizontal: "center", vertical: "center" }, border: thinBorder("FFCBD5E1"),
    });
  }
  for (let r = 7; r < matrix.length; r += 1) {
    const resultAddr = XLSX.utils.encode_cell({ r, c: 10 });
    const value = String(sheet[resultAddr]?.v || "");
    if (value === "일치") {
      sheet[resultAddr].s.fill = { patternType: "solid", fgColor: { rgb: "FFE2F0D9" } };
      sheet[resultAddr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF375623" } };
    } else {
      sheet[resultAddr].s.fill = { patternType: "solid", fgColor: { rgb: value.includes("신청내역 없음") ? "FFFFF2CC" : "FFF4CCCC" } };
      sheet[resultAddr].s.font = { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: value.includes("신청내역 없음") ? "FF7F6000" : "FF9C0006" } };
    }
  }
  workbook.Sheets["해당 월 연차 등록 현황 및 일자"] = sheet;
  if (!workbook.SheetNames.includes("해당 월 연차 등록 현황 및 일자")) workbook.SheetNames.push("해당 월 연차 등록 현황 및 일자");
}



function buildEvidenceDashboardSheet(workbook, result, ctx, year, monthNo) {
  const regionOrder = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const rows = [...(result.missingRows || [])].map((row) => {
    const member = findMember(ctx, row.employeeId, row.store) || {};
    return {
      ...row,
      member,
      region: normalizeDashboardRegion(member.region2 || member.region1 || row.region || "", member.storeName || row.store || ""),
    };
  });

  rows.sort((a, b) => regionOrder.indexOf(a.region) - regionOrder.indexOf(b.region)
    || String(a.member.regionalManager || "").localeCompare(String(b.member.regionalManager || ""), "ko")
    || String(a.member.manager || "").localeCompare(String(b.member.manager || ""), "ko")
    || String(a.member.storeName || a.store || "").localeCompare(String(b.member.storeName || b.store || ""), "ko")
    || String(a.date || "").localeCompare(String(b.date || ""))
    || String(a.name || "").localeCompare(String(b.name || ""), "ko"));

  const uniquePeople = new Set(rows.map((row) => normalizeId(row.employeeId)).filter(Boolean)).size;
  const uniqueStores = new Set(rows.map((row) => row.member.storeName || row.store || "").filter(Boolean)).size;
  const matrix = Array.from({ length: 7 }, () => Array(13).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 출근 미등록`;
  matrix[1][0] = `사번·발생일 기준 증빙 관리 · K열에 O 입력 · 기준일 ${result.cutoffDate || `${year}-${String(monthNo).padStart(2, "0")}-${String(new Date(year, monthNo, 0).getDate()).padStart(2, "0")}`}`;

  const cards = [
    [0, "총 미등록 건수", `${rows.length}건`],
    [2, "대상 인원", `${uniquePeople}명`],
    [4, "대상 매장", `${uniqueStores}개`],
    [6, "미처리", `${rows.length}건`],
    [8, "처리 완료", "0건"],
  ];
  for (const [col, label, value] of cards) {
    matrix[2][col] = label;
    matrix[3][col] = value;
  }
  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 출근 미입력";
  matrix[3][11] = "● 계획 미입력";
  matrix[4][10] = "● 출ㆍ계 미입력";
  matrix[4][11] = "● 처리 완료";
  matrix[6] = ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "발생일", "근무계획", "구분", "증빙여부(O 입력)", "상세사유", "처리상태"];

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
      `출근 미입력 ${typeCounts.clock}건  |  계획 미입력 ${typeCounts.plan}건  |  출ㆍ계 미입력 ${typeCounts.both}건  |  완료 0건`, "", "", "", "",
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
        "",
        row.reason || row.result || "근무계획 또는 실제 출근기록 확인 필요",
        "미처리",
      ]);
      dataRows.push({ row: matrix.length - 1, missingType });
    }
  }

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 12 } },
    { s: { r: 3, c: 11 }, e: { r: 3, c: 12 } },
    { s: { r: 4, c: 11 }, e: { r: 4, c: 12 } },
  ];
  for (const item of regionRows) {
    sheet["!merges"].push(
      { s: { r: item.row, c: 0 }, e: { r: item.row, c: 7 } },
      { s: { r: item.row, c: 8 }, e: { r: item.row, c: 12 } },
    );
  }

  sheet["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 13 }, { wch: 17 }, { wch: 27 }, { wch: 18 }, { wch: 46 }, { wch: 12 },
  ];
  sheet["!rows"] = matrix.map((_, index) => ({ hpt: index === 0 ? 34 : index === 1 ? 24 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 30 : regionRows.some((item) => item.row === index) ? 26 : 24 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false }];

  const rangeStart = 8;
  const rangeEnd = Math.max(rangeStart, matrix.length);
  setFormula(sheet, "A4", `COUNTA($G$${rangeStart}:$G$${rangeEnd})&"건"`, `${rows.length}건`);
  setValue(sheet, "C4", `${uniquePeople}명`);
  setValue(sheet, "E4", `${uniqueStores}개`);
  const doneAllFormula = `(COUNTIF($K$${rangeStart}:$K$${rangeEnd},"O")+COUNTIF($K$${rangeStart}:$K$${rangeEnd},"○")+COUNTIF($K$${rangeStart}:$K$${rangeEnd},"ㅇ"))`;
  setFormula(sheet, "G4", `(COUNTA($G$${rangeStart}:$G$${rangeEnd})-${doneAllFormula})&"건"`, `${rows.length}건`);
  setFormula(sheet, "I4", `${doneAllFormula}&"건"`, "0건");

  for (const item of regionRows) {
    const excelRow = item.row + 1;
    const totalFormula = `COUNTIF($D$${rangeStart}:$D$${rangeEnd},"${item.region}")`;
    const doneFormula = `(COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$K$${rangeStart}:$K$${rangeEnd},"O")+COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$K$${rangeStart}:$K$${rangeEnd},"○")+COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$K$${rangeStart}:$K$${rangeEnd},"ㅇ"))`;
    const clockFormula = `COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$J$${rangeStart}:$J$${rangeEnd},"출근 미입력")`;
    const planFormula = `COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$J$${rangeStart}:$J$${rangeEnd},"계획 미입력")`;
    const bothFormula = `COUNTIFS($D$${rangeStart}:$D$${rangeEnd},"${item.region}",$J$${rangeStart}:$J$${rangeEnd},"출ㆍ계 미입력")`;
    setFormula(sheet, `A${excelRow}`, `"▼  ${item.region} (총 "&${totalFormula}&"건)"`, `▼  ${item.region} (총 ${item.count}건)`);
    setFormula(sheet, `I${excelRow}`, `"출근 미입력 "&${clockFormula}&"건  |  계획 미입력 "&${planFormula}&"건  |  출ㆍ계 미입력 "&${bothFormula}&"건  |  완료 "&${doneFormula}&"건"`, `출근 미입력 ${item.typeCounts.clock}건  |  계획 미입력 ${item.typeCounts.plan}건  |  출ㆍ계 미입력 ${item.typeCounts.both}건  |  완료 0건`);
  }

  dataRows.forEach((item) => {
    const excelRow = item.row + 1;
    setFormula(sheet, `M${excelRow}`, `IF(OR(UPPER(TRIM(K${excelRow}))="O",K${excelRow}="○",K${excelRow}="ㅇ"),"처리 완료","미처리")`, "미처리");
    setNumberFormat(sheet, `H${excelRow}`, "yyyy-mm-dd");
  });

  styleCellRange(sheet, 0, 0, 0, 12, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, 12, {
    fill: { patternType: "solid", fgColor: { rgb: "FFF4F7FB" } },
    font: { name: "맑은 고딕", sz: 10, color: { rgb: "FF40516B" } },
    alignment: { horizontal: "right", vertical: "center" },
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

  styleCellRange(sheet, 6, 0, 6, 12, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
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
    styleCellRange(sheet, item.row, 0, item.row, 12, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "left", vertical: "center" },
      border: thinBorder("FFD6DFEA"),
    });
    const summaryAddr = XLSX.utils.encode_cell({ r: item.row, c: 8 });
    if (sheet[summaryAddr]) sheet[summaryAddr].s.alignment = { horizontal: "right", vertical: "center" };
  }

  dataRows.forEach((item, index) => {
    styleCellRange(sheet, item.row, 0, item.row, 12, {
      fill: { patternType: "solid", fgColor: { rgb: index % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    const reasonAddr = XLSX.utils.encode_cell({ r: item.row, c: 11 });
    if (sheet[reasonAddr]) sheet[reasonAddr].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
    applyMissingTypeStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: 9 }), item.missingType);
    const evidenceAddr = XLSX.utils.encode_cell({ r: item.row, c: 10 });
    if (sheet[evidenceAddr]) {
      sheet[evidenceAddr].s = {
        ...(sheet[evidenceAddr].s || {}),
        fill: { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } },
        font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFC55A11" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: thinBorder("FFF4C27A"),
      };
    }
    applyProcessStatusStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: 12 }), "미처리");
  });

  setRef(sheet, matrix.length, 13);
  workbook.Sheets["출근 미등록"] = sheet;
  if (!workbook.SheetNames.includes("출근 미등록")) workbook.SheetNames.push("출근 미등록");
}

function buildPlanAttendanceMatchSheet(workbook, result, ctx, year, monthNo) {
  const regionOrder = ["서울", "경인", "충청", "경북", "경남", "전라"];
  const evidenceSet = new Set((result.evidenceOverrides || []).map(String));
  const rows = (result.mismatchRows || []).map((row) => {
    const member = findMember(ctx, row.employeeId, row.store) || {};
    const evidenceKey = `${normalizeId(row.employeeId)}|${row.date}`;
    const evidenced = Boolean(row.evidenced || evidenceSet.has(evidenceKey));
    const category = mismatchDashboardCategory(row);
    return {
      ...row,
      member,
      evidenced,
      category,
      region: normalizeDashboardRegion(member.region2 || member.region1 || row.region || "", member.storeName || row.store || ""),
      status: evidenced && row.issueType !== "missing_plan"
        ? "처리 완료"
        : category === "근무인데 출근기록 없음" ? "미처리" : "확인중",
    };
  }).filter((row) => !(row.evidenced && row.category === "근무인데 출근기록 없음"));

  rows.sort((a, b) => regionOrder.indexOf(a.region) - regionOrder.indexOf(b.region)
    || String(a.member.regionalManager || "").localeCompare(String(b.member.regionalManager || ""), "ko")
    || String(a.member.manager || "").localeCompare(String(b.member.manager || ""), "ko")
    || String(a.member.storeName || a.store || "").localeCompare(String(b.member.storeName || b.store || ""), "ko")
    || String(a.date || "").localeCompare(String(b.date || ""))
    || String(a.name || "").localeCompare(String(b.name || ""), "ko"));

  const total = rows.length;
  const missingCount = rows.filter((row) => row.category === "근무인데 출근기록 없음" && row.status !== "처리 완료").length;
  const unexpectedCount = rows.filter((row) => row.category === "휴무·휴가인데 출근기록 있음" && row.status !== "처리 완료").length;
  const reviewCount = rows.filter((row) => row.category === "검토 필요" && row.status !== "처리 완료").length;
  const completedCount = rows.filter((row) => row.status === "처리 완료").length;

  const matrix = Array.from({ length: 7 }, () => Array(13).fill(""));
  matrix[0][0] = `${year}년 ${monthNo}월 계획 & 근태 상이 인원`;
  matrix[1][0] = `사번·발생일 기준 자동 비교 · 기준일 ${result.cutoffDate || `${year}-${String(monthNo).padStart(2, "0")}-${String(new Date(year, monthNo, 0).getDate()).padStart(2, "0")}`} · 증빙 O 반영`;
  const cards = [
    [0, "총 상이 건수", total],
    [2, "근무인데 출근기록 없음", missingCount],
    [4, "휴무·휴가인데 출근기록 있음", unexpectedCount],
    [6, "검토 필요", reviewCount],
    [8, "처리 완료", completedCount],
  ];
  for (const [col, label, value] of cards) {
    matrix[2][col] = label;
    matrix[3][col] = `${value}건`;
  }
  matrix[2][10] = "구분 색상 안내";
  matrix[3][10] = "● 근무인데 출근기록 없음";
  matrix[3][11] = "● 휴무·휴가인데 출근기록 있음";
  matrix[4][10] = "● 검토 필요";
  matrix[4][11] = "● 처리 완료";
  matrix[6] = ["No", "지역장", "매니저", "지역", "매장명", "이름", "사번", "발생일", "근무계획", "실제근태", "구분", "상세사유", "처리상태"];

  const regionRows = [];
  const dataRows = [];
  let number = 1;
  for (const region of regionOrder) {
    const group = rows.filter((row) => row.region === region);
    const counts = {
      missing: group.filter((row) => row.category === "근무인데 출근기록 없음" && row.status !== "처리 완료").length,
      unexpected: group.filter((row) => row.category === "휴무·휴가인데 출근기록 있음" && row.status !== "처리 완료").length,
      review: group.filter((row) => row.category === "검토 필요" && row.status !== "처리 완료").length,
      completed: group.filter((row) => row.status === "처리 완료").length,
    };
    const regionRowIndex = matrix.length;
    matrix.push([
      `▼  ${region} (총 ${group.length}건)`, "", "", "", "", "", "", "",
      `근무미입력 ${counts.missing}건  |  휴무출근 ${counts.unexpected}건  |  검토 ${counts.review}건  |  완료 ${counts.completed}건`, "", "", "", "",
    ]);
    regionRows.push({ row: regionRowIndex, region });
    for (const row of group) {
      const actual = row.evidenced
        ? "출근"
        : row.actualStatus || (row.clockStatus && row.clockStatus !== "미기록" ? "출근" : "미입력");
      matrix.push([
        number++, row.member.regionalManager || "", row.member.manager || "", region,
        row.member.storeName || row.store || "", row.name || row.member.employeeName || "", normalizeId(row.employeeId), row.date || "",
        row.planStatus || "공백", actual, row.category, row.reason || "", row.status,
      ]);
      dataRows.push({ row: matrix.length - 1, category: row.category, status: row.status });
    }
  }

  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 1 } },
    { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, { s: { r: 3, c: 2 }, e: { r: 4, c: 3 } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, { s: { r: 3, c: 4 }, e: { r: 4, c: 5 } },
    { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, { s: { r: 3, c: 6 }, e: { r: 4, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, { s: { r: 3, c: 8 }, e: { r: 4, c: 9 } },
    { s: { r: 2, c: 10 }, e: { r: 2, c: 12 } },
    { s: { r: 3, c: 11 }, e: { r: 3, c: 12 } },
    { s: { r: 4, c: 11 }, e: { r: 4, c: 12 } },
  ];
  for (const item of regionRows) {
    sheet["!merges"].push(
      { s: { r: item.row, c: 0 }, e: { r: item.row, c: 7 } },
      { s: { r: item.row, c: 8 }, e: { r: item.row, c: 12 } },
    );
  }

  sheet["!cols"] = [
    { wch: 6 }, { wch: 11 }, { wch: 11 }, { wch: 9 }, { wch: 18 }, { wch: 11 }, { wch: 13 },
    { wch: 13 }, { wch: 17 }, { wch: 15 }, { wch: 26 }, { wch: 46 }, { wch: 12 },
  ];
  sheet["!rows"] = matrix.map((_, index) => ({ hpt: index === 0 ? 34 : index === 1 ? 24 : index >= 2 && index <= 4 ? 28 : index === 5 ? 8 : index === 6 ? 30 : regionRows.some((item) => item.row === index) ? 26 : 24 }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 7, topLeftCell: "A8", activePane: "bottomLeft", state: "frozen" };
  sheet["!views"] = [{ showGridLines: false }];

  styleCellRange(sheet, 0, 0, 0, 12, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 18, bold: true, color: { rgb: "FFFFFFFF" } },
    alignment: { horizontal: "left", vertical: "center" },
  });
  styleCellRange(sheet, 1, 0, 1, 12, {
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

  styleCellRange(sheet, 6, 0, 6, 12, {
    fill: { patternType: "solid", fgColor: { rgb: "FF0B3B76" } },
    font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FFFFFFFF" } },
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
    styleCellRange(sheet, item.row, 0, item.row, 12, {
      fill: { patternType: "solid", fgColor: { rgb: palette.fill } },
      font: { name: "맑은 고딕", sz: 10, bold: true, color: { rgb: palette.font } },
      alignment: { horizontal: "left", vertical: "center" },
      border: thinBorder("FFD6DFEA"),
    });
    const summaryAddr = XLSX.utils.encode_cell({ r: item.row, c: 8 });
    if (sheet[summaryAddr]) sheet[summaryAddr].s.alignment = { horizontal: "right", vertical: "center" };
  }

  dataRows.forEach((item, index) => {
    styleCellRange(sheet, item.row, 0, item.row, 12, {
      fill: { patternType: "solid", fgColor: { rgb: index % 2 ? "FFF9FBFD" : "FFFFFFFF" } },
      font: { name: "맑은 고딕", sz: 9, color: { rgb: "FF1F2937" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: thinBorder("FFDCE3EC"),
    });
    const reasonAddr = XLSX.utils.encode_cell({ r: item.row, c: 11 });
    if (sheet[reasonAddr]) sheet[reasonAddr].s.alignment = { horizontal: "left", vertical: "center", wrapText: true };
    applyMismatchBadgeStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: 10 }), item.category);
    applyProcessStatusStyle(sheet, XLSX.utils.encode_cell({ r: item.row, c: 12 }), item.status);
  });

  setRef(sheet, matrix.length, 13);
  workbook.Sheets["계획&근태 상이 인원"] = sheet;
  if (!workbook.SheetNames.includes("계획&근태 상이 인원")) workbook.SheetNames.push("계획&근태 상이 인원");
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
    border: thinBorder("FFFFFFFF"),
  };
}

function buildDayoffSubstituteSheet(workbook, result, ctx, year, monthNo) {
  const rows = [...(result.employeeSummaries || [])]
    .filter((row) => Number(row.baseExcess || 0) > 0)
    .sort((a, b) => String(a.store).localeCompare(String(b.store), "ko") || String(a.name).localeCompare(String(b.name), "ko"));
  const matrix = [
    [`${year}년 ${monthNo}월 기본 휴무 초과자`],
    ["경로별 기본 휴무 기준을 초과한 상담사와 대체휴무·보상휴가 사용 및 잔여를 함께 정리했습니다."],
    [],
    ["지역장", "매니저", "지역", "매장명", "이름", "사번", "휴무수", "기본휴무", "휴무초과", "대체휴무 필요", "대체휴무 잔여", "대체휴무 초과", "보상휴가 필요", "보상휴가 잔여", "보상휴가 초과", "판정"],
  ];
  for (const row of rows) {
    const member = findMember(ctx, row.employeeId, row.store) || {};
    matrix.push([
      member.regionalManager || "", member.manager || "", member.region2 || member.region1 || "", member.storeName || row.store || "",
      row.name || member.employeeName || "", normalizeId(row.employeeId), roundHalf(Number(row.basicDayoffUsed || 0)), roundHalf(Number(row.baseAllowance || 0)),
      roundHalf(Number(row.baseExcess || 0)), roundHalf(Number(row.substituteNeeded || 0)), roundHalf(Number(row.remainingSubstitute || 0)), roundHalf(Number(row.shortage || 0)),
      roundHalf(Number(row.compensationNeeded || 0)), roundHalf(Number(row.remainingCompensation || 0)), roundHalf(Number(row.compensationShortage || 0)),
      `${row.judgment || ""}${row.compensationJudgment ? ` / ${row.compensationJudgment}` : ""}`,
    ]);
  }
  if (!rows.length) matrix.push(["", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, "확인 대상 없음"]);
  const sheet = XLSX.utils.aoa_to_sheet(matrix);
  sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 15 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 15 } }];
  sheet["!cols"] = [{wch:11},{wch:11},{wch:10},{wch:18},{wch:11},{wch:13},{wch:10},{wch:10},{wch:10},{wch:13},{wch:13},{wch:13},{wch:13},{wch:13},{wch:13},{wch:55}];
  sheet["!rows"] = [{hpt:30},{hpt:24},{hpt:8},{hpt:30}];
  styleSimpleReportSheet(sheet, matrix.length, 16, [15]);
  workbook.Sheets["휴무 초과자"] = sheet;
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
      if (String(item?.display || "").includes("미입력") && item?.attendance && !item.attendance.hasClockIn) {
        const colLetter = XLSX.utils.encode_col(col0);
        const evidenceCount = `(COUNTIFS('출근 미등록'!$G$8:$G$1000,$I${row},'출근 미등록'!$H$8:$H$1000,${colLetter}$5,'출근 미등록'!$K$8:$K$1000,"O")+COUNTIFS('출근 미등록'!$G$8:$G$1000,$I${row},'출근 미등록'!$H$8:$H$1000,${colLetter}$5,'출근 미등록'!$K$8:$K$1000,"○")+COUNTIFS('출근 미등록'!$G$8:$G$1000,$I${row},'출근 미등록'!$H$8:$H$1000,${colLetter}$5,'출근 미등록'!$K$8:$K$1000,"ㅇ"))`;
        setFormula(sheet, address, `IF(${evidenceCount}>0,"출근","${item.display}")`, item.display);
      } else {
        setValue(sheet, address, item?.display || "");
      }
      applyStatusStyle(sheet, address, item?.display || "");
      if (item?.issues?.length) {
        // 미입력 유형은 각각의 구분색과 8pt 글씨를 유지하고, 그 외 오류만 공통 경고색을 적용합니다.
        if (!String(item?.display || "").includes("미입력")) applyIssueStyle(sheet, address, item.issues);
        issueCount += 1;
      }
      // 증빙 O는 계획 상이 여부와 무관하게 상담사근태 셀을 최종 확정색으로 표시합니다.
      if (item?.evidence) applyEvidenceWorkStyle(sheet, address);
      if (NON_WORK_CODES.has(item?.planStatus) && item?.attendance?.hasClockIn) clockCorrection += 1;
    }

    const planValues = Object.values(daily).map((item) => item.planStatus);
    const displayValues = Object.values(daily).map((item) => item?.display || "");
    const registeredCount = displayValues.filter((value) => value && !String(value).includes("미입력")).length;
    const displayedWorkCount = displayValues.filter((value) => value === "출근").length;
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

    setFormula(sheet, addr(0), `COUNTIFS(${dailyRange},"<>",${dailyRange},"<>미입력")`, registeredCount);
    setFormula(sheet, addr(1), `COUNTIF(${dailyRange},"출근")`, displayedWorkCount);
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
  const planMissing = planStatus === "공백";
  const clockMissing = !attendance.hasClockIn;
  if (planMissing && clockMissing) return "출ㆍ계 미입력";
  if (planMissing) return "계획 미입력";
  const actual = normalizeActual(attendance.actualStatus);
  if (actual) return actual === "근무" ? "출근" : actual;
  if (attendance.hasClockIn) return "출근";
  if (["근무", "근무A", "근무B", "근무C", "교육", "오전반차", "오후반차"].includes(planStatus)) return "출근 미입력";
  return planStatus;
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
    map.set(key, {
      hasClockIn: existing.hasClockIn || Boolean(cleanClock(row.actualIn) || cleanClock(row.changedIn)),
      actualIn: cleanClock(row.actualIn) || existing.actualIn,
      changedIn: cleanClock(row.changedIn) || existing.changedIn,
      actualStatus: cleanPlaceholder(row.actualStatus) || existing.actualStatus,
      location: cleanPlaceholder(row.location) || existing.location,
    });
  }
  return map;
}

function emptyAttendance() {
  return { hasClockIn: false, actualIn: "", changedIn: "", actualStatus: "", location: "" };
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
  if (text === "출근") fill = "FFE2F0D9";
  else if (text === "출근 미입력") { fill = "FFF33E0D"; fontColor = "FFFFFFFF"; bold = true; }
  else if (text === "계획 미입력") { fill = "FFFFC000"; fontColor = "FF7F4100"; bold = true; }
  else if (text === "출ㆍ계 미입력") { fill = "FFC00000"; fontColor = "FFFFFFFF"; bold = true; }
  else if (text === "미입력") { fill = "FFF33E0D"; fontColor = "FFFFFFFF"; bold = true; }
  else if (text === "휴무") fill = "FFE7E6E6";
  else if (text.includes("대체휴일") || text === "대체휴무") fill = "FFBFBFBF";
  else if (text.includes("보상휴가")) fill = "FFFFFF00";
  else if (["연차", "공가", "휴가", "경조", "무급휴가"].includes(text)) fill = "FFDDEBF7";
  else if (["오전반차", "오후반차", "반일근무"].includes(text)) fill = "FFFFF2CC";
  else if (text === "교육") fill = "FFDDEBF7";
  if (!fill) return;
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: fill } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: text.includes("미입력") ? 8 : 10, bold, color: { rgb: fontColor } },
  };
}

function applyEvidenceWorkStyle(sheet, address) {
  if (!sheet[address]) return;
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFA9D08E" } },
    font: { ...(base.font || {}), name: "맑은 고딕", sz: 10, bold: true, color: { rgb: "FF000000" } },
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
  if (typeof JSZip === "undefined") return buffer;
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookPath = "xl/workbook.xml";
    const relsPath = "xl/_rels/workbook.xml.rels";
    const stylesPath = "xl/styles.xml";
    const workbookXml = await zip.file(workbookPath)?.async("string");
    const relsXml = await zip.file(relsPath)?.async("string");
    let stylesXml = await zip.file(stylesPath)?.async("string");
    if (!workbookXml || !relsXml || !stylesXml) return buffer;

    const evidencePath = findWorksheetPath(workbookXml, relsXml, "출근 미등록");
    const attendancePath = findWorksheetPath(workbookXml, relsXml, "상담사근태");
    if (!evidencePath || !attendancePath) return buffer;

    const dxfResult = appendConditionalDxfs(stylesXml);
    stylesXml = dxfResult.xml;
    zip.file(stylesPath, stylesXml);

    let evidenceXml = await zip.file(evidencePath)?.async("string");
    let attendanceXml = await zip.file(attendancePath)?.async("string");
    if (!evidenceXml || !attendanceXml) return buffer;

    evidenceXml = removeManagedConditionalFormatting(evidenceXml, "v24-evidence");
    attendanceXml = removeManagedConditionalFormatting(attendanceXml, "v24-attendance");

    const evidenceRules = `<!--v24-evidence-->
<conditionalFormatting sqref="K8:K1000"><cfRule type="expression" dxfId="${dxfResult.evidence}" priority="1"><formula>OR(UPPER(TRIM(K8))=&quot;O&quot;,K8=&quot;○&quot;,K8=&quot;ㅇ&quot;)</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="M8:M1000"><cfRule type="expression" dxfId="${dxfResult.completed}" priority="2"><formula>$M8=&quot;처리 완료&quot;</formula></cfRule><cfRule type="expression" dxfId="${dxfResult.pending}" priority="3"><formula>$M8=&quot;미처리&quot;</formula></cfRule></conditionalFormatting>
<!--/v24-evidence-->`;
    evidenceXml = insertWorksheetConditionalFormatting(evidenceXml, evidenceRules);

    const [year, month] = String(result?.targetMonth || "").split("-").map(Number);
    const days = year && month ? new Date(year, month, 0).getDate() : 31;
    const lastDayColumn = XLSX.utils.encode_col(14 + days - 1);
    const attendanceFormula = `(COUNTIFS('출근 미등록'!$G$8:$G$1000,$I7,'출근 미등록'!$H$8:$H$1000,O$5,'출근 미등록'!$K$8:$K$1000,&quot;O&quot;)+COUNTIFS('출근 미등록'!$G$8:$G$1000,$I7,'출근 미등록'!$H$8:$H$1000,O$5,'출근 미등록'!$K$8:$K$1000,&quot;○&quot;)+COUNTIFS('출근 미등록'!$G$8:$G$1000,$I7,'출근 미등록'!$H$8:$H$1000,O$5,'출근 미등록'!$K$8:$K$1000,&quot;ㅇ&quot;))&gt;0`;
    const attendanceRules = `<!--v24-attendance-->
<conditionalFormatting sqref="O7:${lastDayColumn}500"><cfRule type="expression" dxfId="${dxfResult.attendance}" priority="1"><formula>${attendanceFormula}</formula></cfRule></conditionalFormatting>
<!--/v24-attendance-->`;
    attendanceXml = insertWorksheetConditionalFormatting(attendanceXml, attendanceRules);

    zip.file(evidencePath, evidenceXml);
    zip.file(attendancePath, attendanceXml);
    return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  } catch (error) {
    console.warn("실시간 증빙 색상 규칙 적용 실패", error);
    return buffer;
  }
}

function findWorksheetPath(workbookXml, relsXml, sheetName) {
  const sheetTags = workbookXml.match(/<sheet\b[^>]*\/>/g) || [];
  const tag = sheetTags.find((item) => item.includes(`name="${sheetName}"`));
  const relationId = tag?.match(/r:id="([^"]+)"/)?.[1];
  if (!relationId) return "";
  const relTags = relsXml.match(/<Relationship\b[^>]*\/>/g) || [];
  const rel = relTags.find((item) => item.includes(`Id="${relationId}"`));
  let target = rel?.match(/Target="([^"]+)"/)?.[1] || "";
  if (!target) return "";
  target = target.replace(/^\//, "");
  if (target.startsWith("xl/")) return target;
  return `xl/${target.replace(/^\.\//, "")}`;
}

function appendConditionalDxfs(stylesXml) {
  const dxfs = [
    `<dxf><font><b/><color rgb="FF107C41"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill></dxf>`,
    `<dxf><font><b/><color rgb="FFC00000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFE4E6"/><bgColor indexed="64"/></patternFill></fill></dxf>`,
    `<dxf><font><b/><color rgb="FF107C41"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill></dxf>`,
    `<dxf><font><b/><color rgb="FF000000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFA9D08E"/><bgColor indexed="64"/></patternFill></fill></dxf>`,
  ];
  let start = 0;
  const normal = stylesXml.match(/<dxfs\b[^>]*count="(\d+)"[^>]*>([\s\S]*?)<\/dxfs>/);
  if (normal) {
    start = Number(normal[1] || 0);
    const replacement = normal[0]
      .replace(/count="\d+"/, `count="${start + dxfs.length}"`)
      .replace(/<\/dxfs>$/, `${dxfs.join("")}</dxfs>`);
    stylesXml = stylesXml.replace(normal[0], replacement);
  } else {
    const selfClosing = stylesXml.match(/<dxfs\b[^>]*count="(\d+)"[^>]*\/>/);
    if (selfClosing) {
      start = Number(selfClosing[1] || 0);
      stylesXml = stylesXml.replace(selfClosing[0], `<dxfs count="${start + dxfs.length}">${dxfs.join("")}</dxfs>`);
    } else {
      const block = `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>`;
      stylesXml = stylesXml.includes("<tableStyles")
        ? stylesXml.replace("<tableStyles", `${block}<tableStyles`)
        : stylesXml.replace("</styleSheet>", `${block}</styleSheet>`);
    }
  }
  return { xml: stylesXml, completed: start, pending: start + 1, evidence: start + 2, attendance: start + 3 };
}

function removeManagedConditionalFormatting(xml, marker) {
  const expression = new RegExp(`<!--${marker}-->[\\s\\S]*?<!--\\/${marker}-->`, "g");
  return xml.replace(expression, "");
}

function insertWorksheetConditionalFormatting(xml, rules) {
  if (xml.includes("<pageMargins")) return xml.replace("<pageMargins", `${rules}<pageMargins`);
  if (xml.includes("<pageSetup")) return xml.replace("<pageSetup", `${rules}<pageSetup`);
  return xml.replace("</worksheet>", `${rules}</worksheet>`);
}

function sanitizeWorkbookForExcel(workbook) {
  for (const sheetName of workbook.SheetNames || []) {
    const sheet = workbook.Sheets?.[sheetName];
    if (!sheet) continue;

    // xlsx-js-style의 쓰기 옵션에 showGridLines를 직접 넘기지 않습니다.
    // 지원되는 시트 보기 속성으로만 저장해 파일 구조 손상을 피합니다.
    sheet["!views"] = [{ showGridLines: false }];

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
