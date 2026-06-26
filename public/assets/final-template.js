const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];
const NON_WORK_CODES = new Set([
  "휴무", "무급휴가", "연차", "공가", "휴가", "경조",
  "대체휴일(1일)", "대체휴일(0.5일)", "보상휴가(1일)", "보상휴가(0.5일)",
]);

export async function buildFinalTemplateWorkbook(result) {
  const response = await fetch("./assets/attendance-final-template.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("최종본 엑셀 양식 파일을 불러오지 못했습니다.");
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellStyles: true, cellDates: true, bookVBA: true });

  const monthNo = Number(String(result.targetMonth || "").slice(5, 7));
  const year = Number(String(result.targetMonth || "").slice(0, 4));
  const daysInMonth = new Date(year, monthNo, 0).getDate();
  const mainName = `${monthNo}월 상담사근태`;
  const planName = `${monthNo}월 근무계획`;
  const annualName = `${monthNo}월 연차사용`;

  renameSheet(workbook, "5월 상담사근태", mainName);
  renameSheet(workbook, "5월 근무계획", planName);
  renameSheet(workbook, "5월 연차사용", annualName);

  const context = buildContext(result, daysInMonth);
  fillMainSheet(workbook.Sheets[mainName], result, context, year, monthNo, daysInMonth);
  fillPlanSheet(workbook.Sheets[planName], result, context, daysInMonth);
  fillAnnualSheet(workbook.Sheets[annualName], result, context);
  fillCompensationSheet(workbook.Sheets["보상휴가(0.5일) 반영 인원"], result, context);
  fillEvidenceSheet(workbook.Sheets["증빙(필수기입)"], result, context);
  fillAttendanceRawSheet(workbook.Sheets["근태RAW"], result);
  fillAnnualRawSheet(workbook.Sheets["연차RAW"], result, context);
  fillEducationRawSheet(workbook.Sheets["교육RAW"], result, context);

  workbook.Props = {
    ...(workbook.Props || {}),
    Title: `${year}년 ${monthNo}월 ${result.routeLabel} 상담사 출퇴근현황`,
    Subject: "근태 관리 시스템 자동 생성 최종본",
    Author: "근태 관리 시스템",
    Comments: "인력·매장매칭, 근무계획, 실제 근태를 기준으로 자동 생성",
  };
  return workbook;
}

export async function buildFinalTemplateFile(result) {
  const workbook = await buildFinalTemplateWorkbook(result);
  const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true, bookVBA: true });
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
  const issueMap = new Map();
  for (const issue of result.mismatchRows || []) addIssue(issueMap, issue.employeeId, issue.date, issue.reason || issue.result);
  for (const summary of result.employeeSummaries || []) {
    for (const event of summary.substituteEvents || []) {
      if (event.source === "기본 휴무 초과") addIssue(issueMap, summary.employeeId, event.date, `기본 휴무 기준 ${daysText(summary.baseAllowance)} 초과분`);
    }
    if (Number(summary.shortage || 0) > 0) {
      for (const event of summary.substituteEvents || []) addIssue(issueMap, summary.employeeId, event.date, `대체휴무 잔여 부족 · 총 ${daysText(summary.shortage)} 초과 사용`);
    }
    if (Number(summary.compensationShortage || 0) > 0) {
      for (const event of summary.compensationEvents || []) addIssue(issueMap, summary.employeeId, event.date, `보상휴가 잔여 부족 · 총 ${daysText(summary.compensationShortage)} 초과 사용`);
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
    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = `${result.targetMonth}-${String(day).padStart(2, "0")}`;
      const planStatus = normalizePlan(person.plan?.plans?.[day]);
      const attendance = attendanceByKey.get(`${person.employeeId}|${date}`) || emptyAttendance();
      const issues = issueMap.get(`${person.employeeId}|${day}`) || [];
      daily[day] = {
        date,
        planStatus,
        attendance,
        display: dailyDisplay(planStatus, attendance),
        issues: [...new Set(issues)],
      };
    }
    dailyByKey.set(person.key, daily);
  }

  return { workforce, workforceById, planById, attendanceByKey, summaryById, issueMap, people, dailyByKey };
}

function fillMainSheet(sheet, result, ctx, year, monthNo, daysInMonth) {
  if (!sheet) throw new Error("상담사근태 최종본 시트를 찾지 못했습니다.");
  setValue(sheet, "B2", `■ ${year}년 ${monthNo}월 출퇴근현황`);
  setValue(sheet, "G2", `${monthNo}월 ${ctx.people.length}명`);
  setValue(sheet, "I2", "");
  setValue(sheet, "AT3", `${monthNo}월`);
  setValue(sheet, "AT4", daysInMonth);
  setValue(sheet, "AU4", countWeekdays(year, monthNo));
  setValue(sheet, "AV4", countWeekendDays(year, monthNo));
  setValue(sheet, "N5", `${monthNo}월1일 \n근무계획\n일치확인`);
  setValue(sheet, "AV5", `${monthNo}월 사용가능\n휴무일수`);

  for (let day = 1; day <= 31; day += 1) {
    const col = 14 + day; // O=15 (1-based)
    const address5 = XLSX.utils.encode_cell({ r: 4, c: col - 1 });
    const address6 = XLSX.utils.encode_cell({ r: 5, c: col - 1 });
    if (day <= daysInMonth) {
      const date = new Date(year, monthNo - 1, day);
      setValue(sheet, address5, Number(`${year}${String(monthNo).padStart(2, "0")}${String(day).padStart(2, "0")}`));
      setValue(sheet, address6, WEEKDAYS[date.getDay()]);
    } else {
      clearCell(sheet, address5);
      clearCell(sheet, address6);
    }
  }

  const startRow = 7;
  const lastTemplateRow = 168;
  clearValues(sheet, startRow, 2, lastTemplateRow, 64);
  let row = startRow;
  for (const person of ctx.people) {
    copyRowStyle(sheet, 7, row, 64);
    const member = person.member;
    const summary = ctx.summaryById.get(person.employeeId) || {};
    const daily = ctx.dailyByKey.get(person.key) || {};
    const values = [
      member.regionalManager, member.manager, member.region1, member.region2,
      member.storeCode, member.storeName, member.portalId, person.employeeId,
      person.name, member.hireDate, member.groupHireDate, member.note,
    ];
    values.forEach((value, index) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: index + 1 }), value));

    setValue(sheet, `N${row}`, daily[1]?.planStatus === "공백" ? "" : daily[1]?.planStatus || "");
    let clockCorrection = 0;
    let issueCount = 0;
    for (let day = 1; day <= 31; day += 1) {
      const col0 = 14 + day - 1;
      const address = XLSX.utils.encode_cell({ r: row - 1, c: col0 });
      if (day > daysInMonth) {
        clearCell(sheet, address);
        continue;
      }
      const item = daily[day];
      setValue(sheet, address, item?.display || "");
      if (item?.issues?.length) {
        applyIssueStyle(sheet, address);
        issueCount += 1;
      }
      if (NON_WORK_CODES.has(item?.planStatus) && item?.attendance?.hasClockIn) clockCorrection += 1;
    }

    const planValues = Object.values(daily).map((item) => item.planStatus);
    const workCount = Object.values(daily).filter((item) => item.attendance.hasClockIn).length;
    const dayoffCount = planValues.filter((value) => value === "휴무").length;
    const educationCount = planValues.filter((value) => value === "교육").length;
    const halfCount = planValues.filter((value) => value === "오전반차" || value === "오후반차").length;
    const annualCount = planValues.filter((value) => value === "연차").length;
    const publicCount = planValues.filter((value) => value === "공가").length;
    const unpaidCount = planValues.filter((value) => value === "무급휴가").length;
    const familyCount = planValues.filter((value) => value === "경조").length;
    const evidenceNeeded = (result.missingRows || []).some((item) => normalizeId(item.employeeId) === person.employeeId);
    const noteParts = [];
    if (member.note) noteParts.push(member.note);
    for (let day = 1; day <= daysInMonth; day += 1) {
      if (daily[day]?.issues?.length) noteParts.push(`${monthNo}/${day} ${daily[day].issues.join("/")}`);
    }
    const expectedChecks = Math.max(1, Object.values(daily).filter((item) => item.planStatus !== "공백" || item.attendance.hasClockIn).length);
    const normalRate = Math.max(0, Math.min(1, (expectedChecks - issueCount) / expectedChecks));

    const summaryValues = {
      AT: workCount,
      AU: dayoffCount,
      AV: roundHalf(Number(summary.baseAllowance || 0) + Number(summary.availableSubstitute || 0)),
      AW: roundHalf(Number(summary.shortage || 0)),
      AX: roundHalf(Number(summary.remainingSubstitute || 0)),
      AY: clockCorrection,
      AZ: educationCount,
      BA: halfCount,
      BB: annualCount,
      BC: publicCount,
      BD: unpaidCount,
      BE: familyCount,
      BF: daysInMonth,
      BG: 0,
      BH: roundHalf(Number(summary.currentAnnualLeave || 0)),
      BI: evidenceNeeded ? "O" : "",
      BJ: "",
      BK: [...new Set(noteParts)].join(" · "),
      BL: normalRate,
    };
    Object.entries(summaryValues).forEach(([col, value]) => setValue(sheet, `${col}${row}`, value));
    setNumberFormat(sheet, `BL${row}`, "0.0%");
    row += 1;
  }
  setRef(sheet, Math.max(lastTemplateRow, row - 1), 64);
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

    const reasonValues = new Array(maxCol).fill("");
    let hasIssue = !person.plan;
    if (!person.plan) reasonValues[0] = "빨간색 체크 사유(확인 필요) · 인력·매장매칭에는 있으나 근무계획에 사번 없음";
    for (const [day, col] of dayColumns.entries()) {
      const messages = issueMap.get(`${person.employeeId}|${day}`) || [];
      if (!messages.length) continue;
      hasIssue = true;
      reasonValues[col] = [...new Set(messages)].join(" / ");
      applyIssueStyle(sheet, XLSX.utils.encode_cell({ r: row - 1, c: col }));
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
  const issues = [...(result.mismatchRows || [])].sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.store).localeCompare(String(b.store)));
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
      issue.date || "",
      issue.reason || issue.result || "",
      "",
    ];
    values.forEach((value, valueIndex) => setValue(sheet, XLSX.utils.encode_cell({ r: row - 1, c: valueIndex + 1 }), value));
    row += 1;
  });
  setRef(sheet, Math.max(25, row - 1), 11);
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
  const actual = normalizeActual(attendance.actualStatus);
  if (actual) return actual === "근무" ? "출근" : actual;
  if (attendance.hasClockIn) return "출근";
  if (planStatus === "공백") return "미입력";
  if (["근무", "근무A", "근무B", "근무C", "교육", "오전반차", "오후반차"].includes(planStatus)) return "미입력";
  return planStatus;
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

function applyIssueStyle(sheet, address) {
  if (!sheet[address]) sheet[address] = { t: "s", v: "" };
  const base = sheet[address].s ? clone(sheet[address].s) : {};
  sheet[address].s = {
    ...base,
    fill: { patternType: "solid", fgColor: { rgb: "FFFF0000" } },
    font: { ...(base.font || {}), bold: true, color: { rgb: "FF000000" } },
    alignment: { ...(base.alignment || {}), horizontal: "center", vertical: "center", wrapText: true },
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

function clone(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}
