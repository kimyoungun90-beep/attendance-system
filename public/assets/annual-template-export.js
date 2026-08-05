export async function buildAnnualTemplateWorkbookFile(analysis) {
  if (typeof JSZip === "undefined") throw new Error("JSZip을 불러오지 못했습니다.");
  if (!analysis.annualTemplateFile) {
    throw new Error("연차 엑셀 생성 전, 회사 보안 승인 후 복호화한 전월 연차촉진 마감본을 ‘전월 연차 마감본’ 항목에 업로드해주세요.");
  }
  const source = await analysis.annualTemplateFile.arrayBuffer();
  const zip = await JSZip.loadAsync(source);
  const wbPath = "xl/workbook.xml";
  const relPath = "xl/_rels/workbook.xml.rels";
  let wbXml = await zip.file(wbPath)?.async("string");
  const relXml = await zip.file(relPath)?.async("string");
  if (!wbXml || !relXml) throw new Error("연차촉진 원본 통합문서 구조가 올바르지 않습니다.");

  const sheets = parseSheets(wbXml, relXml);
  const curOver = latest(sheets, /상담사\(재직\s*1년이상\)_([0-9]{1,2})월/, false);
  const curUnder = latest(sheets, /상담사\(재직\s*1년미만\)_([0-9]{1,2})월/, false);
  const prevOver = latest(sheets, /상담사\(재직\s*1년이상\)_([0-9]{1,2})월/, true);
  const prevUnder = latest(sheets, /상담사\(재직\s*1년미만\)_([0-9]{1,2})월/, true);
  if (!curOver || !curUnder) throw new Error("원본에서 재직 1년 이상/미만 시트를 찾지 못했습니다.");

  const month = Number(analysis.month.monthNo);
  const year = Number(analysis.month.year);
  const oldOver = curOver.name;
  const oldUnder = curUnder.name;
  const newOver = `상담사(재직 1년이상)_${month}월`;
  const newUnder = `상담사(재직 1년미만)_${month}월`;

  if (month !== curOver.month && prevOver && prevUnder) {
    await copyEntry(zip, curOver.path, prevOver.path);
    await copyEntry(zip, curUnder.path, prevUnder.path);
    await copyOptional(zip, curOver.relsPath, prevOver.relsPath);
    await copyOptional(zip, curUnder.relsPath, prevUnder.relsPath);
    const prevOverName = `상담사(재직 1년이상)_${curOver.month}월`;
    const prevUnderName = `상담사(재직 1년미만)_${curUnder.month}월`;
    wbXml = renameSheet(wbXml, prevOver.sheetId, prevOverName);
    wbXml = renameSheet(wbXml, prevUnder.sheetId, prevUnderName);
    zip.file(prevOver.path, replaceAll(await zip.file(prevOver.path).async("string"), oldOver, prevOverName));
    zip.file(prevUnder.path, replaceAll(await zip.file(prevUnder.path).async("string"), oldUnder, prevUnderName));
  }
  wbXml = renameSheet(wbXml, curOver.sheetId, newOver);
  wbXml = renameSheet(wbXml, curUnder.sheetId, newUnder);
  wbXml = forceCalc(wbXml);
  zip.file(wbPath, wbXml);

  for (const info of sheets.filter((s) => !s.name.includes("연차사용(상담사)"))) {
    const f = zip.file(info.path);
    if (!f) continue;
    let xml = await f.async("string");
    xml = replaceAll(xml, oldOver, newOver);
    xml = replaceAll(xml, oldUnder, newUnder);
    zip.file(info.path, xml);
  }

  const allPeople = mergePeople(analysis);
  const over = allPeople.filter((p) => !underOneYear(p, analysis));
  const under = allPeople.filter((p) => underOneYear(p, analysis));
  let overXml = await zip.file(curOver.path).async("string");
  let underXml = await zip.file(curUnder.path).async("string");
  zip.file(curOver.path, fillOverSheet(overXml, over, analysis));
  zip.file(curUnder.path, fillUnderSheet(underXml, under, analysis));

  const usageSheet = sheets.find((s) => new RegExp(`(?:${String(year).slice(-2)}|${year})년\\s*연차사용\\(상담사\\)`).test(s.name));
  if (usageSheet && zip.file(usageSheet.path)) {
    const usageXml = await zip.file(usageSheet.path).async("string");
    zip.file(usageSheet.path, fillUsageSheet(usageXml, allPeople, analysis));
  }
  await fillPromotionSheets(zip, sheets, over, under, analysis);

  const out = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return new File([out], `■ ${analysis.routeLabel}_연차촉진 관리용_${month}월 마감본.xlsx`, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function parseSheets(wbXml, relXml) {
  const rels = new Map();
  for (const m of relXml.matchAll(/<Relationship\b([^>]+)\/?\s*>/g)) {
    const a = attrs(m[1]);
    if (a.Id && a.Target) rels.set(a.Id, xlPath(a.Target));
  }
  const out = [];
  for (const m of wbXml.matchAll(/<sheet\b([^>]+)\/?\s*>/g)) {
    const a = attrs(m[1]);
    const path = rels.get(a["r:id"]);
    if (!path) continue;
    const name = unxml(a.name || "");
    const mm = name.match(/_([0-9]{1,2})월/);
    out.push({ name, path, state: a.state || "visible", sheetId: a.sheetId, month: mm ? Number(mm[1]) : 0, relsPath: sheetRels(path) });
  }
  return out;
}
function latest(sheets, re, hidden) {
  return sheets.filter((s) => re.test(s.name) && (hidden ? s.state === "hidden" : s.state !== "hidden")).sort((a, b) => b.month - a.month)[0] || null;
}
function attrs(raw) {
  const o = {};
  for (const m of raw.matchAll(/([\w:.-]+)="([^"]*)"/g)) o[m[1]] = m[2];
  return o;
}
function xlPath(target) {
  const p = target.replace(/^\//, "").replace(/^\.\//, "");
  return p.startsWith("xl/") ? p : `xl/${p}`;
}
function sheetRels(path) {
  const p = path.split("/");
  const file = p.pop();
  return `${p.join("/")}/_rels/${file}.rels`;
}
async function copyEntry(zip, src, dst) {
  const f = zip.file(src);
  if (!f) throw new Error(`원본 내부 파일 누락: ${src}`);
  zip.file(dst, await f.async("uint8array"));
}
async function copyOptional(zip, src, dst) {
  const f = zip.file(src);
  if (f) zip.file(dst, await f.async("uint8array"));
}
function renameSheet(xml, id, name) {
  const re = new RegExp(`(<sheet\\b(?=[^>]*\\bsheetId="${escRe(id)}")(?=[^>]*\\bname=")[^>]*\\bname=")([^"]*)(")`);
  return xml.replace(re, `$1${xmlText(name)}$3`);
}
function forceCalc(xml) {
  if (/<calcPr\b/.test(xml)) {
    return xml.replace(/<calcPr\b([^>]*)\/?\s*>/, (_, raw) => {
      const a = raw.replace(/\s+calcMode="[^"]*"/g, "").replace(/\s+fullCalcOnLoad="[^"]*"/g, "").replace(/\s+forceFullCalc="[^"]*"/g, "").replace(/\s*\/\s*$/, "");
      return `<calcPr${a} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>`;
    });
  }
  return xml.replace(/<\/workbook>/, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
}

function fillOverSheet(xml, people, analysis) {
  const rows = people.map((p, i) => {
    const c = p.closingPerson || {};
    const grant = numBlank(c.currentYearGranted ?? estimateGrant(p, analysis));
    const used = half(n(c.currentYearUsed) + n(p.monthlyAnnualUsed));
    const priorRemain = numBlank(p.annualLedger?.remaining ?? c.remainingAnnual);
    const remain = priorRemain === "" ? (grant === "" ? "" : half(n(grant) - used)) : half(n(priorRemain) - n(p.monthlyAnnualUsed));
    const period = c.annualPeriod || periodFor(p, analysis, false);
    const end = c.annualExhaustionDate || periodEnd(period);
    return {
      B: i + 1, C: c.employmentGroup || groupLabel(p, analysis), D: c.regionalManager || p.regionalManager, E: c.manager || p.manager,
      F: c.region || p.region, G: c.subRegion || p.subRegion, H: c.storeCode || p.storeCode, I: c.storeName || p.storeName,
      J: c.job || "판매상담사", K: c.portalId || p.portalId || "", L: p.employeeId, M: p.employeeName, N: c.phone || p.phone || "",
      O: c.email || p.email || "", P: c.hireDate || p.hireDate, Q: c.groupHireDate || p.groupHireDate, R: c.job || p.job || "",
      AT: grant, AU: used, AV: remain, AW: period, AX: periodStart(period), BA: c.resignDate || p.resignDate || "",
      BR: c.firstPromotionDate || before(end, 180), BS: c.firstPromotionDone || "", BT: c.secondPromotionDate || before(end, 60),
      BU: c.secondPromotionDone || "", BV: end, BW: c.note || p.annualLedger?.note || p.note || "",
    };
  });
  xml = setCell(xml, "B1", `■ 삼성전자 ${analysis.routeLabel} 상담사 연차(1년 이상자)`);
  xml = setCell(xml, "AT7", `${analysis.month.year}년 \n발생`);
  xml = setCell(xml, "AU7", `${analysis.month.year}년 사용\n(입사월 이후)`);
  return replaceRows(xml, 10, 118, rows, "B", "BW");
}
function fillUnderSheet(xml, people, analysis) {
  const rows = people.map((p, i) => {
    const c = p.closingPerson || {};
    const grant = numBlank(c.currentYearGranted ?? estimateGrant(p, analysis));
    const prior = numBlank(c.priorYearUsed);
    const used = half(n(c.currentYearUsed) + n(p.monthlyAnnualUsed));
    const priorRemain = numBlank(p.annualLedger?.remaining ?? c.remainingAnnual);
    const remain = priorRemain === "" ? (grant === "" ? "" : half(n(grant) - n(prior) - used)) : half(n(priorRemain) - n(p.monthlyAnnualUsed));
    const period = c.annualPeriod || periodFor(p, analysis, true);
    const end = c.annualExhaustionDate || periodEnd(period);
    return {
      B: i + 1, C: c.employmentGroup || groupLabel(p, analysis), D: c.regionalManager || p.regionalManager, E: c.manager || p.manager,
      F: c.region || p.region, G: c.subRegion || p.subRegion, H: c.storeCode || p.storeCode, I: c.storeName || p.storeName,
      J: c.job || "판매상담사", K: c.portalId || p.portalId || "", L: p.employeeId, M: p.employeeName, N: c.phone || p.phone || "",
      O: c.email || p.email || "", P: c.hireDate || p.hireDate, Q: grant, R: prior, S: used, T: remain, U: period,
      V: c.resignDate || p.resignDate || "", X: c.firstPromotionDate || before(end, 90), Y: c.firstPromotionDone || "",
      Z: c.secondPromotionDate || before(end, 30), AA: c.secondPromotionDone || "", AB: end, AC: c.note || p.annualLedger?.note || p.note || "",
    };
  });
  xml = setCell(xml, "B1", `■ 삼성전자 ${analysis.routeLabel} 상담사 연차(1년 미만자)`);
  xml = setCell(xml, "Q10", `${analysis.month.year}년 연차\n(입사일 기준 총 발생연차)`);
  xml = setCell(xml, "S10", `${analysis.month.year}년 사용\n(입사월전)`);
  return replaceRows(xml, 13, 56, rows, "B", "AC");
}
function fillUsageSheet(xml, people, analysis) {
  const old = new Map();
  for (const u of analysis.closingData?.annualUsage || []) if (Number(u.year) === Number(analysis.month.year)) old.set(id(u.employeeId), u);
  const rows = people.map((p) => {
    const u = old.get(id(p.employeeId)) || {};
    const months = { ...(u.months || {}) };
    const key = `${analysis.month.monthNo}월`;
    months[key] = half(n(months[key]) + n(p.monthlyAnnualUsed));
    const row = { A: u.portalId || p.closingPerson?.portalId || p.portalId || "", B: p.employeeId, C: p.employeeName,
      Q: u.hireDate || p.closingPerson?.hireDate || p.hireDate || "", R: u.groupHireDate || p.closingPerson?.groupHireDate || p.groupHireDate || "",
      S: u.resignDate || p.closingPerson?.resignDate || p.resignDate || "", T: u.note || p.closingPerson?.note || p.note || "" };
    let total = 0;
    for (let m = 1; m <= 12; m++) { const v = half(n(months[`${m}월`])); row[colName(3 + m)] = v; total += v; }
    row.P = half(total);
    return row;
  });
  return replaceRows(xml, 6, 195, rows, "A", "T");
}
async function fillPromotionSheets(zip, sheets, over, under, analysis) {
  const end = new Date(analysis.month.year, analysis.month.monthNo, 0);
  const choices = new Map([
    ["1년미만 1차", ["C7", candidate(under, "first", end)]],
    ["1년이상 1차", ["C7", candidate(over, "first", end)]],
  ]);
  for (const s of sheets) {
    let cell = "", p = null;
    if (choices.has(s.name)) [cell, p] = choices.get(s.name);
    else if (s.name.includes("1년미만 2차")) { cell = "C6"; p = candidate(under, "second", end); }
    else if (s.name.includes("1년이상 2차")) { cell = "C6"; p = candidate(over, "second", end); }
    else continue;
    const f = zip.file(s.path); if (!f) continue;
    zip.file(s.path, setCell(await f.async("string"), cell, p?.employeeName || ""));
  }
}
function candidate(people, stage, end) {
  return people.filter((p) => {
    const c = p.closingPerson || {};
    const d = date(stage === "first" ? c.firstPromotionDate : c.secondPromotionDate);
    const done = String(stage === "first" ? c.firstPromotionDone : c.secondPromotionDone).trim();
    return d && d <= end && !isDone(done) && !c.resignDate;
  }).sort((a, b) => a.employeeName.localeCompare(b.employeeName, "ko"))[0] || null;
}

function replaceRows(xml, start, end, rows, startCol, endCol) {
  const requiredEnd = Math.max(end, start + rows.length - 1);
  let template = "";
  const seen = new Set();
  xml = xml.replace(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g, (rowXml, rr) => {
    const r = Number(rr); if (r === start) template = rowXml;
    if (r < start || r > requiredEnd) return rowXml;
    seen.add(r); return applyRow(rowXml, r, rows[r - start] || null, startCol, endCol);
  });
  if (!template) template = `<row r="${start}"></row>`;
  const add = [];
  for (let r = start; r <= requiredEnd; r++) if (!seen.has(r)) add.push(applyRow(changeRow(template, start, r), r, rows[r - start] || null, startCol, endCol));
  if (add.length) xml = xml.replace(/<\/sheetData>/, `${add.join("")}</sheetData>`);
  xml = xml.replace(/(<dimension\b[^>]*\bref="[A-Z]{1,3}\d+:)([A-Z]{1,3})(\d+)(")/, (_, a, c, r, q) => `${a}${c}${Math.max(Number(r), requiredEnd)}${q}`);
  xml = xml.replace(/(<autoFilter\b[^>]*\bref="[A-Z]{1,3}\d+:)([A-Z]{1,3})(\d+)(")/, (_, a, c, r, q) => `${a}${c}${Math.max(start, start + rows.length - 1)}${q}`);
  return xml;
}
function applyRow(rowXml, rowNo, data, startCol, endCol) {
  const wanted = new Map(Object.entries(data || {})), written = new Set();
  const min = colNo(startCol), max = colNo(endCol);
  let boundary = rowXml.length;
  const tags = new RegExp(`<c\\b[^>]*\\br="([A-Z]{1,3})${rowNo}"`, "g");
  let tm; while ((tm = tags.exec(rowXml))) if (colNo(tm[1]) > max) { boundary = tm.index; break; }
  const head = rowXml.slice(0, boundary), tail = rowXml.slice(boundary);
  const cellRe = /<c\b([^>]*)\br="([A-Z]{1,3})(\d+)"([^>]*)\/>|<c\b([^>]*)\br="([A-Z]{1,3})(\d+)"([^>]*)>[\s\S]*?<\/c>/g;
  let out = head.replace(cellRe, (all, a1, c1, r1, a2, a3, c2, r2, a4) => {
    const c = c1 || c2, r = Number(r1 || r2), cn = colNo(c);
    if (r !== rowNo || cn < min || cn > max) return all;
    const at = `${a1 || a3 || ""} r="${c}${rowNo}"${a2 || a4 || ""}`.replace(/\s+t="[^"]*"/g, "").replace(/\s*\/\s*$/, "").replace(/\s+/g, " ");
    if (!wanted.has(c)) return `<c${at}></c>`;
    written.add(c); const body = cellBody(wanted.get(c)); return `<c${at}${body.type}>${body.xml}</c>`;
  });
  const missing = [];
  for (const [c, v] of wanted) if (!written.has(c)) { const b = cellBody(v); missing.push(`<c r="${c}${rowNo}"${b.type}>${b.xml}</c>`); }
  return out + missing.join("") + tail;
}
function changeRow(rowXml, oldR, newR) {
  return rowXml.replace(new RegExp(`(<row\\b[^>]*\\br=")${oldR}(")`), `$1${newR}$2`).replace(new RegExp(`(\\br="[A-Z]{1,3})${oldR}(")`, "g"), `$1${newR}$2`);
}
function setCell(xml, ref, value) {
  const row = Number(ref.match(/\d+/)?.[0]);
  const reRow = new RegExp(`<row\\b[^>]*\\br="${row}"[^>]*>[\\s\\S]*?<\\/row>`);
  const m = xml.match(reRow); if (!m) return xml;
  return xml.replace(reRow, setCellInRow(m[0], ref, value));
}
function setCellInRow(rowXml, ref, value) {
  const re = new RegExp(`<c\\b([^>]*)\\br="${escRe(ref)}"([^>]*)\\/>|<c\\b([^>]*)\\br="${escRe(ref)}"([^>]*)>[\\s\\S]*?<\\/c>`);
  const m = rowXml.match(re);
  const at = m ? `${m[1] || m[3] || ""} r="${ref}"${m[2] || m[4] || ""}`.replace(/\s+t="[^"]*"/g, "").replace(/\s*\/\s*$/, "").replace(/\s+/g, " ") : ` r="${ref}"`;
  const b = cellBody(value), cell = `<c${at}${b.type}>${b.xml}</c>`;
  return m ? rowXml.replace(re, cell) : rowXml.replace(/<\/row>/, `${cell}</row>`);
}
function cellBody(v) {
  if (v === "" || v === null || v === undefined) return { type: "", xml: "" };
  if (typeof v === "number" && Number.isFinite(v)) return { type: "", xml: `<v>${v}</v>` };
  return { type: ' t="inlineStr"', xml: `<is><t xml:space="preserve">${xmlText(String(v))}</t></is>` };
}

function mergePeople(analysis) {
  const map = new Map();
  for (const c of analysis.closingData?.people || []) { const k = id(c.employeeId); if (k) map.set(k, { ...c, employeeId: k, closingPerson: c, monthlyAnnualUsed: 0 }); }
  for (const p of analysis.people || []) { const k = id(p.employeeId); if (!k) continue; const old = map.get(k) || {}; map.set(k, { ...old, ...p, employeeId: k, closingPerson: p.closingPerson || old.closingPerson || null }); }
  return [...map.values()].filter((p) => p.employeeId && p.employeeName).sort(personSort);
}
function underOneYear(p, analysis) {
  const g = String(p.closingPerson?.employmentGroup || p.employmentGroup || "").replace(/\s+/g, "");
  if (g.includes("미만")) return true; if (g.includes("이상")) return false;
  const h = date(p.closingPerson?.hireDate || p.hireDate); if (!h) return false;
  const anniversary = new Date(h.getFullYear() + 1, h.getMonth(), h.getDate());
  return new Date(analysis.month.year, analysis.month.monthNo, 0) < anniversary;
}
function groupLabel(p, analysis) {
  if (underOneYear(p, analysis)) return `${String(analysis.month.year).slice(-2)}년이후`;
  const h = date(p.closingPerson?.hireDate || p.hireDate);
  return h && h.getFullYear() >= analysis.month.year - 1 ? `${String(h.getFullYear()).slice(-2)}년이후` : `${String(analysis.month.year - 3).slice(-2)}년이전`;
}
function estimateGrant(p, analysis) {
  const h = date(p.closingPerson?.groupHireDate || p.groupHireDate || p.closingPerson?.hireDate || p.hireDate); if (!h) return "";
  const ref = new Date(analysis.month.year, analysis.month.monthNo, 0);
  const months = Math.max(0, (ref.getFullYear() - h.getFullYear()) * 12 + ref.getMonth() - h.getMonth());
  if (months < 12) return Math.min(months, 11);
  const years = Math.floor(months / 12); return Math.min(15 + Math.floor(Math.max(0, years - 1) / 2), 25);
}
function periodFor(p, analysis, under) {
  const h = date(p.closingPerson?.hireDate || p.hireDate); if (!h) return "";
  const start = under ? new Date(h) : new Date(analysis.month.year, h.getMonth(), h.getDate());
  if (!under && start > new Date(analysis.month.year, analysis.month.monthNo, 0)) start.setFullYear(start.getFullYear() - 1);
  const end = new Date(start); end.setFullYear(end.getFullYear() + 1); end.setDate(end.getDate() - 1);
  return `${fmt(start)}~${fmt(end)}`;
}
function periodStart(p) { return String(p || "").split("~")[0] || ""; }
function periodEnd(p) { return String(p || "").split("~")[1] || ""; }
function before(end, days) { const d = date(end); if (!d) return ""; d.setDate(d.getDate() - days); return fmt(d); }
function date(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return new Date(v);
  const m = String(v || "").trim().replace(/[.]/g, "-").replace(/[년월]/g, "-").replace(/일/g, "").match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}
function fmt(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; }
function isDone(v) { const s = String(v || "").trim().toUpperCase(); return ["O", "OK", "Y", "YES", "완료", "확인", "문자", "메일"].includes(s) || s.includes("완료"); }
function n(v) { const x = Number(String(v ?? "").replace(/[^\d.-]/g, "")); return Number.isFinite(x) ? x : 0; }
function half(v) { return Math.round(Number(v || 0) * 2) / 2; }
function numBlank(v) { if (v === "" || v === null || v === undefined) return ""; const x = Number(String(v).replace(/[^\d.-]/g, "")); return Number.isFinite(x) ? half(x) : ""; }
function id(v) { const s = String(v || "").trim(); return /^\d+(?:\.0+)?$/.test(s) ? String(Number(s)) : s.replace(/\s+/g, ""); }
function personSort(a, b) { return [a.regionalManager, a.manager, a.region, a.storeName, a.employeeName].join("|").localeCompare([b.regionalManager, b.manager, b.region, b.storeName, b.employeeName].join("|"), "ko"); }
function colNo(s) { let n = 0; for (const c of s) n = n * 26 + c.charCodeAt(0) - 64; return n; }
function colName(n) { let s = ""; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode((n - 1) % 26 + 65) + s; return s; }
function replaceAll(s, a, b) { return a ? s.split(a).join(b) : s; }
function xmlText(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;"); }
function unxml(s) { return String(s).replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&gt;/g, ">").replace(/&lt;/g, "<").replace(/&amp;/g, "&"); }
function escRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
