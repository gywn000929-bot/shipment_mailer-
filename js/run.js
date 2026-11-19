/* 출하 요청 / 긴급 자재 / 주문서 메일 만들기 */
"use strict";
let RESULT = { mails: [], unmatched: [], kind: "" };   // mails: [{kind, vendor, subject, to, cc, body, attachments:[{name,data}], keys, include, noTo}]

// 긴급 자재 표: 에어리스트의 그 업체 품목을 품목 코드별로 합치고, 선박리스트의 미납수량 합계를 붙인다.
// 붉은 칸: 부족일정이 URGENT(날짜가 아닌 글자)이거나 기준 주 안(또는 이미 지남)인 품목.
function urgentItems(airRows, shipAll, weekEnd) {
  const map = new Map();
  for (const r of airRows) {
    const k = norm(r["Item cd"]);
    if (!map.has(k)) map.set(k, { name: r["Part Name"] || r["Item cd"], air: 0, sum: 0, dates: [], texts: [] });
    const it = map.get(k); it.air += toNum(r["Purchase request"]);
    const d = toDate(r["Short Del."]); if (d) it.dates.push(d); else if (clean(r["Short Del."])) it.texts.push(clean(r["Short Del."]));
  }
  for (const r of shipAll) { const it = map.get(norm(r["ITEM CD"])); if (it) it.sum += r._sum || 0; }
  const items = [...map.values()].map((it) => {
    const first = it.dates.sort((a, b) => a - b)[0];
    return { name: it.name, air: it.air, sum: it.sum, short: it.texts[0] || (first ? fmtCellDate(first) : ""), red: it.texts.length > 0 || (!!first && first < weekEnd), _d: it.texts.length ? 0 : first ? +first : 9e15 };
  });
  return items.sort((a, b) => a._d - b._d);
}

// ---------------------------------------------------------------- 출하 요청 실행
async function runShip() {
  if (!CONFIG) { toast("설정을 먼저 불러와야 합니다"); return; }
  const fShip = $("ship").files[0], fAir = $("air").files[0];
  if (!fShip && !fAir) { toast("선박리스트나 에어리스트 중 하나는 넣어야 합니다."); return; }
  clearLog(); setStatus("실행 중...", true); RESULT = { mails: [], unmatched: [], kind: "ship", extraFiles: [] };
  try {
    const S = CONFIG.settings;
    let monday = $("monday").value ? parseISO($("monday").value) : nextMonday();
    const wd = (monday.getDay() + 6) % 7;
    if (wd) { log(`[경고] ${fmtDate(monday)}는 월요일이 아닙니다. 그 주 월요일로 맞춥니다.`, "warn"); monday = addDays(monday, -wd); }
    const [month, week] = weekLabel(monday); RESULT.weekKey = fmtDate(monday);
    const airDate = addDays(monday, 1), shipDate = addDays(monday, 2);
    const subjectBase = (S["제목형식"] || "대아홀딩스 {월}월 {주}주 출하 요청의 건").replace("{월}", month).replace("{주}", week);
    const includeZero = !["N", "NO", "X", "0"].includes(clean(S["수량없는행포함"] ?? "Y").toUpperCase());
    const etdCount = parseInt($("etd").value) || parseInt(S["ETD열개수"]) || 2;
    const only = splitList($("only").value);
    const commonCc = splitList(S["공통참조"]);

    let ship = [], etdSel = [], air = [], hasSum = false;
    const wbShip = fShip ? XLSX.read(await fShip.arrayBuffer(), { type: "array", cellDates: true }) : null;
    const wbAir = fAir ? (fAir === fShip ? wbShip : XLSX.read(await fAir.arrayBuffer(), { type: "array", cellDates: true })) : null;
    if (wbShip) ({ rows: ship, etdSel, hasSum } = parseShip(wbShip, etdCount));
    if (wbAir) air = parseAir(wbAir);
    else if (wbShip) { try { air = parseAir(wbShip); log(`[참고] 에어리스트를 따로 넣지 않아, 선박리스트 파일 안의 에어 표를 함께 사용합니다. (에어를 빼려면 고급 설정 → "같은 파일의 에어 표 자동 사용" 해제)`, "warn"); } catch (e) { log("에어리스트 없음 (선박 파일에 에어 표도 없음) → 선박만 진행"); } }
    if (!wbShip && wbAir) { try { ({ rows: ship, etdSel, hasSum } = parseShip(wbAir, etdCount)); log("[참고] 선박리스트를 따로 넣지 않아, 에어리스트 파일 안의 선박 표를 함께 사용합니다.", "warn"); } catch (e) { log("선박리스트 없음 → 에어만 진행"); } }
    if (!$("autoBoth").checked) { if (!fAir) air = []; if (!fShip) { ship = []; etdSel = []; } }
    const shipAll = ship;
    if (ship.length && !includeZero) ship = ship.filter((r) => r._hasQty);
    const unShip = ship.filter((r) => !r._vendor && r._hasQty), unAir = air.filter((r) => !r._vendor);

    log(""); log("==================== 미리보기 ====================", "head");
    log(`기준 주: ${fmtDate(monday)} (월)   제목: ${subjectBase}`);
    log(`에어 입고 요망: ${koDate(airDate)}   선박 입고 요망: ${koDate(shipDate)}`);
    log(`공통 참조: ${commonCc.join(", ")}`); log("");

    const names = [...new Set([...ship.map((r) => r._vendor), ...air.map((r) => r._vendor)])].filter(Boolean);
    const problems = [], unknownVendors = [];
    for (const n of names) {
      let s = ship.filter((r) => r._vendor === n); const a = air.filter((r) => r._vendor === n);
      const v = CONFIG.vendors[n];
      if (!v) { problems.push(`'${n}' : 리스트의 업체 열 값이 설정 업체 시트에 없습니다 (선박 ${s.length} / 에어 ${a.length})`); unknownVendors.push({ key: n, kind: "vendor", count: s.length + a.length, samples: [...s.map((r) => r["ITEM CD"] + " " + r["ITEM NM"]), ...a.map((r) => r["Item cd"] + " " + r["Part Name"])].slice(0, 3) }); continue; }
      if (only.length && !only.includes(n)) continue;
      if (!v.sendShip) { if (s.some((r) => r._hasQty) || a.length) (RESULT.skipped = RESULT.skipped || []).push(`${n} (선박 ${s.filter((r) => r._hasQty).length} / 에어 ${a.length})`); continue; }
      const noTo = !v.to.length;
      if (noTo) problems.push(`'${n}' : 받는 사람이 비어 있습니다 (선박 ${s.length} / 에어 ${a.length}) → 아래 결과 칸에서 메일 주소를 적어 주세요`);
      const nShipQty = s.filter((r) => r._hasQty).length;
      if (!nShipQty && !a.length) continue;
      const subject = subjectBase + (a.length && !nShipQty ? (S["에어만제목접미"] ?? "(에어)") : "");
      if (!nShipQty) s = [];
      log(`■ ${n}  선박 ${nShipQty}건${s.length !== nShipQty ? `(전체 ${s.length}행)` : ""} / 에어 ${a.length}건  ${v.attach ? "[엑셀 첨부]" : "[본문 표]"}  제목: ${subject}`);
      log(`   받는 사람: ${v.to.join(", ")}`); if (v.cc.length) log(`   추가 참조: ${v.cc.join(", ")}`); if (v.memo) log(`   메모: ${v.memo}`, "warn");
      const attached = v.attach && s.length > 0;
      const attachments = attached ? [{ name: safeName(`${subjectBase}_${v.name}.xlsx`), data: vendorXlsx(s, a, etdSel) }] : [];
      RESULT.mails.push({ kind: "ship", vendor: n, subject, to: v.to, cc: uniq([...commonCc, ...v.cc]), body: buildShipBody(v, a, s, etdSel, airDate, shipDate, S, attached), attachments, note: v.memo ? "메모: " + v.memo : "", ship: s, air: a, noTo, include: !noTo, keys: [`ship|${RESULT.weekKey}|${n}`] });
      RESULT.etdSel = etdSel; RESULT.tag = `${month}월${week}주`;
    }
    // ---- 긴급 자재 메일 (설정 "긴급메일업체"에 적힌 업체. 에어리스트에 올라온 그 업체 품목을 표로)
    for (const un of splitList(S["긴급메일업체"])) {
      const v = CONFIG.vendors[matchAlias(un)]; if (!v || (only.length && !only.includes(v.name))) continue;
      const items = urgentItems(air.filter((r) => r._vendor === v.name), shipAll, addDays(monday, 7));
      if (!items.length) { log(`■ ${v.name} 긴급 자재: 에어리스트에 해당 업체 품목이 없어 만들지 않음`); continue; }
      const subject = (S["긴급제목형식"] || "대아홀딩스 {월}월 {주}주 긴급 자재의 건").replace("{월}", month).replace("{주}", week).replace("{업체}", v.poShort);
      log(`■ ${v.name} 긴급 자재  ${items.length}품목 (붉은 칸 ${items.filter((x) => x.red).length})  제목: ${subject}`);
      if (!hasSum) log("   [참고] 선박리스트에 SUM 열이 없어 '미납수량 합계'를 ETD 열 합으로 계산했습니다.", "warn");
      const noSum = items.filter((x) => !x.sum).length; if (noSum) log(`   [확인] 선박리스트에 없는 품목 ${noSum}건은 미납수량 합계가 비어 있습니다.`, "warn");
      RESULT.mails.push({ kind: "urgent", vendor: v.name, subject, to: v.to, cc: uniq([...commonCc, ...v.cc]), body: buildUrgentBody(v, items, S), attachments: [], note: "긴급 자재 메일: 표의 수량·부족일정을 보내기 전에 한 번 더 확인하세요", noTo: !v.to.length, include: v.to.length > 0, keys: [`urgent|${RESULT.weekKey}|${v.name}`] });
    }
    log("");
    if (RESULT.skipped && RESULT.skipped.length) log(`[출하 요청 메일을 보내지 않는 업체 - 건너뜀] ${RESULT.skipped.join(", ")}`);
    if (problems.length) { log("[문제] 아래 업체는 메일을 만들지 않았습니다:", "warn"); problems.forEach((p) => log("   " + p, "warn")); }
    if (unShip.length || unAir.length) {
      log(`[미지정 품목] 선박 ${unShip.length}건 / 에어 ${unAir.length}건 → 메일 생성 안 함 (아래 미지정 업체 지정하기에서 업체를 지정하세요)`, "warn");
      unShip.forEach((r) => log(`   선박  ${r["ITEM CD"].padEnd(20)} ${r.MAKER.padEnd(18)} ${r["ITEM NM"]}`));
      unAir.forEach((r) => log(`   에어  ${r["Item cd"].padEnd(20)} ${r.Maker.padEnd(18)} ${r["Part Name"]}`));
      RESULT.extraFiles.push({ name: `미지정_품목_${fmtDate(monday).replace(/-/g, "")}_${month}월${week}주.xlsx`, data: unmappedXlsx(unShip, unAir, etdSel) });
      const grp = {};
      for (const r of unShip) { const k = r.MAKER || "(MAKER 없음)"; (grp[k] = grp[k] || []).push(r["ITEM CD"] + " " + r["ITEM NM"]); }
      for (const r of unAir) { const k = r.Maker || "(MAKER 없음)"; (grp[k] = grp[k] || []).push(r["Item cd"] + " " + r["Part Name"]); }
      RESULT.unmatched = Object.entries(grp).map(([k, v]) => ({ key: k, kind: "maker", count: v.length, samples: v.slice(0, 3) }));
    }
    RESULT.unmatched.push(...unknownVendors);
    log("=================================================="); log(`완료: 메일 ${RESULT.mails.length}통 준비됨. 아래에서 받으세요.`, "ok");
    setStatus("완료");
  } catch (e) { console.error(e); log("[오류] " + (e.message || e), "err"); setStatus("오류"); }
  renderResult();
  $("resultCard").scrollIntoView({ behavior: "smooth" });
}

// ---------------------------------------------------------------- 주문서 PDF
async function readPdfPages(buf) {
  const pdf = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const tc = await page.getTextContent();
    const words = tc.items.filter((it) => it.str && it.str.trim()).map((it) => ({ text: it.str, x0: it.transform[4], x1: it.transform[4] + it.width, y: it.transform[5] }));
    const LABELS = ["거래처명", "발주번호", "발주일자", "발주금액", "대표자명", "등록번호", "납기일자", "담당자"];
    const rightOf = (label) => {   // 같은 줄에서 label 오른쪽 토큰들을 다음 label 전까지 이어 붙인 문자열
      const w = words.find((x) => x.text.replace(/\s/g, "") === label);
      if (!w) return "";
      const same = words.filter((x) => Math.abs(x.y - w.y) < 3 && x.x0 > w.x1 - 0.5).sort((a, b) => a.x0 - b.x0);
      const out = [];
      for (const x of same) { if (LABELS.includes(x.text.replace(/\s/g, ""))) break; out.push(x.text.trim()); }
      return out.join(" ").replace(/^[\s:：]+/, "").trim();
    };
    const vendRaw = rightOf("거래처명"), po = rightOf("발주번호");
    let code = ""; const mc = vendRaw.match(/\((\d+)\s*\)/); if (mc) code = mc[1];
    let name = vendRaw.split(/\(\s*\d+/)[0].replace(/\(\s*\)/g, "").replace(/\s+주\s*$/, "").trim();
    let poNo = ""; if (po) { const m = po.match(/PO-?0*(\d+)/); if (m) poNo = m[1]; }
    if (!poNo) { const m = words.map((w) => w.text).join(" ").match(/PO-?0*(\d+)/); poNo = m ? m[1] : ""; }
    // 거래처명이 같은 글자덩어리 안에 있을 때 ("거래처명 : 기성전선(주)")
    if (!name) { const w = words.find((x) => x.text.includes("거래처명")); if (w) { const t = w.text.split("거래처명")[1].replace(/^[\s:：]+/, "").trim(); if (t) name = t.split(/\s+/)[0]; } }
    pages.push({ page: i - 1, vendor: name, code, po: poNo });
  }
  return pages;
}
async function splitPdf(buf, pageIdxs) {
  const src = await PDFLib.PDFDocument.load(buf);
  const out = await PDFLib.PDFDocument.create();
  const copied = await out.copyPages(src, pageIdxs);
  copied.forEach((p) => out.addPage(p));
  return await out.save();
}
function findVendorByPo(poVendorName) {
  const key = normName(poVendorName);
  for (const n of CONFIG.order) { const v = CONFIG.vendors[n]; if (v.poNames.map(normName).includes(key)) return v; }
  for (const n of CONFIG.order) { const vn = normName(n); if (vn && (key.includes(vn) || vn.includes(key))) return CONFIG.vendors[n]; }
  return null;
}
async function runPo() {
  if (!CONFIG) { toast("설정을 먼저 불러와야 합니다"); return; }
  const files = [...$("pdf").files];
  if (!files.length) { toast("주문서 PDF를 넣어야 합니다."); return; }
  clearLog(); setStatus("실행 중...", true); RESULT = { mails: [], unmatched: [], kind: "po", extraFiles: [] };
  try {
    const S = CONFIG.settings, commonCc = splitList(S["공통참조"]);
    const poPages = new Map(), order = [], seenPo = new Map(), warned = new Set();   // key f|po -> {buf, pages:[], vendor, code}
    for (const f of files) {
      const buf = await f.arrayBuffer();
      for (const pg of await readPdfPages(buf)) {
        if (!pg.po) { log(`[경고] ${f.name} ${pg.page + 1}페이지: 발주번호를 못 읽었습니다. 건너뜁니다.`, "warn"); continue; }
        const key = f.name + "|" + pg.po;
        if (!poPages.has(key) && seenPo.has(pg.po)) { if (!warned.has(key)) { warned.add(key); log(`[중복] PO ${pg.po} 가 ${seenPo.get(pg.po)} 와 ${f.name} 두 파일에 모두 있습니다 → ${f.name} 쪽은 건너뜁니다.`, "warn"); } continue; }
        seenPo.set(pg.po, f.name);
        if (!poPages.has(key)) { poPages.set(key, { buf, pages: [], vendor: pg.vendor, code: pg.code, po: pg.po, file: f.name }); order.push(key); }
        poPages.get(key).pages.push(pg.page);
      }
    }
    if (!poPages.size) throw new Error("주문서를 하나도 읽지 못했습니다. (스캔 이미지 PDF면 텍스트가 없어 읽을 수 없습니다)");
    const plan = {}, skipped = [], unknown = []; RESULT.poTable = [];
    for (const key of order) {
      const e = poPages.get(key); const v = findVendorByPo(e.vendor);
      const row = { po: e.po, pages: e.pages.length, pdfVendor: e.vendor, vendor: v ? v.name : "", status: "" };
      RESULT.poTable.push(row);
      if (!v) { unknown.push(e); row.status = "업체 못 찾음"; continue; }
      if (!v.sendPo) { skipped.push(e); row.status = "발주서 미발송 업체(건너뜀)"; continue; }
      row.status = v.to.length ? "메일 생성" : "받는 사람 없음";
      (plan[v.name] = plan[v.name] || { vendor: v, entries: [], poVendor: e.vendor }).entries.push(e);
    }
    log("==================== 미리보기 ====================", "head");
    log(`공통 참조: ${commonCc.join(", ")}`); log("");
    for (const [n, p] of Object.entries(plan)) {
      const v = p.vendor, pos = p.entries.map((e) => e.po);
      const subject = (S["주문서제목형식"] || "대아 주문서-{PO}").replace("{PO}", pos.join(",")).replace("{업체}", v.poShort);
      const attachments = [];
      for (const e of p.entries) attachments.push({ name: safeName((S["주문서첨부이름"] || "{PO}.pdf").replace("{PO}", e.po).replace("{업체}", v.poShort)), data: await splitPdf(e.buf, e.pages) });
      log(`■ ${n} (PDF 거래처명: ${p.poVendor})  PO ${pos.join(", ")}  제목: ${subject}`);
      const note = v.to.length ? "" : "[받는 사람 없음] 설정 업체 시트에 메일을 적어 주세요";
      if (v.to.length) log(`   받는 사람: ${v.to.join(", ")}`); else log("   " + note, "warn");
      if (v.cc.length) log(`   추가 참조: ${v.cc.join(", ")}`);
      RESULT.mails.push({ kind: "po", vendor: n, subject, to: v.to, cc: uniq([...commonCc, ...v.cc]), body: buildPoBody(v, S), attachments, note, noTo: !v.to.length, include: v.to.length > 0, pos, keys: pos.map((p) => "po|" + p) });
    }
    RESULT.skipped = skipped.map((e) => `PO ${e.po} ${e.vendor}`);
    if (skipped.length) { log(""); log("[발주서 미발송 업체 - 건너뜀]"); skipped.forEach((e) => log(`   PO ${e.po}  ${e.vendor}`)); }
    if (unknown.length) {
      log(""); log("[업체를 못 찾음 - 아래 미지정 업체 지정하기에서 지정하세요]", "warn"); unknown.forEach((e) => log(`   PO ${e.po}  거래처명: ${e.vendor}  (코드 ${e.code})`, "warn"));
      const grp = {}; for (const e of unknown) (grp[e.vendor || "(거래처명 없음)"] = grp[e.vendor || "(거래처명 없음)"] || []).push("PO " + e.po);
      RESULT.unmatched = Object.entries(grp).map(([k, v]) => ({ key: k, kind: "po", count: v.length, samples: v }));
      for (const e of unknown) RESULT.extraFiles.push({ name: `미지정_${e.po}.pdf`, data: await splitPdf(e.buf, e.pages) });
    }
    log("=================================================="); log(`완료: 메일 ${RESULT.mails.length}통 준비됨.`, "ok"); setStatus("완료");
  } catch (e) { console.error(e); log("[오류] " + (e.message || e), "err"); setStatus("오류"); }
  renderResult(); $("resultCard").scrollIntoView({ behavior: "smooth" });
}
