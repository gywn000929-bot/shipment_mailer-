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
// fromPaste: "금요일 에어 메일" 칸에 붙여 넣은 표로 만든다(위 엑셀 칸은 보지 않음)
async function runShip(fromPaste = false) {
  if (!CONFIG) { toast("설정을 먼저 불러와야 합니다"); return; }
  const pasted = fromPaste ? clean($("airPaste").value) : "", fRemain = fromPaste ? $("remainFile").files[0] : null;
  const fShip = fromPaste ? null : $("ship").files[0], fAir = fromPaste ? null : $("air").files[0];
  if (fromPaste && !pasted && !fRemain) { toast("먼저 큰 칸에 메일 본문의 표를 붙여 넣거나(Ctrl+V) REMAIN 엑셀을 넣으세요."); $("airPaste").focus(); return; }
  if (!fromPaste && !fShip && !fAir) { toast("선박리스트나 에어리스트 엑셀을 넣어야 합니다. (메일 본문의 표는 아래 '금요일 에어 메일' 칸에 붙여 넣으세요)"); return; }
  clearLog(); setStatus("실행 중...", true); RESULT = { mails: [], unmatched: [], kind: "ship", extraFiles: [] };
  try {
    const S = CONFIG.settings;
    const includeZero = !["N", "NO", "X", "0"].includes(clean(S["수량없는행포함"] ?? "Y").toUpperCase());
    const etdCount = parseInt($("etd").value) || parseInt(S["ETD열개수"]) || 2;
    let ship = [], etdSel = [], air = [], hasSum = false;
    const wbShip = fShip ? XLSX.read(await fShip.arrayBuffer(), { type: "array", cellDates: true }) : null;
    let wbAir = fAir ? (fAir === fShip ? wbShip : XLSX.read(await fAir.arrayBuffer(), { type: "array", cellDates: true })) : null;
    if (fromPaste && fRemain) { wbAir = XLSX.read(await fRemain.arrayBuffer(), { type: "array", cellDates: true }); log(`[수요일 REMAIN] 엑셀 '${fRemain.name}'로 만듭니다.`, "head"); }
    else if (fromPaste) { wbAir = pastedWorkbook($("airPaste").value); log("[수요일 REMAIN] 붙여 넣은 표로 만듭니다.", "head"); }
    // 선박 표가 있으면 월요일 주간 리스트(화요일 에어), 없으면 수요일 REMAIN 리스트(금요일 에어)로 본다
    let shipErr = null;
    const tryShip = (wb) => { try { return parseShip(wb, etdCount); } catch (e) { shipErr = shipErr || e; return null; } };
    let shipRes = wbShip ? tryShip(wbShip) : null, shipFrom = wbShip;
    if (!shipRes && wbAir && !fromPaste && $("autoBoth").checked) { shipRes = tryShip(wbAir); shipFrom = wbAir; }   // 수요일 REMAIN 메일에서는 선박 표를 보지 않는다
    const friday = !shipRes;
    const airWb = wbAir || (friday || $("autoBoth").checked ? wbShip : null);
    if (airWb) {
      try { air = parseAir(airWb, { friday, extraWbs: [wbShip, wbAir].filter((w) => w && w !== airWb) }); }
      catch (e) { if (wbAir || friday) throw shipErr && !wbAir ? new Error(shipErr.message + " / " + e.message) : e; log("에어리스트 없음 (선박 파일에 에어 표도 없음) → 선박만 진행"); }
      if (!wbAir && !friday && air.length) log(`[참고] 에어리스트를 따로 넣지 않아, 선박리스트 파일 안의 에어 표를 함께 사용합니다. (에어를 빼려면 고급 설정 → "같은 파일의 에어 표 자동 사용" 해제)`, "warn");
    }
    const fri = friday && air.mode === "remain";
    if (fromPaste && !fri) throw new Error("넣은 표에 Remain 열(또는 출고수량 열)이 없습니다. 월요일 주간 리스트라면 위에서 '📅 월요일 메일'을 고르고 그 칸에 넣으세요.");
    if (friday && !fri) log("[참고] 선박 표가 없고, 에어 표에도 Remain·출고수량 열이 없어 화요일 에어(월요일 리스트)로 만듭니다.", "warn");
    if (shipRes) { ({ rows: ship, etdSel, hasSum } = shipRes); if (shipFrom !== wbShip) log("[참고] 선박리스트를 따로 넣지 않아, 에어리스트 파일 안의 선박 표를 함께 사용합니다.", "warn"); }
    else if (!fri) log("선박리스트 없음 → 에어만 진행");

    // 기준 주: 금요일 에어(수요일에 받는 REMAIN 리스트)는 이번 주. 날짜 칸을 직접 바꾸지 않았으면 이번 주 월요일로 잡는다.
    let monday = $("monday").value ? parseISO($("monday").value) : nextMonday();
    // 날짜 칸을 직접 바꾸지 않았으면: 리스트의 NEED ETD 날짜로 주를 정한다(없으면 금요일 에어는 이번 주, 월요일 리스트는 날짜 칸 그대로)
    if (!$("monday").dataset.touched) {
      if (air.weekMonday) { monday = air.weekMonday; log(`[기준 주] 리스트의 NEED ETD 날짜로 ${fmtDate(monday)}(월) 주로 잡았습니다. 다른 주면 기준 주 월요일을 직접 바꾸고 다시 누르세요.`, "warn"); }
      else if (fri) { const t = new Date(); monday = addDays(new Date(t.getFullYear(), t.getMonth(), t.getDate()), -((t.getDay() + 6) % 7)); log(`[금요일 에어] 기준 주를 이번 주(${fmtDate(monday)} 월)로 잡았습니다. 다른 주면 기준 주 월요일을 바꾸고 다시 누르세요.`, "warn"); }
      $("monday").value = fmtDate(monday); if ($("monHint")) $("monHint").style.display = "none";
    }
    const wd = (monday.getDay() + 6) % 7;
    if (wd) { log(`[경고] ${fmtDate(monday)}는 월요일이 아닙니다. 그 주 월요일로 맞춥니다.`, "warn"); monday = addDays(monday, -wd); }
    const [month, week] = weekLabel(monday); RESULT.weekKey = fmtDate(monday); RESULT.fri = fri;
    const airDate = addDays(monday, fri ? 4 : 1), shipDate = addDays(monday, 2);
    const subjectBase = (S["제목형식"] || "대아홀딩스 {월}월 {주}째주 출하 요청의 건").replace("{월}", month).replace("{주}", week);
    const typed = splitList($("only").value), mine = $("mineOnly").checked ? myVendors() : null;
    const only = typed.length ? typed : (mine || []);   // "특정 업체만" 칸이 우선, 비어 있으면 "내 업체만"
    if (!typed.length && mine) log(`[내 업체만] ${mine.length}곳만 만듭니다. 전체를 만들려면 "내 업체만" 체크를 끄세요.`, "warn");
    const commonCc = splitList(S["공통참조"]);

    const shipAll = ship;
    if (ship.length && !includeZero) ship = ship.filter((r) => r._hasQty);
    const unShip = ship.filter((r) => !r._vendor && r._hasQty), unAir = air.filter((r) => !r._vendor);

    log(""); log("==================== 미리보기 ====================", "head");
    log(`기준 주: ${fmtDate(monday)} (월)   제목: ${subjectBase}`);
    if (fri) log(`[금요일 에어 · REMAIN 리스트] 에어 입고 요망: ${koDate(airDate)}   (선박·긴급 자재 메일은 만들지 않음)`, "head");
    else log(`에어 입고 요망: ${koDate(airDate)}   선박 입고 요망: ${koDate(shipDate)}`);
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
      const subject = subjectBase + (fri ? (S["금요에어제목접미"] ?? "(금요일 에어)") : a.length && !nShipQty ? (S["에어만제목접미"] ?? "(에어)") : "");
      if (!nShipQty) s = [];
      log(`■ ${n}  선박 ${nShipQty}건${s.length !== nShipQty ? `(전체 ${s.length}행)` : ""} / 에어 ${a.length}건  ${v.attach && s.length ? "[엑셀 첨부]" : "[본문 표]"}  제목: ${subject}`);
      log(`   받는 사람: ${v.to.join(", ")}`); if (v.cc.length) log(`   추가 참조: ${v.cc.join(", ")}`); if (v.memo) log(`   메모: ${v.memo}`, "warn");
      const attached = v.attach && s.length > 0;
      const attachments = attached ? [{ name: safeName(`${subjectBase}_${v.name}.xlsx`), data: vendorXlsx(s, a, etdSel) }] : [];
      // 입고 파일(전산에 올리는 엑셀): 이번 주(첫 ETD) 요청 수량이 있는 줄만, Fac → 1010/1060/1150, 업체코드 넣기
      let inbound = null;
      const inRows = s.filter((r) => toNum(r[etdSel[0]]) > 0).map((r) => ({ fac: r.Fac, cd: r["ITEM CD"], maker: r.MAKER, nm: r["ITEM NM"], req: toNum(r[etdSel[0]]), recv: null }));
      if (inRows.length) {
        const vcode = v.code || store.get("mailmaker_vendor_codes", {})[n] || "";
        inbound = { name: safeName(`${pad2(shipDate.getMonth() + 1)}${pad2(shipDate.getDate())} ${n}.xlsx`), data: inboundXlsx(etdSel[0], vcode, inRows), noCode: !vcode };
        RESULT.extraFiles.push({ name: "입고파일/" + inbound.name, data: inbound.data, inbound: true });
        log(`   입고 파일: ${inbound.name} (${inRows.length}줄)${vcode ? "" : "  [확인] 업체코드가 없어 빈칸입니다 → ⚙ 업체·설정 고치기에서 '전산 업체코드'를 넣어 주세요"}`, vcode ? "" : "warn");
      }
      RESULT.mails.push({ kind: "ship", vendor: n, subject, to: v.to, cc: uniq([...commonCc, ...v.cc]), body: buildShipBody(v, a, s, etdSel, airDate, shipDate, S, attached, fri), attachments, inbound, note: v.memo ? "메모: " + v.memo : "", ship: s, air: a, noTo, include: !noTo, keys: [`${fri ? "airfri" : "ship"}|${RESULT.weekKey}|${n}`] });
      RESULT.etdSel = etdSel; RESULT.tag = `${month}월${week}주`;
    }
    // ---- 긴급 자재 메일 (설정 "긴급메일업체"에 적힌 업체. 에어리스트에 올라온 그 업체 품목을 표로)
    for (const un of fri ? [] : splitList(S["긴급메일업체"])) {
      const v = CONFIG.vendors[matchAlias(un)]; if (!v || (only.length && !only.includes(v.name))) continue;
      const items = urgentItems(air.filter((r) => r._vendor === v.name), shipAll, addDays(monday, 7));
      if (!items.length) { log(`■ ${v.name} 긴급 자재: 에어리스트에 해당 업체 품목이 없어 만들지 않음`); continue; }
      const subject = (S["긴급제목형식"] || "대아홀딩스 {월}월 {주}째주 긴급 자재의 건").replace("{월}", month).replace("{주}", week).replace("{업체}", v.poShort);
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
/** 주문서 PDF의 거래처명(+업체코드)으로 업체 찾기. 돌려주는 값: {v, how} / 못 찾으면 null.
    1) 이름이 정확히 같음  2) PDF의 업체코드가 같음  3) PDF 이름이 등록된 이름(4글자 이상)으로 시작함(가장 긴 것, 한 업체로 정해질 때만).
    메일이 엉뚱한 업체로 가면 안 되므로 3)은 좁게 잡는다: 짧은 이름("엠이")이 다른 업체("에스엠이코리아")를 가로채지 않게 한다. */
function findVendorByPo(poVendorName, poCode) {
  const key = normName(poVendorName);
  const names = (v) => uniq([v.name, ...v.poNames, v.poShort].map(normName).filter(Boolean));
  if (key) for (const n of CONFIG.order) { const v = CONFIG.vendors[n]; if (names(v).includes(key)) return { v, how: "" }; }
  const code = clean(poCode);
  if (code) { const hit = CONFIG.order.filter((n) => clean(CONFIG.vendors[n].code) === code); if (hit.length === 1) return { v: CONFIG.vendors[hit[0]], how: `업체코드 ${code}` }; }
  let best = null, tie = false;
  if (key) for (const n of CONFIG.order) for (const a of names(CONFIG.vendors[n])) {
    if (a.length < 4 || !key.startsWith(a)) continue;
    if (!best || a.length > best.a.length) { best = { a, n }; tie = false; } else if (a.length === best.a.length && n !== best.n) tie = true;
  }
  return best && !tie ? { v: CONFIG.vendors[best.n], how: "이름 앞부분" } : null;
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
      const e = poPages.get(key); const found = findVendorByPo(e.vendor, e.code), v = found && found.v;
      const row = { po: e.po, pages: e.pages.length, pdfVendor: e.vendor, vendor: v ? v.name + (found.how ? ` (${found.how}로 찾음)` : "") : "", status: "" };
      RESULT.poTable.push(row);
      if (!v) { unknown.push(e); row.status = "업체 못 찾음"; continue; }
      if (!v.sendPo) { skipped.push(e); row.status = "발주서 미발송 업체(건너뜀)"; continue; }
      row.status = v.to.length ? "메일 생성" : "받는 사람 없음";
      (plan[v.name] = plan[v.name] || { vendor: v, entries: [], poVendor: e.vendor, how: found.how }).entries.push(e);
    }
    log("==================== 미리보기 ====================", "head");
    log(`공통 참조: ${commonCc.join(", ")}`); log("");
    for (const [n, p] of Object.entries(plan)) {
      const v = p.vendor, pos = p.entries.map((e) => e.po);
      const subject = (S["주문서제목형식"] || "대아 주문서-{PO}").replace("{PO}", pos.join(",")).replace("{업체}", v.poShort);
      const attachments = [];
      for (const e of p.entries) attachments.push({ name: safeName((S["주문서첨부이름"] || "{PO}.pdf").replace("{PO}", e.po).replace("{업체}", v.poShort)), data: await splitPdf(e.buf, e.pages) });
      log(`■ ${n} (PDF 거래처명: ${p.poVendor})  PO ${pos.join(", ")}  제목: ${subject}`);
      if (p.how) log(`   [확인] 이름이 정확히 같지 않아 ${p.how}(으)로 찾았습니다. 업체가 맞는지 확인하세요.`, "warn");
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
