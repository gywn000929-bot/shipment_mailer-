/* 📨 계장님께 보내는 내부 메일 (남효주 → 김민지 계장)
   주간 리스트 엑셀 하나로 세 가지를 만든다. 설정 "내부받는사람"이 있는 설정 파일에서만 보인다.
   ① 월요일 리스트: "{월}월 {주}째주 선박 및 에어리스트의 건" + 엑셀 첨부
   ② 월요일 오후 에어 업데이트: AIR 시트 전체 표(Purchase request·BALANCE·STOCK·B-R …)를 본문에
   ③ 수요일 REMAIN: Remain = Purchase request − 출고수량 > 0 인 품목 표를 본문에 + 엑셀 첨부 */
"use strict";
const INT_KINDS = {
  list: { label: "① 월요일: 선박 및 에어리스트 (엑셀 첨부)", re: false, attach: true, lines: ["수고 많으십니다.", "첨부 금주 선박 및 에어리스트 입니다.", "감사합니다."] },
  air: { label: "② 월요일 오후: 에어 업데이트 전체 리스트 (본문 표)", re: true, attach: false, lines: ["수고 많으십니다.", "하기 금주 F1&F2 에어 업데이트 된 전체 에어 리스트 입니다.", "", "{표}", "", "감사합니다."],
    cols: ["FAC", "Item cd", "Maker", "Part Name", "Purchase request", "BALANCE", "STOCK", "B-R", "Short Del.", "NEED ETD", "REMARK"], hi: "Purchase request" },
  remain: { label: "③ 수요일: 금요일 에어 REMAIN 리스트 (본문 표 + 엑셀 첨부)", re: true, attach: true, lines: ["수고 많으십니다.", "하기 금주 금요일 에어 REMAIN 리스트 입니다.", "", "{표}", "", "감사합니다."],
    cols: ["FAC", "Item cd", "Maker", "Part Name", "Remain", "STOCK", "Short Del.", "NEED ETD", "REMARK"], hi: "Remain" },
};
const intOn = () => !!(CONFIG && splitList(CONFIG.settings["내부받는사람"]).length);

// 엑셀 → 어떤 메일이 맞는지: 출고수량/Remain 열이 있으면 ③, 아니면 ①
function intGuessKind(wb) {
  for (const sn of wb.SheetNames) { const rows = sheetRows(wb, sn), h = findHeaderRow(rows, ["Item cd", "Maker"]); if (h >= 0 && rows[h].map(norm).some((c) => c === "출고수량" || c === "remain")) return "remain"; }
  return "list";
}
// AIR 시트 → 표 행. kind=air: 전체(원래 순서), kind=remain: Remain>0 만, FAC→MAKER→품명 순
function intAirRows(wb, kind) {
  const sh = loadSheet(wb, ["Item cd", "Maker"], "에어리스트", ["Remain", "Purchase request"], [], kind === "remain" ? AIR_REMAIN_COLS : [], "", kind === "remain" ? [] : AIR_REMAIN_COLS);
  const c = (names) => col(sh, names, { required: false });
  const cFac = c(["FAC", "Fac"]), cCd = c(["Item cd", "ITEM CD"]), cMk = c(["Maker", "MAKER"]), cNm = c(["Part Name", "ITEM NM"]), cPr = c(["Purchase request"]), cRem = c(["Remain"]), cOut = c(["출고수량"]);
  const cBal = c(["BALANCE"]), cSt = c(["STOCK"]), cBr = c(["B-R"]), cSh = c(["Short Del.", "Short Del"]), cNeed = c(["NEED ETD", "Need ETD"]), cRm = c(["REMARK", "Remark"]);
  const out = [];
  for (const r of sh.rows) {
    const cd = clean(r[cCd]); if (!cd) continue;
    const g = (k) => (k ? r[k] : "");
    const row = { FAC: clean(g(cFac)), "Item cd": cd, Maker: clean(g(cMk)) || makerFromName(g(cNm)), "Part Name": clean(g(cNm)), "Purchase request": g(cPr), BALANCE: g(cBal), STOCK: kind === "remain" ? "" : g(cSt), "B-R": g(cBr), "Short Del.": clean(g(cSh)), "NEED ETD": clean(g(cNeed)), REMARK: clean(g(cRm)) };
    if (kind === "remain") { row.Remain = cRem ? toNum(g(cRem)) : toNum(g(cPr)) - toNum(g(cOut)); if (row.Remain <= 0) continue; }
    out.push(row);
  }
  if (kind === "remain") out.sort(airRowSort);
  out.sheet = sh.sheet; out.hasOut = !!(cRem || cOut);
  return out;
}
const INT_NUM = ["Purchase request", "BALANCE", "STOCK", "B-R", "Remain"];
function intCell(v, k) {
  if (!INT_NUM.includes(k)) return clean(v);
  if (v == null || clean(v) === "") return "";
  const n = toNum(v); return Number.isInteger(n) ? n.toLocaleString("en-US") : n.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

async function runInternal() {
  if (!CONFIG) return;
  const f = $("intFile").files[0]; if (!f) { toast("주간 리스트 엑셀을 먼저 넣으세요."); return; }
  clearLog(); setStatus("실행 중...", true); RESULT = { mails: [], unmatched: [], kind: "internal", extraFiles: [] };
  try {
    const S = CONFIG.settings, buf = await f.arrayBuffer(), wb = XLSX.read(buf, { type: "array", cellDates: true });
    let kind = $("intKind").value; if (kind === "auto") { kind = intGuessKind(wb); log(`[자동 판단] ${kind === "remain" ? "출고수량(또는 Remain) 열이 있어 ③ 수요일 REMAIN" : "① 월요일 리스트"} 메일로 만듭니다. 다른 메일이면 위에서 고르세요.`, "warn"); }
    const K = INT_KINDS[kind];
    // 기준 주: 에어 표의 NEED ETD 날짜로(없으면 이번 주)
    let monday = null; try { monday = weekFromNeedEtd(intAirRows(wb, "air")); } catch (e) {}
    if (!monday) { const t = new Date(); monday = addDays(new Date(t.getFullYear(), t.getMonth(), t.getDate()), -((t.getDay() + 6) % 7)); }
    const [month, week] = weekLabel(monday);
    const subject = (K.re ? "Re: " : "") + (S["내부제목형식"] || "{월}월 {주}째주 선박 및 에어리스트의 건").replace("{월}", month).replace("{주}", week);
    let table = "";
    if (K.cols) {
      const rows = intAirRows(wb, kind);
      if (kind === "remain" && !rows.hasOut) throw new Error("이 엑셀의 AIR 시트에 출고수량(또는 Remain) 열이 없습니다. ③ REMAIN 메일은 출고수량을 적은 뒤에 만드세요.");
      log(`[${kind === "remain" ? "REMAIN" : "에어 업데이트"}] 시트 '${rows.sheet}' · ${rows.length}품목${kind === "remain" ? " (Remain 0 이하는 뺌, STOCK 칸은 비워 둠 — 필요하면 ✏ 고치기로 적으세요)" : ""}`);
      table = htmlTable(K.cols, rows.map((r) => K.cols.map((k) => intCell(r[k], k))), [K.hi]);
    }
    const body = templateHtml(K.lines.join("\n"), {}, table);
    const to = splitList(S["내부받는사람"]), cc = splitList(S["내부참조"]);
    RESULT.mails.push({ kind: "internal", vendor: "김민지 계장", subject, to, cc, body, attachments: K.attach ? [{ name: f.name, data: new Uint8Array(buf) }] : [], note: "", noTo: !to.length, include: true, keys: [`internal|${fmtDate(monday)}|${kind}`] });
    log(`■ ${subject} → ${to.join(", ")}${cc.length ? " (참조 " + cc.join(", ") + ")" : ""}${K.attach ? " · 첨부 " + f.name : ""}`);
    log("완료: 메일 1통. 아래에서 받으세요.", "ok"); setStatus("완료");
  } catch (e) { console.error(e); log("[오류] " + (e.message || e), "err"); setStatus("오류"); }
  renderResult(); $("resultCard").scrollIntoView({ behavior: "smooth" });
}
