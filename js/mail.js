/* 메일 본문 · 첨부 엑셀 · .eml 파일 · (Outlook에서 안 열릴 때) 복사해서 보내기 */
"use strict";
const RED_BG = "FFC7CE";
// ---------------------------------------------------------------- 본문 HTML
function htmlTable(headers, rows, highlight = [], redRows = []) {
  const th = headers.map((h) => `<th style="border:1px solid #999;padding:3px 8px;background:#f2f2f2;white-space:nowrap">${esc(h)}</th>`).join("");
  const trs = rows.map((row, ri) => "<tr>" + row.map((v, i) => {
    const h = headers[i];
    const st = ["border:1px solid #999", "padding:3px 8px"];
    if (redRows[ri]) st.push("background:#" + RED_BG);
    else if (highlight.includes(h)) st.push("background:#" + YELLOW);
    if (typeof v === "number" || /^[\d,]+$/.test(String(v))) st.push("text-align:right");
    return `<td style="${st.join(";")}">${esc(String(v ?? ""))}</td>`;
  }).join("") + "</tr>").join("");
  return `<table style="border-collapse:collapse;font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:10pt">${"<tr>" + th + "</tr>"}${trs}</table>`;
}
const P = (s) => `<p style="margin:0">${s}</p>`;
const BR = '<p style="margin:0">&nbsp;</p>';
function bodyWrap(settings, inner) {   // 설정: 본문크기(pt), 본문굵게(Y/N)
  const size = parseInt(settings["본문크기"]) || 10, bold = /^(Y|YES|O|1)$/i.test(clean(settings["본문굵게"]));
  return `<div style="font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:${size}pt${bold ? ";font-weight:bold" : ""}">${inner}</div>`;
}
// 서명: 설정 시트의 "서명". Outlook이 서명을 자동으로 넣는 PC에서는 화면의 "본문 끝에 서명 넣기"를 끈다(이 PC에 기억).
const sigOn = () => !pref("sigOff", CONFIG_NAME);
function sigHtml(settings) {
  const sig = settings["서명"] || "";
  return sig && sigOn() ? [BR, ...sig.split("\n").map((l) => P(esc(l)))].join("") : "";
}
// 메일 한 통의 최종 본문 (본문 + 서명). 미리보기·.eml·복사에서 모두 이것을 쓴다.
// m.customHtml: 결과 화면의 "✏ 고치기"로 이 메일만 고친 본문(서명 포함 전체)
const mailHtml = (m) => m.customHtml != null ? m.customHtml : bodyWrap(CONFIG.settings, m.body + sigHtml(CONFIG.settings));
// 줄글 양식(설정 시트 값) → 문단. {표} 자리에 표를 넣는다.
// blocks: 줄 전체가 {이름}이면 그 자리에 통째로 넣는 HTML 덩어리(표 등). 비어 있으면 그 줄은 빠진다.
function templateHtml(tpl, vars = {}, blocks = "") {
  if (typeof blocks === "string") blocks = { 표: blocks };
  let out = [];
  for (let line of String(tpl).replace(/\r/g, "").split("\n")) {
    const bk = line.trim().match(/^\{([^{}]+)\}$/);
    if (bk && bk[1] in blocks) { if (blocks[bk[1]]) out.push(blocks[bk[1]]); continue; }
    let skip = false;
    for (const [k, val] of Object.entries(vars)) if (line.includes(`{${k}}`)) { if (!val && line.trim() === `{${k}}`) skip = true; line = line.replace(`{${k}}`, val || ""); }
    if (skip) continue;
    out.push(line.trim() ? P(esc(line)) : BR);
  }
  out = out.filter((x, i) => !(x === BR && out[i - 1] === BR));
  return out.join("");
}

// 출하 요청 본문. 문구는 설정 "출하요청본문"(월요일) / "금요에어본문"(수요일 REMAIN)에서 고친다.
//  {인사} = 호칭 + 출하요청인사 + 발신자소개, {안내} = "첨부/하기 내용으로 금주 … 리스트 송부 드립니다."
//  {선박} = 선박리스트 제목 줄 + 표(또는 첨부 안내), {에어} = 에어리스트 제목 줄 + 표. 줄 전체가 {선박}/{에어}여야 한다.
const SHIP_TPL_DEFAULT = "{인사}\n{안내}\n\n{선박}\n\n{에어}\n\n감사합니다.";
const FRI_TPL_DEFAULT = "{인사}\n하기 내용으로 금주 금요일 에어 리스트 송부 드립니다.\n\n{에어}\n\n감사합니다.";
function buildShipBody(v, air, ship, etdSel, airDate, shipDate, settings, attached, friday = false) {
  const hasShip = ship.length > 0, hasAir = air.length > 0;
  const greet = (v.title ? v.title + " " : "") + (clean(settings["출하요청인사"]) || "안녕하세요.") + " " + (settings["발신자소개"] || "대아홀딩스 남효주 입니다.");
  const guide = hasShip ? (attached ? "첨부" : "하기") + " 내용으로 금주 선박 리스트 송부 드립니다." : `하기 내용으로 금주 ${friday ? "금요일 " : ""}에어 리스트 송부 드립니다.`;
  let shipBlk = "", airBlk = "";
  if (hasShip) {
    shipBlk = P(`선박리스트 (노란색 수량, <b>${shortDate(shipDate)}</b> 입고 요망)`) + (attached ? P("첨부 선박 리스트는 자재 입고 전 입고수량 기재하여 회신 부탁드립니다.")
      : htmlTable(["Fac", "ITEM CD", "MAKER", "ITEM NM", ...etdSel], ship.map((r) => [r.Fac, r["ITEM CD"], r.MAKER, r["ITEM NM"], ...etdSel.map((c) => fmtNum(r[c]))]), [etdSel[0]]));
  }
  if (hasAir) {
    const qtyLabel = air[0]._qtyCol || "Purchase request", hasShort = air.some((r) => clean(r["Short Del."]));
    const hdr = ["FAC", "Item cd", "Maker", "Part Name", qtyLabel, ...(hasShort ? ["Short Del."] : []), "NEED ETD"];
    airBlk = P(`에어리스트 (노란색 수량, <b>${shortDate(airDate)}</b> 입고 요망)`) + htmlTable(hdr, air.map((r) => [r.FAC, r["Item cd"], r.Maker, r["Part Name"], fmtNum(r["Purchase request"]), ...(hasShort ? [r["Short Del."]] : []), r["NEED ETD"]]), [qtyLabel]);
  }
  let tpl = clean(settings[friday ? "금요에어본문" : "출하요청본문"]) ? settings[friday ? "금요에어본문" : "출하요청본문"] : friday ? FRI_TPL_DEFAULT : SHIP_TPL_DEFAULT;
  // 고친 문구에서 {선박}/{에어} 줄을 지웠어도 표가 빠지지 않게 맨 끝에 붙인다
  for (const [k, blk] of [["선박", shipBlk], ["에어", airBlk]]) if (blk && !new RegExp("^\\s*\\{" + k + "\\}\\s*$", "m").test(tpl)) tpl += `\n\n{${k}}`;
  const memo = v.memo ? `<p style="color:red;font-weight:bold;background:#fff3cd;padding:6px;border:2px solid red">※ ${esc(v.memo)} (보내기 전 이 줄은 지워 주세요)</p>` : "";
  return memo + (v.recvLine ? P(esc(v.recvLine)) + BR : "") + templateHtml(tpl, { 인사: greet, 안내: guide, 호칭: v.title ? v.title + " " : "" }, { 선박: shipBlk, 에어: airBlk });
}
function buildPoBody(v, settings) {
  let tpl = settings["주문서본문"] || ("안녕하세요. " + (settings["발신자소개"] || "대아홀딩스 남효주 입니다.") + "\n수고 많으십니다.\n\n상기 제목에 따라 주문서 송부하오니\n첨부파일 확인 및 수신 여부 회신 부탁드리겠습니다.\n\n감사합니다.");
  if (!tpl.includes("{추가문구}") && v.poExtra) tpl = tpl.replace("감사합니다", "{추가문구}\n\n감사합니다");
  return (v.recvLine ? P(esc(v.recvLine)) + BR : "") + templateHtml(tpl, { 추가문구: v.poExtra, 호칭: v.title ? v.title + " " : "" });
}
// 긴급 자재 메일 (KET 주간). items: [{name, sum, air, short, red}]
const URGENT_HDR = ["NO", "품목명", "미납수량 합계", "AIR 수량", "부족일정"];
function buildUrgentBody(v, items, settings) {
  const tpl = settings["긴급본문"] || ("수고많으십니다. " + (settings["발신자소개"] || "") + "\n\n하기 긴급 자재 건 납기 확인 후 회신 부탁드립니다.\n\n{표}\n\n감사합니다.");
  const table = htmlTable(URGENT_HDR, items.map((it, i) => [i + 1, it.name, fmtNum(it.sum), fmtNum(it.air), it.short]), [], items.map((it) => it.red));
  return (v.recvLine ? P(esc(v.recvLine)) + BR : "") + templateHtml(tpl, { 호칭: v.title ? v.title + " " : "" }, table);
}

// ---------------------------------------------------------------- 첨부 엑셀 (xlsx-js-style)
function styleSheet(ws, nrows, ncols, hiCol) {
  const border = { top: { style: "thin", color: { rgb: "999999" } }, bottom: { style: "thin", color: { rgb: "999999" } }, left: { style: "thin", color: { rgb: "999999" } }, right: { style: "thin", color: { rgb: "999999" } } };
  const widths = [];
  for (let r = 0; r < nrows; r++) for (let c = 0; c < ncols; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    if (!ws[addr]) ws[addr] = { t: "s", v: "" };
    const cell = ws[addr];
    cell.s = { border };
    if (r === 0) cell.s = { border, font: { bold: true }, fill: { fgColor: { rgb: "F2F2F2" } }, alignment: { horizontal: "center" } };
    if (hiCol != null && c === hiCol) cell.s.fill = { fgColor: { rgb: YELLOW } };
    if (r > 0 && typeof cell.v === "number") cell.z = "#,##0";
    const len = String(cell.v ?? "").length;
    widths[c] = Math.max(widths[c] || 0, len);
  }
  ws["!cols"] = widths.map((w) => ({ wch: Math.min(Math.max(8, w * 1.3 + 2), 60) }));
}
function vendorXlsx(ship, air, etdSel) {
  const wb = XLSX.utils.book_new();
  if (ship.length) {
    const hdr = ["Fac", "ITEM CD", "MAKER", "ITEM NM", etdSel[0], "입고수량", ...etdSel.slice(1)];
    const aoa = [hdr, ...ship.map((r) => [r.Fac, r["ITEM CD"], r.MAKER, r["ITEM NM"], toNum(r[etdSel[0]]) || 0, null, ...etdSel.slice(1).map((c) => toNum(r[c]) || 0)])];
    const ws = XLSX.utils.aoa_to_sheet(aoa); styleSheet(ws, aoa.length, hdr.length, 4);
    XLSX.utils.book_append_sheet(wb, ws, "선박");
  }
  if (air.length) {
    const hdr = ["FAC", "Item cd", "Maker", "Part Name", "Purchase request", "Short Del.", "NEED ETD"];
    const aoa = [hdr, ...air.map((r) => [r.FAC, r["Item cd"], r.Maker, r["Part Name"], toNum(r["Purchase request"]), r["Short Del."], r["NEED ETD"]])];
    const ws = XLSX.utils.aoa_to_sheet(aoa); styleSheet(ws, aoa.length, hdr.length, 4);
    XLSX.utils.book_append_sheet(wb, ws, "에어");
  }
  return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}
function unmappedXlsx(ship, air, etdSel) {
  const hdr = ["구분", "Fac", "ITEM CD", "MAKER", "ITEM NM", "수량", "비고"];
  const aoa = [hdr,
    ...ship.map((r) => ["선박", r.Fac, r["ITEM CD"], r.MAKER, r["ITEM NM"], etdSel.filter((c) => toNum(r[c]) > 0).map((c) => `${c}: ${fmtNum(r[c])}`).join(" / "), "config.xlsx 업체 시트의 MAKER 또는 품목지정 시트에 추가하세요"]),
    ...air.map((r) => ["에어", r.FAC, r["Item cd"], r.Maker, r["Part Name"], fmtNum(r["Purchase request"]), ""])];
  const ws = XLSX.utils.aoa_to_sheet(aoa); styleSheet(ws, aoa.length, hdr.length, null);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "미지정 품목");
  return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}

// ---------------------------------------------------------------- .eml 만들기
function b64(bytes) {
  let s = ""; const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  return btoa(s);
}
const b64str = (s) => b64(new TextEncoder().encode(s));
const wrap76 = (s) => s.replace(/(.{76})/g, "$1\r\n");
const encWord = (s) => /^[\x20-\x7e]*$/.test(s) ? s : `=?UTF-8?B?${b64str(s)}?=`;
const MIME = { xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", pdf: "application/pdf" };
// X-Unsent: 1 → 클래식 Outlook에서 더블클릭하면 "보내기"가 있는 새 메일 창으로 열린다.
function buildEml(m) {
  const { to, cc, subject, attachments } = m, html = mailHtml(m);
  const bd = "----=_MailMaker_" + Math.random().toString(36).slice(2);
  const L = [];
  L.push("X-Unsent: 1");
  L.push("MIME-Version: 1.0");
  L.push("To: " + to.join(", "));
  if (cc.length) L.push("Cc: " + cc.join(", "));
  L.push("Subject: " + encWord(subject));
  L.push("Date: " + new Date().toUTCString());
  L.push(`Content-Type: multipart/mixed; boundary="${bd}"`, "");
  L.push(`--${bd}`, 'Content-Type: text/html; charset="utf-8"', "Content-Transfer-Encoding: base64", "");
  L.push(wrap76(b64str(`<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body>${html}</body></html>`)), "");
  for (const a of attachments) {
    const ext = a.name.split(".").pop().toLowerCase();
    const fnStar = "UTF-8''" + encodeURIComponent(a.name);
    L.push(`--${bd}`, `Content-Type: ${MIME[ext] || "application/octet-stream"}; name="${encWord(a.name)}"`,
      `Content-Disposition: attachment; filename="${encWord(a.name)}"; filename*=${fnStar}`,
      "Content-Transfer-Encoding: base64", "", wrap76(b64(a.data)), "");
  }
  L.push(`--${bd}--`, "");
  return new Blob([L.join("\r\n")], { type: "message/rfc822" });
}
const emlName = (m) => safeName(m.subject + "_" + m.vendor) + ".eml";

// ---------------------------------------------------------------- .eml이 안 열리는 PC용: 새 메일 창 + 본문 복사
// 받는 사람·참조·제목이 채워진 새 메일 창을 연다(PC의 기본 메일 프로그램). 본문·첨부는 사람이 붙인다.
function mailtoUrl(m) {
  const q = [];
  if (m.cc.length) q.push("cc=" + encodeURIComponent(m.cc.join(",")));
  q.push("subject=" + encodeURIComponent(m.subject));
  return "mailto:" + m.to.map(encodeURIComponent).join(",") + "?" + q.join("&");
}
async function copyMailBody(m) {
  const html = mailHtml(m);
  try {
    const tmp = document.createElement("div"); tmp.innerHTML = html;
    await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([tmp.innerText || tmp.textContent], { type: "text/plain" }) })]);
    return true;
  } catch (e) {   // 권한이 없는 브라우저: 화면에 잠깐 그려서 복사
    const box = document.createElement("div"); box.innerHTML = html; box.style.cssText = "position:fixed;left:-9999px;top:0;background:#fff";
    document.body.appendChild(box);
    const sel = getSelection(), rg = document.createRange(); rg.selectNodeContents(box); sel.removeAllRanges(); sel.addRange(rg);
    let ok = false; try { ok = document.execCommand("copy"); } catch {}
    sel.removeAllRanges(); box.remove(); return ok;
  }
}
// 처음 한 번 확인용 시험 메일 (받는 사람 = 본인)
function testMail() {
  const S = CONFIG.settings, me = parseEmails((S["서명"] || "").replace(/\n/g, " "))[0] || "";
  const hdr = ["Fac", "ITEM CD", "ITEM NM", "수량"], rows = [["F1", "TEST-001", "시험 품목 가", 1000], ["F2", "TEST-002", "시험 품목 나", 250]];
  const ws = XLSX.utils.aoa_to_sheet([hdr, ...rows]); styleSheet(ws, 3, 4, 3);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "시험");
  const body = [P("이 메일은 메일 작성기가 내 PC에서 잘 되는지 보는 시험용입니다. 보내지 않고 닫아도 됩니다."), BR,
    P("확인할 것: 1. 위에 '보내기' 버튼이 있는지 2. 받는 사람·제목이 채워져 있는지 3. 아래 표의 노란 칸이 보이는지 4. 첨부 엑셀이 붙어 있는지 5. 맨 아래 서명이 한 번만 나오는지"), BR,
    htmlTable(hdr, rows.map((r) => [r[0], r[1], r[2], fmtNum(r[3])]), ["수량"]), BR, P("감사합니다.")].join("");
  return { vendor: "시험", subject: "[시험] 메일 작성기 확인용", to: me ? [me] : [], cc: [], body, attachments: [{ name: "시험_첨부.xlsx", data: new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" })) }] };
}
