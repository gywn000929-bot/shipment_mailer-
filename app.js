/* 대아홀딩스 메일 작성기 - 브라우저 전용 버전 (서버/Python 없음)
   엑셀·PDF 읽기, 본문 생성, 첨부 생성, PDF 분리를 모두 브라우저에서 처리하고
   Outlook에서 더블클릭하면 바로 '보내기'만 누르면 되는 .eml 파일을 내려준다. */
"use strict";
const $ = (id) => document.getElementById(id);
const YELLOW = "FFFF00";
const WEEKDAY_KO = "월화수목금토일";
const LS_KEY = "mailmaker_overlay_v1";

// ---------------------------------------------------------------- 유틸
const norm = (s) => (s == null ? "" : String(s).replace(/\s+/g, "").toLowerCase());
const clean = (v) => {
  if (v == null) return "";
  if (v instanceof Date) return isNaN(v) ? "" : fmtCellDate(v);
  const s = String(v).trim();
  return s.toLowerCase() === "nan" ? "" : s;
};
const toNum = (v) => {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "" || s === "-") return 0;
  const n = Number(s);
  return isNaN(n) ? 0 : n;
};
const fmtNum = (v) => {
  const n = toNum(v);
  if (n === 0) return "";
  return Number.isInteger(n) ? n.toLocaleString("en-US") : n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const splitList = (s) => clean(s).split(/[,;/\n]/).map((x) => x.trim()).filter(Boolean);
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const makerFromName = (name) => { const m = clean(name).match(/\(([^()]+)\)\s*$/); return m ? m[1].trim() : ""; };
const pad2 = (n) => String(n).padStart(2, "0");
const fmtDate = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const fmtCellDate = (d) => `${pad2(d.getDate())}-${MON[d.getMonth()]}-${String(d.getFullYear()).slice(2)}`;   // 엑셀에 보이는 형태
const parseISO = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const weekdayKo = (d) => WEEKDAY_KO[(d.getDay() + 6) % 7];
const koDate = (d) => `${d.getMonth() + 1}월 ${d.getDate()}일(${weekdayKo(d)})`;
const shortDate = (d) => `${d.getMonth() + 1}/${d.getDate()}일(${weekdayKo(d)})`;
const safeName = (s) => s.replace(/[\\/:*?"<>|]/g, "_");
const normName = (s) => clean(s).replace(/\(주\)|주식회사|㈜|\(\s*\d+\s*\)|\[[^\]]*\]/g, "").replace(/[\s()\-_.]/g, "").toLowerCase();
function matchAlias(name) {   // 리스트의 업체 칸 값 → 설정 업체명. 1) 정확히 일치 2) 별칭 포함(긴 별칭 우선)
  const key = normName(name); if (!key) return name;
  if (key in CONFIG.aliasMap) return CONFIG.aliasMap[key];
  let best = null;
  for (const [a, v] of Object.entries(CONFIG.aliasMap)) if (a.length >= 2 && (key.includes(a) || a.includes(key)) && (!best || a.length > best[0].length)) best = [a, v];
  return best ? best[1] : name;
}

function nextMonday(today = new Date()) {
  const wd = (today.getDay() + 6) % 7; // 월=0
  return addDays(today, (7 - wd) % 7 || 7);
}
function weekLabel(monday) {
  const thu = addDays(monday, 3);
  return [thu.getMonth() + 1, Math.ceil(thu.getDate() / 7)];
}

// ---------------------------------------------------------------- 로그
const LOG = [];
function log(t, k = "info") {
  const d = document.createElement("div"); d.className = k; d.textContent = t;
  $("log").appendChild(d); $("log").scrollTop = $("log").scrollHeight;
}
function clearLog() { $("log").innerHTML = ""; }
function toast(m) { const t = $("toast"); t.textContent = m; t.classList.add("show"); setTimeout(() => t.classList.remove("show"), 3000); }
function setStatus(t, run) { $("status").textContent = t; $("status").className = "status" + (run ? " run" : ""); }

// ---------------------------------------------------------------- 설정 (config.xlsx)
let CONFIG = null;      // {vendors:{}, makerMap, itemMap, settings, aliasMap, name, raw(workbook)}
let CONFIG_NAME = "config.xlsx";

function loadOverlay() { try { return JSON.parse(localStorage.getItem(LS_KEY) || "{}"); } catch { return {}; } }
function saveOverlay(o) { try { localStorage.setItem(LS_KEY, JSON.stringify(o)); } catch {} }

function sheetRows(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}
function sheetObjs(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
}

function parseConfig(wb, name) {
  const vendorsRaw = sheetObjs(wb, "업체");
  const ovRaw = sheetObjs(wb, "품목지정");
  const setRows = sheetRows(wb, "설정");
  const settings = {};
  for (const r of setRows) { const k = clean(r[0]); if (k) settings[k] = clean(r[1]).replace(/\\n/g, "\n"); }

  // 브라우저에 저장된 추가 지정(미지정 업체 지정하기에서 저장한 것) 합치기
  const overlay = (loadOverlay()[name] || {});
  const vendors = {}, makerMap = {}, itemMap = {}, aliasMap = {};
  const order = [];
  const addVendor = (r) => {
    const vname = clean(r["업체"]);
    if (!vname) return;
    if (["N", "NO", "X", "0"].includes(clean(r["활성"] ?? "Y").toUpperCase())) return;
    const ov = overlay.vendors?.[vname] || {};
    const v = {
      name: vname,
      to: splitList(r["받는사람"]).concat(ov.to || []),
      cc: splitList(r["추가참조"]),
      recvLine: clean(r["수신줄"]),
      title: clean(r["호칭"]),
      makers: splitList(r["MAKER"]).concat(ov.makers || []).map((m) => m.toUpperCase()),
      attach: !["N", "NO", "X", "0"].includes(clean(r["엑셀첨부"] ?? "Y").toUpperCase()),
      memo: clean(r["메모"]),
      poNames: splitList(r["거래처명(주문서)"]).concat(ov.poNames || []),
      sendPo: !["N", "NO", "X", "0"].includes(clean(r["발주서발송"] ?? "Y").toUpperCase()),
      poExtra: clean(r["주문서추가문구"]),
      poShort: clean(r["주문서약칭"]) || vname,
      _row: r,
    };
    v.to = [...new Set(v.to)];
    vendors[vname] = v; order.push(vname);
    for (const m of v.makers) if (!(m in makerMap)) makerMap[m] = vname;
    for (const a of [vname, ...v.poNames, v.poShort]) if (normName(a) && !(normName(a) in aliasMap)) aliasMap[normName(a)] = vname;
  };
  vendorsRaw.forEach(addVendor);
  for (const nv of overlay.newVendors || []) if (!vendors[nv.name]) addVendor({ 업체: nv.name, 받는사람: (nv.to || []).join(","), MAKER: (nv.makers || []).join(","), "거래처명(주문서)": (nv.poNames || []).join(",") });
  for (const r of ovRaw) { const c = clean(r["ITEM CD"]), v = clean(r["업체"]); if (c && v) itemMap[norm(c)] = v; }
  return { vendors, order, makerMap, itemMap, settings, aliasMap, name, wb };
}

async function loadConfigFromUrl(name) {
  const res = await fetch(name, { cache: "no-store" });
  if (!res.ok) throw new Error(`${name} 을(를) 불러오지 못했습니다 (${res.status})`);
  const wb = XLSX.read(await res.arrayBuffer(), { type: "array", cellDates: true });
  return parseConfig(wb, name);
}
async function loadConfigFromFile(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  return parseConfig(wb, file.name);
}
async function reloadConfig() {
  try {
    const sel = $("config").value;
    if (sel === "__file__") {
      const f = $("configFile").files[0];
      if (!f) { toast("설정 엑셀 파일을 골라 주세요"); return; }
      CONFIG = await loadConfigFromFile(f);
    } else {
      CONFIG = await loadConfigFromUrl(sel);
    }
    CONFIG_NAME = CONFIG.name;
    try { if (sel !== "__file__") localStorage.setItem("mailmaker_cfg", sel); } catch {}
    const who = (CONFIG.settings["발신자소개"] || "").replace(/^대아홀딩스\s*/, "").replace(/\s*(입니다|이다)\.?$/, "").trim();
    $("whoBadge").textContent = "보내는 사람: " + (who || CONFIG_NAME); $("whoBadge").className = "badge " + (/김민지/.test(who) ? "kim" : "nam");
    document.title = `메일 작성기 - ${who || CONFIG_NAME}`;
    const ov = loadOverlay()[CONFIG_NAME];
    const n = ov ? ((ov.newVendors || []).length + Object.keys(ov.vendors || {}).length) : 0;
    $("cfgInfo").textContent = `업체 ${CONFIG.order.length}개 · 발신: ${CONFIG.settings["발신자소개"] || ""}` + (n ? ` · 이 브라우저에 추가 지정 ${n}건` : "");
    $("btnDlCfg").style.display = n ? "" : "none";
  } catch (e) { toast(String(e.message || e)); $("cfgInfo").textContent = "설정을 못 읽음: " + (e.message || e); }
}

function resolveVendor(rowVendor, itemCd, maker) {
  if (rowVendor) return matchAlias(rowVendor);
  if (norm(itemCd) in CONFIG.itemMap) return CONFIG.itemMap[norm(itemCd)];
  return CONFIG.makerMap[clean(maker).toUpperCase()] || "";
}

// ---------------------------------------------------------------- 엑셀 표 읽기
function findHeaderRow(rows, mustHave, anyOf = [], prefixAny = []) {
  const need = mustHave.map(norm), any = anyOf.map(norm), pre = prefixAny.map(norm);
  for (let i = 0; i < Math.min(40, rows.length); i++) {
    const cells = (rows[i] || []).map(norm);
    if (!need.every((n) => cells.includes(n))) continue;
    if (any.length && !any.some((a) => cells.includes(a))) continue;
    if (pre.length && !cells.some((c) => pre.some((p) => c.startsWith(p)))) continue;
    return i;
  }
  return -1;
}
function loadSheet(wb, mustHave, label, anyOf = [], prefixAny = [], prefer = [], sheetName = "") {
  const found = [];
  for (const sn of wb.SheetNames) {
    if (sheetName && norm(sn) !== norm(sheetName)) continue;
    const rows = sheetRows(wb, sn);
    const h = findHeaderRow(rows, mustHave, anyOf, prefixAny);
    if (h < 0) continue;
    const headers = rows[h].map(clean);
    const data = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!r.some((c) => c != null && c !== "")) continue;
      const o = {};
      headers.forEach((hd, j) => { if (hd) o[hd] = r[j]; });
      data.push(o);
    }
    const hn = headers.map(norm);
    found.push({ sheet: sn, headers: headers.filter(Boolean), rows: data, score: prefer.filter((p) => hn.includes(norm(p))).length });
  }
  if (!found.length) throw new Error(`${label} 파일에서 ${label} 표(헤더 ${mustHave.join(", ")}${anyOf.length ? " + " + anyOf.join("/") : ""}${prefixAny.length ? " + ETD열" : ""})가 있는 시트를 못 찾았습니다. 시트: ${wb.SheetNames.join(", ")}${sheetName ? " / 지정한 시트: " + sheetName : ""}. 선박리스트와 에어리스트에 같은 파일을 고른 것은 아닌지 확인하세요.`);
  found.sort((a, b) => b.score - a.score);
  if (found.length > 1) log(`[참고] ${label} 표가 있는 시트가 ${found.length}개: ${found.map((f) => f.sheet).join(", ")} → '${found[0].sheet}' 사용${prefer.length ? " (Remain 열이 있는 시트 우선)" : ""}. 다른 시트를 쓰려면 파일 아래 시트 선택에서 고르세요.`, "warn");
  return found[0];
}
// 파일을 고르면 시트 목록을 보여 주고 고를 수 있게
async function fillSheetSelect(input, selId, mustHave, anyOf, prefixAny, prefer) {
  const sel = $(selId); sel.innerHTML = '<option value="">(자동 선택)</option>'; sel.style.display = "none"; sel.parentElement.style.display = "none";
  const f = input.files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    const cands = [];
    for (const sn of wb.SheetNames) { const rows = sheetRows(wb, sn); const h = findHeaderRow(rows, mustHave, anyOf, prefixAny); if (h >= 0) { const hn = rows[h].map(norm); cands.push({ sn, score: prefer.filter((p) => hn.includes(norm(p))).length }); } }
    if (cands.length > 1) {
      cands.sort((a, b) => b.score - a.score);
      sel.innerHTML = `<option value="">(자동: ${cands[0].sn})</option>` + wb.SheetNames.map((sn) => `<option value="${esc(sn)}">${esc(sn)}${cands.some((c) => c.sn === sn) ? "" : " (표 없음)"}</option>`).join("");
      sel.style.display = ""; sel.parentElement.style.display = "";
    }
  } catch (e) { console.warn(e); }
}


function col(sheet, cands, { required = true, prefix = false } = {}) {
  const lookup = {}; sheet.headers.forEach((c) => (lookup[norm(c)] = c));
  for (const c of cands) if (norm(c) in lookup) return lookup[norm(c)];
  if (prefix) for (const c of cands) for (const [k, v] of Object.entries(lookup)) if (k.startsWith(norm(c))) return v;
  if (required) throw new Error(`열을 찾을 수 없습니다: ${cands.join("/")}  (실제 열: ${sheet.headers.join(", ")})`);
  return null;
}
const VENDOR_COLS = ["업체", "거래처명(매칭)", "거래처명", "거래처", "VENDER", "VENDOR"];

function parseShip(wb, etdCount) {
  const sh = loadSheet(wb, ["ITEM CD", "MAKER"], "선박리스트", [], ["ETD"], [], $("shipSheet").value);
  const cFac = col(sh, ["Fac", "FAC"]), cCd = col(sh, ["ITEM CD"]), cMk = col(sh, ["MAKER"]), cNm = col(sh, ["ITEM NM", "Part Name", "품명"]);
  const cV = col(sh, VENDOR_COLS, { required: false, prefix: true });
  log(`선박리스트 시트: ${sh.sheet}` + (cV ? ` / 업체 열: '${cV}' 사용` : " / 업체 열 없음 → MAKER로 매칭"));
  const etdAll = sh.headers.filter((c) => norm(c).startsWith("etd"));
  if (!etdAll.length) throw new Error("선박리스트에 ETD 열이 없습니다.");
  const etdSel = etdAll.slice(0, etdCount);
  log(`선박리스트 ETD 열(전체): ${etdAll.join(", ")}`);
  log(`선박리스트 ETD 열(선택): ${etdSel.join(", ")}  ← 첫 열 '${etdSel[0]}'이 이번 주(노란색)`);
  const rows = [];
  for (const r of sh.rows) {
    const cd = clean(r[cCd]); if (!cd) continue;
    const qtys = etdSel.map((c) => toNum(r[c]));
    const maker = clean(r[cMk]) || makerFromName(r[cNm]);
    const row = { Fac: clean(r[cFac]), "ITEM CD": cd, MAKER: maker, "ITEM NM": clean(r[cNm]), _hasQty: qtys.some((q) => q > 0), _src: "선박" };
    etdSel.forEach((c) => (row[c] = r[c]));
    row._vendor = resolveVendor(cV ? clean(r[cV]) : "", cd, maker);
    rows.push(row);
  }
  return { rows, etdSel, etdAll };
}
function parseAir(wb) {
  const sh = loadSheet(wb, ["Item cd", "Maker"], "에어리스트", ["Remain", "Purchase request"], [], ["Remain"], $("airSheet").value);
  const cFac = col(sh, ["FAC", "Fac"]), cCd = col(sh, ["Item cd", "ITEM CD"]), cMk = col(sh, ["Maker", "MAKER"]), cNm = col(sh, ["Part Name", "ITEM NM", "품명"]);
  const cRem = col(sh, ["Remain", "Purchase request"]), cShort = col(sh, ["Short Del.", "Short Del"], { required: false }), cNeed = col(sh, ["NEED ETD", "Need ETD"], { required: false });
  const cV = col(sh, VENDOR_COLS, { required: false, prefix: true });
  log(`에어리스트 시트: ${sh.sheet} / 수량 열: '${cRem}'` + (cV ? ` / 업체 열: '${cV}' 사용` : ""));
  const rows = [];
  for (const r of sh.rows) {
    const cd = clean(r[cCd]); if (!cd || toNum(r[cRem]) <= 0) continue;
    const maker = clean(r[cMk]) || makerFromName(r[cNm]);
    rows.push({
      FAC: clean(r[cFac]), "Item cd": cd, Maker: maker, "Part Name": clean(r[cNm]), "Purchase request": r[cRem],
      "Short Del.": cShort ? clean(r[cShort]) : "", "NEED ETD": cNeed ? clean(r[cNeed]) : "", _qtyCol: cRem,
      _vendor: resolveVendor(cV ? clean(r[cV]) : "", cd, maker), _src: "에어",
    });
  }
  return rows;
}

// ---------------------------------------------------------------- 본문 HTML
function htmlTable(headers, rows, highlight = []) {
  const th = headers.map((h) => `<th style="border:1px solid #999;padding:3px 8px;background:#f2f2f2;white-space:nowrap">${esc(h)}</th>`).join("");
  const trs = rows.map((row) => "<tr>" + row.map((v, i) => {
    const h = headers[i];
    const st = ["border:1px solid #999", "padding:3px 8px"];
    if (highlight.includes(h)) st.push("background:#" + YELLOW);
    if (typeof v === "number" || /^[\d,]+$/.test(String(v))) st.push("text-align:right");
    return `<td style="${st.join(";")}">${esc(String(v ?? ""))}</td>`;
  }).join("") + "</tr>").join("");
  return `<table style="border-collapse:collapse;font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:10pt"><tr>${th}</tr>${trs}</table>`;
}
const P = (s) => `<p style="margin:0">${s}</p>`;
function bodyWrap(settings, inner) {   // 설정: 본문크기(pt), 본문굵게(Y/N)
  const size = parseInt(settings["본문크기"]) || 10, bold = /^(Y|YES|O|1)$/i.test(clean(settings["본문굵게"]));
  return `<div style="font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:${size}pt${bold ? ";font-weight:bold" : ""}">${inner}</div>`;
}
const BR = '<p style="margin:0">&nbsp;</p>';
function sigHtml(settings) {
  const sig = settings["서명"] || "";
  return sig ? [BR, ...sig.split("\n").map((l) => P(esc(l)))].join("") : "";
}
function buildShipBody(v, air, ship, etdSel, airDate, shipDate, settings, attached) {
  const out = [];
  const hasShip = ship.length > 0, hasAir = air.length > 0;
  if (v.memo) out.push(`<p style="color:red;font-weight:bold;background:#fff3cd;padding:6px;border:2px solid red">※ ${esc(v.memo)} (보내기 전 이 줄은 지워 주세요)</p>`);
  if (v.recvLine) out.push(P(esc(v.recvLine)), BR);
  out.push(P(esc((v.title ? v.title + " " : "") + "안녕하세요. " + (settings["발신자소개"] || "대아홀딩스 남효주 입니다."))));
  out.push(P(hasShip ? (attached ? "첨부" : "하기") + " 내용으로 금주 선박 리스트 송부 드립니다." : "하기 내용으로 금주 에어 리스트 송부 드립니다."));
  if (hasShip) {
    out.push(P(`선박리스트 (노란색 수량, <b>${shortDate(shipDate)}</b> 입고 요망)`));
    if (attached) out.push(P("첨부 선박 리스트는 자재 입고 전 입고수량 기재하여 회신 부탁드립니다."));
    else out.push(htmlTable(["Fac", "ITEM CD", "MAKER", "ITEM NM", ...etdSel], ship.map((r) => [r.Fac, r["ITEM CD"], r.MAKER, r["ITEM NM"], ...etdSel.map((c) => fmtNum(r[c]))]), [etdSel[0]]));
    out.push(BR);
  }
  if (hasAir) {
    out.push(P(`에어리스트 (노란색 수량, <b>${shortDate(airDate)}</b> 입고 요망)`));
    const qtyLabel = air[0]._qtyCol || "Purchase request";
    const hdr = ["FAC", "Item cd", "Maker", "Part Name", qtyLabel];
    const hasShort = air.some((r) => clean(r["Short Del."]));
    if (hasShort) hdr.push("Short Del.");
    hdr.push("NEED ETD");
    out.push(htmlTable(hdr, air.map((r) => { const row = [r.FAC, r["Item cd"], r.Maker, r["Part Name"], fmtNum(r["Purchase request"])]; if (hasShort) row.push(r["Short Del."]); row.push(r["NEED ETD"]); return row; }), [qtyLabel]), BR);
  }
  out.push(P("감사합니다."), sigHtml(settings));
  return bodyWrap(settings, out.join(""));
}
function buildPoBody(v, settings) {
  let tpl = settings["주문서본문"] || ("안녕하세요. " + (settings["발신자소개"] || "대아홀딩스 남효주 입니다.") + "\n수고 많으십니다.\n\n상기 제목에 따라 주문서 송부하오니\n첨부파일 확인 및 수신 여부 회신 부탁드리겠습니다.\n\n감사합니다.");
  if (!tpl.includes("{추가문구}") && v.poExtra) tpl = tpl.replace("감사합니다", "{추가문구}\n\n감사합니다");
  let out = [];
  if (v.recvLine) out.push(P(esc(v.recvLine)), BR);
  for (let line of tpl.replace(/\r/g, "").split("\n")) {
    if (line.includes("{추가문구}")) { if (!v.poExtra) continue; line = line.replace("{추가문구}", v.poExtra); }
    if (line.includes("{호칭}")) line = line.replace("{호칭}", v.title ? v.title + " " : "");
    out.push(line.trim() ? P(esc(line)) : BR);
  }
  out = out.filter((x, i) => !(x === BR && out[i - 1] === BR));
  out.push(sigHtml(settings));
  return bodyWrap(settings, out.join(""));
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
function buildEml({ to, cc, subject, html, attachments }) {
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
function download(blob, name) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}

// ---------------------------------------------------------------- 결과 표시
let RESULT = { mails: [], unmatched: [], kind: "" };   // mails: [{vendor, subject, to, cc, html, attachments:[{name,data}]}]
function sentLog() { try { return JSON.parse(localStorage.getItem("mailmaker_sent") || "{}"); } catch { return {}; } }
function markSent(subjects) { const l = sentLog(); const today = fmtDate(new Date()); for (const s of subjects) l[s] = today; const keys = Object.keys(l); if (keys.length > 400) keys.slice(0, keys.length - 400).forEach((k) => delete l[k]); try { localStorage.setItem("mailmaker_sent", JSON.stringify(l)); } catch {} }
function renderResult() {
  const box = $("mails"); box.innerHTML = "";
  const ready = RESULT.mails.filter((m) => !m.noTo), noTo = RESULT.mails.filter((m) => m.noTo);
  $("resultCard").style.display = RESULT.mails.length || RESULT.unmatched.length ? "" : "none";
  $("btnZip").style.display = ready.length ? "" : "none";
  $("btnXlsx").style.display = "none";   // 전체 엑셀 받기는 에어 컨펌 도우미에서만 사용
  $("resultTitle").textContent = ready.length ? `만들어진 메일 ${ready.length}통` : "만들어진 메일 없음";
  // 요약 배너
  const sl = sentLog(); const dup = ready.filter((m) => sl[m.subject + "|" + m.vendor]);
  const parts = [];
  if (ready.length) parts.push(`<span class="ok">✅ 메일 ${ready.length}통 준비됨</span> → 오른쪽 위 <b>📦 전부 받기</b>를 누르고, 받은 .eml 파일을 더블클릭해 Outlook에서 <b>보내기</b>`);
  if (noTo.length) parts.push(`<span class="warn">⚠ 받는 사람이 없는 업체 ${noTo.length}곳</span> (${esc(noTo.map((m) => m.vendor).join(", "))}) → 아래 해당 칸에 메일 주소를 적고 <b>저장</b>하면 다음부터 기억합니다`);
  if (RESULT.unmatched.length) parts.push(`<span class="warn">⚠ 업체를 모르는 품목 ${RESULT.unmatched.reduce((a, b) => a + b.count, 0)}건</span> → 아래 <b>③ 미지정 업체 지정하기</b>에서 업체를 고른 뒤 다시 ①`);
  if (RESULT.skipped && RESULT.skipped.length) parts.push(`<span class="hint">건너뜀(발주서 미발송 업체): ${esc(RESULT.skipped.join(", "))}</span>`);
  if (dup.length) parts.push(`<span class="warn">⚠ 같은 제목의 메일 파일을 이미 받은 적이 있습니다 (${esc(dup.map((m) => m.vendor + " " + sl[m.subject + "|" + m.vendor]).join(", "))}) → 중복 발송이 아닌지 보낸 편지함을 확인하세요</span>`);
  if (RESULT.poTable && RESULT.poTable.length) {
    const cls = (st) => st === "메일 생성" ? "ok" : st.includes("건너뜀") ? "hint" : "warn";
    parts.push(`<details><summary>PDF에서 읽은 주문서 ${RESULT.poTable.length}건 보기</summary><table><tr><th>PO</th><th>쪽수</th><th>PDF 거래처명</th><th>매칭 업체</th><th>처리</th></tr>${RESULT.poTable.map((r) => `<tr><td>${esc(r.po)}</td><td>${r.pages}</td><td>${esc(r.pdfVendor)}</td><td>${esc(r.vendor || "-")}</td><td class="${cls(r.status)}">${esc(r.status)}</td></tr>`).join("")}</table></details>`);
  }
  $("summary").innerHTML = parts.map((p) => `<div>${p}</div>`).join("");
  RESULT.mails.forEach((m, i) => {
    const d = document.createElement("div"); d.className = "mail" + (m.noTo ? " noto" : "");
    d.innerHTML = `<div class="mh"><label class="inc"><input type="checkbox" data-i="${i}" data-act="inc" ${m.include ? "checked" : ""} ${m.noTo ? "disabled" : ""}> 포함</label><b>${esc(m.vendor)}</b><span class="subj">${esc(m.subject)}</span>
      ${m.noTo ? "" : `<button class="sm pri" data-i="${i}" data-act="eml">📩 메일 파일 받기</button>`}<button class="sm" data-i="${i}" data-act="view">본문 보기</button></div>
      <div class="meta">${m.noTo ? `<span class="warn">받는 사람 없음</span> → <input type="text" class="toin" data-i="${i}" placeholder="메일 주소 (여러 개면 쉼표)" style="width:300px"> <button class="sm pri" data-i="${i}" data-act="saveto">저장</button>` : "받는 사람: " + esc(m.to.join(", "))}${m.cc.length ? "<br>참조: " + esc(m.cc.join(", ")) : ""}${m.attachments.length ? "<br>첨부: " + esc(m.attachments.map((a) => a.name).join(", ")) : ""}${m.note && !m.noTo ? `<br><span class="warn">${esc(m.note)}</span>` : ""}</div>
      <iframe class="prev" style="display:none"></iframe>`;
    box.appendChild(d);
  });
  box.onclick = (e) => {
    const b = e.target.closest("button,input[type=checkbox]"); if (!b) return;
    const m = RESULT.mails[+b.dataset.i];
    if (b.dataset.act === "inc") { m.include = b.checked; return; }
    if (b.dataset.act === "eml") { download(buildEml(m), safeName(m.subject + "_" + m.vendor) + ".eml"); markSent([m.subject + "|" + m.vendor]); }
    else if (b.dataset.act === "saveto") {
      const emails = box.querySelector(`.toin[data-i="${b.dataset.i}"]`).value.split(/[,;\s]+/).map((x) => x.trim()).filter((x) => x.includes("@"));
      if (!emails.length) { toast("메일 주소를 적어 주세요"); return; }
      const all = loadOverlay(); const ov = all[CONFIG_NAME] || (all[CONFIG_NAME] = { vendors: {}, newVendors: [] }); ov.vendors = ov.vendors || {};
      const t = ov.vendors[m.vendor] = ov.vendors[m.vendor] || { to: [], makers: [], poNames: [] }; t.to = [...new Set([...(t.to || []), ...emails])]; saveOverlay(all);
      m.to = emails; m.noTo = false; m.include = true; log(`[설정 반영 · 이 브라우저에 저장] ${m.vendor}: 받는사람 + ${emails.join(",")}`, "ok");
      reloadConfig().then(renderResult); toast(`${m.vendor} 받는 사람을 저장했습니다`);
    } else { const f = b.closest(".mail").querySelector("iframe"); if (f.style.display === "none") { f.srcdoc = `<meta charset="utf-8"><div style="padding:8px">${m.html}</div>`; f.style.display = ""; } else f.style.display = "none"; }
  };
  renderAssign();
}
async function downloadZip() {
  const zip = new JSZip();
  const sel = RESULT.mails.filter((m) => m.include && !m.noTo);
  if (!sel.length) { toast("포함된 메일이 없습니다 (왼쪽 '포함' 체크 확인)"); return; }
  for (const m of sel) zip.file(safeName(m.subject + "_" + m.vendor) + ".eml", buildEml(m));
  markSent(sel.map((m) => m.subject + "|" + m.vendor));
  if (RESULT.extraFiles) for (const f of RESULT.extraFiles) zip.file(f.name, f.data);
  const blob = await zip.generateAsync({ type: "blob" });
  download(blob, (RESULT.kind === "po" ? "주문서메일_" : "출하요청메일_") + fmtDate(new Date()) + ".zip");
  toast(`메일 ${sel.length}통을 ZIP으로 받았습니다. 압축을 풀고 .eml 파일을 더블클릭하면 Outlook 새 메일 창이 열립니다.`);
  $("afterZip").style.display = "";
}

// ---------------------------------------------------------------- 전체 엑셀 (F1,2 / F5 시트)
function downloadSummaryXlsx() {
  const etdSel = RESULT.etdSel || [];
  const hdr = ["구분", "업체", "Fac", "ITEM CD", "MAKER", "ITEM NM", etdSel[0] || "요청수량", "입고수량", ...etdSel.slice(1), "Short Del.", "NEED ETD"];
  const groups = { "F1,2": [], "F5": [] };
  for (const m of RESULT.mails) {
    for (const r of m.ship || []) {
      const row = ["선박", m.vendor, r.Fac, r["ITEM CD"], r.MAKER, r["ITEM NM"], toNum(r[etdSel[0]]) || 0, null, ...etdSel.slice(1).map((c) => toNum(r[c]) || 0), "", ""];
      (norm(r.Fac).startsWith("f5") ? groups["F5"] : groups["F1,2"]).push(row);
    }
    for (const r of m.air || []) {
      const row = ["에어", m.vendor, r.FAC, r["Item cd"], r.Maker, r["Part Name"], toNum(r["Purchase request"]) || 0, null, ...etdSel.slice(1).map(() => ""), r["Short Del."], r["NEED ETD"]];
      (norm(r.FAC).startsWith("f5") ? groups["F5"] : groups["F1,2"]).push(row);
    }
  }
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(groups)) {
    const aoa = [hdr, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa); styleSheet(ws, aoa.length, hdr.length, 6);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  download(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })]), `출하요청_${RESULT.tag || fmtDate(new Date())}.xlsx`);
  toast(`엑셀을 받았습니다 (F1,2: ${groups["F1,2"].length}행 / F5: ${groups["F5"].length}행)`);
}

// ---------------------------------------------------------------- 미지정 업체 지정
function renderAssign() {
  const card = $("assignCard"); const tb = $("assignTbl").querySelector("tbody"); tb.innerHTML = "";
  if (!RESULT.unmatched.length) { card.style.display = "none"; return; }
  const kindName = { maker: "MAKER", po: "주문서 거래처명", vendor: "리스트 업체명" };
  const opts = ["(건너뜀)", "(새 업체)", ...CONFIG.order];
  RESULT.unmatched.forEach((it, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><span class="tag">${kindName[it.kind] || ""}</span><b>${esc(it.key)}</b> <span class="hint">(${it.count}건)</span></td><td>${esc((it.samples || []).slice(0, 2).join("; "))}</td>
      <td><select id="ch${i}">${opts.map((o) => `<option>${esc(o)}</option>`).join("")}</select></td><td><input type="text" id="ex${i}" style="width:100%" placeholder="업체이름, 메일주소"></td>`;
    tb.appendChild(tr);
  });
  card.style.display = "";
}
function saveAssign() {
  const all = loadOverlay(); const ov = all[CONFIG_NAME] || (all[CONFIG_NAME] = { vendors: {}, newVendors: [] });
  ov.vendors = ov.vendors || {}; ov.newVendors = ov.newVendors || [];
  const changes = [];
  const getNew = (name) => { let nv = ov.newVendors.find((x) => x.name === name); if (!nv) { nv = { name, to: [], makers: [], poNames: [] }; ov.newVendors.push(nv); changes.push("업체 추가: " + name); } return nv; };
  RESULT.unmatched.forEach((it, i) => {
    const choice = $("ch" + i).value, extra = $("ex" + i).value.trim();
    if (!choice || choice === "(건너뜀)") return;
    let target;
    if (choice === "(새 업체)") {
      const name = extra.split(",")[0].trim() || it.key;
      target = CONFIG.vendors[name] ? (ov.vendors[name] = ov.vendors[name] || { to: [], makers: [], poNames: [] }) : getNew(name);
      target._name = name;
      const emails = extra.split(",").slice(1).map((x) => x.trim()).filter((x) => x.includes("@"));
      if (emails.length) { target.to = [...new Set([...(target.to || []), ...emails])]; changes.push(`${name}: 받는사람 + ${emails.join(",")}`); }
    } else {
      target = ov.vendors[choice] = ov.vendors[choice] || { to: [], makers: [], poNames: [] }; target._name = choice;
      const emails = extra.split(",").map((x) => x.trim()).filter((x) => x.includes("@"));
      if (emails.length) { target.to = [...new Set([...(target.to || []), ...emails])]; changes.push(`${choice}: 받는사람 + ${emails.join(",")}`); }
    }
    if (it.kind === "maker" && it.key !== "(MAKER 없음)") { if (!target.makers.includes(it.key)) { target.makers.push(it.key); changes.push(`${target._name}: MAKER + ${it.key}`); } }
    else if (it.kind === "po" || it.kind === "vendor") { if (!target.poNames.includes(it.key)) { target.poNames.push(it.key); changes.push(`${target._name}: 거래처명 + ${it.key}`); } }
    delete target._name;
  });
  saveOverlay(all);
  log("[설정 반영 · 이 브라우저에 저장] " + (changes.length ? changes.join(" / ") : "변경 없음"), "ok");
  toast(changes.length ? "저장했습니다. ① 미리보기를 다시 누르세요." : "변경 없음");
  reloadConfig();
}
function downloadMergedConfig() {
  // 원본 config.xlsx 에 브라우저 저장분을 합쳐서 내려줌 (서식은 단순화됨)
  const wb = CONFIG.wb; const ws = wb.Sheets["업체"];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" , raw: true});
  const ov = loadOverlay()[CONFIG_NAME] || {};
  const addTo = (r, k, vals) => { const cur = splitList(r[k]); for (const v of vals || []) if (!cur.includes(v)) cur.push(v); r[k] = cur.join(", "); };
  for (const r of rows) { const o = ov.vendors?.[clean(r["업체"])]; if (o) { addTo(r, "받는사람", o.to); addTo(r, "MAKER", o.makers); addTo(r, "거래처명(주문서)", o.poNames); } }
  for (const nv of ov.newVendors || []) rows.push({ ...Object.fromEntries(Object.keys(rows[0] || {}).map((k) => [k, ""])), 업체: nv.name, 받는사람: nv.to.join(", "), MAKER: nv.makers.join(", "), "거래처명(주문서)": nv.poNames.join(", "), 엑셀첨부: "N", 활성: "Y", 발주서발송: "Y" });
  const out = XLSX.utils.book_new();
  for (const sn of wb.SheetNames) XLSX.utils.book_append_sheet(out, sn === "업체" ? XLSX.utils.json_to_sheet(rows) : wb.Sheets[sn], sn);
  download(new Blob([XLSX.write(out, { type: "array", bookType: "xlsx" })]), CONFIG_NAME);
  toast("합쳐진 설정 파일을 받았습니다. GitHub의 " + CONFIG_NAME + " 에 올리면 다른 PC에서도 적용됩니다.");
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
    const [month, week] = weekLabel(monday);
    const airDate = addDays(monday, 1), shipDate = addDays(monday, 2);
    const subjectBase = (S["제목형식"] || "대아홀딩스 {월}월 {주}주 출하 요청의 건").replace("{월}", month).replace("{주}", week);
    const includeZero = !["N", "NO", "X", "0"].includes(clean(S["수량없는행포함"] ?? "Y").toUpperCase());
    const etdCount = parseInt($("etd").value) || parseInt(S["ETD열개수"]) || 2;
    const only = splitList($("only").value);
    const commonCc = splitList(S["공통참조"]);

    let ship = [], etdSel = [], air = [];
    const wbShip = fShip ? XLSX.read(await fShip.arrayBuffer(), { type: "array", cellDates: true }) : null;
    const wbAir = fAir ? (fAir === fShip ? wbShip : XLSX.read(await fAir.arrayBuffer(), { type: "array", cellDates: true })) : null;
    if (wbShip) ({ rows: ship, etdSel } = parseShip(wbShip, etdCount));
    if (wbAir) air = parseAir(wbAir);
    else if (wbShip) { try { air = parseAir(wbShip); log(`[참고] 에어리스트를 따로 넣지 않아, 선박리스트 파일 안의 에어 표를 함께 사용합니다. (에어를 빼려면 고급 설정 → "같은 파일의 에어 표 자동 사용" 해제)`, "warn"); } catch (e) { log("에어리스트 없음 (선박 파일에 에어 표도 없음) → 선박만 진행"); } }
    if (!wbShip && wbAir) { try { ({ rows: ship, etdSel } = parseShip(wbAir, etdCount)); log("[참고] 선박리스트를 따로 넣지 않아, 에어리스트 파일 안의 선박 표를 함께 사용합니다.", "warn"); } catch (e) { log("선박리스트 없음 → 에어만 진행"); } }
    if (!$("autoBoth").checked) { if (!fAir) air = []; if (!fShip) { ship = []; etdSel = []; } }
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
      RESULT.mails.push({ vendor: n, subject, to: v.to, cc: [...new Set([...commonCc, ...v.cc])], html: buildShipBody(v, a, s, etdSel, airDate, shipDate, S, attached), attachments, note: v.memo ? "메모: " + v.memo : "", ship: s, air: a, noTo, include: !noTo });
      RESULT.etdSel = etdSel; RESULT.tag = `${month}월${week}주`;
    }
    log("");
    if (problems.length) { log("[문제] 아래 업체는 메일을 만들지 않았습니다:", "warn"); problems.forEach((p) => log("   " + p, "warn")); }
    if (unShip.length || unAir.length) {
      log(`[미지정 품목] 선박 ${unShip.length}건 / 에어 ${unAir.length}건 → 메일 생성 안 함 (아래 ③에서 업체를 지정하세요)`, "warn");
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
    const poPages = new Map(), order = [];   // key f|po -> {buf, pages:[], vendor, code}
    for (const f of files) {
      const buf = await f.arrayBuffer();
      for (const pg of await readPdfPages(buf)) {
        if (!pg.po) { log(`[경고] ${f.name} ${pg.page + 1}페이지: 발주번호를 못 읽었습니다. 건너뜁니다.`, "warn"); continue; }
        const key = f.name + "|" + pg.po;
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
      RESULT.mails.push({ vendor: n, subject, to: v.to, cc: [...new Set([...commonCc, ...v.cc])], html: buildPoBody(v, S), attachments, note, noTo: !v.to.length, include: v.to.length > 0, pos });
    }
    RESULT.skipped = skipped.map((e) => `PO ${e.po} ${e.vendor}`);
    if (skipped.length) { log(""); log("[발주서 미발송 업체 - 건너뜀]"); skipped.forEach((e) => log(`   PO ${e.po}  ${e.vendor}`)); }
    if (unknown.length) {
      log(""); log("[업체를 못 찾음 - 아래 ③에서 지정하세요]", "warn"); unknown.forEach((e) => log(`   PO ${e.po}  거래처명: ${e.vendor}  (코드 ${e.code})`, "warn"));
      const grp = {}; for (const e of unknown) (grp[e.vendor || "(거래처명 없음)"] = grp[e.vendor || "(거래처명 없음)"] || []).push("PO " + e.po);
      RESULT.unmatched = Object.entries(grp).map(([k, v]) => ({ key: k, kind: "po", count: v.length, samples: v }));
      for (const e of unknown) RESULT.extraFiles.push({ name: `미지정_${e.po}.pdf`, data: await splitPdf(e.buf, e.pages) });
    }
    log("=================================================="); log(`완료: 메일 ${RESULT.mails.length}통 준비됨.`, "ok"); setStatus("완료");
  } catch (e) { console.error(e); log("[오류] " + (e.message || e), "err"); setStatus("오류"); }
  renderResult(); $("resultCard").scrollIntoView({ behavior: "smooth" });
}

// ---------------------------------------------------------------- UI 연결
function bindDrop(zone, input, nameEl) {
  const z = $(zone), i = $(input);
  const show = () => { const fs = [...i.files]; $(nameEl).textContent = fs.length ? fs.map((f) => f.name).join(", ") : "아직 없음"; z.classList.toggle("has", fs.length > 0); };
  z.addEventListener("click", () => i.click());
  z.addEventListener("dragover", (e) => { e.preventDefault(); z.classList.add("over"); });
  z.addEventListener("dragleave", () => z.classList.remove("over"));
  z.addEventListener("drop", (e) => { e.preventDefault(); z.classList.remove("over"); i.files = e.dataTransfer.files; show(); i.dispatchEvent(new Event("change")); });
  i.addEventListener("change", show);
}
window.addEventListener("DOMContentLoaded", () => {
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  bindDrop("dropShip", "ship", "shipName"); bindDrop("dropAir", "air", "airName"); bindDrop("dropPdf", "pdf", "pdfName");
  $("ship").addEventListener("change", () => fillSheetSelect($("ship"), "shipSheet", ["ITEM CD", "MAKER"], [], ["ETD"], []));
  $("air").addEventListener("change", () => fillSheetSelect($("air"), "airSheet", ["Item cd", "Maker"], ["Remain", "Purchase request"], [], ["Remain"]));
  $("monday").value = fmtDate(nextMonday());
  { // 설정 파일 기본값: 주소의 ?who=kim → 계장님, 아니면 지난번 선택
    const q = new URLSearchParams(location.search).get("who"); let last = null; try { last = localStorage.getItem("mailmaker_cfg"); } catch {}
    const want = q === "kim" ? "config_kim.xlsx" : q === "nam" ? "config.xlsx" : last;
    if (want && [...$("config").options].some((o) => o.value === want)) $("config").value = want;
  }
  $("advToggle").onclick = () => { const a = $("adv"); a.style.display = a.style.display === "none" ? "" : "none"; $("advToggle").textContent = (a.style.display === "none" ? "▸" : "▾") + " 고급 설정"; };
  $("config").addEventListener("change", () => { $("configFile").style.display = $("config").value === "__file__" ? "" : "none"; if ($("config").value !== "__file__") reloadConfig(); });
  $("configFile").addEventListener("change", reloadConfig);
  $("btnShip").onclick = runShip; $("btnPo").onclick = runPo; $("btnZip").onclick = downloadZip; $("btnXlsx").onclick = downloadSummaryXlsx;
  $("btnAssign").onclick = saveAssign; $("btnDlCfg").onclick = downloadMergedConfig;
  $("btnReset").onclick = () => { if (confirm("이 브라우저에 저장된 추가 지정을 모두 지울까요?")) { const a = loadOverlay(); delete a[CONFIG_NAME]; saveOverlay(a); reloadConfig(); } };
  $("btnHelp").onclick = () => $("helpCard").scrollIntoView({ behavior: "smooth" });
  reloadConfig();
});
