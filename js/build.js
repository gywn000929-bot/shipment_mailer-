/* 주간 리스트 만들기: 받은 요청 파일(F1·F2·F5 선박, 에어 리포트) + 전산 Balance/Stock + 거래처 표
   → "선박 및 에어리스트" 엑셀. (엑셀 매크로 "선박리스트_만들기"와 같은 규칙)
   규칙
   - 선박: 각 요청 파일의 SUMMARY 시트에서 ITEM NM 바로 뒤 열들이 ETD 수량. SUM=ETD 합, SUM이 0인 줄은 뺀다.
   - BALANCE/STOCK: 공장별 시트에서 품목 코드로 찾는다. MAKER가 JST이고 값이 비어 있으면 JST 시트에서 다시 찾는다.
   - BALANCE 1 = BALANCE + STOCK, B-R = BALANCE 1 - SUM (에어는 B-R = BALANCE 1 - 요청 수량)
   - 거래처: 1. 예외표에서 M코드가 같고 품목명이 예외 품목명으로 시작 2. 품목명이 예외 품목명과 같음 3. MAKER별 기본 거래처
   - 에어: 요청 수량이 "600>>2,100"이면 2,100, "500>>0"이면 그 줄은 뺀다.
 - 정렬: FAC → MAKER → 품목명 (엑셀처럼 하이픈은 무시) */
"use strict";
const SLOTS = ["F1 Bal", "F2 Bal", "F5 Bal", "F1 JST Bal", "F2 JST Bal", "F5 JST Bal", "F1 Sto", "F2 Sto", "F5 Sto", "F1 JST Sto", "F2 JST Sto", "F5 JST Sto"];
const BUILD = { inputs: [], out: null };
const nsp = (s) => clean(s).replace(/\s+/g, "").toUpperCase();
const isCode = (v) => /^[MS]0/i.test(clean(v));

// ---------------------------------------------------------------- 파일·시트 알아보기
function slotOf(sheetName, fileName) {   // 매크로와 같은 시트 이름 규칙 + 파일 이름
  const tryName = (raw) => {
    const s = nsp(raw).replace(/[_\-.()]/g, "");
    const f = (s.match(/F([125])/) || s.match(/(?:BALANCE|JST|STO|STOCK)([125])/) || [])[1];
    if (!f) return "";
    const jst = s.includes("JST"), sto = /STO|재고/.test(s), bal = /BAL|미납/.test(s);
    if (!sto && !bal && !jst) return "";
    return `F${f} ${jst ? "JST " : ""}${sto ? "Sto" : "Bal"}`;
  };
  return tryName(sheetName) || tryName(fileName.replace(/\.[^.]+$/, ""));
}
function headerRowOf(rows, needs, max = 12) {
  for (let i = 0; i < Math.min(max, rows.length); i++) { const c = (rows[i] || []).map(nsp); if (needs.every((n) => c.includes(n))) return i; }
  return -1;
}
function readInputs(fileName, wb) {
  const out = [];
  for (const sn of wb.SheetNames) {
    const rows = sheetRows(wb, sn); if (!rows.length) continue;
    const base = { file: fileName, sheet: sn, rows, sheetNames: wb.SheetNames, ws: wb.Sheets[sn], styled: !!wb.__styled };
    if (headerRowOf(rows, ["매입업체", "MAKER"], 3) >= 0) { out.push({ ...base, kind: "vendor" }); continue; }
    if (/^SUMMARY/i.test(sn.trim()) && headerRowOf(rows, ["ITEMCD"]) >= 0) { let facs = []; try { facs = parseSea(base).facs; } catch (e) { console.warn(e); } out.push({ ...base, kind: "sea", facs }); continue; }
    if (headerRowOf(rows, ["ITEMCD", "PARTNAME", "PURCHASEREQUEST"]) >= 0) { out.push({ ...base, kind: "air" }); continue; }
    const hb = headerRowOf(rows, ["거래처품목코드", "미납수량합계"], 5);
    if (hb >= 0) {   // 전산에서 받은 미납수량 합계 원본: 수주처 칸(F1/F2/F5)으로 공장을 알고, JST 여부는 파일 이름으로
      const cF = rows[hb].map(nsp).indexOf("수주처"), fac = cF >= 0 ? nsp((rows.slice(hb + 1).find((r) => r && clean(r[cF])) || [])[cF]) : "";
      const byName = slotOf(sn, "") || slotOf(fileName.replace(/\.[^.]+$/, ""), "");
      out.push({ ...base, kind: "bal", slot: /^F[125]$/.test(fac) ? `${fac} ${/JST/i.test(sn + fileName) ? "JST " : ""}Bal` : byName }); continue;
    }
    const hf = headerRowOf(rows, ["FAC", "ITEMCD"], 5);
    if (hf >= 0) {   // 전산 Stock(출고수량 합계) 모양: FAC | ITEM CD | Part Name | Shipped Q'ty. 공장은 FAC 칸으로, Stock/Balance는 머리글·이름으로
      const hd = rows[hf].map(nsp), cF = hd.indexOf("FAC"), cC = hd.indexOf("ITEMCD");
      const cV = hd.findIndex((c, i) => i > cC && /SHIPPED|QTY|Q'TY|수량|STOCK|BALANCE/.test(c));
      if (cV >= 0 && !hd.includes("MAKER") && hd.filter(Boolean).length <= 6) {   // 지난주 리스트(MAKER·ETD 열이 많은 표)는 여기서 받지 않는다
        const nm = nsp(sn + " " + fileName), isBal = /BAL|미납/.test(nm + hd[cV]) && !/SHIPPED|STO|재고|출고/.test(nm + hd[cV]);
        const jst = /JST/.test(nm), kind = isBal ? "Bal" : "Sto";
        const facs = uniq(rows.slice(hf + 1).map((r) => nsp((r || [])[cF])).filter((f) => /^F[125]$/.test(f)));
        for (const f of facs) {
          const slot = `${f} ${jst ? "JST " : ""}${kind}`;
          const own = facs.length === 1;   // 공장 하나만 든 시트(F1·F2·F5)가 합계 시트보다 먼저
          const prev = out.findIndex((x) => x.slot === slot);
          if (prev >= 0 && !own) continue;
          if (prev >= 0) out.splice(prev, 1);
          out.push({ ...base, kind: "bal", slot, codeCol: cC, valCol: cV, rows: rows.filter((r, i) => i > hf && nsp((r || [])[cF]) === f) });
        }
        if (facs.length) continue;
      }
    }
    const slot = slotOf(sn, wb.SheetNames.length === 1 ? fileName : "");
    const looksBal = rows.slice(0, 30).some((r) => (r || []).slice(0, 10).some(isCode)) && rows.length > 1 && (rows[0] || []).length <= 12;
    if (slot && looksBal) { out.push({ ...base, kind: "bal", slot }); continue; }
  }
  return out;
}
async function addBuildFiles(files) {
  for (const f of files) {
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true, cellStyles: true, bookFiles: true });
      try { enrichStyles(wb); } catch (e) { console.warn("칸 서식을 읽지 못함", e); }
      delete wb.files; delete wb.keys;
      addWorkbook(f.name, wb);
    } catch (e) { console.error(e); toast(`${f.name}: 읽지 못했습니다 (${e.message || e})`); }
  }
  afterAdd();
}
function addWorkbook(name, wb) {
  const got = readInputs(name, wb);
  if (!got.length) { toast(`${name}: 쓸 수 있는 표를 찾지 못했습니다`); return 0; }
  for (const g of got) {   // 같은 종류·같은 자리는 새 파일로 바꾼다
    BUILD.inputs = BUILD.inputs.filter((x) => !(x.kind === g.kind && (g.kind === "vendor" || (g.kind === "bal" && g.slot && x.slot === g.slot) || (x.file === g.file && x.sheet === g.sheet))));
    BUILD.inputs.push(g);
  }
  return got.length;
}
// 메일 본문의 표를 복사해 붙여넣기: 아웃룩에서 복사하면 HTML 표가 같이 오므로 그것을 먼저 쓰고, 없으면 탭으로 나뉜 글자를 쓴다
function pastedRows(html, text) {
  const cellTxt = (el) => el.textContent.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  if (html && /<t[dh][\s>]/i.test(html)) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const tables = [...doc.querySelectorAll("table")].map((t) => [...t.rows].map((r) => [...r.cells]));
    const best = tables.filter((t) => t.some((r) => r.map((el) => nsp(cellTxt(el))).includes("ITEMCD"))).sort((a, b) => a.length - b.length)[0]   // 표 안에 표가 있으면 가장 안쪽(머리글이 든 가장 작은 표)
      || tables.sort((a, b) => b.length - a.length)[0];
    if (best && best.length) { const out = best.map((r) => r.map(cellTxt)); out.styles = best.map((r) => r.map(htmlCellStyle)); return out; }
  }
  const lines = String(text || "").replace(/\u00a0/g, " ").split(/\r?\n/);
  if (lines.some((l) => l.includes("\t"))) return lines.filter((l) => l.trim()).map((l) => l.split("\t").map((c) => c.trim()));
  // 탭 없이 한 칸이 한 줄씩 온 경우(메일 글자만 복사): "No" 머리글부터 첫 번호 줄 전까지가 머리글, 그 칸 수만큼 묶는다
  const t = lines.map((l) => l.trim()), h0 = t.findIndex((l, i) => nsp(l) === "NO" && nsp(t[i + 1] || "") === "FAC");
  if (h0 < 0) return [];
  let h1 = h0 + 1; while (h1 < t.length && !/^\d+$/.test(t[h1])) h1++;
  const hdr = t.slice(h0, h1), n = hdr.length, out = [hdr];
  for (let i = h1; i + 1 < t.length && /^\d+$/.test(t[i]) && /^F\d$/i.test(t[i + 1]); i += n) out.push(t.slice(i, i + n));
  return out;
}
// 메일 표 칸의 색: 바탕색(td·안쪽 글자의 background), 글자색(color), 굵게. 검정 글자·흰 바탕은 기본이라 뺀다
function cssHex(v) {
  if (!v || /windowtext|auto|transparent|inherit|initial/i.test(v)) return "";
  const d = document.createElement("div"); d.style.color = ""; d.style.color = v.trim(); if (!d.style.color) return "";
  document.body.appendChild(d); const m = getComputedStyle(d).color.match(/\d+/g); d.remove();
  return m ? m.slice(0, 3).map((x) => (+x).toString(16).padStart(2, "0")).join("").toUpperCase() : "";
}
function htmlCellStyle(td) {
  const els = [td, ...td.querySelectorAll("*")], get = (re) => { for (const el of els) { const st = el.getAttribute("style") || "", m = st.match(re); if (m) return m[1]; } return ""; };
  const bg = cssHex(td.getAttribute("bgcolor") || get(/background(?:-color)?\s*:\s*([^;]+)/i));
  const fc = cssHex(get(/(?:^|;)\s*color\s*:\s*([^;]+)/i) || ((td.querySelector("font[color]") || { getAttribute: () => "" }).getAttribute("color")));
  const bold = !!td.querySelector("b,strong") || /font-weight\s*:\s*(bold|[6-9]00)/i.test(els.map((e) => e.getAttribute("style") || "").join(";"));
  const s = {};
  if (bg && bg !== "FFFFFF") s.fill = { patternType: "solid", fgColor: { rgb: bg } };
  if ((fc && fc !== "000000") || bold) s.font = { ...(fc && fc !== "000000" ? { color: { rgb: fc } } : {}), ...(bold ? { bold: true } : {}) };
  return s.fill || s.font ? s : null;
}
// 엑셀(xlsx) 칸의 색: cellStyles 옵션은 바탕색만 주므로, 칸마다 적힌 서식 번호로 styles.xml의 글꼴(글자색·굵게)까지 찾아 붙인다. xls는 색을 읽지 못한다
function enrichStyles(wb) {
  const files = wb.files, St = wb.Styles; if (!files || !St || !St.CellXf || !files["xl/workbook.xml"] && !files["/xl/workbook.xml"]) return;
  const txt = (p) => { const f = files[p] || files["/" + p]; return !f ? "" : typeof f.content === "string" ? f.content : new TextDecoder().decode(f.content); };
  const rels = {}; for (const m of txt("xl/_rels/workbook.xml.rels").matchAll(/<Relationship\b[^>]*>/g)) { const id = (m[0].match(/\bId="([^"]+)"/) || [])[1], t = (m[0].match(/\bTarget="([^"]+)"/) || [])[1]; if (id && t) rels[id] = t.replace(/^\/?(xl\/)?/, "xl/"); }
  const unx = (s) => s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  for (const m of txt("xl/workbook.xml").matchAll(/<sheet\b[^>]*>/g)) {
    const name = unx((m[0].match(/\bname="([^"]*)"/) || [])[1] || ""), rid = (m[0].match(/\br:id="([^"]+)"/) || [])[1], ws = wb.Sheets[name], xml = rid && rels[rid] ? txt(rels[rid]) : "";
    if (!ws || !xml) continue;
    for (const c of xml.matchAll(/<c\b[^>]*>/g)) {
      const a = (c[0].match(/\br="([A-Z]+\d+)"/) || [])[1], si = (c[0].match(/\bs="(\d+)"/) || [])[1]; if (!a || !si || !ws[a]) continue;
      const xf = St.CellXf[+si] || {}, font = St.Fonts && St.Fonts[+xf.fontId] || {}, fill = St.Fills && St.Fills[+xf.fillId] || {}, s = {};
      if (fill.patternType === "solid" && fill.fgColor) s.fill = { patternType: "solid", fgColor: fill.fgColor };
      const fc = font.color && (font.color.rgb ? (/^(FF)?000000$/i.test(font.color.rgb) ? null : font.color) : font.color.theme != null && font.color.theme !== 1 ? font.color : null);
      if (fc || font.bold || font.italic) s.font = { ...(fc ? { color: fc } : {}), ...(font.bold ? { bold: true } : {}), ...(font.italic ? { italic: true } : {}) };
      ws[a].s = s.fill || s.font ? s : undefined;
    }
  }
  wb.__styled = true;
}
// Short Del.·NEED ETD 날짜 글자("02 December 2026", "27-Nov-26", "2026-12-02")를 날짜로. URGENT 같은 글자는 그대로
function asDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const t = toDate(v); if (t) return t;
  const s = clean(v), m = s.match(/^(\d{1,2})[\s\-.]+([A-Za-z]{3,})[\s\-.,]+(\d{2,4})$/) || s.match(/^([A-Za-z]{3,})[\s\-.]+(\d{1,2}),?[\s\-.]+(\d{2,4})$/);
  if (m) { const [dd, mon, yy] = /^\d/.test(m[1]) ? [m[1], m[2], m[3]] : [m[2], m[1], m[3]]; const mi = MON.findIndex((x) => x.toLowerCase() === mon.slice(0, 3).toLowerCase()); if (mi >= 0) return new Date(yy.length === 2 ? 2000 + +yy : +yy, mi, +dd); }
  const iso = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/); if (iso) return new Date(+iso[1], iso[2] - 1, +iso[3]);
  return null;
}
// 요청 수량 읽기: 칸 앞뒤 공백·nbsp, 전각 숫자, "72.000"(점 천 단위), "72 000", "72,000 PCS" 같은 모양도 읽는다. 숫자가 없으면 NaN
function qtyOf(v) {
  if (typeof v === "number") return v;
  let s = String(v == null ? "" : v).replace(/[\uFF10-\uFF19]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xFEE0)).replace(/[\s\u00a0\u200b\u3000]/g, "");
  s = s.split(">>").pop();   // "600>>2,100"처럼 수량이 바뀐 칸은 바뀐 뒤 수량
  if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, "");
  const m = s.replace(/,/g, "").match(/-?\d+(\.\d+)?/);
  return m ? Number(m[0]) : NaN;
}
let pasteNo = 0;
function addPasted(html, text) {
  const rows = pastedRows(html, text);
  if (rows.length < 2) { toast("붙여넣은 내용에서 표를 찾지 못했습니다. 메일의 표를 머리글(No, FAC, Item cd …)부터 끝까지 골라 복사해 주세요."); return; }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  if (rows.styles) rows.styles.forEach((r, i) => r.forEach((st, j) => { const a = XLSX.utils.encode_cell({ r: i, c: j }); if (st && ws[a]) ws[a].s = st; }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "메일 표"); wb.__styled = !!rows.styles;
  const facs = uniq(rows.map((r) => nsp(r[1])).filter((f) => /^F[125]$/.test(f)));
  const name = `붙여넣기 ${++pasteNo}${facs.length ? " (" + facs.join(",") + ")" : ""}`;
  if (addWorkbook(name, wb)) toast(`${name}: ${rows.length - 1}줄을 넣었습니다`);
  afterAdd();
}
function afterAdd() {
  const v = BUILD.inputs.find((x) => x.kind === "vendor");
  if (v) store.set(LS_VTBL, { rows: v.rows.map((r) => (r || []).slice(0, 8).map((c) => (c == null ? "" : clean(c)))), from: v.file, at: fmtDate(new Date()) });
  BUILD.out = null; renderBuild();
}

// ---------------------------------------------------------------- 선박 요청 파일 읽기
const etdKey = (s) => { const m = clean(s).toUpperCase().match(/(\d{1,2})[\s\-]*([A-Z]{3})/); return m ? `${+m[1]}-${m[2]}` : ""; };   // "ETD 04-DEC-26", "ETD 4 DEC" → "4-DEC"
function parseSea(inp) {
  const rows = inp.rows, h = headerRowOf(rows, ["ITEMCD"]);
  const hd = rows[h].map(nsp), cCd = hd.indexOf("ITEMCD"), cFac = cCd - 1;
  const cMk = hd.indexOf("MAKER"), cNm = hd.findIndex((c) => c === "ITEMNM");
  if (cMk < 0 || cNm < 0) throw new Error(`${inp.file}: SUMMARY 시트에서 MAKER / ITEM NM 열을 찾지 못했습니다`);
  let names = [];
  for (let c = cNm + 1; c < hd.length && hd[c].startsWith("ETD"); c++) names.push(clean(rows[h][c]));
  if (!names.length) names = inp.sheetNames.filter((s) => /ETD\s*\d/i.test(s)).map((s) => "ETD " + (s.toUpperCase().match(/\d{1,2}[\s\-]*[A-Z]{3}(?:[\s\-]*\d{2})?/) || [""])[0].trim());   // F2처럼 머리글에 날짜가 없으면 시트 이름에서
  if (!names.length) throw new Error(`${inp.file}: ETD 열을 알아내지 못했습니다`);
  const data = [];
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i] || [], fac = nsp(r[cFac]);
    if (!/^F\d$/.test(fac) || !clean(r[cCd])) continue;
    data.push({ fac, cd: clean(r[cCd]), maker: clean(r[cMk]), nm: clean(r[cNm]), q: names.map((_, k) => toNum(r[cNm + 1 + k])) });
  }
  return { names, data, facs: uniq(data.map((d) => d.fac)) };
}

// ---------------------------------------------------------------- Balance / Stock / 거래처 표
function lookupTable(inp) {   // 매크로처럼: 처음 나오는 M0…/S0… 칸의 열이 코드, 그 오른쪽 두 번째 열이 값. 같은 코드는 첫 줄
  const rows = inp.rows; let cc = -1;
  if (inp.valCol != null) { const m = new Map(); for (const r of rows) { const k = clean((r || [])[inp.codeCol]); if (k && !m.has(k)) m.set(k, (r || [])[inp.valCol]); } return m; }
  for (let r = 0; r < Math.min(30, rows.length) && cc < 0; r++) for (let c = 0; c < 10; c++) if (isCode((rows[r] || [])[c])) { cc = c; break; }
  const map = new Map(); if (cc < 0) return map;
  for (const r of rows) { const k = clean((r || [])[cc]); if (isCode(k) && !map.has(k)) map.set(k, (r || [])[cc + 2]); }
  return map;
}
const vendorRule = vtRule;   // js/vendor_table.js
const xlSortKey = (s) => clean(s).toUpperCase().replace(/[-']/g, "");
const cmpStr = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
const rowSort = (a, b) => cmpStr(a.fac, b.fac) || cmpStr(xlSortKey(a.maker), xlSortKey(b.maker)) || cmpStr(xlSortKey(a.nm), xlSortKey(b.nm)) || cmpStr(a.nm, b.nm);

// ---------------------------------------------------------------- 만들기
function buildLists() {
  const notes = [], I = BUILD.inputs;
  const seas = I.filter((x) => x.kind === "sea").map(parseSea);
  const airIns = I.filter((x) => x.kind === "air");
  if (!seas.length && !airIns.length) throw new Error("선박 요청 파일(F1·F2·F5)이나 에어 리포트를 넣어 주세요.");
  const vt = I.find((x) => x.kind === "vendor") || currentVendorTable();   // 넣은 표 → 이 PC에 기억한 새 표 → 내장 표
  const vendorOf = vt ? vendorRule(vt.rows) : () => "";
  const tbl = {}; for (const b of I.filter((x) => x.kind === "bal" && x.slot)) tbl[b.slot] = lookupTable(b);
  for (const b of I.filter((x) => x.kind === "bal" && !x.slot)) notes.push(`${b.file}: 어느 공장의 Balance인지 몰라 쓰지 않았습니다. 위 표에서 골라 주세요.`);
  const find = (fac, cd, maker, kind) => {   // kind: "Bal" | "Sto"
    let v = (tbl[`${fac} ${kind}`] || new Map()).get(cd);
    if (clean(maker).toUpperCase() === "JST" && (v == null || clean(v) === "" || toNum(v) === 0)) { const j = (tbl[`${fac} JST ${kind}`] || new Map()).get(cd); if (j != null && clean(j) !== "") v = j; }
    return v == null || clean(v) === "" ? null : toNum(v);
  };
  // ---- 선박
  let etd = [], ship = [];
  if (seas.length) {
    const hasYear = (s) => (s.names.every((n) => /\d{2}\s*$/.test(n.replace(/^ETD\s*\d+/i, ""))) ? 1 : 0);
    const master = seas.slice().sort((a, b) => b.names.length - a.names.length || hasYear(b) - hasYear(a))[0];
    etd = master.names.slice();
    const pos = Object.fromEntries(etd.map((n, i) => [etdKey(n), i]));
    for (const s of seas) {
      const idx = s.names.map((n) => pos[etdKey(n)]);
      if (idx.some((x) => x == null)) notes.push(`${s.facs.join(",")} 요청 파일의 ETD(${s.names.join(", ")}) 중 기준 날짜(${etd.join(", ")})에 없는 것이 있어 순서대로 맞췄습니다. 날짜를 확인하세요.`);
      for (const d of s.data) {
        const q = etd.map(() => 0); d.q.forEach((v, k) => { const p = idx[k] == null ? k : idx[k]; if (p < q.length) q[p] += v; });
        const sum = q.reduce((a, b) => a + b, 0); if (sum === 0) continue;
        const bal = find(d.fac, d.cd, d.maker, "Bal"), sto = find(d.fac, d.cd, d.maker, "Sto"), b1 = (bal || 0) + (sto || 0);
        ship.push({ ...d, q, sum, bal, sto, b1, br: b1 - sum, vendor: vendorOf(d.cd, d.nm, d.maker) });
      }
      for (const f of s.facs) if (!tbl[`${f} Bal`]) notes.push(`${f} Balance 파일이 없어 ${f} 품목의 BALANCE가 비어 있습니다.`);
    }
    ship.sort(rowSort);
  }
  // ---- 에어
  let air = [];
  const seen = new Set();   // 같은 에어 요청을 파일로도 넣고 붙여넣기로도 넣었을 때 한 번만
  for (const airIn of airIns) {
    const rows = airIn.rows, h = headerRowOf(rows, ["ITEMCD", "PARTNAME", "PURCHASEREQUEST"]), hd = rows[h].map(nsp), c = (n) => hd.indexOf(n);
    const rg = airIn.ws && airIn.ws["!ref"] ? XLSX.utils.decode_range(airIn.ws["!ref"]) : { s: { r: 0, c: 0 } };
    const stAt = (i, col) => { if (col < 0 || !airIn.ws) return null; const x = airIn.ws[XLSX.utils.encode_cell({ r: rg.s.r + i, c: rg.s.c + col })]; return x && x.s && (x.s.fill || x.s.font) ? x.s : null; };
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [], fac = nsp(r[c("FAC")]), cd = clean(r[c("ITEMCD")]); if (!cd || !/^F\d$/.test(fac)) continue;
      const maker = clean(r[c("MAKER")]), nm = clean(r[c("PARTNAME")]), rawReq = r[c("PURCHASEREQUEST")];
      let req = qtyOf(rawReq);
      if (req === 0 || clean(rawReq) === "") continue;   // "500>>0": 요청 취소, 빈 칸은 전처럼 뺀다
      if (isNaN(req)) { notes.push(`에어 ${fac} ${cd} ${nm}: 요청 수량 "${clean(rawReq)}"을 숫자로 읽지 못해 0으로 넣었습니다. 엑셀에서 직접 고쳐 주세요.`); req = 0; }
      const key = `${fac}|${cd}|${req}|${clean(r[c("NEEDETD")])}`; if (seen.has(key)) continue; seen.add(key);
      const bal = find(fac, cd, maker, "Bal"), sto = find(fac, cd, maker, "Sto"), b1 = (bal || 0) + (sto || 0);
      const dOr = (v) => asDate(v) || v;   // 메일 표의 "02 December 2026"도 F1·F2 파일처럼 날짜 칸으로
      air.push({ fac, cd, maker, nm, req, bal, sto, b1, br: b1 - req, short: dOr(r[c("SHORTDEL.")]), need: dOr(r[c("NEEDETD")]),
        shortS: stAt(i, c("SHORTDEL.")), needS: stAt(i, c("NEEDETD")), styled: airIn.styled, remark: c("REMARK") >= 0 ? clean(r[c("REMARK")]) : "", vendor: vendorOf(cd, nm, maker) });
    }
  }
  // F5 칸 색은 F1·F2 리스트에 맞춘다: F1·F2(색을 읽을 수 있는 파일)에서 같은 값 → 같은 종류(URGENT/날짜/글자) 칸의 색을 쓰고, 없으면 F5 원래 색
  const vKey = (v) => (v instanceof Date ? "D:" + fmtCellDate(dayOf(v)) : "T:" + clean(v).toUpperCase());
  const vCls = (v) => (v instanceof Date ? "D" : /URGENT/i.test(clean(v)) ? "U" : clean(v) ? "T" : "E");
  for (const [f, sk] of [["short", "shortS"], ["need", "needS"]]) {
    const ref = new Map();
    for (const r of air) if (r.fac !== "F5" && r.styled) for (const k of [vKey(r[f]), vCls(r[f])]) if (!ref.has(k)) ref.set(k, r[sk]);
    if (ref.size) for (const r of air) if (r.fac === "F5") { const k = ref.has(vKey(r[f])) ? vKey(r[f]) : vCls(r[f]); if (ref.has(k)) r[sk] = ref.get(k); }
  }
  air.sort(rowSort);
  const noVendor = uniq([...ship, ...air].filter((r) => !r.vendor).map((r) => r.maker || "(MAKER 없음)"));
  if (vt && noVendor.length) notes.push(`거래처를 못 붙인 MAKER: ${noVendor.join(", ")} → 거래처 표에 MAKER를 추가하세요.`);
  BUILD.out = { etd, ship, air, notes, vt, tbl };
  return BUILD.out;
}

// ---------------------------------------------------------------- 엑셀로 (매크로와 같은 모양)
function listWorkbook(o) {
  const thin = { style: "thin", color: { rgb: "000000" } }, border = { top: thin, bottom: thin, left: thin, right: thin };
  const font = (x = {}) => ({ name: "맑은 고딕", sz: 10, ...x }), NF = "#,##0;[Red]-#,##0;";
  const FAC = { F1: "FFFFFF", F2: "FFFF00", F5: "00B0F0" };
  const cell = (v, st = {}) => {
    const c = { v: v == null ? "" : v, t: typeof v === "number" ? "n" : "s", s: { border, font: font(st.bold ? { bold: true } : {}), alignment: { horizontal: st.align || "left", vertical: "center" } } };
    if (st.bg) c.s.fill = { patternType: "solid", fgColor: { rgb: st.bg } };
    if (c.t === "n") c.z = NF;
    return c;
  };
  const num = (v, st) => cell(v == null ? "" : v, { align: "center", ...st });
  const wb = XLSX.utils.book_new();
  if (o.ship.length) {
    const hdr = ["Fac", "ITEM CD", "MAKER", "ITEM NM", ...o.etd, "SUM", "BALANCE", "STOCK", "BALANCE 1", "B-R", "거래처명(매칭)"];
    let grp = 0, prev = ""; const tint = o.etd.map((n) => { const mo = (n.toUpperCase().match(/JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC/) || [""])[0]; if (prev && mo !== prev) grp++; prev = mo; return grp % 2 ? "E4F0DC" : "DBE9F7"; });
    const aoa = [hdr.map((h, i) => cell(h, { bold: true, align: "center", bg: i >= 4 && i < 4 + o.etd.length ? tint[i - 4] : "FFFFFF" }))];
    for (const r of o.ship) aoa.push([cell(r.fac, { align: "center", bg: FAC[r.fac] }), cell(r.cd), cell(r.maker), cell(r.nm), ...r.q.map((v, k) => num(v, { bg: tint[k] })), num(r.sum, { bold: true }), num(r.bal), num(r.sto), num(r.b1), num(r.br, { bold: true }), cell(r.vendor)]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [5, 14, 12, 34, ...o.etd.map(() => 11), 11, 11, 11, 11, 11, 24].map((w) => ({ wch: w }));
    ws["!rows"] = [{ hpt: 26 }]; ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: hdr.length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws, "선박");
  }
  if (o.air.length) {
    const hdr = ["FAC", "Item cd", "Maker", "Part Name", "Purchase request", "BALANCE", "STOCK", "BALANCE 1", "B-R", "Short Del.", "NEED ETD", "REMARK", "거래처명(매칭)"];
    const aoa = [hdr.map((h) => cell(h, { bold: true, align: "center", bg: "FFFFFF" }))];
    const dv0 = (v) => { if (!(v instanceof Date) || isNaN(v)) return cell(clean(v), { align: "center" }); const d = dayOf(v), c = cell((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000, { align: "center" }); c.z = "dd-mmm-yy"; return c; };
    const dv = (v, st) => {   // 받은 파일의 칸 색(바탕·글자색·굵게)을 그대로 옮긴다
      const c = dv0(v); if (!st) return c;
      if (st.fill) c.s.fill = st.fill;
      if (st.font) c.s.font = { ...c.s.font, ...(st.font.color ? { color: st.font.color } : {}), ...(st.font.bold ? { bold: true } : {}), ...(st.font.italic ? { italic: true } : {}) };
      return c;
    };
    for (const r of o.air) aoa.push([cell(r.fac, { align: "center", bg: FAC[r.fac] }), cell(r.cd), cell(r.maker), cell(r.nm), num(r.req, { bold: true }), num(r.bal), num(r.sto), num(r.b1), num(r.br, { bold: true }), dv(r.short, r.shortS), dv(r.need, r.needS), cell(r.remark), cell(r.vendor)]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [5, 14, 14, 34, 14, 11, 11, 11, 11, 12, 12, 16, 24].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, "Air");
  }
  if (o.vt) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(o.vt.rows), "거래처코드");
  for (const b of BUILD.inputs.filter((x) => x.kind === "bal" && x.slot)) if (!wb.SheetNames.includes(b.slot)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(b.rows.map((r) => (r || []).slice(0, 12))), b.slot);
  return wb;
}
const listBytes = () => new Uint8Array(XLSX.write(listWorkbook(BUILD.out), { type: "array", bookType: "xlsx" }));
function listFileName() { const d = BUILD.out.etd[0] ? BUILD.out.etd[0].replace(/^ETD\s*/i, "").replace(/\s+/g, "-") : fmtDate(new Date()); return `선박 및 에어리스트_ETD ${d}.xlsx`; }

// ---------------------------------------------------------------- 화면
const KIND_KO = { sea: "선박 요청", air: "에어 리포트", bal: "전산 Balance/Stock", vendor: "거래처 표" };
function renderBuild() {
  const I = BUILD.inputs, saved = currentVendorTable();
  $("bdList").innerHTML = I.length ? "<tr><th>파일</th><th>시트</th><th>종류</th><th>자리</th></tr>" + I.map((x, i) => `<tr><td>${esc(x.file)}</td><td>${esc(x.sheet)}</td><td>${KIND_KO[x.kind]}</td><td>${x.kind === "bal"
    ? `<select data-b="${i}"><option value="">(골라 주세요)</option>${SLOTS.map((s) => `<option ${s === x.slot ? "selected" : ""}>${s}</option>`).join("")}</select>`
    : x.kind === "sea" ? esc((x.facs || []).join(", ")) : ""}</td></tr>`).join("") : "";
  const has = (k) => I.some((x) => x.kind === k), slots = I.filter((x) => x.kind === "bal" && x.slot).map((x) => x.slot);
  const chip = (ok, t) => `<span class="chip ${ok ? "on" : ""}">${ok ? "✔" : "·"} ${t}</span>`;
  const facs = uniq(I.filter((x) => x.kind === "sea").flatMap((x) => x.facs || []));
  $("bdNeed").innerHTML = ["F1", "F2", "F5"].map((f) => chip(facs.includes(f), `${f} 선박 요청`)).join("") + chip(has("air"), "에어 리포트")
    + ["F1", "F2", "F5"].map((f) => chip(slots.includes(`${f} Bal`), `${f} Balance`)).join("") + chip(slots.some((s) => s.includes("JST")), "JST Balance") + chip(slots.some((s) => s.endsWith("Sto")), "Stock")
    + chip(true, has("vendor") ? "거래처 표 (넣은 파일)" : saved === BUILTIN_VTBL ? `거래처 표 (내장 · ${saved.at})` : `거래처 표 (이 PC에 넣은 것 · ${saved.at})`);
  $("btnBuild").disabled = !(has("sea") || has("air"));
  const o = BUILD.out; $("bdResult").style.display = o ? "" : "none";
  if (o) {
    const byFac = (rows) => ["F1", "F2", "F5"].map((f) => `${f} ${rows.filter((r) => r.fac === f).length}`).join(" · ");
    $("bdSummary").innerHTML = `<div><span class="ok">✅ 리스트를 만들었습니다.</span> 선박 <b>${o.ship.length}</b>줄 (${byFac(o.ship)})${o.etd.length ? ` · ETD ${o.etd.length}개 (${esc(o.etd[0].replace(/^ETD\s*/i, ""))} ~)` : ""} / 에어 <b>${o.air.length}</b>줄</div>`
      + o.notes.map((n) => `<div class="warn">⚠ ${esc(n)}</div>`).join("")
      + `<div class="hint">B-R이 음수(부족)인 품목: 선박 ${o.ship.filter((r) => r.br < 0).length}줄 / 에어 ${o.air.filter((r) => r.br < 0).length}줄 · BALANCE를 못 찾은 품목: 선박 ${o.ship.filter((r) => r.bal == null).length}줄</div>`;
  }
}
function bindBuild() {
  bindDrop("dropBuild", "bdFiles", "bdName");
  $("bdFiles").addEventListener("change", () => addBuildFiles([...$("bdFiles").files]));
  const pb = $("bdPaste");
  if (pb) {
    pb.addEventListener("paste", (e) => { const d = e.clipboardData; if (!d) return; e.preventDefault(); addPasted(d.getData("text/html"), d.getData("text/plain")); pb.value = ""; });
    $("btnPaste").onclick = () => { addPasted("", pb.value); pb.value = ""; };
  }
  $("bdList").addEventListener("change", (e) => { const t = e.target; if (t.dataset.b == null) return; BUILD.inputs[+t.dataset.b].slot = t.value; BUILD.out = null; renderBuild(); });
  $("btnBuildClear").onclick = () => { BUILD.inputs = []; BUILD.out = null; $("bdFiles").value = ""; $("bdName").textContent = "아직 없음"; $("dropBuild").classList.remove("has"); renderBuild(); };
  $("btnBuild").onclick = () => { try { buildLists(); } catch (e) { console.error(e); toast(String(e.message || e)); BUILD.out = null; } renderBuild(); };
  $("btnBuildDl").onclick = () => download(new Blob([listBytes()]), listFileName());
  $("btnBuildUse").onclick = () => {   // 만든 리스트를 메일 작성기 화면으로 넘긴다 (이 PC의 브라우저 안에서만 전달)
    try {
      let s = ""; const bytes = listBytes(); for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      localStorage.setItem(HANDOFF_KEY, JSON.stringify({ name: listFileName(), b64: btoa(s) }));
      location.href = "index.html" + location.search;
    } catch (e) { console.error(e); toast("바로 넘기지 못했습니다. '리스트 엑셀 받기'로 받은 뒤 메일 작성기에 넣어 주세요."); }
  };
  renderBuild();
}
