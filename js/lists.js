/* 선박리스트 · 에어리스트 엑셀 읽기 */
"use strict";
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
// 표가 여러 시트에 있을 때 고르는 점수: Remain 같은 우선 열이 있으면 +, ETD 열이 많을수록 +(업체 첨부용 작은 표보다 원본 리스트 우선)
const sheetScore = (hn, prefer, prefixAny) => prefer.filter((p) => hn.includes(norm(p))).length * 100 + hn.filter((c) => prefixAny.some((p) => c.startsWith(norm(p)))).length;
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
    found.push({ sheet: sn, headers: headers.filter(Boolean), rows: data, score: sheetScore(hn, prefer, prefixAny) });
  }
  if (!found.length) throw new Error(`${label} 파일에서 ${label} 표(헤더 ${mustHave.join(", ")}${anyOf.length ? " + " + anyOf.join("/") : ""}${prefixAny.length ? " + ETD열" : ""})가 있는 시트를 못 찾았습니다. 시트: ${wb.SheetNames.join(", ")}${sheetName ? " / 지정한 시트: " + sheetName : ""}. 선박리스트와 에어리스트에 같은 파일을 고른 것은 아닌지 확인하세요.`);
  found.sort((a, b) => b.score - a.score);
  if (found.length > 1) log(`[참고] ${label} 표가 있는 시트가 ${found.length}개: ${found.map((f) => f.sheet).join(", ")} → '${found[0].sheet}' 사용${prefer.length ? " (Remain 열이 있는 시트 우선)" : prefixAny.length ? " (ETD 열이 가장 많은 시트)" : ""}. 다른 시트를 쓰려면 파일 아래 시트 선택에서 고르세요.`, "warn");
  return found[0];
}
// 파일을 고르면 시트 목록을 보여 주고 고를 수 있게
async function fillSheetSelect(input, selId, mustHave, anyOf, prefixAny, prefer) {
  const sel = $(selId); sel.innerHTML = '<option value="">(자동 선택)</option>'; sel.style.display = "none"; sel.parentElement.style.display = "none";
  const f = input.files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    const cands = [];
    for (const sn of wb.SheetNames) { const rows = sheetRows(wb, sn); const h = findHeaderRow(rows, mustHave, anyOf, prefixAny); if (h >= 0) { const hn = rows[h].map(norm); cands.push({ sn, score: sheetScore(hn, prefer, prefixAny) }); } }
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
  const cSum = col(sh, ["SUM", "합계"], { required: false });
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
    row._sum = cSum ? toNum(r[cSum]) : etdAll.reduce((a, c) => a + toNum(r[c]), 0);   // 미납수량 합계(긴급 자재 메일용)
    row._vendor = resolveVendor(cV ? clean(r[cV]) : "", cd, maker);
    rows.push(row);
  }
  return { rows, etdSel, etdAll, hasSum: !!cSum };
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
