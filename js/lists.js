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
// avoid: 이 열이 있는 시트는 뒤로 (월요일 리스트에서는 Remain/출고수량 시트보다 원래 Air 시트를 쓴다)
const sheetScore = (hn, prefer, prefixAny, avoid = []) => prefer.filter((p) => hn.includes(norm(p))).length * 100 - avoid.filter((p) => hn.includes(norm(p))).length * 100 + hn.filter((c) => prefixAny.some((p) => c.startsWith(norm(p)))).length;
function loadSheet(wb, mustHave, label, anyOf = [], prefixAny = [], prefer = [], sheetName = "", avoid = []) {
  const found = [];
  for (const sn of wb.SheetNames) {
    if (sheetName && norm(sn) !== norm(sheetName)) continue;
    if (!sheetName && /장납/.test(sn)) continue;   // "에어 장납" 시트는 품목 목록일 뿐 에어 표가 아님
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
    found.push({ sheet: sn, headers: headers.filter(Boolean), rows: data, score: sheetScore(hn, prefer, prefixAny, avoid) + data.length / 100000 });   // 점수가 같으면 줄이 많은(전체) 시트
  }
  if (!found.length) throw new Error(`${label} 파일에서 ${label} 표(헤더 ${mustHave.join(", ")}${anyOf.length ? " + " + anyOf.join("/") : ""}${prefixAny.length ? " + ETD열" : ""})가 있는 시트를 못 찾았습니다. 시트: ${wb.SheetNames.join(", ")}${sheetName ? " / 지정한 시트: " + sheetName : ""}. 선박리스트와 에어리스트에 같은 파일을 고른 것은 아닌지 확인하세요.`);
  found.sort((a, b) => b.score - a.score);
  if (found.length > 1) log(`[참고] ${label} 표가 있는 시트가 ${found.length}개: ${found.map((f) => f.sheet).join(", ")} → '${found[0].sheet}' 사용${prefer.length ? ` (${prefer.join("/")} 열이 있는 시트 우선)` : avoid.length ? ` (${avoid.join("/")} 열이 없는 원래 시트 우선)` : prefixAny.length ? " (ETD 열이 가장 많은 시트)" : ""}. 다른 시트를 쓰려면 파일 아래 시트 선택에서 고르세요.`, "warn");
  return found[0];
}
// 파일을 고르면 시트 목록을 보여 주고 고를 수 있게
async function fillSheetSelect(input, selId, mustHave, anyOf, prefixAny, prefer, avoid = []) {
  const sel = $(selId); sel.innerHTML = '<option value="">(자동 선택)</option>'; sel.style.display = "none"; sel.parentElement.style.display = "none";
  const f = input.files[0]; if (!f) return;
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    const cands = [];
    for (const sn of wb.SheetNames) { const rows = sheetRows(wb, sn); const h = findHeaderRow(rows, mustHave, anyOf, prefixAny); if (h >= 0) { const hn = rows[h].map(norm); cands.push({ sn, score: sheetScore(hn, prefer, prefixAny, avoid) }); } }
    if (cands.length > 1) {
      cands.sort((a, b) => b.score - a.score);
      const auto = selId === "airSheet" ? "자동: 월요일 리스트는 원래 Air 시트, 수요일 REMAIN 리스트는 출고수량·Remain 열이 있는 시트" : "자동: " + cands[0].sn;
      sel.innerHTML = `<option value="">(${esc(auto)})</option>` + wb.SheetNames.map((sn) => `<option value="${esc(sn)}">${esc(sn)}${cands.some((c) => c.sn === sn) ? "" : " (표 없음)"}</option>`).join("");
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
    row._vendor = resolveVendor(cV ? clean(r[cV]) : typeof vendorFromTable === "function" ? vendorFromTable(cd, r[cNm], maker) : "", cd, maker);
    rows.push(row);
  }
  if (cV) rememberItemVendors(sh.rows.filter((r) => clean(r[cCd]) && clean(r[cV])).map((r) => [ivKey(r[cFac], r[cCd]), clean(r[cV])]));
  return { rows, etdSel, etdAll, hasSum: !!cSum };
}
// ---------------------------------------------------------------- 에어리스트
// 두 가지 에어가 있다.
//  · 화요일 에어(월요일에 받는 주간 리스트): Purchase request 수량. KET 품목·"에어 장납" 시트 품목은 NEED ETD와 상관없이 전부,
//    나머지 업체는 NEED ETD가 URGENT인 품목만 요청한다.
//  · 금요일 에어(수요일에 받는 REMAIN 리스트): 시트에 출고수량 열이 있으면 Remain = Purchase request − 출고수량,
//    Remain 열이 있으면 그 값. 0보다 큰 품목 전부.
const AIR_REMAIN_COLS = ["Remain", "출고수량"];
const isUrgent = (v) => /urgent/i.test(clean(v));
// "에어 장납" 시트(이름에 '장납'이 들어간 시트)의 품목 코드. FAC 열이 있으면 FAC+품목, 없으면 품목만.
// 이번 파일에 없으면 이 PC에 기억해 둔 지난 목록을 쓴다.
const LS_LONGLEAD = "mailmaker_air_longlead";
function longLeadItems(wbs) {
  const items = [], sheets = [];
  for (const wb of wbs) for (const sn of wb.SheetNames) {
    if (!/장납/.test(sn)) continue;
    const rows = sheetRows(wb, sn); sheets.push(sn);
    const h = findHeaderRow(rows, ["Item cd"]), fi = h >= 0 ? rows[h].map(norm).findIndex((c) => c === "fac") : -1;
    for (let i = h + 1; i < rows.length; i++) for (const c of rows[i] || []) {
      const cd = clean(c); if (!/^[A-Z]\d{3}-\d{4,7}$/i.test(cd)) continue;
      items.push({ fac: fi >= 0 ? clean((rows[i] || [])[fi]).toUpperCase() : "", cd: norm(cd) });
    }
  }
  if (sheets.length) { store.set(LS_LONGLEAD, { at: new Date().toISOString(), sheets, items }); return { items, from: `시트 '${sheets.join(", ")}'` }; }
  const saved = store.get(LS_LONGLEAD, null);
  return saved ? { items: saved.items, from: `이 PC에 기억해 둔 목록(${koStamp(saved.at)}에 읽은 '${saved.sheets.join(", ")}')`, old: true } : { items: [], from: "" };
}
// 품목 → 업체 기억: 업체 열(거래처명(매칭))이 있는 엑셀을 읽을 때마다 FAC+품목코드별 업체 이름을 이 PC에 기억한다.
// 메일 본문의 표를 붙여 넣으면 업체 열이 없으므로, 이 기억을 MAKER보다 먼저 쓴다.
const LS_ITEMVEND = "mailmaker_item_vendor";
const ivKey = (fac, cd) => clean(fac).toUpperCase() + "|" + norm(cd);
function rememberItemVendors(pairs) {
  if (!pairs.length) return;
  const m = store.get(LS_ITEMVEND, {});
  for (const [k, v] of pairs) { m[k] = v; m["*|" + k.split("|")[1]] = v; }
  store.set(LS_ITEMVEND, m);
}
const airSortKey = (s) => clean(s).toUpperCase().replace(/[-']/g, "");
const airCmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
const airRowSort = (a, b) => airCmp(a.FAC, b.FAC) || airCmp(airSortKey(a.Maker), airSortKey(b.Maker)) || airCmp(airSortKey(a["Part Name"]), airSortKey(b["Part Name"]));

/** friday=true: 금요일 에어(REMAIN)용 시트를 우선. false: 월요일 주간 리스트의 원래 Air 시트를 우선.
    돌려주는 값: 에어 행 배열 + .mode ("remain" | "tue") */
function parseAir(wb, { friday = false, extraWbs = [] } = {}) {
  const sh = loadSheet(wb, ["Item cd", "Maker"], "에어리스트", ["Remain", "Purchase request"], [], friday ? AIR_REMAIN_COLS : [], $("airSheet").value, friday ? [] : AIR_REMAIN_COLS);
  const cFac = col(sh, ["FAC", "Fac"]), cCd = col(sh, ["Item cd", "ITEM CD"]), cMk = col(sh, ["Maker", "MAKER"]), cNm = col(sh, ["Part Name", "ITEM NM", "품명"]);
  const cRemain = col(sh, ["Remain"], { required: false }), cPr = col(sh, ["Purchase request"], { required: false }), cOut = col(sh, ["출고수량", "출고"], { required: false });
  const cShort = col(sh, ["Short Del.", "Short Del"], { required: false }), cNeed = col(sh, ["NEED ETD", "Need ETD"], { required: false });
  const cV = col(sh, VENDOR_COLS, { required: false, prefix: true });
  const mode = cRemain || (cOut && cPr) ? "remain" : "tue";
  const qty = cRemain ? (r) => toNum(r[cRemain]) : cOut && cPr ? (r) => toNum(r[cPr]) - toNum(r[cOut]) : (r) => toNum(r[cPr]);
  const qtyLabel = mode === "remain" ? "Remain" : "Purchase request";
  log(`에어리스트 시트: ${sh.sheet} / ` + (cRemain ? `수량: '${cRemain}' 열` : cOut && cPr ? `수량: Remain = '${cPr}' − '${cOut}'` : `수량 열: '${cPr}'`) + (cV ? ` / 업체 열: '${cV}' 사용` : ""));
  let rows = [], fromMem = 0, fromTbl = 0;
  const mem = cV ? null : store.get(LS_ITEMVEND, {}), seen = [];
  for (const r of sh.rows) {
    const cd = clean(r[cCd]); if (cd && cV && clean(r[cV])) seen.push([ivKey(r[cFac], cd), clean(r[cV])]);
    const q = qty(r); if (!cd || q <= 0) continue;
    const maker = clean(r[cMk]) || makerFromName(r[cNm]);
    const memV = mem && (mem[ivKey(r[cFac], cd)] || mem["*|" + norm(cd)]);
    const tblV = !cV && !memV && typeof vendorFromTable === "function" ? vendorFromTable(cd, r[cNm], maker) : "";
    if (memV) fromMem++; else if (tblV) fromTbl++;
    rows.push({
      FAC: clean(r[cFac]), "Item cd": cd, Maker: maker, "Part Name": clean(r[cNm]), "Purchase request": q,
      "Short Del.": cShort ? clean(r[cShort]) : "", "NEED ETD": cNeed ? clean(r[cNeed]) : "", _qtyCol: qtyLabel,
      _vendor: resolveVendor(cV ? clean(r[cV]) : memV || tblV || "", cd, maker), _src: "에어",
    });
  }
  rememberItemVendors(seen);
  if (!cV && rows.length && !fromMem && !fromTbl) log("[주의] 이 PC는 아직 업체 이름(거래처명)이 있는 리스트 엑셀을 읽은 적이 없어, 모두 MAKER로 업체를 정했습니다. KET 품목이 태양단자로 가는 것처럼 틀릴 수 있습니다. 이번 주 월요일 리스트 엑셀을 한 번 넣고 미리보기를 누른 뒤 다시 하거나, 엑셀 파일로 만드세요.", "err");
  const byMk = rows.length - fromMem - fromTbl;
  if (!cV) log(`[업체 정하기] 업체 열이 없어 ${fromMem}건은 이 PC가 전에 읽은 엑셀의 거래처명, ${fromTbl}건은 거래처 표(내장)${byMk ? `, ${byMk}건은 MAKER` : ""}로 업체를 정했습니다.` + (byMk ? " MAKER로 정한 업체가 맞는지 메일 목록에서 확인하세요." : ""), byMk ? "warn" : "info");
  if (mode === "remain") {
    rows.sort(airRowSort);   // 리스트 아래에 덧붙인 줄도 FAC → MAKER → 품명 순서 자리로
    log(`[금요일 에어] Remain이 있는 품목 ${rows.length}건 (Remain 0 이하는 뺌)`);
  } else if ($("tueRule") ? $("tueRule").checked : true) {
    // 화요일 에어 규칙
    const allVend = splitList(CONFIG.settings["화요일에어전체업체"] ?? "KET").map(matchAlias), allMk = splitList(CONFIG.settings["화요일에어전체업체"] ?? "KET").map((x) => x.toUpperCase());
    const ll = longLeadItems([wb, ...extraWbs]);
    const inLL = (r) => ll.items.some((x) => x.cd === norm(r["Item cd"]) && (!x.fac || x.fac === r.FAC.toUpperCase()));
    const before = rows.length, why = { all: 0, ll: 0, urg: 0 };
    rows = rows.filter((r) => {
      if (allVend.includes(r._vendor) || allMk.includes(r.Maker.toUpperCase())) return ++why.all;
      if (inLL(r)) return ++why.ll;
      if (!cNeed || isUrgent(r["NEED ETD"])) return ++why.urg;
      return false;
    });
    log(`[화요일 에어] ${before}건 중 ${rows.length}건 요청: ${allVend.join("·")} 전체 ${why.all}건 + 에어 장납 ${why.ll}건 + NEED ETD URGENT ${why.urg}건 (나머지 ${before - rows.length}건은 URGENT가 아니라 뺌)`);
    if (!cNeed) log("   [확인] NEED ETD 열이 없어 URGENT로 거를 수 없습니다. 전부 넣었습니다.", "warn");
    if (!ll.from) log("   [확인] '에어 장납' 시트를 못 찾았습니다(시트 이름에 '장납'이 들어가야 함). 장납 품목도 URGENT가 아니면 빠졌습니다.", "warn");
    else log(`   에어 장납 목록: ${ll.from}, ${ll.items.length}품목`, ll.old ? "warn" : "info");
  }
  rows.mode = mode;
  return rows;
}

// 메일 본문에서 복사한 표(탭으로 나뉜 글자) → 엑셀과 같은 모양의 통합 문서. 표 머리글(FAC, Item cd …)이 있어야 한다.
function pastedWorkbook(text) {
  const lines = String(text || "").replace(/\r/g, "").split("\n").filter((l) => l.trim());
  if (!lines.length) return null;
  const aoa = lines.map((l) => l.split("\t").map((c) => c.replace(/[\u00a0\u3000]/g, " ").trim()));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "붙여넣은 표");
  return wb;
}
