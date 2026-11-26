/* 연호 발주 화면 — 파일 읽기 · 표 · 엑셀 받기 (계산은 yeonho_calc.js) */
"use strict";
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const fmt = (n) => (n ? Number(n).toLocaleString("en-US") : "");
const nrm = (s) => String(s == null ? "" : s).replace(/\s+/g, "").toLowerCase();
const store = { get(k, d) { try { const v = localStorage.getItem("yh." + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } }, set(k, v) { try { localStorage.setItem("yh." + k, JSON.stringify(v)); } catch (e) { /* 저장 못 해도 동작 */ } } };

const S = { before: null, list: null, stock: null, after: null, master: YH_MASTER, terminal: YH_TERMINAL, masterNote: `기본 마스터 (${YH_MASTER_DATE}, ${YH_MASTER.length}줄)`, result: null, edits: new Map(), filter: "all", afterDate: "" };

// ---------------------------------------------------------------- 파일 → 표(행 배열) 목록
async function readTables(file) {
  const buf = new Uint8Array(await file.arrayBuffer());
  const head = new TextDecoder("latin1").decode(buf.slice(0, 2000)).trimStart();
  if (head.startsWith("<")) {   // 연호 KMS 에서 받은 .xls 는 사실 HTML 표
    const cs = (head.match(/charset\s*=\s*["']?([\w-]+)/i) || [])[1] || "euc-kr";
    let text; try { text = new TextDecoder(cs).decode(buf); } catch (e) { text = new TextDecoder("euc-kr").decode(buf); }
    const doc = new DOMParser().parseFromString(text, "text/html");
    return [...doc.querySelectorAll("table")].map((t, i) => ({ name: `표${i + 1}`, rows: [...t.rows].map((tr) => [...tr.cells].map((c) => c.textContent.replace(/ /g, " ").trim())) }));
  }
  const wb = XLSX.read(buf, { type: "array", cellDates: true, cellStyles: true });
  return wb.SheetNames.map((sn) => ({ name: sn, ws: wb.Sheets[sn], rows: XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: "" }) }));
}
// 머리글 줄 찾기: need 의 이름들이 모두 들어 있는 첫 줄(앞 15줄)
function headerAt(rows, need) {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const h = (rows[i] || []).map(nrm);
    if (need.every((alts) => alts.some((a) => h.some((x) => x === nrm(a) || x.startsWith(nrm(a)))))) return i;
  }
  return -1;
}
function colOf(hrow, names) { const h = hrow.map(nrm); for (const n of names) { const i = h.findIndex((x) => x === nrm(n)); if (i >= 0) return i; } for (const n of names) { const i = h.findIndex((x) => x.startsWith(nrm(n))); if (i >= 0) return i; } return -1; }
const isYeonho = (maker, name) => /YEONHO/i.test(maker) || /YEONHO/i.test(name);

/** 연호 미납조회 → [{name, qty(잔량), ord(수주수량), memo(적요), date(수주일자)}] */
function parseErp(tables) {
  for (const t of tables) {
    const h = headerAt(t.rows, [["품명"], ["잔량", "발주수량", "수주수량"]]); if (h < 0) continue;
    const hr = t.rows[h], cN = colOf(hr, ["품명"]), cR = colOf(hr, ["잔량"]), cO = colOf(hr, ["수주수량", "발주수량"]), cM = colOf(hr, ["적요", "Fac", "공장"]), cD = colOf(hr, ["수주일자"]);
    const memoCol = cM >= 0 ? cM : 2;   // 작업양식 '연호 ERP' 시트처럼 C열에 F1·F2·F5만 적은 경우
    const out = [];
    for (const r of t.rows.slice(h + 1)) {
      const name = String(r[cN] == null ? "" : r[cN]).trim(); if (!name) continue;
      const rem = yhNum(r[cR >= 0 ? cR : cO]), ord = yhNum(r[cO >= 0 ? cO : cR]);
      let d = cD >= 0 ? r[cD] : ""; if (d instanceof Date) d = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
      out.push({ name, qty: rem, ord, memo: String(r[memoCol] == null ? "" : r[memoCol]).trim(), date: String(d || "").trim() });
    }
    if (out.length) return out;
  }
  throw new Error("연호 미납조회 표(품명·잔량 머리글)를 찾지 못했습니다.");
}

/** 주간 선박 및 에어리스트(또는 작업양식) → {ship, air, etd, stock?, master?} */
function parseList(tables, weeks) {
  let ship = null, air = null, stock = null, master = null, terminal = null, etdNames = [], shipSheet = "", airSheet = "";
  // 선박: ITEM CD + ETD 열이 있는 시트 중 ETD 열이 가장 많은 것
  let best = 0;
  for (const t of tables) {
    const h = headerAt(t.rows, [["ITEM CD", "Item cd"], ["ITEM NM", "Part Name"]]); if (h < 0) continue;
    const hr = t.rows[h].map((x) => String(x == null ? "" : x).trim()), etd = hr.map((x, i) => [x, i]).filter(([x]) => /^ETD/i.test(x));
    if (etd.length > best) { best = etd.length; ship = { t, h, hr, etd }; }
  }
  if (ship) {
    const { t, h, hr, etd } = ship, cF = colOf(hr, ["Fac", "FAC"]), cC = colOf(hr, ["ITEM CD"]), cMk = colOf(hr, ["MAKER"]), cN = colOf(hr, ["ITEM NM", "Part Name"]), cB = colOf(hr, ["BALANCE"]);
    const use = etd.slice(0, weeks); etdNames = use.map(([x]) => x.replace(/^ETD\s*/i, "")); shipSheet = t.name;
    ship = t.rows.slice(h + 1).filter((r) => String(r[cN] || "").trim() && isYeonho(cMk >= 0 ? r[cMk] : "", r[cN])).map((r) => ({
      fac: r[cF], code: String(r[cC] || "").trim(), name: String(r[cN]).trim(), qty: use.reduce((a, [, i]) => a + yhNum(r[i]), 0), wait: 0, bal: cB >= 0 ? r[cB] : 0 }));
  }
  // 에어: Purchase request 열이 있는 시트 우선, 없으면 Remain
  let airT = null;
  for (const key of ["Purchase request", "Remain"]) {
    for (const t of tables) { const h = headerAt(t.rows, [["Part Name", "ITEM NM"], [key]]); if (h >= 0) { airT = { t, h, key }; break; } }
    if (airT) break;
  }
  if (airT) {
    const { t, h, key } = airT, hr = t.rows[h], cF = colOf(hr, ["FAC", "Fac"]), cC = colOf(hr, ["Item cd", "ITEM CD"]), cMk = colOf(hr, ["Maker", "MAKER"]), cN = colOf(hr, ["Part Name", "ITEM NM"]), cQ = colOf(hr, [key]), cB = colOf(hr, ["BALANCE"]);
    airSheet = `${t.name} (${key})`;
    air = t.rows.slice(h + 1).filter((r) => String(r[cN] || "").trim() && isYeonho(cMk >= 0 ? r[cMk] : "", r[cN])).map((r) => ({ fac: r[cF], code: String(r[cC] || "").trim(), name: String(r[cN]).trim(), qty: r[cQ], bal: cB >= 0 ? r[cB] : 0 }));
  }
  for (const t of tables) {
    if (!stock && (/재고/.test(t.name) || headerAt(t.rows, [["Shipped"]]) >= 0)) { try { stock = parseStock([t]); } catch (e) { /* 재고 표 아님 */ } }
    if (!master && /마스터/.test(t.name)) master = readMaster(t);
    if (!terminal && /터미널/.test(t.name)) terminal = readTerminal(t);
  }
  return { ship: ship || [], air: air || [], etdNames, shipSheet, airSheet, stock, master, terminal };
}
function parseStock(tables) {
  for (const t of tables) {
    const h = headerAt(t.rows, [["FAC", "Fac", "공장"], ["Part Name", "ITEM NM", "품목명", "품명"]]); if (h < 0) continue;
    const hr = t.rows[h], cF = colOf(hr, ["FAC", "Fac", "공장"]), cC = colOf(hr, ["ITEM CD", "Item cd", "거래처 품목코드", "품목코드"]), cN = colOf(hr, ["Part Name", "ITEM NM", "품목명", "품명"]);
    let cQ = colOf(hr, ["Shipped", "재고수량", "재고", "STOCK", "수량", "Q'ty"]); if (cQ < 0) cQ = 3;
    const out = t.rows.slice(h + 1).filter((r) => String(r[cN] || "").trim()).map((r) => ({ fac: r[cF], code: cC >= 0 ? r[cC] : "", name: String(r[cN]).trim(), qty: r[cQ] }));
    if (out.length) return out;
  }
  throw new Error("재고 표(FAC·품목명·수량)를 찾지 못했습니다.");
}
// 마스터 시트: A 연호명, B 대아명, B칸에 색이 칠해져 있으면 터미널
function readMaster(t) {
  const out = [];
  for (let i = 1; i < t.rows.length; i++) {
    const [a, b] = t.rows[i]; if (!String(b || "").trim()) continue;
    const c = t.ws && t.ws[XLSX.utils.encode_cell({ r: i, c: 1 })], s = c && c.s, fill = s && (s.fgColor || (s.fill && s.fill.fgColor));
    const painted = !!(s && (s.patternType && s.patternType !== "none" || fill && fill.rgb && !/^(FF)?FFFFFF$/i.test(fill.rgb)));
    out.push([String(a || "").trim(), String(b).trim(), painted ? 1 : 0]);
  }
  return out.length ? out : null;
}
function readTerminal(t) {
  const h = headerAt(t.rows, [["CODE"], ["BOX"]]); if (h < 0) return null;
  const hr = t.rows[h], cC = colOf(hr, ["CODE"]), cN = colOf(hr, ["품목명"]), cB = colOf(hr, ["BOX"]);
  const out = t.rows.slice(h + 1).filter((r) => r[cC] || r[cN]).map((r) => [String(r[cC] || "").trim(), String(r[cN] || "").trim(), yhNum(r[cB])]);
  return out.length ? out : null;
}

// ---------------------------------------------------------------- 파일 넣기
const BOXES = {
  before: { label: "연호 미납조회 (발주 전)", read: async (f) => { const rows = parseErp(await readTables(f)); return { rows, note: `${rows.length}줄` }; } },
  list: { label: "선박 및 에어리스트", read: async (f) => { const tabs = await readTables(f); const p = parseList(tabs, weeks()); p.tabs = tabs; if (!p.ship.length && !p.air.length) throw new Error("YEONHO 줄이 있는 선박·에어 표를 찾지 못했습니다."); return p; } },
  stock: { label: "재고 (선택)", read: async (f) => { const rows = parseStock(await readTables(f)); return { rows, note: `${rows.length}줄` }; } },
  after: { label: "연호 미납조회 (발주 후)", read: async (f) => { const rows = parseErp(await readTables(f)); return { rows, note: `${rows.length}줄` }; } },
};
function weeks() { return Math.max(1, +($("weeks").value || 3)); }
function listNote(p) {
  const bits = [`선박 ${p.ship.length}줄 (${p.shipSheet || "-"}, ETD ${p.etdNames.join(" · ") || "없음"})`, `에어 ${p.air.length}줄 (${p.airSheet || "-"})`];
  if (p.stock) bits.push(`재고 시트 ${p.stock.length}줄`);
  if (p.master) bits.push(`마스터 ${p.master.length}줄`);
  return bits.join(" / ");
}
async function takeFile(kind, file) {
  const box = document.querySelector(`.drop[data-k="${kind}"]`), st = box.querySelector(".st");
  try {
    const v = await BOXES[kind].read(file); v.file = file.name; S[kind] = v;
    if (kind === "list") {
      if (v.master) { S.master = v.master; S.masterNote = `${file.name} 의 마스터 (${v.master.length}줄, 터미널 ${v.master.filter((m) => m[2]).length})`; }
      if (v.terminal) S.terminal = v.terminal;
      const sb = document.querySelector('.drop[data-k="stock"]');
      if (v.stock && !S.stock) { sb.classList.add("ok"); sb.querySelector(".st").innerHTML = `②의 '재고' 시트 사용 중 (${v.stock.length}줄)<br>다른 재고 파일을 넣으면 그것을 씁니다.`; }
    }
    box.classList.add("ok"); box.classList.remove("err");
    st.innerHTML = `<b>${esc(file.name)}</b><br>${esc(kind === "list" ? listNote(v) : v.note)}`;
  } catch (e) {
    S[kind] = null; box.classList.add("err"); box.classList.remove("ok");
    st.innerHTML = `<b>${esc(file.name)}</b><br>읽지 못했습니다: ${esc(e.message)}`;
  }
  $("mnote").textContent = S.masterNote;
  if (kind === "after") renderAfterDates();
  compute();
}
function wireDrops() {
  for (const box of document.querySelectorAll(".drop")) {
    const k = box.dataset.k, inp = box.querySelector("input");
    inp.addEventListener("change", () => { if (inp.files[0]) takeFile(k, inp.files[0]); inp.value = ""; });
    box.addEventListener("dragover", (e) => { e.preventDefault(); box.classList.add("over"); });
    box.addEventListener("dragleave", () => box.classList.remove("over"));
    box.addEventListener("drop", (e) => { e.preventDefault(); box.classList.remove("over"); if (e.dataTransfer.files[0]) takeFile(k, e.dataTransfer.files[0]); });
  }
}

// ---------------------------------------------------------------- 발주 계산
function stockRows() { return S.stock ? S.stock.rows : S.list && S.list.stock ? S.list.stock : null; }
function compute() {
  if (!S.list) { S.result = null; renderOrder(); renderCheck(); return; }
  const r = yhOrder({ erp: S.before ? S.before.rows : [], ship: S.list.ship, air: S.list.air, stock: stockRows(), master: S.master, terminal: S.terminal }, { urgentPool: $("upool").checked });
  for (const x of r.rows) x.boxUp = x.terminal && x.box && x.need > 0 ? Math.ceil(x.need / x.box) * x.box : 0;
  S.result = r;
  renderOrder(); renderCheck();
}
const finalQty = (x) => (S.edits.has(x.key) ? S.edits.get(x.key) : x.need);
function renderOrder() {
  const el = $("order");
  if (!S.result) { el.innerHTML = `<p class="muted">② 선박 및 에어리스트를 넣으면 여기 최종발주 표가 나옵니다.</p>`; $("osum").innerHTML = ""; return; }
  const rows = S.result.rows, warn = [];
  if (!S.before) warn.push("① 발주 전 미납조회가 없어 <b>연호대기를 0</b>으로 계산했습니다.");
  if (!stockRows()) warn.push("재고 파일이 없어 <b>본사대기를 0</b>으로 계산했습니다.");
  const unreg = rows.filter((x) => x.status === "미등록");
  if (unreg.length) warn.push(`마스터에 없는 품명 ${unreg.length}개(표에서 <span class="tag un">미등록</span>) — 연호 품명과 맞춰지지 않아 연호대기가 빠질 수 있습니다.`);
  if (S.result.urgentLeft.length) warn.push(`쓰고 남은 긴급 재고: ${S.result.urgentLeft.map(([k, q]) => `${esc(k)} ${fmt(q)}`).join(", ")}`);
  const need = rows.filter((x) => finalQty(x) > 0), edited = rows.filter((x) => S.edits.has(x.key)).length;
  $("osum").innerHTML = `<span class="chip">전체 <b>${rows.length}</b></span><span class="chip pri">발주할 품목 <b>${need.length}</b></span><span class="chip">발주 합계 <b>${fmt(need.reduce((a, x) => a + finalQty(x), 0))}</b></span><span class="chip or">터미널 <b>${rows.filter((x) => x.terminal).length}</b></span><span class="chip red">에어 부족 <b>${rows.filter((x) => x.airShort).length}</b></span>${edited ? `<span class="chip">손으로 고친 칸 <b>${edited}</b></span>` : ""}`
    + (warn.length ? `<div class="box warn">${warn.join("<br>")}</div>` : "");
  const show = rows.filter((x) => S.filter === "all" || (S.filter === "need" && finalQty(x) > 0) || (S.filter === "term" && x.terminal) || (S.filter === "edit" && (S.edits.has(x.key) || x.boxUp && x.boxUp !== x.need)));
  el.innerHTML = `<div class="tw"><table class="grid"><thead><tr><th>FAC</th><th>품명</th><th>선박</th><th>AIR</th><th>필요합</th><th>본사대기</th><th>연호대기</th><th class="hq">발주필요량</th><th>발란스</th><th>발란스 남음</th><th>박스</th></tr></thead><tbody>${show.map((x) => {
    const q = finalQty(x), left = (x.bal || 0) - q - x.yh - x.wait, ed = S.edits.has(x.key);
    return `<tr class="${q > 0 ? "" : "zero"}"><td class="c">${esc(x.fac)}</td><td class="nm">${esc(x.name)}${x.status === "미등록" ? ` <span class="tag un">미등록</span>` : ""}${x.urgent ? ` <span class="tag ug">긴급 ${fmt(x.urgent)} 차감</span>` : ""}</td>`
      + `<td>${fmt(x.ship)}</td><td>${fmt(x.air)}</td><td>${fmt(x.total)}</td><td>${fmt(x.wait)}</td><td>${fmt(x.yh)}</td>`
      + `<td class="hq ${x.terminal ? "term" : ""}"><input data-k="${esc(x.key)}" value="${q ? q.toLocaleString("en-US") : "0"}" class="${x.airShort ? "red" : ""} ${ed ? "ed" : ""}" inputmode="numeric" title="${ed ? `계산값 ${fmt(x.need) || 0} → 고침` : "칸을 눌러 고칠 수 있어요"}"></td>`
      + `<td>${fmt(x.bal)}</td><td class="${left < 0 ? "neg" : ""}">${left ? left.toLocaleString("en-US") : ""}</td>`
      + `<td>${x.box ? `${fmt(x.box)}${x.boxUp && x.boxUp !== q ? ` <button class="mini" data-up="${esc(x.key)}" title="${fmt(x.boxUp)}로 올리기">→ ${fmt(x.boxUp)}</button>` : ""}` : ""}</td></tr>`;
  }).join("")}</tbody></table></div>`;
  el.querySelectorAll("input[data-k]").forEach((inp) => inp.addEventListener("change", () => {
    const x = rows.find((r) => r.key === inp.dataset.k), v = yhNum(inp.value);
    if (v === x.need) S.edits.delete(x.key); else S.edits.set(x.key, v);
    renderOrder(); renderCheck();
  }));
  el.querySelectorAll("button[data-up]").forEach((b) => b.addEventListener("click", () => { const x = rows.find((r) => r.key === b.dataset.up); S.edits.set(x.key, x.boxUp); renderOrder(); renderCheck(); }));
}

// ---------------------------------------------------------------- 점검
function renderAfterDates() {
  const sel = $("adate"); if (!S.after) { sel.innerHTML = ""; return; }
  const cnt = new Map(); for (const r of S.after.rows) cnt.set(r.date || "(날짜 없음)", (cnt.get(r.date || "(날짜 없음)") || 0) + 1);
  const ds = [...cnt.keys()].sort().reverse();
  S.afterDate = ds[0] || "";
  sel.innerHTML = `<option value="*">전부 (${S.after.rows.length}줄)</option>` + ds.map((d) => `<option value="${esc(d)}" ${d === S.afterDate ? "selected" : ""}>수주일자 ${esc(d)} (${cnt.get(d)}줄)</option>`).join("");
}
function daeaRows() { return S.result ? S.result.rows.filter((x) => finalQty(x) > 0).map((x) => ({ fac: x.fac, name: x.name, qty: finalQty(x) })) : []; }
function yhRows() {
  if (!S.after) return [];
  return S.after.rows.filter((r) => S.afterDate === "*" || (r.date || "(날짜 없음)") === S.afterDate).map((r) => { const m = String(r.memo).toUpperCase().match(/F\d/); return { fac: m ? m[0] : yhFac(r.memo), name: r.name, qty: r.ord || r.qty }; });
}
function renderCheck() {
  const el = $("check");
  if (!S.after || !S.result) { el.innerHTML = `<p class="muted">${!S.result ? "② 리스트로 최종발주를 먼저 만들고, " : ""}③ 연호 미납조회(발주 후)를 넣으면 대조 결과가 나옵니다.</p>`; return; }
  const res = yhCheck(yhRows(), daeaRows(), S.master); S.check = res;
  const bad = res.filter((x) => x.bad), sum = (a, f) => a.reduce((s, x) => s + x[f], 0);
  el.innerHTML = `<div class="box ${bad.length ? "stop" : "good"}"><b>${bad.length ? `불일치 ${bad.length}개 — 빨간 줄을 확인하세요` : "모두 일치합니다 ✔"}</b>연호 ${res.filter((x) => x.yh).length}품목 ${fmt(sum(res, "yh"))} / 대아 ${res.filter((x) => x.daea).length}품목 ${fmt(sum(res, "daea"))}</div>`
    + `<label class="opt"><input type="checkbox" id="onlybad" ${store.get("onlybad", false) ? "checked" : ""}> 불일치만 보기</label>`
    + `<div class="tw"><table class="grid"><thead><tr><th>FAC</th><th>품명(연호기준)</th><th>연호 수량</th><th>대아 수량</th><th>차이(연호−대아)</th><th>상태</th></tr></thead><tbody>${res.filter((x) => !store.get("onlybad", false) || x.bad).map((x) =>
      `<tr class="${x.bad ? "bad" : ""}"><td class="c">${esc(x.fac)}</td><td class="nm">${esc(x.part)}<small>${x.name !== x.part ? esc(x.name) : ""}</small></td><td class="${x.terminal ? "term" : ""}">${fmt(x.yh) || 0}</td><td class="${x.terminal ? "term" : ""}">${fmt(x.daea) || 0}</td><td class="${x.terminal ? "term" : ""}">${x.bad ? x.diff.toLocaleString("en-US") : ""}</td><td>${x.bad ? (x.onlyY ? "불일치 · 대아에 없음" : x.onlyD ? "불일치 · 연호에 없음" : "불일치") : ""}</td></tr>`).join("")}</tbody></table></div>`;
  $("onlybad").addEventListener("change", (e) => { store.set("onlybad", e.target.checked); renderCheck(); });
}

// ---------------------------------------------------------------- 엑셀 받기 · 복사
const XS = {
  head: { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "2B3746" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true } },
  headH: { font: { bold: true, color: { rgb: "000000" }, sz: 12 }, fill: { fgColor: { rgb: "FFD54F" } }, alignment: { horizontal: "center", vertical: "center" } },
  num: { numFmt: "#,##0" },
};
const border = { top: { style: "thin", color: { rgb: "D2D7DE" } }, bottom: { style: "thin", color: { rgb: "D2D7DE" } }, left: { style: "thin", color: { rgb: "D2D7DE" } }, right: { style: "thin", color: { rgb: "D2D7DE" } } };
function sheetOf(aoa, styleFn, widths) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let r = 0; r < aoa.length; r++) for (let c = 0; c < aoa[r].length; c++) {
    const a = XLSX.utils.encode_cell({ r, c }); if (!ws[a]) ws[a] = { t: "s", v: "" };
    ws[a].s = Object.assign({ border, font: { name: "맑은 고딕", sz: 11 }, alignment: { vertical: "center" } }, typeof aoa[r][c] === "number" ? XS.num : {}, styleFn(r, c) || {});
  }
  ws["!cols"] = widths.map((w) => ({ wch: w }));
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
  ws["!freeze"] = { xSplit: 2, ySplit: 1 };
  return ws;
}
function stamp() { const d = new Date(); return `${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`; }
function downloadOrder() {
  if (!S.result) return alert("먼저 ② 리스트를 넣어 주세요.");
  const rows = S.result.rows, wb = XLSX.utils.book_new();
  const aoa = [["FAC", "품명", "선박", "AIR", "필요합", "본사대기", "연호대기", "발주필요량", "발란스", "발란스 남음", "REMARK(박스)"]];
  for (const x of rows) { const q = finalQty(x); aoa.push([x.fac, x.name, x.ship, x.air, x.total, x.wait, x.yh, q, x.bal || "", (x.bal || 0) - q - x.yh - x.wait, x.box || ""]); }
  XLSX.utils.book_append_sheet(wb, sheetOf(aoa, (r, c) => {
    if (r === 0) return c === 7 ? XS.headH : XS.head;
    const x = rows[r - 1], s = { fill: { fgColor: { rgb: r % 2 ? "FFFFFF" : "EFF3F8" } } };
    if (c === 0) return Object.assign(s, { font: { bold: true }, alignment: { horizontal: "center", vertical: "center" } });
    if (c === 1) return Object.assign(s, { font: { bold: true, sz: 12 } });
    if (c === 7) return { numFmt: "#,##0", fill: { fgColor: { rgb: x.terminal ? "FFA500" : "FFF8CC" } }, font: { bold: true, sz: 16, color: { rgb: x.airShort ? "FF0000" : "000000" } } };
    return s;
  }, [7, 38, 11, 11, 11, 11, 11, 15, 11, 12, 12]), "최종발주");
  const d = [["FAC", "품명", "발주필요량"], ...daeaRows().map((x) => [x.fac, x.name, x.qty])];
  XLSX.utils.book_append_sheet(wb, sheetOf(d, (r) => (r === 0 ? XS.head : null), [7, 38, 14]), "대아 ERP");
  XLSX.writeFile(wb, `연호 최종발주_${stamp()}.xlsx`);
}
function downloadCheck() {
  if (!S.check) return alert("③ 발주 후 미납조회를 넣어 주세요.");
  const aoa = [["Fac", "품명(연호기준)", "연호 수량", "대아 수량", "차이(연호-대아)", "상태"], ...S.check.map((x) => [x.fac, x.part, x.yh, x.daea, x.bad ? x.diff : "", x.bad ? "불일치" : ""])];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheetOf(aoa, (r, c) => {
    if (r === 0) return XS.head; const x = S.check[r - 1], s = {};
    if (x.bad) s.font = { color: { rgb: "C80000" }, bold: true };
    if (x.terminal && c >= 2 && c <= 4) s.fill = { fgColor: { rgb: "FFC000" } };
    return s;
  }, [6, 30, 12, 12, 16, 10]), "최종수량대조_결과");
  XLSX.writeFile(wb, `연호 입고점검_${stamp()}.xlsx`);
}
async function copyOrder() {
  const d = daeaRows(); if (!d.length) return alert("발주할 품목이 없습니다.");
  const tsv = ["FAC\t품명\t발주필요량", ...d.map((x) => `${x.fac}\t${x.name}\t${x.qty}`)].join("\n");
  try { await navigator.clipboard.writeText(tsv); flash("copyo", "복사했습니다 ✔"); } catch (e) { prompt("아래 내용을 Ctrl+C 로 복사하세요", tsv); }
}
function flash(id, t) { const b = $(id), o = b.textContent; b.textContent = t; setTimeout(() => (b.textContent = o), 1500); }

// ---------------------------------------------------------------- 시작
document.addEventListener("DOMContentLoaded", () => {
  $("weeks").value = store.get("weeks", 3);
  $("upool").checked = store.get("upool", false);
  $("mnote").textContent = S.masterNote;
  wireDrops();
  $("weeks").addEventListener("change", () => { store.set("weeks", weeks()); if (S.list && S.list.tabs) { const p = parseList(S.list.tabs, weeks()); Object.assign(S.list, { ship: p.ship, etdNames: p.etdNames }); document.querySelector('.drop[data-k="list"] .st').innerHTML = `<b>${esc(S.list.file)}</b><br>${esc(listNote(S.list))}`; } compute(); });
  $("upool").addEventListener("change", () => { store.set("upool", $("upool").checked); compute(); });
  $("adate").addEventListener("change", (e) => { S.afterDate = e.target.value; renderCheck(); });
  document.querySelectorAll("[data-f]").forEach((b) => b.addEventListener("click", () => { S.filter = b.dataset.f; document.querySelectorAll("[data-f]").forEach((x) => x.classList.toggle("on", x === b)); renderOrder(); }));
  $("boxall").addEventListener("click", () => { if (!S.result) return; let n = 0; for (const x of S.result.rows) if (x.boxUp && x.boxUp !== finalQty(x)) { S.edits.set(x.key, x.boxUp); n++; } renderOrder(); renderCheck(); flash("boxall", n ? `${n}칸 올렸습니다` : "올릴 칸 없음"); });
  $("reset").addEventListener("click", () => { S.edits.clear(); renderOrder(); renderCheck(); });
  $("dlo").addEventListener("click", downloadOrder);
  $("dlc").addEventListener("click", downloadCheck);
  $("copyo").addEventListener("click", copyOrder);
  renderOrder(); renderCheck();
});
