/* 수주 등록 확인: 태국 주문서 원본(Order No|Seq|Item cd|…|Qty|…|Item nm|Item spec) ↔ 전산 수주현황 엑셀 비교.
   모든 처리는 브라우저 안에서만 한다. */
"use strict";

// 주문번호 앞 2자리(공장) → 수주현황 거래처코드. 여기 없는 공장은 수주현황에서 가장 많이 쓰인 거래처를 기준으로 삼는다.
const CK_FACTORY = { F1: "1010", F2: "1060" };
const CK_OPT_KEY = "ck_opts_v1";

const ckState = { orders: [], status: null, result: null };

const ckS = (v) => (v === null || v === undefined ? "" : String(v)).trim();
const ckNorm = (v) => ckS(v).normalize("NFKC").replace(/\s+/g, " ").toUpperCase();
const ckKeyNorm = (v) => ckNorm(v).replace(/[\s.\-_/()△▽▲▼↑↓]/g, ""); // △▽: 전산 화면의 정렬 표시
const ckEsc = (s) => ckS(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function ckNum(v) {
  if (typeof v === "number") return v;
  const s = ckS(v).replace(/,/g, "");
  if (!s || isNaN(Number(s))) return null;
  return Number(s);
}
/** 엑셀 날짜(일련번호·Date·문자열) → "YYYY-MM-DD" */
function ckDate(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    const d = new Date(Math.round((Math.floor(v) - 25569) * 86400000));
    return d.toISOString().slice(0, 10);
  }
  if (v instanceof Date) return new Date(v.getTime() - v.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  const m = ckS(v).match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}` : ckS(v);
}
/** 글자 비교: 공백·대소문자·전각 차이는 무시. 수주현황의 '?'(전산에서 깨진 글자)는 아무 글자 하나로 본다. */
function ckSameText(orderVal, statusVal) {
  const a = ckNorm(orderVal), b = ckNorm(statusVal);
  if (a === b) return true;
  if (!b.includes("?")) return false;
  const re = new RegExp("^" + b.replace(/[.*+^${}()|[\]\\]/g, "\\$&").replace(/\?/g, ".?") + "$");
  return re.test(a);
}
/* 품명: 태국 주문서와 전산 품목 마스터는 같은 품목도 표기가 조금 다르다
   (예: "SRGW 0.3SQ WHT (BOSUNG)" ↔ "SRGW 0.3SQ(12/0.18) WHT"). 괄호 안(메이커·심선 구성)과 색 약어 차이는 같은 품목으로 본다. */
const CK_COLOR = { WHITE: "WHT", BLACK: "BLK", YELLOW: "YEL", BLUE: "BLU", GREEN: "GRN", ORANGE: "ORG", GRAY: "GRY", GREY: "GRY", BROWN: "BRN", PINK: "PNK", VIOLET: "VIO", PURPLE: "PUR", CLEAR: "CLR", NATURAL: "NAT" };
function ckNameCore(v) {
  return ckNorm(v).replace(/\([^()]*\)/g, " ").replace(/\([^()]*\)/g, " ").split(/[^A-Z0-9.#가-힣?]+/).filter(Boolean)
    .map((t) => CK_COLOR[t] || t).sort().join(" ");
}
/** 0 = 글자까지 같음, 1 = 표기만 다름(같은 품목으로 봄), 2 = 다름 */
function ckNameLevel(orderVal, statusVal) {
  if (ckSameText(orderVal, statusVal)) return 0;
  const a = ckNameCore(orderVal), b = ckNameCore(statusVal);
  return a && a === b ? 1 : 2;
}
const ckSameNum = (a, b) => a !== null && b !== null && Math.abs(a - b) < 1e-6;

/** 시트에서 머리글 줄을 찾아 {열이름(정규화): 열번호}와 데이터 줄을 돌려준다. */
function ckFindTable(wb, needs) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
      const head = (rows[i] || []).map(ckKeyNorm);
      if (needs.every((n) => [].concat(n).some((a) => head.includes(ckKeyNorm(a))))) {
        const col = {};
        head.forEach((h, c) => { if (h && !(h in col)) col[h] = c; });
        return { col, rows: rows.slice(i + 1), headRow: i };
      }
    }
  }
  return null;
}
/** 열 이름은 여러 후보를 줄 수 있다(전산 화면마다 머리글이 조금 다름). 처음 찾은 열의 값을 돌려준다. */
const ckGet = (t, row, ...names) => { for (const n of names) { const c = t.col[ckKeyNorm(n)]; if (c !== undefined) return row[c]; } return null; };

function ckParseOrder(t, fileName) {
  const lines = [];
  for (const r of t.rows) {
    const no = ckS(ckGet(t, r, "Order No")), seq = ckNum(ckGet(t, r, "Seq"));
    if (!no || seq === null) continue;
    lines.push({
      file: fileName, no: no.toUpperCase(), seq,
      itemCd: ckS(ckGet(t, r, "Item cd")), type: ckS(ckGet(t, r, "DO/EEI/EX")),
      date: ckDate(ckGet(t, r, "Delivery date")), unit: ckS(ckGet(t, r, "Unit")),
      qty: ckNum(ckGet(t, r, "Qty")), price: ckNum(ckGet(t, r, "Price")),
      name: ckS(ckGet(t, r, "Item nm")), spec: ckS(ckGet(t, r, "Item spec")), maker: ckS(ckGet(t, r, "Maker Nm")),
    });
  }
  return lines;
}
function ckParseStatus(t, fileName) {
  const lines = [];
  t.rows.forEach((r, i) => {
    const no = ckS(ckGet(t, r, "수주번호"));
    if (!no) return; // 거래처 계 / 합계 줄
    lines.push({
      file: fileName, row: t.headRow + 2 + i, no: no.toUpperCase(), seq: ckNum(ckGet(t, r, "수주순번", "순번")),
      custNo: ckS(ckGet(t, r, "고객주문번호")).toUpperCase(),
      vendCd: ckS(ckGet(t, r, "거래처코드")), vend: ckS(ckGet(t, r, "거래처명")),
      itemCd: ckS(ckGet(t, r, "품목코드")), name: ckS(ckGet(t, r, "품목명")), spec: ckS(ckGet(t, r, "규격")),
      unit: ckS(ckGet(t, r, "단위", "수주단위")), qty: ckNum(ckGet(t, r, "수주수량")), price: ckNum(ckGet(t, r, "수주단가")), hasPrice: t.col[ckKeyNorm("수주단가")] !== undefined,
      date: ckDate(ckGet(t, r, "납기일자")), prodSite: ckS(ckGet(t, r, "생산사업장")),
    });
  });
  return lines;
}

/** 전산이 내보낸 .xls 중에는 EXTERNSHEET 기록이 규격과 달라 읽기가 멈추는 것이 있다.
    값과는 상관없는 기록이라, 그 기록만 건너뛰게 바꿔서 다시 읽는다. */
function ckFixXls(buf) {
  const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: "array" });
  const ent = XLSX.CFB.find(cfb, "Workbook") || XLSX.CFB.find(cfb, "Book");
  const d = ent.content;
  for (let i = 0; i + 4 <= d.length;) {
    const id = d[i] | (d[i + 1] << 8), len = d[i + 2] | (d[i + 3] << 8);
    if (id === 0x17) { d[i] = 0xff; d[i + 1] = 0x0f; } // 0x0017 EXTERNSHEET → 쓰이지 않는 번호
    i += 4 + len;
  }
  return XLSX.CFB.write(cfb, { type: "array" });
}

/** 파일 하나를 읽어 주문서인지 수주현황인지 알아낸다. */
async function ckReadFile(file) {
  const buf = await file.arrayBuffer();
  let wb;
  try { wb = XLSX.read(buf, { type: "array" }); } catch (e) {
    try { wb = XLSX.read(ckFixXls(buf), { type: "array" }); } catch (e2) { return { kind: "bad", name: file.name, msg: "엑셀로 열 수 없음 (엑셀에서 열어 .xlsx로 다시 저장해 넣어 보세요)" }; }
  }
  const st = ckFindTable(wb, ["수주번호", ["수주순번", "순번"], "수주수량"]);
  if (st) return { kind: "status", name: file.name, lines: ckParseStatus(st, file.name) };
  const od = ckFindTable(wb, ["Order No", "Seq", "Qty"]);
  if (od) return { kind: "order", name: file.name, lines: ckParseOrder(od, file.name) };
  return { kind: "bad", name: file.name, msg: "주문서도 수주현황도 아님 (머리글을 못 찾음)" };
}

/** 비교. 결과: {issues:[…], okCount, orders:[…요약], noFile:[…], notUploaded:[…]} */
function ckCompare(orderFiles, status, opt) {
  const issues = [];
  const add = (level, kind, no, seq, msg, o, s) => issues.push({ level, kind, no, seq, msg, o: o || null, s: s || null });

  // 주문서: 같은 주문번호가 여러 파일에 있으면 처음 것만 쓴다
  const orders = new Map(); // no -> {file, lines: Map(seq->line)}
  const dupFiles = [];
  for (const f of orderFiles) {
    for (const l of f.lines) {
      let o = orders.get(l.no);
      if (!o) { o = { file: f.name, lines: new Map() }; orders.set(l.no, o); }
      if (o.file !== f.name) { if (!dupFiles.includes(`${l.no} (${o.file} / ${f.name})`)) dupFiles.push(`${l.no} (${o.file} / ${f.name})`); continue; }
      if (o.lines.has(l.seq)) add("err", "주문서 이상", l.no, l.seq, "주문서 원본 안에 같은 순번이 두 번 있습니다", l, null);
      else o.lines.set(l.seq, l);
    }
  }

  // 수주현황 색인 + 자체 점검
  const sIdx = new Map(); // no -> Map(seq -> [lines])
  for (const s of status.lines) {
    if (!sIdx.has(s.no)) sIdx.set(s.no, new Map());
    const m = sIdx.get(s.no);
    if (!m.has(s.seq)) m.set(s.seq, []);
    m.get(s.seq).push(s);
  }
  // 공장별 기준 거래처: 정해진 값이 없으면 수주현황에서 가장 많이 쓰인 거래처
  const facCount = {};
  for (const s of status.lines) { const f = s.no.slice(0, 2); (facCount[f] = facCount[f] || {})[s.vendCd] = (facCount[f][s.vendCd] || 0) + 1; }
  const facVend = {}, vendName = {};
  for (const s of status.lines) vendName[s.vendCd] = s.vend;
  for (const f of Object.keys(facCount)) facVend[f] = CK_FACTORY[f] || Object.entries(facCount[f]).sort((a, b) => b[1] - a[1])[0][0];

  for (const s of status.lines) {
    const f = s.no.slice(0, 2);
    if (s.vendCd !== facVend[f])
      add("err", "공장 다름", s.no, s.seq, `${f} 주문인데 거래처가 ${s.vendCd} ${s.vend} 로 올라가 있습니다 (${f}는 ${facVend[f]} ${vendName[facVend[f]] || ""})`, null, s);
    if (s.custNo && s.custNo !== s.no)
      add("err", "주문번호 다름", s.no, s.seq, `고객주문번호(${s.custNo})와 수주번호(${s.no})가 다릅니다`, null, s);
  }
  for (const [no, m] of sIdx) for (const [seq, arr] of m)
    if (arr.length > 1) add("err", "중복 등록", no, seq, `수주현황에 같은 수주번호·순번이 ${arr.length}번 있습니다 (엑셀 ${arr.map((x) => x.row).join(", ")}행)`, null, arr[0]);

  // 줄 단위 비교
  let okCount = 0;
  const loose = []; // 표기만 다른 품명(같은 품목으로 본 줄)
  const summary = [], notUploaded = [];
  for (const [no, o] of orders) {
    const m = sIdx.get(no);
    if (!m) { notUploaded.push({ no, file: o.file, n: o.lines.size }); continue; }
    let bad = 0;
    const flat = [...m.values()].flat();
    for (const [seq, l] of o.lines) {
      const arr = m.get(seq);
      if (!arr) {
        const moved = flat.find((x) => ckNameLevel(l.name, x.name) < 2 && ckSameNum(l.qty, x.qty) && !o.lines.has(x.seq));
        add("err", moved ? "순번 다름" : "빠짐", no, seq, moved ? `주문서 순번은 ${seq}인데 수주현황에는 순번 ${moved.seq}으로 올라가 있습니다` : "주문서에는 있는데 수주현황에 없습니다", l, moved || null);
        bad++; continue;
      }
      const s = arr[0], diffs = [];
      if (!ckSameNum(l.qty, s.qty)) diffs.push(["수량 다름", `수량: 주문서 ${ckFmt(l.qty)} → 수주현황 ${ckFmt(s.qty)}`]);
      const nl = ckNameLevel(l.name, s.name);
      if (nl === 2) diffs.push(["품명 다름", `품명: 주문서 "${l.name}" → 수주현황 "${s.name}"`]);
      else if (nl === 1) loose.push({ no, seq, o: l.name, s: s.name });
      if (opt.spec && !ckSameText(l.spec, s.spec)) diffs.push(["규격 다름", `규격: 주문서 "${l.spec}" → 수주현황 "${s.spec}"`]);
      if (opt.unit && !ckSameText(l.unit, s.unit)) diffs.push(["단위 다름", `단위: 주문서 ${l.unit} → 수주현황 ${s.unit}`]);
      if (opt.date && l.date !== s.date) diffs.push(["납기 다름", `납기: 주문서 ${l.date} → 수주현황 ${s.date}`]);
      if (opt.price && s.hasPrice && !ckSameNum(l.price, s.price)) diffs.push(["단가 다름", `단가: 주문서 ${l.price} → 수주현황 ${s.price}`]);
      if (diffs.length) { bad++; for (const [k, msg] of diffs) add("err", k, no, seq, msg, l, s); } else okCount++;
    }
    for (const [seq, arr] of m) if (!o.lines.has(seq)) {
      const already = issues.some((x) => x.kind === "순번 다름" && x.no === no && x.s === arr[0]);
      if (!already) { add("err", "남는 줄", no, seq, "수주현황에는 있는데 주문서 원본에 없는 순번입니다", null, arr[0]); bad++; }
    }
    summary.push({ no, file: o.file, orderN: o.lines.size, statusN: flat.length, bad });
  }
  const noFile = [...sIdx.keys()].filter((no) => !orders.has(no)).map((no) => ({ no, n: [...sIdx.get(no).values()].flat().length }));
  issues.sort((a, b) => a.no.localeCompare(b.no) || (a.seq ?? 0) - (b.seq ?? 0));
  return { issues, okCount, loose, summary, noFile, notUploaded, dupFiles, statusN: status.lines.length };
}
const ckFmt = (n) => (n === null || n === undefined ? "(없음)" : Number(n).toLocaleString("en-US", { maximumFractionDigits: 6 }));

/* ------------------------------------------------------------------ 화면 */
const ck$ = (id) => document.getElementById(id);
function ckOpts() {
  const o = {};
  document.querySelectorAll("[data-opt]").forEach((el) => { o[el.dataset.opt] = el.checked; });
  return o;
}
function ckToast(msg) { const t = ck$("toast"); t.textContent = msg; t.classList.add("show"); setTimeout(() => t.classList.remove("show"), 2500); }

async function ckAddFiles(fileList) {
  const bad = [];
  for (const file of fileList) {
    if (!/\.xlsx?$/i.test(file.name)) { bad.push(`${file.name}: 엑셀 파일이 아님`); continue; }
    const r = await ckReadFile(file);
    if (r.kind === "status") ckState.status = r;
    else if (r.kind === "order") { ckState.orders = ckState.orders.filter((o) => o.name !== r.name); ckState.orders.push(r); }
    else bad.push(`${r.name}: ${r.msg}`);
  }
  ck$("badFiles").innerHTML = bad.map((b) => `<div class="warn">⚠ ${ckEsc(b)}</div>`).join("");
  ckRender();
}

function ckRender() {
  const { orders, status } = ckState;
  const nOrderNo = new Set(orders.flatMap((o) => o.lines.map((l) => l.no))).size;
  ck$("chips").innerHTML =
    `<span class="chip ${orders.length ? "on" : ""}">주문서 원본 ${orders.length ? `${orders.length}개 (주문 ${nOrderNo}건, ${orders.reduce((a, o) => a + o.lines.length, 0)}줄)` : "아직 없음"}</span>` +
    `<span class="chip ${status ? "on" : ""}">수주현황 ${status ? `${ckEsc(status.name)} (${status.lines.length}줄)` : "아직 없음"}</span>`;
  ck$("drop").classList.toggle("has", !!(orders.length && status));
  const out = ck$("result");
  if (!status || !orders.length) {
    out.style.display = "none"; ckState.result = null;
    return;
  }
  const r = ckState.result = ckCompare(orders, status, ckOpts());
  out.style.display = "";
  const errN = r.issues.length;
  let h = "";
  if (!errN && !r.notUploaded.length && !r.noFile.length)
    h += `<div class="verdict good">✅ 틀린 곳 없음 — 주문서 ${r.okCount}줄이 수주현황과 모두 같습니다.</div>`;
  else if (!errN)
    h += `<div class="verdict mid">🟡 비교한 ${r.okCount}줄은 모두 같습니다. 다만 아래에 비교하지 못한 주문이 있습니다.</div>`;
  else
    h += `<div class="verdict bad">❌ 확인할 곳 ${errN}건 <small>(같은 줄 ${r.okCount}줄)</small></div>`;

  if (r.notUploaded.length)
    h += `<div class="box warnbox"><b>수주현황에 없는 주문서 ${r.notUploaded.length}건</b> — 아직 안 올렸거나, 다른 날·다른 회사 전산에 올린 주문입니다.<ul>` +
      r.notUploaded.map((x) => `<li><b>${ckEsc(x.no)}</b> (${x.n}줄) <span class="mut">${ckEsc(x.file)}</span></li>`).join("") + `</ul></div>`;
  if (r.noFile.length)
    h += `<div class="box warnbox"><b>주문서 원본을 안 넣어서 비교 못 한 주문 ${r.noFile.length}건</b> — 원본 파일을 추가로 끌어다 놓으면 비교됩니다. (공장·중복 점검만 했습니다)<ul>` +
      r.noFile.map((x) => `<li><b>${ckEsc(x.no)}</b> (수주현황 ${x.n}줄)</li>`).join("") + `</ul></div>`;
  if (r.dupFiles.length)
    h += `<div class="box warnbox"><b>같은 주문번호가 여러 파일에 있습니다</b> — 먼저 넣은 파일만 썼습니다.<ul>${r.dupFiles.map((x) => `<li>${ckEsc(x)}</li>`).join("")}</ul></div>`;

  if (errN) {
    h += `<table class="res"><thead><tr><th>수주번호</th><th>순번</th><th>무엇이</th><th>내용</th><th>주문서 원본</th><th>수주현황</th></tr></thead><tbody>` +
      r.issues.map((x) => `<tr><td class="nw">${ckEsc(x.no)}</td><td>${x.seq ?? ""}</td><td class="nw"><span class="tag red">${ckEsc(x.kind)}</span></td><td>${ckEsc(x.msg)}</td>` +
        `<td class="mut">${x.o ? `${ckEsc(x.o.name)}<br>${ckFmt(x.o.qty)} ${ckEsc(x.o.unit)}` : ""}</td>` +
        `<td class="mut">${x.s ? `${ckEsc(x.s.name)}<br>${ckFmt(x.s.qty)} ${ckEsc(x.s.unit)} · 엑셀 ${x.s.row}행` : ""}</td></tr>`).join("") + `</tbody></table>`;
  }
  if (r.loose.length)
    h += `<details class="box loosebox"><summary>품명 표기만 다른 줄 ${r.loose.length}줄 — 같은 품목으로 보고 넘어갔습니다 (눌러서 보기)</summary><table class="res"><thead><tr><th>수주번호</th><th>순번</th><th>주문서 품명</th><th>수주현황 품명</th></tr></thead><tbody>` +
      r.loose.map((x) => `<tr><td class="nw">${ckEsc(x.no)}</td><td>${x.seq}</td><td>${ckEsc(x.o)}</td><td>${ckEsc(x.s)}</td></tr>`).join("") + `</tbody></table></details>`;
  if (r.summary.length)
    h += `<h3>주문별 요약</h3><table class="res"><thead><tr><th>수주번호</th><th>주문서 줄</th><th>수주현황 줄</th><th>결과</th><th>파일</th></tr></thead><tbody>` +
      r.summary.map((x) => `<tr><td class="nw">${ckEsc(x.no)}</td><td>${x.orderN}</td><td>${x.statusN}</td><td>${x.bad ? `<span class="tag red">확인 ${x.bad}줄</span>` : `<span class="tag grn">같음</span>`}</td><td class="mut">${ckEsc(x.file)}</td></tr>`).join("") + `</tbody></table>`;
  ck$("resBody").innerHTML = h;
  ck$("btnDl").style.display = errN ? "" : "none";
}

function ckDownload() {
  const r = ckState.result; if (!r) return;
  const rows = [["수주번호", "순번", "무엇이", "내용", "주문서 품명", "주문서 수량", "수주현황 품명", "수주현황 수량", "수주현황 엑셀 행"]];
  for (const x of r.issues) rows.push([x.no, x.seq, x.kind, x.msg, x.o ? x.o.name : "", x.o ? x.o.qty : "", x.s ? x.s.name : "", x.s ? x.s.qty : "", x.s ? x.s.row : ""]);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [16, 6, 12, 60, 30, 10, 30, 10, 8].map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "확인할 곳");
  XLSX.writeFile(wb, "수주 확인 결과.xlsx");
}

function initCheck() {
  try { const saved = JSON.parse(localStorage.getItem(CK_OPT_KEY) || "{}"); document.querySelectorAll("[data-opt]").forEach((el) => { if (el.dataset.opt in saved) el.checked = !!saved[el.dataset.opt]; }); } catch (e) { /* 저장값 없음 */ }
  document.querySelectorAll("[data-opt]").forEach((el) => el.addEventListener("change", () => {
    try { localStorage.setItem(CK_OPT_KEY, JSON.stringify(ckOpts())); } catch (e) { /* 저장 불가 */ }
    ckRender();
  }));
  const drop = ck$("drop"), inp = ck$("files");
  drop.addEventListener("click", () => inp.click());
  inp.addEventListener("change", () => { ckAddFiles([...inp.files]); inp.value = ""; });
  ["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => ckAddFiles([...e.dataTransfer.files]));
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());
  ck$("btnClear").addEventListener("click", () => { ckState.orders = []; ckState.status = null; ck$("badFiles").innerHTML = ""; ckRender(); ckToast("비웠습니다"); });
  ck$("btnDl").addEventListener("click", ckDownload);
  ckRender();
}
