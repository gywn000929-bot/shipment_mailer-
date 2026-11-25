/* 🧮 발주 점검 — 주간 선박리스트 + ERP 발주 미납(.xls) 으로 업체(KET·경신 등) 발주 점검표를 만든다.
   계장님이 엑셀로 하던 일: 리스트에서 업체 줄 골라내기 → ERP 미납을 공장·납기월별 피벗 → VLOOKUP → 품목마다 부족 계산.
   여기서는 파일 두 개만 넣으면
   · 품목(공장+M코드)별로 월별 필요량(ETD 월) · 월별 기발주(ERP 납기일자 월) · 재고 → 달마다 모자라는 양
   · 모자라는 양을 포장단위(그 품목의 지난 발주수량들의 최대공약수)로 올린 "추천 추가 발주" (칸에서 고칠 수 있음)
   · 받기: 점검표 엑셀(계장님 작업본과 같은 열 + 공장·납기월별 추가 발주 시트), 업체에 보낼 잔량 일정 확인 메일(.eml) */
"use strict";
let ORD = null;   // {rows, months, vendors, edits:{key|m: qty}}

// ---------------------------------------------------------------- 파일 읽기
async function ordReadWb(file) {
  const buf = await file.arrayBuffer();
  try { return XLSX.read(buf, { type: "array", cellDates: true }); } catch (e) {
    // 전산 .xls 중 EXTERNSHEET 기록이 규격과 달라 못 읽는 파일: 그 기록만 건너뛰고 다시 읽는다(check.js 와 같은 방법)
    const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: "array" }), ent = XLSX.CFB.find(cfb, "Workbook") || XLSX.CFB.find(cfb, "Book"), d = ent.content;
    for (let i = 0; i + 4 <= d.length;) { const id = d[i] | (d[i + 1] << 8), len = d[i + 2] | (d[i + 3] << 8); if (id === 0x17) { d[i] = 0xff; d[i + 1] = 0x0f; } i += 4 + len; }
    return XLSX.read(XLSX.CFB.write(cfb, { type: "array" }), { type: "array", cellDates: true });
  }
}
// 날짜 칸: 날짜 / 엑셀 일련번호 / "2026-11-20" "11/20/2026" "2026.11.20" "20-Nov-26"
function ordDate(v) {
  if (v instanceof Date) return toDate(v);
  if (typeof v === "number" && v > 20000 && v < 80000) { const p = XLSX.SSF.parse_date_code(v); return new Date(p.y, p.m - 1, p.d); }
  const s = clean(v); let m;
  if ((m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = s.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})/))) return new Date(+m[3], +m[1] - 1, +m[2]);
  return toDate(s);
}
const ORD_CODE = /^M\d{3}-\d{5,7}$/i;
// 이 엑셀이 ERP 발주 미납(조회) 파일인가: 머리글에 미납수량·납기일자
function ordIsErp(wb) { return wb.SheetNames.some((sn) => findHeaderRow(sheetRows(wb, sn), ["미납수량", "납기일자"]) >= 0); }
/** ERP 발주 미납 → [{fac, cd, name, qty(미납), ord(발주수량), due, vendor}]
    머리글이 두 줄(업체/품목코드, 수주처/품목코드)이라, M코드 열은 값을 보고 찾는다. 공장은 수주번호(F1DTP…) 앞 두 글자. */
function ordParseErp(wb, fileName) {
  for (const sn of wb.SheetNames) {
    const rows = sheetRows(wb, sn), h = findHeaderRow(rows, ["미납수량", "납기일자"]); if (h < 0) continue;
    const hd = rows[h].map(norm), at = (...names) => { for (const n of names) { const i = hd.indexOf(norm(n)); if (i >= 0) return i; } return -1; };
    const cQty = at("미납수량"), cOrd = at("발주수량"), cDue = at("납기일자"), cName = at("품목명", "품명"), cVend = at("거래처명"), cSo = at("수주번호");
    const body = rows.slice(h + 1).filter((r) => r && r.some((c) => clean(c)));
    const width = Math.max(...body.map((r) => r.length), hd.length);
    let cCd = -1, best = 0;
    for (let c = 0; c < width; c++) { const n = body.filter((r) => ORD_CODE.test(clean(r[c]))).length; if (n > best) { best = n; cCd = c; } }
    if (cCd < 0) throw new Error(`${fileName}: 수주처 품목코드(M…) 열을 찾지 못했습니다.`);
    const out = [];
    for (const r of body) {
      const cd = clean(r[cCd]).toUpperCase(); if (!ORD_CODE.test(cd)) continue;
      const so = cSo >= 0 ? clean(r[cSo]) : "", fm = so.match(/^(F\d+)/i);
      out.push({ fac: fm ? fm[1].toUpperCase() : "", cd, name: cName >= 0 ? clean(r[cName]) : "", qty: toNum(r[cQty]), ord: cOrd >= 0 ? toNum(r[cOrd]) : 0, due: cDue >= 0 ? ordDate(r[cDue]) : null, vendor: cVend >= 0 ? clean(r[cVend]).replace(/\s+/g, " ") : "" });
    }
    return out;
  }
  throw new Error(`${fileName}: ERP 발주 미납 표(머리글 미납수량·납기일자)를 찾지 못했습니다.`);
}
// 주간 선박리스트 → [{fac, cd, maker, name, etd:[{d, q}], sum, bal, stock, br}]
function ordParseList(wb) {
  const sh = loadSheet(wb, ["ITEM CD", "MAKER"], "선박리스트", [], ["ETD"], [], "");
  const c = (names) => col(sh, names, { required: false });
  const cFac = c(["Fac", "FAC"]), cCd = c(["ITEM CD"]), cMk = c(["MAKER"]), cNm = c(["ITEM NM", "Part Name"]), cSum = c(["SUM", "sum"]), cBal = c(["BALANCE"]), cSt = c(["STOCK"]), cBr = c(["B-R"]), cB1 = c(["BALANCE 1"]);
  const etd = sh.headers.filter((h) => /^etd/i.test(clean(h))).map((h) => ({ h, d: toDate(clean(h).replace(/^ETD\s*/i, "")) })).filter((x) => x.d);
  // ETD 열은 매주 하나씩이다. 머리글의 연도 오타(예: 8-MAY-27)처럼 앞 열 + 7일에서 크게 벗어나면 앞 열 + 7일로 바로잡는다
  for (let i = 1; i < etd.length; i++) { const want = addDays(etd[i - 1].d, 7); if (Math.abs(etd[i].d - want) > 20 * 864e5) { log(`[확인] ETD 머리글 '${etd[i].h}' 날짜가 주 순서와 맞지 않아 ${fmtCellDate(want)}로 보고 계산합니다.`, "warn"); etd[i].d = want; } }
  const out = [];
  for (const r of sh.rows) {
    const cd = clean(r[cCd]).toUpperCase(); if (!cd) continue;
    const e = etd.map((x) => ({ d: x.d, q: toNum(r[x.h]) }));
    out.push({ fac: clean(r[cFac]).toUpperCase(), cd, maker: clean(r[cMk]), name: clean(r[cNm]), etd: e, sum: cSum ? toNum(r[cSum]) : e.reduce((a, x) => a + x.q, 0), bal: toNum(r[cBal]), stock: toNum(r[cSt]), br: cBr ? toNum(r[cBr]) : (cB1 ? toNum(r[cB1]) : toNum(r[cBal]) + toNum(r[cSt])) - (cSum ? toNum(r[cSum]) : 0), raw: r });
  }
  out.etdHeads = etd; out.sheet = sh.sheet;
  return out;
}

// ---------------------------------------------------------------- 계산
const gcd = (a, b) => { a = Math.round(a); b = Math.round(b); while (b) [a, b] = [b, a % b]; return Math.abs(a); };
const monthKey = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
const monthLabel = (k) => `${+k.slice(5)}월납`;
function ordCompute(list, erp) {
  // 업체: ERP 파일의 거래처명 → 설정 업체. 대상 품목 = ERP에 있는 품목 + 그 업체 MAKER의 리스트 품목
  const vendNames = uniq(erp.map((e) => e.vendor).filter(Boolean));
  const vends = uniq(vendNames.map((n) => matchAlias(n)).filter((n) => CONFIG.vendors[n]));
  const makers = new Set(vends.flatMap((n) => CONFIG.vendors[n].makers));
  const erpCodes = new Set(erp.map((e) => e.cd));
  const rows = list.filter((r) => erpCodes.has(r.cd) || makers.has(r.maker.toUpperCase()));
  // 달: 리스트 ETD 월(필요가 있는 달) / ERP 납기월. 첫 달 이전 납기(밀린 것)는 첫 달 공급으로, 마지막 달 이후는 "이후"로 따로 보여 준다
  const months = uniq(list.etdHeads.map((x) => monthKey(x.d))).sort();
  const first = months[0], last = months[months.length - 1];
  // 포장단위: 그 품목의 발주수량들의 최대공약수(발주 2건 이상) → 같은 계열(품명 첫 낱말: UL3271, AVSS, MG6…)의 최대공약수 → 1(올림 없음)
  const fam = (name) => (clean(name).toUpperCase().match(/^[A-Z]+\d{0,3}/) || [""])[0];
  const ordQ = new Map(), famG = new Map(), nameOf = new Map();
  for (const e of erp) if (e.ord > 0) { (ordQ.get(e.cd) || ordQ.set(e.cd, []).get(e.cd)).push(e.ord); const f = fam(e.name); famG.set(f, gcd(famG.get(f) || 0, e.ord)); nameOf.set(e.cd, e.name); }
  const unitOf = (cd, name) => { const q = ordQ.get(cd) || []; const g = q.length >= 2 ? q.reduce(gcd, 0) : 0; if (g > 1) return g; const f = famG.get(fam(nameOf.get(cd) || name)); return f > 1 ? f : 1; };
  const out = [];
  for (const r of rows) {
    const mine = erp.filter((e) => e.cd === r.cd && (!e.fac || !r.fac || e.fac === r.fac));
    const need = {}, sup = {}; let later = 0, nodue = 0;
    for (const m of months) { need[m] = 0; sup[m] = 0; }
    for (const x of r.etd) need[monthKey(x.d)] += x.q;
    for (const e of mine) {
      if (!e.due) { nodue += e.qty; continue; }
      const k = monthKey(e.due);
      if (k < first) sup[first] += e.qty; else if (k > last) later += e.qty; else sup[k] += e.qty;
    }
    if (nodue) sup[first] += nodue;
    const unit = unitOf(r.cd, r.name), short = {}, sugg = {};
    let avail = r.stock;
    for (const m of months) {
      avail += sup[m] - need[m];
      short[m] = avail < 0 ? -avail : 0;
      sugg[m] = short[m] ? Math.ceil(short[m] / unit) * unit : 0;
      avail += sugg[m];
    }
    const firstNeed = r.etd.find((x) => x.q > 0);
    out.push({ ...r, need, sup, later, short, sugg, unit, erpQty: mine.reduce((a, e) => a + e.qty, 0), vendor: vends[0] || "", firstNeed: firstNeed ? firstNeed.d : null });
  }
  return { rows: out, months, vends, vendNames, onlyErp: [...erpCodes].filter((cd) => !rows.some((r) => r.cd === cd)) };
}

// ---------------------------------------------------------------- 실행 · 화면
async function runOrder() {
  if (!CONFIG) return;
  const fl = $("ordList").files[0], fe = [...$("ordErp").files];
  if (!fl || !fe.length) { toast(!fl ? "주간 선박리스트 엑셀을 넣어 주세요." : "ERP 발주 미납 파일(.xls)을 넣어 주세요."); return; }
  clearLog(); setStatus("실행 중...", true); RESULT = { mails: [], unmatched: [], kind: "order", extraFiles: [] };
  try {
    const list = ordParseList(await ordReadWb(fl));
    log(`[리스트] 시트 '${list.sheet}' · ${list.length}품목 · ETD ${list.etdHeads.length}주 (${list.etdHeads.map((x) => fmtCellDate(x.d)).join(", ")})`);
    let erp = [];
    for (const f of fe) { const rows = ordParseErp(await ordReadWb(f), f.name); log(`[ERP 미납] ${f.name}: ${rows.length}줄 · 거래처 ${uniq(rows.map((r) => r.vendor)).join(", ") || "?"} · 미납 합계 ${fmtNum(rows.reduce((a, r) => a + r.qty, 0))}`); erp = erp.concat(rows); }
    ORD = { ...ordCompute(list, erp), edits: {} };
    if (!ORD.vends.length) log(`[확인] ERP 파일의 거래처명(${ORD.vendNames.join(", ")})을 설정 업체에서 찾지 못했습니다. 품목 코드가 같은 줄만 봅니다. 확인 메일은 만들지 않습니다.`, "warn");
    const nShort = ORD.rows.filter((r) => ORD.months.some((m) => r.sugg[m])).length;
    log(`[점검] ${ORD.vends.join(", ") || "업체 미확인"} · 대상 ${ORD.rows.length}품목 · 부족(추가 발주 추천) ${nShort}품목${ORD.onlyErp.length ? ` · 리스트에 없는 ERP 품목 ${ORD.onlyErp.length}개(이번 리스트에 필요 없음)` : ""}`, "ok");
    renderOrder(); ordMakeMail();
    setStatus("완료");
  } catch (e) { console.error(e); log("[오류] " + (e.message || e), "err"); setStatus("오류"); }
  renderResult();
}
const ordQty = (r, m) => { const k = r.fac + "|" + r.cd + "|" + m; return k in ORD.edits ? ORD.edits[k] : r.sugg[m]; };
function renderOrder() {
  const box = $("ordOut"); if (!ORD) { box.innerHTML = ""; return; }
  const onlyShort = $("ordOnlyShort").checked, ms = ORD.months;
  const rows = ORD.rows.filter((r) => !onlyShort || ms.some((m) => r.sugg[m] || ordQty(r, m)));
  const tot = (m) => ORD.rows.reduce((a, r) => a + toNum(ordQty(r, m)), 0);
  box.innerHTML = `<div class="ordSum">${ms.map((m) => `<span><b>${monthLabel(m)}</b> 추가 발주 ${ORD.rows.filter((r) => ordQty(r, m) > 0).length}품목 · ${fmtNum(tot(m)) || 0}</span>`).join("")}
     <button class="ok" id="ordDl">📥 점검표·발주 목록 엑셀 받기</button></div>
   <div class="ordWrap"><table class="ordTbl"><thead><tr><th rowspan="2">Fac</th><th rowspan="2">ITEM CD</th><th rowspan="2">ITEM NM</th><th rowspan="2">재고</th>${ms.map((m) => `<th colspan="3">${monthLabel(m)}</th>`).join("")}<th rowspan="2">이후 납기</th><th rowspan="2">B-R</th><th rowspan="2">단위</th></tr>
     <tr>${ms.map(() => "<th>필요</th><th>기발주</th><th class='sg'>추가 발주</th>").join("")}</tr></thead><tbody>
     ${rows.map((r) => `<tr class="${ms.some((m) => r.short[m]) ? "short" : ""}"><td>${esc(r.fac)}</td><td>${esc(r.cd)}</td><td class="nm">${esc(r.name)}</td><td class="n">${fmtNum(r.stock)}</td>
       ${ms.map((m) => `<td class="n">${fmtNum(r.need[m])}</td><td class="n">${fmtNum(r.sup[m])}</td><td class="sg"><input type="text" inputmode="numeric" data-k="${esc(r.fac + "|" + r.cd + "|" + m)}" value="${ordQty(r, m) ? fmtNum(ordQty(r, m)) : ""}" placeholder="${r.short[m] ? "부족 " + fmtNum(r.short[m]) : ""}"></td>`).join("")}
       <td class="n">${fmtNum(r.later)}</td><td class="n ${r.br < 0 ? "neg" : ""}">${r.br ? r.br.toLocaleString("en-US") : ""}</td><td class="n mut">${fmtNum(r.unit)}</td></tr>`).join("") || `<tr><td colspan="${7 + ms.length * 3}" class="hint">부족한 품목이 없습니다. "부족 품목만" 체크를 끄면 전체가 보입니다.</td></tr>`}
   </tbody></table></div>
   <div class="hint">추가 발주 = 재고 + 그달까지의 기발주로 그달까지의 필요량이 모자라는 만큼을 포장단위(지난 발주수량들의 최대공약수)로 올린 값입니다. 칸을 고치면 엑셀에도 고친 값이 들어갑니다. 빨간 줄 = 부족이 있는 품목.</div>`;
  $("ordDl").onclick = ordDownload;
}
function ordBind() {
  $("ordOut").addEventListener("change", (e) => { const k = e.target.dataset && e.target.dataset.k; if (!k) return; ORD.edits[k] = toNum(e.target.value); renderOrder(); ordMakeMail(); renderResult(); });
  $("ordOnlyShort").addEventListener("change", renderOrder);
  $("btnOrd").onclick = runOrder;
  bindDrop("dropOrdList", "ordList", "ordListName"); bindDrop("dropOrdErp", "ordErp", "ordErpName");
  for (const id of ["ordList", "ordErp"]) $(id).addEventListener("change", () => { if ($("ordList").files.length && $("ordErp").files.length) setTimeout(runOrder, 50); });
}

// ---------------------------------------------------------------- 받기: 엑셀 · 메일
function ordDownload() {
  const ms = ORD.months, wb = XLSX.utils.book_new(), etdH = ORD.rows[0] ? ORD.rows[0].etd.map((x) => "ETD " + fmtCellDate(x.d).toUpperCase()) : [];
  const hdr = ["Fac", "ITEM CD", "MAKER", "ITEM NM", ...etdH, "SUM", "BALANCE", "STOCK", "B-R", ...ms.map((m) => "기발주 " + monthLabel(m)), "이후 납기", ...ms.map((m) => "추가 발주 " + monthLabel(m))];
  const aoa = [hdr, ...ORD.rows.map((r) => [r.fac, r.cd, r.maker, r.name, ...r.etd.map((x) => x.q || null), r.sum, r.bal, r.stock || null, r.br, ...ms.map((m) => r.sup[m] || null), r.later || null, ...ms.map((m) => ordQty(r, m) || null)])];
  const ws = XLSX.utils.aoa_to_sheet(aoa); styleSheet(ws, aoa.length, hdr.length, null);
  for (let i = 1; i < aoa.length; i++) if (ms.some((m) => ordQty(ORD.rows[i - 1], m) > 0)) for (let c = hdr.length - ms.length; c < hdr.length; c++) { const a = XLSX.utils.encode_cell({ r: i, c }); if (ws[a]) ws[a].s = { ...ws[a].s, fill: { fgColor: { rgb: YELLOW } } }; }
  XLSX.utils.book_append_sheet(wb, ws, "발주 점검");
  // 공장·납기월별 추가 발주 목록 (ERP 입력용)
  for (const fac of uniq(ORD.rows.map((r) => r.fac)).sort()) for (const m of ms) {
    const rs = ORD.rows.filter((r) => r.fac === fac && ordQty(r, m) > 0); if (!rs.length) continue;
    const a2 = [["Ord Item Cd", "품목명", "발주수량"], ...rs.map((r) => [r.cd, r.name, ordQty(r, m)])];
    const w2 = XLSX.utils.aoa_to_sheet(a2); styleSheet(w2, a2.length, 3, 2);
    let nm = `${fac} ${monthLabel(m)} 발주`; if (wb.SheetNames.includes(nm)) nm = `${fac} ${m} 발주`; XLSX.utils.book_append_sheet(wb, w2, nm.slice(0, 31));
  }
  download(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })]), safeName(`${ORD.vends[0] || "업체"} 발주 점검 ${fmtDate(new Date()).slice(5).replace("-", "")}.xlsx`));
}
// 업체에 보내는 잔량 일정 확인 메일: 이번 리스트에 필요가 있고 기발주(미납)가 있는 품목, 필요 시점 빠른 순
function ordMakeMail() {
  RESULT.mails = RESULT.mails.filter((m) => m.kind !== "order");
  if (!ORD || !ORD.vends.length) return;
  const S = CONFIG.settings, v = CONFIG.vendors[ORD.vends[0]];
  const t = new Date(), monday = addDays(new Date(t.getFullYear(), t.getMonth(), t.getDate()), -((t.getDay() + 6) % 7)), [month, week] = weekLabel(monday);
  const rows = ORD.rows.filter((r) => r.firstNeed && r.erpQty > 0).sort((a, b) => a.firstNeed - b.firstNeed);
  if (!rows.length) return;
  const hdr = ["FAC", "Part Name", "미납수량", "필요 수량", "필요 시점", "입고 예정일"];
  const table = htmlTable(hdr, rows.map((r) => [r.fac, r.name, fmtNum(r.erpQty), fmtNum(r.sum), fmtCellDate(r.firstNeed), ""]), ["입고 예정일"]);
  const greet = (clean(S["출하요청인사"]) || "안녕하세요.") + " " + (S["발신자소개"] || "");
  const tpl = S["발주확인본문"] || `${greet}\n\n하기 품목 발주 잔량의 입고 예정일 확인 부탁드립니다.\n필요 시점이 빠른 순서입니다.\n\n{표}\n\n감사합니다.`;
  const subject = (S["발주확인제목형식"] || "대아홀딩스 {월}월 {주}째주 발주 잔량 일정 확인의 건").replace("{월}", month).replace("{주}", week).replace("{업체}", v.poShort);
  RESULT.mails.push({ kind: "order", vendor: v.name, subject, to: v.to, cc: uniq([...splitList(S["공통참조"]), ...v.cc]), body: templateHtml(tpl, {}, table), attachments: [], note: `발주 잔량 확인: ${rows.length}품목 (이번 리스트에 필요가 있고 기발주가 남은 품목)`, noTo: !v.to.length, include: v.to.length > 0, keys: [`ordchk|${fmtDate(monday)}|${v.name}`] });
}
