/* 입고 파일 만들기: 업체가 입고수량을 적어 보낸 엑셀(또는 주간 리스트) → 전산에 올리는 엑셀 ("1118 태양기전.xlsx" 모양)
   규칙 (실제 파일로 확인)
   - 이번 주(첫 ETD 열) 요청 수량이 있는 줄만, 원래 순서대로
   - Fac → 수주처 코드 (설정 "공장코드": F1=1010, F2=1060, F5=1150)
   - 열: 수주처 코드 | ITEM CD | MAKER | 업체코드 | ITEM NM | 요청 수량(첫 ETD) | 입고수량 */
"use strict";
const LS_CODES = "mailmaker_vendor_codes";
const VCOLS = ["업체", "거래처명(매칭)", "거래처명", "거래처", "VENDER", "VENDOR"].map((s) => s.replace(/\s+/g, "").toUpperCase());
const INB = { items: [] };   // {id, file, vendor, code, etd, rows:[{fac, cd, maker, nm, req, recv}], open}
let inbSeq = 1;
const up = (s) => clean(s).replace(/\s+/g, "").toUpperCase();

function codeOf(vendor) {
  const saved = store.get(LS_CODES, {});
  return (CONFIG.vendors[vendor] && CONFIG.vendors[vendor].code) || saved[vendor] || "";
}
function vendorFromFileName(name) {   // 파일 이름에 들어 있는 업체 이름(별칭 포함, 가장 긴 것)
  const key = normName(name.replace(/\.[^.]+$/, "")); let best = null;
  for (const [a, v] of Object.entries(CONFIG.aliasMap)) if (a.length >= 2 && key.includes(a) && (!best || a.length > best[0].length)) best = [a, v];
  return best ? best[1] : "";
}

// ---------------------------------------------------------------- 파일 읽기
function readInbound(fileName, wb) {
  let pick = null;
  for (const sn of wb.SheetNames) {
    const rows = sheetRows(wb, sn);
    for (let h = 0; h < Math.min(15, rows.length); h++) {
      const hd = (rows[h] || []).map(up), cCd = hd.indexOf("ITEMCD"), cMk = hd.indexOf("MAKER"), cEtd = hd.findIndex((c) => c.startsWith("ETD"));
      if (cCd < 0 || cMk < 0 || cEtd < 0) continue;
      const cand = { sn, rows, h, hd, cCd, cMk, cEtd, cV: hd.findIndex((c) => VCOLS.includes(c)), nEtd: hd.filter((c) => c.startsWith("ETD")).length };
      if (!pick || (cand.cV >= 0 && pick.cV < 0) || ((cand.cV >= 0) === (pick.cV >= 0) && cand.nEtd > pick.nEtd)) pick = cand;
      break;
    }
  }
  if (!pick) return [];
  const { rows, h, hd, cCd, cMk, cEtd, cV } = pick, cFac = hd.findIndex((c) => c === "FAC"), cNm = hd.findIndex((c) => c === "ITEMNM" || c === "PARTNAME"), cRecv = hd.indexOf("입고수량");
  const etd = clean(rows[h][cEtd]), groups = new Map();
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i] || [], cd = clean(r[cCd]), req = toNum(r[cEtd]);
    if (!cd || req <= 0) continue;
    const vendor = cV >= 0 ? matchAlias(clean(r[cV])) || "(거래처 없음)" : "";
    const rv = cRecv >= 0 ? r[cRecv] : null;
    if (!groups.has(vendor)) groups.set(vendor, []);
    groups.get(vendor).push({ fac: up(r[cFac]), cd, maker: clean(r[cMk]), nm: clean(r[cNm]), req, recv: rv == null || clean(rv) === "" ? null : toNum(rv) });
  }
  const out = [];
  for (const [v, rws] of groups) {
    const vendor = v || vendorFromFileName(fileName);
    out.push({ id: inbSeq++, file: fileName, vendor, code: vendor ? codeOf(vendor) : "", etd, rows: rws, open: cV < 0 });
  }
  return out;
}
async function addInboundFiles(files) {
  for (const f of files) {
    try {
      const got = readInbound(f.name, XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true }));
      if (!got.length) { toast(`${f.name}: Fac / ITEM CD / MAKER / ETD 열이 있는 표를 찾지 못했습니다`); continue; }
      INB.items.push(...got);
    } catch (e) { console.error(e); toast(`${f.name}: 읽지 못했습니다 (${e.message || e})`); }
  }
  renderInbound();
}

// ---------------------------------------------------------------- 전산 올리는 엑셀
function inboundBytes(it) { return inboundXlsx(it.etd, it.code, it.rows); }
function inboundName(it) { const d = $("inDate").value ? parseISO($("inDate").value) : new Date(); return safeName(`${pad2(d.getMonth() + 1)}${pad2(d.getDate())} ${it.vendor || "업체미정"}.xlsx`); }
const itemProblems = (it) => [!it.vendor && "업체를 골라 주세요", !it.code && "업체코드가 없습니다", it.rows.some((r) => !facCodes()[r.fac]) && "공장 코드를 모르는 줄이 있습니다", it.rows.every((r) => r.recv == null) && "입고수량이 모두 비어 있습니다"].filter(Boolean);

// ---------------------------------------------------------------- 화면
function renderInbound() {
  const box = $("inItems"); box.innerHTML = "";
  $("inTop").style.display = INB.items.length ? "" : "none";
  $("inCount").textContent = INB.items.length ? `업체 ${INB.items.length}곳` : "";
  const opts = (sel) => ['<option value="">(업체 고르기)</option>', ...CONFIG.order.map((n) => `<option ${n === sel ? "selected" : ""}>${esc(n)}</option>`), sel && !CONFIG.order.includes(sel) ? `<option selected>${esc(sel)}</option>` : ""].join("");
  for (const it of INB.items) {
    const d = document.createElement("div"); d.className = "mail"; d.dataset.id = it.id;
    const diff = it.rows.filter((r) => r.recv != null && r.recv !== r.req).length, empty = it.rows.filter((r) => r.recv == null).length, prob = itemProblems(it);
    d.innerHTML = `<div class="mh"><select data-act="vendor">${opts(it.vendor)}</select><label class="hint">업체코드 <input type="text" data-act="code" value="${esc(it.code)}" style="width:70px"></label>
       <span class="subj">${esc(it.file)} · ${esc(it.etd)} · ${it.rows.length}줄${diff ? ` · <b class="warn">요청과 다른 줄 ${diff}</b>` : ""}${empty ? ` · 빈칸 ${empty}` : ""}</span>
       <button class="sm" data-act="fill" title="입고수량이 빈 줄을 요청 수량으로 채웁니다">빈칸을 요청 수량대로</button><button class="sm" data-act="toggle">${it.open ? "접기" : "수량 보기·고치기"}</button><button class="sm pri" data-act="dl">📥 받기</button><button class="sm" data-act="del" title="목록에서 빼기">✕</button></div>
       ${prob.length ? `<div class="meta"><span class="warn">⚠ ${esc(prob.join(" · "))}</span></div>` : ""}
       ${it.open ? `<table class="intbl"><tr><th>Fac</th><th>ITEM CD</th><th>MAKER</th><th>ITEM NM</th><th style="text-align:right">요청</th><th>입고수량</th></tr>${it.rows.map((r, i) => `<tr class="${r.recv != null && r.recv !== r.req ? "chg" : ""}"><td>${esc(r.fac)}</td><td>${esc(r.cd)}</td><td>${esc(r.maker)}</td><td>${esc(r.nm)}</td><td style="text-align:right">${fmtNum(r.req)}</td><td><input type="text" data-row="${i}" value="${r.recv == null ? "" : r.recv}" style="width:110px;text-align:right"></td></tr>`).join("")}</table>` : ""}`;
    box.appendChild(d);
  }
}
function bindInbound() {
  bindDrop("dropIn", "inFiles", "inName");
  $("inFiles").addEventListener("change", () => addInboundFiles([...$("inFiles").files]));
  $("inDate").value = fmtDate(new Date());
  const find = (el) => INB.items.find((x) => x.id === +el.closest(".mail").dataset.id);
  $("inItems").addEventListener("change", (e) => {
    const t = e.target, it = find(t); if (!it) return;
    if (t.dataset.act === "vendor") { it.vendor = t.value; it.code = codeOf(it.vendor); renderInbound(); }
    else if (t.dataset.act === "code") { it.code = t.value.trim(); if (it.vendor && it.code && !(CONFIG.vendors[it.vendor] || {}).code) { const s = store.get(LS_CODES, {}); s[it.vendor] = it.code; store.set(LS_CODES, s); } renderInbound(); }
    else if (t.dataset.row != null) { const r = it.rows[+t.dataset.row], v = t.value.replace(/,/g, "").trim(); r.recv = v === "" ? null : toNum(v); t.closest("tr").classList.toggle("chg", r.recv != null && r.recv !== r.req); }
  });
  $("inItems").addEventListener("click", (e) => {
    const b = e.target.closest("button"); if (!b) return; const it = find(b), act = b.dataset.act;
    if (act === "fill") { it.rows.forEach((r) => { if (r.recv == null) r.recv = r.req; }); renderInbound(); }
    else if (act === "toggle") { it.open = !it.open; renderInbound(); }
    else if (act === "del") { INB.items = INB.items.filter((x) => x !== it); renderInbound(); }
    else if (act === "dl") { const p = itemProblems(it); if (p.length && !confirm(`${it.vendor || it.file}: ${p.join(", ")}\n그래도 받을까요?`)) return; download(new Blob([inboundBytes(it)]), inboundName(it)); }
  });
  $("btnInFillAll").onclick = () => { INB.items.forEach((it) => it.rows.forEach((r) => { if (r.recv == null) r.recv = r.req; })); renderInbound(); toast("입고수량이 빈 줄을 모두 요청 수량으로 채웠습니다. 실제와 다른 줄은 고쳐 주세요."); };
  $("btnInClear").onclick = () => { INB.items = []; $("inFiles").value = ""; $("inName").textContent = "아직 없음"; $("dropIn").classList.remove("has"); renderInbound(); };
  $("btnInZip").onclick = async () => {
    const bad = INB.items.filter((it) => itemProblems(it).length);
    if (bad.length && !confirm(`확인이 필요한 업체가 ${bad.length}곳 있습니다:\n${bad.map((it) => " - " + (it.vendor || it.file) + ": " + itemProblems(it).join(", ")).join("\n")}\n\n그래도 전부 받을까요?`)) return;
    const zip = new JSZip(), used = new Set();
    for (const it of INB.items) { let n = inboundName(it), k = 2; while (used.has(n)) n = inboundName(it).replace(/\.xlsx$/, ` (${k++}).xlsx`); used.add(n); zip.file(n, inboundBytes(it)); }
    const d = parseISO($("inDate").value || fmtDate(new Date()));
    download(await zip.generateAsync({ type: "blob" }), `입고파일_${pad2(d.getMonth() + 1)}${pad2(d.getDate())}.zip`);
  };
  renderInbound();
}
async function initInbound() {
  const q = new URLSearchParams(location.search).get("who"); let last = null; try { last = localStorage.getItem("mailmaker_cfg"); } catch {}
  const name = q === "kim" ? "config_kim.xlsx" : q === "nam" ? "config.xlsx" : last && /^config.*\.xlsx$/.test(last) ? last : "config.xlsx";
  try {
    const res = await fetch(name, { cache: "no-store" }); if (!res.ok) throw new Error(`${name} 을(를) 불러오지 못했습니다 (${res.status})`);
    CONFIG = parseConfig(XLSX.read(await res.arrayBuffer(), { type: "array", cellDates: true }), name); CONFIG_NAME = name;
    $("inCfg").textContent = `업체 ${CONFIG.order.length}곳 · 업체코드가 있는 업체 ${CONFIG.order.filter((n) => CONFIG.vendors[n].code).length}곳`;
  } catch (e) { console.error(e); $("inCfg").textContent = "설정을 못 읽음: " + (e.message || e); return; }
  bindInbound();
}
