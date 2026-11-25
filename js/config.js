/* 설정 엑셀(config*.xlsx) 읽기, 이 PC에만 있는 변경, 설정 고치기 화면, 설정 내려받기 */
"use strict";
const HDR_DEFAULT = ["업체", "받는사람", "추가참조", "수신줄", "호칭", "MAKER", "엑셀첨부", "메모", "활성", "거래처명(주문서)", "발주서발송", "주문서약칭", "주문서추가문구", "출하요청발송", "업체코드"];
const CHK_FIELDS = ["활성", "엑셀첨부", "발주서발송", "출하요청발송"];
const LIST_FIELDS = { to: "받는사람", makers: "MAKER", poNames: "거래처명(주문서)" };
const LS_OVERLAY = "mailmaker_overlay_v1";   // 결과 화면에서 바로 적은 것(받는 사람 추가, 미지정 업체 지정)
const LS_PATCH = "mailmaker_patch_v1";       // "설정 고치기" 화면에서 바꾼 것
const DEFAULT_REPO = "gywn000929-bot/shipment_mailer-";

let CONFIG = null;      // {vendors, order, makerMap, itemMap, settings, aliasMap, name, wb, base, eff, pending}
let CONFIG_NAME = "config.xlsx";

function matchAlias(name) {   // 리스트의 업체 칸 값 → 설정 업체명. 1) 정확히 일치 2) 별칭 포함(긴 별칭 우선)
  const key = normName(name); if (!key) return name;
  if (key in CONFIG.aliasMap) return CONFIG.aliasMap[key];
  let best = null;
  for (const [a, v] of Object.entries(CONFIG.aliasMap)) if (a.length >= 2 && (key.includes(a) || a.includes(key)) && (!best || a.length > best[0].length)) best = [a, v];
  return best ? best[1] : name;
}
function resolveVendor(rowVendor, itemCd, maker) {
  if (rowVendor) return matchAlias(rowVendor);
  if (norm(itemCd) in CONFIG.itemMap) return CONFIG.itemMap[norm(itemCd)];
  return CONFIG.makerMap[clean(maker).toUpperCase()] || "";
}

// ---------------------------------------------------------------- 엑셀 → 기본 설정(BASE)
function readBase(wb) {
  let hdr = (sheetRows(wb, "업체")[0] || []).map(clean).filter(Boolean);
  for (const h of HDR_DEFAULT) if (!hdr.includes(h)) hdr.push(h);
  const rows = [], notes = [];
  for (const r of sheetObjs(wb, "업체")) {
    const name = clean(r["업체"]);
    if (!name) continue;
    if (name.startsWith("※")) { notes.push(name); continue; }   // 시트 아래쪽 설명 줄
    rows.push(Object.fromEntries(hdr.map((h) => [h, clean(r[h])])));
  }
  const settings = [];
  for (const r of sheetRows(wb, "설정")) {
    const k = clean(r[0]); if (!k || k === "항목") continue;
    settings.push([k, clean(r[1]).replace(/\\n/g, "\n").replace(/\r/g, ""), clean(r[2])]);
  }
  return { hdr, rows, notes, settings };
}

// ---------------------------------------------------------------- 이 PC에만 있는 변경
function getLocal(name) {
  const o = store.get(LS_OVERLAY, {})[name] || {}, p = store.get(LS_PATCH, {})[name] || {};
  return { ov: { vendors: o.vendors || {}, newVendors: o.newVendors || [] }, patch: { vendors: p.vendors || {}, newRows: p.newRows || [], settings: p.settings || {} } };
}
function setLocal(name, local) {
  const o = store.get(LS_OVERLAY, {}), p = store.get(LS_PATCH, {});
  o[name] = local.ov; p[name] = local.patch;
  if (!Object.keys(local.ov.vendors).length && !local.ov.newVendors.length) delete o[name];
  if (!Object.keys(local.patch.vendors).length && !local.patch.newRows.length && !Object.keys(local.patch.settings).length) delete p[name];
  store.set(LS_OVERLAY, o); store.set(LS_PATCH, p);
}
const sameField = (f, a, b) => CHK_FIELDS.includes(f) ? isNo(a) === isNo(b) : clean(a) === clean(b);
const hasItem = (field, list, v) => field === "MAKER" ? list.map((x) => x.toUpperCase()).includes(String(v).toUpperCase()) : list.includes(v);

// 올려 둔 설정 파일에 이미 들어간 변경은 이 PC 저장분에서 지운다 → 남은 것이 "아직 안 올린 변경"
function reconcile(base, local) {
  const byName = Object.fromEntries(base.rows.map((r) => [r["업체"], r]));
  const { ov, patch } = local;
  patch.newRows = patch.newRows.filter((nr) => {
    const b = byName[nr["업체"]]; if (!b) return true;
    const p = patch.vendors[nr["업체"]] = patch.vendors[nr["업체"]] || {};
    for (const [f, v] of Object.entries(nr)) if (f !== "업체" && !(f in p)) p[f] = v;
    return false;
  });
  for (const [n, p] of Object.entries(patch.vendors)) {
    const b = byName[n];
    if (!b) { delete patch.vendors[n]; continue; }
    for (const f of Object.keys(p)) if (sameField(f, p[f], b[f])) delete p[f];
    if (!Object.keys(p).length) delete patch.vendors[n];
  }
  const baseSet = Object.fromEntries(base.settings.map((s) => [s[0], s[1]]));
  for (const k of Object.keys(patch.settings)) if (k in baseSet && baseSet[k] === patch.settings[k]) delete patch.settings[k];
  ov.newVendors = ov.newVendors.filter((nv) => {
    if (!byName[nv.name]) return true;
    const t = ov.vendors[nv.name] = ov.vendors[nv.name] || {};
    for (const k of Object.keys(LIST_FIELDS)) t[k] = uniq([...(t[k] || []), ...(nv[k] || [])]);
    return false;
  });
  const newNames = [...ov.newVendors.map((x) => x.name), ...patch.newRows.map((x) => x["업체"])];
  for (const [n, o] of Object.entries(ov.vendors)) {
    const b = byName[n];
    if (!b) { if (!newNames.includes(n)) delete ov.vendors[n]; continue; }
    let left = 0;
    for (const [k, f] of Object.entries(LIST_FIELDS)) { const cur = splitList(b[f]); o[k] = (o[k] || []).filter((v) => !hasItem(f, cur, v)); left += o[k].length; }
    if (!left) delete ov.vendors[n];
  }
  return local;
}
// 기본 설정 + 이 PC 변경 = 실제로 쓰는 설정
function effective(base, local) {
  const { ov, patch } = local;
  const blank = Object.fromEntries(base.hdr.map((h) => [h, ""]));
  const rows = base.rows.map((r) => ({ ...r, ...(patch.vendors[r["업체"]] || {}) }));
  const has = (n) => rows.some((r) => r["업체"] === n);
  for (const nr of patch.newRows) if (nr["업체"] && !has(nr["업체"])) rows.push({ ...blank, ...nr });
  for (const nv of ov.newVendors) if (!has(nv.name)) rows.push({ ...blank, 업체: nv.name, 받는사람: (nv.to || []).join(", "), MAKER: (nv.makers || []).join(", "), "거래처명(주문서)": (nv.poNames || []).join(", "), 엑셀첨부: "N", 활성: "Y", 발주서발송: "Y", 주문서약칭: nv.name });
  for (const r of rows) {
    const o = ov.vendors[r["업체"]]; if (!o) continue;
    for (const [k, f] of Object.entries(LIST_FIELDS)) { const cur = splitList(r[f]); for (const v of o[k] || []) if (!hasItem(f, cur, v)) cur.push(v); r[f] = cur.join(", "); }
  }
  const settings = base.settings.map(([k, v, d]) => [k, k in patch.settings ? patch.settings[k] : v, d]);
  for (const [k, v] of Object.entries(patch.settings)) if (!settings.some((s) => s[0] === k)) settings.push([k, v, ""]);
  return { rows, settings };
}
// 아직 GitHub에 안 올린 변경을 사람이 읽을 수 있게
function pendingList(base, eff) {
  const out = [], byName = Object.fromEntries(base.rows.map((r) => [r["업체"], r]));
  for (const r of eff.rows) {
    const b = byName[r["업체"]];
    if (!b) { out.push(`새 업체: ${r["업체"]}`); continue; }
    const f = base.hdr.filter((h) => !sameField(h, r[h], b[h]));
    if (f.length) out.push(`${r["업체"]}: ${f.join(", ")}`);
  }
  const bs = Object.fromEntries(base.settings.map((s) => [s[0], s[1]]));
  for (const [k, v] of eff.settings) if (bs[k] !== v) out.push(`설정: ${k}`);
  return out;
}

function parseConfig(wb, name) {
  const base = readBase(wb);
  const had = hadLocal(getLocal(name));
  const local = reconcile(base, getLocal(name));
  setLocal(name, local);
  const eff = effective(base, local);
  const settings = Object.fromEntries(eff.settings.map((s) => [s[0], s[1]]));
  const vendors = {}, makerMap = {}, itemMap = {}, aliasMap = {}, order = [];
  for (const r of eff.rows) {
    const vname = clean(r["업체"]);
    if (!vname || isNo(r["활성"])) continue;
    const v = {
      name: vname,
      to: uniq(splitList(r["받는사람"])),
      cc: splitList(r["추가참조"]),
      recvLine: clean(r["수신줄"]),
      title: clean(r["호칭"]),
      makers: splitList(r["MAKER"]).map((m) => m.toUpperCase()),
      attach: !isNo(r["엑셀첨부"]),
      memo: clean(r["메모"]),
      poNames: splitList(r["거래처명(주문서)"]),
      sendPo: !isNo(r["발주서발송"]),
      sendShip: !isNo(r["출하요청발송"]),
      code: clean(r["업체코드"]),
      poExtra: clean(r["주문서추가문구"]),
      poShort: clean(r["주문서약칭"]) || vname,
    };
    vendors[vname] = v; order.push(vname);
    for (const m of v.makers) if (!(m in makerMap)) makerMap[m] = vname;
    for (const a of [vname, ...v.poNames, v.poShort]) if (normName(a) && !(normName(a) in aliasMap)) aliasMap[normName(a)] = vname;
  }
  for (const r of sheetObjs(wb, "품목지정")) { const c = clean(r["ITEM CD"]), v = clean(r["업체"]); if (c && v) itemMap[norm(c)] = v; }
  const pending = pendingList(base, eff);
  return { vendors, order, makerMap, itemMap, settings, aliasMap, name, wb, base, eff, pending, synced: had && !pending.length };
}
const hadLocal = (l) => Object.keys(l.ov.vendors).length + l.ov.newVendors.length + Object.keys(l.patch.vendors).length + l.patch.newRows.length + Object.keys(l.patch.settings).length > 0;

let CONFIG_SRC = null;   // 마지막으로 읽은 설정 엑셀(다시 계산할 때 재사용)
async function reloadConfig({ refetch = true } = {}) {
  try {
    const sel = $("config").value;
    if (refetch || !CONFIG_SRC) {
      if (sel === "__file__") {
        const f = $("configFile").files[0];
        if (!f) { toast("설정 엑셀 파일을 골라 주세요"); return; }
        CONFIG_SRC = { wb: XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true }), name: f.name };
      } else {
        const res = await fetch(sel, { cache: "no-store" });
        if (!res.ok) throw new Error(`${sel} 을(를) 불러오지 못했습니다 (${res.status})`);
        CONFIG_SRC = { wb: XLSX.read(await res.arrayBuffer(), { type: "array", cellDates: true }), name: sel };
      }
    }
    CONFIG = parseConfig(CONFIG_SRC.wb, CONFIG_SRC.name);
    CONFIG_NAME = CONFIG.name;
    try { if (sel !== "__file__") localStorage.setItem("mailmaker_cfg", sel); } catch {}
    const who = senderName();
    $("whoBadge").textContent = "보내는 사람: " + (who || CONFIG_NAME); $("whoBadge").className = "badge " + (/김민지/.test(who) ? "kim" : "nam");
    document.title = `메일 작성기 - ${who || CONFIG_NAME}`;
    const noMail = CONFIG.order.filter((n) => !CONFIG.vendors[n].to.length);
    $("cfgInfo").textContent = `업체 ${CONFIG.order.length}곳` + (noMail.length ? ` · 메일 주소 없는 업체 ${noMail.length}곳` : "");
    renderPending();
    if (CONFIG.synced) toast("이 PC에서 고친 내용이 올려 둔 설정 파일에 모두 들어 있습니다. 이제 다른 PC에서도 똑같이 적용됩니다.");
    if (typeof onConfigLoaded === "function") onConfigLoaded();
  } catch (e) { console.error(e); toast(String(e.message || e)); $("cfgInfo").textContent = "설정을 못 읽음: " + (e.message || e); }
}
// "내 업체만": 설정 "내업체"(이 업체들만) 또는 "내업체제외"(이 업체들 빼고 전부). 둘 다 없으면 null
function myVendors() {
  const S = CONFIG.settings, inc = splitList(S["내업체"]).map(matchAlias), exc = splitList(S["내업체제외"]).map(matchAlias);
  if (inc.length) return CONFIG.order.filter((n) => inc.includes(n));
  if (exc.length) return CONFIG.order.filter((n) => !exc.includes(n));
  return null;
}
const senderName = () => (CONFIG.settings["발신자소개"] || "").replace(/^대아홀딩스\s*/, "").replace(/\s*(입니다|이다)\.?$/, "").trim();
function repoUrl() {
  const m = location.hostname.match(/^([^.]+)\.github\.io$/), seg = location.pathname.split("/").filter(Boolean)[0];
  return "https://github.com/" + (m && seg ? `${m[1]}/${seg}` : DEFAULT_REPO);
}

// ---------------------------------------------------------------- "이 PC에만 저장됨" 알림 띠
function renderPending() {
  const bar = $("pendingBar"), p = CONFIG.pending;
  if (!p.length) { bar.style.display = "none"; return; }
  bar.style.display = "";
  $("pendingText").innerHTML = `<b>이 PC에만 저장된 설정 변경 ${p.length}건</b> — 다른 PC에서 열면 보이지 않습니다. <span class="hint">(${esc(p.slice(0, 4).join(" / "))}${p.length > 4 ? " …" : ""})</span>`;
}
function showUploadGuide() {
  const fromFile = $("config").value === "__file__";
  $("uploadGuide").style.display = "";
  $("uploadGuide").innerHTML = `<b>방금 받은 ${esc(CONFIG_NAME)} 파일을 올리는 방법</b> (한 번 올리면 어느 PC에서나 같은 설정이 됩니다)
   <ol><li><a href="${repoUrl()}/upload/main" target="_blank" rel="noopener"><b>여기를 눌러 올리는 화면 열기</b></a> (GitHub에 로그인되어 있어야 합니다)</li>
   <li>열린 화면 가운데 네모 칸에 방금 받은 <b>${esc(CONFIG_NAME)}</b> 파일을 끌어다 놓습니다. <span class="hint">파일 이름이 "${esc(CONFIG_NAME.replace(".xlsx", " (1).xlsx"))}"처럼 바뀌어 있으면 먼저 "${esc(CONFIG_NAME)}"로 고쳐 주세요.</span></li>
   <li>아래 초록색 <b>Commit changes</b> 버튼을 누릅니다.</li>
   <li>2~3분 뒤 이 화면을 새로 고치면(F5) 위의 노란 알림이 사라집니다. 사라졌으면 끝입니다.</li></ol>
   <div class="hint">GitHub 계정이 없거나 올리기가 어려우면, 받은 파일을 남효주에게 메일로 보내 주세요. 그 전까지는 이 PC에서는 지금 고친 대로 계속 쓸 수 있습니다.${fromFile ? " (지금은 내 PC의 설정 파일을 쓰는 중이라, 받은 파일을 다음에 다시 올려 써도 됩니다.)" : ""}</div>`;
  $("uploadGuide").scrollIntoView({ behavior: "smooth", block: "center" });
}

// ---------------------------------------------------------------- 설정 엑셀 내려받기 (이 PC 변경을 합친 것)
function configWorkbook() {
  const { base, eff, wb } = CONFIG;
  const head = { font: { bold: true }, fill: { fgColor: { rgb: "DDEBF7" } } };
  const out = XLSX.utils.book_new();
  const sheet = (aoa, wrapCol) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa), widths = [];
    aoa.forEach((row, r) => row.forEach((v, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (!cell) return;
      if (r === 0) cell.s = head;
      else if (c === wrapCol && String(v).includes("\n")) cell.s = { alignment: { wrapText: true, vertical: "top" } };
      if (!String(v).startsWith("※")) widths[c] = Math.max(widths[c] || 0, String(v ?? "").split("\n")[0].length);
    }));
    ws["!cols"] = widths.map((w) => ({ wch: Math.min(Math.max(10, (w || 0) * 1.4 + 2), 70) }));
    return ws;
  };
  XLSX.utils.book_append_sheet(out, sheet([base.hdr, ...eff.rows.map((r) => base.hdr.map((h) => r[h] ?? "")), [], ...base.notes.map((n) => [n])]), "업체");
  for (const sn of wb.SheetNames) {
    if (sn === "업체") continue;
    if (sn === "설정") XLSX.utils.book_append_sheet(out, sheet([["항목", "값", "설명"], ...eff.settings], 1), sn);
    else XLSX.utils.book_append_sheet(out, wb.Sheets[sn], sn);
  }
  if (!wb.SheetNames.includes("설정")) XLSX.utils.book_append_sheet(out, sheet([["항목", "값", "설명"], ...eff.settings], 1), "설정");
  return out;
}
function downloadConfig() {
  download(new Blob([XLSX.write(configWorkbook(), { type: "array", bookType: "xlsx" })]), CONFIG_NAME);
  showUploadGuide();
}

// ---------------------------------------------------------------- 설정 고치기 화면
const ED_COLS = [   // [엑셀 열 이름, 화면에 보일 이름, 종류, 너비]
  ["활성", "사용", "chk", 44], ["업체", "업체", "name", 120], ["받는사람", "받는 사람 메일 (여러 명은 쉼표)", "text", 280], ["추가참조", "이 업체만 참조 추가", "text", 200],
  ["MAKER", "리스트의 MAKER 이름", "text", 170], ["거래처명(주문서)", "주문서·리스트에 적힌 업체 이름", "text", 200], ["출하요청발송", "출하 요청 메일", "chk", 60], ["발주서발송", "주문서 메일", "chk", 60],
  ["엑셀첨부", "선박 엑셀 첨부", "chk", 60], ["주문서약칭", "제목용 짧은 이름", "text", 100], ["주문서추가문구", "주문서 본문 추가 한 줄", "text", 200],
  ["업체코드", "전산 업체코드", "text", 80], ["수신줄", "본문 첫 줄(수신: …)", "text", 170], ["호칭", "호칭", "text", 80], ["메모", "메모", "text", 240],
];
let ED = null;   // {rows:[{...,_new}], settings:[[k,v,d]]}
function edShowTab(t) {
  document.querySelectorAll(".edTab").forEach((b) => b.classList.toggle("on", b.dataset.edtab === t));
  $("edMail").style.display = t === "mail" ? "" : "none"; $("edVend").style.display = t === "vend" ? "" : "none";
  if (t === "mail" && typeof renderMailText === "function") renderMailText();
}
function openEditor(filterNoMail, tab) {
  if (!CONFIG) return;
  ED = { rows: CONFIG.eff.rows.map((r) => ({ ...r })), settings: CONFIG.eff.settings.map((s) => [...s]) };
  const names = CONFIG.base.rows.map((r) => r["업체"]);
  ED.rows.forEach((r) => { r._new = !names.includes(r["업체"]); });
  $("edNoMail").checked = !!filterNoMail; $("edFilter").value = "";
  $("editCard").style.display = ""; renderEditor(); edShowTab(tab || "vend");
  $("editCard").scrollIntoView({ behavior: "smooth" });
}
function edChanged(r, f) { const b = CONFIG.base.rows.find((x) => x["업체"] === r["업체"]); return r._new || !b || !sameField(f, r[f], b[f]); }
function renderEditor() {
  const q = norm($("edFilter").value), onlyNo = $("edNoMail").checked;
  const head = "<tr>" + ED_COLS.map((c) => `<th style="min-width:${c[3]}px">${esc(c[1])}</th>`).join("") + "<th></th></tr>";
  const body = ED.rows.map((r, i) => {
    if (onlyNo && (splitList(r["받는사람"]).length || isNo(r["활성"])) && !r._touched) return "";
    if (q && !norm(ED_COLS.map((c) => r[c[0]]).join(" ")).includes(q)) return "";
    return `<tr class="${isNo(r["활성"]) ? "off" : ""}">` + ED_COLS.map(([f, , kind]) => {
      const ch = edChanged(r, f) ? " chg" : "";
      if (kind === "chk") return `<td class="c${ch}"><input type="checkbox" data-r="${i}" data-f="${esc(f)}" ${isNo(r[f]) ? "" : "checked"}></td>`;
      if (kind === "name" && !r._new) return `<td class="nm"><b>${esc(r[f])}</b></td>`;
      const miss = f === "받는사람" && !splitList(r[f]).length && !isNo(r["활성"]) ? " miss" : "";
      return `<td class="${ch}${miss}"><input type="text" data-r="${i}" data-f="${esc(f)}" value="${esc(r[f] ?? "")}"${f === "받는사람" ? ' placeholder="메일 주소를 적어 주세요"' : f === "업체" ? ' placeholder="새 업체 이름"' : ""}></td>`;
    }).join("") + `<td>${r._new ? `<button class="sm" data-del="${i}" title="새로 넣은 줄 지우기">✕</button>` : ""}</td></tr>`;
  }).join("");
  $("edTbl").innerHTML = head + (body || `<tr><td colspan="${ED_COLS.length + 1}" class="hint">해당하는 업체가 없습니다.</td></tr>`);
  $("edSet").innerHTML = ED.settings.map(([k, v, d], i) => {
    if (typeof MT_KEYS !== "undefined" && MT_KEYS.includes(k)) return "";   // 메일 문구 화면에서 고치는 칸
    const b = CONFIG.base.settings.find((s) => s[0] === k), ch = !b || b[1] !== v ? " chg" : "";
    return `<div class="setrow"><label><b>${esc(k)}</b><span class="hint">${esc(d)}</span></label><textarea class="${ch}" data-s="${i}" rows="${Math.min(8, Math.max(1, String(v).split("\n").length))}">${esc(v)}</textarea></div>`;
  }).join("");
  edCountUpdate();
}
function edCountUpdate() { const n = editorPatch().count; $("edCount").textContent = n ? `올려 둔 설정과 다른 곳 ${n}군데 (노란 칸)` : "올려 둔 설정과 같습니다"; }
function editorPatch() {
  const patch = { vendors: {}, newRows: [], settings: {} }; let count = 0; const seen = new Set();
  for (const r of ED.rows) {
    const name = clean(r["업체"]); if (!name || seen.has(name)) continue; seen.add(name);
    const b = CONFIG.base.rows.find((x) => x["업체"] === name);
    const row = Object.fromEntries(CONFIG.base.hdr.map((h) => [h, clean(r[h])]));
    if (!b) { patch.newRows.push({ ...row, 활성: row["활성"] || "Y", 엑셀첨부: row["엑셀첨부"] || "N", 발주서발송: row["발주서발송"] || "Y", 주문서약칭: row["주문서약칭"] || name }); count++; continue; }
    const p = {}; for (const h of CONFIG.base.hdr) if (h !== "업체" && !sameField(h, row[h], b[h])) { p[h] = row[h]; count++; }
    if (Object.keys(p).length) patch.vendors[name] = p;
  }
  for (const [k, v] of ED.settings) { const b = CONFIG.base.settings.find((s) => s[0] === k); if (!b || b[1] !== v) { patch.settings[k] = v; count++; } }
  return { patch, count };
}
function bindEditor() {
  const tbl = $("edTbl");
  tbl.addEventListener("input", (e) => {
    const t = e.target; if (t.dataset.r == null) return;
    const r = ED.rows[+t.dataset.r], f = t.dataset.f;
    r[f] = t.type === "checkbox" ? (t.checked ? "Y" : "N") : t.value; r._touched = true;
    const td = t.closest("td"); td.classList.toggle("chg", edChanged(r, f));
    if (f === "받는사람") td.classList.toggle("miss", !splitList(r[f]).length);
    if (f === "활성") t.closest("tr").classList.toggle("off", !t.checked);
    edCountUpdate();
  });
  tbl.addEventListener("click", (e) => { const b = e.target.closest("[data-del]"); if (b) { ED.rows.splice(+b.dataset.del, 1); renderEditor(); } });
  $("edSet").addEventListener("input", (e) => { const t = e.target; if (t.dataset.s == null) return; const st = ED.settings[+t.dataset.s]; st[1] = t.value.replace(/\r/g, ""); const b = CONFIG.base.settings.find((x) => x[0] === st[0]); t.classList.toggle("chg", !b || b[1] !== st[1]); edCountUpdate(); });
  document.querySelectorAll(".edTab").forEach((b) => (b.onclick = () => edShowTab(b.dataset.edtab)));
  if (typeof bindMailText === "function") bindMailText();
  $("edFilter").addEventListener("input", renderEditor); $("edNoMail").addEventListener("change", renderEditor);
  $("edAdd").onclick = () => { ED.rows.push({ ...Object.fromEntries(CONFIG.base.hdr.map((h) => [h, ""])), 활성: "Y", 엑셀첨부: "N", 발주서발송: "Y", _new: true, _touched: true }); $("edFilter").value = ""; renderEditor(); const ins = $("edTbl").querySelectorAll('input[data-f="업체"]'); if (ins.length) ins[ins.length - 1].focus(); };
  const save = async (dl) => {
    const bad = ED.rows.filter((r) => clean(r["업체"]) && splitList(r["받는사람"]).some((m) => !parseEmails(m).length));
    if (bad.length && !confirm(`메일 주소 모양이 이상한 업체가 있습니다: ${bad.map((r) => r["업체"]).join(", ")}\n그래도 저장할까요?`)) return;
    const { patch, count } = editorPatch();
    setLocal(CONFIG_NAME, { ov: { vendors: {}, newVendors: [] }, patch });
    await reloadConfig({ refetch: false });
    $("editCard").style.display = "none";
    if (dl && CONFIG.pending.length) downloadConfig();
    else toast(count ? "이 PC에 저장했습니다. 다른 PC에서도 쓰려면 위 노란 알림의 '설정 파일 받아서 올리기'를 누르세요." : "바뀐 내용이 없습니다.");
    if (typeof RESULT !== "undefined" && RESULT.mails.length) toast("설정을 고쳤습니다. 미리보기 · 메일 만들기를 다시 눌러 주세요.");
  };
  $("edSave").onclick = () => save(false); $("edSaveDl").onclick = () => save(true);
  $("edCancel").onclick = () => { $("editCard").style.display = "none"; };
}
function discardLocal() {
  if (!confirm("이 PC에만 저장된 설정 변경을 모두 버리고, 올려 둔 설정 파일대로 되돌릴까요?")) return;
  setLocal(CONFIG_NAME, { ov: { vendors: {}, newVendors: [] }, patch: { vendors: {}, newRows: [], settings: {} } });
  reloadConfig({ refetch: false });
}
