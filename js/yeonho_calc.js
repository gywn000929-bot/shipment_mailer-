/* 연호 발주 · 점검 계산 (엑셀 매크로 두 개를 그대로 옮김, 화면과 분리된 순수 계산)
   · yhOrder  = 매크로 '연호_발주_재고포함'  → 최종발주 표
   · yhCheck  = 매크로 '연호_입고점검'       → 최종수량대조 결과
   품명 정리 규칙·공장 구분·순서·긴급 차감 모두 매크로와 같다. */
"use strict";

// 매크로 GetPerfectCleanKey (발주 계산용)
function yhKeyP(v) {
  let s = String(v == null ? "" : v).trim().toUpperCase();
  s = s.split("(YEONHO)").join("").split("YEONHO").join("");
  s = s.split("RED").join("RE").split("GRN").join("GN").split("BLU").join("BL").split("YEL").join("YE");
  s = s.split("WHT").join("").split("WH").join("");
  return s.replace(/[ \-()[\]\/.]/g, "");
}
// 매크로 GetCleanKey (점검용, 더 적게 지움)
function yhKeyC(v) {
  let s = String(v == null ? "" : v).trim().toUpperCase();
  s = s.split("(YEONHO)").join("").split("YEONHO").join("").split("WHT").join("");
  return s.replace(/[ \-]/g, "");
}
// 매크로 GetFacCodeFinal
function yhFac(v) {
  const s = String(v == null ? "" : v).trim().toUpperCase();
  if (s.includes("긴급")) return "긴급";
  for (const f of ["F1", "F2", "F5"]) if (s.includes(f)) return f;
  return s;
}
const yhNum = (v) => { if (typeof v === "number") return isNaN(v) ? 0 : v; const n = Number(String(v == null ? "" : v).replace(/,/g, "").trim()); return isNaN(n) ? 0 : n; };
const yhAdd = (m, k, q) => m.set(k, (m.get(k) || 0) + q);

/** 발주 계산
   inp.erp   : [{name, qty, memo}]            연호 미납조회(발주 전) — memo=적요(F2(11/26), F2(긴급)…)
   inp.ship  : [{fac, code, name, qty, wait, bal}]  선박 YEONHO 줄 (qty=선박 수량, wait=본사대기, bal=BALANCE)
   inp.air   : [{fac, code, name, qty, bal}]   Air YEONHO 줄 (qty=Purchase request)
   inp.stock : null | [{fac, code, name, qty}]  재고(선택). 있으면 본사대기를 이 값으로
   inp.master: [[연호명, 대아명, 터미널0/1]]
   inp.terminal: [[CODE, 품목명, BOX]]
   opt.urgentPool: true 면 적요에 '긴급'인 줄을 공장 구분 없는 긴급 재고로(매크로 방식), false 면 적요 앞의 공장으로 */
function yhOrder(inp, opt = {}) {
  const master = new Map(), color = new Set();
  for (const [y, d, t] of inp.master) {
    if (!String(d || "").trim()) continue;
    const k = yhKeyP(d); master.set(k, String(y || "").trim());
    if (t) color.add(yhKeyP(master.get(k)));
  }
  const erp = new Map(), urgent = new Map();
  for (const r of inp.erp) {
    const nm = yhKeyP(r.name);
    let fac = yhFac(r.memo);
    if (fac === "긴급" && !opt.urgentPool) { const m = String(r.memo).toUpperCase().match(/F\d/); fac = m ? m[0] : "긴급"; }
    if (fac === "긴급") yhAdd(urgent, nm, r.qty);
    else if (fac) yhAdd(erp, fac + "|" + nm, r.qty);
  }
  const shipD = new Map(), airD = new Map(), incoming = new Map(), bal = new Map(), stock = new Map();
  const code = new Map(), shipKeys = new Map(), airOnly = new Map(), status = new Map(), rawName = new Map(), itemCd = new Map();
  const mapName = (raw) => { const k = yhKeyP(raw); return master.has(k) ? [master.get(k), "매칭완료"] : [raw, "미등록"]; };
  for (const r of inp.ship) {
    if (!String(r.name || "").trim()) continue;
    const fac = yhFac(r.fac), [yName, st] = mapName(r.name), cy = yhKeyP(yName), key = fac + "|" + cy;
    yhAdd(shipD, key, yhNum(r.qty)); yhAdd(incoming, key, yhNum(r.wait));
    if (String(r.code || "").trim()) code.set(yhKeyP(r.code), cy);
    if (yhNum(r.bal) !== 0) bal.set(key, yhNum(r.bal));
    shipKeys.set(key, true); status.set(key, st);
    if (!rawName.has(key)) rawName.set(key, r.name);
    if (!itemCd.has(key) && r.code) itemCd.set(key, String(r.code).trim());
  }
  for (const r of inp.air) {
    if (!String(r.name || "").trim()) continue;
    const fac = yhFac(r.fac), [yName, st] = mapName(r.name), cy = yhKeyP(yName), key = fac + "|" + cy;
    yhAdd(airD, key, yhNum(r.qty));
    if (String(r.code || "").trim()) { const sc = yhKeyP(r.code); if (!code.has(sc)) code.set(sc, cy); }
    if (yhNum(r.bal) !== 0) bal.set(key, yhNum(r.bal));
    if (!shipKeys.has(key)) airOnly.set(key, true);
    if (!status.has(key)) status.set(key, st);
    if (!rawName.has(key)) rawName.set(key, r.name);
    if (!itemCd.has(key) && r.code) itemCd.set(key, String(r.code).trim());
  }
  if (inp.stock) for (const r of inp.stock) {
    if (!String(r.name || "").trim()) continue;
    const fac = yhFac(r.fac), sc = yhKeyP(r.code), kr = yhKeyP(r.name);
    const cy = code.has(sc) ? code.get(sc) : master.has(kr) ? yhKeyP(master.get(kr)) : kr;
    yhAdd(stock, fac + "|" + cy, yhNum(r.qty));
  }
  const termCode = new Map(), termName = new Map();
  for (const [c, n, box] of inp.terminal || []) { if (c) termCode.set(String(c).trim().toUpperCase(), box); if (n) termName.set(yhKeyP(n), box); }
  const out = [];
  for (const key of [...shipKeys.keys(), ...airOnly.keys()]) {
    const [fac, cy] = key.split("|");
    const ship = shipD.get(key) || 0, air = airD.get(key) || 0, total = ship + air;
    const inc = inp.stock ? stock.get(key) || 0 : incoming.get(key) || 0;
    const arr = erp.get(key) || 0;
    let remain = Math.max(0, total - inc - arr), used = 0;
    if (remain > 0 && urgent.has(cy)) { const u = urgent.get(cy); used = Math.min(u, remain); urgent.set(cy, u - used); }
    const need = Math.max(0, remain - used);
    const cd = (itemCd.get(key) || "").toUpperCase();
    out.push({ key, fac, keyName: cy, name: rawName.get(key), code: itemCd.get(key) || "", status: status.get(key),
      ship, air, total, wait: inc, yh: arr, urgent: used, need, bal: bal.get(key) || 0,
      terminal: color.has(cy), box: termCode.has(cd) ? termCode.get(cd) : termName.get(cy) || 0,
      airShort: inc + arr < air });
  }
  return { rows: out, urgentLeft: [...urgent].filter(([, q]) => q > 0) };
}

/** 점검: 연호 ERP(발주 후 넣은 수량) ↔ 대아 ERP(최종발주 발주필요량)
   yh: [{fac, name, qty}]  daea: [{fac, name, qty}]  master: [[연호명, 대아명, 터미널]] */
function yhCheck(yh, daea, masterRows) {
  const m = new Map(), color = new Set();
  for (const [y, d, t] of masterRows) {
    if (!String(d || "").trim()) continue;
    const cd = yhKeyC(d), cy = yhKeyC(y); m.set(cd, cy);
    if (t) color.add(cy);
  }
  const Y = new Map(), D = new Map(), shown = new Map();
  for (const r of yh) { const k = String(r.fac).trim() + "|" + yhKeyC(r.name); yhAdd(Y, k, yhNum(r.qty)); if (!shown.has(k)) shown.set(k, r.name); }
  for (const r of daea) {
    const cd = yhKeyC(r.name), y = m.has(cd) ? m.get(cd) : cd, k = String(r.fac).trim() + "|" + y;
    yhAdd(D, k, yhNum(r.qty)); if (!shown.has(k)) shown.set(k, r.name);
  }
  const keys = [...new Set([...Y.keys(), ...D.keys()])];
  return keys.map((k) => {
    const [fac, part] = k.split("|"), qy = Y.get(k) || 0, qd = D.get(k) || 0, diff = qy - qd;
    return { fac, part, name: shown.get(k), yh: qy, daea: qd, diff, bad: Math.abs(diff) > 0.0001, terminal: color.has(part), onlyY: !D.has(k), onlyD: !Y.has(k) };
  });
}

if (typeof module !== "undefined") module.exports = { yhKeyP, yhKeyC, yhFac, yhOrder, yhCheck };
