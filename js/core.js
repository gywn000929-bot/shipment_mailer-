/* 대아홀딩스 메일 작성기 - 브라우저 전용 버전 (서버/Python 없음)
   파일 구성 (빌드 없음, index.html 이 아래 순서로 읽는다)
     core.js   공통 도구 · 로그 · 이 PC(브라우저) 저장소
     config.js 설정 엑셀 읽기 · 이 PC에만 있는 변경 · 설정 고치기 화면 · 설정 내려받기
     lists.js  선박/에어 리스트 엑셀 읽기
     mail.js   메일 본문 · 첨부 엑셀 · .eml 파일 · 복사해서 보내기
     run.js    출하 요청 / 긴급 자재 / 주문서 메일 만들기
     ui.js     결과 화면 · 중복 방지 기록 · 처음 한 번 확인 · 버튼 연결 */
"use strict";
const $ = (id) => document.getElementById(id);
const YELLOW = "FFFF00";
const WEEKDAY_KO = "월화수목금토일";
const NO_SET = ["N", "NO", "X", "0"];

// ---------------------------------------------------------------- 유틸
const norm = (s) => (s == null ? "" : String(s).replace(/\s+/g, "").toLowerCase());
const clean = (v) => {
  if (v == null) return "";
  if (v instanceof Date) return isNaN(v) ? "" : fmtCellDate(dayOf(v));
  const s = String(v).trim();
  return s.toLowerCase() === "nan" ? "" : s;
};
const isNo = (v) => NO_SET.includes(clean(v).toUpperCase());
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
// 엑셀 날짜 칸은 시간대에 따라 자정에서 몇 초 어긋나게 읽힐 수 있어, 날짜만 쓸 때는 1시간 더해 그날로 맞춘다
const dayOf = (d) => new Date(d.getTime() + 3600000);
const fmtCellDate = (d) => `${pad2(d.getDate())}-${MON[d.getMonth()]}-${String(d.getFullYear()).slice(2)}`;   // 엑셀에 보이는 형태
const parseISO = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const weekdayKo = (d) => WEEKDAY_KO[(d.getDay() + 6) % 7];
const koDate = (d) => `${d.getMonth() + 1}월 ${d.getDate()}일(${weekdayKo(d)})`;
const shortDate = (d) => `${d.getMonth() + 1}/${d.getDate()}일(${weekdayKo(d)})`;
const koStamp = (iso) => { const d = new Date(iso); return isNaN(d) ? String(iso) : `${d.getMonth() + 1}/${d.getDate()}(${weekdayKo(d)})` + (String(iso).length > 10 ? ` ${pad2(d.getHours())}:${pad2(d.getMinutes())}` : ""); };
const safeName = (s) => s.replace(/[\\/:*?"<>|]/g, "_");
const normName = (s) => clean(s).replace(/\(주\)|주식회사|㈜|\(\s*\d+\s*\)|\[[^\]]*\]/g, "").replace(/[\s()\-_.]/g, "").toLowerCase();
const uniq = (a) => [...new Set(a)];
const parseEmails = (s) => String(s || "").split(/[,;\s]+/).map((x) => x.trim()).filter((x) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));

function nextMonday(today = new Date()) {
  const wd = (today.getDay() + 6) % 7; // 월=0
  return addDays(today, (7 - wd) % 7 || 7);
}
function weekLabel(monday) {
  const thu = addDays(monday, 3);
  return [thu.getMonth() + 1, Math.ceil(thu.getDate() / 7)];
}
// "02-Dec-26" 같은 글자 또는 날짜 셀 → Date (아니면 null)
function toDate(v) {
  if (v instanceof Date) { if (isNaN(v)) return null; const d = dayOf(v); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  const m = clean(v).match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (!m) return null;
  const mi = MON.findIndex((x) => x.toLowerCase() === m[2].toLowerCase());
  if (mi < 0) return null;
  return new Date(m[3].length === 2 ? 2000 + +m[3] : +m[3], mi, +m[1]);
}

// ---------------------------------------------------------------- 이 PC(브라우저) 저장소
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
};
const LS_PREFS = "mailmaker_prefs";
function pref(group, name, value) {   // pref("sigOff","config_kim.xlsx") 읽기 / 세 번째 값을 주면 쓰기
  const p = store.get(LS_PREFS, {});
  if (value === undefined) return (p[group] || {})[name];
  (p[group] = p[group] || {})[name] = value; store.set(LS_PREFS, p);
}

// ---------------------------------------------------------------- 로그 · 알림
function log(t, k = "info") {
  const d = document.createElement("div"); d.className = k; d.textContent = t;
  $("log").appendChild(d); $("log").scrollTop = $("log").scrollHeight;
}
function clearLog() { $("log").innerHTML = ""; }
let toastTimer = null;
function toast(m) { const t = $("toast"); t.textContent = m; t.classList.add("show"); clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), 3500); }
function setStatus(t, run) { $("status").textContent = t; $("status").className = "status" + (run ? " run" : ""); }
function download(blob, name) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
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
