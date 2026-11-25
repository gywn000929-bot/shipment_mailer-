/* 메일 문구 고치기 — Outlook 새 메일 창처럼 보이는 편집 화면
   받는 사람(자동) · 참조(공통참조) · 제목 · 본문 · 서명을 그 자리에서 고친다.
   자동으로 채워지는 부분(인사, 표, 날짜, 업체 이름 …)은 회색 칸으로 보이고, 지우거나 옮길 수만 있다.
   저장하면 설정 시트의 같은 칸(제목형식, 출하요청본문 …)에 {인사} 같은 자리 표시로 바꿔 넣는다. */
"use strict";

// 어떤 메일을 어떤 설정 칸으로 고치는지
const MT_MAILS = [
  { id: "mon", label: "📅 월요일 출하 요청", subj: "제목형식", body: "출하요청본문", def: () => SHIP_TPL_DEFAULT, to: "업체마다 자동 (업체 시트의 받는사람)" },
  { id: "fri", label: "✈ 금요일 에어", subj: "금요에어제목접미", subjBase: true, body: "금요에어본문", def: () => FRI_TPL_DEFAULT, to: "업체마다 자동 (업체 시트의 받는사람)" },
  { id: "po", label: "📄 주문서", subj: "주문서제목형식", body: "주문서본문", def: () => "안녕하세요. " + (mtVal("발신자소개") || "대아홀딩스 남효주 입니다.") + "\n수고 많으십니다.\n\n상기 제목에 따라 주문서 송부하오니\n첨부파일 확인 및 수신 여부 회신 부탁드리겠습니다.\n\n{추가문구}\n\n감사합니다.", to: "업체마다 자동 (주문서 PDF의 거래처)" },
  { id: "urg", label: "🚨 KET 긴급 자재", subj: "긴급제목형식", body: "긴급본문", def: () => "수고많으십니다. " + (mtVal("발신자소개") || "") + "\n\n하기 긴급 자재 건 납기 확인 후 회신 부탁드립니다.\n\n{표}\n\n감사합니다.", to: "KET (설정 '긴급메일업체')" },
];
const MT_KEYS = ["제목형식", "금요에어제목접미", "출하요청본문", "금요에어본문", "주문서제목형식", "주문서본문", "긴급제목형식", "긴급본문", "서명", "공통참조", "본문굵게", "본문크기"];
const MT_SUBJ_DEF = { 제목형식: "대아홀딩스 {월}월 {주}째주 출하 요청의 건", 금요에어제목접미: "(금요일 에어)", 주문서제목형식: "대아 주문서-{PO}", 긴급제목형식: "대아홀딩스 {월}월 {주}째주 긴급 자재의 건" };

// 자동으로 채워지는 자리: 화면에 보일 글자와 모양(block = 한 줄 전체 차지)
function mtPh(name) {
  const [m, w] = weekLabel(nextMonday());
  const greet = (clean(mtVal("출하요청인사")) || "안녕하세요.") + " " + (mtVal("발신자소개") || "");
  const P = {
    인사: { text: greet, tip: "인사 — 눌러서 글자로 바꾸면 직접 고칠 수 있습니다", editable: true },
    안내: { text: "첨부/하기 내용으로 금주 선박 리스트 송부 드립니다.", tip: "선박 표가 본문에 있으면 '하기', 엑셀 첨부면 '첨부', 에어만 있으면 '에어 리스트'로 자동으로 바뀝니다" },
    선박: { text: "📋 선박리스트 (노란색 수량, ○/○일(수) 입고 요망) + 선박 표 또는 첨부 안내", block: true, tip: "이 자리에 선박 표가 자동으로 들어갑니다" },
    에어: { text: "✈ 에어리스트 (노란색 수량, ○/○일 입고 요망) + 에어 표", block: true, tip: "이 자리에 에어 표가 자동으로 들어갑니다" },
    표: { text: "📋 긴급 자재 표 (NO · 품목명 · 미납수량 합계 · AIR 수량 · 부족일정)", block: true, tip: "이 자리에 표가 자동으로 들어갑니다" },
    추가문구: { text: "업체별 추가 문구 (있는 업체만)", tip: "업체 시트의 '주문서추가문구'. 없는 업체는 이 줄이 빠집니다" },
    호칭: { text: "호칭", tip: "업체 시트의 '호칭' (없으면 빠짐)" },
    월: { text: String(m), tip: "월 (자동)" }, 주: { text: String(w), tip: "주 (자동)" },
    업체: { text: "업체약칭", tip: "업체 시트의 '주문서약칭' (예: 기성)" }, PO: { text: "48010", tip: "주문서 번호 (자동, 여러 개면 쉼표로 이어짐)" },
    __base: { text: "월요일 메일 제목", tip: "월요일 출하 요청 제목과 같음 (그 탭에서 고침)" },
  };
  return P[name] || { text: name, tip: "자동" };
}

// ---------------------------------------------------------------- ED.settings 읽고 쓰기
const mtVal = (k) => { const s = ED && ED.settings.find((x) => x[0] === k); return s ? s[1] : (CONFIG.settings[k] ?? ""); };
function mtSet(k, v) {
  let s = ED.settings.find((x) => x[0] === k);
  if (!s) ED.settings.push((s = [k, "", "메일 문구 고치기 화면에서 추가"]));
  if (s[1] === v) return;
  s[1] = v; edCountUpdate(); mtMarkChanged();
}
const mtBase = (k) => { const b = CONFIG.base.settings.find((s) => s[0] === k); return b ? b[1] : undefined; };

// ---------------------------------------------------------------- 글자 ↔ 편집 칸
function mtChip(name) {
  const p = mtPh(name);
  return `<span class="mtChip${p.block ? " blk" : ""}${p.editable ? " ed" : ""}" contenteditable="false" data-ph="${esc(name)}" title="${esc(p.tip)} — 지우거나 옮길 수 있습니다">${esc(p.text)}${p.editable ? ' <b class="mtPen">✎ 고치기</b>' : ""}</span>`;
}
function mtLineHtml(line) {
  return esc(line).replace(/\{([^{}\n]+)\}/g, (all, n) => mtChip(n.replace(/&amp;/g, "&")));
}
function mtToHtml(tpl) {
  return String(tpl ?? "").replace(/\r/g, "").split("\n").map((l) => (l.trim() ? `<div>${mtLineHtml(l)}</div>` : "<div><br></div>")).join("");
}
// 편집 칸 → 줄 글자 (회색 칸은 {이름}으로)
function mtFromEl(el) {
  const lines = []; let cur = "";
  const flush = () => { lines.push(cur); cur = ""; };
  const walk = (n) => {
    for (const c of n.childNodes) {
      if (c.nodeType === 3) { cur += c.nodeValue.replace(/ /g, " ").replace(/\n/g, ""); continue; }
      if (c.nodeType !== 1) continue;
      if (c.dataset && c.dataset.ph) { if (c.classList.contains("blk")) { if (cur.trim()) flush(); else cur = ""; lines.push(`{${c.dataset.ph}}`); } else cur += `{${c.dataset.ph}}`; continue; }
      if (c.tagName === "BR") { flush(); continue; }
      const block = /^(DIV|P|LI|H\d)$/.test(c.tagName);
      if (block && cur) flush();
      const before = lines.length;
      walk(c);
      if (block && (cur || lines.length === before)) flush();
    }
  };
  walk(el); if (cur) flush();
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  while (lines.length && !lines[0].trim()) lines.shift();
  return lines.map((l) => l.replace(/\s+$/, "")).join("\n");
}

// ---------------------------------------------------------------- 화면
let MT_CUR = "mon";
function renderMailText() {
  const box = $("mtBox"); if (!box || !ED) return;
  const m = MT_MAILS.find((x) => x.id === MT_CUR);
  const bold = /^(Y|YES|O|1)$/i.test(clean(mtVal("본문굵게"))), size = parseInt(mtVal("본문크기")) || 10;
  const subjTpl = m.subjBase ? "{__base}" + (mtVal(m.subj) || MT_SUBJ_DEF[m.subj]) : (mtVal(m.subj) || MT_SUBJ_DEF[m.subj]);
  const bodyTpl = clean(mtVal(m.body)) ? mtVal(m.body) : m.def();
  box.innerHTML = `
  <div class="mtTabs">${MT_MAILS.map((x) => `<button class="mtTab${x.id === MT_CUR ? " on" : ""}" data-mt="${x.id}">${esc(x.label)}</button>`).join("")}</div>
  <div class="mtWin">
   <div class="mtRow"><span class="mtLbl">받는 사람</span><span class="mtAuto">${esc(m.to)}</span></div>
   <div class="mtRow"><span class="mtLbl">참조</span><input type="text" id="mtCc" value="${esc(mtVal("공통참조"))}" placeholder="모든 메일에 들어가는 참조 (쉼표로 구분)"><span class="hint" style="margin:0">모든 메일 공통</span></div>
   <div class="mtRow"><span class="mtLbl">제목</span><div class="mtSubj" id="mtSubj" contenteditable="true" spellcheck="false">${mtLineHtml(subjTpl)}</div></div>
   <div class="mtTools"><label><input type="checkbox" id="mtBold" ${bold ? "checked" : ""}> 본문 굵게</label>
     <label>글자 크기 <select id="mtSize">${[9, 10, 10.5, 11, 12, 14].map((n) => `<option ${n == size ? "selected" : ""}>${n}</option>`).join("")}</select> pt</label>
     <span class="hint" style="margin:0">모든 메일 공통</span><span class="sp" style="flex:1"></span>
     <button class="sm" id="mtReset" title="이 메일의 제목·본문을 올려 둔 설정 파일의 문구로 되돌립니다">↺ 이 메일 원래대로</button></div>
   <div class="mtBody" id="mtBody" contenteditable="true" spellcheck="false" style="font-size:${size}pt;${bold ? "font-weight:bold" : ""}">${mtToHtml(bodyTpl)}</div>
   <div class="mtSigLbl">서명 <span class="hint" style="margin:0">(모든 메일 공통 · Outlook이 서명을 스스로 넣는 PC면 비워 두세요)</span></div>
   <div class="mtBody mtSig" id="mtSig" contenteditable="true" spellcheck="false" style="font-size:${size}pt;${bold ? "font-weight:bold" : ""}">${mtToHtml(mtVal("서명"))}</div>
  </div>
  <div class="hint">글자를 눌러 바로 고치세요. <span class="mtChip">회색 칸</span>은 메일을 만들 때 자동으로 채워지는 부분입니다(표·날짜·업체 이름 등). 필요 없으면 지우고, 다른 곳으로 끌어다 옮길 수 있습니다. 실수로 지웠으면 "↺ 이 메일 원래대로".</div>`;
  mtMarkChanged();
}
function mtMarkChanged() {
  const m = MT_MAILS.find((x) => x.id === MT_CUR), ch = (k) => mtBase(k) !== undefined ? mtVal(k) !== mtBase(k) : !!clean(mtVal(k));
  const set = (id, on) => { const e = $(id); if (e) e.classList.toggle("chg", on); };
  set("mtSubj", ch(m.subj)); set("mtBody", ch(m.body)); set("mtSig", ch("서명")); set("mtCc", ch("공통참조"));
  document.querySelectorAll(".mtTab").forEach((b) => { const x = MT_MAILS.find((y) => y.id === b.dataset.mt); b.classList.toggle("chg", ch(x.subj) || ch(x.body)); });
}
function bindMailText() {
  const box = $("mtBox"); if (!box) return;
  box.addEventListener("click", (e) => {
    const t = e.target.closest("[data-mt]"); if (t) { MT_CUR = t.dataset.mt; renderMailText(); return; }
    const chip = e.target.closest(".mtChip.ed");   // 인사 칸을 누르면 그 글자를 그대로 적은 보통 글자로 바꾼다
    if (chip) { chip.replaceWith(document.createTextNode(mtPh(chip.dataset.ph).text)); $("mtBody").dispatchEvent(new Event("input", { bubbles: true })); toast("인사 문구를 글자로 바꿨습니다. 바로 고치세요."); return; }
    if (e.target.id === "mtReset") {
      const m = MT_MAILS.find((x) => x.id === MT_CUR);
      for (const k of [m.subj, m.body]) mtSet(k, mtBase(k) ?? "");
      renderMailText(); toast("이 메일의 제목·본문을 올려 둔 설정대로 되돌렸습니다.");
    }
  });
  box.addEventListener("input", (e) => {
    const m = MT_MAILS.find((x) => x.id === MT_CUR), id = e.target.id || (e.target.closest("[id]") || {}).id;
    if (id === "mtCc") mtSet("공통참조", e.target.value);
    else if (id === "mtSubj") { let v = mtFromEl($("mtSubj")).replace(/\n/g, " "); if (m.subjBase) v = v.replace("{__base}", ""); mtSet(m.subj, v); }
    else if (id === "mtBody") mtSet(m.body, mtFromEl($("mtBody")));
    else if (id === "mtSig") mtSet("서명", mtFromEl($("mtSig")));
  });
  box.addEventListener("change", (e) => {
    if (e.target.id === "mtBold") { mtSet("본문굵게", e.target.checked ? "Y" : "N"); renderMailText(); }
    if (e.target.id === "mtSize") { mtSet("본문크기", e.target.value); renderMailText(); }
  });
  // 제목은 한 줄, 붙여넣기는 글자만(다른 곳의 서식이 딸려 오지 않게)
  box.addEventListener("keydown", (e) => { if (e.target.id === "mtSubj" && e.key === "Enter") e.preventDefault(); });
  box.addEventListener("paste", (e) => {
    if (!e.target.closest || !e.target.closest("[contenteditable=true]")) return;
    e.preventDefault(); let t = (e.clipboardData || window.clipboardData).getData("text/plain");
    if (e.target.closest("#mtSubj")) t = t.replace(/\s*\n\s*/g, " ");
    document.execCommand("insertText", false, t);
  });
}
