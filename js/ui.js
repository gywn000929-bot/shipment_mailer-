/* 결과 화면 · 중복 방지 기록 · 처음 한 번 확인 · 버튼 연결 */
"use strict";
// ---------------------------------------------------------------- 받은 메일 파일 기록 (중복 발송 방지)
// 기록 한 건: {keys:[...], vendor, subject, at(받은 시각), sent(Outlook에서 보냈다고 표시), who}
// keys: 출하 요청 = "ship|기준 주 월요일|업체", 긴급 자재 = "urgent|…", 주문서 = "po|PO번호"(PO마다 하나)
const LS_HIST = "mailmaker_hist_v2";
function histAll() {
  let h = store.get(LS_HIST, null);
  if (!h) {   // 예전 기록(제목|업체 → 날짜) 옮기기
    h = Object.entries(store.get("mailmaker_sent", {})).map(([k, d]) => ({ keys: ["old|" + k], subject: k.split("|")[0], vendor: k.split("|")[1] || "", at: d, sent: false }));
    store.set(LS_HIST, h);
  }
  return h;
}
const mailKeys = (m) => [...(m.keys || []), `old|${m.subject}|${m.vendor}`];
const histFor = (m) => { const ks = mailKeys(m); return histAll().filter((r) => r.keys.some((k) => ks.includes(k))); };
function recordDownload(mails) {
  const h = histAll(), at = new Date().toISOString();
  for (const m of mails) { h.push({ keys: mailKeys(m), vendor: m.vendor, subject: m.subject, at, sent: false, who: senderName() }); m.got = true; }
  store.set(LS_HIST, h.slice(-600)); renderHist();
}
function setSent(m, sent) {
  const h = histAll(), ks = mailKeys(m); let last = null;
  for (const r of h) if (r.keys.some((k) => ks.includes(k))) last = r;
  if (last) { last.sent = sent; store.set(LS_HIST, h); renderHist(); }
}
function dupText(m) {   // 예전에 받은 기록 설명
  const recs = m.dupRecs || []; if (!recs.length) return "";
  const last = recs[recs.length - 1];
  const pos = m.kind === "po" ? m.pos.filter((p) => recs.some((r) => r.keys.includes("po|" + p))) : [];
  return `${pos.length ? "PO " + pos.join(", ") + " · " : ""}${koStamp(last.at)}에 이미 메일 파일을 받았습니다 (${last.who && last.who !== senderName() ? last.who + " 이름으로 · " : ""}${recs.some((r) => r.sent) ? "보냈음 표시 있음" : "보냈음 표시 없음"})`;
}
function renderHist() {
  const h = histAll().slice().reverse().slice(0, 80);
  $("histCount").textContent = h.length ? `${histAll().length}건` : "없음";
  $("histTbl").innerHTML = h.length ? "<tr><th>받은 때</th><th>업체</th><th>제목</th><th>보냈음</th></tr>" + h.map((r) => `<tr><td>${esc(koStamp(r.at))}</td><td>${esc(r.vendor)}</td><td>${esc(r.subject)}</td><td>${r.sent ? "✔" : "-"}</td></tr>`).join("") : '<tr><td class="hint">아직 기록이 없습니다.</td></tr>';
}

// ---------------------------------------------------------------- 결과 표시
function renderResult() {
  const box = $("mails"); box.innerHTML = "";
  if (!RESULT._dupChecked) {   // 처음 한 번: 예전에 이미 받은 메일은 기본으로 빼 둔다
    for (const m of RESULT.mails) { m.dupRecs = histFor(m); if (m.dupRecs.length) m.include = false; }
    RESULT._dupChecked = true;
  }
  const ready = RESULT.mails.filter((m) => !m.noTo), noTo = RESULT.mails.filter((m) => m.noTo);
  const dup = ready.filter((m) => m.dupRecs.length), copyMode = pref("check", CONFIG_NAME) === "copy";
  $("resultCard").style.display = RESULT.mails.length || RESULT.unmatched.length ? "" : "none";
  $("btnZip").style.display = ready.length ? "" : "none";
  $("resultTitle").textContent = ready.length ? `만들어진 메일 ${ready.length}통` : "만들어진 메일 없음";
  $("resultHint").innerHTML = copyMode
    ? '이 PC는 <b>복사해서 보내기</b> 방식입니다. 메일마다 <b>① 새 메일 창 열기 → ② 본문 복사 후 붙여넣기(Ctrl+V) → ③ 첨부 받아서 끌어다 놓기</b> 순서로 하고, 확인 후 보내세요. (자동 발송 없음)'
    : '받은 <b>.eml 파일을 더블클릭</b>하면 Outlook에 받는 사람·제목·본문·첨부가 채워진 <b>새 메일 창</b>이 열립니다. 내용 확인 후 <b>보내기</b>만 누르세요. (자동 발송 없음) · 보내지 않을 업체는 왼쪽 "포함" 체크를 끄세요.';
  const parts = [];
  if (ready.length) parts.push(`<span class="ok">✅ 메일 ${ready.length}통 준비됨</span> → 오른쪽 위 <b>📦 전부 받기</b>를 누르고, 받은 .eml 파일을 더블클릭해 Outlook에서 <b>보내기</b>`);
  if (dup.length) parts.push(`<span class="warn">⚠ 이미 만든 적이 있는 메일 ${dup.length}통</span> (${esc(dup.map((m) => m.vendor).join(", "))}) → 두 번 보내지 않도록 <b>"포함"을 꺼 두었습니다.</b> 정말 다시 보내야 하면 Outlook 보낸 편지함을 확인한 뒤 "포함"을 켜세요.`);
  if (noTo.length) parts.push(`<span class="warn">⚠ 받는 사람이 없는 업체 ${noTo.length}곳</span> (${esc(noTo.map((m) => m.vendor).join(", "))}) → 아래 해당 칸에 메일 주소를 적고 <b>저장</b>하면 다음부터 기억합니다`);
  if (RESULT.unmatched.length) parts.push(`<span class="warn">⚠ 업체를 모르는 품목 ${RESULT.unmatched.reduce((a, b) => a + b.count, 0)}건</span> → 아래 <b>③ 미지정 업체 지정하기</b>에서 업체를 고른 뒤 다시 ①`);
  if (RESULT.skipped && RESULT.skipped.length) parts.push(`<span class="hint">건너뜀(메일을 보내지 않는 업체): ${esc(RESULT.skipped.join(", "))}</span>`);
  if (RESULT.poTable && RESULT.poTable.length) {
    const cls = (st) => st === "메일 생성" ? "ok" : st.includes("건너뜀") ? "hint" : "warn";
    parts.push(`<details><summary>PDF에서 읽은 주문서 ${RESULT.poTable.length}건 보기</summary><table><tr><th>PO</th><th>쪽수</th><th>PDF 거래처명</th><th>매칭 업체</th><th>처리</th></tr>${RESULT.poTable.map((r) => `<tr><td>${esc(r.po)}</td><td>${r.pages}</td><td>${esc(r.pdfVendor)}</td><td>${esc(r.vendor || "-")}</td><td class="${cls(r.status)}">${esc(r.status)}</td></tr>`).join("")}</table></details>`);
  }
  $("summary").innerHTML = parts.map((p) => `<div>${p}</div>`).join("");
  RESULT.mails.forEach((m, i) => {
    const d = document.createElement("div"); d.className = "mail" + (m.noTo ? " noto" : m.dupRecs.length ? " dup" : "");
    const alt = `<div class="alt" style="display:${copyMode && !m.noTo ? "" : "none"}"><b>복사해서 보내기</b>
       <a class="btn sm" href="${esc(mailtoUrl(m))}" data-i="${i}" data-act="mailto">① 새 메일 창 열기</a>
       <button class="sm" data-i="${i}" data-act="copy">② 본문 복사</button>
       ${m.attachments.map((a, j) => `<button class="sm" data-i="${i}" data-j="${j}" data-act="att">③ 첨부 받기: ${esc(a.name)}</button>`).join("")}
       <div class="hint">①로 열린 새 메일 창의 본문 칸을 누르고 Ctrl+V → 받은 첨부 파일을 메일 창에 끌어다 놓기 → 확인 후 보내기</div></div>`;
    d.innerHTML = `<div class="mh"><label class="inc"><input type="checkbox" data-i="${i}" data-act="inc" ${m.include ? "checked" : ""} ${m.noTo ? "disabled" : ""}> 포함</label><b>${esc(m.vendor)}</b>${m.kind === "urgent" ? '<span class="tag red">긴급 자재</span>' : ""}<span class="subj">${esc(m.subject)}</span>
      ${m.noTo ? "" : `<button class="sm pri" data-i="${i}" data-act="eml">📩 메일 파일 받기</button><button class="sm" data-i="${i}" data-act="alt" title=".eml 파일이 Outlook 새 메일 창으로 열리지 않을 때">다른 방법</button>`}<button class="sm" data-i="${i}" data-act="view">본문 보기</button></div>
      <div class="meta">${m.noTo ? `<span class="warn">받는 사람 없음</span> → <input type="text" class="toin" data-i="${i}" placeholder="메일 주소 (여러 개면 쉼표)" style="width:300px"> <button class="sm pri" data-i="${i}" data-act="saveto">저장</button>` : "받는 사람: " + esc(m.to.join(", "))}${m.cc.length ? "<br>참조: " + esc(m.cc.join(", ")) : ""}${m.attachments.length ? "<br>첨부: " + esc(m.attachments.map((a) => a.name).join(", ")) : ""}${m.note && !m.noTo ? `<br><span class="warn">${esc(m.note)}</span>` : ""}${m.dupRecs.length ? `<br><span class="warn">⚠ ${esc(dupText(m))}</span>` : ""}
      ${m.got || m.dupRecs.length ? `<br><label class="sent"><input type="checkbox" data-i="${i}" data-act="sent" ${histFor(m).some((r) => r.sent) ? "checked" : ""}> Outlook에서 보냈음</label>${m.got ? ' <span class="ok">✔ 방금 받음</span>' : ""}` : ""}</div>
      ${m.noTo ? "" : alt}<iframe class="prev" style="display:none"></iframe>`;
    box.appendChild(d);
  });
  box.onclick = onMailClick;
  renderAssign();
}
function onMailClick(e) {
  const b = e.target.closest("button,input[type=checkbox],a.btn"); if (!b || b.dataset.i == null) return;
  const box = $("mails"), m = RESULT.mails[+b.dataset.i], act = b.dataset.act;
  if (act === "inc") { m.include = b.checked; return; }
  if (act === "sent") { setSent(m, b.checked); return; }
  if (act === "mailto") return;   // 링크 그대로 열림
  if (act === "eml") {
    if (m.dupRecs.length && !m.got && !confirm(`${m.vendor}: ${dupText(m)}.\n두 번 보내는 것이 아닌지 Outlook 보낸 편지함을 확인하셨나요? 그래도 받을까요?`)) return;
    download(buildEml(m), emlName(m)); recordDownload([m]); renderResult();
  } else if (act === "alt") { const a = b.closest(".mail").querySelector(".alt"); a.style.display = a.style.display === "none" ? "" : "none"; }
  else if (act === "copy") { copyMailBody(m).then((ok) => toast(ok ? `${m.vendor} 본문을 복사했습니다. 새 메일 창 본문 칸에서 Ctrl+V 하세요.` : "복사하지 못했습니다. '본문 보기'를 눌러 직접 드래그해서 복사하세요.")); if (!m.got) { recordDownload([m]); renderResult(); } }
  else if (act === "att") { const a = m.attachments[+b.dataset.j]; download(new Blob([a.data], { type: "application/octet-stream" }), a.name); }
  else if (act === "saveto") {
    const emails = parseEmails(box.querySelector(`.toin[data-i="${b.dataset.i}"]`).value);
    if (!emails.length) { toast("메일 주소를 적어 주세요 (예: name@company.co.kr)"); return; }
    const local = getLocal(CONFIG_NAME), t = local.ov.vendors[m.vendor] = local.ov.vendors[m.vendor] || { to: [], makers: [], poNames: [] };
    t.to = uniq([...(t.to || []), ...emails]); setLocal(CONFIG_NAME, local);
    m.to = emails; m.noTo = false; m.include = !m.dupRecs.length; log(`[설정 반영 · 이 PC에 저장] ${m.vendor}: 받는사람 + ${emails.join(",")}`, "ok");
    reloadConfig({ refetch: false }).then(renderResult); toast(`${m.vendor} 받는 사람을 저장했습니다`);
  } else if (act === "view") { const f = b.closest(".mail").querySelector("iframe"); if (f.style.display === "none") { f.srcdoc = `<meta charset="utf-8"><div style="padding:8px">${mailHtml(m)}</div>`; f.style.display = ""; } else f.style.display = "none"; }
}
async function downloadZip() {
  const sel = RESULT.mails.filter((m) => m.include && !m.noTo);
  if (!sel.length) { toast("포함된 메일이 없습니다 (왼쪽 '포함' 체크 확인)"); return; }
  const again = sel.filter((m) => m.dupRecs.length && !m.got);
  if (again.length && !confirm(`이미 만든 적이 있는 메일 ${again.length}통이 포함되어 있습니다:\n${again.map((m) => " - " + m.vendor + ": " + dupText(m)).join("\n")}\n\n두 번 보내는 것이 아닌지 확인하셨나요? 그래도 받을까요?`)) return;
  const zip = new JSZip();
  for (const m of sel) zip.file(emlName(m), buildEml(m));
  if (RESULT.extraFiles) for (const f of RESULT.extraFiles) zip.file(f.name, f.data);
  const blob = await zip.generateAsync({ type: "blob" });
  download(blob, (RESULT.kind === "po" ? "주문서메일_" : "출하요청메일_") + fmtDate(new Date()) + ".zip");
  recordDownload(sel); renderResult();
  toast(`메일 ${sel.length}통을 ZIP으로 받았습니다. 압축을 풀고 .eml 파일을 더블클릭하면 Outlook 새 메일 창이 열립니다.`);
  $("afterZip").style.display = "";
}

// ---------------------------------------------------------------- 미지정 업체 지정
function renderAssign() {
  const card = $("assignCard"); const tb = $("assignTbl").querySelector("tbody"); tb.innerHTML = "";
  if (!RESULT.unmatched.length) { card.style.display = "none"; return; }
  const kindName = { maker: "MAKER", po: "주문서 거래처명", vendor: "리스트 업체명" };
  const opts = ["(건너뜀)", "(새 업체)", ...CONFIG.order];
  RESULT.unmatched.forEach((it, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><span class="tag">${kindName[it.kind] || ""}</span><b>${esc(it.key)}</b> <span class="hint">(${it.count}건)</span></td><td>${esc((it.samples || []).slice(0, 2).join("; "))}</td>
      <td><select id="ch${i}">${opts.map((o) => `<option>${esc(o)}</option>`).join("")}</select></td><td><input type="text" id="ex${i}" style="width:100%" placeholder="업체이름, 메일주소"></td>`;
    tb.appendChild(tr);
  });
  card.style.display = "";
}
function saveAssign() {
  const local = getLocal(CONFIG_NAME), ov = local.ov;
  const changes = [];
  const getNew = (name) => { let nv = ov.newVendors.find((x) => x.name === name); if (!nv) { nv = { name, to: [], makers: [], poNames: [] }; ov.newVendors.push(nv); changes.push("업체 추가: " + name); } return nv; };
  const slot = (name) => { const t = ov.vendors[name] = ov.vendors[name] || {}; for (const k of Object.keys(LIST_FIELDS)) t[k] = t[k] || []; return t; };
  RESULT.unmatched.forEach((it, i) => {
    const choice = $("ch" + i).value, extra = $("ex" + i).value.trim();
    if (!choice || choice === "(건너뜀)") return;
    let target, tname, emails;
    if (choice === "(새 업체)") {
      tname = extra.split(",")[0].trim() || it.key;
      target = CONFIG.eff.rows.some((r) => r["업체"] === tname) ? slot(tname) : getNew(tname);
      emails = parseEmails(extra.split(",").slice(1).join(","));
    } else { tname = choice; target = slot(choice); emails = parseEmails(extra); }
    if (emails.length) { target.to = uniq([...(target.to || []), ...emails]); changes.push(`${tname}: 받는사람 + ${emails.join(",")}`); }
    if (it.kind === "maker" && it.key !== "(MAKER 없음)") { if (!target.makers.includes(it.key)) { target.makers.push(it.key); changes.push(`${tname}: MAKER + ${it.key}`); } }
    else if (it.kind === "po" || it.kind === "vendor") { if (!target.poNames.includes(it.key)) { target.poNames.push(it.key); changes.push(`${tname}: 거래처명 + ${it.key}`); } }
  });
  setLocal(CONFIG_NAME, local);
  log("[설정 반영 · 이 PC에 저장] " + (changes.length ? changes.join(" / ") : "변경 없음"), "ok");
  toast(changes.length ? "저장했습니다. ① 미리보기를 다시 누르세요." : "변경 없음");
  reloadConfig({ refetch: false });
}

// ---------------------------------------------------------------- 처음 한 번 확인 (내 PC에서 .eml이 잘 열리는지)
const FC_MSG = {
  ok: `<b>좋습니다. 하나만 더 봐 주세요.</b> 열린 메일 맨 아래에 서명(회사 이름·전화번호)이 몇 번 보이나요?
     <div class="btns"><button class="pri sm" data-fc="sig1">한 번 보인다 (정상)</button><button class="sm" data-fc="sig2">두 번 보인다 (겹친다)</button><button class="sm" data-fc="sig0">서명이 없다</button></div>`,
  sig1: `<span class="ok">✅ 확인 끝.</span> 이제 평소대로 쓰시면 됩니다. 시험 메일 창은 보내지 말고 닫으세요. (이 안내는 위쪽 "처음 확인" 버튼으로 다시 볼 수 있습니다)`,
  sig2: `<span class="ok">✅ 서명을 넣지 않도록 바꿨습니다.</span> (Outlook이 서명을 스스로 넣어 주는 PC입니다) <b>시험 메일 파일을 다시 받아</b> 서명이 한 번만 나오는지 확인하면 끝입니다. 이 설정은 이 PC에 기억됩니다.`,
  sig0: `서명을 본문 끝에 넣도록 켜 두었습니다. 그래도 없다면 위쪽 <b>⚙ 업체·설정 고치기</b> → 아래 "서명" 칸에 서명 내용을 적고 저장한 뒤, 시험 메일을 다시 받아 보세요.`,
  read: `<b>"보내기" 버튼이 없는 경우</b> — "새 Outlook"을 쓰는 PC일 가능성이 큽니다. 아래 둘 중 하나를 고르세요.
     <ol><li><b>예전 Outlook 화면으로 바꾸기</b>: Outlook 창 오른쪽 위에 "새 Outlook" 스위치가 켜져 있으면 끕니다. Outlook이 다시 열리면 시험 파일을 다시 더블클릭해 보세요. "보내기"가 보이면 위에서 첫 번째 버튼을 누르면 됩니다.</li>
     <li><b>스위치가 없거나 바꾸기 싫으면</b>: <button class="pri sm" data-fc="copy">복사해서 보내기 방식으로 쓰기</button><br><span class="hint">메일마다 "새 메일 창 열기 → 본문 복사해서 붙여넣기 → 첨부 끌어다 놓기"를 하는 방식입니다. 손이 조금 더 가지만 어느 메일 프로그램에서도 됩니다.</span></li></ol>`,
  other: `<b>다른 프로그램이 열리거나 아무 일도 없는 경우</b> — .eml 파일을 Outlook으로 열도록 한 번만 정해 주면 됩니다.
     <ol><li>받은 시험 파일(.eml)을 <b>마우스 오른쪽 버튼</b>으로 누릅니다.</li><li><b>연결 프로그램</b> → <b>다른 앱 선택</b>을 누릅니다.</li><li>목록에서 <b>Outlook</b>을 고르고 <b>항상</b>을 누릅니다. <span class="hint">("Outlook (classic)"이 보이면 그것을 고르세요)</span></li><li>파일을 다시 더블클릭합니다. 새 메일 창이 열리면 위에서 맞는 버튼을 다시 눌러 주세요.</li></ol>
     목록에 Outlook이 없거나 그래도 안 되면: <button class="sm" data-fc="copy">복사해서 보내기 방식으로 쓰기</button>`,
  copy: `<span class="ok">✅ 이 PC는 "복사해서 보내기" 방식으로 바꿨습니다.</span> 메일을 만들면 메일마다 ① 새 메일 창 열기 ② 본문 복사 ③ 첨부 받기 버튼이 나옵니다. 나중에 .eml이 잘 열리게 되면 여기서 첫 번째 버튼을 다시 누르면 원래 방식으로 돌아갑니다.`,
};
function fcShow(key) { $("fcResult").innerHTML = FC_MSG[key] || ""; $("fcResult").style.display = key ? "" : "none"; }
function onFirstCheck(e) {
  const b = e.target.closest("[data-fc]"); if (!b) return;
  const k = b.dataset.fc;
  if (k === "test") { const m = testMail(); download(buildEml(m), "시험메일.eml"); toast("시험메일.eml 을 받았습니다. 더블클릭해 보세요."); return; }
  if (k === "hide") { $("firstCard").style.display = "none"; pref("fcSeen", CONFIG_NAME, true); return; }
  if (k === "ok") { if (pref("check", CONFIG_NAME) === "copy") pref("check", CONFIG_NAME, ""); }
  if (k === "sig1") pref("check", CONFIG_NAME, "ok");
  if (k === "sig2") { pref("sigOff", CONFIG_NAME, true); pref("check", CONFIG_NAME, "ok"); }
  if (k === "sig0") pref("sigOff", CONFIG_NAME, false);
  if (k === "copy") pref("check", CONFIG_NAME, "copy");
  fcShow(k); syncPrefsUi();
  if (RESULT.mails.length) renderResult();
}
function syncPrefsUi() {
  $("sigOn").checked = sigOn();
  const st = pref("check", CONFIG_NAME);
  $("btnFirst").textContent = st === "ok" ? "처음 확인 ✔" : st === "copy" ? "처음 확인 (복사 방식)" : "🔰 처음 확인";
}
function onConfigLoaded() {
  syncPrefsUi(); renderHist();
  if (!pref("check", CONFIG_NAME) && !pref("fcSeen", CONFIG_NAME)) $("firstCard").style.display = "";
}

// ---------------------------------------------------------------- UI 연결
function bindDrop(zone, input, nameEl) {
  const z = $(zone), i = $(input);
  const show = () => { const fs = [...i.files]; $(nameEl).textContent = fs.length ? fs.map((f) => f.name).join(", ") : "아직 없음"; z.classList.toggle("has", fs.length > 0); };
  z.addEventListener("click", () => i.click());
  z.addEventListener("dragover", (e) => { e.preventDefault(); z.classList.add("over"); });
  z.addEventListener("dragleave", () => z.classList.remove("over"));
  z.addEventListener("drop", (e) => { e.preventDefault(); z.classList.remove("over"); i.files = e.dataTransfer.files; show(); i.dispatchEvent(new Event("change")); });
  i.addEventListener("change", show);
}
window.addEventListener("DOMContentLoaded", () => {
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  bindDrop("dropShip", "ship", "shipName"); bindDrop("dropAir", "air", "airName"); bindDrop("dropPdf", "pdf", "pdfName");
  $("ship").addEventListener("change", () => fillSheetSelect($("ship"), "shipSheet", ["ITEM CD", "MAKER"], [], ["ETD"], []));
  $("air").addEventListener("change", () => fillSheetSelect($("air"), "airSheet", ["Item cd", "Maker"], ["Remain", "Purchase request"], [], ["Remain"]));
  $("monday").value = fmtDate(nextMonday());
  if (new Date().getDay() === 1) {   // 월요일에 열면 기준 주가 "다음 주"로 잡혀 있으므로 알려 준다
    $("monHint").style.display = "";
    $("btnThisWeek").onclick = () => { $("monday").value = fmtDate(new Date()); $("monHint").style.display = "none"; toast("기준 주를 이번 주(오늘)로 바꿨습니다."); };
  }
  { // 설정 파일 기본값: 주소의 ?who=kim → 계장님, 아니면 지난번 선택
    const q = new URLSearchParams(location.search).get("who"); let last = null; try { last = localStorage.getItem("mailmaker_cfg"); } catch {}
    const want = q === "kim" ? "config_kim.xlsx" : q === "nam" ? "config.xlsx" : last;
    if (want && [...$("config").options].some((o) => o.value === want)) $("config").value = want;
  }
  $("advToggle").onclick = () => { const a = $("adv"); a.style.display = a.style.display === "none" ? "" : "none"; $("advToggle").textContent = (a.style.display === "none" ? "▸" : "▾") + " 고급 설정"; };
  $("config").addEventListener("change", () => { $("configFile").style.display = $("config").value === "__file__" ? "" : "none"; if ($("config").value !== "__file__") reloadConfig(); });
  $("configFile").addEventListener("change", () => reloadConfig());
  $("btnShip").onclick = runShip; $("btnPo").onclick = runPo; $("btnZip").onclick = downloadZip;
  $("btnAssign").onclick = saveAssign;
  $("btnDlCfg").onclick = downloadConfig; $("btnDiscard").onclick = discardLocal;
  $("btnEdit").onclick = () => openEditor(false); $("btnEditNoMail").onclick = () => openEditor(true);
  bindEditor();
  $("btnHelp").onclick = () => $("helpCard").scrollIntoView({ behavior: "smooth" });
  $("btnFirst").onclick = () => { $("firstCard").style.display = ""; fcShow(""); $("firstCard").scrollIntoView({ behavior: "smooth" }); };
  $("firstCard").addEventListener("click", onFirstCheck);
  $("sigOn").addEventListener("change", () => { pref("sigOff", CONFIG_NAME, !$("sigOn").checked); if (RESULT.mails.length) renderResult(); toast($("sigOn").checked ? "본문 끝에 서명을 넣습니다." : "본문 끝에 서명을 넣지 않습니다."); });
  $("btnHistClear").onclick = () => { if (confirm("받은 메일 파일 기록을 모두 지울까요? (지우면 이미 만든 메일인지 알려 주지 못합니다)")) { store.set(LS_HIST, []); store.set("mailmaker_sent", {}); renderHist(); } };
  reloadConfig();
});
