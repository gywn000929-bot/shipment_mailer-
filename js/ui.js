/* 결과 화면 · 중복 방지 기록 · 처음 한 번 확인 · 버튼 연결 */
"use strict";
// ---------------------------------------------------------------- 받은 메일 파일 기록 (중복 발송 방지)
// 기록 한 건: {keys:[...], vendor, subject, at(받은 시각), sent(Outlook에서 보냈다고 표시), who}
// keys: 출하 요청 = "ship|기준 주 월요일|업체", 긴급 자재 = "urgent|…", 주문서 = "po|PO번호"(PO마다 하나)
const LS_HIST = "mailmaker_hist_v2";
// ?solo=ord → 발주 점검만(balju.html), ?solo=int → 계장님께 메일만(gyejang.html) 보이는 전용 화면. 메일 작성기 본 화면에는 두 메일이 없다
const SOLO = ["ord", "int"].includes(new URLSearchParams(location.search).get("solo")) ? new URLSearchParams(location.search).get("solo") : "";
if (SOLO) document.documentElement.classList.add("solo-" + SOLO);
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
  // 처음 보이는 메일마다 한 번: 예전에 이미 받은 메일은 기본으로 빼 둔다 (발주 점검을 다시 만들어 나중에 더해진 메일도)
  for (const m of RESULT.mails) if (!m.dupRecs) { m.dupRecs = histFor(m); if (m.dupRecs.length) m.include = false; }
  const ready = RESULT.mails.filter((m) => !m.noTo), noTo = RESULT.mails.filter((m) => m.noTo);
  const dup = ready.filter((m) => m.dupRecs.length), copyMode = pref("check", CONFIG_NAME) === "copy";
  $("resultCard").style.display = RESULT.mails.length || RESULT.unmatched.length ? "" : "none";
  $("btnZip").style.display = ready.length ? "" : "none";
  $("resultTitle").textContent = ready.length ? `만들어진 메일 ${ready.length}통` : "만들어진 메일 없음";
  $("resultHint").innerHTML = copyMode
    ? '이 PC는 <b>복사해서 보내기</b> 방식입니다. 메일마다 <b>1. 새 메일 창 열기 → 2. 본문 복사 후 붙여넣기(Ctrl+V) → 3. 첨부 받아서 끌어다 놓기</b> 순서로 하고, 확인 후 보내세요. (자동 발송 없음)'
    : '받은 <b>.eml 파일을 더블클릭</b>하면 Outlook에 받는 사람·제목·본문·첨부가 채워진 <b>새 메일 창</b>이 열립니다. 내용 확인 후 <b>보내기</b>만 누르세요. (자동 발송 없음) · 보내지 않을 업체는 왼쪽 "포함" 체크를 끄세요.';
  const parts = [];
  if (ready.length) parts.push(`<span class="ok">✅ 메일 ${ready.length}통 준비됨</span> → 오른쪽 위 <b>📦 전부 받기</b>를 누르고, 받은 .eml 파일을 더블클릭해 Outlook에서 <b>보내기</b>`);
  if (dup.length) parts.push(`<span class="warn">⚠ 이미 만든 적이 있는 메일 ${dup.length}통</span> (${esc(dup.map((m) => m.vendor).join(", "))}) → 두 번 보내지 않도록 <b>"포함"을 꺼 두었습니다.</b> 정말 다시 보내야 하면 Outlook 보낸 편지함을 확인한 뒤 "포함"을 켜세요.`);
  if (noTo.length) parts.push(`<span class="warn">⚠ 받는 사람이 없는 업체 ${noTo.length}곳</span> (${esc(noTo.map((m) => m.vendor).join(", "))}) → 아래 해당 칸에 메일 주소를 적고 <b>저장</b>하면 다음부터 기억합니다`);
  if (RESULT.unmatched.length) parts.push(`<span class="warn">⚠ 업체를 모르는 품목 ${RESULT.unmatched.reduce((a, b) => a + b.count, 0)}건</span> → 아래 <b>미지정 업체 지정하기</b>에서 업체를 고른 뒤 다시 미리보기`);
  if (RESULT.skipped && RESULT.skipped.length) parts.push(`<span class="hint">건너뜀(메일을 보내지 않는 업체): ${esc(RESULT.skipped.join(", "))}</span>`);
  if (RESULT.poTable && RESULT.poTable.length) {
    const cls = (st) => st === "메일 생성" ? "ok" : st.includes("건너뜀") ? "hint" : "warn";
    parts.push(`<details><summary>PDF에서 읽은 주문서 ${RESULT.poTable.length}건 보기</summary><table><tr><th>PO</th><th>쪽수</th><th>PDF 거래처명</th><th>매칭 업체</th><th>처리</th></tr>${RESULT.poTable.map((r) => `<tr><td>${esc(r.po)}</td><td>${r.pages}</td><td>${esc(r.pdfVendor)}</td><td>${esc(r.vendor || "-")}</td><td class="${cls(r.status)}">${esc(r.status)}</td></tr>`).join("")}</table></details>`);
  }
  $("summary").innerHTML = parts.map((p) => `<div>${p}</div>`).join("");
  RESULT.mails.forEach((m, i) => {
    if (m._origSubject == null) m._origSubject = m.subject;
    const d = document.createElement("div"); d.className = "mail" + (m.noTo ? " noto" : m.dupRecs.length ? " dup" : "");
    const alt = `<div class="alt" style="display:${copyMode && !m.noTo ? "" : "none"}"><b>복사해서 보내기</b>
       <a class="btn sm" href="${esc(mailtoUrl(m))}" data-i="${i}" data-act="mailto">1. 새 메일 창 열기</a>
       <button class="sm" data-i="${i}" data-act="copy">2. 본문 복사</button>
       ${m.attachments.map((a, j) => `<button class="sm" data-i="${i}" data-j="${j}" data-act="att">3. 첨부 받기: ${esc(a.name)}</button>`).join("")}
       <div class="hint">1번으로 열린 새 메일 창의 본문 칸을 누르고 Ctrl+V → 받은 첨부 파일을 메일 창에 끌어다 놓기 → 확인 후 보내기</div></div>`;
    d.innerHTML = `<div class="mh"><label class="inc"><input type="checkbox" data-i="${i}" data-act="inc" ${m.include ? "checked" : ""} ${m.noTo ? "disabled" : ""}> 포함</label><b>${esc(m.vendor)}</b>${m.kind === "urgent" ? '<span class="tag red">긴급 자재</span>' : ""}<span class="subj">${esc(m.subject)}</span>
      ${m.noTo ? "" : `<button class="sm pri" data-i="${i}" data-act="eml">📩 메일 파일 받기</button><button class="sm" data-i="${i}" data-act="alt" title=".eml 파일이 Outlook 새 메일 창으로 열리지 않을 때">다른 방법</button>`}<button class="sm" data-i="${i}" data-act="view">본문 보기</button><button class="sm" data-i="${i}" data-act="edit" title="이 메일 한 통의 제목·본문을 직접 고칩니다">✏ 고치기</button>${m.customHtml != null || m.subjEdited ? `<span class="tag">✏ 고친 메일</span><button class="sm" data-i="${i}" data-act="edreset">원래대로</button>` : ""}</div>
      <div class="meta">${m.noTo ? `<span class="warn">받는 사람 없음</span> → <input type="text" class="toin" data-i="${i}" placeholder="메일 주소 (여러 개면 쉼표)" style="width:300px"> <button class="sm pri" data-i="${i}" data-act="saveto">저장</button>` : "받는 사람: " + esc(m.to.join(", "))}${m.cc.length ? "<br>참조: " + esc(m.cc.join(", ")) : ""}${m.attachments.length ? "<br>첨부: " + esc(m.attachments.map((a) => a.name).join(", ")) : ""}${m.note && !m.noTo ? `<br><span class="warn">${esc(m.note)}</span>` : ""}${m.dupRecs.length ? `<br><span class="warn">⚠ ${esc(dupText(m))}</span>` : ""}
      ${m.got || m.dupRecs.length ? `<br><label class="sent"><input type="checkbox" data-i="${i}" data-act="sent" ${histFor(m).some((r) => r.sent) ? "checked" : ""}> Outlook에서 보냈음</label>${m.got ? ' <span class="ok">✔ 방금 받음</span>' : ""}` : ""}</div>
      ${m.noTo ? "" : alt}<iframe class="prev" style="display:none"></iframe>
      <div class="editBox" style="display:none"><div class="row" style="margin-top:6px"><label>제목</label><input type="text" class="edSubj" style="flex:1;min-width:260px"></div>
       <div class="hint">아래 본문을 눌러 바로 고치세요(글자 지우기·적기, Ctrl+B 굵게). 표 안 숫자도 고칠 수 있습니다. 고친 내용은 이 메일 한 통에만 들어가고, 미리보기를 다시 누르면 없어집니다.</div>
       <iframe class="edFrame"></iframe>
       <div class="btns" style="margin-top:6px"><button class="sm pri" data-i="${i}" data-act="edsave">고친 내용 저장</button><button class="sm" data-i="${i}" data-act="edreset">원래대로</button><button class="sm" data-i="${i}" data-act="edclose">닫기</button></div></div>`;
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
  } else if (act === "edit" || act === "edclose") {
    const box2 = b.closest(".mail").querySelector(".editBox"), f = box2.querySelector("iframe");
    if (act === "edclose" || box2.style.display !== "none") { box2.style.display = "none"; return; }
    box2.querySelector(".edSubj").value = m.subject;
    f.onload = () => { try { f.contentDocument.designMode = "on"; } catch {} };
    f.srcdoc = `<meta charset="utf-8"><body style="margin:8px">${mailHtml(m)}</body>`;
    box2.style.display = "";
  } else if (act === "edsave") {
    const box2 = b.closest(".mail").querySelector(".editBox"), f = box2.querySelector("iframe"), subj = clean(box2.querySelector(".edSubj").value);
    if (subj && subj !== m.subject) { m.subject = subj; m.subjEdited = true; }
    m.customHtml = f.contentDocument.body.innerHTML;
    log(`[고침] ${m.vendor}: 이 메일 한 통의 ${m.subjEdited ? "제목·" : ""}본문을 고쳤습니다.`, "ok"); toast(`${m.vendor} 메일을 고친 내용으로 저장했습니다. 메일 파일 받기/ZIP에 고친 내용이 들어갑니다.`); renderResult();
  } else if (act === "edreset") {
    if (m._origSubject) m.subject = m._origSubject; delete m.customHtml; m.subjEdited = false; toast(`${m.vendor} 메일을 처음 만든 내용으로 되돌렸습니다.`); renderResult();
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
  download(blob, (RESULT.kind === "po" ? "주문서메일_" : RESULT.kind === "order" ? "발주확인메일_" : RESULT.fri ? "금요일에어메일_" : "출하요청메일_") + fmtDate(new Date()) + ".zip");
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
  toast(changes.length ? "저장했습니다. 미리보기를 다시 누르세요." : "변경 없음");
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
  copy: `<span class="ok">✅ 이 PC는 "복사해서 보내기" 방식으로 바꿨습니다.</span> 메일을 만들면 메일마다 1. 새 메일 창 열기 2. 본문 복사 3. 첨부 받기 버튼이 나옵니다. 나중에 .eml이 잘 열리게 되면 여기서 첫 번째 버튼을 다시 누르면 원래 방식으로 돌아갑니다.`,
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
function syncMine() {   // "내 업체만" 칸: 설정에 내 업체 목록이 있을 때만 보인다. 켠 상태는 이 PC에 기억
  const mine = myVendors(); $("mineRow").style.display = mine ? "" : "none";
  if (!mine) { $("mineOnly").checked = false; return; }
  if (new URLSearchParams(location.search).get("mine") === "1") pref("mineOnly", CONFIG_NAME, true);
  $("mineOnly").checked = !!pref("mineOnly", CONFIG_NAME);
  const exc = splitList(CONFIG.settings["내업체제외"]);
  $("mineInfo").textContent = exc.length ? `켜면 ${mine.length}곳만 만듭니다 (제외 ${exc.length}곳: ${exc.join(", ")})` : `켜면 ${mine.length}곳만 만듭니다 (${mine.join(", ")})`;
}
function onConfigLoaded() {
  syncMine();
  syncPrefsUi(); renderHist();
  if (!pref("check", CONFIG_NAME) && !pref("fcSeen", CONFIG_NAME)) $("firstCard").style.display = "";
}

// ---------------------------------------------------------------- UI 연결
// "주간 리스트 만들기" 화면에서 넘겨준 리스트가 있으면 선박리스트 칸에 넣는다 (같은 PC·같은 브라우저 안에서만)
function takeHandoff() {
  let h = null; try { h = JSON.parse(sessionStorage.getItem(HANDOFF_KEY) || localStorage.getItem(HANDOFF_KEY) || "null"); localStorage.removeItem(HANDOFF_KEY); sessionStorage.removeItem(HANDOFF_KEY); } catch {}
  if (!h || !h.b64) return;
  const bytes = Uint8Array.from(atob(h.b64), (c) => c.charCodeAt(0)), dt = new DataTransfer();
  if (window.setMode) setMode("mon", true);
  dt.items.add(new File([bytes], h.name || "선박 및 에어리스트.xlsx")); $("ship").files = dt.files; $("ship").dispatchEvent(new Event("change"));
  $("firstCard").style.display = "none"; setTimeout(() => { $("dropShip").scrollIntoView({ behavior: "smooth", block: "center" }); toast("만든 리스트를 선박리스트 칸에 넣었습니다. 기준 주를 확인하고 미리보기 · 메일 만들기를 누르세요."); }, 600);
}
// ---------------------------------------------------------------- 화면 아무 곳에나 놓기 / 붙여넣기
// 파일을 보고 메일 종류를 고른다: PDF → 주문서, 선박 표가 있는 엑셀 → 월요일, 출고수량·Remain 열만 있는 엑셀 → 수요일 REMAIN.
// "📨 계장님께"를 고른 상태(남효주 설정)에서는 그 칸으로 넣는다.
async function routeFiles(files) {
  files = [...files]; if (!files.length) return;
  const put = (id, fs) => { const dt = new DataTransfer(); fs.forEach((f) => dt.items.add(f)); $(id).files = dt.files; $(id).dispatchEvent(new Event("change")); };
  const pdfs = files.filter((f) => /\.pdf$/i.test(f.name)), xls = files.filter((f) => /\.xlsx?$/i.test(f.name));
  if (SOLO && !xls.length) { toast(`여기는 ${SOLO === "int" ? "계장님께 메일" : "발주 점검"} 화면입니다. 엑셀(.xls·.xlsx)을 놓아 주세요.`); return; }
  if (pdfs.length && !SOLO) { setMode("po", true); put("pdf", pdfs); toast(`주문서 PDF ${pdfs.length}개 → 주문서 메일을 만듭니다.`); return; }
  if (!xls.length) { toast("엑셀(.xlsx)이나 주문서 PDF를 놓아 주세요."); return; }
  const f = xls[0];
  if (SOLO === "int") { put("intFile", [f]); return; }
  // ERP 발주 미납 파일이 섞여 있으면 → 발주 점검 (리스트 엑셀은 ① 칸, 미납 파일은 ② 칸)
  const wbs = []; for (const x of xls) { try { wbs.push(await ordReadWb(x)); } catch (e) { wbs.push(null); } }
  const erpIdx = wbs.map((w, i) => (w && ordIsErp(w) ? i : -1)).filter((i) => i >= 0);
  if (erpIdx.length && !SOLO) { toast("ERP 발주 미납 파일입니다. 발주 점검은 따로 엽니다 → ☰ 더보기의 🧮 발주 점검"); return; }
  if (SOLO === "ord") {
    setMode("ord", true);
    const lists = xls.filter((_, i) => !erpIdx.includes(i)), erps = xls.filter((_, i) => erpIdx.includes(i));
    if (lists.length) put("ordList", [lists[0]]);
    if (erps.length) put("ordErp", erps);
    toast(erps.length ? `ERP 발주 미납 ${erps.length}개 → 발주 점검` : "선박리스트 → 발주 점검 ①칸"); return;
  }
  let wb = wbs[0]; if (!wb) { toast("엑셀을 읽지 못했습니다: " + f.name); return; }
  const hasShip = wb.SheetNames.some((sn) => findHeaderRow(sheetRows(wb, sn), ["ITEM CD", "MAKER"], [], ["ETD"]) >= 0);
  const hasRemain = wb.SheetNames.some((sn) => { const rows = sheetRows(wb, sn), h = findHeaderRow(rows, ["Item cd", "Maker"]); return h >= 0 && rows[h].map(norm).some((c) => c === "출고수량" || c === "remain"); });
  if (!hasShip && hasRemain) { setMode("wed", true); put("remainFile", [f]); toast("REMAIN 엑셀 → 금요일 에어 메일을 만듭니다."); }
  else { setMode("mon", true); put("ship", [f]); toast("주간 리스트 → 월요일 메일을 만듭니다."); }
}
function bindDropAnywhere() {
  let n = 0; const ov = $("dropAll");
  const isFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes("Files");
  document.addEventListener("dragenter", (e) => { if (isFiles(e)) { n++; ov.classList.add("on"); } });
  document.addEventListener("dragleave", () => { if (--n <= 0) { n = 0; ov.classList.remove("on"); } });
  document.addEventListener("dragover", (e) => { if (isFiles(e)) e.preventDefault(); });
  document.addEventListener("drop", (e) => { n = 0; ov.classList.remove("on"); if (e.defaultPrevented || !isFiles(e)) return; e.preventDefault(); routeFiles(e.dataTransfer.files); });
}
// 입력 칸 밖에서 Ctrl+V: REMAIN 표(머리글에 Item cd·Maker)면 수요일 칸에 넣고 바로 만든다
function bindPasteAnywhere() {
  document.addEventListener("paste", (e) => {
    if (SOLO) return;
    const t = e.target; if (t && (t.closest && t.closest("input,textarea,[contenteditable=true],#editCard"))) return;
    const text = (e.clipboardData || window.clipboardData).getData("text/plain") || "";
    const { aoa } = parsePastedTable(text);
    if (!(aoa.length > 1 && aoa[0].some((c) => /item\s*cd/i.test(c)) && aoa[0].some((c) => /maker/i.test(c)))) return;
    e.preventDefault(); setMode("wed", true); $("airPaste").value = text; $("airPaste").dispatchEvent(new Event("input")); $("remainFile").value = "";
    toast("붙여 넣은 REMAIN 표로 금요일 에어 메일을 만듭니다."); setTimeout(() => runShip(true), 50);
  });
}

// 새 버전 확인: 브라우저가 예전 화면을 기억해 두고 보여 주는 일이 있어서, 열 때마다 서버의 최신 index.html 의
// app-ver 를 보고 다르면 한 번 새로 불러온다(같은 버전에서 두 번 다시 불러오지 않도록 sessionStorage 로 막음).
async function checkNewVersion() {
  const mine = (document.querySelector('meta[name="app-ver"]') || {}).content; if (!mine || location.protocol === "file:") return;
  try {
    const html = await (await fetch(location.pathname + "?ver=" + Date.now(), { cache: "no-store" })).text();
    const m = html.match(/<meta name="app-ver" content="([^"]+)"/); if (!m || m[1] === mine) return;
    const key = "mailmaker_reloaded_" + m[1]; if (sessionStorage.getItem(key)) { toast("새 버전이 있습니다. Ctrl+F5 를 눌러 새로고침해 주세요."); return; }
    sessionStorage.setItem(key, "1"); location.reload();
  } catch (e) { /* 인터넷이 안 되면 그냥 지금 화면을 쓴다 */ }
}
window.addEventListener("DOMContentLoaded", () => {
  checkNewVersion();
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  bindDrop("dropShip", "ship", "shipName"); bindDrop("dropAir", "air", "airName"); bindDrop("dropPdf", "pdf", "pdfName");
  $("ship").addEventListener("change", () => fillSheetSelect($("ship"), "shipSheet", ["ITEM CD", "MAKER"], [], ["ETD"], []));
  $("air").addEventListener("change", () => fillSheetSelect($("air"), "airSheet", ["Item cd", "Maker"], ["Remain", "Purchase request"], [], AIR_REMAIN_COLS));
  $("monday").value = fmtDate(nextMonday());
  $("monday").addEventListener("input", () => { $("monday").dataset.touched = "1"; });   // 직접 바꾼 날짜는 금요일 에어에서도 그대로 쓴다
  if (new Date().getDay() === 1) {   // 월요일에 열면 기준 주가 "다음 주"로 잡혀 있으므로 알려 준다
    $("monHint").style.display = "";
    $("btnThisWeek").onclick = () => { $("monday").value = fmtDate(new Date()); $("monday").dataset.touched = "1"; $("monHint").style.display = "none"; toast("기준 주를 이번 주(오늘)로 바꿨습니다."); };
  }
  { // 설정 파일 기본값: 주소의 ?who=kim → 계장님, 아니면 지난번 선택
    const q = new URLSearchParams(location.search).get("who"); let last = null; try { last = localStorage.getItem("mailmaker_cfg"); } catch {}
    const want = q === "kim" ? "config_kim.xlsx" : q === "nam" ? "config.xlsx" : last;
    if (want && [...$("config").options].some((o) => o.value === want)) $("config").value = want;
  }
  { // 주소의 ?only=업체1,업체2 → "특정 업체만" 칸을 미리 채운다 (내 담당 업체 화면에서 넘어올 때)
    const only = new URLSearchParams(location.search).get("only");
    if (only) { $("only").value = only; $("adv").style.display = ""; $("onlyNote").style.display = ""; $("onlyNote").textContent = `지금은 ${only.split(",").length}개 업체만 만들도록 되어 있습니다 (고급 설정의 "특정 업체만"). 전체를 만들려면 그 칸을 비우세요.`; }
  }
  $("mineOnly").addEventListener("change", () => { pref("mineOnly", CONFIG_NAME, $("mineOnly").checked); toast($("mineOnly").checked ? "내 업체만 만듭니다. 미리보기 · 메일 만들기를 누르세요." : "모든 업체를 만듭니다."); if (RESULT.mails.length && RESULT.kind === "ship") toast("내 업체만 설정을 바꿨습니다. 미리보기 · 메일 만들기를 다시 누르세요."); });
  $("advToggle").onclick = () => { const a = $("adv"); a.style.display = a.style.display === "none" ? "" : "none"; $("advToggle").textContent = (a.style.display === "none" ? "▸" : "▾") + " 고급 설정"; };
  $("config").addEventListener("change", () => { $("configFile").style.display = $("config").value === "__file__" ? "" : "none"; if ($("config").value !== "__file__") reloadConfig(); });
  $("configFile").addEventListener("change", () => reloadConfig());
  const pasteInfo = () => {   // 붙여 넣은 표가 몇 줄인지 바로 보여 준다
    const t = $("airPaste").value, { aoa, vertical } = parsePastedTable(t);
    const hasHdr = aoa.length && aoa[0].some((c) => /item\s*cd/i.test(c)) && aoa[0].some((c) => /maker/i.test(c));
    const items = aoa.filter((r) => PASTE_FAC.test(r[0] || "")).length;
    $("airPaste").classList.toggle("has", !!(hasHdr && items));
    $("pasteInfo").innerHTML = !t.trim() ? "아직 붙여 넣지 않았습니다"
      : !hasHdr ? '<span class="warn">⚠ 맨 윗줄(FAC, Item cd, Maker …)이 빠졌거나 표가 아닙니다. 표의 머리글 줄부터 마지막 줄까지 다시 복사해 붙여 넣으세요.</span>'
      : `✔ 품목 <b>${items}</b>줄을 읽었습니다${vertical ? " (한 칸씩 세로로 붙여졌지만 표로 다시 맞췄습니다)" : ""}. 아래 <b>미리보기 · 메일 만들기</b>를 누르세요.`;
  };
  $("airPaste").addEventListener("input", pasteInfo);
  $("btnPasteClear").onclick = () => { $("airPaste").value = ""; pasteInfo(); toast("붙여넣은 표를 지웠습니다."); };
  $("btnPasteRun").onclick = () => runShip(true);
  bindDrop("dropRemain", "remainFile", "remainName");
  bindDrop("dropInt", "intFile", "intName"); $("btnInt").onclick = runInternal; $("intKind").addEventListener("change", () => { if ($("intFile").files.length) runInternal(); });
  // 칸에 놓으면 바로 만든다
  const auto = (id, fn) => $(id).addEventListener("change", () => { if ($(id).files.length) setTimeout(fn, 50); });
  auto("ship", () => runShip(false)); auto("remainFile", () => runShip(true)); auto("pdf", runPo); auto("intFile", runInternal);
  bindDropAnywhere(); bindPasteAnywhere(); ordBind();
  // 맨 위 "오늘 만들 메일은?" — 고른 메일에 필요한 칸만 보여 준다. 처음엔 요일로 고른다(수·목 = 수요일 REMAIN, 그 밖 = 월요일)
  const setMode = window.setMode = (m, quiet) => {
    if (SOLO) m = SOLO;   // 전용 링크(balju.html·gyejang.html): 다른 메일 칸은 안 보인다
    document.querySelectorAll(".modeBtn").forEach((b) => b.classList.toggle("on", b.dataset.mode === m));
    document.querySelectorAll("[data-modes]").forEach((el) => el.classList.toggle("mhide", !el.dataset.modes.split(" ").includes(m)));
    if (!quiet) { const c = { mon: "shipCard", wed: "pasteCard", po: "poCard", int: "intCard", ord: "ordCard" }[m]; $(c).scrollIntoView({ behavior: "smooth", block: "start" }); }
  };
  document.querySelectorAll(".modeBtn").forEach((b) => (b.onclick = () => setMode(b.dataset.mode)));
  { const d = new Date().getDay(), q = new URLSearchParams(location.search).get("mode");
    setMode(["mon", "wed", "po"].includes(q) ? q : d === 3 || d === 4 ? "wed" : "mon", true); }
  $("btnShip").onclick = () => runShip(false); $("btnPo").onclick = runPo; $("btnZip").onclick = downloadZip;
  $("btnAssign").onclick = saveAssign;
  $("btnDlCfg").onclick = downloadConfig; $("btnDiscard").onclick = discardLocal;
  $("btnEdit").onclick = () => openEditor(false, "vend"); $("btnMailText").onclick = () => openEditor(false, "mail"); $("btnEditNoMail").onclick = () => openEditor(true);
  bindEditor(); takeHandoff();
  $("lnkGuide").href = "guide.html" + location.search;
  $("btnHelp").onclick = () => { $("helpDetails").open = true; $("helpCard").scrollIntoView({ behavior: "smooth" }); };
  document.querySelectorAll(".menuBox button, .menuBox a").forEach((el) => el.addEventListener("click", () => { const m = el.closest("details"); if (m) m.open = false; }));
  $("btnFirst").onclick = () => { $("firstCard").style.display = ""; fcShow(""); $("firstCard").scrollIntoView({ behavior: "smooth" }); };
  $("firstCard").addEventListener("click", onFirstCheck);
  $("sigOn").addEventListener("change", () => { pref("sigOff", CONFIG_NAME, !$("sigOn").checked); if (RESULT.mails.length) renderResult(); toast($("sigOn").checked ? "본문 끝에 서명을 넣습니다." : "본문 끝에 서명을 넣지 않습니다."); });
  $("btnHistClear").onclick = () => { if (confirm("받은 메일 파일 기록을 모두 지울까요? (지우면 이미 만든 메일인지 알려 주지 못합니다)")) { store.set(LS_HIST, []); store.set("mailmaker_sent", {}); renderHist(); } };
  reloadConfig().then(() => {   // 주소의 ?open=mail|vend → 메일 문구 / 업체 주소 화면을 바로 연다 (계장님 할 일 페이지에서 넘어올 때)
    const o = new URLSearchParams(location.search).get("open");
    if (o === "mail" || o === "vend") openEditor(false, o);
  });
  if (SOLO) {   // 전용 화면(발주 점검 / 계장님께): 제목·번호를 바꾸고 메일 작성기로 돌아가는 링크
    $("appTitle").textContent = SOLO === "int" ? "📨 계장님께 메일" : "🧮 발주 점검";
    const st = (id, n) => { const e = document.querySelector(`#${id} .step`); if (e) e.textContent = n; };
    st(SOLO === "int" ? "intCard" : "ordCard", "1"); st("resultCard", "2");
    const w = new URLSearchParams(location.search).get("who"); $("lnkMain").href = "index.html" + (w ? "?who=" + encodeURIComponent(w) : "");
  }
  { const w = new URLSearchParams(location.search).get("who"); if ($("lnkTodo")) $("lnkTodo").href = w === "kim" ? "kim.html" : "my.html?who=" + (w || "nam"); }
});
