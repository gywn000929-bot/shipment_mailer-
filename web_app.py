# -*- coding: utf-8 -*-
"""
브라우저(HTML) 화면 버전.  더블클릭(0_run_web.bat) → 브라우저가 열림 → 파일 끌어다 놓고 버튼.
- 뒤에서 이 PC 안에서만 도는 작은 서버(127.0.0.1)가 make_drafts / make_po_mails 를 실행
- 파일은 브라우저가 업로드 → 서버가 폴더(uploads)에 저장 → 처리
"""
import io, json, os, re, runpy, sys, threading, time, webbrowser, socket, datetime as dt
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
UPLOAD = os.path.join(HERE, "uploads")
STATE = {"log": [], "running": False, "rc": None, "unmatched": [], "seq": 0}
LOCK = threading.Lock()


def log(text, kind="info"):
    with LOCK:
        STATE["seq"] += 1
        STATE["log"].append({"i": STATE["seq"], "t": text, "k": kind})
        STATE["log"] = STATE["log"][-2000:]


def next_monday():
    t = dt.date.today()
    return (t + dt.timedelta(days=(7 - t.weekday()) % 7 or 7)).isoformat()


def list_configs():
    return sorted(f for f in os.listdir(HERE) if f.lower().startswith("config") and f.lower().endswith(".xlsx"))


# ---------------------------------------------------------------- 스크립트 실행 (같은 프로세스 안에서)
class _Writer(io.TextIOBase):
    def __init__(self):
        self.buf = ""
    def write(self, t):
        self.buf += t
        while "\n" in self.buf:
            line, self.buf = self.buf.split("\n", 1)
            kind = "err" if re.search(r"오류|Error|Traceback", line) else ("warn" if re.search(r"경고|문제|미지정|못 찾", line) else "info")
            log(line, kind)
        return len(t)
    def flush(self):
        if self.buf:
            log(self.buf); self.buf = ""


def run_script(script, argv, label):
    if STATE["running"]:
        return False
    STATE.update(running=True, rc=None, log=[], unmatched=[])
    log(f"[{label}] 시작", "head")

    def worker():
        rc = 0
        old = (sys.stdout, sys.stderr, sys.argv, os.getcwd())
        w = _Writer()
        try:
            sys.stdout = sys.stderr = w
            sys.argv = [script] + list(argv)
            os.chdir(HERE)
            if HERE not in sys.path:
                sys.path.insert(0, HERE)
            runpy.run_path(os.path.join(HERE, script), run_name="__main__")
        except SystemExit as e:
            rc = 0 if e.code in (None, 0) else 1
            if e.code not in (None, 0) and not isinstance(e.code, int):
                print(str(e.code))
        except Exception:
            import traceback; traceback.print_exc(); rc = 1
        finally:
            w.flush()
            sys.stdout, sys.stderr, sys.argv = old[0], old[1], old[2]
            os.chdir(old[3])
        STATE["unmatched"] = load_unmatched()
        log("끝났습니다." if rc == 0 else "오류로 끝났습니다. 위 내용을 확인하세요.", "ok" if rc == 0 else "err")
        STATE.update(running=False, rc=rc)
    threading.Thread(target=worker, daemon=True).start()
    return True


def load_unmatched():
    items = []
    for d, fn, kind in ((os.path.join(HERE, "output"), "unmatched_ship.json", "maker"),
                        (os.path.join(HERE, "output_po"), "unmatched_po.json", "po")):
        f = os.path.join(d, fn)
        if os.path.exists(f):
            try:
                with open(f, encoding="utf-8") as fh:
                    for it in json.load(fh):
                        it.setdefault("kind", kind); items.append(it)
            except Exception:
                pass
    return items


# ---------------------------------------------------------------- 업로드 파싱 (multipart)
def parse_multipart(body, ctype):
    m = re.search(r"boundary=(.+)", ctype)
    if not m:
        return {}, {}
    boundary = m.group(1).strip('"').encode()
    fields, files = {}, {}
    for part in body.split(b"--" + boundary):
        if b"\r\n\r\n" not in part:
            continue
        head, data = part.split(b"\r\n\r\n", 1)
        data = data.rstrip(b"\r\n")
        if data.endswith(b"--"):
            data = data[:-2].rstrip(b"\r\n")
        h = head.decode("utf-8", "replace")
        nm = re.search(r'name="([^"]*)"', h)
        fnm = re.search(r'filename="([^"]*)"', h)
        if not nm:
            continue
        if fnm and fnm.group(1):
            files.setdefault(nm.group(1), []).append((fnm.group(1), data))
        else:
            fields[nm.group(1)] = data.decode("utf-8", "replace")
    return fields, files


def save_upload(name, data):
    os.makedirs(UPLOAD, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', "_", os.path.basename(name))
    p = os.path.join(UPLOAD, safe)
    with open(p, "wb") as f:
        f.write(data)
    return p


# ---------------------------------------------------------------- HTTP
class H(BaseHTTPRequestHandler):
    def log_message(self, *a):  # 콘솔 조용히
        pass

    def _json(self, obj, code=200):
        b = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code); self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query)
        if u.path == "/":
            b = INDEX_HTML.encode()
            self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)
        elif u.path == "/api/state":
            since = int(q.get("since", ["0"])[0])
            with LOCK:
                lines = [l for l in STATE["log"] if l["i"] > since]
            self._json({"running": STATE["running"], "rc": STATE["rc"], "log": lines,
                        "unmatched": STATE["unmatched"], "configs": list_configs(), "monday": next_monday()})
        elif u.path == "/api/vendors":
            try:
                from config_tools import VendorSheet
                self._json({"names": VendorSheet(os.path.join(HERE, q.get("config", ["config.xlsx"])[0])).names()})
            except Exception as e:
                self._json({"names": [], "error": str(e)})
        elif u.path == "/api/open":
            what = q.get("what", [""])[0]
            path = {"output": os.path.join(HERE, "output"), "output_po": os.path.join(HERE, "output_po"),
                    "config": os.path.join(HERE, q.get("config", ["config.xlsx"])[0]), "folder": HERE}.get(what, HERE)
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)
                self._json({"ok": True})
            except Exception as e:
                self._json({"ok": False, "error": str(e)})
        else:
            self.send_error(404)

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(n)
        ctype = self.headers.get("Content-Type", "")
        if ctype.startswith("multipart/form-data"):
            fields, files = parse_multipart(body, ctype)
        else:
            fields, files = (json.loads(body.decode() or "{}"), {})
        if u.path == "/api/run_ship":
            ship = save_upload(*files["ship"][0]) if files.get("ship") else fields.get("ship_path", "")
            air = save_upload(*files["air"][0]) if files.get("air") else fields.get("air_path", "")
            if not ship and not air:
                return self._json({"ok": False, "error": "선박리스트나 에어리스트 중 하나는 넣어야 합니다."})
            args = ["--config", os.path.join(HERE, fields.get("config", "config.xlsx")), "-y"]
            if ship: args += ["--ship", ship]
            if air: args += ["--air", air]
            if fields.get("monday"): args += ["--week-monday", fields["monday"]]
            if fields.get("etd"): args += ["--etd", fields["etd"]]
            if fields.get("only"): args += ["--only", fields["only"]]
            if fields.get("mode") != "real": args.append("--dry-run")
            ok = run_script("make_drafts.py", args, "출하 요청 " + ("미리보기" if "--dry-run" in args else "Outlook 초안 생성"))
            self._json({"ok": ok, "error": None if ok else "이미 실행 중입니다."})
        elif u.path == "/api/run_po":
            pdfs = [save_upload(n_, d) for n_, d in files.get("pdf", [])]
            if not pdfs:
                return self._json({"ok": False, "error": "주문서 PDF를 넣어야 합니다."})
            args = ["--config", os.path.join(HERE, fields.get("config", "config.xlsx")), "-y"] + pdfs
            if fields.get("mode") != "real": args.append("--dry-run")
            ok = run_script("make_po_mails.py", args, "주문서 " + ("미리보기" if "--dry-run" in args else "Outlook 초안 생성"))
            self._json({"ok": ok, "error": None if ok else "이미 실행 중입니다."})
        elif u.path == "/api/learn":
            ok = run_script("learn_from_outlook.py", ["--config", os.path.join(HERE, fields.get("config", "config.xlsx")), "--months", "6"], "Outlook 업체 정보 자동 수집")
            self._json({"ok": ok})
        elif u.path == "/api/quit":
            self._json({"ok": True})
            threading.Thread(target=lambda: (time.sleep(0.3), os._exit(0)), daemon=True).start()
        elif u.path == "/api/assign":
            try:
                from config_tools import VendorSheet
                vs = VendorSheet(os.path.join(HERE, fields.get("config", "config.xlsx")))
                for a in fields.get("items", []):
                    choice, extra, it = a.get("choice", ""), (a.get("extra") or "").strip(), a.get("item", {})
                    if not choice or choice == "(건너뜀)":
                        continue
                    if choice == "(새 업체)":
                        name = extra.split(",")[0].strip() or it["key"]
                        r = vs.find_by_name(name) or vs.add_vendor(name)
                        emails = [x.strip() for x in extra.split(",")[1:] if "@" in x]
                        if emails: vs.add_emails(r, emails)
                    else:
                        r = vs.find_by_name(choice)
                        if extra and "@" in extra:
                            vs.add_emails(r, [x.strip() for x in extra.split(",") if "@" in x])
                    if it.get("kind") == "maker" and it["key"] != "(MAKER 없음)":
                        vs.add_maker(r, it["key"])
                    elif it.get("kind") in ("po", "vendor"):
                        vs.add_po_name(r, it["key"])
                ch = vs.save()
                STATE["unmatched"] = []
                log("[설정 반영] " + (" / ".join(ch) if ch else "변경 없음"), "ok")
                self._json({"ok": True, "changes": ch})
            except Exception as e:
                self._json({"ok": False, "error": str(e)})
        else:
            self.send_error(404)


INDEX_HTML = r"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>대아홀딩스 메일 작성기</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
:root{--bg:#f4f6fa;--card:#fff;--line:#e3e7ee;--txt:#1f2937;--mut:#6b7280;--pri:#2563eb;--pri2:#1d4ed8;--ok:#16a34a;--warn:#d97706;--err:#dc2626;--chip:#eef2ff}
*{box-sizing:border-box}body{margin:0;font-family:"Malgun Gothic","맑은 고딕",system-ui,sans-serif;background:var(--bg);color:var(--txt);font-size:14px}
header{background:#fff;border-bottom:1px solid var(--line);padding:14px 24px;display:flex;align-items:center;gap:16px}
header h1{font-size:18px;margin:0}header .sp{flex:1}
select,input[type=text],input[type=date]{border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:14px;background:#fff}
main{max-width:1100px;margin:20px auto;padding:0 16px;display:grid;grid-template-columns:1fr 1fr;gap:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px;box-shadow:0 1px 2px rgba(0,0,0,.03)}
.card h2{margin:0 0 12px;font-size:16px;display:flex;align-items:center;gap:8px}
.step{display:inline-flex;width:22px;height:22px;border-radius:50%;background:var(--pri);color:#fff;font-size:12px;align-items:center;justify-content:center}
.drop{border:2px dashed #c7d2fe;border-radius:12px;padding:16px;text-align:center;color:var(--mut);background:#fafbff;cursor:pointer;transition:.15s}
.drop.over{background:#eef2ff;border-color:var(--pri)}.drop.has{border-style:solid;border-color:var(--ok);background:#f0fdf4;color:var(--txt)}
.drop b{color:var(--txt)}.drop small{display:block;margin-top:4px}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px}.row label{color:var(--mut);font-size:13px}
.btns{display:flex;gap:8px;margin-top:14px;flex-wrap:wrap}
button{border:0;border-radius:10px;padding:10px 14px;font-size:14px;cursor:pointer;background:#e5e7eb;color:var(--txt)}
button.pri{background:var(--pri);color:#fff}button.pri:hover{background:var(--pri2)}button.ok{background:var(--ok);color:#fff}
button:disabled{opacity:.5;cursor:not-allowed}
.full{grid-column:1/-1}
#log{background:#0f172a;color:#e2e8f0;border-radius:12px;padding:14px;height:340px;overflow:auto;font-family:Consolas,"D2Coding",monospace;font-size:13px;white-space:pre-wrap;line-height:1.5}
#log .head{color:#93c5fd;font-weight:bold}#log .ok{color:#86efac}#log .warn{color:#fcd34d}#log .err{color:#fca5a5}
.status{font-size:13px;color:var(--mut)}.status.run{color:var(--pri)}
table{width:100%;border-collapse:collapse;margin-top:8px}th,td{text-align:left;padding:8px 6px;border-bottom:1px solid var(--line);font-size:13px}
.tag{display:inline-block;background:var(--chip);color:var(--pri2);border-radius:6px;padding:2px 8px;font-size:12px;margin-right:6px}
.hint{color:var(--mut);font-size:12px;margin-top:6px}
.toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:#111827;color:#fff;padding:10px 16px;border-radius:10px;opacity:0;transition:.2s}
.toast.show{opacity:1}
</style></head><body>
<header><h1>📮 대아홀딩스 메일 작성기</h1><span class="sp"></span>
<label class="status">설정 파일&nbsp;</label><select id="config"></select>
<button onclick="openIt('config')">설정 열기</button><button onclick="learn()">Outlook에서 업체 정보 수집</button><button onclick="quitApp()" title="뒤에서 도는 프로그램을 끝냅니다">종료</button></header>
<main>
<section class="card"><h2><span class="step">1</span> 출하 요청 메일 (선박 / 에어)</h2>
 <div class="drop" id="dropShip" onclick="pick('ship')"><b>선박리스트</b> 엑셀을 여기에 끌어다 놓거나 클릭<small id="shipName">아직 없음</small></div>
 <div style="height:8px"></div>
 <div class="drop" id="dropAir" onclick="pick('air')"><b>에어리스트</b> 엑셀을 여기에 끌어다 놓거나 클릭<small id="airName">아직 없음 (에어가 없는 주는 비워 두세요)</small></div>
 <input type="file" id="ship" accept=".xlsx,.xls" hidden><input type="file" id="air" accept=".xlsx,.xls" hidden>
 <div class="row"><label>기준 주 월요일</label><input type="date" id="monday"><label>ETD 열 개수</label><input type="text" id="etd" style="width:50px" placeholder="설정값"><label>특정 업체만</label><input type="text" id="only" style="width:160px" placeholder="예: 대일CST,태양단자"></div>
 <div class="btns"><button class="pri" onclick="runShip('dry')">① 미리보기</button><button class="ok" onclick="runShip('real')">② Outlook 임시 보관함에 초안 만들기</button></div>
 <div class="hint">미리보기는 Outlook을 쓰지 않습니다. 초안은 임시 보관함(Drafts)에만 생기고 자동 발송되지 않습니다.</div>
</section>
<section class="card"><h2><span class="step">1</span> 주문서(발주서) 메일</h2>
 <div class="drop" id="dropPdf" onclick="document.getElementById('pdf').click()"><b>주문서 PDF</b>를 여기에 끌어다 놓거나 클릭 (여러 개 가능)<small id="pdfName">아직 없음</small></div>
 <input type="file" id="pdf" accept=".pdf" multiple hidden>
 <div class="btns"><button class="pri" onclick="runPo('dry')">① 미리보기 (PDF 분리만)</button><button class="ok" onclick="runPo('real')">② Outlook 임시 보관함에 초안 만들기</button></div>
 <div class="hint">한 PDF에 여러 업체·여러 PO가 있어도 됩니다. 업체별로 잘라서 "업체-PO번호" 메일을 만듭니다. 한국단자공업·연호·경신은 자동 제외.</div>
 <div class="btns"><button onclick="openIt('output')">출하 결과 폴더</button><button onclick="openIt('output_po')">주문서 결과 폴더</button></div>
</section>
<section class="card full"><h2><span class="step">2</span> 진행 내용 <span class="status" id="status">준비됨</span></h2><div id="log"></div></section>
<section class="card full" id="assignCard" style="display:none"><h2><span class="step">3</span> 미지정 업체 지정하기</h2>
 <div class="hint">아래 항목을 어느 업체로 보낼지 고르고 저장하세요. 새 업체면 "(새 업체)"를 고르고 오른쪽에 <b>업체이름, 메일주소</b>를 적습니다. 저장 후 ① 미리보기를 다시 누르세요.</div>
 <table id="assignTbl"><thead><tr><th>항목</th><th>예시</th><th>업체</th><th>새 업체 이름 / 메일</th></tr></thead><tbody></tbody></table>
 <div class="btns"><button class="pri" onclick="saveAssign()">저장</button></div>
</section>
</main><div class="toast" id="toast"></div>
<script>
let since=0, unmatchedShown="";
const $=id=>document.getElementById(id);
function toast(m){const t=$('toast');t.textContent=m;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2500)}
function pick(id){$(id).click()}
function bindDrop(zone,input,nameEl,multi){
 const z=$(zone),i=$(input);
 z.addEventListener('dragover',e=>{e.preventDefault();z.classList.add('over')});
 z.addEventListener('dragleave',()=>z.classList.remove('over'));
 z.addEventListener('drop',e=>{e.preventDefault();z.classList.remove('over');i.files=e.dataTransfer.files;show()});
 i.addEventListener('change',show);
 function show(){const fs=[...i.files];$(nameEl).textContent=fs.length?fs.map(f=>f.name).join(', '):'아직 없음';z.classList.toggle('has',fs.length>0)}
}
bindDrop('dropShip','ship','shipName');bindDrop('dropAir','air','airName');bindDrop('dropPdf','pdf','pdfName',true);
async function post(url,fd){const r=await fetch(url,{method:'POST',body:fd});return r.json()}
function cfg(){return $('config').value}
async function runShip(mode){
 const fd=new FormData();fd.append('config',cfg());fd.append('mode',mode);fd.append('monday',$('monday').value);fd.append('etd',$('etd').value);fd.append('only',$('only').value);
 if($('ship').files[0])fd.append('ship',$('ship').files[0]);if($('air').files[0])fd.append('air',$('air').files[0]);
 if(mode==='real'&&!confirm('출하 요청 초안을 Outlook 임시 보관함에 만들까요?\n(미리보기를 먼저 보셨나요?)'))return;
 start();const r=await post('/api/run_ship',fd);if(!r.ok){toast(r.error);$('status').textContent='준비됨'}
}
async function runPo(mode){
 const fd=new FormData();fd.append('config',cfg());fd.append('mode',mode);[...$('pdf').files].forEach(f=>fd.append('pdf',f));
 if(mode==='real'&&!confirm('주문서 초안을 Outlook 임시 보관함에 만들까요?'))return;
 start();const r=await post('/api/run_po',fd);if(!r.ok){toast(r.error);$('status').textContent='준비됨'}
}
async function learn(){if(!confirm('Outlook 보낸 편지함(최근 6개월)을 읽어 비어 있는 받는사람·MAKER·거래처명을 설정에 채웁니다.\n이미 값이 있는 칸은 건드리지 않습니다. 진행할까요?'))return;
 start();await post('/api/learn',new Blob([JSON.stringify({config:cfg()})],{type:'application/json'}))}
function start(){$('log').innerHTML='';since=0;unmatchedShown="";$('assignCard').style.display='none';$('status').textContent='실행 중...';$('status').className='status run'}
function quitApp(){if(!confirm('메일 작성기를 종료할까요? (브라우저 탭은 직접 닫으세요)'))return;fetch('/api/quit',{method:'POST'});document.body.innerHTML='<p style="padding:40px;font-size:18px">종료되었습니다. 이 탭을 닫으세요.</p>'}
function openIt(what){fetch('/api/open?what='+what+'&config='+encodeURIComponent(cfg()))}
async function poll(){
 try{const r=await fetch('/api/state?since='+since);const s=await r.json();
  if(!$('config').options.length){s.configs.forEach(c=>{const o=document.createElement('option');o.value=o.textContent=c;$('config').appendChild(o)});if(!$('monday').value)$('monday').value=s.monday}
  const L=$('log');for(const l of s.log){const d=document.createElement('div');d.className=l.k;d.textContent=l.t;L.appendChild(d);since=l.i}
  if(s.log.length)L.scrollTop=L.scrollHeight;
  if(!s.running&&$('status').classList.contains('run')){$('status').textContent=s.rc===0?'완료':'오류';$('status').className='status'}
  const key=JSON.stringify(s.unmatched);if(!s.running&&s.unmatched.length&&key!==unmatchedShown){unmatchedShown=key;showAssign(s.unmatched)}
 }catch(e){}
 setTimeout(poll,700)
}
async function showAssign(items){
 const v=await (await fetch('/api/vendors?config='+encodeURIComponent(cfg()))).json();
 const tb=$('assignTbl').querySelector('tbody');tb.innerHTML='';
 const kindName={maker:'MAKER',po:'주문서 거래처명',vendor:'리스트 업체명'};
 items.forEach((it,idx)=>{const tr=document.createElement('tr');
  tr.innerHTML=`<td><span class="tag">${kindName[it.kind]||''}</span><b>${it.key}</b> <span class="hint">(${it.count}건)</span></td><td>${(it.samples||[]).slice(0,2).join('; ')}</td>
  <td><select id="ch${idx}"><option>(건너뜀)</option><option>(새 업체)</option>${v.names.map(n=>`<option>${n}</option>`).join('')}</select></td><td><input type="text" id="ex${idx}" style="width:100%" placeholder="업체이름, 메일주소"></td>`;
  tb.appendChild(tr)});
 $('assignCard').dataset.items=JSON.stringify(items);$('assignCard').style.display='block';$('assignCard').scrollIntoView({behavior:'smooth'})
}
async function saveAssign(){
 const items=JSON.parse($('assignCard').dataset.items||'[]');
 const payload={config:cfg(),items:items.map((it,i)=>({item:it,choice:$('ch'+i).value,extra:$('ex'+i).value}))};
 const r=await post('/api/assign',new Blob([JSON.stringify(payload)],{type:'application/json'}));
 if(r.ok){toast('설정에 저장했습니다. ① 미리보기를 다시 누르세요.');$('assignCard').style.display='none'}else toast('저장 실패: '+r.error)
}
poll();
</script></body></html>"""


def main():
    port = 8765
    with socket.socket() as s:
        while True:
            try:
                s.bind(("127.0.0.1", port)); break
            except OSError:
                port += 1
    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    url = f"http://127.0.0.1:{port}/"
    print("메일 작성기 화면:", url, "(이 창은 닫지 마세요. 끝나면 닫으면 됩니다)")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
