# -*- coding: utf-8 -*-
"""
출하 요청 / 주문서 메일 작성기 - 창(GUI) 버전
더블클릭(0_run_app.bat)으로 열고, 파일 고르고, 버튼만 누르면 됩니다.
Python 기본 tkinter 사용 → 추가 설치 없음.
"""
import os
import sys
import subprocess
import threading
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

FROZEN = getattr(sys, "frozen", False) or os.environ.get("APP_INPROC") == "1"
HERE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
PY = sys.executable
if PY.lower().endswith("pythonw.exe"):
    PY = PY[:-len("pythonw.exe")] + "python.exe"


def next_monday():
    t = dt.date.today()
    return t + dt.timedelta(days=(7 - t.weekday()) % 7 or 7)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("대아홀딩스 메일 작성기")
        self.geometry("860x640")
        self.minsize(760, 520)
        self.proc = None
        self._build()

    # ---------------------------------------------------------------- UI
    def _build(self):
        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self); top.pack(fill="x", **pad)
        ttk.Label(top, text="설정 파일").grid(row=0, column=0, sticky="w")
        self.cfg = tk.StringVar(value=self._default_config())
        ttk.Entry(top, textvariable=self.cfg, width=70).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(top, text="찾기", command=lambda: self._pick(self.cfg, [("Excel", "*.xlsx")])).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        nb = ttk.Notebook(self); nb.pack(fill="x", **pad)

        # --- 탭 1: 출하 요청
        f1 = ttk.Frame(nb); nb.add(f1, text="  출하 요청 메일 (선박/에어)  ")
        self.ship = tk.StringVar(); self.air = tk.StringVar()
        ttk.Label(f1, text="선박리스트").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.ship, width=70).grid(row=0, column=1, sticky="we")
        ttk.Button(f1, text="찾기", command=lambda: self._pick(self.ship, [("Excel", "*.xlsx *.xls")])).grid(row=0, column=2, padx=4)
        ttk.Label(f1, text="에어리스트").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.air, width=70).grid(row=1, column=1, sticky="we")
        ttk.Button(f1, text="찾기", command=lambda: self._pick(self.air, [("Excel", "*.xlsx *.xls")])).grid(row=1, column=2, padx=4)

        opt = ttk.Frame(f1); opt.grid(row=2, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(opt, text="기준 주 월요일").pack(side="left")
        self.monday = tk.StringVar(value=next_monday().isoformat())
        ttk.Entry(opt, textvariable=self.monday, width=12).pack(side="left", padx=(4, 12))
        ttk.Label(opt, text="ETD 열 개수").pack(side="left")
        self.etd = tk.StringVar(value="")
        ttk.Entry(opt, textvariable=self.etd, width=4).pack(side="left", padx=(4, 12))
        ttk.Label(opt, text="(비우면 설정값)").pack(side="left")
        ttk.Label(opt, text="   특정 업체만").pack(side="left")
        self.only = tk.StringVar()
        ttk.Entry(opt, textvariable=self.only, width=24).pack(side="left", padx=4)

        btn = ttk.Frame(f1); btn.grid(row=3, column=0, columnspan=3, sticky="w", **pad)
        ttk.Button(btn, text="① 미리보기 (Outlook 안 씀)", command=lambda: self.run_ship(dry=True)).pack(side="left", padx=4)
        ttk.Button(btn, text="② Outlook 임시 보관함에 초안 만들기", command=lambda: self.run_ship(dry=False)).pack(side="left", padx=4)
        f1.columnconfigure(1, weight=1)

        # --- 탭 2: 주문서
        f2 = ttk.Frame(nb); nb.add(f2, text="  주문서(발주서) 메일  ")
        self.pdfs = []
        ttk.Label(f2, text="주문서 PDF").grid(row=0, column=0, sticky="nw", **pad)
        self.pdf_list = tk.Listbox(f2, height=4)
        self.pdf_list.grid(row=0, column=1, sticky="we", pady=4)
        pb = ttk.Frame(f2); pb.grid(row=0, column=2, sticky="n", padx=4)
        ttk.Button(pb, text="추가", command=self._add_pdfs).pack(fill="x")
        ttk.Button(pb, text="비우기", command=self._clear_pdfs).pack(fill="x", pady=2)
        btn2 = ttk.Frame(f2); btn2.grid(row=1, column=0, columnspan=3, sticky="w", **pad)
        ttk.Button(btn2, text="① 미리보기 (PDF 분리만)", command=lambda: self.run_po(dry=True)).pack(side="left", padx=4)
        ttk.Button(btn2, text="② Outlook 임시 보관함에 초안 만들기", command=lambda: self.run_po(dry=False)).pack(side="left", padx=4)
        f2.columnconfigure(1, weight=1)

        # --- 로그
        lf = ttk.LabelFrame(self, text="진행 내용"); lf.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(lf, wrap="word", font=("Malgun Gothic", 10))
        sb = ttk.Scrollbar(lf, command=self.log.yview); self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.log.tag_configure("err", foreground="red"); self.log.tag_configure("ok", foreground="green")

        bottom = ttk.Frame(self); bottom.pack(fill="x", **pad)
        self.btn_assign = ttk.Button(bottom, text="③ 미지정 업체 지정하기", command=self.assign_dialog, state="disabled")
        self.btn_assign.pack(side="left", padx=(0, 8))
        ttk.Button(bottom, text="Outlook에서 업체 정보 자동 수집", command=self.run_learn).pack(side="left", padx=(0, 8))
        ttk.Button(bottom, text="설정 파일 열기", command=lambda: self._open(self.cfg.get())).pack(side="left")
        ttk.Button(bottom, text="output 폴더 열기", command=lambda: self._open(os.path.join(HERE, "output"))).pack(side="left", padx=4)
        ttk.Button(bottom, text="output_po 폴더 열기", command=lambda: self._open(os.path.join(HERE, "output_po"))).pack(side="left")
        self.status = ttk.Label(bottom, text="준비됨"); self.status.pack(side="right")

    # ---------------------------------------------------------------- helpers
    def _default_config(self):
        for n in ("config.xlsx", "config_kim.xlsx"):
            p = os.path.join(HERE, n)
            if os.path.exists(p):
                return p
        return ""

    def _pick(self, var, types):
        p = filedialog.askopenfilename(filetypes=types, initialdir=HERE)
        if p:
            var.set(p)

    def _add_pdfs(self):
        for p in filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")], initialdir=HERE):
            if p not in self.pdfs:
                self.pdfs.append(p); self.pdf_list.insert("end", os.path.basename(p))

    def _clear_pdfs(self):
        self.pdfs = []; self.pdf_list.delete(0, "end")

    def _open(self, path):
        if not os.path.exists(path):
            messagebox.showinfo("안내", f"아직 없습니다:\n{path}"); return
        if sys.platform.startswith("win"):
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])

    def _write(self, text, tag=None):
        self.log.insert("end", text, tag); self.log.see("end")

    # ---------------------------------------------------------------- run
    def run_ship(self, dry):
        if not self.ship.get() and not self.air.get():
            messagebox.showwarning("확인", "선박리스트나 에어리스트 중 하나는 골라야 합니다."); return
        args = [PY, "-u", os.path.join(HERE, "make_drafts.py"), "--config", self.cfg.get(), "-y"]
        if self.ship.get(): args += ["--ship", self.ship.get()]
        if self.air.get(): args += ["--air", self.air.get()]
        if self.monday.get().strip(): args += ["--week-monday", self.monday.get().strip()]
        if self.etd.get().strip(): args += ["--etd", self.etd.get().strip()]
        if self.only.get().strip(): args += ["--only", self.only.get().strip()]
        if dry: args.append("--dry-run")
        self._run(args, dry, "출하 요청")

    def run_po(self, dry):
        if not self.pdfs:
            messagebox.showwarning("확인", "주문서 PDF를 추가하세요."); return
        args = [PY, "-u", os.path.join(HERE, "make_po_mails.py"), "--config", self.cfg.get(), "-y"] + self.pdfs
        if dry: args.append("--dry-run")
        self._run(args, dry, "주문서")

    def _run(self, args, dry, label):
        if self.proc is not None:
            messagebox.showinfo("안내", "이미 실행 중입니다. 끝날 때까지 기다려 주세요."); return
        if not dry and not messagebox.askyesno("확인", f"{label} 초안을 Outlook 임시 보관함에 만들까요?\n(미리보기를 먼저 보셨나요?)"):
            return
        self.log.delete("1.0", "end")
        self._write(("[미리보기] " if dry else "[Outlook 초안 생성] ") + label + "\n\n", "ok")
        self.status.config(text="실행 중...")
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")

        if FROZEN:
            self._run_inproc(args)
            return

        def worker():
            try:
                self.proc = subprocess.Popen(args, cwd=HERE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                             text=True, encoding="utf-8", errors="replace", env=env,
                                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                for line in self.proc.stdout:
                    tag = "err" if ("오류" in line or "Error" in line or "Traceback" in line) else None
                    self.after(0, self._write, line, tag)
                rc = self.proc.wait()
                self.after(0, self._done, rc)
            except Exception as e:
                self.after(0, self._write, f"\n실행 실패: {e}\n", "err")
                self.after(0, self._done, 1)
        threading.Thread(target=worker, daemon=True).start()

    def run_learn(self):
        if not messagebox.askyesno("확인", "Outlook 보낸 편지함(최근 6개월)을 읽어서\n비어 있는 받는사람·MAKER·거래처명을 설정 파일에 채웁니다.\n이미 값이 있는 칸은 건드리지 않습니다. 진행할까요?"):
            return
        args = [PY, "-u", os.path.join(HERE, "learn_from_outlook.py"), "--config", self.cfg.get(), "--months", "6"]
        self._run(args, True, "업체 정보 자동 수집")

    def _unmatched_items(self):
        import json
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

    def assign_dialog(self):
        from config_tools import VendorSheet
        items = self._unmatched_items()
        if not items:
            messagebox.showinfo("안내", "지정할 미지정 항목이 없습니다."); return
        vs = VendorSheet(self.cfg.get())
        names = vs.names()
        win = tk.Toplevel(self); win.title("미지정 업체 지정"); win.grab_set()
        ttk.Label(win, text="각 항목을 어느 업체로 보낼지 고르세요. 새 업체면 '(새 업체)'를 고르고 이름을 적으세요.").grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=6)
        ttk.Label(win, text="항목").grid(row=1, column=0, sticky="w", padx=8)
        ttk.Label(win, text="예시 품목").grid(row=1, column=1, sticky="w", padx=8)
        ttk.Label(win, text="업체").grid(row=1, column=2, sticky="w", padx=8)
        ttk.Label(win, text="새 업체 이름 / 메일").grid(row=1, column=3, sticky="w", padx=8)
        rows = []
        for i, it in enumerate(items, start=2):
            label = {"maker": "MAKER: ", "po": "주문서 거래처명: ", "vendor": "리스트 업체명: "}.get(it["kind"], "") + it["key"]
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="w", padx=8, pady=2)
            ttk.Label(win, text="; ".join(it["samples"][:2])[:60]).grid(row=i, column=1, sticky="w", padx=8)
            cb = ttk.Combobox(win, values=["(건너뜀)", "(새 업체)"] + names, width=18, state="readonly"); cb.current(0)
            cb.grid(row=i, column=2, padx=8)
            ent = ttk.Entry(win, width=34); ent.grid(row=i, column=3, padx=8)
            rows.append((it, cb, ent))

        def save():
            for it, cb, ent in rows:
                choice = cb.get(); extra = ent.get().strip()
                if choice == "(건너뜀)":
                    continue
                if choice == "(새 업체)":
                    name = extra.split(",")[0].strip() or it["key"]
                    r = vs.find_by_name(name) or vs.add_vendor(name)
                    emails = [x.strip() for x in extra.split(",")[1:] if "@" in x]
                    if emails:
                        vs.add_emails(r, emails)
                else:
                    r = vs.find_by_name(choice)
                    if extra and "@" in extra:
                        vs.add_emails(r, [x.strip() for x in extra.split(",") if "@" in x])
                if it["kind"] == "maker" and it["key"] != "(MAKER 없음)":
                    vs.add_maker(r, it["key"])
                elif it["kind"] in ("po", "vendor"):
                    vs.add_po_name(r, it["key"])
            ch = vs.save()
            self._write("\n[설정 반영]\n" + ("\n".join("  - " + c for c in ch) if ch else "  변경 없음") + "\n\n다시 ① 미리보기를 누르면 반영된 결과가 나옵니다.\n", "ok")
            win.destroy()
        ttk.Button(win, text="저장", command=save).grid(row=len(items) + 2, column=3, sticky="e", padx=8, pady=8)
        ttk.Button(win, text="취소", command=win.destroy).grid(row=len(items) + 2, column=2, sticky="e", padx=8, pady=8)

    def _run_inproc(self, args):
        """exe로 묶였을 때: 별도 python 없이 같은 프로세스 안에서 스크립트 실행."""
        import runpy, io
        script, argv = args[2], args[3:]
        app = self

        class Writer(io.TextIOBase):
            def write(self, t):
                if t:
                    tag = "err" if ("오류" in t or "Error" in t or "Traceback" in t) else None
                    app.after(0, app._write, t, tag)
                return len(t)

        def worker():
            rc = 0
            old_out, old_err, old_argv, old_cwd = sys.stdout, sys.stderr, sys.argv, os.getcwd()
            try:
                sys.stdout = sys.stderr = Writer()
                sys.argv = [script] + list(argv)
                os.chdir(HERE)
                if HERE not in sys.path:
                    sys.path.insert(0, HERE)
                runpy.run_path(script, run_name="__main__")
            except SystemExit as e:
                rc = 0 if (e.code in (None, 0)) else 1
                if e.code not in (None, 0) and not isinstance(e.code, int):
                    print(str(e.code))
            except Exception:
                import traceback; traceback.print_exc(); rc = 1
            finally:
                sys.stdout, sys.stderr, sys.argv = old_out, old_err, old_argv
                os.chdir(old_cwd)
            app.after(0, app._done, rc)
        self.proc = True
        threading.Thread(target=worker, daemon=True).start()

    def _done(self, rc):
        self.proc = None
        try:
            self.btn_assign.config(state="normal" if self._unmatched_items() else "disabled")
        except Exception:
            pass
        if rc == 0:
            self._write("\n끝났습니다.\n", "ok"); self.status.config(text="완료")
        else:
            self._write("\n오류로 끝났습니다. 위 내용을 캡처해서 보내 주세요.\n", "err"); self.status.config(text="오류")


if __name__ == "__main__" and "--web" in sys.argv:
    import web_app
    web_app.main()
    raise SystemExit
if __name__ == "__main__":
    App().mainloop()
