# -*- coding: utf-8 -*-
"""
주문서(발주서) 메일 자동 작성기
- ERP에서 한 번에 뽑은 주문서 PDF(여러 업체, 여러 PO가 한 파일) -> 업체별로 PO PDF를 나눠 Outlook 초안 생성
- 제목: 대아 주문서-48010,48011   첨부: 48010.pdf, 48011.pdf
- config.xlsx 업체 시트의 '거래처명(주문서)' 열로 업체를 찾고, '발주서발송'=N 인 업체(한국단자공업, 연호전자, 경신전선 등)는 건너뜀

사용법:
    python make_po_mails.py 주문서.pdf             # 미리보기 후 Outlook 초안 생성
    python make_po_mails.py 주문서.pdf --dry-run   # Outlook 없이 PDF 분리 + HTML 미리보기만
"""
import argparse
import html
import os
import re
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG = os.path.join(HERE, "config.xlsx")


def clean(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def split_list(s):
    return [x.strip() for x in re.split(r"[,;/\n]", clean(s)) if x.strip()]


def norm_name(s):
    """거래처명 비교용: (주), 주식회사, 공백, 괄호 제거."""
    s = clean(s)
    s = re.sub(r"\(\s*주\s*\)|주식회사|㈜|\(\s*\d+\s*\)|\[[^\]]*\]", " ", s)
    s = re.sub(r"(^|[\s(])주(?=[\s)]|$)", r"\1", s)  # PDF에서 띄어져 읽힌 "( 주 반도기전 )", "반도기전 주 김해지점"
    return re.sub(r"[\s()\-_.]", "", s).lower()


# ---------------------------------------------------------------- 설정
def load_config(path):
    xl = pd.ExcelFile(path)
    df = pd.read_excel(xl, "업체", dtype=object).fillna("")
    settings_df = pd.read_excel(xl, "설정", dtype=object, header=None).fillna("")
    settings = {clean(r[0]): clean(r[1]) for _, r in settings_df.iterrows() if clean(r[0])}
    vendors = []
    for _, r in df.iterrows():
        name = clean(r.get("업체"))
        if not name or name.startswith("※"):
            continue
        vendors.append({
            "name": name,
            "to": split_list(r.get("받는사람")),
            "cc": split_list(r.get("추가참조")),
            "recv_line": clean(r.get("수신줄")),
            "title": clean(r.get("호칭")),
            "po_names": [norm_name(x) for x in split_list(r.get("거래처명(주문서)"))],
            "send_po": clean(r.get("발주서발송", "Y")).upper() not in ("N", "NO", "X", "0"),
            "po_extra": clean(r.get("주문서추가문구")),
            "po_short": clean(r.get("주문서약칭")) or name,
        })
    return vendors, settings


def find_vendor(po_vendor_name, vendors):
    key = norm_name(po_vendor_name)
    if not key:
        return None
    names = lambda v: [x for x in dict.fromkeys([norm_name(v["name"]), norm_name(v["po_short"])] + v["po_names"]) if x]
    for v in vendors:                      # 1) 이름이 정확히 같음
        if key in names(v):
            return v
    # 2) 등록된 이름(4글자 이상)이 거래처명 안에 들어 있음 - 가장 긴 것, 한 업체로 정해질 때만
    #    (짧은 이름 "엠이"가 "에스엠이코리아"를 가로채 메일이 엉뚱한 업체로 가지 않게)
    best, tie = None, False
    for v in vendors:
        for a in names(v):
            if len(a) < 4 or not key.startswith(a):
                continue
            if best is None or len(a) > len(best[0]):
                best, tie = (a, v), False
            elif len(a) == len(best[0]) and v is not best[1]:
                tie = True
    return best[1] if best and not tie else None


# ---------------------------------------------------------------- PDF
def read_pages(pdf_path):
    """페이지별 (거래처명, 거래처코드, PO번호) 추출."""
    try:
        import pdfplumber
    except ImportError:
        raise SystemExit("pdfplumber가 없습니다. 설치: pip install pdfplumber")
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, pg in enumerate(pdf.pages):
            words = pg.extract_words(keep_blank_chars=True, x_tolerance=2)

            def right_of(label, n=3):
                w = [x for x in words if x["text"].replace(" ", "") == label]
                if not w:
                    return []
                y, x0 = w[0]["top"], w[0]["x1"]
                same = sorted([x for x in words if abs(x["top"] - y) < 3 and x["x0"] > x0], key=lambda x: x["x0"])
                return [x["text"].strip() for x in same[:n]]

            vend = right_of("거래처명")
            po = right_of("발주번호", 1)
            name = vend[0] if vend else ""
            code = ""
            m = re.search(r"\((\d+)\s*\)", " ".join(vend[1:2]))
            if m:
                code = m.group(1)
            po_no = ""
            if po:
                m = re.search(r"PO-?0*(\d+)", po[0])
                if m:
                    po_no = m.group(1)
            if not po_no:
                text = pg.extract_text() or ""
                m = re.search(r"PO-?0*(\d+)", text)
                po_no = m.group(1) if m else ""
            pages.append({"page": i, "vendor": name, "code": code, "po": po_no})
    return pages


def write_po_pdfs(pdf_path, groups, out_dir):
    """groups: {po_no: [page_idx...]} -> out_dir/{po_no}.pdf  (pypdf 우선, 없으면 pypdfium2)"""
    paths = {}
    try:
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(pdf_path)
        for po_no, idxs in groups.items():
            w = PdfWriter()
            for i in idxs:
                w.add_page(reader.pages[i])
            p = os.path.join(out_dir, f"{po_no}.pdf")
            with open(p, "wb") as f:
                w.write(f)
            paths[po_no] = p
        return paths
    except ImportError:
        pass
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise SystemExit("PDF를 나눌 패키지가 없습니다. 설치: pip install pypdf")
    src = pdfium.PdfDocument(pdf_path)
    for po_no, idxs in groups.items():
        dst = pdfium.PdfDocument.new()
        dst.import_pages(src, pages=list(idxs))
        p = os.path.join(out_dir, f"{po_no}.pdf")
        dst.save(p)
        paths[po_no] = p
    return paths


# ---------------------------------------------------------------- 메일
def build_body(vendor, settings, with_sig):
    """설정 시트 '주문서본문' 템플릿 사용. 빈 줄은 그대로 빈 줄. {추가문구} 자리에 업체별 '주문서추가문구'가 들어감."""
    p = lambda t: f'<p style="margin:0">{t}</p>'
    br = '<p style="margin:0">&nbsp;</p>'
    tpl = settings.get("주문서본문") or ("안녕하세요. " + settings.get("발신자소개", "대아홀딩스 남효주 입니다.")
                                      + "\n수고 많으십니다.\n\n상기 제목에 따라 주문서 송부하오니\n첨부파일 확인 및 수신 여부 회신 부탁드리겠습니다.\n\n감사합니다.")
    if "{추가문구}" not in tpl and vendor["po_extra"]:
        tpl = tpl.replace("감사합니다", "{추가문구}\n\n감사합니다", 1)
    out = []
    if vendor["recv_line"]:
        out += [p(html.escape(vendor["recv_line"])), br]
    for line in tpl.replace("\r", "").split("\n"):
        if "{추가문구}" in line:
            if not vendor["po_extra"]:
                continue
            line = line.replace("{추가문구}", vendor["po_extra"])
        if "{호칭}" in line:
            line = line.replace("{호칭}", (vendor["title"] + " ") if vendor["title"] else "")
        out.append(p(html.escape(line)) if line.strip() else br)
    # 연속 빈 줄은 하나로
    dedup = []
    for x in out:
        if x == br and dedup and dedup[-1] == br:
            continue
        dedup.append(x)
    out = dedup
    if with_sig and settings.get("서명"):
        out += [br] + [p(html.escape(l)) for l in settings["서명"].split("\n")]
    size = clean(settings.get("본문크기")) or "10"
    bold = clean(settings.get("본문굵게")).upper() in ("Y", "YES", "O", "1")
    return (f'<div style="font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:{size}pt{";font-weight:bold" if bold else ""}">'
            + "".join(out) + "</div>")


def make_subject(settings, vendor, pos):
    tpl = settings.get("주문서제목형식") or "대아 주문서-{PO}"
    return tpl.replace("{PO}", ",".join(pos)).replace("{업체}", vendor["po_short"])


def att_name(settings, vendor, po):
    tpl = settings.get("주문서첨부이름") or "{PO}.pdf"
    name = tpl.replace("{PO}", po).replace("{업체}", vendor["po_short"])
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def outlook_app():
    try:
        import win32com.client
    except ImportError:
        raise SystemExit("pywin32가 없습니다. 설치: pip install pywin32   (또는 --dry-run 사용)")
    return win32com.client.Dispatch("Outlook.Application")


def existing_drafts(app, po_nos):
    ns = app.GetNamespace("MAPI")
    found = []
    for item in ns.GetDefaultFolder(16).Items:
        try:
            subj = item.Subject or ""
        except Exception:
            continue
        if any(re.search(r"(?<!\d)" + n + r"(?!\d)", subj) for n in po_nos):
            found.append((subj, item.To))
    return found


def create_draft(app, to, cc, subject, body_html, attachments, settings):
    mail = app.CreateItem(0)
    mail.To = "; ".join(to)
    mail.CC = "; ".join(cc)
    mail.Subject = subject
    try:
        _ = mail.GetInspector
    except Exception:
        pass
    sig_html = mail.HTMLBody or ""
    has_sig = len(re.sub(r"<[^>]+>|&nbsp;|\s", "", sig_html)) > 0
    if has_sig and re.search(r"<body[^>]*>", sig_html, re.I):
        mail.HTMLBody = re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + body_html, sig_html, count=1, flags=re.I)
    else:
        has_sig = False
        fallback = settings.get("서명", "")
        extra = "".join(f'<p style="margin:0">{html.escape(l)}</p>' for l in fallback.split("\n")) if fallback else ""
        mail.HTMLBody = body_html + ('<p style="margin:0">&nbsp;</p>' + extra if extra else "")
    for a in attachments:
        mail.Attachments.Add(os.path.abspath(a))
    mail.Save()
    try:
        mail.Close(0)
    except Exception:
        pass
    return has_sig


# ---------------------------------------------------------------- 메인
def main():
    ap = argparse.ArgumentParser(description="주문서 PDF -> 업체별 주문서 메일 초안 (Outlook 임시 보관함)")
    ap.add_argument("pdf", nargs="+", help="주문서 PDF (여러 개 가능)")
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--out", default=os.path.join(HERE, "output_po"), help="분리한 PDF/미리보기 저장 폴더")
    ap.add_argument("--dry-run", action="store_true", help="Outlook 없이 PDF 분리 + HTML 미리보기만")
    ap.add_argument("--yes", "-y", action="store_true")
    ap.add_argument("--only", help="이 업체만 (쉼표 구분)")
    args = ap.parse_args()

    for f in args.pdf:
        if not os.path.isfile(f):
            print(f"[오류] 파일을 찾을 수 없습니다: {f}   (현재 폴더: {os.getcwd()})")
            sys.exit(1)
    vendors, settings = load_config(args.config)
    common_cc = split_list(settings.get("공통참조", ""))
    os.makedirs(args.out, exist_ok=True)

    # PDF 읽기: PO별 페이지 묶기, PO -> 거래처
    po_pages, po_vendor, order = {}, {}, []
    pdf_paths = {}
    for f in args.pdf:
        for pg in read_pages(f):
            if not pg["po"]:
                print(f"[경고] {os.path.basename(f)} {pg['page'] + 1}페이지: 발주번호를 못 읽었습니다. 건너뜁니다.")
                continue
            key = (f, pg["po"])
            if key not in po_pages:
                po_pages[key] = []
                order.append(key)
                po_vendor[key] = (pg["vendor"], pg["code"])
            po_pages[key].append(pg["page"])
    if not po_pages:
        raise SystemExit("주문서를 하나도 읽지 못했습니다. (스캔 이미지 PDF면 텍스트가 없어 읽을 수 없습니다)")

    # PO PDF 분리
    for f in args.pdf:
        groups = {po: idxs for (ff, po), idxs in po_pages.items() if ff == f}
        pdf_paths.update(write_po_pdfs(f, groups, args.out))

    # 업체별 묶기
    plan, skipped, unknown = {}, [], []
    for key in order:
        vname, code = po_vendor[key]
        v = find_vendor(vname, vendors)
        po = key[1]
        if v is None:
            unknown.append((po, vname, code))
            continue
        if not v["send_po"]:
            skipped.append((po, vname, v["name"]))
            continue
        if args.only and v["name"] not in split_list(args.only):
            continue
        plan.setdefault(v["name"], {"vendor": v, "pos": [], "po_vendor": vname})
        plan[v["name"]]["pos"].append(po)

    print("\n==================== 미리보기 ====================")
    for name, item in plan.items():
        v = item["vendor"]
        print(f"■ {name} ({item['po_vendor']})  PO: {', '.join(item['pos'])}")
        print(f"   제목: {make_subject(settings, v, item['pos'])}")
        print(f"   받는 사람: {', '.join(v['to']) or '(없음!)'}   첨부: {', '.join(att_name(settings, v, po) for po in item['pos'])}")
    if skipped:
        print("\n[발주서 미발송 업체 - 건너뜀]")
        for po, vname, name in skipped:
            print(f"   PO {po}  {vname} → {name}")
    if unknown:
        print("\n[업체를 못 찾음 - config.xlsx 업체 시트 '거래처명(주문서)' 열에 추가하세요]")
        for po, vname, code in unknown:
            print(f"   PO {po}  거래처명: {vname}  (코드 {code})")
    print("==================================================")
    import json
    with open(os.path.join(args.out, "unmatched_po.json"), "w", encoding="utf-8") as f:
        json.dump([{"key": vname, "samples": [f"PO {po} (코드 {code})"], "count": 1} for po, vname, code in unknown], f, ensure_ascii=False)
    no_mail = [n for n, it in plan.items() if not it["vendor"]["to"]]
    if no_mail:
        print(f"[받는 사람 없음 → 초안 생성 안 함] {', '.join(no_mail)}  (config 업체 시트에 메일 주소를 채우세요)")
    plan = {k: v for k, v in plan.items() if v["vendor"]["to"]}
    if not plan:
        print("생성할 초안이 없습니다.")
        return

    app = None
    if not args.dry_run:
        app = outlook_app()
        all_pos = [po for it in plan.values() for po in it["pos"]]
        dup = existing_drafts(app, all_pos)
        if dup:
            print(f"\n[경고] 임시 보관함에 같은 PO 번호가 들어간 주문서 초안이 이미 {len(dup)}건 있습니다:")
            for subj, to in dup:
                print(f"   - {subj}  → {to}")
    if not args.yes:
        ans = input(f"\n{len(plan)}개 업체 주문서 초안을 {'HTML로 저장' if args.dry_run else 'Outlook 임시 보관함에 생성'}할까요? [y/N] ").strip().lower()
        if ans not in ("y", "yes"):
            print("취소했습니다.")
            return

    for name, item in plan.items():
        v = item["vendor"]
        subject = make_subject(settings, v, item["pos"])
        atts = []
        for po in item["pos"]:
            want = os.path.join(args.out, att_name(settings, v, po))
            if os.path.abspath(want) != os.path.abspath(pdf_paths[po]):
                import shutil
                shutil.copyfile(pdf_paths[po], want)
            atts.append(want)
        cc = list(dict.fromkeys(common_cc + v["cc"]))
        body = build_body(v, settings, with_sig=args.dry_run)
        if args.dry_run:
            hp = os.path.join(args.out, f"{subject}_{name}.html")
            head = (f"<p><b>To:</b> {html.escape('; '.join(v['to']))}<br><b>CC:</b> {html.escape('; '.join(cc))}<br>"
                    f"<b>Subject:</b> {html.escape(subject)}<br><b>첨부:</b> {html.escape(', '.join(os.path.basename(a) for a in atts))}</p><hr>")
            with open(hp, "w", encoding="utf-8") as f:
                f.write('<html><head><meta charset="utf-8"></head><body>' + head + body + "</body></html>")
            print(f"  저장: {hp}")
        else:
            has_sig = create_draft(app, v["to"], cc, subject, body, atts, settings)
            print(f"  초안 생성: {name}  {subject}  ({'Outlook 서명' if has_sig else '설정 서명'})")
    print("\n완료. Outlook 임시 보관함(Drafts)에서 확인 후 직접 보내세요." if not args.dry_run else "\n완료 (dry-run).")


if __name__ == "__main__":
    main()
