# -*- coding: utf-8 -*-
"""
Outlook 보낸 편지함을 훑어서 config.xlsx 업체 표를 자동으로 채운다.
  - 주문서 메일(PDF 첨부): PDF의 거래처명 → 받는사람/참조, 거래처명(주문서), 호칭, 수신줄
  - 출하 요청 메일(본문 표): 표의 MAKER 값 → 그 메일 받는 업체의 MAKER
자동으로 채우는 건 '비어 있는 칸'뿐. 이미 값이 있으면 건드리지 않고 '발견' 메모만 출력.
사용: python learn_from_outlook.py [--months 6] [--config config.xlsx] [--dry-run]
"""
import argparse, os, re, sys, tempfile, datetime as dt
from config_tools import VendorSheet, _split

HERE = os.path.dirname(os.path.abspath(__file__))
INTERNAL = ("kdaea.com",)


def internal(addr):
    return any(addr.lower().endswith("@" + d) for d in INTERNAL)


def smtp_of(recipient):
    """Outlook Recipient -> SMTP 주소 (Exchange 주소는 변환)."""
    try:
        ae = recipient.AddressEntry
        if ae.Type == "EX":
            ex = ae.GetExchangeUser()
            if ex and ex.PrimarySmtpAddress:
                return ex.PrimarySmtpAddress
        return recipient.Address
    except Exception:
        return getattr(recipient, "Address", "") or ""


def recipients(mail):
    to, cc = [], []
    for r in mail.Recipients:
        a = smtp_of(r)
        if not a or "@" not in a or internal(a):
            continue
        (to if r.Type == 1 else cc).append(a)
    return to, cc


def greeting_info(body):
    """본문 앞부분에서 '수신: ...' 줄과 호칭('OO님 안녕하세요')을 뽑는다."""
    head = (body or "")[:300]
    recv = ""
    m = re.search(r"(수신\s*[:：].*?)(?:\r?\n|$)", head)
    if m:
        recv = m.group(1).strip()
    title = ""
    m = re.search(r"(?:^|\n)\s*([가-힣A-Za-z ,]{1,20}?님)\s*안녕하세요", head)
    if m:
        title = m.group(1).strip()
        if "," in title:
            title = ""
    return recv, title


def po_vendor_from_pdf(path):
    try:
        from make_po_mails import read_pages
        pages = read_pages(path)
        names = [p["vendor"] for p in pages if p["vendor"]]
        return names[0] if names else ""
    except Exception as e:
        return ""


def makers_from_html(html_body):
    """본문 표에서 MAKER/Maker 열 값 모으기."""
    makers = set()
    for table in re.findall(r"<table.*?</table>", html_body or "", re.S | re.I):
        rows = re.findall(r"<tr.*?</tr>", table, re.S | re.I)
        if not rows:
            continue
        cells0 = [re.sub(r"<[^>]+>|&nbsp;", "", c).strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", rows[0], re.S | re.I)]
        idx = next((i for i, c in enumerate(cells0) if c.replace(" ", "").upper() == "MAKER"), None)
        if idx is None:
            continue
        for r in rows[1:]:
            cells = [re.sub(r"<[^>]+>|&nbsp;", "", c).strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", r, re.S | re.I)]
            if len(cells) > idx and cells[idx] and cells[idx] != "　":
                makers.add(cells[idx])
    return makers


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.path.join(HERE, "config.xlsx"))
    ap.add_argument("--months", type=int, default=6)
    ap.add_argument("--dry-run", action="store_true", help="설정 파일에 저장하지 않고 결과만 출력")
    args = ap.parse_args()
    try:
        import win32com.client
    except ImportError:
        sys.exit("pywin32가 없습니다. 이 기능은 Outlook이 있는 PC에서만 됩니다.")
    app = win32com.client.Dispatch("Outlook.Application")
    sent = app.GetNamespace("MAPI").GetDefaultFolder(5)  # olFolderSentMail
    items = sent.Items
    items.Sort("[SentOn]", True)
    since = dt.datetime.now() - dt.timedelta(days=30 * args.months)
    items = items.Restrict(f"[SentOn] >= '{since:%m/%d/%Y}'")

    vs = VendorSheet(args.config)
    tmp = tempfile.mkdtemp(prefix="po_learn_")
    seen_po, seen_ship = 0, 0
    for item in items:
        try:
            if getattr(item, "Class", 0) != 43:
                continue
            subj = item.Subject or ""
            if re.search(r"주문서|발주서", subj) or re.search(r"-\d{5}\b", subj):
                for att in item.Attachments:
                    if not att.FileName.lower().endswith(".pdf"):
                        continue
                    p = os.path.join(tmp, re.sub(r'[\\/:*?"<>|]', "_", att.FileName))
                    att.SaveAsFile(p)
                    vend = po_vendor_from_pdf(p)
                    if not vend:
                        continue
                    seen_po += 1
                    to, cc = recipients(item)
                    recv, title = greeting_info(item.Body)
                    r = vs.find_by_po_name(vend)
                    if r is None:
                        r = vs.add_vendor(vend, 받는사람="", 메모="Outlook에서 자동 추가")
                    vs.add_po_name(r, vend)
                    vs.add_emails(r, to, cc)
                    vs.set_field_if_empty(r, "수신줄", recv)
                    vs.set_field_if_empty(r, "호칭", title)
                    break
            elif re.search(r"출하 요청", subj):
                to, cc = recipients(item)
                if not to:
                    continue
                makers = makers_from_html(item.HTMLBody)
                if not makers:
                    continue
                seen_ship += 1
                # 받는 주소로 업체 찾기
                r = None
                for rr in vs.rows():
                    mine = [x.lower() for x in _split(vs.get(rr, "받는사람"))]
                    if any(a.lower() in mine for a in to):
                        r = rr; break
                if r is None:
                    continue
                for mk in makers:
                    if vs.find_by_maker(mk) is None:
                        vs.add_maker(r, mk)
        except Exception as e:
            print(f"[건너뜀] {getattr(item, 'Subject', '')}: {e}")

    print(f"읽은 메일: 주문서 {seen_po}건, 출하 요청 {seen_ship}건")
    if not vs.changes:
        print("바뀐 내용 없음."); return
    print("\n[변경 내용]")
    for c in vs.changes:
        print("  - " + c)
    if args.dry_run:
        print("\n(--dry-run: 저장 안 함)")
    else:
        vs.save(); print(f"\n저장: {args.config}")


if __name__ == "__main__":
    main()
