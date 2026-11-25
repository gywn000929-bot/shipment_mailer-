# -*- coding: utf-8 -*-
"""
출하 요청 메일 자동 작성기
- 선박리스트/에어리스트 엑셀 -> 업체별 Outlook 초안(임시 보관함) 생성
- 자동 발송 없음. 사람이 Outlook에서 확인 후 직접 보냄.

사용법 (Windows, Outlook 설치 PC):
    python make_drafts.py --ship 선박리스트.xlsx --air 에어리스트.xlsx
    python make_drafts.py --ship 선박리스트.xlsx --air 에어리스트.xlsx --dry-run   # Outlook 없이 HTML만 생성
자세한 옵션: python make_drafts.py -h
"""
import argparse
import datetime as dt
import html
import math
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG = os.path.join(HERE, "config.xlsx")
YELLOW = "FFFF00"
WEEKDAY_KO = "월화수목금토일"


# ---------------------------------------------------------------- 유틸
def norm(s):
    """헤더 비교용: 공백 제거, 소문자."""
    return re.sub(r"\s+", "", str(s)).lower() if s is not None else ""


def to_num(v):
    """'3,000' 같은 문자열도 숫자로. 숫자 아니면 0."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return 0
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return 0
    try:
        return float(s)
    except ValueError:
        return 0


def fmt_num(v):
    n = to_num(v)
    if n == 0:
        return ""
    if float(n).is_integer():
        return f"{int(n):,}"
    return f"{n:,.2f}"


def clean(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def split_list(s):
    return [x.strip() for x in re.split(r"[,;/\n]", clean(s)) if x.strip()]


def maker_from_name(name):
    """품명 끝의 '(메이커)'에서 메이커 추출."""
    m = re.search(r"\(([^()]+)\)\s*$", clean(name))
    return m.group(1).strip() if m else ""


def find_header_row(df, must_have, max_rows=40, any_of=(), prefix_any=()):
    """처음 max_rows행 안에서 must_have 헤더가 모두 있고, any_of 중 하나(또는 prefix_any로 시작하는 열)가 있는 행. 없으면 None."""
    need = [norm(x) for x in must_have]
    anyn = [norm(x) for x in any_of]
    pre = [norm(x) for x in prefix_any]
    for i in range(min(max_rows, len(df))):
        cells = [norm(c) for c in df.iloc[i].tolist()]
        if not all(n in cells for n in need):
            continue
        if anyn and not any(a in cells for a in anyn):
            continue
        if pre and not any(c.startswith(p) for c in cells for p in pre):
            continue
        return i
    return None


def load_sheet(path, must_have, label="", any_of=(), prefix_any=(), prefer=(), sheet_name=None):
    """파일의 모든 시트를 돌며 조건에 맞는 헤더가 있는 시트를 찾는다 (선박/에어가 한 파일에 있어도 됨).
    여러 시트가 맞으면 prefer 열(예: Remain)이 있는 시트를 먼저, sheet_name을 주면 그 시트만."""
    xl = pd.ExcelFile(path)
    found = []
    for sheet in xl.sheet_names:
        if sheet_name and norm(sheet) != norm(sheet_name):
            continue
        raw = pd.read_excel(xl, sheet, header=None, dtype=object)
        h = find_header_row(raw, must_have, any_of=any_of, prefix_any=prefix_any)
        if h is None:
            continue
        headers = [clean(c) for c in raw.iloc[h].tolist()]
        df = raw.iloc[h + 1:].copy()
        df.columns = headers
        df = df.loc[:, [c for c in df.columns if c != ""]]
        df = df.dropna(how="all").reset_index(drop=True)
        df.attrs["sheet"] = sheet
        score = sum(1 for p in prefer if norm(p) in [norm(c) for c in df.columns])
        found.append((score, len(found), df))
    if not found:
        raise SystemExit(f"[오류] {label} 파일에서 {label} 표(헤더 {must_have}{' + ' + '/'.join(any_of) if any_of else ''}{' + ETD열' if prefix_any else ''})가 있는 시트를 못 찾았습니다.\n"
                         f"       파일: {os.path.basename(path)} / 시트: {xl.sheet_names}" + (f" / 지정한 시트: {sheet_name}" if sheet_name else "") + "\n"
                         f"       선박리스트와 에어리스트에 같은 파일을 고른 것은 아닌지 확인하세요.")
    found.sort(key=lambda x: (-x[0], x[1]))
    df = found[0][2]
    if len(found) > 1:
        print(f"[참고] {label} 표가 있는 시트가 {len(found)}개: {[f[2].attrs['sheet'] for f in found]} → '{df.attrs['sheet']}' 사용"
              + (f" (Remain 열이 있는 시트 우선)" if prefer else "") + ". 다른 시트를 쓰려면 --ship-sheet/--air-sheet 로 지정")
    return df


def col(df, *candidates, required=True, prefix=False):
    """후보 이름 중 (공백/대소문자 무시) 일치하는 실제 열 이름. prefix=True면 앞부분 일치도 허용."""
    lookup = {norm(c): c for c in df.columns}
    for cand in candidates:
        if norm(cand) in lookup:
            return lookup[norm(cand)]
    if prefix:
        for cand in candidates:
            for k, v in lookup.items():
                if k.startswith(norm(cand)):
                    return v
    if required:
        raise SystemExit(f"[오류] 열을 찾을 수 없습니다: {candidates}\n       실제 열: {list(df.columns)}")
    return None


# ---------------------------------------------------------------- 설정
def load_config(path):
    xl = pd.ExcelFile(path)
    vendors = pd.read_excel(xl, "업체", dtype=object).fillna("")
    overrides = pd.read_excel(xl, "품목지정", dtype=object).fillna("") if "품목지정" in xl.sheet_names else pd.DataFrame()
    settings_df = pd.read_excel(xl, "설정", dtype=object, header=None).fillna("")
    settings = {clean(r[0]): clean(r[1]) for _, r in settings_df.iterrows() if clean(r[0])}

    vend = {}
    maker_map = {}
    for _, r in vendors.iterrows():
        name = clean(r.get("업체"))
        if not name or name.startswith("※") or clean(r.get("활성", "Y")).upper() in ("N", "NO", "X", "0"):
            continue
        v = {
            "name": name,
            "to": split_list(r.get("받는사람")),
            "cc": split_list(r.get("추가참조")),
            "recv_line": clean(r.get("수신줄")),
            "title": clean(r.get("호칭")),
            "makers": [m.upper() for m in split_list(r.get("MAKER"))],
            "attach": clean(r.get("엑셀첨부", "Y")).upper() not in ("N", "NO", "X", "0"),
            "memo": clean(r.get("메모")),
            "send_ship": clean(r.get("출하요청발송", "Y")).upper() not in ("N", "NO", "X", "0"),
        }
        vend[name] = v
        for m in v["makers"]:
            maker_map.setdefault(m, name)
        v["aliases"] = [name] + split_list(r.get("거래처명(주문서)")) + split_list(r.get("주문서약칭"))

    item_map = {}
    if not overrides.empty and "ITEM CD" in [c.strip() for c in overrides.columns]:
        for _, r in overrides.iterrows():
            code, vname = clean(r.get("ITEM CD")), clean(r.get("업체"))
            if code and vname:
                item_map[norm(code)] = vname
    alias_map = {}
    for name, v in vend.items():
        for a in v["aliases"]:
            if norm_name(a):
                alias_map.setdefault(norm_name(a), name)
    return vend, maker_map, item_map, settings, alias_map


def norm_name(s):
    """업체명 비교용: (주), 주식회사, [코드], 공백, 괄호 제거, 소문자."""
    s = clean(s)
    s = re.sub(r"\(\s*주\s*\)|주식회사|㈜|\(\s*\d+\s*\)|\[[^\]]*\]", " ", s)
    s = re.sub(r"(^|[\s(])주(?=[\s)]|$)", r"\1", s)  # PDF에서 띄어져 읽힌 "( 주 반도기전 )", "반도기전 주 김해지점"
    return re.sub(r"[\s()\-_.]", "", s).lower()


def match_alias(name, alias_map):
    """리스트의 업체 칸 값 → config 업체명. 1) 정확히 일치 2) 별칭이 포함(긴 별칭 우선)."""
    key = norm_name(name)
    if not key:
        return name
    if key in alias_map:
        return alias_map[key]
    best = None
    for a, v in alias_map.items():
        if len(a) >= 2 and (a in key or key in a) and (best is None or len(a) > len(best[0])):
            best = (a, v)
    return best[1] if best else name


def resolve_vendor(row_vendor_col, item_cd, maker, maker_map, item_map, alias_map=None):
    """우선순위: 선박리스트 '업체' 열 > 품목지정 시트 > MAKER 매핑."""
    if row_vendor_col:
        if alias_map:
            return match_alias(row_vendor_col, alias_map)
        return row_vendor_col
    if norm(item_cd) in item_map:
        return item_map[norm(item_cd)]
    return maker_map.get(clean(maker).upper(), "")


# ---------------------------------------------------------------- 입력 파싱
def parse_ship(path, etd_count, etd_names, maker_map, item_map, alias_map=None, sheet_name=None):
    df = load_sheet(path, ["ITEM CD", "MAKER"], "선박리스트", prefix_any=("ETD",), sheet_name=sheet_name)
    c_fac = col(df, "Fac", "FAC")
    c_cd = col(df, "ITEM CD")
    c_mk = col(df, "MAKER")
    c_nm = col(df, "ITEM NM", "Part Name", "품명")
    c_vendor = col(df, "업체", "거래처명(매칭)", "거래처명", "거래처", "VENDER", "VENDOR", required=False, prefix=True)
    print(f"선박리스트 시트: {df.attrs.get('sheet')}" + (f" / 업체 열: '{c_vendor}' 사용" if c_vendor else " / 업체 열 없음 → MAKER로 매칭"))
    etd_all = [c for c in df.columns if norm(c).startswith("etd")]
    if not etd_all:
        raise SystemExit("선박리스트에 ETD 열이 없습니다.")
    if etd_names:
        etd_sel = []
        for n in etd_names:
            m = [c for c in etd_all if norm(c) == norm(n) or norm(n) in norm(c)]
            if not m:
                raise SystemExit(f"ETD 열 '{n}' 없음. 있는 열: {etd_all}")
            etd_sel.append(m[0])
    else:
        etd_sel = etd_all[:etd_count]

    rows = []
    for _, r in df.iterrows():
        cd = clean(r[c_cd])
        if not cd:
            continue
        qtys = [to_num(r[c]) for c in etd_sel]
        has_qty = any(q > 0 for q in qtys)
        maker = clean(r[c_mk]) or maker_from_name(r[c_nm])
        vendor = resolve_vendor(clean(r[c_vendor]) if c_vendor else "", cd, maker, maker_map, item_map, alias_map)
        rows.append({
            "Fac": clean(r[c_fac]), "ITEM CD": cd, "MAKER": maker, "ITEM NM": clean(r[c_nm]),
            **{c: r[c] for c in etd_sel}, "_vendor": vendor, "_src": "선박", "_has_qty": has_qty,
        })
    return pd.DataFrame(rows), etd_sel, etd_all


def parse_air(path, maker_map, item_map, alias_map=None, sheet_name=None):
    df = load_sheet(path, ["Item cd", "Maker"], "에어리스트", any_of=("Remain", "Purchase request"), prefer=("Remain",), sheet_name=sheet_name)
    c_fac = col(df, "FAC", "Fac")
    c_cd = col(df, "Item cd", "ITEM CD")
    c_mk = col(df, "Maker", "MAKER")
    c_nm = col(df, "Part Name", "ITEM NM", "품명")
    c_rem = col(df, "Remain", "Purchase request")
    c_short = col(df, "Short Del.", "Short Del", required=False)
    c_need = col(df, "NEED ETD", "Need ETD", required=False)
    c_vendor = col(df, "업체", "거래처명(매칭)", "거래처명", "거래처", "VENDER", "VENDOR", required=False, prefix=True)
    print(f"에어리스트 시트: {df.attrs.get('sheet')} / 수량 열: '{c_rem}'" + (f" / 업체 열: '{c_vendor}' 사용" if c_vendor else ""))

    rows = []
    for _, r in df.iterrows():
        cd = clean(r[c_cd])
        if not cd or to_num(r[c_rem]) <= 0:
            continue
        maker = clean(r[c_mk]) or maker_from_name(r[c_nm])
        vendor = resolve_vendor(clean(r[c_vendor]) if c_vendor else "", cd, maker, maker_map, item_map, alias_map)
        rows.append({
            "FAC": clean(r[c_fac]), "Item cd": cd, "Maker": maker, "Part Name": clean(r[c_nm]),
            "Purchase request": r[c_rem],
            "Short Del.": fmt_date_cell(r[c_short]) if c_short else "",
            "NEED ETD": fmt_date_cell(r[c_need]) if c_need else "",
            "_qty_col": c_rem,
            "_vendor": vendor, "_src": "에어",
        })
    return pd.DataFrame(rows)


def fmt_date_cell(v):
    """날짜 셀은 엑셀에서 보이는 형태(20-Nov-26)로, 글자(URGENT 등)는 그대로."""
    if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)):
        t = pd.Timestamp(v)
        return f"{t.day:02d}-{t.strftime('%b')}-{t.strftime('%y')}"
    return clean(v)


# ---------------------------------------------------------------- 날짜/주차
def next_monday(today=None):
    today = today or dt.date.today()
    return today + dt.timedelta(days=(7 - today.weekday()) % 7 or 7)


def week_label(monday):
    """해당 주 월요일+3일(목요일)이 속한 달과 ceil(일/7)."""
    thu = monday + dt.timedelta(days=3)
    return thu.month, math.ceil(thu.day / 7)


def ko_date(d):
    return f"{d.month}월 {d.day}일({WEEKDAY_KO[d.weekday()]})"


# ---------------------------------------------------------------- 본문
def html_table(headers, rows, highlight_cols=()):
    th = "".join(
        f'<th style="border:1px solid #999;padding:3px 8px;background:#f2f2f2;white-space:nowrap">{html.escape(h)}</th>'
        for h in headers)
    trs = []
    for row in rows:
        tds = []
        for h, v in zip(headers, row):
            bg = f"background:{'#' + YELLOW}" if h in highlight_cols else ""
            align = "text-align:right" if isinstance(v, (int, float)) or re.fullmatch(r"[\d,]+", str(v)) else ""
            style = ";".join(x for x in ["border:1px solid #999", "padding:3px 8px", bg, align] if x)
            tds.append(f'<td style="{style}">{html.escape(str(v))}</td>')
        trs.append("<tr>" + "".join(tds) + "</tr>")
    return (f'<table style="border-collapse:collapse;font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:10pt">'
            f"<tr>{th}</tr>{''.join(trs)}</table>")


def build_body(vendor, air_df, ship_df, etd_sel, month, week, air_date, ship_date, settings, with_sig, attached):
    """남효주 담당자 기존 메일 양식 (필룩스/태양기전 메일 기준)."""
    p = lambda s: f'<p style="margin:0">{s}</p>'
    br = '<p style="margin:0">&nbsp;</p>'
    has_ship = ship_df is not None and len(ship_df)
    has_air = air_df is not None and len(air_df)
    out = []
    if vendor["memo"]:
        out.append(f'<p style="color:red;font-weight:bold;background:#fff3cd;padding:6px;border:2px solid red">'
                   f'※ {html.escape(vendor["memo"])} (보내기 전 이 줄은 지워 주세요)</p>')
    if vendor["recv_line"]:
        out += [p(html.escape(vendor["recv_line"])), br]
    greet = (vendor["title"] + " " if vendor["title"] else "") + (str(settings.get("출하요청인사") or "").strip() or "안녕하세요.") + " " + settings.get("발신자소개", "대아홀딩스 남효주 입니다.")
    out.append(p(html.escape(greet)))
    if has_ship:
        out.append(p(("첨부" if attached else "하기") + " 내용으로 금주 선박 리스트 송부 드립니다."))
    else:
        out.append(p("하기 내용으로 금주 에어 리스트 송부 드립니다."))

    if has_ship:
        out.append(p(f"선박리스트 (노란색 수량, <b>{short_date(ship_date)}</b> 입고 요망)"))
        if attached:
            out.append(p("첨부 선박 리스트는 자재 입고 전 입고수량 기재하여 회신 부탁드립니다."))
        else:
            hdr = ["Fac", "ITEM CD", "MAKER", "ITEM NM"] + list(etd_sel)
            rows = [[r["Fac"], r["ITEM CD"], r["MAKER"], r["ITEM NM"]] + [fmt_num(r[c]) for c in etd_sel]
                    for _, r in ship_df.iterrows()]
            out.append(html_table(hdr, rows, highlight_cols=(etd_sel[0],)))
        out.append(br)

    if has_air:
        out.append(p(f"에어리스트 (노란색 수량, <b>{short_date(air_date)}</b> 입고 요망)"))
        qty_label = air_df.iloc[0]["_qty_col"] if "_qty_col" in air_df.columns else "Purchase request"
        hdr = ["FAC", "Item cd", "Maker", "Part Name", qty_label]
        if any(clean(r["Short Del."]) for _, r in air_df.iterrows()):
            hdr.append("Short Del.")
        hdr.append("NEED ETD")
        rows = []
        for _, r in air_df.iterrows():
            row = [r["FAC"], r["Item cd"], r["Maker"], r["Part Name"], fmt_num(r["Purchase request"])]
            if "Short Del." in hdr:
                row.append(r["Short Del."])
            row.append(r["NEED ETD"])
            rows.append(row)
        out += [html_table(hdr, rows, highlight_cols=(qty_label,)), br]

    out.append(p("감사합니다."))
    if with_sig:
        sig = settings.get("서명", "")
        if sig:
            out += [br] + [p(html.escape(line)) for line in sig.split("\n")]
    size = clean(settings.get("본문크기")) or "10"
    bold = clean(settings.get("본문굵게")).upper() in ("Y", "YES", "O", "1")
    return (f'<div style="font-family:맑은 고딕,Malgun Gothic,sans-serif;font-size:{size}pt{";font-weight:bold" if bold else ""}">'
            + "".join(out) + "</div>")


def short_date(d):
    """11/18일(수) 형식."""
    return f"{d.month}/{d.day}일({WEEKDAY_KO[d.weekday()]})"


# ---------------------------------------------------------------- 엑셀 첨부/미지정 목록
def _style_ws(ws, highlight_col_idx=None):
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="F2F2F2")
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = border
    if highlight_col_idx:
        for r in range(1, ws.max_row + 1):
            ws.cell(row=r, column=highlight_col_idx).fill = PatternFill("solid", fgColor=YELLOW)
    for col_cells in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(8, width * 1.3 + 2), 60)


def write_vendor_xlsx(path, ship_df, air_df, etd_sel):
    """기존 첨부 양식: Fac | ITEM CD | MAKER | ITEM NM | 이번주 ETD(노란색) | 입고수량(업체 기재) | 다음 ETD..."""
    wb = Workbook()
    wb.remove(wb.active)
    if ship_df is not None and len(ship_df):
        ws = wb.create_sheet("선박")
        hdr = ["Fac", "ITEM CD", "MAKER", "ITEM NM", etd_sel[0], "입고수량"] + list(etd_sel[1:])
        ws.append(hdr)
        for _, r in ship_df.iterrows():
            ws.append([r["Fac"], r["ITEM CD"], r["MAKER"], r["ITEM NM"], to_num(r[etd_sel[0]]) or 0, None]
                      + [to_num(r[c]) or 0 for c in etd_sel[1:]])
        for row in ws.iter_rows(min_row=2, min_col=5, max_col=len(hdr)):
            for c in row:
                c.number_format = "#,##0"
        _style_ws(ws, highlight_col_idx=5)
    if air_df is not None and len(air_df):
        ws = wb.create_sheet("에어")
        ws.append(["FAC", "Item cd", "Maker", "Part Name", "Purchase request", "Short Del.", "NEED ETD"])
        for _, r in air_df.iterrows():
            ws.append([r["FAC"], r["Item cd"], r["Maker"], r["Part Name"], to_num(r["Purchase request"]),
                       r["Short Del."], r["NEED ETD"]])
        for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
            for c in row:
                c.number_format = "#,##0"
        _style_ws(ws, highlight_col_idx=5)
    wb.save(path)


def write_unmapped(path, ship_df, air_df, etd_sel):
    wb = Workbook()
    ws = wb.active
    ws.title = "미지정 품목"
    ws.append(["구분", "Fac", "ITEM CD", "MAKER", "ITEM NM", "수량", "비고"])
    for _, r in ship_df.iterrows():
        ws.append(["선박", r["Fac"], r["ITEM CD"], r["MAKER"], r["ITEM NM"],
                   " / ".join(f"{c}: {fmt_num(r[c])}" for c in etd_sel if to_num(r[c]) > 0),
                   "config.xlsx 업체 시트의 MAKER 또는 품목지정 시트에 추가하세요"])
    for _, r in air_df.iterrows():
        ws.append(["에어", r["FAC"], r["Item cd"], r["Maker"], r["Part Name"], fmt_num(r["Purchase request"]), ""])
    _style_ws(ws)
    wb.save(path)


# ---------------------------------------------------------------- Outlook
def outlook_app():
    try:
        import win32com.client  # noqa
    except ImportError:
        raise SystemExit("pywin32가 없습니다. 설치: pip install pywin32   (또는 --dry-run 사용)")
    import win32com.client
    return win32com.client.Dispatch("Outlook.Application")


def existing_drafts(app, subject_prefix):
    """임시 보관함에서 제목이 subject_prefix로 시작하는 초안 목록."""
    ns = app.GetNamespace("MAPI")
    drafts = ns.GetDefaultFolder(16)  # olFolderDrafts
    found = []
    for item in drafts.Items:
        try:
            subj = item.Subject or ""
        except Exception:
            continue
        if subj.startswith(subject_prefix):
            found.append((subj, item.To))
    return found


def create_draft(app, to, cc, subject, body_html, attachments, settings):
    mail = app.CreateItem(0)  # olMailItem
    mail.To = "; ".join(to)
    mail.CC = "; ".join(cc)
    mail.Subject = subject
    # GetInspector에 접근하면 기본 서명이 HTMLBody에 들어온다 (창은 뜨지 않음).
    try:
        _ = mail.GetInspector
    except Exception:
        pass
    sig_html = mail.HTMLBody or ""
    has_sig = len(re.sub(r"<[^>]+>|&nbsp;|\s", "", sig_html)) > 0
    if has_sig and re.search(r"<body[^>]*>", sig_html, re.I):
        # 본문을 서명 앞(body 바로 뒤)에 삽입
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
        mail.Close(0)  # olSave
    except Exception:
        pass
    return has_sig


# ---------------------------------------------------------------- 메인
def main():
    ap = argparse.ArgumentParser(description="업체별 출하 요청 메일 초안 생성 (Outlook 임시 보관함)")
    ap.add_argument("--ship", help="선박리스트 .xlsx")
    ap.add_argument("--ship-sheet", help="선박 표가 있는 시트 이름 (한 파일에 여러 개일 때)")
    ap.add_argument("--air-sheet", help="에어 표가 있는 시트 이름 (기본: Remain 열이 있는 시트 우선, 예: 'Air up')")
    ap.add_argument("--air", help="에어리스트 .xlsx")
    ap.add_argument("--config", default=DEFAULT_CONFIG, help="설정 엑셀 (기본 config.xlsx)")
    ap.add_argument("--etd", type=int, default=None, help="메일에 넣을 ETD 열 개수 (기본: 설정 시트 값 또는 2)")
    ap.add_argument("--etd-cols", help="ETD 열 이름 직접 지정, 쉼표 구분 (예: '20-NOV-26,27-NOV-26,04-DEC-26')")
    ap.add_argument("--week-monday", help="기준 주 월요일 YYYY-MM-DD (기본: 다음 주 월요일)")
    ap.add_argument("--air-date", help="에어 입고 요망일 YYYY-MM-DD (기본: 기준 월요일+1, 화요일)")
    ap.add_argument("--ship-date", help="선박 입고 요망일 YYYY-MM-DD (기본: 기준 월요일+2, 수요일)")
    ap.add_argument("--out", default=os.path.join(HERE, "output"), help="첨부/미리보기 저장 폴더")
    ap.add_argument("--dry-run", action="store_true", help="Outlook을 쓰지 않고 HTML 미리보기만 저장")
    ap.add_argument("--yes", "-y", action="store_true", help="확인 질문 없이 진행")
    ap.add_argument("--only", help="이 업체만 생성 (쉼표 구분)")
    args = ap.parse_args()

    if not args.ship and not args.air:
        ap.error("--ship 또는 --air 중 하나는 필요합니다.")
    for label, fp in (("선박리스트", args.ship), ("에어리스트", args.air), ("설정 파일", args.config)):
        if fp and not os.path.isfile(fp):
            print(f"[오류] {label} 파일을 찾을 수 없습니다: {fp}")
            print(f"       현재 폴더: {os.getcwd()}")
            print("       파일을 이 폴더에 복사하거나, 전체 경로를 따옴표로 감싸서 지정하세요.")
            print('       예: python make_drafts.py --ship "C:\\경로\\선박리스트.xlsx" --air "C:\\경로\\에어리스트.xlsx"')
            sys.exit(1)

    vendors, maker_map, item_map, settings, alias_map = load_config(args.config)
    common_cc = split_list(settings.get("공통참조", ""))
    etd_count = args.etd or int(to_num(settings.get("ETD열개수", 2)) or 2)
    etd_names = split_list(args.etd_cols) if args.etd_cols else None

    monday = dt.date.fromisoformat(args.week_monday) if args.week_monday else next_monday()
    if monday.weekday() != 0:
        print(f"[경고] {monday}는 월요일이 아닙니다. 그 주 월요일({monday - dt.timedelta(days=monday.weekday())})로 맞춥니다.")
        monday -= dt.timedelta(days=monday.weekday())
    month, week = week_label(monday)
    air_date = dt.date.fromisoformat(args.air_date) if args.air_date else monday + dt.timedelta(days=1)
    ship_date = dt.date.fromisoformat(args.ship_date) if args.ship_date else monday + dt.timedelta(days=2)
    subject_tpl = settings.get("제목형식", "대아홀딩스 {월}월 {주}주 출하 요청의 건")
    subject_base = subject_tpl.replace("{월}", str(month)).replace("{주}", str(week))
    include_zero = clean(settings.get("수량없는행포함", "Y")).upper() not in ("N", "NO", "X", "0")

    ship_df, etd_sel = pd.DataFrame(), []
    if args.ship:
        ship_df, etd_sel, etd_all = parse_ship(args.ship, etd_count, etd_names, maker_map, item_map, alias_map, sheet_name=args.ship_sheet)
        print(f"선박리스트 ETD 열(전체): {etd_all}")
        print(f"선박리스트 ETD 열(선택): {etd_sel}  ← 첫 열 '{etd_sel[0]}'이 이번 주(노란색)")
    air_df = parse_air(args.air, maker_map, item_map, alias_map, sheet_name=args.air_sheet) if args.air else pd.DataFrame()

    os.makedirs(args.out, exist_ok=True)
    tag = f"{monday:%Y%m%d}_{month}월{week}주"

    # 미지정
    if len(ship_df) and not include_zero:
        ship_df = ship_df[ship_df["_has_qty"]].reset_index(drop=True)
    un_ship = ship_df[(ship_df["_vendor"] == "") & ship_df["_has_qty"]] if len(ship_df) else ship_df
    un_air = air_df[air_df["_vendor"] == ""] if len(air_df) else air_df
    unmapped_path = os.path.join(args.out, f"미지정_품목_{tag}.xlsx")
    import json
    um = {}
    for _, r in (un_ship.iterrows() if len(un_ship) else []):
        um.setdefault(clean(r["MAKER"]) or "(MAKER 없음)", []).append(f"{r['ITEM CD']} {r['ITEM NM']}")
    for _, r in (un_air.iterrows() if len(un_air) else []):
        um.setdefault(clean(r["Maker"]) or "(MAKER 없음)", []).append(f"{r['Item cd']} {r['Part Name']}")
    unmatched_json_path = os.path.join(args.out, "unmatched_ship.json")
    unmatched_json = [{"key": k, "samples": v[:5], "count": len(v), "kind": "maker"} for k, v in um.items()]
    if len(un_ship) or len(un_air):
        write_unmapped(unmapped_path, un_ship, un_air, etd_sel)

    # 업체별 묶기
    names = sorted(set(ship_df["_vendor"].tolist() if len(ship_df) else []) |
                   set(air_df["_vendor"].tolist() if len(air_df) else []))
    names = [n for n in names if n]
    if args.only:
        only = split_list(args.only)
        names = [n for n in names if n in only]

    plan = []
    print("\n==================== 미리보기 ====================")
    print(f"기준 주: {monday} (월)   제목: {subject_base}")
    print(f"에어 입고 요망: {ko_date(air_date)}   선박 입고 요망: {ko_date(ship_date)}")
    print(f"공통 참조: {', '.join(common_cc)}\n")
    problems = []
    unknown_vendor_names = []
    for n in names:
        s = ship_df[ship_df["_vendor"] == n] if len(ship_df) else pd.DataFrame()
        a = air_df[air_df["_vendor"] == n] if len(air_df) else pd.DataFrame()
        v = vendors.get(n)
        if v is None:
            problems.append(f"'{n}' : 리스트의 업체 열 값이 config 업체 시트에 없습니다 (선박 {len(s)} / 에어 {len(a)}) → ③ 미지정 업체 지정하기로 연결하세요")
            unknown_vendor_names.append((n, len(s) + len(a)))
            continue
        if not v.get("send_ship", True):
            if (int(s["_has_qty"].sum()) if len(s) else 0) or len(a):
                print(f"[건너뜀] {n} : 출하 요청 메일을 보내지 않는 업체 (출하요청발송=N)")
            continue
        if not v["to"]:
            problems.append(f"'{n}' : 받는 사람이 비어 있습니다 (선박 {len(s)} / 에어 {len(a)}) → 초안 생성 안 함")
            continue
        n_ship_qty = int(s["_has_qty"].sum()) if len(s) else 0
        if n_ship_qty == 0 and not len(a):
            continue  # 이번 주 수량이 하나도 없으면 메일 안 만듦
        subject = subject_base + (settings.get("에어만제목접미", "(에어)") if len(a) and not n_ship_qty else "")
        if n_ship_qty == 0:
            s = s.iloc[0:0]
        plan.append((v, s, a, subject))
        print(f"■ {n}  선박 {n_ship_qty}건{'(전체 ' + str(len(s)) + '행)' if len(s) != n_ship_qty else ''} / 에어 {len(a)}건  "
              f"{'[엑셀 첨부]' if v['attach'] else '[본문 표]'}  제목: {subject}")
        print(f"   받는 사람: {', '.join(v['to'])}")
        if v["cc"]:
            print(f"   추가 참조: {', '.join(v['cc'])}")
        if v["memo"]:
            print(f"   메모: {v['memo']}")
    for n, cnt in unknown_vendor_names:
        unmatched_json.append({"key": n, "samples": [f"리스트의 업체 열 값 ({cnt}건)"], "count": cnt, "kind": "vendor"})
    with open(unmatched_json_path, "w", encoding="utf-8") as f:
        json.dump(unmatched_json, f, ensure_ascii=False)
    if problems:
        print("\n[문제]")
        for pr in problems:
            print("  - " + pr)
    if len(un_ship) or len(un_air):
        print(f"\n[미지정 품목] 선박 {len(un_ship)}건 / 에어 {len(un_air)}건 → 메일 생성 안 함. 목록: {unmapped_path}")
        for _, r in un_ship.head(15).iterrows():
            print(f"   선박  {r['ITEM CD']:<20} {r['MAKER']:<18} {r['ITEM NM']}")
        for _, r in un_air.head(15).iterrows():
            print(f"   에어  {r['Item cd']:<20} {r['Maker']:<18} {r['Part Name']}")
    print("==================================================")
    if not plan:
        print("생성할 초안이 없습니다.")
        return

    app = None
    if not args.dry_run:
        app = outlook_app()
        dup = existing_drafts(app, subject_base)
        if dup:
            print(f"\n[경고] 임시 보관함에 같은 주차 초안이 이미 {len(dup)}건 있습니다:")
            for subj, to in dup:
                print(f"   - {subj}  → {to}")
            print("   계속하면 같은 업체 초안이 중복으로 생깁니다. 기존 초안을 먼저 지우거나 --only로 일부만 만드세요.")

    if not args.yes:
        ans = input(f"\n{len(plan)}개 업체 초안을 {'HTML로 저장' if args.dry_run else 'Outlook 임시 보관함에 생성'}할까요? [y/N] ").strip().lower()
        if ans not in ("y", "yes"):
            print("취소했습니다.")
            return

    for v, s, a, subject in plan:
        atts = []
        if v["attach"]:
            xp = os.path.join(args.out, f"{subject_base}_{v['name']}.xlsx")
            write_vendor_xlsx(xp, s, a, etd_sel)
            atts.append(xp)
        body = build_body(v, a, s, etd_sel, month, week, air_date, ship_date, settings,
                          with_sig=args.dry_run, attached=bool(atts) and len(s) > 0)
        cc = list(dict.fromkeys(common_cc + v["cc"]))
        if args.dry_run:
            hp = os.path.join(args.out, f"{v['name']}_{tag}.html")
            head = (f"<p><b>To:</b> {html.escape('; '.join(v['to']))}<br><b>CC:</b> {html.escape('; '.join(cc))}<br>"
                    f"<b>Subject:</b> {html.escape(subject)}<br><b>첨부:</b> {html.escape(', '.join(os.path.basename(x) for x in atts) or '없음')}</p><hr>")
            with open(hp, "w", encoding="utf-8") as f:
                f.write('<html><head><meta charset="utf-8"></head><body>' + head + body + "</body></html>")
            print(f"  저장: {hp}")
        else:
            has_sig = create_draft(app, v["to"], cc, subject, body, atts, settings)
            print(f"  초안 생성: {v['name']}  ({'Outlook 서명' if has_sig else '설정 서명'} 사용, 첨부 {len(atts)})")
    print("\n완료. Outlook 임시 보관함(Drafts)에서 내용 확인 후 직접 보내세요." if not args.dry_run else "\n완료 (dry-run).")


if __name__ == "__main__":
    main()
