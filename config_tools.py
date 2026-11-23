# -*- coding: utf-8 -*-
"""config.xlsx 읽기/수정 공용 함수 (업체 추가, 받는사람/MAKER/거래처명 채우기)"""
import re
from openpyxl import load_workbook


def _norm(s):
    s = str(s or "")
    s = re.sub(r"\(\s*주\s*\)|주식회사|㈜|\(\s*\d+\s*\)|\[[^\]]*\]", " ", s)
    s = re.sub(r"(^|[\s(])주(?=[\s)]|$)", r"\1", s)  # PDF에서 띄어져 읽힌 "( 주 반도기전 )", "반도기전 주 김해지점"
    return re.sub(r"[\s()\-_.]", "", s).lower()


def _split(s):
    return [x.strip() for x in re.split(r"[,;/\n]", str(s or "")) if x.strip()]


def _join(items):
    return ", ".join(dict.fromkeys(x for x in items if x))


class VendorSheet:
    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb["업체"]
        hdr = [str(c.value).strip() if c.value else "" for c in self.ws[1]]
        self.col = {h: i + 1 for i, h in enumerate(hdr) if h}
        for need in ("거래처명(주문서)", "발주서발송", "주문서약칭", "주문서추가문구"):
            if need not in self.col:
                self.ws.cell(row=1, column=self.ws.max_column + 1, value=need)
                self.col[need] = self.ws.max_column
        self.changes = []

    # --- 조회
    def rows(self):
        out = []
        for r in range(2, self.ws.max_row + 1):
            name = self.ws.cell(r, self.col["업체"]).value
            if not name or str(name).startswith("※"):
                continue
            out.append(r)
        return out

    def get(self, r, key):
        return str(self.ws.cell(r, self.col[key]).value or "").strip() if key in self.col else ""

    def set(self, r, key, val):
        if key in self.col:
            self.ws.cell(r, self.col[key], val)

    def names(self):
        return [self.get(r, "업체") for r in self.rows()]

    def find_by_name(self, name):
        key = _norm(name)
        for r in self.rows():
            if _norm(self.get(r, "업체")) == key or key in [_norm(x) for x in _split(self.get(r, "주문서약칭"))]:
                return r
        return None

    def find_by_po_name(self, po_vendor):
        key = _norm(po_vendor)
        for r in self.rows():
            if key in [_norm(x) for x in _split(self.get(r, "거래처명(주문서)"))]:
                return r
        best, tie = None, False            # 등록된 이름(4글자 이상)이 들어 있음 - 가장 긴 것, 한 업체일 때만
        for r in self.rows():
            for a in [_norm(self.get(r, "업체"))] + [_norm(x) for x in _split(self.get(r, "거래처명(주문서)"))]:
                if len(a) < 4 or not key.startswith(a):
                    continue
                if best is None or len(a) > len(best[0]):
                    best, tie = (a, r), False
                elif len(a) == len(best[0]) and r != best[1]:
                    tie = True
        return best[1] if best and not tie else None

    def find_by_maker(self, maker):
        key = str(maker or "").strip().upper()
        for r in self.rows():
            if key in [x.upper() for x in _split(self.get(r, "MAKER"))]:
                return r
        return None

    # --- 수정
    def add_vendor(self, name, **fields):
        r = self.ws.max_row + 1
        # 비고(※) 줄 위에 끼워 넣기
        for rr in range(2, self.ws.max_row + 1):
            v = self.ws.cell(rr, 1).value
            if v is None or str(v).startswith("※"):
                self.ws.insert_rows(rr); r = rr; break
        self.set(r, "업체", name)
        self.set(r, "활성", "Y"); self.set(r, "엑셀첨부", "N"); self.set(r, "발주서발송", "Y")
        self.set(r, "거래처명(주문서)", name); self.set(r, "주문서약칭", name)
        for k, v in fields.items():
            self.set(r, k, v)
        self.changes.append(f"업체 추가: {name}")
        return r

    def add_emails(self, r, to_list, cc_list=()):
        cur = _split(self.get(r, "받는사람"))
        new = [x for x in to_list if x.lower() not in [c.lower() for c in cur]]
        if not cur and new:
            self.set(r, "받는사람", _join(new)); self.changes.append(f"{self.get(r,'업체')}: 받는사람 ← {_join(new)}")
        elif new:
            self.changes.append(f"{self.get(r,'업체')}: 받는사람에 없는 주소 발견(미반영) {_join(new)}")
        curcc = _split(self.get(r, "추가참조"))
        newcc = [x for x in cc_list if x.lower() not in [c.lower() for c in cur + curcc]]
        if newcc and not curcc:
            self.set(r, "추가참조", _join(newcc)); self.changes.append(f"{self.get(r,'업체')}: 추가참조 ← {_join(newcc)}")

    def add_po_name(self, r, po_vendor):
        cur = _split(self.get(r, "거래처명(주문서)"))
        if _norm(po_vendor) not in [_norm(x) for x in cur]:
            self.set(r, "거래처명(주문서)", _join(cur + [po_vendor]))
            self.changes.append(f"{self.get(r,'업체')}: 거래처명(주문서) + {po_vendor}")

    def add_maker(self, r, maker):
        cur = _split(self.get(r, "MAKER"))
        if maker.upper() not in [x.upper() for x in cur]:
            self.set(r, "MAKER", _join(cur + [maker]))
            self.changes.append(f"{self.get(r,'업체')}: MAKER + {maker}")

    def set_field_if_empty(self, r, key, val):
        if val and not self.get(r, key):
            self.set(r, key, val); self.changes.append(f"{self.get(r,'업체')}: {key} ← {val}")

    def save(self):
        self.wb.save(self.path)
        return self.changes
