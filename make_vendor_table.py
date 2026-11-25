"""거래처 표(거래처코드 시트) 엑셀 → js/vendor_table.js (주간 리스트 만들기·메일 작성기에 미리 넣어 두는 표)
사용: python make_vendor_table.py 거래처표.xlsx   → js/vendor_table.js 를 새로 쓰고 GitHub에 올린다"""
import sys, json, datetime
from openpyxl import load_workbook

HEAD = '/* 거래처 표 (내장) — 같은 MAKER를 여러 업체가 쓰는 품목을 업체로 나누는 기준.\n   원본: 거래처코드 시트 (매입업체 | MAKER | 열1 | 열2 | 예외 거래처 | 예외 품목명 | 예외 MAKER | M코드). 앞 8열만.\n   바꾸려면: 주간 리스트 만들기 화면에 새 거래처 표 파일을 넣으면 그 PC에서는 새 표를 쓴다(이 파일보다 우선).\n   모두에게 바꾸려면 이 파일을 새 표로 다시 만든다(README 참고).\n   규칙(vtRule): ① 예외 표의 M코드가 같고 품목명이 예외 품목명으로 시작 ② 품목명이 예외 품목명과 같음 ③ MAKER별 기본 거래처 */\n"use strict";\n'
TAIL = 'const vtNsp = (s) => clean(s).replace(/\\s+/g, "").toUpperCase();\nfunction vtRule(tblRows) {\n  const maker = new Map(), ex = [];\n  for (const r of tblRows.slice(1)) {\n    const mk = vtNsp(r[1]) ? clean(r[1]).toUpperCase() : "";\n    if (mk && !maker.has(mk)) maker.set(mk, clean(r[0]));\n    if (clean(r[4]) && clean(r[5])) ex.push({ vendor: clean(r[4]), nm: vtNsp(r[5]), codes: clean(r[7]).split(",").map((x) => x.trim()).filter(Boolean) });\n  }\n  return (cd, nm, mk) => {\n    const n = vtNsp(nm);\n    for (const e of ex) if (e.codes.includes(cd) && n.startsWith(e.nm)) return e.vendor;\n    for (const e of ex) if (n === e.nm) return e.vendor;\n    return maker.get(clean(mk).toUpperCase()) || "";\n  };\n}\n// 쓸 거래처 표: 이 PC에 새로 넣은 표가 내장 표보다 새 것이면 그것, 아니면 내장 표\nconst LS_VTBL = "mailmaker_vendor_table_v1";\nfunction currentVendorTable() {\n  const saved = store.get(LS_VTBL, null);\n  return saved && saved.rows && saved.at >= BUILTIN_VTBL.at ? saved : BUILTIN_VTBL;\n}\nlet _vtRuleCache = null;\n// 메일 작성기: 리스트에 업체 열이 없을 때 품목코드·품명·MAKER로 거래처 이름을 찾는다 (없으면 "")\nfunction vendorFromTable(cd, nm, mk) {\n  const t = currentVendorTable();\n  if (!_vtRuleCache || _vtRuleCache.t !== t.at + t.from) _vtRuleCache = { t: t.at + t.from, f: vtRule(t.rows) };\n  return _vtRuleCache.f(clean(cd), nm, mk);\n}\n'

def cell(v):
    if v is None: return ""
    if isinstance(v, datetime.datetime): return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return str(v).strip()

def main(path):
    wb = load_workbook(path, data_only=True)
    ws = wb["거래처코드"] if "거래처코드" in wb.sheetnames else wb.worksheets[0]
    rows = [r for r in ([cell(x) for x in row[:8]] for row in ws.iter_rows(values_only=True)) if any(r)]
    meta = {"from": "거래처 표 (내장)", "at": datetime.date.today().isoformat()}
    body = 'const BUILTIN_VTBL = { from: %s, at: %s, rows: %s };\n' % (json.dumps(meta["from"], ensure_ascii=False), json.dumps(meta["at"]), json.dumps(rows, ensure_ascii=False))
    open("js/vendor_table.js", "w", encoding="utf-8").write(HEAD + body + TAIL)
    print(f"js/vendor_table.js: {len(rows)}줄 ({meta['at']})")

if __name__ == "__main__":
    if len(sys.argv) < 2: sys.exit(__doc__)
    main(sys.argv[1])
