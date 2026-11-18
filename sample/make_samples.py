# 테스트용 합성 입력 파일 (실제 데이터 아님)
import os
from openpyxl import Workbook
H = os.path.dirname(os.path.abspath(__file__))

wb = Workbook(); ws = wb.active; ws.title = "선박"
ws.append(["주간 선박 리스트"]); ws.append([]); ws.append(["작성일", "2026-11-16"])
ws.append(["Fac", "ITEM CD", "MAKER", "ITEM NM", "ETD 20-NOV-26", "ETD 27-NOV-26", "ETD 04-DEC-26", "ETD 11-DEC-26", "sum", "BALANCE"])
rows = [
 ["A", "MX-50001", "MOLEX", "CONNECTOR 2P", "3,000", 1500, "", 2000, 6500, 100],
 ["A", "MX-50002", "", "HOUSING 4P (MOLEX)", "", "", "", 500, 500, 0],      # 선택 ETD 밖 → 제외
 ["B", "TE-77001", "TE", "TERMINAL 0.5SQ", 10000, "", 5000, "", 15000, 0],
 ["B", "AMP-1234", "AMP", "WIRE SEAL", "", 2000, "", "", 2000, 0],
 ["A", "BY-001", "BUNYOUNG", "TUBE 3MM", 800, 800, 800, "", 2400, 0],
 ["C", "DC-9", "DAEIL CST", "CABLE TIE", "", "", 1200, "", 1200, 0],
 ["A", "KET-12345", "KET", "TERMINAL KET A", 500, "", "", "", 500, 0],       # 품목지정 → 태양단자
 ["A", "KET-99999", "KET", "TERMINAL KET B", 300, 300, "", "", 600, 0],      # MAKER → 태양단자
 ["A", "WS-1", "DONG-A BESTECH", "TERMINAL W", 100, "", "", "", 100, 0],
 ["A", "UNK-1", "SOMEMAKER", "UNKNOWN PART", 50, "", "", "", 50, 0],         # 미지정
 ["A", "FL-1", "FEELUX", "LED MODULE", "", "", "", "", 0, 0],                 # 수량 없음 → 제외
 ["A", "DA-1", "", "SHRINK TUBE (DAEA)", 700, "", "", "", 700, 0],            # 업체 열 없음, 메이커 DAEA 미지정
]
for r in rows: ws.append(r)
wb.save(os.path.join(H, "ship_sample.xlsx"))

wb = Workbook(); ws = wb.active; ws.title = "에어"
ws.append(["에어 리스트"]); ws.append([])
ws.append(["FAC", "Item cd", "Maker", "Part Name", "Remain", "STOCK", "Short Del.", "NEED ETD"])
rows = [
 ["A", "MX-50010", "MOLEX", "CONNECTOR 6P", "1,200", 300, "11/18", "2026-11-25"],
 ["B", "TE-77002", "TE", "TERMINAL 1.25SQ", 0, 500, "", ""],                  # Remain 0 → 제외
 ["B", "BY-002", "BUNYOUNG", "TUBE 5MM", 400, 0, "11/19", "2026-11-26"],
 ["A", "FL-2", "FEELUX", "LED BAR", 60, 0, "", "2026-11-30"],                 # 에어만 → (에어) 제목
 ["A", "UNK-2", "", "MYSTERY PART", 10, 0, "", ""],                           # 미지정
]
for r in rows: ws.append(r)
wb.save(os.path.join(H, "air_sample.xlsx"))
print("ok")
