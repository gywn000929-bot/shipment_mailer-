# -*- coding: utf-8 -*-
"""
업체 마스터 한 곳에서 두 설정 파일을 만든다.
  config.xlsx        : 남효주 기준 (출하 요청 메일 + 주문서)
  config_kim.xlsx : 김민지 계장 기준 (주문서 전 업체 + 출하 요청)
업체 시트는 둘 다 같고, 설정 시트(이름/서명/제목형식/본문)만 다르다.
실행: python make_config.py --force
"""
import os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
HDR = ["업체", "받는사람", "추가참조", "수신줄", "호칭", "MAKER", "엑셀첨부", "메모", "활성", "거래처명(주문서)", "발주서발송", "주문서약칭", "주문서추가문구"]

# 업체, 받는사람, 추가참조, 수신줄, 호칭, MAKER, 엑셀첨부, 메모, 활성, 거래처명(주문서), 발주서발송, 주문서약칭, 주문서추가문구
VENDORS = [
 # ---- 출하 요청 메일 주 대상 (기존 발송 메일 기준) ----
 ["대일CST",      "skim@daeilcst.co.kr", "", "", "계장님", "DAEIL CST", "Y", "", "Y", "대일CST, 대일", "Y", "대일CST", ""],
 ["반도기전 김해", "sclee@bandoamp.com, kjkim@bandoamp.com", "", "", "", "AMP, TE", "Y", "", "Y", "반도기전 김해, 반도기전(김해)", "Y", "반도김해", ""],
 ["반도기전 진영", "sales4@bandomk.com, sales1@bandomk.com", "jongho.lee@bandomk.com", "", "선임님", "MOLEX", "Y", "", "Y", "반도기전, 반도기전 진영", "Y", "반도진영", ""],
 ["엠이시시스텍",  "smlee@mec.co.kr, yuk7400@mec.co.kr", "", "", "차장님", "MEC", "N", "", "Y", "엠이시시스텍, MEC", "Y", "MEC", ""],
 ["위성단자",     "wesung-1989@daum.net", "", "", "과장님", "CAIS ELECTRONICS, JIN", "N", "", "Y", "위성단자", "Y", "위성", ""],
 ["정우 RTI",     "jungwoo4050@hanmail.net", "", "", "대표님", "JUNGWOO", "N", "", "Y", "정우 RTI, 정우", "Y", "정우", ""],
 ["태양기전",     "kyc8391@hanmail.net, hanyang78@hanmail.net, hoongidance@gmail.com", "", "", "",
  "DAEMYUNG, DUCKSUNG, TAEYOUNG, TAEYANG, DONG-A, DONG-A BESTECH, BAEG SAN, KDK", "Y", "", "Y", "태양기전", "Y", "태양기전", ""],
 ["태양단자",     "kwak@tae-yang.co.kr", "", "", "대리님", "KET", "N", "KET 품목 중 태양단자 경유분. 품목지정 시트로 나눌 수 있음", "Y", "태양단자", "Y", "태양단자", ""],
 ["성우실업",     "swcfuse@thermalcutoff.com", "", "", "과장님", "SUNG WOO", "N", "", "Y", "성우실업, 성우", "Y", "성우", ""],
 ["연호",        "yh0511@yeonho.com", "", "", "", "YEONHO", "N", "주문서 메일 안 보냄", "Y", "연호전자, 연호", "N", "연호", ""],
 ["보성",        "sales@bosungsilitech.co.kr", "", "수신: 김진수 이사님", "이사님", "", "N", "", "Y", "보성실리텍, 보성", "Y", "보성", ""],
 ["삼풍",        "", "", "", "", "", "N", "메일 주소 확인 필요", "Y", "삼풍", "Y", "삼풍", ""],
 ["번영하이테크",  "by_hitech@naver.com", "", "수신: 이미경 실장님", "실장님", "BUNYOUNG", "N", "", "Y", "번영하이테크, 번영", "Y", "번영", ""],
 ["대아산업",     "sale@daeatube.com, songdh40@daeatube.com", "", "수신: 송대호 과장님, 조정화 대리님", "대리님", "", "N", "출하 요청 메일에 주문서(PDF) 첨부 필요", "Y", "대아산업", "Y", "대아산업", ""],
 ["뉴콘",        "", "", "", "", "", "N", "메일 주소 확인 필요", "Y", "뉴콘", "Y", "뉴콘", ""],
 ["대성테크",     "daesung3570@naver.com", "", "", "대표님", "", "N", "최석주 대표 (계장님 메일 기준)", "Y", "대성테크", "Y", "대성테크", ""],
 # ---- 계장님 주문서 대상 (하단 업체) ----
 ["KET",         "kkh@ket.com", "hkko@ket.com, ksj95@ket.com, kim.js@ket.com, grace@ket.com, nsc@ket.com, ketsale@ket.com, tskim@ket.com", "", "", "KET", "N",
  "주문서 메일 안 보냄 (주간 긴급 자재 메일은 별도)", "Y", "한국단자공업, 한국단자, KET", "N", "KET", ""],
 ["BIT",         "gb0831@nate.com, aaa0802@nate.com", "", "", "대리님", "BIT", "N", "", "Y", "비아이티 이앤엠, 비아이티, BIT, 비아이티이엔엠, BITENM", "Y", "BIT", ""],
 ["경신",        "hdlee@kyungshin.co.kr, hksong@kyungshin.co.kr", "", "", "매니저님", "KYUNG SHIN", "N", "주문서 메일 안 보냄(엑셀 발주서 별도). 이헌동 선임매니저·송현경 매니저", "Y", "경신전선, 경신", "N", "경신", ""],
 ["필룩스",       "vtrboy@feelux.com", "rhkdqja2801@feelux.com, gotn@feelux.com, ryu80@feelux.com", "", "부장님", "FEELUX", "N", "", "Y", "필룩스, KH필룩스, 케이에이치필룩스", "Y", "필룩스", ""],
 ["TSE",         "tse2023@naver.com", "", "", "대표님", "TSE", "N", "", "Y", "티에스이, TSE, 태성일렉트론", "Y", "TSE", ""],
 ["한림",        "hanlimsale@daum.net", "", "", "", "HANLIM", "N", "", "Y", "한림전자, 한림", "Y", "한림", ""],
 ["JST",         "theconn@theconnintl.co.kr", "", "수신: 유진서 주임님", "주임님", "JST", "N", "더콘INTL 경유 (더콘과 같은 주소). 거래처명 확인 필요", "Y", "JST, 한국제이에스티", "Y", "JST", ""],
 ["SL",          "", "", "", "", "SL", "N", "메일 주소 확인 필요", "Y", "SL, 에스엘", "Y", "SL", ""],
 ["더콘",        "theconn@theconnintl.co.kr", "", "수신: 유진서 주임님", "주임님", "", "N", "AMPHENOL 품목은 텔콤 업체로 분리함", "Y", "더콘, 더콘INTL", "Y", "더콘", ""],
 ["텔콤",        "cslim@telcom.kr", "chlee@telcom.kr, seunghwan.lee@telcom.kr, kslee@telcom.kr", "", "", "AMPHENOL", "N", "임충신 담당(미납잔량 메일 기준). 더콘과 같은 회사인지 확인 필요", "Y", "텔콤, 텔콤씨앤에스", "Y", "텔콤", ""],
 ["태진",        "taejin@taejin2.com", "", "", "", "TAEJIN", "N", "", "Y", "태진아이엔티, 태진", "Y", "태진", ""],
 ["삼화전자공업",  "cbc2728@samwha.com, hms@samwha.com", "", "", "", "SAMWHA", "N", "", "Y", "삼화전자공업, 삼화", "Y", "삼화", ""],
 ["행성사",       "runnertec@runnertec.com", "", "수신: 최성호 팀장님", "팀장님", "", "N", "런너테크 경유 (행성사 리드타임 5~6주)", "Y", "행성사, 런너테크", "Y", "행성사", ""],
 ["이테크",       "etech67@korea.com", "", "", "", "ETECH", "N", "", "Y", "이테크전자, 이테크", "Y", "이테크", ""],
 ["기성전선",     "ygje@kiseong.co.kr", "", "", "", "KISEONG", "N", "", "Y", "기성전선, 기성", "Y", "기성", ""],
 ["명보케이블",    "pkk@mbcable.co.kr, kch@mbcable.co.kr", "", "", "", "MYUNGBO", "N", "", "Y", "명보케이블, 명보", "Y", "명보", ""],
 ["엔에이치티",    "jhl@nhtco.kr", "", "", "", "NHT", "N", "담당자 변경(2026.10~ 이지현)", "Y", "엔에이치티, 뉴한테크, NHT", "Y", "NHT", ""],
 ["부천공업",     "", "", "", "", "", "N", "메일 주소 확인 필요", "Y", "부천공업, 부천", "Y", "부천", ""],
 ["에코로아",     "ecoroa05@ecoroakorea.com", "", "", "책임님", "ECOROA", "N", "", "Y", "에코로아, 주식회사 에코로아 (UJU)", "Y", "에코로아", ""],
 ["상아 TRADING", "ls9783@naver.com", "", "", "", "LITTEL FUSE", "N", "", "Y", "상아 TRADING, 상아트레이딩, 상아", "Y", "상아", ""],
 ["세일테크",     "seilcom3@hanmail.net", "", "", "", "SEIL", "N", "", "Y", "세일테크, 세일", "Y", "세일", "해당 납기일에 입고 부탁드립니다."],
 ["하이테크",     "jhg@ihitech.co.kr, lgy@ihitech.co.kr", "", "", "계장님", "HITECH", "N", "", "Y", "하이테크전자부품, 하이테크", "Y", "하이테크", ""],
 ["ASIREL",      "", "", "", "", "ASIREL", "N", "메일 주소 확인 필요", "Y", "ASIREL", "Y", "ASIREL", ""],
 ["엠이",        "", "", "", "", "", "N", "엠이시시스텍과 다른 업체인지 확인 필요", "Y", "엠이", "Y", "엠이", ""],
]

NOTES = [
 "※ MAKER: 선박/에어리스트의 MAKER 칸 값(쉼표로 여러 개). 비어 있으면 그 업체 품목은 '미지정'으로 나옴. 받는사람/참조도 쉼표 구분.",
 "※ 엑셀첨부=Y: 출하 요청 메일에 선박 표 대신 엑셀 첨부. 메모: 출하 요청 본문 맨 위 빨간 글씨(보내기 전 삭제). 활성=N: 무시.",
 "※ 거래처명(주문서): 주문서 PDF 거래처명 칸의 이름(쉼표로 여러 개). 발주서발송=N: 주문서 메일 안 만듦. 주문서약칭: 제목/첨부 파일명. 주문서추가문구: 본문 '감사합니다' 앞 한 줄.",
 "※ 선박리스트에 '업체' 열이 있으면 MAKER보다 그 값을 먼저 씀. 값은 이 시트의 '업체' 이름과 같아야 함.",
]


def build(path, settings_rows, overrides_rows, strip_greeting=False):
    wb = Workbook(); ws = wb.active; ws.title = "업체"
    ws.append(HDR)
    for r in VENDORS:
        r = list(r)
        if strip_greeting:      # 계장님 메일은 '수신: ○○○' 줄·호칭을 쓰지 않음 (기존 메일 기준)
            r[3] = ""; r[4] = ""
        ws.append(r)
    ws.append([])
    for n in NOTES:
        ws.append([n])
    ws2 = wb.create_sheet("품목지정"); ws2.append(["ITEM CD", "업체", "비고"])
    for r in overrides_rows:
        ws2.append(r)
    ws3 = wb.create_sheet("설정"); ws3.append(["항목", "값", "설명"])
    for r in settings_rows:
        ws3.append(r)
    for w in (ws, ws2, ws3):
        for c in w[1]:
            c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDEBF7")
        for col_cells in w.columns:
            width = max(len(str(c.value).split("\n")[0]) if c.value else 0 for c in col_cells)
            w.column_dimensions[col_cells[0].column_letter].width = min(max(10, width * 1.4 + 2), 70)
    for row in ws3.iter_rows(min_row=2, min_col=2, max_col=2):
        for c in row:
            if c.value and "\n" in str(c.value):
                c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path); print("생성:", path)


COMMON = [
 ["제목형식", "대아홀딩스 {월}월 {주}주 출하 요청의 건", "출하 요청 메일 제목. {월}, {주}가 숫자로 바뀜"],
 ["에어만제목접미", "(에어)", "선박 없이 에어만 있을 때 제목 끝"],
 ["ETD열개수", "2", "메일에 넣을 ETD 열 개수. --etd 옵션으로 덮어쓸 수 있음"],
 ["수량없는행포함", "Y", "Y: 업체 전체 품목(수량 0 포함)을 표/첨부에 넣음(기존 메일 방식). N: 이번 주 수량 있는 행만"],
]

NAM = [
 ["공통참조", "ejkim@kdaea.com, minji3907@kdaea.com, hwanggj@kdaea.com", "모든 메일에 들어가는 참조"],
 ["발신자소개", "대아홀딩스 남효주 입니다.", "'안녕하세요. ' 뒤에 붙는 문구"],
 *COMMON,
 ["주문서제목형식", "대아 주문서-{PO}", "주문서 메일 제목"],
 ["주문서첨부이름", "{PO}.pdf", "첨부 PDF 파일명"],
 ["주문서본문", "{호칭}안녕하세요. 대아홀딩스 남효주 입니다.\n수고 많으십니다.\n\n상기 제목에 따라 주문서 송부하오니\n첨부파일 확인 및 수신 여부 회신 부탁드리겠습니다.\n\n{추가문구}\n\n감사합니다.",
  "주문서 메일 본문(줄바꿈 Alt+Enter). {호칭}=업체 시트 호칭, {추가문구}=업체 시트 주문서추가문구(없으면 그 줄 삭제)"],
 ["서명", "D.E.TECH & DAE-A HOLDINGS CO.,LTD\nTEL 055-296-1677~8 / FAX 055-296-1679\nhjnam@kdaea.com", "Outlook 기본 서명이 없을 때만 사용"],
]

KIM = [
 ["공통참조", "ejkim@kdaea.com, hwanggj@kdaea.com, hjnam@kdaea.com", "모든 메일에 들어가는 참조 (기존 메일 기준)"],
 ["발신자소개", "대아홀딩스 김민지 계장입니다.", "'안녕하세요. ' 뒤에 붙는 문구"],
 *COMMON,
 ["주문서제목형식", "{업체}-{PO}", "주문서 메일 제목. 예: 기성-47198 (PO 여러 개면 쉼표로 이어짐)"],
 ["주문서첨부이름", "{업체}-{PO}.pdf", "첨부 PDF 파일명. 예: 기성-47198.pdf"],
 ["주문서본문", "수고많으십니다. 대아홀딩스 김민지 계장입니다.\n\n발주서 첨부드립니다.\n\n{추가문구}\n\n감사합니다.",
  "주문서 메일 본문(줄바꿈 Alt+Enter). {추가문구}=업체 시트 주문서추가문구(없으면 그 줄 삭제)"],
 ["서명", "D.E.TECH / DAE-A HOLDINGS CO.,LTD\nMIN JI KIM(金 敏 智)\nTEL : 055-296-1677~8\nFAX : 055-296-1679\nMOBILE : 010-2992-3907\nMAIL : minji3907@kdaea.com", "Outlook 기본 서명이 없을 때만 사용"],
 ["본문굵게", "Y", "Y: 본문 전체 굵게 (계장님 기존 메일 스타일)"],
 ["본문크기", "11", "본문 글자 크기(pt). 계장님 기존 메일 11pt"],
]

OVR = [["(예) KET-12345", "태양단자", "KET 품목 중 태양단자 경유 건. 품목 코드로 업체를 직접 지정 (MAKER보다 우선)"]]

if __name__ == "__main__":
    for name, st in (("config.xlsx", NAM), ("config_kim.xlsx", KIM)):
        p = os.path.join(HERE, name)
        if os.path.exists(p) and "--force" not in sys.argv:
            print(f"{p} 가 이미 있습니다. 다시 만들려면 --force"); continue
        build(p, st, OVR, strip_greeting=(name == "config_kim.xlsx"))
