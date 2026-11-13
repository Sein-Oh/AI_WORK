# -*- coding: utf-8 -*-
r"""pandas-kit 시험용 복잡한 MBOM 샘플 두 판 (집 PC 2026-11-13, 사용자와 함께 회사 BOM과 비슷하게 다듬음).

    python AIKIT/dev/pandas-kit/tools/make_mbom_samples.py [--out <폴더>]

<폴더>(기본: 이 저장소의 AIKIT/dev/work/pandas, git 제외)에 만듦:
  MBOM_DR200.xlsx     제품 전동 드릴 DR-200, 시트 MBOM·Routing·변경이력. MBOM은 상위 열 없이 Level과 행 순서로만 구조
                      (어떤 행의 상위 = 위쪽 행 중 가장 가까운, Level이 더 작은 행). 맨 아래 TRAPS가 일부러 넣은 함정.
  MBOM_DR200_v2.xlsx  v1에서 세 곳만 바꾼 판: 100100 아래 000125 Qty 4 -> 5, 두 번째 300210 아래 000125 Qty 3 -> 4,
                      110050(LED 창) 행 삭제. 경로로 비교하면 "삭제 1행, 바뀐 칸 2곳".
  MBOM_DR200_v2_mod.xlsx  v2를 회사 BOM 개정처럼 고친 판 (MOD_CHANGES: 바뀐 칸 8, 추가 3행, 삭제 3행, 대체·이동 포함,
                      MOD_TRAPS: 행 밀림·순서만 바뀜·글자 숫자->숫자·제목 블록은 바뀐 것이 아님).
도구가 아니라 시험용 (실행 파일만 두는 AIKIT/pandas-kit/에 넣지 않음). xlsx는 반입 규칙상 저장소에 넣지 않고 이 스크립트로 만든다.
"""
import argparse
import copy
import datetime as dt
import os

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

DEFAULT_OUT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "work", "pandas"))

HEAD = ["Level", "Part Number", "Rev", "Description", "Make/Buy", "Item Type",
        "Qty", "UoM", "Scrap %", "Op No", "Work Center", "Plant", "Supplier", "Alt Group",
        "Eff From", "Eff To", "Note"]

D0 = dt.date(2026, 1, 1)


# 트리: (PN, Rev, 설명, M/B/P, 종류, 수량, 단위, 스크랩%, 공정, 작업장, 공급사, 대체그룹, 시작, 끝, 비고, 자식들)
def n(pn, desc, mb, typ, qty, uom="EA", rev="A", scrap=0, op="", wc="", sup="", alt="", eff_from=D0, eff_to=None,
      note="", kids=()):
    return dict(pn=pn, rev=rev, desc=desc, mb=mb, typ=typ, qty=qty, uom=uom, scrap=scrap, op=op, wc=wc, sup=sup,
                alt=alt, eff_from=eff_from, eff_to=eff_to, note=note, kids=list(kids))


# 공용 부품 (여러 조립체에서 다시 씀)
def screw_m3(q, op, wc):
    return n("000123", "나사 M3x8 (아연)", "B", "PART", q, op=op, wc=wc, sup="한빛볼트")


def screw_m4(q, op, wc):
    return n("000124", "나사 M4x12 (아연)", "B", "PART", q, op=op, wc=wc, sup="한빛볼트")


def washer(q, op, wc):
    return n("000456", "와셔 M3", "B", "PART", q, op=op, wc=wc, sup="한빛볼트")


def grease(q, op, wc):
    return n("900010", "그리스 (리튬계)", "B", "CONS", q, uom="KG", op=op, wc=wc, sup="대성윤활")


def oring(pn, size, q, op, wc):
    return n(pn, f"O-링 {size}", "B", "PART", q, op=op, wc=wc, sup="미래실링")


# 원자재
def steel(q, op, wc):
    return n("RM-1001", "봉강 SCM440 D12", "B", "RAW", q, uom="KG", scrap=3, op=op, wc=wc, sup="포항소재")


def copper(q, op, wc):
    return n("RM-2001", "에나멜 동선 0.5mm", "B", "RAW", q, uom="KG", scrap=2, op=op, wc=wc, sup="LS전선")


def resin(pn, desc, q, op, wc):
    return n(pn, desc, "B", "RAW", q, uom="KG", scrap=4, op=op, wc=wc, sup="롯데케미칼")


# 2단 유성기어 스테이지: 같은 PN이 기어박스 아래 두 번 (Op No만 다름)
def planet_stage(find_note, op):
    return n("300210", "유성기어 스테이지", "M", "ASSY", 1, op=op, wc="GB-02", note=find_note, kids=[
        n("300211", "유성기어 Z12", "M", "PART", 3, op="0010", wc="MC-03", kids=[
            steel(0.012, "0010", "MC-03")]),
        n("300212", "캐리어", "M", "PART", 1, op="0010", wc="MC-03", kids=[
            steel(0.045, "0010", "MC-03")]),
        n("300213", "핀 D3", "B", "PART", 3, op="0020", wc="GB-02", sup="정밀핀"),
        grease(0.004, "0020", "GB-02"),
    ])


TOP = n("DR-200", "전동 드릴 DR-200", "M", "FG", 1, rev="C", op="0100", wc="FA-01", kids=[
    n("100100", "하우징 조립", "M", "ASSY", 1, rev="B", op="0010", wc="ASM-01", kids=[
        n("110010", "하우징 좌", "M", "PART", 1, op="0010", wc="IM-01", kids=[
            resin("RM-3001", "PA66 GF30 (검정)", 0.085, "0010", "IM-01")]),
        n("110020", "하우징 우", "M", "PART", 1, op="0010", wc="IM-01", kids=[
            resin("RM-3001", "PA66 GF30 (검정)", 0.083, "0010", "IM-01")]),
        n("110030", "손잡이 고무", "M", "PART", 2, op="0020", wc="IM-02", kids=[
            resin("RM-3002", "TPE 쇼어 60A", 0.012, "0020", "IM-02")]),
        n("110040", "명판 라벨", "B", "PART", 1, op="0030", wc="ASM-01", sup="대한인쇄"),
        n("110050", "LED 창", "B", "PART", 1, op="0030", wc="ASM-01", sup="투명광학"),
        screw_m3(8, "0030", "ASM-01"),
        n("110060", "스위치 트리거", "B", "PART", 1, rev="C", op="0030", wc="ASM-01", sup="오므론코리아"),
        n("110070", "정역 레버", "M", "PART", 1, op="0020", wc="IM-02", kids=[
            resin("RM-3003", "POM (흰색)", 0.004, "0020", "IM-02")]),
    ]),
    n("100200", "모터 조립", "M", "ASSY", 1, rev="D", op="0020", wc="MT-01", kids=[
        n("200100", "로터 조립", "M", "ASSY", 1, op="0010", wc="MT-02", kids=[
            n("200110", "샤프트", "M", "PART", 1, op="0010", wc="MC-01", kids=[steel(0.06, "0010", "MC-01")]),
            n("200120", "적층 코어", "B", "PART", 1, op="0020", wc="MT-02", sup="동양코어"),
            copper(0.11, "0030", "MT-02"),
            n("200130", "정류자", "B", "PART", 1, op="0020", wc="MT-02", sup="한일정류"),
            n("200140", "베어링 608ZZ", "B", "PART", 2, op="0040", wc="MT-02", sup="SKF코리아"),
        ]),
        n("200200", "스테이터 조립", "M", "ASSY", 1, op="0020", wc="MT-03", kids=[
            n("200210", "요크", "M", "PART", 1, op="0010", wc="MC-02", kids=[steel(0.14, "0010", "MC-02")]),
            n("200220", "페라이트 자석", "B", "PART", 2, op="0020", wc="MT-03", sup="성림첨단"),
            n("200230", "자석 클립", "B", "PART", 2, op="0020", wc="MT-03", sup="성림첨단"),
        ]),
        n("200150", "카본 브러시 (구형)", "B", "PART", 2, op="0030", wc="MT-01", sup="한일정류",
          eff_to=dt.date(2026, 12, 31), note="2027-01부터 200151로 바뀜"),
        n("200151", "카본 브러시 (신형)", "B", "PART", 2, op="0030", wc="MT-01", sup="한일정류",
          eff_from=dt.date(2027, 1, 1), note="200150 대체"),
        n("200160", "브러시 홀더", "B", "PART", 2, op="0030", wc="MT-01", sup="한일정류"),
        n("200170", "냉각 팬", "M", "PART", 1, op="0030", wc="IM-01", kids=[
            resin("RM-3001", "PA66 GF30 (검정)", 0.009, "0030", "IM-01")]),
        screw_m4(2, "0030", "MT-01"),
        washer(2, "0030", "MT-01"),
    ]),
    n("100300", "기어박스 조립", "M", "ASSY", 1, rev="B", op="0030", wc="GB-01", kids=[
        n("300100", "기어 케이스", "M", "PART", 1, op="0010", wc="DC-01", kids=[
            n("RM-4001", "알루미늄 ADC12", "B", "RAW", 0.12, uom="KG", scrap=6, op="0010", wc="DC-01", sup="알루텍")]),
        n("300110", "링기어 Z45", "M", "PART", 1, op="0010", wc="MC-03", kids=[steel(0.05, "0010", "MC-03")]),
        n("300120", "선기어 Z9", "M", "PART", 1, op="0010", wc="MC-03", kids=[steel(0.008, "0010", "MC-03")]),
        planet_stage("1단", "0020"),
        planet_stage("2단", "0030"),
        n("300130", "클러치 링", "M", "PART", 1, op="0040", wc="MC-03", kids=[steel(0.02, "0040", "MC-03")]),
        n("300140", "클러치 스프링", "B", "PART", 6, op="0040", wc="GB-01", sup="대원스프링"),
        n("300150", "토크 조절 링", "M", "PART", 1, op="0040", wc="IM-02", kids=[
            resin("RM-3003", "POM (흰색)", 0.006, "0040", "IM-02")]),
        oring("000781", "P10", 1, "0040", "GB-01"),
        grease(0.015, "0040", "GB-01"),
        screw_m3(4, "0050", "GB-01"),
    ]),
    n("100400", "배터리팩 18V 2.0Ah", "B", "ASSY", 1, rev="A", op="0090", wc="FA-01", sup="삼성SDI",
      alt="B1", note="기본"),
    n("100401", "배터리팩 18V 4.0Ah", "B", "ASSY", 0, rev="A", op="0090", wc="FA-01", sup="삼성SDI",
      alt="B1", note="대체 (주문 사양)"),
    n("100500", "척 조립", "M", "ASSY", 1, op="0040", wc="ASM-02", kids=[
        n("500100", "척 바디", "M", "PART", 1, op="0010", wc="MC-01", kids=[steel(0.09, "0010", "MC-01")]),
        n("500110", "조 (jaw)", "M", "PART", 3, op="0010", wc="MC-01", kids=[steel(0.01, "0010", "MC-01")]),
        n("500120", "척 슬리브", "M", "PART", 1, op="0020", wc="IM-02", kids=[
            resin("RM-3002", "TPE 쇼어 60A", 0.015, "0020", "IM-02")]),
        n("500130", "척 고정 나사 (왼나사)", "B", "PART", 1, op="0030", wc="ASM-02", sup="한빛볼트"),
        screw_m3(2, "0030", "ASM-02"),
    ]),
    n("100600", "제어 PCB 조립", "M", "ASSY", 1, rev="E", op="0050", wc="SMT-01", kids=[
        n("600100", "PCB 생기판 2층", "B", "PART", 1, op="0010", wc="SMT-01", sup="대덕전자"),
        n("600110", "MOSFET 40V", "B", "PART", 4, op="0010", wc="SMT-01", sup="인피니언", alt="M1", note="기본"),
        n("600111", "MOSFET 40V (대체)", "B", "PART", 0, op="0010", wc="SMT-01", sup="온세미", alt="M1",
          note="대체: 600110 품절 때"),
        n("600120", "MCU 8bit", "B", "PART", 1, op="0010", wc="SMT-01", sup="마이크로칩"),
        n("600130", "적층 세라믹 콘덴서 0603", "B", "PART", 18, op="0010", wc="SMT-01", sup="삼성전기", scrap=1),
        n("600140", "저항 0603", "B", "PART", 24, op="0010", wc="SMT-01", sup="삼성전기", scrap=1),
        n("600150", "솔더 페이스트", "B", "CONS", 0.002, uom="KG", op="0010", wc="SMT-01", sup="희성소재"),
        n("600160", "LED 흰색", "B", "PART", 1, op="0010", wc="SMT-01", sup="서울반도체"),
        n("600170", "단자대", "B", "PART", 1, op="0020", wc="SMT-02", sup="커넥터원"),
        screw_m3(2, "0030", "ASM-01"),
    ]),
    n("100700", "배선 하네스", "P", "PHANTOM", 1, op="0060", wc="ASM-01", note="팬텀: 하위 품목을 상위 공정에서 바로 씀", kids=[
        n("700100", "전선 AWG18 빨강", "B", "RAW", 0.35, uom="M", scrap=5, op="0060", wc="ASM-01", sup="LS전선"),
        n("700110", "전선 AWG18 검정", "B", "RAW", 0.35, uom="M", scrap=5, op="0060", wc="ASM-01", sup="LS전선"),
        n("700120", "단자 링형", "B", "PART", 4, op="0060", wc="ASM-01", sup="커넥터원"),
        n("700130", "수축 튜브", "B", "CONS", 0.08, uom="M", op="0060", wc="ASM-01", sup="커넥터원"),
    ]),
    screw_m4(4, "0100", "FA-01"),
    washer(4, "0100", "FA-01"),
    n("900020", "나사 고정제", "B", "CONS", 0.5, uom="ML", op="0100", wc="FA-01", sup="록타이트"),
    n("100800", "포장", "P", "PHANTOM", 1, op="0110", wc="PK-01", kids=[
        n("800100", "단상자", "B", "PKG", 1, op="0110", wc="PK-01", sup="태림포장"),
        n("800110", "완충재", "B", "PKG", 2, op="0110", wc="PK-01", sup="태림포장"),
        n("800120", "사용설명서 (한/영)", "B", "PKG", 1, rev="F", op="0110", wc="PK-01", sup="대한인쇄"),
        n("800130", "PE 봉투", "B", "PKG", 1, op="0110", wc="PK-01", sup="태림포장"),
        n("800140", "외박스 (4입)", "B", "PKG", 0.25, op="0110", wc="PK-01", sup="태림포장", note="4대에 1개"),
    ]),
])


# M3x12 볼트 (사용자 요청 2026-11-13): 1레벨 조립체 6개 바로 아래(2레벨)와 몇몇 2레벨 조립체 아래(3레벨)에 끼워 넣음.
# (상위 품번, 끼울 자리, 수량, 공정, 작업장). 300210은 기어박스 아래 두 번 나오므로 두 곳 모두에 들어감.
BOLT_L2 = [("100100", 3, 4, "0030", "ASM-01"), ("100200", 2, 2, "0030", "MT-01"), ("100300", 5, 6, "0050", "GB-01"),
           ("100500", 1, 3, "0030", "ASM-02"), ("100600", 9, 2, "0030", "ASM-01"), ("100700", 2, 2, "0060", "ASM-01")]
BOLT_L3 = [("200100", 4, 2, "0040", "MT-02"), ("200200", 1, 4, "0020", "MT-03"), ("300210", 2, 3, "0020", "GB-02")]


def bolt_m3x12(q, op, wc):
    return n("000125", "볼트 M3x12 (아연)", "B", "PART", q, op=op, wc=wc, sup="한빛볼트")


def add_bolts(node, level=0):
    for parent_pn, at, q, op, wc in (BOLT_L2 if level == 1 else BOLT_L3 if level == 2 else []):
        if node["pn"] == parent_pn:
            node["kids"].insert(at, bolt_m3x12(q, op, wc))
    for k in list(node["kids"]):
        if k["pn"] != "000125":
            add_bolts(k, level + 1)


add_bolts(TOP)


def flatten(node, level=0, parent="", find=0, out=None):
    out = [] if out is None else out
    out.append((level, find, parent, node))
    for i, k in enumerate(node["kids"], 1):
        flatten(k, level + 1, node["pn"], i * 10, out)
    return out


ROUTING = [
    ("DR-200", "0100", "FA-01", "최종 조립", 5, 2.4), ("DR-200", "0110", "PK-01", "포장", 2, 0.8),
    ("100100", "0010", "ASM-01", "하우징 반조립", 3, 0.9), ("100100", "0030", "ASM-01", "스위치·라벨 장착", 2, 1.1),
    ("100200", "0010", "MT-02", "로터 권선", 15, 1.6), ("100200", "0020", "MT-03", "스테이터 조립", 8, 0.7),
    ("100200", "0030", "MT-01", "모터 최종 조립", 6, 1.2), ("100300", "0010", "DC-01", "케이스 다이캐스팅", 45, 0.5),
    ("100300", "0020", "GB-02", "1단 스테이지", 4, 0.6), ("100300", "0030", "GB-02", "2단 스테이지", 4, 0.6),
    ("100300", "0040", "GB-01", "클러치·그리스", 5, 0.9), ("100500", "0010", "MC-01", "척 가공", 20, 1.3),
    ("100600", "0010", "SMT-01", "SMT", 30, 0.35), ("100600", "0020", "SMT-02", "수삽·납땜", 10, 0.8),
]

HISTORY = [
    ("2026-03-02", "C", "100600", "PCB Rev D -> E (MOSFET 대체 그룹 M1 추가)", "개발팀"),
    ("2026-05-18", "C", "200151", "카본 브러시 신형 추가, 2027-01-01부터 적용", "품질팀"),
    ("2026-07-07", "C", "800140", "외박스 4입으로 바뀌어 수량 0.25", "물류팀"),
]

TRAPS = [
    "1~4행 제목 블록(병합 칸), 5행 빈 행, 머리글은 6행 B열부터 (A열 비움)",
    "상위 품목 열 없음: 상위는 Level로만 알 수 있음 (바로 위쪽 행 중 Level이 하나 작은 행)",
    "같은 상위 아래 같은 품번 두 번: 100300 아래 300210 (1단·2단, Op No만 다름) -> 상위+Part Number는 유일하지 않음",
    "공용 부품 여러 상위에: 000123 나사(하우징·기어박스·척·PCB), 000456 와셔, RM-1001 봉강(11곳), RM-3001 수지(3곳)",
    "000125 볼트 M3x12: 2레벨 6곳(100100·100200·100300·100500·100600·100700 아래), 3레벨 4곳(200100·200200·300210 1단·2단 아래)",
    "앞자리 0 품번: 000123, 000124, 000456, 000781 (글자 서식)",
    "단위 섞임: EA, KG, M, ML / 소수 수량 0.002~0.35 / 0.25 (4대에 외박스 1개)",
    "팬텀(P): 100700 배선 하네스, 100800 포장",
    "대체 그룹: B1(배터리 100400 기본·100401 수량 0), M1(MOSFET 600110 기본·600111 수량 0)",
    "유효 기간: 200150(~2026-12-31)과 200151(2027-01-01~)이 함께 있음",
    "숫자가 글자로 저장된 칸: 600130의 Qty '18'",
    "숨긴 행: 폐기 품목 110099 (Eff To 2025-12-31)",
    "자동 필터 걸림, 머리글 굵게·색",
    "표 아래 빈 행 2줄 뒤 메모 블록 (표가 아님)",
    "시트 3개: MBOM, Routing, 변경이력",
]


def make_v1(path):
    return write_book(path, TOP)


def write_book(path, top, bom_rev="C", made=dt.date(2026, 11, 13), qty_text=True, swap=()):
    """top 트리를 MBOM·Routing·변경이력 세 시트로 씀 (v1과 v2_mod가 같은 모양). swap: 순서를 바꿔 쓸 형제 품번 두 개."""
    rows = flatten(top)
    if swap:
        i, j = (next(k for k, r in enumerate(rows) if r[3]["pn"] == pn) for pn in swap)
        rows[i], rows[j] = rows[j], rows[i]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MBOM"
    ws["B1"] = "제조 BOM (MBOM) - 전동 드릴 DR-200"
    ws["B1"].font = Font(size=14, bold=True)
    ws.merge_cells("B1:H1")
    ws["B2"], ws["D2"] = "공장", "P100 (평택)"
    ws["B3"], ws["D3"] = "BOM Rev", bom_rev
    ws["B4"], ws["D4"] = "작성일", made
    for r in (2, 3, 4):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    hr = 6
    fill = PatternFill("solid", start_color="DDEBF7")
    for j, h in enumerate(HEAD):
        c = ws.cell(row=hr, column=2 + j, value=h)
        c.font = Font(bold=True)
        c.fill = fill
        c.alignment = Alignment(horizontal="center")

    def put(r, vals):
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=2 + j, value=v)
            if HEAD[j] in ("Part Number", "Op No"):
                c.number_format = "@"
            if HEAD[j] in ("Eff From", "Eff To") and v is not None:
                c.number_format = "yyyy-mm-dd"

    r = hr
    for level, _find, _parent, nd in rows:
        r += 1
        vals = [level, nd["pn"], nd["rev"], nd["desc"], nd["mb"], nd["typ"], nd["qty"],
                nd["uom"], nd["scrap"] or None, nd["op"] or None, nd["wc"] or None, "P100", nd["sup"] or None,
                nd["alt"] or None, nd["eff_from"], nd["eff_to"], nd["note"] or None]
        if nd["pn"] == "600130" and qty_text:
            vals[HEAD.index("Qty")] = "18"                     # 숫자가 글자로 저장된 칸
        put(r, vals)
        if nd["pn"] == "100100":                               # 하우징 조립 바로 뒤에 숨긴 폐기 품목
            r += 1
            put(r, [2, "110099", "A", "하우징 나사 커버 (폐기)", "B", "PART", 2, "EA", None, "0030",
                    "ASM-01", "P100", "한빛볼트", None, dt.date(2024, 1, 1), dt.date(2025, 12, 31), "폐기"])
            ws.row_dimensions[r].hidden = True
    last = r
    ws.auto_filter.ref = f"B{hr}:{openpyxl.utils.get_column_letter(1 + len(HEAD))}{last}"
    ws.freeze_panes = f"B{hr + 1}"
    ws.cell(row=last + 3, column=2, value="메모")
    ws.cell(row=last + 4, column=2, value="1. 팬텀 품목은 재고를 두지 않음")
    ws.cell(row=last + 5, column=2, value="2. 대체 그룹은 수량 0인 품목이 대체품")
    for col, w in zip("BCDEFGHIJKLMNOPQR", (6, 13, 5, 26, 9, 10, 7, 6, 8, 7, 11, 7, 12, 9, 11, 11, 30)):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet("Routing")
    for j, h in enumerate(["Item", "Op No", "Work Center", "Operation", "Setup min", "Run min"]):
        ws2.cell(row=1, column=1 + j, value=h).font = Font(bold=True)
    for i, row in enumerate(ROUTING, 2):
        for j, v in enumerate(row):
            c = ws2.cell(row=i, column=1 + j, value=v)
            if j in (0, 1):
                c.number_format = "@"
    ws3 = wb.create_sheet("변경이력")
    for j, h in enumerate(["날짜", "BOM Rev", "품번", "내용", "작성"]):
        ws3.cell(row=1, column=1 + j, value=h).font = Font(bold=True)
    for i, row in enumerate(HISTORY, 2):
        for j, v in enumerate(row):
            ws3.cell(row=i, column=1 + j, value=v)
    wb.save(path)
    return last - hr, f"B{hr}:{openpyxl.utils.get_column_letter(1 + len(HEAD))}{last}"


def make_v2(v1_path, path):
    """v1에서 세 곳만 바꿈 (행 번호가 아니라 내용으로 찾음)."""
    wb = openpyxl.load_workbook(v1_path)
    ws = wb["MBOM"]
    pn_col, qty_col = 2 + HEAD.index("Part Number"), 2 + HEAD.index("Qty")
    pns = [(r, ws.cell(row=r, column=pn_col).value) for r in range(7, ws.max_row + 1)]
    first_bolt = next(r for r, pn in pns if pn == "000125")                          # 100100 아래
    stage2 = [r for r, pn in pns if pn == "300210"][1]
    stage2_bolt = next(r for r, pn in pns if pn == "000125" and r > stage2)          # 두 번째 300210 아래
    assert ws.cell(row=first_bolt, column=qty_col).value == 4 and ws.cell(row=stage2_bolt, column=qty_col).value == 3
    ws.cell(row=first_bolt, column=qty_col, value=5)
    ws.cell(row=stage2_bolt, column=qty_col, value=4)
    ws.delete_rows(next(r for r, pn in pns if pn == "110050"))                         # LED 창
    wb.save(path)


def kid(node, *pns):
    """경로로 찾기: kid(TOP, "100300", "300210#2") (같은 상위 아래 같은 품번은 #순번)."""
    for pn in pns:
        pn, _, no = pn.partition("#")
        node = [k for k in node["kids"] if k["pn"] == pn][int(no or 1) - 1]
    return node


def v2_tree():
    """make_v2가 v1 파일에 한 세 가지 변경을 트리에 똑같이."""
    top = copy.deepcopy(TOP)
    kid(top, "100100", "000125")["qty"] = 5
    kid(top, "100300", "300210#2", "000125")["qty"] = 4
    hs = kid(top, "100100")
    hs["kids"].remove(kid(hs, "110050"))
    return top


# v2 -> v2_mod: 회사 BOM 개정에서 흔한 변경 (사용자 요청 2026-11-14). 시험이 이 표로 diff 결과를 맞춰 봄.
# (구분, 경로, 열, 옛 값, 새 값). 경로 = 최상위부터 품번을 /로 이은 것, 같은 상위 아래 같은 품번은 #순번.
MOD_CHANGES = [
    ("바뀜", "DR-200/100500/000123", "Qty", 2, 3),                                   # 공용 나사, 척 아래 한 곳만
    ("바뀜", "DR-200/100300/300210#1/300213", "Qty", 3, 4),                          # 같은 품번 두 스테이지 중 1단만
    ("바뀜", "DR-200/100300/300210#2/300212/RM-1001", "Qty", 0.045, 0.05),           # 2단 아래 원자재만
    ("바뀜", "DR-200/100300/300140", "Supplier", "대원스프링", "신영스프링"),
    ("바뀜", "DR-200/100100/110060", "Rev", "C", "D"),
    ("바뀜", "DR-200/100200/200150", "Eff To", "2026-12-31", "2027-03-31"),         # 적용 시점 미룸
    ("바뀜", "DR-200/100200/200150", "Note", "2027-01부터 200151로 바뀜", "2027-04부터 200151로 바뀜"),
    ("바뀜", "DR-200/100200/200151", "Eff From", "2027-01-01", "2027-04-01"),
    ("추가", "DR-200/100600/600180", None, None, None),                             # 새 부품 TVS 다이오드
    ("삭제", "DR-200/900020", None, None, None),                                     # 나사 고정제 빠짐
    ("삭제", "DR-200/100600/600160", None, None, None),                              # 대체: 600160 -> 600161 (같은 자리)
    ("추가", "DR-200/100600/600161", None, None, None),
    ("삭제", "DR-200/100100/110040", None, None, None),                              # 이동: 하우징 아래 -> 최종 조립
    ("추가", "DR-200/110040", None, None, None),
]
# 바뀌지 않았는데 바뀐 것처럼 보이기 쉬운 것 (비교 결과에 나오면 틀림)
MOD_TRAPS = [
    "행 번호: 행이 추가·삭제되어 그 아래 행의 엑셀 행 번호가 모두 밀림 (행 번호로 짝을 맞추면 거의 모든 행이 바뀜)",
    "순서만 바뀜: 100700 아래 700100과 700110의 순서",
    "적는 방식만 바뀜: 600130의 Qty가 글자 '18'에서 숫자 18로",
    "표 밖: 제목 블록의 BOM Rev C -> D, 작성일 2026-11-20 (표 비교 대상 아님)",
    "숨긴 행 110099는 그대로",
]


def v2_mod_tree():
    top = v2_tree()
    kid(top, "100500", "000123")["qty"] = 3
    kid(top, "100300", "300210#1", "300213")["qty"] = 4
    kid(top, "100300", "300210#2", "300212", "RM-1001")["qty"] = 0.05
    kid(top, "100300", "300140")["sup"] = "신영스프링"
    kid(top, "100100", "110060")["rev"] = "D"
    old, new = kid(top, "100200", "200150"), kid(top, "100200", "200151")
    old["eff_to"], old["note"], new["eff_from"] = dt.date(2027, 3, 31), "2027-04부터 200151로 바뀜", dt.date(2027, 4, 1)
    pcb = kid(top, "100600")
    at = pcb["kids"].index(kid(pcb, "600170")) + 1
    pcb["kids"].insert(at, n("600180", "TVS 다이오드", "B", "PART", 1, op="0010", wc="SMT-01", sup="서울반도체"))
    led = kid(pcb, "600160")
    pcb["kids"][pcb["kids"].index(led)] = n("600161", "LED 주광색", "B", "PART", 1, op="0010", wc="SMT-01", sup="서울반도체")
    top["kids"].remove(kid(top, "900020"))
    hs = kid(top, "100100")
    label = kid(hs, "110040")
    hs["kids"].remove(label)
    label.update(op="0100", wc="FA-01")
    top["kids"].insert(top["kids"].index(kid(top, "100800")), label)
    return top


def make_v2_mod(path):
    write_book(path, v2_mod_tree(), bom_rev="D", made=dt.date(2026, 11, 20), qty_text=False, swap=("700100", "700110"))


def main():
    ap = argparse.ArgumentParser(description="pandas-kit 시험용 MBOM 샘플 두 판")
    ap.add_argument("--out", default=DEFAULT_OUT, help="만들 폴더 (기본: 저장소 AIKIT/dev/work/pandas)")
    out = ap.parse_args().out
    os.makedirs(out, exist_ok=True)
    v1, v2 = os.path.join(out, "MBOM_DR200.xlsx"), os.path.join(out, "MBOM_DR200_v2.xlsx")
    n, area = make_v1(v1)
    make_v2(v1, v2)
    v2m = os.path.join(out, "MBOM_DR200_v2_mod.xlsx")
    make_v2_mod(v2m)
    print(f"{v1}: MBOM 데이터 {n}행 ({area}), 최대 레벨 {max(l for l, *_ in flatten(TOP))}")
    print(f"{v2}: v1에서 Qty 두 칸(000125) 바꾸고 110050 행 삭제")
    for t in TRAPS:
        print(" -", t)
    print(f"{v2m}: v2에서 바꾼 것 (경로 기준, 바뀐 칸 {sum(c[0] == '바뀜' for c in MOD_CHANGES)}, "
          f"추가 {sum(c[0] == '추가' for c in MOD_CHANGES)}행, 삭제 {sum(c[0] == '삭제' for c in MOD_CHANGES)}행)")
    for c in MOD_CHANGES:
        print(" -", " | ".join("" if v is None else str(v) for v in c))
    print("  바뀌지 않았는데 바뀐 것처럼 보이기 쉬운 것:")
    for t in MOD_TRAPS:
        print(" -", t)


if __name__ == "__main__":
    main()
