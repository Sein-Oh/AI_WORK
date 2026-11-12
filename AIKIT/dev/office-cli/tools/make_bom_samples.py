"""BOM 비교 시험 샘플 (정답을 아는 두 판, 계층 BOM). 사람이 실행하는 개발 도구 (에이전트 계약 대상 아님).

    python make_bom_samples.py [--dir C:\\AIKIT\\dev\\work\\bom]
    python make_bom_samples.py --check "<..._비교결과.csv>"     # 맞대기 결과를 정답과 비교

만드는 파일 (--dir, 기본 C:\\AIKIT\\dev\\work\\bom):
  BOM_A_Rev1.xlsx, BOM_A_Rev2.xlsx   제목 행·빈 행 아래 머리글이 있는 두 판 (함정 포함, 아래 TRAPS)
  BOM_A_Rev2_중복.xlsx              새 판의 한 중간 조립체 아래 같은 Part Number가 두 번 (맞대기가 멈추고 사용자에게 묻는지)
  BOM_A_정답.csv                    맞대기가 내야 할 결과 (구분, 상위 조립체, Part Number, 바뀐 열)

모양: Level 0 최상위 조립체 1개 -> Level 1 중간 조립체 12개 -> Level 2 부품 (조립체마다 5~8개).
공통 부품(볼트, 와셔, 너트 등)은 여러 중간 조립체에 반복해서 쓰인다. 같은 자리 = 상위 조립체 + Part Number.

집 PC에서 BOM 비교 스킬 만들기 가이드(tools/BOM비교_스킬만들기_가이드.md)를 연습할 때 쓴다. 사내 LLM이 만든 맞대기 결과 CSV의
앞 네 열(구분, 상위 조립체, Part Number, 바뀐 열)을 --check로 정답과 비교한다. 회사 BOM 모양에 가깝게 바꾸려면 아래 COLUMNS,
PARTS, ASSEMBLIES, CHANGES를 고친다.
"""
import argparse
import csv
import os
import sys

COLUMNS = ["Level", "Part Number", "Description", "Spec", "Qty", "Unit", "Remark"]
TITLE = "제품 A 부품표 (BOM, 계층)"
TOP = "TA-1000"

# 부품 마스터: Part Number -> (Description, Spec, Unit). 여러 조립체에서 같은 값으로 쓰인다.
PARTS = {
    "P-0001": ("육각 볼트", "M4x10", "EA"), "P-0002": ("육각 볼트", "M5x12", "EA"), "P-0003": ("평와셔", "M4", "EA"),
    "P-0004": ("스프링 와셔", "M5", "EA"), "P-0005": ("육각 너트", "M4", "EA"), "P-0006": ("케이블 타이", "100mm", "EA"),
    "P-0007": ("고무 부싱", "D12", "EA"), "P-0008": ("브래킷 L", "SPCC t1.6", "EA"), "P-0009": ("커넥터", "4P 2.54mm", "EA"),
    "P-0010": ("O링", "P20 NBR", "EA"), "P-0011": ("스페이서", "M3x8", "EA"), "P-0012": ("방열판", "AL 40x40", "EA"),
    "P-0013": ("모터", "DC 24V 30W", "EA"), "P-0014": ("기어", "M1 Z20", "EA"), "P-0015": ("베어링", "6201ZZ", "EA"),
    "P-0016": ("샤프트", "D12 L120", "EA"), "P-0017": ("센서", "근접 M12", "EA"), "P-0018": ("PCB 조립", "MAIN-V2", "EA"),
    "P-0019": ("하네스", "8P 300mm", "EA"), "P-0020": ("라벨", "50x20", "EA"),
    "0102": ("브래킷 U", "SUS304 t1.0", "EA"), "0207": ("가스켓", "EPDM 50A", "EA"),
    "P-0025": ("서지 보호기", "24V 1kA", "EA"),          # 새 판에서 처음 쓰는 부품
    "102": ("브래킷 U(신)", "SUS304 t1.2", "EA"),         # 0102와 다른 품번 (앞자리 0)
}

# 옛 판 중간 조립체: (Part Number, Description, Qty) 와 그 아래 부품 [(Part Number, Qty)]
ASSEMBLIES = [
    (("SA-01", "프레임 조립", 1), [("P-0008", 2), ("P-0002", 8), ("P-0004", 8), ("P-0011", 4), ("P-0020", 1)]),
    (("SA-02", "구동부 조립", 1), [("P-0013", 1), ("P-0014", 2), ("P-0001", 4), ("P-0003", 4), ("P-0010", 2), ("P-0004", 3)]),
    (("SA-03", "축 조립", 2), [("P-0016", 1), ("P-0015", 2), ("P-0001", 4), ("P-0003", 4), ("P-0005", 4)]),
    (("SA-04", "제어부 조립", 1), [("P-0018", 1), ("P-0009", 3), ("P-0012", 1), ("P-0011", 4), ("P-0006", 5), ("P-0019", 1)]),
    (("SA-05", "센서 조립", 4), [("P-0017", 1), ("P-0005", 2), ("P-0009", 1), ("P-0006", 2), ("P-0007", 1)]),
    (("SA-06", "덮개 조립", 1), [("P-0007", 4), ("P-0001", 6), ("P-0003", 6), ("P-0020", 2), ("0207", 1)]),
    (("SA-07", "배선 조립", 1), [("P-0019", 2), ("P-0006", 10), ("P-0009", 2), ("P-0005", 2), ("P-0020", 3)]),
    (("SA-08", "냉각부 조립", 1), [("P-0012", 2), ("P-0010", 2), ("P-0001", 4), ("P-0003", 4), ("P-0006", 2), ("0207", 2)]),
    (("SA-09", "지지대 조립", 2), [("0102", 2), ("P-0002", 4), ("P-0004", 4), ("P-0007", 4), ("P-0011", 2)]),
    (("SA-10", "도어 조립", 1), [("P-0008", 1), ("P-0002", 4), ("P-0004", 4), ("P-0015", 2), ("P-0020", 1), ("P-0007", 2)]),
    (("SA-11", "전원부 조립", 1), [("P-0018", 1), ("P-0019", 1), ("P-0009", 2), ("P-0012", 1), ("P-0011", 4)]),
    (("SA-12", "브래킷 조립", 3), [("0102", 1), ("P-0001", 2), ("P-0003", 2), ("P-0005", 2), ("P-0008", 1)]),
]

# 새 판에서 바뀌는 것. 정답은 answer()가 여기서 만든다. 함정(TRAPS)은 정답에 넣지 않는다.
CHANGES = {
    "Qty": {("SA-03", "P-0001"): 6, ("SA-07", "P-0005"): 3, (TOP, "SA-05"): 2},   # 같은 부품의 다른 자리(SA-02 등 P-0001)는 그대로
    "Spec": {"P-0010": "P20 FKM"},            # 부품 마스터 변경: 쓰인 자리마다 (SA-02, SA-08) 한 줄씩
    "Description": {"P-0018": "PCB 조립(개정)"},  # 쓰인 자리마다 (SA-04, SA-11)
    "삭제": [("SA-06", "P-0007"), ("SA-10", "P-0002"), ("SA-12", "0102")],
    "추가": [("SA-04", "P-0025", 1), ("SA-09", "P-0003", 4), ("SA-12", "102", 1)],
    "이동": [("P-0012", "SA-04", "SA-06")],   # SA-04에서 SA-06으로: 옛 자리 삭제 + 새 자리 추가
}
TRAPS = [
    "중간 조립체 순서: 새 판은 SA-08과 SA-09 묶음의 순서가 바뀜 (바뀐 것 아님)",
    "조립체 안 부품 순서: 새 판 SA-01 아래 부품 순서가 거꾸로 (바뀐 것 아님)",
    "Remark 열: SA-02 아래 P-0013 Remark '확인' 추가 (비교하지 않는 열)",
    "Part Number 공백·소문자: 새 판 SA-10 아래 P-0004는 ' p-0004 ' (같은 부품)",
    "수량 표기: 새 판 SA-01 아래 P-0002 수량 8 -> 8.0 (같은 값)",
    "Spec 대소문자: 새 판 SA-03 아래 P-0016 Spec 'd12 l120' (같은 규격)",
    "반복 부품: P-0001·P-0003·P-0006 등이 여러 조립체에 반복 (중복 아님, 자리마다 비교)",
    "앞자리 0: 옛 판 SA-12의 0102와 새 판 SA-12의 102는 다른 품번 (0102 삭제, 102 추가)",
]


def part_row(level, pn, qty, remark=""):
    name, spec, unit = PARTS[pn]
    return [level, pn, name, spec, qty, unit, remark]


def old_rows():
    rows = [[0, TOP, "제품 A", "", 1, "SET", ""]]
    for (sa, name, qty), kids in ASSEMBLIES:
        rows.append([1, sa, name, "", qty, "SET", ""])
        rows += [part_row(2, pn, q) for pn, q in kids]
    return rows


def new_rows():
    moved = {(src, pn) for pn, src, _ in CHANGES["이동"]}
    blocks = []
    for (sa, name, qty), kids in ASSEMBLIES:
        qty = CHANGES["Qty"].get((TOP, sa), qty)
        body = []
        for pn, q in kids:
            if (sa, pn) in CHANGES["삭제"] or (sa, pn) in moved:
                continue
            row = part_row(2, pn, CHANGES["Qty"].get((sa, pn), q))
            row[3] = CHANGES["Spec"].get(pn, row[3])
            row[2] = CHANGES["Description"].get(pn, row[2])
            if (sa, pn) == ("SA-02", "P-0013"):
                row[6] = "확인"
            if (sa, pn) == ("SA-10", "P-0004"):
                row[1] = " p-0004 "
            if (sa, pn) == ("SA-01", "P-0002"):
                row[4] = float(row[4])
            if (sa, pn) == ("SA-03", "P-0016"):
                row[3] = row[3].lower()
            body.append(row)
        body += [part_row(2, pn, q) for s, pn, q in CHANGES["추가"] if s == sa]
        body += [part_row(2, pn, dict(ASSEMBLIES_KIDS[src])[pn]) for pn, src, dst in CHANGES["이동"] if dst == sa]
        if sa == "SA-01":
            body.reverse()
        blocks.append([[1, sa, name, "", qty, "SET", ""]] + body)
    i = next(k for k, b in enumerate(blocks) if b[0][1] == "SA-08")
    blocks[i], blocks[i + 1] = blocks[i + 1], blocks[i]
    return [[0, TOP, "제품 A", "", 1, "SET", ""]] + [r for b in blocks for r in b]


ASSEMBLIES_KIDS = {sa: kids for (sa, _, _), kids in ASSEMBLIES}


def answer():
    """(구분, 상위 조립체, Part Number, 바뀐 열). 부품 마스터 변경(Spec, Description)은 그 부품이 남아 있는 자리마다 한 줄."""
    out = [("추가", sa, pn, "") for sa, pn, _ in CHANGES["추가"]]
    out += [("삭제", sa, pn, "") for sa, pn in CHANGES["삭제"]]
    for pn, src, dst in CHANGES["이동"]:
        out += [("삭제", src, pn, ""), ("추가", dst, pn, "")]
    out += [("변경", parent, pn, "Qty") for parent, pn in CHANGES["Qty"]]
    gone = set(CHANGES["삭제"]) | {(src, pn) for pn, src, _ in CHANGES["이동"]}
    for col in ("Spec", "Description"):
        for pn in CHANGES[col]:
            out += [("변경", sa, pn, col) for sa, kids in ASSEMBLIES_KIDS.items()
                    if any(k == pn for k, _ in kids) and (sa, pn) not in gone]
    return sorted(out)


def save_xlsx(path, rows):
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws["A1"] = TITLE
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(COLUMNS)
    for c in ws[3]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=4, min_col=2, max_col=2):
        row[0].number_format = "@"                       # Part Number는 글자 (0102가 102로 바뀌지 않게)
    wb.save(path)


def check(result_csv):
    with open(result_csv, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))[1:]
    got = sorted((r[0], r[1].strip().upper(), r[2].strip().upper(), r[3]) for r in rows)
    want = sorted((a, b.upper(), c.upper(), d) for a, b, c, d in answer())
    missing = [w for w in want if w not in got]
    extra = [g for g in got if g not in want]
    print(f"정답 {len(want)}건, 결과 {len(got)}건")
    print("[OK] 정답과 같습니다" if not missing and not extra else "[FAIL] 정답과 다릅니다")
    for label, items in (("빠진 것", missing), ("정답에 없는 것 (함정에 걸렸을 수 있음)", extra)):
        for it in items:
            print(f"  {label}: {it}")
    return 0 if not missing and not extra else 1


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dir", default=r"C:\AIKIT\dev\work\bom")
    p.add_argument("--check", help="맞대기 결과 CSV를 정답과 비교")
    a = p.parse_args()
    if a.check:
        return check(a.check)
    os.makedirs(a.dir, exist_ok=True)
    rev1, rev2 = old_rows(), new_rows()
    i = next(k for k, r in enumerate(rev2) if r[1] == "SA-03")
    dup = rev2[:i + 2] + [list(rev2[i + 1])] + rev2[i + 2:]       # SA-03 아래 첫 부품이 두 번
    save_xlsx(os.path.join(a.dir, "BOM_A_Rev1.xlsx"), rev1)
    save_xlsx(os.path.join(a.dir, "BOM_A_Rev2.xlsx"), rev2)
    save_xlsx(os.path.join(a.dir, "BOM_A_Rev2_중복.xlsx"), dup)
    with open(os.path.join(a.dir, "BOM_A_정답.csv"), "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([["구분", "상위 조립체", "Part Number", "바뀐 열"]] + answer())
    n_sa = sum(1 for r in rev1 if r[0] == 1)
    print(f"샘플을 만들었습니다: {a.dir}")
    print(f"  옛 판 {len(rev1)}행(중간 조립체 {n_sa}개, 부품 자리 {len(rev1) - 1 - n_sa}개), 새 판 {len(rev2)}행, 정답 {len(answer())}건")
    print(f"  중복 판: SA-03 아래 {rev2[i + 1][1]}이(가) 두 번")
    print("  함정 (정답에 넣지 않음):")
    for t in TRAPS:
        print(f"   - {t}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
