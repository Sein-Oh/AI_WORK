"""BOM 비교 시험 샘플 (정답을 아는 두 판). 사람이 실행하는 개발 도구 (에이전트 계약 대상 아님).

    python make_bom_samples.py [--dir C:\\AIKIT\\dev\\work\\bom]
    python make_bom_samples.py --check "<..._비교결과.csv>"     # 맞대기 결과를 정답과 비교

만드는 파일 (--dir, 기본 C:\\AIKIT\\dev\\work\\bom):
  BOM_A_Rev1.xlsx, BOM_A_Rev2.xlsx   제목 행·빈 행 아래 머리글이 있는 두 판 (함정 포함, 아래 TRAPS)
  BOM_A_Rev2_중복.xlsx              새 판에 같은 품번이 두 번 (맞대기가 멈추고 사용자에게 묻는지)
  BOM_A_정답.csv                    맞대기가 내야 할 결과 (구분, 품번, 바뀐 열)

회사 BOM과 비슷하게 바꾸려면 아래 COLUMNS, PARTS, CHANGES를 고친다. 열 이름을 바꾸면
bom-compare\\scripts\\compare.py의 기준 상수(KEY, COMPARE)도 같이 고친다.
"""
import argparse
import csv
import os
import sys

COLUMNS = ["번호", "품번", "품명", "규격", "수량", "단위", "비고"]
TITLE = "제품 A 부품표 (BOM)"

# 옛 판 부품: (품번, 품명, 규격, 수량, 단위, 비고)
PARTS = [(f"P-{i:04d}", f"부품{i}", f"M{3 + i % 3}x{10 + i % 4 * 5}", i % 5 + 1, "EA", "") for i in range(1, 41)]
PARTS += [("0102", "브래킷", "SUS304 t1.0", 2, "EA", ""), ("0207", "가스켓", "EPDM 50A", 4, "EA", "")]

# 새 판에서 바뀌는 것과 정답 (구분, 품번, 바뀐 열). 함정(TRAPS)은 정답에 넣지 않는다.
CHANGES = {
    "수량": {"P-0003": 9, "P-0011": 7, "P-0025": 5},
    "규격": {"P-0008": "M6x20", "0207": "NBR 50A"},
    "품명": {"P-0014": "부품14-개정"},
    "삭제": ["P-0030", "P-0031", "0102"],
    "추가": [("P-0101", "센서 커버", "PC 2.0t", 1, "EA", "신규"), ("102", "브래킷(신)", "SUS304 t1.2", 2, "EA", "0102 대체")],
}
TRAPS = [
    "행 순서: 새 판은 P-0016~P-0020의 순서가 거꾸로 (바뀐 것 아님)",
    "번호 열: 새 판은 번호를 다시 매김 (비교하지 않는 열)",
    "비고 열: P-0005 비고 '확인' 추가 (비교하지 않는 열)",
    "품번 공백·소문자: 새 판 P-0006은 ' p-0006 ' (같은 부품)",
    "수량 표기: 새 판 P-0007 수량 3 -> 3.0 (같은 값, P-0007 옛 수량 3)",
    "규격 대소문자: 새 판 P-0009 규격 대문자 X (같은 규격)",
    "앞자리 0: 옛 판 0102와 새 판 102는 다른 품번 (0102 삭제, 102 추가)",
]


def new_parts():
    rows = []
    for pn, name, spec, qty, unit, note in PARTS:
        if pn in CHANGES["삭제"]:
            continue
        qty = CHANGES["수량"].get(pn, qty)
        spec = CHANGES["규격"].get(pn, spec)
        name = CHANGES["품명"].get(pn, name)
        if pn == "P-0005":
            note = "확인"
        if pn == "P-0006":
            pn = " p-0006 "
        if pn == "P-0007":
            qty = float(qty)
        if pn == "P-0009":
            spec = spec.upper()
        rows.append((pn, name, spec, qty, unit, note))
    i = next(k for k, r in enumerate(rows) if r[0] == "P-0016")
    rows[i:i + 5] = rows[i:i + 5][::-1]
    return rows + CHANGES["추가"]


def answer():
    out = [("추가", pn, "") for pn, *_ in CHANGES["추가"]] + [("삭제", pn, "") for pn in CHANGES["삭제"]]
    for col in ("품명", "규격", "수량"):
        out += [("변경", pn, col) for pn in CHANGES[col]]
    return sorted(out)


def save_xlsx(path, parts):
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
    for n, (pn, name, spec, qty, unit, note) in enumerate(parts, start=1):
        ws.append([n, pn, name, spec, qty, unit, note])
    for row in ws.iter_rows(min_row=4, min_col=2, max_col=2):
        row[0].number_format = "@"                       # 품번은 글자 (0102가 102로 바뀌지 않게)
    wb.save(path)


def check(result_csv):
    with open(result_csv, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))[1:]
    got = sorted((r[0], r[1].strip(), r[2]) for r in rows)
    want = answer()
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
    rev2 = new_parts()
    save_xlsx(os.path.join(a.dir, "BOM_A_Rev1.xlsx"), PARTS)
    save_xlsx(os.path.join(a.dir, "BOM_A_Rev2.xlsx"), rev2)
    save_xlsx(os.path.join(a.dir, "BOM_A_Rev2_중복.xlsx"), rev2 + [rev2[2]])
    with open(os.path.join(a.dir, "BOM_A_정답.csv"), "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([["구분", "품번", "바뀐 열"]] + answer())
    print(f"샘플을 만들었습니다: {a.dir}")
    print(f"  옛 판 {len(PARTS)}개, 새 판 {len(rev2)}개 부품, 정답 {len(answer())}건")
    print("  함정 (정답에 넣지 않음):")
    for t in TRAPS:
        print(f"   - {t}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
