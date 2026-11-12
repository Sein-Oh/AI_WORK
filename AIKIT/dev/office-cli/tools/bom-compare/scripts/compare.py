"""bom-compare 맞대기 (사용자 조합 스킬 예시의 스크립트, office-cli 도구가 아님).

excel.py read --save로 저장한 두 판의 CSV를 "상위 조립체 + Part Number"(자리)로 맞대어 추가·삭제·변경을 찾는다.
상위 조립체는 Level 열로 정한다 (자기보다 Level이 작은 가장 가까운 위 행). 같은 부품이 여러 조립체에 쓰이면 자리마다 따로 비교한다.
같은 판에서 같은 자리가 두 번 나오면 맞대지 않고 멈춘다 (판단 기준 4).
판단 기준은 SKILL.md "판단 기준"과 같고, 회사 BOM에 맞게 바꿀 곳은 아래 "기준" 상수뿐이다.

입력: --old/--new  read --save가 만든 CSV (열 고르기 --columns 없이 저장한 것)
      --new-range  새 판 read 결과 JSON의 range 값 그대로 (예: A1:G45). CSV 첫 칸이 시트의 어디인지 알려 줌
출력 (새 판 CSV와 같은 폴더):
  <새 판>_비교결과.csv     구분, 상위 조립체, Part Number, 바뀐 열, 옛 값, 새 값, 옛 행, 새 행 (같은 자리는 넣지 않음)
  <새 판>_칠할칸_변경.csv   바뀐 칸 주소 (excel.py style --file)
  <새 판>_칠할칸_추가.csv   추가된 자리의 Part Number 칸 주소 (excel.py style --file)
  <새 판>_메모.csv         칸, 메모 (excel.py comment --action add --file)
stdout: 마지막 줄에 JSON 1줄 (status, preview, counts, 파일 경로). 진행 글은 stderr.
"""
import argparse
import csv
import json
import os
import re
import sys

# ---- 기준 (회사 BOM에 맞게 여기만 고친다) ------------------------------------------
LEVEL = "Level"                    # 계층 열 (0 최상위, 1, 2 ...). 평면 BOM이면 None (상위 조립체 없이 품번만으로 맞댐)
KEY = "Part Number"                # 부품을 가르는 열
COMPARE = ["Description", "Spec", "Qty"]   # 비교할 열 (그 밖의 열: Unit, Remark 등은 보지 않음)
NUMBER_COLS = {"Qty"}              # 숫자로 비교할 열 (1과 1.0은 같음)
IGNORE_CASE_COLS = {"Spec"}        # 공백·대소문자를 무시할 열 (M3x10 = M3X10)
HEADER_SCAN = 15                   # 머리글 행을 찾을 위쪽 행 수 (제목 행이 있어도 됨)
# -----------------------------------------------------------------------------------


def out(payload, code=0):
    print(json.dumps(payload, ensure_ascii=False))
    sys.exit(code)


def fail(message):
    out({"status": "error", "agent_directive": "STOP_RETRY: " + message, "message": message}, 1)


def read_csv(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.reader(f))
        except UnicodeDecodeError:
            continue
    fail(f"CSV를 읽지 못했습니다: {path}")


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def norm_key(v):
    return str(v).strip().upper()             # 공백·대소문자 무시, 앞자리 0은 그대로 (0102와 102는 다른 품번)


def same(col, a, b):
    a, b = str(a).strip(), str(b).strip()
    if col in NUMBER_COLS:
        try:
            return float(a.replace(",", "")) == float(b.replace(",", ""))
        except ValueError:
            pass
    if col in IGNORE_CASE_COLS:
        return re.sub(r"\s+", "", a).upper() == re.sub(r"\s+", "", b).upper()
    return a == b


def load(path, label):
    rows = read_csv(path)
    need = ([LEVEL] if LEVEL else []) + [KEY] + COMPARE
    for hi, row in enumerate(rows[:HEADER_SCAN]):
        names = [str(c).strip() for c in row]
        if all(n in names for n in need):
            idx = {n: names.index(n) for n in need}
            break
    else:
        fail(f"{label} CSV의 위 {HEADER_SCAN}행에서 머리글({', '.join(need)})을 찾지 못했습니다. 열 이름이 다르면 compare.py의 기준 상수를 고치세요.")
    items, dup, stack = {}, {}, []          # stack: 지금 행의 상위 조립체 후보 [(level, Part Number)]
    for i in range(hi + 1, len(rows)):
        row = rows[i] + [""] * (len(rows[hi]) - len(rows[i]))
        pn = row[idx[KEY]].strip()
        if not pn:
            continue
        level = 0
        if LEVEL:
            try:
                level = int(float(str(row[idx[LEVEL]]).strip()))
            except ValueError:
                fail(f"{label} {i + 1}번째 줄의 {LEVEL} 값 '{row[idx[LEVEL]]}'이(가) 숫자가 아닙니다. {LEVEL}은 0, 1, 2 같은 숫자여야 합니다.")
            while stack and stack[-1][0] >= level:
                stack.pop()
        parent = stack[-1][1] if stack else ""
        if LEVEL:
            stack.append((level, pn))
        key = (norm_key(parent), norm_key(pn))
        if key in items:
            dup.setdefault(key, [items[key]["row"]]).append(i)
            continue
        items[key] = {"row": i, "raw": pn, "parent": parent, **{c: row[idx[c]] for c in COMPARE}}
    return items, idx, dup


def main():
    p = argparse.ArgumentParser(description="BOM 두 판 맞대기 (bom-compare 스킬)")
    p.add_argument("--old", required=True, help="옛 판 CSV (excel.py read --save)")
    p.add_argument("--new", required=True, help="새 판 CSV (excel.py read --save)")
    p.add_argument("--new-range", required=True, help="새 판 read 결과의 range 값 그대로 (예: A1:G45)")
    p.add_argument("--old-range", default="A1", help="옛 판 read 결과의 range 값 (행 번호 표시용, 기본 A1)")
    a = p.parse_args()
    for f in (a.old, a.new):
        if not os.path.isfile(f):
            fail(f"파일이 없습니다: {f}")
    m_new, m_old = (re.match(r"\$?([A-Za-z]+)\$?(\d+)", r) for r in (a.new_range, a.old_range))
    if not m_new or not m_old:
        fail("--new-range·--old-range는 read 결과의 range 값 그대로 주세요 (예: A1:G45).")
    new_c0, new_r0 = col_number(m_new.group(1)), int(m_new.group(2))
    old_r0 = int(m_old.group(2))

    old, _, dup_old = load(a.old, "옛 판")
    new, idx_new, dup_new = load(a.new, "새 판")
    if dup_old or dup_new:
        shown = "; ".join(f"{side} {(k[0] + ' 아래 ') if k[0] else ''}{k[1]}: 엑셀 {', '.join(str(r0 + r) for r in v)}행"
                          for side, d, r0 in (("옛 판", dup_old, old_r0), ("새 판", dup_new, new_r0)) for k, v in list(d.items())[:10])
        fail(f"같은 판 안에서 같은 상위 조립체 아래 같은 {KEY}가 두 번 이상 나옵니다 ({shown}). "
             "맞대지 않았습니다. 어느 줄이 맞는지 사용자가 정한 뒤 다시 비교하세요.")

    def cell(item, col):
        return f"{col_letter(new_c0 + idx_new[col])}{new_r0 + item['row']}"

    result, changed_cells, added_cells, memos = [], [], [], []
    for key, n in new.items():
        if key not in old:
            result.append(["추가", n["parent"], n["raw"], "", "", "", "", new_r0 + n["row"]])
            added_cells.append(cell(n, KEY))
            continue
        o = old[key]
        for col in COMPARE:
            if not same(col, o[col], n[col]):
                result.append(["변경", n["parent"], n["raw"], col, o[col], n[col], old_r0 + o["row"], new_r0 + n["row"]])
                changed_cells.append(cell(n, col))
                where = f"{n['parent']} 아래, " if n["parent"] else ""
                memos.append([cell(n, col), f"[BOM 비교] {col}: {o[col]} -> {n[col]} ({where}옛 판 {old_r0 + o['row']}행)"])
    for key, o in old.items():
        if key not in new:
            result.append(["삭제", o["parent"], o["raw"], "", "", "", old_r0 + o["row"], ""])

    base = os.path.splitext(a.new)[0]
    files = {"result_file": base + "_비교결과.csv", "changed_cells_file": base + "_칠할칸_변경.csv",
             "added_cells_file": base + "_칠할칸_추가.csv", "memo_file": base + "_메모.csv"}
    for key, header, rows in (("result_file", ["구분", "상위 조립체", KEY, "바뀐 열", "옛 값", "새 값", "옛 행", "새 행"], result),
                              ("changed_cells_file", ["칸"], [[c] for c in changed_cells]),
                              ("added_cells_file", ["칸"], [[c] for c in added_cells]),
                              ("memo_file", ["칸", "메모"], memos)):
        with open(files[key], "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows([header] + rows)
    counts = {k: sum(1 for r in result if r[0] == k) for k in ("추가", "삭제", "변경")}
    parts = len({(r[1], r[2]) for r in result if r[0] == "변경"})
    preview = f"추가 {counts['추가']}, 삭제 {counts['삭제']}, 변경 {parts}자리({counts['변경']}칸)"
    print(f"[INFO] 옛 판 {len(old)}자리, 새 판 {len(new)}자리를 상위 조립체 + {KEY}로 맞댐", file=sys.stderr)
    out({"status": "ok", "preview": preview, "counts": counts, "changed_parts": parts,
         "agent_directive": "VERIFIED: 맞대기를 마쳤습니다. preview와 result_file의 앞 10줄을 사용자에게 보여 주세요." if result
         else "VERIFIED: 0건. 두 판의 비교 열이 모두 같습니다 (찾은 항목 없음).",
         "added_cells": len(added_cells), "changed_cells": len(changed_cells), **files})


if __name__ == "__main__":
    main()
