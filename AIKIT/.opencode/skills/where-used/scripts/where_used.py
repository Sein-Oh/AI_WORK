# -*- coding: utf-8 -*-
r"""BOM에서 품목이 어느 상위에 쓰였는지 찾는다 (where-used 스킬의 스크립트, pandas-kit 기반 명령과 함께 씀).

    python "<스킬 폴더>\scripts\where_used.py" --file "<BOM export .csv>" --find "<품번이나 품명>"

BOM에 상위 열이 없을 때, 어떤 행의 상위 = 위쪽 행 중 가장 가까운, Level이 더 작은 행 (업무 규칙이라 도구가 아니라 스킬에 둠).
경로 = 최상위부터 품번을 /로 이은 것. 같은 상위 아래 같은 품번이 둘 이상이면 #순번 (#1, #2).
제품 1대당 수량 = 그 행과 위쪽 상위들의 수량을 곱한 값 (스크랩·유효 기간·대체 그룹은 보지 않음).
결과: 입력 옆 <이름>_쓰인곳.csv, stdout은 1줄 JSON (preview에 사용자에게 그대로 전할 줄).
"""
import argparse
import csv
import json
import os
import re
import sys

ROW = "엑셀 행"
LEVEL_NAMES = ("level", "lvl", "레벨", "수준")
PART_NAMES = ("part number", "part no", "partnumber", "item", "품번", "자재번호", "자재 번호", "품목번호")
QTY_NAMES = ("qty", "quantity", "qpa", "수량", "소요량")
DESC_NAMES = ("description", "desc", "품명", "품목명", "내역", "규격")
UOM_NAMES = ("uom", "unit", "단위")
LINE_MAX = 30


def out(obj, code=0):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.exit(code)


def stop(msg):
    out({"status": "error", "error_type": "input", "message": msg,
         "agent_directive": "STOP_RETRY: message를 사용자에게 알리고 멈추세요."}, 1)


class Parser(argparse.ArgumentParser):
    def error(self, message):
        out({"status": "error", "error_type": "usage", "message": f"인자 오류: {message}",
             "agent_directive": "STOP_RETRY: 스킬의 명령을 그대로 다시 확인하세요."}, 2)

    def print_help(self, file=None):
        sys.stderr.write(self.format_help())


def find_col(cols, names, given, what, needed=True):
    if given:
        if given in cols:
            return given
        stop(f"'{given}' 열이 없습니다. 있는 열: {', '.join(cols)}")
    for c in cols:
        if c.strip().lower() in names:
            return c
    if needed:
        stop(f"{what} 열을 찾지 못했습니다. 있는 열: {', '.join(cols)}. --level/--part/--qty로 열 이름을 주세요.")
    return None


def read_rows(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    stop(f"{os.path.basename(path)}을(를) 읽지 못했습니다 (UTF-8, CP949 아님).")


def level_of(v):
    try:
        f = float(str(v).strip())
    except ValueError:
        return None
    return int(f) if f == int(f) and f >= 0 else None


def qty_of(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, AttributeError):
        return None


def number(v):
    return str(int(v)) if v == int(v) else f"{v:g}"


def compute(rows, lcol, pcol, qcol):
    """행마다 (엑셀 행, 상위 품번, 경로, 그 자리 수량, 제품 1대당 수량). 상위는 위쪽의 더 작은 Level."""
    stack, count, raw, bad, jumps, noqty = [], {}, [], [], [], []       # stack: [(level, 경로 조각, 누적 수량)]
    for r in rows:
        lv, pn = level_of(r.get(lcol)), str(r.get(pcol) or "").strip()
        if lv is None or not pn:
            bad.append(r.get(ROW) or "?")
            continue
        while stack and stack[-1][0] >= lv:
            stack.pop()
        if stack and lv > stack[-1][0] + 1:
            jumps.append(r.get(ROW) or "?")
        parent, above = (stack[-1][1], stack[-1][2]) if stack else ((), 1.0)
        n = count[(parent, pn)] = count.get((parent, pn), 0) + 1
        parts = parent + (f"{pn}#{n}",)
        q = qty_of(r.get(qcol))
        if q is None:
            noqty.append(r.get(ROW) or "?")
        total = None if q is None or above is None else above * q
        stack.append((lv, parts, total))
        raw.append((r, parts, q, total))

    def shown(parts):       # 같은 상위 아래 한 번만 나오는 품번은 #1을 뗌
        return "/".join(p if count[(parts[:i], p.rpartition("#")[0])] > 1 else p.rpartition("#")[0]
                        for i, p in enumerate(parts))
    res = [(r, parts[-2].rpartition("#")[0] if len(parts) > 1 else "", shown(parts), q, total)
           for r, parts, q, total in raw]
    return res, bad, jumps, noqty


def main():
    ap = Parser(description="BOM에서 품목이 쓰인 상위와 제품 1대당 수량 (where-used 스킬)")
    ap.add_argument("--file", required=True, help="pandas-kit export로 만든 BOM .csv")
    ap.add_argument("--find", required=True, help="찾을 품번(그대로 같은 것) 또는 품명에 든 글자")
    ap.add_argument("--level", help="Level 열 이름 (생략하면 찾음)")
    ap.add_argument("--part", help="품번 열 이름 (생략하면 찾음)")
    ap.add_argument("--qty", help="수량 열 이름 (생략하면 찾음)")
    a = ap.parse_args()
    path = os.path.abspath(a.file.strip().strip('"'))
    if not os.path.isfile(path):
        stop(f"파일이 없습니다: {path}")
    rows = read_rows(path)
    if not rows:
        stop(f"{os.path.basename(path)}에 행이 없습니다.")
    cols = list(rows[0].keys())
    if ROW not in cols:
        stop(f"{os.path.basename(path)}에 '{ROW}' 열이 없습니다. pandas-kit export로 만든 csv를 주세요.")
    pcol = find_col(cols, PART_NAMES, a.part, "품번")
    qcol = find_col(cols, QTY_NAMES, a.qty, "수량")
    dcol = find_col(cols, DESC_NAMES, None, "품명", needed=False)
    ucol = find_col(cols, UOM_NAMES, None, "단위", needed=False)
    res, bad, jumps, noqty = compute(rows, find_col(cols, LEVEL_NAMES, a.level, "Level"), pcol, qcol)

    key = a.find.strip().lower()
    hits = [(r, par, path_, q, total) for r, par, path_, q, total in res
            if str(r.get(pcol) or "").strip().lower() == key
            or (dcol and key in str(r.get(dcol) or "").lower())]
    name = {}
    for r, *_ in hits:
        name.setdefault(str(r.get(pcol) or "").strip(), str(r.get(dcol) or "").strip() if dcol else "")
    save = os.path.splitext(path)[0] + "_쓰인곳.csv"
    head = [ROW, "품번", "품명", "상위 품번", "경로", "그 자리 수량", "제품 1대당 수량", "단위"]
    with open(save, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(head)
        for r, par, path_, q, total in hits:
            w.writerow([r.get(ROW), r.get(pcol), r.get(dcol) if dcol else "", par, path_,
                        "" if q is None else number(q), "" if total is None else number(total),
                        r.get(ucol) if ucol else ""])
    warnings = []
    if bad:
        warnings.append(f"Level이 숫자가 아니거나 품번이 빈 행 {len(bad)}개는 보지 않았습니다 (엑셀 행 {', '.join(map(str, bad[:10]))}).")
    if jumps:
        warnings.append(f"Level이 두 단계 이상 건너뛴 행 {len(jumps)}개 (엑셀 행 {', '.join(map(str, jumps[:10]))}). "
                        "상위는 바로 위의 더 작은 Level로 봤습니다.")
    if noqty:
        warnings.append(f"수량이 숫자가 아닌 행 {len(noqty)}개 (엑셀 행 {', '.join(map(str, noqty[:10]))})는 "
                        "제품 1대당 수량을 비웠습니다.")
    if not hits:
        out({"status": "ok", "tool": "where-used", "command": "where_used", "verified": True, "count": 0,
             "preview": f"'{a.find}'을(를) 품번이나 품명에서 찾지 못했습니다 (BOM {len(res)}행).",
             "lines": [], "output_file": save, "warnings": warnings or None,
             "agent_directive": "VERIFIED: 0건. 찾은 품목이 없습니다. preview를 그대로 알리고 다시 실행하지 마세요. "
                                "품번이나 품명의 일부를 --find에 주세요."})

    lines, totals = [], {}
    for r, par, path_, q, total in hits:
        pn = str(r.get(pcol) or "").strip()
        unit = (str(r.get(ucol) or "").strip() if ucol else "")
        parent_name = ""
        for rr, *_ in res:                                   # 상위 품번의 품명 (있으면)
            if dcol and str(rr.get(pcol) or "").strip() == par:
                parent_name = f" {str(rr.get(dcol) or '').strip()}"
                break
        where = f"{par}{parent_name} 아래" if par else "최상위"
        amount = "수량 모름" if total is None else f"{number(total)}{(' ' + unit) if unit else ''}"
        lines.append(f"[쓰임] {pn}: {where} {amount} (경로 {path_}, 엑셀 행 {r.get(ROW)})")
        if total is not None:
            totals.setdefault((pn, unit), [0.0, 0])
            totals[(pn, unit)][0] += total
            totals[(pn, unit)][1] += 1
    sums = [f"{pn} {name.get(pn) or ''}".strip() + f": {len(hits) if len(totals) == 1 else n}곳, "
            f"제품 1대당 {number(t)}{(' ' + u) if u else ''}" for (pn, u), (t, n) in totals.items()]
    preview = f"'{a.find}': {len(hits)}곳에 쓰입니다. " + " / ".join(sums)
    shown_lines = lines[:LINE_MAX]
    out({"status": "ok", "tool": "where-used", "command": "where_used", "verified": True,
         "preview": preview + ("\n" + "\n".join(shown_lines) if shown_lines else ""),
         "count": len(hits), "lines": shown_lines, "totals": sums, "truncated": len(lines) > LINE_MAX,
         "output_file": save, "warnings": warnings or None,
         "agent_directive": "VERIFIED: preview의 줄을 사용자에게 그대로 알리세요 (줄을 직접 세거나 고쳐 쓰지 마세요). "
                            "같은 명령을 다시 실행하지 마세요." + (" 잘렸으면 전체는 output_file에 있습니다." if len(lines) > LINE_MAX else "")
                            + (" warnings도 그대로 알리세요." if warnings else "")})


if __name__ == "__main__":
    main()
