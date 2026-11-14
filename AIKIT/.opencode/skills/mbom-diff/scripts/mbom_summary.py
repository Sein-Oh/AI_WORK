# -*- coding: utf-8 -*-
r"""MBOM 두 판 비교 결과(pandas-kit diff를 경로 키로 export한 csv)를 사람이 읽을 변경 목록으로 묶는다 (mbom-diff 스킬).

    python "<스킬 폴더>\scripts\mbom_summary.py" --file "<변경내용.csv>"

- 대체 후보: 같은 상위 아래 같은 자리(앞이나 뒤의 형제 품번이 같음)에서 품번이 삭제되고 다른 품번이 추가됨 (품번만 봄, 품명은 보지 않음.
  삭제와 추가에 모두 있는 품번은 옮겨진 것이라 짝짓지 않음)
  (옆의 <표>_경로.csv로 형제 순서를 봄, 없으면 삭제 하나·추가 하나일 때만. 짐작이라 사용자 확인이 필요).
- 추가·삭제: 나머지. 다른 상위 아래로 옮겨진 품목도 삭제·추가로 둔다 (사용자 결정 2026-11-14: 이동으로 묶지 않음). 바뀐 칸: 품번마다 열: 옛 값 -> 새 값.
결과: 입력 옆 <이름>_요약.txt, stdout은 1줄 JSON (preview를 사용자에게 그대로 알림).
"""
import argparse
import csv
import json
import os
import sys

PREVIEW_LINES = 30


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


def read_rows(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    stop(f"{os.path.basename(path)}을(를) 읽지 못했습니다 (UTF-8, CP949 아님).")


def num(v):
    """2.0 -> 2 (pandas가 정수 열을 소수로 적은 것만), 나머지는 그대로."""
    v = (v or "").strip()
    try:
        f = float(v)
    except ValueError:
        return v
    return str(int(f)) if f == int(f) and "e" not in v.lower() else v


def split(path):
    parent, _, last = path.rpartition("/")
    return parent, last.partition("#")[0]


def where(parent):
    return f"{parent} 아래" if parent else "최상위"


def summarize(rows, kind, key, ocol, ncol, order=None):
    dels = [(r[key], r.get(ocol) or "") for r in rows if r[kind] == "삭제"]
    adds = [(r[key], r.get(ncol) or "") for r in rows if r[kind] == "추가"]
    lines, counts = [], {"대체 후보": 0, "추가": 0, "삭제": 0, "바뀐 칸": 0}
    # 대체 후보: 같은 상위 아래, 앞이나 뒤의 형제 품번이 같은 자리 (형제 순서를 모르면 삭제 하나·추가 하나일 때만)
    # 같은 품번이 삭제와 추가에 모두 있으면 옮겨진 품목이라 대체 짝에서 뺌 (삭제·추가로 둠)
    moved = {split(x[0])[1] for x in dels} & {split(x[0])[1] for x in adds}
    for d in [x for x in dels if split(x[0])[1] not in moved]:
        parent = split(d[0])[0]
        same_a = [x for x in adds if split(x[0])[0] == parent and split(x[0])[1] not in moved]
        if order:
            spot = order[0].get(d[0])
            same_a = [x for x in same_a if spot and (order[1].get(x[0]) or (None, None))[0] == spot[0]] or \
                     [x for x in same_a if spot and (order[1].get(x[0]) or (None, None))[1] == spot[1]]
        elif len([x for x in dels if split(x[0])[0] == parent and split(x[0])[1] not in moved]) != 1:
            continue
        if len(same_a) != 1:
            continue
        a = same_a[0]
        lines.append(f"[대체 후보] {split(d[0])[1]} -> {split(a[0])[1]}: {where(parent)} 같은 자리 "
                     f"(옛 엑셀 행 {d[1]}, 새 엑셀 행 {a[1]})")
        dels.remove(d)
        adds.remove(a)
        counts["대체 후보"] += 1
    for path, row in adds:
        parent, pn = split(path)
        lines.append(f"[추가] {pn}: {where(parent)} (새 엑셀 행 {row})")
    for path, row in dels:
        parent, pn = split(path)
        lines.append(f"[삭제] {pn}: {where(parent)} (옛 엑셀 행 {row})")
    counts["추가"], counts["삭제"] = len(adds), len(dels)
    changed = {}
    for r in rows:
        if r[kind] == "바뀜":
            changed.setdefault(r[key], []).append(r)
    for path, rs in changed.items():
        parent, pn = split(path)
        cells = ", ".join(f"{r['열']} {num(r['옛 값']) or '(빈 칸)'} -> {num(r['새 값']) or '(빈 칸)'}" for r in rs)
        lines.append(f"[바뀜] {pn} ({where(parent)}): {cells} (옛 엑셀 행 {rs[0].get(ocol)}, 새 엑셀 행 {rs[0].get(ncol)})")
        counts["바뀐 칸"] += len(rs)
    return lines, counts, len(changed)


def siblings(path):
    """<표>_경로.csv -> {경로: (앞 형제 품번, 뒤 형제 품번)} (같은 상위 아래 표 순서)."""
    kids = {}
    for r in read_rows(path):
        kids.setdefault(split(r["경로"])[0], []).append(r["경로"])
    spot = {}
    for ks in kids.values():
        for i, k in enumerate(ks):
            spot[k] = (split(ks[i - 1])[1] if i else "(처음)", split(ks[i + 1])[1] if i + 1 < len(ks) else "(끝)")
    return spot


def main():
    ap = Parser(description="MBOM 비교 결과를 변경 목록으로 (mbom-diff 스킬)")
    ap.add_argument("--file", required=True, help="pandas-kit diff 결과를 export한 .csv")
    a = ap.parse_args()
    path = os.path.abspath(a.file.strip().strip('"'))
    if not os.path.isfile(path):
        stop(f"파일이 없습니다: {path}")
    rows = read_rows(path)
    cols = list(rows[0].keys()) if rows else []
    rowcols = [c for c in cols if c.endswith(" 행")]
    if rows and ("구분" not in cols or "경로" not in cols or len(rowcols) != 2):
        stop(f"pandas-kit diff --on 경로 결과를 export한 csv가 아닙니다 (있는 열: {', '.join(cols)}).")
    # mbom_path가 만든 <표>_경로.csv (diff의 'oldp 행' -> old_경로.csv)가 옆에 있으면 형제 순서로 대체 자리를 찾음
    files = [os.path.join(os.path.dirname(path), c[:-2].removesuffix("p") + "_경로.csv") for c in rowcols]
    order = [siblings(f) for f in files] if rows and all(os.path.isfile(f) for f in files) else None
    lines, counts, nrows = summarize(rows, "구분", "경로", *rowcols, order=order) if rows else ([], {}, 0)
    head = ("두 판이 같습니다 (찾은 변경 없음)." if not lines else
            "변경 " + ", ".join(f"{k} {v}" + ("곳" if k == "바뀐 칸" else "건") for k, v in counts.items() if v)
            + (f" ({nrows}개 품목)" if counts.get("바뀐 칸") else "") + ".")
    save = os.path.splitext(path)[0] + "_요약.txt"
    with open(save, "w", encoding="utf-8-sig") as f:
        f.write("\n".join([head] + lines) + "\n")
    with open(save, encoding="utf-8-sig") as f:
        verified = len(f.read().splitlines()) == len(lines) + 1
    shown = lines[:PREVIEW_LINES]
    preview = "\n".join([head] + shown + ([f"(나머지 {len(lines) - len(shown)}줄은 {save})"] if len(lines) > len(shown) else []))
    directive = ("preview를 그대로 사용자에게 알리세요." + (" [대체 후보]는 도구의 짐작이니 맞는지 사용자에게 확인하세요."
                                                    if counts.get("대체 후보") else ""))
    out({"status": "ok", "command": "mbom_summary", "preview": preview, "count": len(lines), "counts": counts,
         "lines": shown, "truncated": len(lines) > len(shown), "output_file": save,
         "agent_directive": ("VERIFIED: " if verified else "UNVERIFIED: 저장한 요약 파일을 확인하지 못함. ") + directive})


if __name__ == "__main__":
    main()
