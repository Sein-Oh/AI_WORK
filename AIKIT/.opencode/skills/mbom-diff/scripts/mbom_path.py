# -*- coding: utf-8 -*-
r"""MBOM 행마다 상위 품번과 경로를 계산한다 (mbom-diff 스킬의 스크립트, pandas-kit 기반 명령과 함께 씀).

    python "<스킬 폴더>\scripts\mbom_path.py" --old "<옛 판 export .csv>" --new "<새 판 export .csv>"

MBOM에 상위 열이 없을 때, 어떤 행의 상위 = 위쪽 행 중 가장 가까운, Level이 더 작은 행 (업무 규칙이라 도구가 아니라 스킬에 둠).
경로 = 최상위부터 품번을 /로 이은 것. 같은 상위 아래 같은 품번이 둘 이상이면 모두 #순번 (#1, #2).
결과: 입력 옆 <이름>_경로.csv (열: 엑셀 행, 상위 품번, 경로). stdout은 1줄 JSON, 다음에 실행할 pandas-kit 명령
(경로 붙이기, 경로로 diff, 결과 export)을 next_commands로 준다. csv 파일 이름이 pandas-kit 표 이름 (export --save old.csv).
"""
import argparse
import csv
import json
import os
import re
import sys



def pandas_kit_command():
    r"""이 스크립트 위쪽 폴더에서 pandas-kit/pandas_kit.py를 찾음 (C:\AIKIT\.opencode\skills\<스킬>\scripts -> C:\AIKIT)."""
    d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(8):
        p = os.path.join(d, "pandas-kit", "pandas_kit.py")
        if os.path.isfile(p):
            return f'python "{p}"'
        d = os.path.dirname(d)
    return r'python "C:\AIKIT\pandas-kit\pandas_kit.py"'


ROW = "엑셀 행"
LEVEL_NAMES = ("level", "lvl", "레벨", "수준")
PART_NAMES = ("part number", "part no", "partnumber", "item", "품번", "자재번호", "자재 번호", "품목번호")


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


def find_col(cols, names, given, what):
    if given:
        if given in cols:
            return given
        stop(f"'{given}' 열이 없습니다. 있는 열: {', '.join(cols)}")
    for c in cols:
        if c.strip().lower() in names:
            return c
    stop(f"{what} 열을 찾지 못했습니다. 있는 열: {', '.join(cols)}. --level/--part로 열 이름을 주세요.")


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


def compute(rows, lcol, pcol):
    """(엑셀 행, Level, 상위 품번, 경로) 목록, Level·품번이 빈 행, Level을 건너뛴 행."""
    stack, count, raw, bad, jumps = [], {}, [], [], []     # stack: [(level, 경로 조각 목록)]
    for r in rows:
        lv, pn = level_of(r.get(lcol)), str(r.get(pcol) or "").strip()
        if lv is None or not pn:
            bad.append(r.get(ROW) or "?")
            continue
        while stack and stack[-1][0] >= lv:
            stack.pop()
        if stack and lv > stack[-1][0] + 1:
            jumps.append(r.get(ROW) or "?")
        parent = stack[-1][1] if stack else ()
        n = count[(parent, pn)] = count.get((parent, pn), 0) + 1
        parts = parent + (f"{pn}#{n}",)
        stack.append((lv, parts))
        raw.append((r.get(ROW), lv, parts))

    def shown(parts):       # 같은 상위 아래 한 번만 나오는 품번은 #1을 뗌
        return "/".join(p if count[(parts[:i], p.rpartition("#")[0])] > 1 else p.rpartition("#")[0]
                        for i, p in enumerate(parts))
    res = [(row, lv, parts[-2].rpartition("#")[0] if len(parts) > 1 else "", shown(parts)) for row, lv, parts in raw]
    return res, bad, jumps


def one(path, level, part):
    """csv 하나 -> <이름>_경로.csv. (표 이름, 저장 경로, 행 수, #순번 품번, warnings)."""
    path = os.path.abspath(path.strip().strip('"'))
    if not os.path.isfile(path):
        stop(f"파일이 없습니다: {path}")
    # 표 이름 = 파일 이름 (export가 같은 이름 파일이 있어 old_20261114_101500.csv로 저장했으면 그 꼬리는 뗌)
    table = re.sub(r"_\d{8}_\d{6}$", "", os.path.splitext(os.path.basename(path))[0])
    rows = read_rows(path)
    if not rows:
        stop(f"{os.path.basename(path)}에 행이 없습니다.")
    cols = list(rows[0].keys())
    if ROW not in cols:
        stop(f"{os.path.basename(path)}에 '{ROW}' 열이 없습니다. pandas-kit export로 만든 csv를 주세요.")
    res, bad, jumps = compute(rows, find_col(cols, LEVEL_NAMES, level, "Level"), find_col(cols, PART_NAMES, part, "품번"))
    save = os.path.splitext(path)[0] + "_경로.csv"
    with open(save, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([ROW, "상위 품번", "경로"])
        w.writerows((r, p, k) for r, _lv, p, k in res)
    with open(save, encoding="utf-8-sig", newline="") as f:
        back = sum(1 for _ in csv.reader(f)) - 1
    warnings = []
    if bad:
        warnings.append(f"{table}: Level이 숫자가 아니거나 품번이 빈 행 {len(bad)}개는 경로를 만들지 않았습니다 "
                        f"(엑셀 행 {', '.join(map(str, bad[:10]))}).")
    if jumps:
        warnings.append(f"{table}: Level이 두 단계 이상 건너뛴 행 {len(jumps)}개 (엑셀 행 {', '.join(map(str, jumps[:10]))}). "
                        "상위는 바로 위의 더 작은 Level로 봤습니다.")
    multi = sorted({p.rpartition("/")[2] for *_, p in res if "#" in p.rpartition("/")[2]})
    return table, save, len(res), back == len(res), multi, warnings


def main():
    ap = Parser(description="MBOM 두 판의 행마다 상위 품번과 경로 (mbom-diff 스킬)")
    ap.add_argument("--old", required=True, help="옛 판: pandas-kit export로 만든 .csv (파일 이름 = 표 이름)")
    ap.add_argument("--new", required=True, help="새 판: pandas-kit export로 만든 .csv")
    ap.add_argument("--level", help="Level 열 이름 (기본: 찾음)")
    ap.add_argument("--part", help="품번 열 이름 (기본: 찾음)")
    a = ap.parse_args()
    done = [one(f, a.level, a.part) for f in (a.old, a.new)]
    if done[0][0] == done[1][0]:
        stop("두 csv의 파일 이름(표 이름)이 같습니다. 옛 판과 새 판을 따로 export하세요.")
    pk, q = pandas_kit_command(), (lambda s: '"' + s + '"')
    nxt = []
    for table, save, *_ in done:
        nxt += [{"command": f"{pk} select --file {q(save)} --name {table}k", "reason": f"{table} 경로 표 읽기"},
                {"command": f'{pk} join --left {table} --right {table}k --on "엑셀 행=엑셀 행번호" --name {table}p',
                 "reason": f"경로를 {table}에 붙이기"}]
    (o, *_), (n, *_) = done
    nxt += [{"command": f"{pk} diff --old {o}p --new {n}p --on 경로 --name chg", "reason": "경로로 두 판 비교"},
            {"command": f"{pk} export --from chg --save 변경내용.csv", "reason": "비교 결과를 csv로 (요약 스크립트가 읽음)"}]
    warnings = [w for d in done for w in d[5]]
    multi = sorted({m for d in done for m in d[4]})
    preview = (f"두 판의 상위 품번과 경로를 계산했습니다 ({o} {done[0][2]}행, {n} {done[1][2]}행)."
               + (f" 같은 상위 아래 두 번 이상 나와 #순번을 붙인 품번: {', '.join(multi[:5])}." if multi else ""))
    verified = all(d[3] for d in done)
    out({"status": "ok", "command": "mbom_path", "preview": preview, "count": done[0][2] + done[1][2],
         "output_files": [d[1] for d in done], "warnings": warnings,
         "agent_directive": ("VERIFIED: " if verified else "UNVERIFIED: 저장한 파일의 행 수를 확인하지 못함. ")
         + "next_commands를 처음부터 끝까지 순서대로 모두 실행하세요 (중간 명령이 주는 next_commands는 따르지 않음).",
         "next_commands": nxt})


if __name__ == "__main__":
    main()
