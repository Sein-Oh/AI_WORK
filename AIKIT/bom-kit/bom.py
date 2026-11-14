# -*- coding: utf-8 -*-
r"""
bom.py - bom-kit: BOM 업무 계산 (mbom-diff·where-used 스킬). 명령 하나로 끝난다 (Gemma 같은 경량 모델이 여러 단계를 잇지 않게).

    python "C:\AIKIT\bom-kit\bom.py" diff --old "<옛 판.xlsx>" --new "<새 판.xlsx>"
    python "C:\AIKIT\bom-kit\bom.py" where-used --file "<BOM.xlsx>" --find "<품번이나 품명>"

표 읽기·비교·내보내기는 pandas-kit 명령을 차례로 실행해서 한다 (C:\AIKIT\pandas-kit\pandas_kit.py, import하지 않고 실행만,
pandas-kit의 지금 작업은 바꾸지 않게 데이터 폴더를 따로 씀). 결과 csv·txt는 pandas-kit 작업 폴더(원본 옆)에 만든다. 원본 엑셀은 쓰지 않는다.
BOM 업무 규칙은 이 파일 한 곳에만 둔다:
- 상위 = 위쪽 행 중 가장 가까운, Level이 더 작은 행 (상위 품번 열이 없는 BOM).
- 경로 = 최상위부터 품번을 /로 이은 것. 같은 상위 아래 같은 품번이 둘 이상이면 모두 #순번 (#1, #2).
- 제품 1대당 수량 = 그 행과 위쪽 상위들의 수량을 곱한 값 (스크랩·유효 기간·대체 그룹은 보지 않음).
- 두 판 비교: 경로(품번 기준)로 짝을 맞춤. 다른 상위로 옮겨진 품목은 삭제·추가, 같은 자리에서 품번이 바뀐 것만 대체 후보
  (품명은 보지 않음, 삭제와 추가에 모두 있는 품번은 옮겨진 것이라 짝짓지 않음) (사용자 결정 2026-11-14).
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import shlex
import subprocess
import sys
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _bom_runtime as rt  # noqa: E402

ROW = "엑셀 행"                       # pandas-kit export의 원본 엑셀 행 번호 열
LEVEL_NAMES = ("level", "lvl", "레벨", "수준")
PART_NAMES = ("part number", "part no", "partnumber", "item", "품번", "자재번호", "자재 번호", "품목번호")
QTY_NAMES = ("qty", "quantity", "qpa", "수량", "소요량")
DESC_NAMES = ("description", "desc", "품명", "품목명", "내역", "규격")
UOM_NAMES = ("uom", "unit", "단위")
LINE_MAX = 30                         # JSON lines·preview에 넣는 줄 수 (넘으면 truncated + output_file)
PANDAS_KIT = os.path.join(os.path.dirname(rt.TOOL_DIR), "pandas-kit", "pandas_kit.py")   # C:\AIKIT\pandas-kit


# ==============================================================================
# 1. csv 읽기와 열 찾기
# ==============================================================================

def load_csv(text_path: str) -> Tuple[str, List[Dict[str, str]], List[str]]:
    """pandas-kit export csv -> (전체 경로, 행, 열). 인코딩은 자동 판별 (read_text_file)."""
    path = os.path.abspath(str(text_path).strip().strip('"'))
    if not os.path.isfile(path):
        raise FileNotFoundError(f"파일이 없습니다: {path}")
    rows = list(csv.DictReader(io.StringIO(rt.read_text_file(path))))
    if not rows:
        raise rt.InputError(f"{os.path.basename(path)}에 행이 없습니다.")
    cols = list(rows[0].keys())
    if ROW not in cols:
        raise rt.InputError(f"{os.path.basename(path)}에 '{ROW}' 열이 없습니다. pandas-kit export로 만든 csv를 주세요.")
    return path, rows, cols


def argv_setting(option: str, value: str) -> List[str]:
    """지금 실행 중인 bom.py 명령에서 option 값을 value로 (없으면 더함): [후보]·[추천]의 전체 명령."""
    args = [str(x) for x in (rt.CURRENT_ARGV or [""])[1:]]
    if option in args and args.index(option) + 1 < len(args):
        i = args.index(option)
        return args[:i + 1] + [value] + args[i + 2:]
    return args + [option, value]


def find_col(cols: List[str], names: Tuple[str, ...], given: Optional[str], what: str, option: str,
             needed: bool = True) -> Optional[str]:
    """열 이름 찾기. 못 찾으면 열들을 [후보]로, 고른 번호의 confirm_commands는 option을 넣은 bom.py 전체 명령."""
    if given and given in cols:
        return given
    if not given:
        for c in cols:
            if c.strip().lower() in names:
                return c
        if not needed:
            return None
    cands = [c for c in cols if c != ROW][:10]
    head = f"'{given}' 열이 없습니다." if given else f"{what} 열을 찾지 못했습니다."
    err = rt.ask(cands, [argv_setting(option, c) for c in cands], f"{what} 열")
    err.args = (f"{head} {err.args[0]}",) + err.args[1:]
    raise err


def level_of(v: Any) -> Optional[int]:
    try:
        f = float(str(v).strip())
    except ValueError:
        return None
    return int(f) if f == int(f) and f >= 0 else None


def qty_of(v: Any) -> Optional[float]:
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, AttributeError):
        return None


def number(v: float) -> str:
    return str(int(v)) if v == int(v) else f"{v:g}"


def rows_text(rows: List[Any]) -> str:
    return ", ".join(map(str, rows[:10])) + (" 등" if len(rows) > 10 else "")


# ==============================================================================
# 2. 구조: Level로 상위·경로·제품 1대당 수량 (BOM 업무 규칙, 이 한 곳에만)
# ==============================================================================

def tree(rows: List[Dict[str, str]], lcol: str, pcol: str, qcol: Optional[str] = None) -> Dict[str, Any]:
    """행마다 {r, level, parent, path, qty, total}. bad: Level·품번이 빈 행, jumps: Level을 건너뛴 행, noqty: 수량이 숫자가 아닌 행."""
    stack: List[Tuple[int, Tuple[str, ...], Optional[float]]] = []     # (level, 경로 조각, 누적 수량)
    count: Dict[Tuple[Tuple[str, ...], str], int] = {}
    raw, bad, jumps, noqty = [], [], [], []
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
        q = qty_of(r.get(qcol)) if qcol else None
        if qcol and q is None:
            noqty.append(r.get(ROW) or "?")
        total = None if q is None or above is None else above * q
        stack.append((lv, parts, total))
        raw.append((r, lv, parts, q, total))

    def shown(parts: Tuple[str, ...]) -> str:       # 같은 상위 아래 한 번만 나오는 품번은 #1을 뗌
        return "/".join(p if count[(parts[:i], p.rpartition("#")[0])] > 1 else p.rpartition("#")[0]
                        for i, p in enumerate(parts))
    items = [{"r": r, "level": lv, "parent": parts[-2].rpartition("#")[0] if len(parts) > 1 else "",
              "path": shown(parts), "qty": q, "total": total} for r, lv, parts, q, total in raw]
    return {"items": items, "bad": bad, "jumps": jumps, "noqty": noqty}


def tree_warnings(t: Dict[str, Any], who: str = "") -> List[str]:
    pre = f"{who}: " if who else ""
    out = []
    if t["bad"]:
        out.append(f"{pre}Level이 숫자가 아니거나 품번이 빈 행 {len(t['bad'])}개는 보지 않았습니다 (엑셀 행 {rows_text(t['bad'])}).")
    if t["jumps"]:
        out.append(f"{pre}Level이 두 단계 이상 건너뛴 행 {len(t['jumps'])}개 (엑셀 행 {rows_text(t['jumps'])}). "
                   "상위는 바로 위의 더 작은 Level로 봤습니다.")
    if t["noqty"]:
        out.append(f"{pre}수량이 숫자가 아닌 행 {len(t['noqty'])}개 (엑셀 행 {rows_text(t['noqty'])})는 제품 1대당 수량을 비웠습니다.")
    return out


def split(path: str) -> Tuple[str, str]:
    """'A/B#2/C' -> ('A/B#2', 'C')."""
    parent, _, last = path.rpartition("/")
    return parent, last.partition("#")[0]


def where(parent: str) -> str:
    return f"{parent} 아래" if parent else "최상위"


def write_csv(path: str, head: List[str], body: List[List[Any]], what: str) -> Tuple[str, bool]:
    """csv(UTF-8 BOM) 저장 -> (실제 경로, 다시 읽어 행 수가 같은지). 잠겨 있으면 새 이름 (save_with_lock_fallback)."""
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(head)
    w.writerows(body)
    text = buf.getvalue()

    def writer(p: str) -> None:
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            f.write(text)
    saved = rt.save_with_lock_fallback(path, writer, what)
    try:
        back = len(list(csv.reader(io.StringIO(rt.read_text_file(saved))))) - 1
    except OSError:
        back = -1
    return saved, back == len(body)


# ==============================================================================
# 3. pandas-kit 실행 (import하지 않고 전체 명령으로만, 데이터 폴더는 따로: 사용자의 지금 pandas-kit 작업을 바꾸지 않음)
# ==============================================================================

class PandasKit:
    """bom-kit 한 번 실행 동안의 pandas-kit 작업 하나. 명령마다 마지막 stdout 1줄 JSON을 읽고, 실패는 bom.py 오류로 바꾼다."""

    def __init__(self) -> None:
        if not os.path.isfile(PANDAS_KIT):
            raise rt.InputError(f"pandas-kit이 없습니다: {PANDAS_KIT} (C:\\AIKIT\\pandas-kit 폴더를 함께 두세요).")
        self.env = dict(os.environ, PANDAS_KIT_DATA_DIR=os.path.join(rt.get_data_dir(), "pandas-kit"),
                        PANDAS_KIT_SHELL=rt.command_shell(), PYTHONIOENCODING="utf-8")
        self.warnings: List[str] = []
        self.work: Optional[str] = None

    def run(self, args: List[str], who: str = "", shell: Optional[str] = None) -> Dict[str, Any]:
        env = dict(self.env, **({"PANDAS_KIT_SHELL": shell} if shell else {}))
        sys.stderr.write(f"[PROGRESS] pandas-kit {' '.join(args[:3])}\n")
        p = subprocess.run([sys.executable, PANDAS_KIT, *args], capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env)
        lines = [x for x in p.stdout.splitlines() if x.strip()]
        try:
            out = json.loads(lines[-1])
        except (IndexError, ValueError):
            raise RuntimeError(f"pandas-kit {args[0]}의 결과를 읽지 못했습니다 (exit {p.returncode}): {p.stderr[-300:]}") from None
        if out.get("status") != "ok":
            raise _pk_error(out, who)
        self.work = out.get("work") or self.work
        for w in (out.get("warnings") or []) + (out.get("recovered") or []):
            w = _plain_warning(w)
            if w:
                self.warnings.append(f"{who}: {w}" if who else w)
        return out


def _plain_warning(text: str) -> str:
    """pandas-kit 경고에서 pandas-kit 옵션·명령을 권하는 문장은 뺌 (경량 모델이 그 명령으로 여러 단계를 다시 잇지 않게).
    사실을 알리는 문장(숨긴 행, 글자로 저장된 숫자, 표 위의 내용)은 그대로."""
    keep = [t for t in re.split(r"(?<=[.)])\s+", str(text).strip()) if t and "python " not in t and "--" not in t]
    return " ".join(keep).strip()


# 표를 읽을 때 pandas-kit이 [후보]·[추천]으로 묻는 옵션 -> 같은 값을 넣을 bom.py 옵션 (옛 판·새 판 따로)
_READ_OPTIONS = ("--sheet", "--range")


def _pk_error(out: Dict[str, Any], who: str) -> Exception:
    """pandas-kit 실패 -> bom.py 오류. 시트·범위를 묻는 것이면 그 값을 넣은 bom.py 전체 명령으로 바꿔 줌."""
    msg = str(out.get("message") or "pandas-kit 실패")
    pre = f"{who}을(를) 읽다가 멈췄습니다: " if who else "pandas-kit이 멈췄습니다: "
    side = {"옛 판": "old", "새 판": "new", "BOM": ""}.get(who)
    cmds = [c.get("command", "") for c in out.get("confirm_commands") or []]
    if "[추천] " in msg:
        cmds = [msg.split("[추천] ", 1)[1]]
    fixed = []
    for c in cmds:
        try:
            parts = shlex.split(c)
        except ValueError:
            parts = []
        opt = next((o for o in _READ_OPTIONS if o in parts[:-1]), None)
        if side is None or not opt:
            fixed = []
            break
        # 옛 판에서 시트를 물으면 두 판 공통 --sheet를 고침 (새 판도 같은 시트일 때가 많음, 새 판에서 또 묻지 않게),
        # 새 판에서 물으면 새 판만 (--new-sheet), 범위는 판마다 다름 (--old-range, --new-range)
        name = opt if not side or (side == "old" and opt == "--sheet") else f"--{side}-{opt[2:]}"
        fixed.append(argv_setting(name, parts[parts.index(opt) + 1]))
    if fixed and "[추천] " in msg:
        return rt.InputError(pre + msg.split("[추천] ", 1)[0] + f"[추천] {rt.full_command(fixed[0])}")
    if fixed:
        labels = re.findall(r"\d+\) (.*?)(?= \d+\) | \(사용자가|$)", msg.split("[후보] ", 1)[1]) if "[후보] " in msg else []
        if len(labels) == len(fixed):
            err = rt.ask(labels, fixed, msg.split(" [후보]", 1)[0].removesuffix("을(를) 정해야 합니다."))
            err.args = (pre + err.args[0],) + err.args[1:]
            return err
    # 그 밖의 실패(파일 없음, 저장 안 된 문서 등)는 그대로 알림 (pandas-kit 명령은 주지 않음: 이어서 할 일은 bom.py 하나)
    text = re.sub(r" ?\[(추천|후보)\].*$", "", msg)
    return rt.InputError(pre + text)


def _table_name(path: str) -> str:
    """표 이름 = 파일 이름 (export가 같은 이름 파일이 있어 old_20261114_101500.csv로 저장했으면 그 꼬리는 뗌)."""
    return re.sub(r"_\d{8}_\d{6}$", "", os.path.splitext(os.path.basename(path))[0])


def read_bom(pk: PandasKit, file: str, name: str, who: str, sheet: Optional[str], rng: Optional[str]) -> Tuple[str, str]:
    """pandas-kit으로 BOM 읽기 -> export csv. (csv 경로, 읽은 범위 preview)."""
    args = ["select", "--file", file, "--name", name] + (["--sheet", sheet] if sheet else []) + (["--range", rng] if rng else [])
    got = pk.run(args, who, shell="bash")                      # [후보] 명령을 나눠 읽으려고 bash 따옴표 (_pk_error)
    out = pk.run(["export", "--from", name, "--save", f"{name}.csv"], who)
    return out["output_file"], got.get("preview", "")


# ==============================================================================
# 4. diff: 두 판 비교 (읽기 -> 경로 -> 경로로 diff -> 대체 후보·추가·삭제·바뀜 줄)
# ==============================================================================

def _paths(csv_path: str, level: Optional[str], part: Optional[str], who: str) -> Dict[str, Any]:
    path, rows, cols = load_csv(csv_path)
    t = tree(rows, find_col(cols, LEVEL_NAMES, level, "Level", "--level"), find_col(cols, PART_NAMES, part, "품번", "--part"))
    save, ok_ = write_csv(os.path.splitext(path)[0] + "_경로.csv", [ROW, "상위 품번", "경로"],
                          [[i["r"].get(ROW), i["parent"], i["path"]] for i in t["items"]], "경로 파일")
    multi = sorted({i["path"].rpartition("/")[2] for i in t["items"] if "#" in i["path"].rpartition("/")[2]})
    return {"table": _table_name(path), "save": save, "n": len(t["items"]), "ok": ok_, "multi": multi,
            "warnings": tree_warnings(t, who)}


def _num(v: Any) -> str:
    """2.0 -> 2 (pandas가 정수 열을 소수로 적은 것만), 나머지는 그대로."""
    v = str(v or "").strip()
    try:
        f = float(v)
    except ValueError:
        return v
    return str(int(f)) if f == int(f) and "e" not in v.lower() else v


def _siblings(path: str) -> Dict[str, Tuple[str, str]]:
    """<표>_경로.csv -> {경로: (앞 형제 품번, 뒤 형제 품번)} (같은 상위 아래 표 순서)."""
    kids: Dict[str, List[str]] = {}
    for r in csv.DictReader(io.StringIO(rt.read_text_file(path))):
        kids.setdefault(split(r["경로"])[0], []).append(r["경로"])
    spot = {}
    for ks in kids.values():
        for i, k in enumerate(ks):
            spot[k] = (split(ks[i - 1])[1] if i else "(처음)", split(ks[i + 1])[1] if i + 1 < len(ks) else "(끝)")
    return spot


def summarize(rows: List[Dict[str, str]], ocol: str, ncol: str, order: Optional[List[Dict[str, Tuple[str, str]]]] = None):
    dels = [(r["경로"], r.get(ocol) or "") for r in rows if r["구분"] == "삭제"]
    adds = [(r["경로"], r.get(ncol) or "") for r in rows if r["구분"] == "추가"]
    lines, counts = [], {"대체 후보": 0, "추가": 0, "삭제": 0, "바뀐 칸": 0}
    # 대체 후보: 같은 상위 아래, 앞이나 뒤의 형제 품번이 같은 자리 (형제 순서를 모르면 삭제 하나·추가 하나일 때만).
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
        add = same_a[0]
        lines.append(f"[대체 후보] {split(d[0])[1]} -> {split(add[0])[1]}: {where(parent)} 같은 자리 "
                     f"(옛 엑셀 행 {d[1]}, 새 엑셀 행 {add[1]})")
        dels.remove(d)
        adds.remove(add)
        counts["대체 후보"] += 1
    for path, row in adds:
        parent, pn = split(path)
        lines.append(f"[추가] {pn}: {where(parent)} (새 엑셀 행 {row})")
    for path, row in dels:
        parent, pn = split(path)
        lines.append(f"[삭제] {pn}: {where(parent)} (옛 엑셀 행 {row})")
    counts["추가"], counts["삭제"] = len(adds), len(dels)
    changed: Dict[str, List[Dict[str, str]]] = {}
    for r in rows:
        if r["구분"] == "바뀜":
            changed.setdefault(r["경로"], []).append(r)
    for path, rs in changed.items():
        parent, pn = split(path)
        cells = ", ".join(f"{r['열']} {_num(r['옛 값']) or '(빈 칸)'} -> {_num(r['새 값']) or '(빈 칸)'}" for r in rs)
        lines.append(f"[바뀜] {pn} ({where(parent)}): {cells} (옛 엑셀 행 {rs[0].get(ocol)}, 새 엑셀 행 {rs[0].get(ncol)})")
        counts["바뀐 칸"] += len(rs)
    return lines, counts, len(changed)


def cmd_diff(a: Any) -> Dict[str, Any]:
    pk = PandasKit()
    pk.run(["new"])
    o_csv, o_read = read_bom(pk, a.old, "old", "옛 판", a.old_sheet or a.sheet, a.old_range)
    n_csv, n_read = read_bom(pk, a.new, "new", "새 판", a.new_sheet or a.sheet, a.new_range)
    o, n = _paths(o_csv, a.level, a.part, "옛 판"), _paths(n_csv, a.level, a.part, "새 판")
    for d in (o, n):
        pk.run(["select", "--file", d["save"], "--name", f"{d['table']}k"])
        pk.run(["join", "--left", d["table"], "--right", f"{d['table']}k", "--on", "엑셀 행=엑셀 행번호", "--name", f"{d['table']}p"])
    pk.warnings = [w for w in pk.warnings if "엑셀 행번호" not in w]      # 경로 표의 '엑셀 행' 이름 바꿈 안내는 bom-kit 안의 일
    diff = pk.run(["diff", "--old", "oldp", "--new", "newp", "--on", "경로", "--name", "chg"])
    changes = pk.run(["export", "--from", "chg", "--save", "변경내용.csv"])["output_file"]
    rows = list(csv.DictReader(io.StringIO(rt.read_text_file(changes))))
    order = [_siblings(o["save"]), _siblings(n["save"])]
    lines, counts, nrows = summarize(rows, "oldp 행", "newp 행", order) if rows else ([], {}, 0)
    head = ("두 판이 같습니다 (찾은 변경 없음)." if not lines else
            "변경 " + ", ".join(f"{k} {v}" + ("곳" if k == "바뀐 칸" else "건") for k, v in counts.items() if v)
            + (f" ({nrows}개 품목)" if counts.get("바뀐 칸") else "") + ".")
    text = "\n".join([head] + lines) + "\n"

    def writer(p: str) -> None:
        with open(p, "w", encoding="utf-8-sig") as f:
            f.write(text)
    save = rt.save_with_lock_fallback(os.path.splitext(changes)[0] + "_요약.txt", writer, "요약 파일")
    verified = len(rt.read_text_file(save).splitlines()) == len(lines) + 1 and o["ok"] and n["ok"]
    shown = lines[:LINE_MAX]
    read = f"읽은 표: 옛 판 {o_read.split(' = ', 1)[-1]}, 새 판 {n_read.split(' = ', 1)[-1]}."
    preview = "\n".join([head] + shown + ([f"(나머지 {len(lines) - len(shown)}줄은 {save})"] if len(lines) > len(shown) else [])
                        + [read])
    then = ("preview를 그대로 사용자에게 알리세요 (줄을 직접 세거나 고쳐 쓰지 마세요). 같은 명령을 다시 실행하지 마세요."
            + (" [대체 후보]는 도구의 짐작이니 맞는지 사용자에게 확인하세요." if counts.get("대체 후보") else "")
            + ("" if lines else " 0건. 찾은 변경이 없습니다."))
    out = rt.ok("diff", preview, verified, then, None if verified else ["저장한 요약·경로 파일을 확인하지 못함"],
                count=len(lines), counts=counts, lines=shown, truncated=len(lines) > len(shown), output_file=save,
                changes_file=changes, work=pk.work)
    paint = diff.get("confirm_commands") or []
    if paint:                                    # 원본에 바뀐 칸 칠하기 (office-cli, 사용자가 원할 때만)
        out["confirm_commands"] = paint
        out["agent_directive"] += " confirm_commands는 원본 엑셀에 바뀐 칸을 칠하는 명령입니다. 사용자가 칠해 달라고 할 때만 1번 실행하세요."
    warnings = pk.warnings + o["warnings"] + n["warnings"]
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


# ==============================================================================
# 5. where-used: 품목이 쓰인 상위와 제품 1대당 수량
# ==============================================================================

def cmd_where_used(a: Any) -> Dict[str, Any]:
    pk = PandasKit()
    pk.run(["new"])
    bom_csv, read = read_bom(pk, a.file, "bom", "BOM", a.sheet, a.range)
    path, rows, cols = load_csv(bom_csv)
    pcol = find_col(cols, PART_NAMES, a.part, "품번", "--part")
    qcol = find_col(cols, QTY_NAMES, a.qty, "수량", "--qty")
    dcol = find_col(cols, DESC_NAMES, None, "품명", "", needed=False)
    ucol = find_col(cols, UOM_NAMES, None, "단위", "", needed=False)
    t = tree(rows, find_col(cols, LEVEL_NAMES, a.level, "Level", "--level"), pcol, qcol)
    items = t["items"]
    key = a.find.strip().lower()
    hits = [i for i in items if str(i["r"].get(pcol) or "").strip().lower() == key
            or (dcol and key in str(i["r"].get(dcol) or "").lower())]
    name: Dict[str, str] = {}
    for i in hits:
        name.setdefault(str(i["r"].get(pcol) or "").strip(), str(i["r"].get(dcol) or "").strip() if dcol else "")
    save, ok_ = write_csv(os.path.splitext(path)[0] + "_쓰인곳.csv",
                          [ROW, "품번", "품명", "상위 품번", "경로", "그 자리 수량", "제품 1대당 수량", "단위"],
                          [[i["r"].get(ROW), i["r"].get(pcol), i["r"].get(dcol) if dcol else "", i["parent"], i["path"],
                            "" if i["qty"] is None else number(i["qty"]), "" if i["total"] is None else number(i["total"]),
                            i["r"].get(ucol) if ucol else ""] for i in hits], "쓰인 곳 파일")
    warnings = pk.warnings + tree_warnings(t)
    unverified = None if ok_ else ["저장한 쓰인 곳 파일의 행 수를 확인하지 못함"]
    if not hits:
        out = rt.ok("where-used", f"'{a.find}'을(를) 품번이나 품명에서 찾지 못했습니다 (BOM {len(items)}행, {read.split(' = ', 1)[-1]}).", ok_,
                    "0건. 찾은 품목이 없습니다. preview를 그대로 알리고 다시 실행하지 마세요. 품번이나 품명의 일부를 --find에 주세요.",
                    unverified, count=0, lines=[], output_file=save, work=pk.work)
    else:
        parent_name: Dict[str, str] = {}
        if dcol:
            for i in items:                                  # 상위 품번의 품명 (처음 나온 것)
                parent_name.setdefault(str(i["r"].get(pcol) or "").strip(), str(i["r"].get(dcol) or "").strip())
        lines, totals = [], {}
        for i in hits:
            pn = str(i["r"].get(pcol) or "").strip()
            unit = str(i["r"].get(ucol) or "").strip() if ucol else ""
            pname = f" {parent_name[i['parent']]}" if parent_name.get(i["parent"]) else ""
            place = f"{i['parent']}{pname} 아래" if i["parent"] else "최상위"
            amount = "수량 모름" if i["total"] is None else f"{number(i['total'])}{(' ' + unit) if unit else ''}"
            lines.append(f"[쓰임] {pn}: {place} {amount} (경로 {i['path']}, 엑셀 행 {i['r'].get(ROW)})")
            if i["total"] is not None:
                totals.setdefault((pn, unit), [0.0, 0])
                totals[(pn, unit)][0] += i["total"]
                totals[(pn, unit)][1] += 1
        sums = [f"{pn} {name.get(pn) or ''}".strip() + f": {len(hits) if len(totals) == 1 else c}곳, "
                f"제품 1대당 {number(tot)}{(' ' + u) if u else ''}" for (pn, u), (tot, c) in totals.items()]
        shown = lines[:LINE_MAX]
        preview = f"'{a.find}': {len(hits)}곳에 쓰입니다. " + " / ".join(sums) + ("\n" + "\n".join(shown) if shown else "")
        out = rt.ok("where-used", preview, ok_,
                    "preview의 줄을 사용자에게 그대로 알리세요 (줄을 직접 세거나 고쳐 쓰지 마세요). 같은 명령을 다시 실행하지 마세요."
                    + (" 잘렸으면 전체는 output_file에 있습니다." if len(lines) > LINE_MAX else ""),
                    unverified, count=len(hits), lines=shown, totals=sums, truncated=len(lines) > LINE_MAX, output_file=save,
                    work=pk.work)
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


# ==============================================================================
# 6. Parser
# ==============================================================================

def build_cli() -> rt.JsonArgumentParser:
    root, sub = rt.make_parser("bom-kit: BOM 업무 계산 (표 읽기는 pandas-kit, 원본 엑셀은 그대로)")
    p = rt.add_command(sub, root, "diff", "MBOM 두 판 비교: 대체 후보·추가·삭제·바뀐 칸 (명령 하나로)")
    p.add_argument("--old", required=True, help="옛 판 파일 (.xlsx, .xlsm, .csv)")
    p.add_argument("--new", required=True, help="새 판 파일")
    p.add_argument("--sheet", help="두 판의 시트 이름 (생략하면 파일을 열었을 때 보이는 시트)")
    p.add_argument("--old-sheet", help="옛 판만 다른 시트일 때 (도구가 [후보]로 줌)")
    p.add_argument("--new-sheet", help="새 판만 다른 시트일 때 (도구가 [후보]로 줌)")
    p.add_argument("--old-range", help="옛 판 표 범위 (도구가 [후보]로 줌, 생략하면 표를 찾음)")
    p.add_argument("--new-range", help="새 판 표 범위 (도구가 [후보]로 줌)")
    p.add_argument("--level", help="Level 열 이름 (생략하면 찾음)")
    p.add_argument("--part", help="품번 열 이름 (생략하면 찾음)")
    p = rt.add_command(sub, root, "where-used", "품목이 쓰인 상위와 제품 1대당 수량 (명령 하나로)")
    p.add_argument("--file", required=True, help="BOM 파일 (.xlsx, .xlsm, .csv)")
    p.add_argument("--find", required=True, help="찾을 품번(그대로 같은 것) 또는 품명에 든 글자")
    p.add_argument("--sheet", help="시트 이름 (생략하면 파일을 열었을 때 보이는 시트)")
    p.add_argument("--range", help="표 범위 (도구가 [후보]로 줌, 생략하면 표를 찾음)")
    p.add_argument("--level", help="Level 열 이름 (생략하면 찾음)")
    p.add_argument("--part", help="품번 열 이름 (생략하면 찾음)")
    p.add_argument("--qty", help="수량 열 이름 (생략하면 찾음)")
    return root


COMMANDS = {"diff": cmd_diff, "where-used": cmd_where_used}


def main(argv: Optional[List[str]] = None) -> None:
    rt.setup_console()
    rt.CURRENT_ARGV = [os.path.abspath(__file__)] + list(sys.argv[1:] if argv is None else argv)
    args = build_cli().parse_args(rt.CURRENT_ARGV[1:])
    rt.run_cli(lambda: COMMANDS[args.cmd](args), hint="확인할 것: 파일 경로, 시트, 열 이름.")


if __name__ == "__main__":
    main()
