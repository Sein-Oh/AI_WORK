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
import math
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _bom_runtime as rt  # noqa: E402

ROW = "엑셀 행"                       # pandas-kit export의 원본 엑셀 행 번호 열
LEVEL_NAMES = ("level", "lvl", "레벨", "수준")
PART_NAMES = ("part number", "part no", "partnumber", "item", "품번", "자재번호", "자재 번호", "품목번호")
QTY_NAMES = ("qty", "quantity", "qpa", "수량", "소요량")
DESC_NAMES = ("description", "desc", "품명", "품목명", "내역", "규격")
UOM_NAMES = ("uom", "unit", "단위")
KEY_SEP, KEY_NUM = " >> ", " ##"     # 짝 맞추는 경로 키: 품번에 '/'·'#'이 있어도 겹치지 않게 (화면용 경로는 '/'와 '#순번')
LINE_MAX = 30                         # JSON lines·preview에 넣는 줄 수 (넘으면 truncated + output_file)
MARKER = ".created_by_pandas-kit"      # pandas-kit이 만든 작업 폴더의 표시 파일 (이것이 있는 폴더만 지움)
RESULT_MARKER = ".created_by_bom-kit"  # 결과 폴더(<원본 이름>_bom)의 표시 파일
WORKS_MAX = 200                       # works.json에 기억하는 못 지운 폴더 수
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
    """Level 칸 -> 깊이. 숫자(0, 1, 2.0), 점 표기(.1, ..2 -> 뒤의 숫자, 숫자 없이 점만 '..' -> 점 수).
    'inf', 'nan' 글자와 음수·소수는 Level이 아님 (Fable 검토 2026-11-14: 점 표기를 모두 버리고 '두 판이 같습니다'였음)."""
    t = str(v if v is not None else "").strip()
    dots = re.fullmatch(r"(\.+)(\d+)?", t)
    if dots:
        return int(dots.group(2)) if dots.group(2) else len(dots.group(1))
    try:
        f = float(t)
    except ValueError:
        return None
    return int(f) if math.isfinite(f) and f == int(f) and f >= 0 else None      # "inf", "nan" 글자는 Level이 아님


def qty_of(v: Any) -> Optional[float]:
    try:
        f = float(str(v).replace(",", "").strip())
    except (ValueError, AttributeError):
        return None
    return f if math.isfinite(f) else None


def _int(v: Any) -> Optional[int]:
    """csv의 엑셀 행 번호 ('9', '9.0') -> 9."""
    try:
        return int(float(str(v).strip()))
    except (TypeError, ValueError):
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
        parts = parent + (f"{pn}{KEY_NUM}{n}",)
        q = qty_of(r.get(qcol)) if qcol else None
        if qcol and q is None:
            noqty.append(r.get(ROW) or "?")
        total = None if q is None or above is None else above * q
        stack.append((lv, parts, total))
        raw.append((r, lv, parts, q, total))

    def shown(parts: Tuple[str, ...]) -> str:       # 화면용: 같은 상위 아래 한 번만 나오는 품번은 #1을 뗌
        out = []
        for i, p in enumerate(parts):
            pn, _, n = p.rpartition(KEY_NUM)
            out.append(f"{pn}#{n}" if count[(parts[:i], pn)] > 1 else pn)
        return "/".join(out)
    # 짝 맞추는 키는 늘 #순번 (Fable 검토 2026-11-14: 같은 상위 아래 둘이던 품번이 하나가 되면 #1을 떼어 짝이 어긋났음)
    items = [{"r": r, "level": lv, "parent": parts[-2].rpartition(KEY_NUM)[0] if len(parts) > 1 else "",
              "path": shown(parts), "key": KEY_SEP.join(parts), "qty": q, "total": total} for r, lv, parts, q, total in raw]
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


def split(key: str) -> Tuple[str, str]:
    """경로 키 -> (상위의 경로 키, 품번)."""
    parent, _, last = key.rpartition(KEY_SEP)
    return parent, last.rpartition(KEY_NUM)[0]


SHOWN: Dict[str, str] = {}                 # 경로 키 -> 화면용 경로 (두 판의 _paths가 채움)


def where(parent: str) -> str:
    return f"{SHOWN.get(parent, parent)} 아래" if parent else "최상위"


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
        # 실행마다 pandas-kit 데이터 폴더를 따로 (Fable 검토 2026-11-14 A2: 두 실행이 지금 작업 포인터 하나를 같이 써
        # 둘 다 멈췄음). 사용자의 pandas-kit 작업도, 다른 bom-kit 실행도 건드리지 않는다. 오류 로그는 끝에 공용 로그로 옮김.
        self.base = os.path.join(rt.get_data_dir(), "pandas-kit")
        _drop_old_runs(os.path.join(self.base, "runs"))
        self.data = os.path.join(self.base, "runs", f"{os.getpid()}_{int(time.time() * 1000)}")
        os.makedirs(self.data, exist_ok=True)
        self.env = dict(os.environ, PANDAS_KIT_DATA_DIR=self.data, PANDAS_KIT_SHELL=rt.command_shell(), PYTHONIOENCODING="utf-8")
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

    # --- 작업 폴더(캐시) 정리: 결과는 결과 폴더에 두고 pandas-kit 작업 폴더는 실행마다 지움 (0.1.2) ----------------
    def current(self) -> Optional[str]:
        """이 실행의 pandas-kit이 만든 작업 폴더 (멈춘 실행이 만든 폴더도)."""
        try:
            with open(os.path.join(self.data, "current.json"), encoding="utf-8") as f:
                return json.load(f).get("work") or None
        except (OSError, ValueError):
            return self.work

    def close(self) -> None:
        """작업 폴더와 이 실행의 데이터 폴더 지우기. pandas-kit 오류 로그는 공용 로그(<bom-kit 데이터>/pandas-kit/logs)에 덧붙임."""
        cur = self.current()
        if cur and not self.drop(cur, "작업 폴더(캐시)"):
            _remember_pending(cur)
        log = os.path.join(self.data, "logs", "errors.txt")
        try:
            if os.path.isfile(log):
                os.makedirs(os.path.join(self.base, "logs"), exist_ok=True)
                with open(log, encoding="utf-8", errors="replace") as f,                         open(os.path.join(self.base, "logs", "errors.txt"), "a", encoding="utf-8") as g:
                    g.write(f.read())
        except OSError:
            pass
        shutil.rmtree(self.data, ignore_errors=True)

    @staticmethod
    def drop(folder: Optional[str], why: str, marked: bool = True) -> bool:
        """bom-kit이 만든 작업 폴더 지우기 -> 다 지웠거나 없으면 True, 못 지워 그대로 두면 False.
        marked: pandas-kit 표시 파일이 있는 폴더만 (표시 파일을 보고 기억해 둔 폴더를 다시 지울 때는 False).
        안의 파일이 열려 있으면 반쯤 지우지 않게 먼저 폴더 이름을 바꿔 본다 (Windows: 열린 파일이 든 폴더는 이름을 못 바꿈,
        집 PC 2026-11-14 B9: rmtree가 표시 파일부터 지우고 열린 csv에서 멈춰 다시는 못 지우는 폴더가 남았음)."""
        if not folder or not os.path.isdir(folder):
            return True
        if marked and not os.path.isfile(os.path.join(folder, MARKER)):
            return True                                  # 우리가 만든 폴더가 아님: 손대지 않고 잊음
        gone = folder.rstrip("\\/") + ".지우는중"
        n = 2
        while os.path.exists(gone):                      # 앞에서 다 못 지운 같은 이름이 있으면 번호
            gone, n = folder.rstrip("\\/") + f".지우는중{n}", n + 1
        try:
            rt.retry_fs(lambda: os.rename(folder, gone), attempts=3)
        except OSError as e:
            sys.stderr.write(f"[WARN] {why}: 작업 폴더 안의 파일이 열려 있어 그대로 둡니다. 다음 실행 때 다시 지웁니다 "
                             f"({folder}, {type(e).__name__})\n")
            return False
        try:
            rt.retry_fs(lambda: shutil.rmtree(gone), attempts=3)
        except OSError as e:                             # 이름은 바뀌었는데 못 지움 (드묾): 바뀐 이름을 기억해 다음에 다시
            sys.stderr.write(f"[WARN] {why}: 작업 폴더를 다 지우지 못했습니다. 다음 실행 때 다시 지웁니다 ({gone}, {type(e).__name__})\n")
            _remember_pending(gone)
            return True
        sys.stderr.write(f"[INFO] {why}: 작업 폴더를 지웠습니다 ({folder})\n")
        return True


def _drop_old_runs(runs: str) -> None:
    """끝나지 못한(중단된) 실행의 데이터 폴더 중 하루 지난 것은 지움."""
    try:
        for n in os.listdir(runs):
            p = os.path.join(runs, n)
            if time.time() - os.path.getmtime(p) > 86400:
                shutil.rmtree(p, ignore_errors=True)
    except OSError:
        pass


def _works_path() -> str:
    return os.path.join(rt.get_data_dir(), "works.json")


def _works() -> Dict[str, Any]:
    try:
        with open(_works_path(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


PENDING = "_못 지운 폴더"                # works.json에서 다음 실행 때 다시 지울 폴더 목록의 키


def _save_works(data: Dict[str, Any]) -> None:
    try:
        os.makedirs(rt.get_data_dir(), exist_ok=True)
        text = json.dumps(data, ensure_ascii=False)
        tmp = _works_path() + f".{os.getpid()}.tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(text)
        rt.retry_fs(lambda: os.replace(tmp, _works_path()))
    except OSError:
        pass                                   # 기억하지 못해도 결과는 그대로 (다음에 폴더 하나가 더 남을 뿐)


def _remember_pending(folder: str) -> None:
    data = _works()
    pend = [p for p in data.get(PENDING) or [] if isinstance(p, str)]
    if folder not in pend:
        pend.append(folder)
    data[PENDING] = pend[-WORKS_MAX:]
    _save_works(data)


def _retry_pending() -> None:
    """앞 실행에서 열려 있어 못 지운 폴더를 다시 지움 (아직 열려 있으면 그대로 기억)."""
    data = _works()
    pend = [p for p in data.get(PENDING) or [] if isinstance(p, str)]
    if not pend:
        return
    left = [p for p in pend if not PandasKit.drop(p, "앞 실행에서 못 지운 폴더", marked=False)]
    data = _works()                            # drop이 바꾼 이름(.지우는중)을 더했을 수 있음
    data[PENDING] = left + [p for p in data.get(PENDING) or [] if p not in pend and p not in left]
    if not data[PENDING]:
        del data[PENDING]
    _save_works(data)


def result_dir(original: str) -> str:
    """결과 폴더: 원본 옆 <원본 이름>_bom (사용자 결정 2026-11-14). 같은 이름의 다른 폴더(표시 파일 없음)가 있으면 _bom2, _bom3."""
    stem = os.path.splitext(os.path.abspath(original))[0]
    n = 1
    while True:
        folder = stem + ("_bom" if n == 1 else f"_bom{n}")
        if not os.path.exists(folder):
            rt.retry_fs(lambda: os.makedirs(folder, exist_ok=True))
            with open(os.path.join(folder, RESULT_MARKER), "w", encoding="utf-8") as f:
                f.write("bom-kit 결과 폴더 (실행마다 시각을 붙인 새 파일, 덮어쓰지 않음)\n")
            return folder
        if os.path.isfile(os.path.join(folder, RESULT_MARKER)):
            return folder
        n += 1


def beside(target: str, what: str) -> str:
    """사용자가 볼 결과 파일: 대상 파일과 같은 폴더의 <대상 이름>_<what>_<시각> (전역 지침 2장 '결과 파일 위치', 사용자 지적 2026-11-14).
    _bom 결과 폴더에는 도구가 다시 쓰는 것(비교 기록, 칠할 칸 목록, 변경내용 csv)만 둔다."""
    t = os.path.abspath(target)
    return stamped(os.path.dirname(t), f"{file_part(os.path.splitext(os.path.basename(t))[0], 60)}_{what}")


def stamped(folder: str, base: str) -> str:
    """결과 파일 이름 앞부분: <base>_<YYYYMMDD_HHMMSS> (사용자 결정 2026-11-14: 결과 파일은 덮어쓰지 않고 실행마다 새 이름, R4).
    같은 초에 같은 이름이 있으면 _2, _3."""
    head = f"{base}_{datetime.now():%Y%m%d_%H%M%S}"
    name, n = head, 2
    while any(f.startswith(name + "_") or f.startswith(name + ".") for f in os.listdir(folder)):
        name, n = f"{head}_{n}", n + 1
    return os.path.join(folder, name)


def file_part(text: str, limit: int = 40) -> str:
    """파일 이름에 넣을 말 (찾는 말, 새 판 이름): 파일 이름에 못 쓰는 글자는 _."""
    t = re.sub(r'[\\/:*?"<>|\s]+', "_", str(text).strip()).strip("._")
    return (t or "_")[:limit]


def run_once(body: Any) -> Dict[str, Any]:
    """작업 하나 (0.1.2, 사용자 결정 2026-11-14). 결과 파일은 원본 옆 결과 폴더(<원본 이름>_bom)에 실행마다 시각을 붙인 새 이름으로
    둔다 (덮어쓰지 않음: 알려 준 output_file이 사라지지 않고 diff 결과가 쌓임),
    pandas-kit 작업 폴더는 캐시라 실행 끝(멈춘 실행 포함)에 지운다. 안의 파일이 열려 있어 못 지운 폴더는 기억했다가 다음 실행 때 다시."""
    _retry_pending()
    pk = PandasKit()
    try:
        return body(pk)
    finally:
        pk.close()


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
    fixed: List[List[str]] = []
    picks: List[Tuple[str, str]] = []
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
        picks.append((name, parts[parts.index(opt) + 1]))
    if fixed and "[추천] " in msg:
        return rt.InputError(pre + msg.split("[추천] ", 1)[0] + f"[추천] {rt.full_command(fixed[0])}")
    if fixed:
        labels = re.findall(r"\d+\) (.*?)(?= \d+\) | \(사용자가|$)", msg.split("[후보] ", 1)[1]) if "[후보] " in msg else []
        if len(labels) == len(fixed):
            err = rt.ask(labels, fixed, msg.split(" [후보]", 1)[0].removesuffix("을(를) 정해야 합니다."))
            err.args = (pre + err.args[0],) + err.args[1:]
            err.choices = list(zip(labels, picks))          # _diff가 두 판의 범위를 한 번에 물을 때 씀
            return err
    # 그 밖의 실패(파일 없음, 저장 안 된 문서 등)는 그대로 알림 (pandas-kit 명령은 주지 않음: 이어서 할 일은 bom.py 하나)
    text = re.sub(r" ?\[(추천|후보)\].*$", "", msg)
    keep = [t for t in re.split(r"(?<=[.)])\s+", text.strip()) if t and "python " not in t and "pandas_kit" not in t]
    return rt.InputError(pre + (" ".join(keep) or text))     # pandas-kit 명령을 권하는 문장은 뺌 (Fable 검토 D2)


def _table_name(path: str) -> str:
    """표 이름 = 파일 이름 (export가 같은 이름 파일이 있어 old_20261114_101500.csv로 저장했으면 그 꼬리는 뗌)."""
    return re.sub(r"_\d{8}_\d{6}$", "", os.path.splitext(os.path.basename(path))[0])


def read_bom(pk: PandasKit, file: str, name: str, who: str, sheet: Optional[str], rng: Optional[str]) -> Tuple[str, str, List[int]]:
    """pandas-kit으로 BOM 읽기 -> export csv. (csv 경로, 읽은 범위 preview, 읽은 숨긴 행의 엑셀 행)."""
    args = ["select", "--file", file, "--name", name] + (["--sheet", sheet] if sheet else []) + (["--range", rng] if rng else [])
    try:
        got = pk.run(args, who, shell="bash")                  # [후보] 명령을 나눠 읽으려고 bash 따옴표 (_pk_error)
    except rt.InputError as e:
        # 표 안에 빈 행이 있어 범위를 물으면 묻지 않고 빈 행까지 한 표로 읽고 알림 (사용자 결정 2026-11-14:
        # BOM 가운데 빈 행은 표가 끝난 것이 아님. 전에는 옛 판·새 판마다 [후보]로 물었음)
        whole = next((c for c in getattr(e, "choices", None) or [] if c[0].startswith(WHOLE)), None)
        if rng or not whole:
            raise
        got = pk.run(args + ["--range", whole[1][1]], who, shell="bash")
        pk.warnings = [w for w in pk.warnings if not (w.startswith(f"{who}: ") and "범위 안의 빈 행" in w)]
        pk.warnings.append(f"{who}: 표 안에 빈 행이 있어 빈 행까지 한 표로 읽었습니다 ({whole[0][len(WHOLE):].strip(' ()')}).")
    out = pk.run(["export", "--from", name, "--save", f"{name}.csv"], who)
    return out["output_file"], got.get("preview", ""), [int(r) for r in got.get("hidden_rows") or []]


# ==============================================================================
# 4. diff: 두 판 비교 (읽기 -> 경로 -> 경로로 diff -> 대체 후보·추가·삭제·바뀜 줄)
# ==============================================================================

def _paths(csv_path: str, level: Optional[str], part: Optional[str], who: str) -> Dict[str, Any]:
    path, rows, cols = load_csv(csv_path)
    t = tree(rows, find_col(cols, LEVEL_NAMES, level, "Level", "--level"), find_col(cols, PART_NAMES, part, "품번", "--part"))
    save, ok_ = write_csv(os.path.splitext(path)[0] + "_경로.csv", [ROW, "상위 품번", "경로"],
                          [[i["r"].get(ROW), i["parent"], i["key"]] for i in t["items"]], "경로 파일")
    SHOWN.update({i["key"]: i["path"] for i in t["items"]})
    return {"table": _table_name(path), "save": save, "n": len(t["items"]), "ok": ok_, "bad": len(t["bad"]),
            "rows": len(rows), "cols": cols, "data": rows, "warnings": tree_warnings(t, who)}


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


def _pns(key: str) -> List[str]:
    return [p.rpartition(KEY_NUM)[0] for p in key.split(KEY_SEP)]


def _row_no(v: Any) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return float("inf")


def _removed_parents(dels: List[Tuple[str, str]], adds: List[Tuple[str, str]]):
    """상위 행이 지워져 그 하위가 새 판에서 다른 상위 아래로 붙은 것 (Level만 보고 상위를 정하므로 바로 위 형제에 붙음).
    지운 상위마다 하위를 행 순서대로 같은 상대 경로(품번)의 추가 행과 짝지음 (집 PC 2026-11-14 sample1·2: 상위 2개를 지웠는데
    하위 수십 행이 '삭제 + 추가'로 나와 추가 38, 삭제 39였음). -> [(상위 키, 옛 행, 새 상위 키, [(옛 키, 옛 행, 새 키, 새 행)])]"""
    used_d, used_a, out = set(), set(), []
    adds_sorted = sorted(adds, key=lambda x: _row_no(x[1]))
    for pk_, prow in sorted(dels, key=lambda x: (len(x[0].split(KEY_SEP)), _row_no(x[1]))):
        if pk_ in used_d:
            continue
        plen = len(pk_.split(KEY_SEP))
        desc = sorted([(k, r) for k, r in dels if k.startswith(pk_ + KEY_SEP) and k not in used_d], key=lambda x: _row_no(x[1]))
        if not desc:
            continue
        best: Dict[str, List[Tuple[str, str, str, str]]] = {}
        for x in {KEY_SEP.join(ak.split(KEY_SEP)[:len(_pns(ak)) - len(_pns(dk)[plen:])])
                  for dk, _ in desc for ak, _ in adds_sorted
                  if ak not in used_a and len(_pns(ak)) > len(_pns(dk)[plen:]) and _pns(ak)[-len(_pns(dk)[plen:]):] == _pns(dk)[plen:]}:
            pairs, taken, last = [], set(), -1.0
            for dk, dr in desc:                          # 행 순서를 지키며 같은 상대 경로의 추가 행과 짝
                rel = _pns(dk)[plen:]
                for ak, ar in adds_sorted:
                    if ak in used_a or ak in taken or _row_no(ar) <= last or not ak.startswith(x + KEY_SEP):
                        continue
                    if _pns(ak)[len(x.split(KEY_SEP)):] == rel:
                        pairs.append((dk, dr, ak, ar))
                        taken.add(ak)
                        last = _row_no(ar)
                        break
            if len(pairs) > len(best.get("p", [])):
                best = {"x": [x], "p": pairs}
        pairs = best.get("p", [])
        if not pairs or len(pairs) * 2 < len(desc):     # 하위의 절반도 못 찾으면 묶지 않음 (짐작하지 않음)
            continue
        used_d.add(pk_)
        used_d.update(p[0] for p in pairs)
        used_a.update(p[2] for p in pairs)
        out.append((pk_, prow, best["x"][0], pairs))
    return out


def summarize(rows: List[Dict[str, str]], ocol: str, ncol: str, order: Optional[List[Dict[str, Tuple[str, str]]]] = None):
    dels = [(r["경로"], r.get(ocol) or "") for r in rows if r["구분"] == "삭제"]
    adds = [(r["경로"], r.get(ncol) or "") for r in rows if r["구분"] == "추가"]
    lines, counts = [], {"상위 삭제": 0, "대체 후보": 0, "추가": 0, "삭제": 0, "바뀐 칸": 0}
    moved_rows: Dict[str, set] = {"old": set(), "new": set()}
    items: List[Dict[str, Any]] = []

    def item(kind: str, pn: str, parent: str, col: str = "", ov: str = "", nv: str = "", orow: Any = "", nrow: Any = "",
             note: str = "") -> None:
        items.append({"구분": kind, "품번": pn, "위치": SHOWN.get(parent, parent) or "(최상위)", "열": col, "옛 값": ov, "새 값": nv,
                      "옛 엑셀 행": _int(orow) or "", "새 엑셀 행": _int(nrow) or "", "비고": note})
    summarize.items = items
    for pk_, prow, new_parent, pairs in _removed_parents(dels, adds):
        parent, pn = split(pk_)
        lines.append(f"[상위 삭제] {pn}: {where(parent)} (옛 엑셀 행 {prow}). 그 하위 {len(pairs)}행은 새 판에서 "
                     f"{SHOWN.get(new_parent, new_parent)} 아래에 그대로 있음 (옛 엑셀 행 {pairs[0][1]}~{pairs[-1][1]}, "
                     f"새 엑셀 행 {pairs[0][3]}~{pairs[-1][3]})")
        counts["상위 삭제"] += 1
        item("상위 삭제", pn, parent, orow=prow, note=f"하위 {len(pairs)}행은 새 판에서 {SHOWN.get(new_parent, new_parent)} 아래에 그대로 "
             f"(옛 엑셀 행 {pairs[0][1]}~{pairs[-1][1]}, 새 엑셀 행 {pairs[0][3]}~{pairs[-1][3]})")
        gone = {pk_} | {p[0] for p in pairs}
        dels = [d for d in dels if d[0] not in gone]
        adds = [x for x in adds if x[0] not in {p[2] for p in pairs}]
        moved_rows["old"].update(str(p[1]) for p in pairs)
        moved_rows["new"].update(str(p[3]) for p in pairs)
    summarize.moved_rows = moved_rows                    # _diff가 칠할 칸 목록에서 뺌
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
        item("대체 후보", f"{split(d[0])[1]} -> {split(add[0])[1]}", parent, orow=d[1], nrow=add[1], note="같은 자리 (도구의 짐작)")
        dels.remove(d)
        adds.remove(add)
        counts["대체 후보"] += 1
    for path, row in adds:
        parent, pn = split(path)
        lines.append(f"[추가] {pn}: {where(parent)} (새 엑셀 행 {row})")
        item("추가", pn, parent, nrow=row)
    for path, row in dels:
        parent, pn = split(path)
        lines.append(f"[삭제] {pn}: {where(parent)} (옛 엑셀 행 {row})")
        item("삭제", pn, parent, orow=row)
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
        for r in rs:
            item("바뀜", pn, parent, r["열"], _num(r["옛 값"]), _num(r["새 값"]), rs[0].get(ocol), rs[0].get(ncol))
    return lines, counts, len(changed)


def _xml(v: Any) -> str:
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(v))
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def write_xlsx(path: str, sheets: List[Tuple[str, List[str], List[List[Any]], List[int]]], what: str) -> Tuple[str, bool]:
    """시트들 [(이름, 머리글, 행들, 열 너비)] -> xlsx (Excel 2007+ 최소 구조, 글자는 inlineStr, 정수만 숫자, 머리글 굵게·틀 고정).
    잠겨 있으면 시각을 붙인 이름. (실제 경로, 다시 열어 시트마다 행 수가 같은지)."""
    import zipfile

    def col(i: int) -> str:
        out = ""
        while i:
            i, r = divmod(i - 1, 26)
            out = chr(65 + r) + out
        return out

    def cell(ref: str, v: Any, bold: bool = False) -> str:
        if isinstance(v, int) and not isinstance(v, bool):
            return f'<c r="{ref}"><v>{v}</v></c>'
        return f'<c r="{ref}" t="inlineStr"{" s=\"1\"" if bold else ""}><is><t xml:space="preserve">{_xml(v)}</t></is></c>'

    def sheet_xml(head: List[str], body: List[List[Any]], widths: List[int]) -> str:
        rows = [f'<row r="1">' + "".join(cell(f"{col(j)}1", h, True) for j, h in enumerate(head, 1)) + "</row>"]
        for i, r in enumerate(body, 2):
            rows.append(f'<row r="{i}">' + "".join(cell(f"{col(j)}{i}", v) for j, v in enumerate(r, 1) if v not in (None, "")) + "</row>")
        cols = "".join(f'<col min="{j}" max="{j}" width="{w}" customWidth="1"/>' for j, w in enumerate(widths, 1))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                f'</sheetView></sheetViews>{"<cols>" + cols + "</cols>" if cols else ""}<sheetData>{"".join(rows)}</sheetData></worksheet>')
    ns = "http://schemas.openxmlformats.org"
    files = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{ns}/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                  for i in range(1, len(sheets) + 1)) + "</Types>",
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{ns}/package/2006/relationships"><Relationship Id="rId1" '
        f'Type="{ns}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{ns}/spreadsheetml/2006/main" xmlns:r="{ns}/officeDocument/2006/relationships"><sheets>'
        + "".join(f'<sheet name="{_xml(n[:31])}" sheetId="{i}" r:id="rId{i}"/>' for i, (n, _h, _b, _w) in enumerate(sheets, 1))
        + "</sheets></workbook>",
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{ns}/package/2006/relationships">'
        + "".join(f'<Relationship Id="rId{i}" Type="{ns}/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                  for i in range(1, len(sheets) + 1))
        + f'<Relationship Id="rId{len(sheets) + 1}" Type="{ns}/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{ns}/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="맑은 고딕"/></font>'
        '<font><b/><sz val="10"/><name val="맑은 고딕"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0"/><xf fontId="1" applyFont="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    }
    for i, (_n, head, body, widths) in enumerate(sheets, 1):
        files[f"xl/worksheets/sheet{i}.xml"] = sheet_xml(head, body, widths)

    def writer(p_: str) -> None:
        with zipfile.ZipFile(p_, "w", zipfile.ZIP_DEFLATED) as z:
            for name, text in files.items():
                z.writestr(name, text.encode("utf-8"))
    saved = rt.save_with_lock_fallback(path, writer, what)
    try:
        with zipfile.ZipFile(saved) as z:
            ok_ = all(z.read(f"xl/worksheets/sheet{i}.xml").decode("utf-8").count("<row ") == len(b) + 1
                      for i, (_n, _h, b, _w) in enumerate(sheets, 1))
    except (OSError, KeyError, zipfile.BadZipFile):
        ok_ = False
    return saved, ok_


REPORT_COLS = ["구분", "품번", "위치", "열", "옛 값", "새 값", "옛 엑셀 행", "새 엑셀 행", "비고"]


def cmd_diff(a: Any) -> Dict[str, Any]:
    return run_once(lambda pk: _diff(a, pk))


def _keep(src: str, dst: str, what: str) -> str:
    """작업 폴더(캐시)의 결과 파일을 결과 폴더로 (이름은 stamped가 정한 새 이름)."""
    with open(src, "rb") as f:
        data = f.read()

    def writer(p: str) -> None:
        with open(p, "wb") as f:
            f.write(data)
    return rt.save_with_lock_fallback(dst, writer, what)


CATEGORIES = (("Level", LEVEL_NAMES), ("품번", PART_NAMES), ("수량", QTY_NAMES), ("품명", DESC_NAMES), ("단위", UOM_NAMES))
ONLY_COLS = "한쪽 표에만 있는 열은 비교하지 않았습니다: "


def _paired_columns(ocols: List[str], ncols: List[str]) -> List[Tuple[str, str]]:
    """한 판만 열 이름이 다른 같은 뜻의 열 (Qty / 수량): 같은 열로 보고 bom-kit이 직접 비교 (Fable 검토 2026-11-14:
    pandas-kit diff가 '한쪽에만 있는 열'로 빼서 수량이 바뀌었는데 '두 판이 같습니다'였음)."""
    out = []
    for _what, names in CATEGORIES:
        oc = next((c for c in ocols if c.strip().lower() in names and c not in ncols), None)
        nc = next((c for c in ncols if c.strip().lower() in names and c not in ocols), None)
        if oc and nc:
            out.append((oc, nc))
    return out


def _paired_changes(o: Dict[str, Any], n: Dict[str, Any], pairs: List[Tuple[str, str]], gone: set) -> List[Dict[str, str]]:
    """짝 지은 열의 바뀐 칸 -> 변경내용 행 모양 ({구분, 경로, 열, 옛 값, 새 값, oldp 행, newp 행})."""
    def by_key(d: Dict[str, Any]) -> Dict[str, Dict[str, str]]:
        row_of = {r[ROW]: r for r in d["data"]}
        out = {}
        for r in csv.DictReader(io.StringIO(rt.read_text_file(d["save"]))):
            if r[ROW] in row_of:
                out[r["경로"]] = row_of[r[ROW]]
        return out
    ok_, nk = by_key(o), by_key(n)
    rows = []
    for key in [k for k in ok_ if k in nk and k not in gone]:
        for oc, nc in pairs:
            a, b = _num(ok_[key].get(oc)), _num(nk[key].get(nc))
            if a != b:
                rows.append({"구분": "바뀜", "경로": key, "열": f"{oc}(새 판 {nc})", "옛 값": a, "새 값": b,
                             "oldp 행": ok_[key].get(ROW, ""), "newp 행": nk[key].get(ROW, "")})
    return rows


WHOLE = "빈 행까지 한 표로"                 # pandas-kit이 빈 행으로 나뉜 표에서 주는 첫 [후보]


def _diff(a: Any, pk: PandasKit) -> Dict[str, Any]:
    o_csv, o_read, _ = read_bom(pk, a.old, "old", "옛 판", a.old_sheet or a.sheet, a.old_range)
    n_csv, n_read, _ = read_bom(pk, a.new, "new", "새 판", a.new_sheet or a.sheet, a.new_range)
    o, n = _paths(o_csv, a.level, a.part, "옛 판"), _paths(n_csv, a.level, a.part, "새 판")
    for d in (o, n):
        pk.run(["select", "--file", d["save"], "--name", f"{d['table']}k"])
        pk.run(["join", "--left", d["table"], "--right", f"{d['table']}k", "--on", "엑셀 행=엑셀 행번호", "--name", f"{d['table']}p"])
    pk.warnings = [w for w in pk.warnings if "엑셀 행번호" not in w]      # 경로 표의 '엑셀 행' 이름 바꿈 안내는 bom-kit 안의 일
    diff = pk.run(["diff", "--old", "oldp", "--new", "newp", "--on", "경로", "--name", "chg"])
    folder = result_dir(a.old)                            # 결과는 옛 판 옆 <옛 판 이름>_bom, 실행마다 시각을 붙인 새 이름
    stem = stamped(folder, f"diff_{file_part(os.path.splitext(os.path.basename(a.new))[0])}")
    raw = pk.run(["export", "--from", "chg", "--save", "변경내용.csv"])["output_file"]
    rows = list(csv.DictReader(io.StringIO(rt.read_text_file(raw))))
    head_cols = list(rows[0].keys()) if rows else ["구분", "경로", "열", "옛 값", "새 값", "oldp 행", "newp 행"]
    # 한 판만 이름이 다른 같은 뜻의 열은 bom-kit이 비교, 남은 '한쪽에만 있는 열'은 비교하지 못한 것으로 알림
    pairs = _paired_columns(o["cols"], n["cols"])
    paired = {c for p in pairs for c in p}
    limits: List[str] = []
    for w in [w for w in pk.warnings if w.startswith(ONLY_COLS)]:
        pk.warnings.remove(w)
        left = [c.strip() for c in w[len(ONLY_COLS):].split(",") if c.strip() and c.strip() not in paired]
        if left:
            limits.append(f"한쪽 판에만 있는 열 {', '.join(left)}은(는) 비교하지 않았습니다")
    if pairs:
        gone = {r["경로"] for r in rows if r["구분"] in ("추가", "삭제")}
        rows += _paired_changes(o, n, pairs, gone)
        pk.warnings.append("열 이름이 다른 같은 뜻의 열을 짝지어 비교했습니다: " + ", ".join(f"옛 판 {x} = 새 판 {y}" for x, y in pairs)
                           + " (이 열의 바뀐 칸은 칠할 칸 목록에 없음).")
    for d, who in ((o, "옛 판"), (n, "새 판")):
        if d["bad"]:
            limits.append(f"{who} {d['rows']}행 중 {d['bad']}행은 Level·품번을 읽지 못해 비교하지 않았습니다")
    # 변경내용 파일은 화면용 경로로 (짝 맞추는 키는 도구 안의 일)
    shown_rows = [dict(r, 경로=SHOWN.get(r["경로"], r["경로"])) for r in rows]
    changes, _ok = write_csv(stem + "_변경내용.csv", head_cols, [[r.get(c, "") for c in head_cols] for r in shown_rows], "변경내용 파일")
    order = [_siblings(o["save"]), _siblings(n["save"])]
    summarize.moved_rows, summarize.items = {"old": set(), "new": set()}, []
    lines, counts, nrows = summarize(rows, "oldp 행", "newp 행", order) if rows else ([], {}, 0)
    moved, items = summarize.moved_rows, summarize.items
    head = ("두 판이 같습니다 (찾은 변경 없음)." if not lines else
            "변경 " + ", ".join(f"{k} {v}" + ("곳" if k == "바뀐 칸" else "건") for k, v in counts.items() if v)
            + (f" ({nrows}개 품목)" if counts.get("바뀐 칸") else "") + ".")
    if limits:                                            # 비교하지 못한 것은 첫 줄에 (warnings는 뒤에 있어 빠뜨리기 쉬움)
        head = "주의: " + "; ".join(limits) + ". 아래 결과는 그것을 뺀 비교입니다.\n" + head
    files_ok = o["ok"] and n["ok"] and _ok                 # 요약은 비교 결과 엑셀의 요약 시트 (txt는 만들지 않음)
    verified = files_ok and not limits
    unverified = None if verified else (limits or ["저장한 경로·변경내용 파일을 확인하지 못함"])
    shown = lines[:LINE_MAX]
    read = f"읽은 표: 옛 판 {o_read.split(' = ', 1)[-1]}, 새 판 {n_read.split(' = ', 1)[-1]}."
    warnings = pk.warnings + o["warnings"] + n["warnings"]
    # 비교 결과 엑셀 파일 (사용자 결정 2026-11-14): 요약 시트(건수·줄·읽은 표·알림)와 변경내용 시트(한 줄에 한 변경, 바뀐 칸은 칸마다)
    summary_rows = [[x] for x in head.splitlines() + lines + [read]] + [[f"알림: {w}"] for w in warnings]
    new_name = file_part(os.path.splitext(os.path.basename(a.new))[0], 60)
    report, report_ok = write_xlsx(beside(a.old, f"{new_name}_비교결과") + ".xlsx", [
        ("요약", [f"{os.path.basename(a.old)} -> {os.path.basename(a.new)} 비교 ({datetime.now():%Y-%m-%d %H:%M})"], summary_rows, [120]),
        ("변경내용", REPORT_COLS, [[it[c] for c in REPORT_COLS] for it in items], [10, 18, 40, 14, 18, 18, 10, 10, 60])], "비교 결과 파일")
    verified = verified and report_ok
    if not report_ok:
        unverified = (unverified or []) + ["저장한 비교 결과 엑셀 파일을 다시 읽어 확인하지 못함"]
    preview = "\n".join([head] + shown + ([f"(나머지 {len(lines) - len(shown)}줄은 {report})"] if len(lines) > len(shown) else [])
                        + [read, f"비교 결과 파일: {report}"])
    then = ("preview를 그대로 사용자에게 알리세요 (줄을 직접 세거나 고쳐 쓰지 마세요). 같은 명령을 다시 실행하지 마세요."
            + (" [대체 후보]는 도구의 짐작이니 맞는지 사용자에게 확인하세요." if counts.get("대체 후보") else "")
            + ("" if lines else " 0건. 찾은 변경이 없습니다."))
    out = rt.ok("diff", preview, verified, then, unverified,
                count=len(lines), counts=counts, lines=shown, truncated=len(lines) > len(shown), output_file=report,
                changes_file=changes)                     # 변경내용 csv (_bom 결과 폴더, history가 모음)
    paint = diff.get("confirm_commands") or []
    for f in diff.get("paint_files") or []:              # 칠할 칸 목록도 결과 폴더로 (confirm_commands가 다음 턴에도 되게)
        side = "옛판" if "_칠할칸_old_" in os.path.basename(f) else "새판"
        drop = moved["old" if side == "옛판" else "new"]   # 상위만 지워져 다른 상위에 붙은 하위 행은 칠하지 않음
        cells = [x for x in rt.read_text_file(f).splitlines() if x.strip()
                 and not set(re.findall(r"[A-Z]+(\d+)", x)) & drop]
        cmd = next((c for c in paint if rt.quote_arg(f) in c["command"]), None)
        if not cells:
            if cmd:
                paint.remove(cmd)
            continue
        text = "\n".join(cells) + "\n"

        def writer(p_: str, text: str = text) -> None:
            with open(p_, "w", encoding="utf-8-sig") as h:
                h.write(text)
        kept = rt.save_with_lock_fallback(f"{stem}_칠할칸_{side}.csv", writer, "칠할 칸 파일")
        if cmd:
            cmd["command"] = cmd["command"].replace(rt.quote_arg(f), rt.quote_arg(kept))
            if drop:
                orig = os.path.basename(a.old if side == "옛판" else a.new)
                cmd["purpose"] = f"원본 {orig}의 {'삭제된' if side == '옛판' else '추가된'} 행과 바뀐 칸을 노랗게 칠하기 (사용자가 원할 때만, office-cli)"
    kept_paint = [f"{stem}_칠할칸_{x}.csv" for x in ("옛판", "새판") if os.path.isfile(f"{stem}_칠할칸_{x}.csv")]
    _write_record(stem, a, counts, head, report, changes, kept_paint, verified)     # history가 읽는 기록
    if paint:                                    # 원본에 바뀐 칸 칠하기 (office-cli, 사용자가 원할 때만)
        out["confirm_commands"] = paint
        out["agent_directive"] += " confirm_commands는 원본 엑셀에 바뀐 칸을 칠하는 명령입니다. 사용자가 칠해 달라고 할 때만 1번 실행하세요."
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


# ==============================================================================
# 5. where-used: 품목이 쓰인 상위와 제품 1대당 수량
# ==============================================================================

def cmd_where_used(a: Any) -> Dict[str, Any]:
    return run_once(lambda pk: _where_used(a, pk))


def _where_used(a: Any, pk: PandasKit) -> Dict[str, Any]:
    bom_csv, read, hidden = read_bom(pk, a.file, "bom", "BOM", a.sheet, a.range)
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
    save, ok_ = write_csv(beside(a.file, f"쓰인곳_{file_part(a.find)}") + ".csv",
                          [ROW, "품번", "품명", "상위 품번", "경로", "그 자리 수량", "제품 1대당 수량", "단위"],
                          [[i["r"].get(ROW), i["r"].get(pcol), i["r"].get(dcol) if dcol else "", i["parent"], i["path"],
                            "" if i["qty"] is None else number(i["qty"]), "" if i["total"] is None else number(i["total"]),
                            i["r"].get(ucol) if ucol else ""] for i in hits], "쓰인 곳 파일")
    warnings = pk.warnings + tree_warnings(t)
    hid = [i for i in hits if _int(i["r"].get(ROW)) in set(hidden)]
    if hid:                                      # 숨긴 폐기 품목이 '쓰인 곳'으로 세어지는 것 (집 PC 2026-11-14 B4)
        warnings.append(f"찾은 곳 중 {len(hid)}곳은 숨긴 행(또는 자동 필터로 가려진 행)입니다 (엑셀 행 "
                        f"{rows_text([i['r'].get(ROW) for i in hid])}, 품번 {', '.join(dict.fromkeys(str(i['r'].get(pcol) or '').strip() for i in hid))}). "
                        "결과와 제품 1대당 수량에 넣었습니다.")
    # 읽지 못한 행이 있으면 곳 수·합계가 모자랄 수 있음: 첫 줄에 알리고 UNVERIFIED (diff와 같은 규칙)
    limit = f"주의: BOM {len(rows)}행 중 {len(t['bad'])}행은 Level·품번을 읽지 못해 찾지 않았습니다.\n" if t["bad"] else ""
    unverified = ([limit.strip()[4:]] if limit else []) + ([] if ok_ else ["저장한 쓰인 곳 파일의 행 수를 확인하지 못함"])
    ok_ = ok_ and not limit
    unverified = unverified or None
    if not hits:
        out = rt.ok("where-used", limit + f"'{a.find}'을(를) 품번이나 품명에서 찾지 못했습니다 (BOM {len(items)}행, {read.split(' = ', 1)[-1]}).", ok_,
                    "0건. 찾은 품목이 없습니다. preview를 그대로 알리고 다시 실행하지 마세요. 품번이나 품명의 일부를 --find에 주세요.",
                    unverified, count=0, lines=[], output_file=save)
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
        more = [f"(나머지 {len(lines) - len(shown)}줄은 {save})"] if len(lines) > len(shown) else []   # diff와 같은 모양
        preview = limit + f"'{a.find}': {len(hits)}곳에 쓰입니다. " + " / ".join(sums) + ("\n" + "\n".join(shown + more) if shown else "")
        out = rt.ok("where-used", preview, ok_,
                    "preview의 줄을 사용자에게 그대로 알리세요 (줄을 직접 세거나 고쳐 쓰지 마세요). 같은 명령을 다시 실행하지 마세요."
                    + (" 잘렸으면 전체는 output_file에 있습니다." if len(lines) > LINE_MAX else ""),
                    unverified, count=len(hits), lines=shown, totals=sums, truncated=len(lines) > LINE_MAX, output_file=save)
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


# ==============================================================================
# 6. history: 한 BOM의 지난 diff 결과를 시간순으로 (변경 이력)
# ==============================================================================

def _file_id(path: str) -> str:
    return os.path.normcase(os.path.abspath(str(path)))


def _stat(path: str) -> Dict[str, Any]:
    try:
        st = os.stat(path)
        return {"mtime": int(st.st_mtime), "size": st.st_size}
    except OSError:
        return {}


def _write_record(stem: str, a: Any, counts: Dict[str, int], head: str, summary: str, changes: str,
                  paint: List[str], verified: bool) -> None:
    """diff 1번 = 기록 하나 (<결과 폴더>/diff_<새 판>_<시각>_기록.json). history가 결과 폴더들에서 모아 읽음."""
    rec = {"tool": rt.TOOL_NAME, "version": rt.VERSION, "at": datetime.now().isoformat(timespec="seconds"),
           "old": os.path.abspath(a.old), "new": os.path.abspath(a.new), "old_file": _stat(a.old), "new_file": _stat(a.new),
           "sheet": a.sheet or a.old_sheet or "", "new_sheet": a.new_sheet or "", "counts": counts,
           "head": head.splitlines()[-1], "verified": verified, "summary": summary, "changes": changes, "paint": paint}
    try:
        text = json.dumps(rec, ensure_ascii=False, indent=1)
        with open(stem + "_기록.json", "w", encoding="utf-8") as f:
            f.write(text)
    except OSError as e:
        sys.stderr.write(f"[WARN] 비교 기록을 저장하지 못했습니다 ({type(e).__name__}). 이 비교는 history에 나오지 않습니다.\n")


def _records(start: str) -> List[Dict[str, Any]]:
    """start 파일과 비교로 이어진 모든 판의 기록 (옛 판·새 판을 따라 이어 감). 기록은 각 옛 판 옆 결과 폴더에 있음."""
    found: Dict[str, Dict[str, Any]] = {}
    seen_dirs: set = set()
    files = {_file_id(start)}
    todo = [_file_id(os.path.dirname(os.path.abspath(start)))]
    while todo:
        d = todo.pop()
        if d in seen_dirs:
            continue
        seen_dirs.add(d)
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for n in names:
            folder = os.path.join(d, n)
            if not os.path.isfile(os.path.join(folder, RESULT_MARKER)):
                continue
            for f in os.listdir(folder):
                p = _file_id(os.path.join(folder, f))
                if f.endswith("_기록.json") and p not in found:
                    try:
                        with open(p, encoding="utf-8") as h:
                            found[p] = json.load(h)
                    except (OSError, ValueError):
                        continue
        grew = True
        while grew:                                      # 비교로 이어진 판 모으기
            grew = False
            for rec in found.values():
                ends = {_file_id(rec.get("old", "")), _file_id(rec.get("new", ""))}
                if ends & files and not ends <= files:
                    files |= ends
                    grew = True
        todo.extend(os.path.dirname(f) for f in files)
    out = [dict(r, record=p) for p, r in found.items() if {_file_id(r.get("old", "")), _file_id(r.get("new", ""))} & files]
    last: Dict[str, Dict[str, Any]] = {}                 # 파일이 그대로인 같은 비교를 여러 번 했으면 마지막 것만
    for r in sorted(out, key=lambda r: r.get("at", "")):
        same = json.dumps([_file_id(r.get("old", "")), _file_id(r.get("new", "")), r.get("old_file"), r.get("new_file"),
                           r.get("sheet"), r.get("new_sheet")], sort_keys=True)
        last.pop(same, None)
        last[same] = r
    return sorted(last.values(), key=lambda r: r.get("at", ""))


def cmd_history(a: Any) -> Dict[str, Any]:
    _retry_pending()
    start = os.path.abspath(str(a.file).strip().strip('"'))
    if not os.path.isfile(start):
        raise FileNotFoundError(f"파일이 없습니다: {start}")
    recs = _records(start)
    name = os.path.basename(start)
    if not recs:
        return rt.ok("history", f"'{name}'로 한 두 판 비교 기록이 없습니다 (찾은 비교 0건).", True,
                     "0건. 이 파일로 한 비교 기록이 없습니다. preview를 그대로 알리고 다시 실행하지 마세요. "
                     "두 판을 비교하려면 mbom-diff 스킬을 쓰세요.", None, count=0, lines=[])
    lines: List[str] = []
    warnings: List[str] = []
    head_cols = ["비교 시각", "옛 판", "새 판"]
    body: List[Dict[str, str]] = []
    for r in recs:
        at = r.get("at", "").replace("T", " ")[:16]
        old, new = os.path.basename(r.get("old", "")), os.path.basename(r.get("new", ""))
        lines.append(f"[비교] {at} {old} -> {new}: {r.get('head', '')}")
        for side in ("old", "new"):
            now, then_ = _stat(r.get(side, "")), r.get(f"{side}_file") or {}
            if not now:
                warnings.append(f"{at} 비교의 {os.path.basename(r.get(side, ''))}이(가) 지금 없습니다.")
            elif then_ and now != then_:
                warnings.append(f"{at} 비교 뒤 {os.path.basename(r.get(side, ''))}이(가) 바뀌었습니다 (기록은 그때의 비교).")
        if not r.get("verified", True):
            warnings.append(f"{at} 비교는 확인하지 못한 것이 있었습니다 (UNVERIFIED, 그 비교의 요약 파일 첫 줄).")
        path = r.get("changes", "")
        if not os.path.isfile(path):
            warnings.append(f"{at} 비교의 변경내용 파일이 없습니다: {path}")
            continue
        for row in csv.DictReader(io.StringIO(rt.read_text_file(path))):
            for c in row:
                if c not in head_cols:
                    head_cols.append(c)
            body.append(dict(row, **{"비교 시각": at, "옛 판": old, "새 판": new}))
    save, ok_ = write_csv(beside(start, "변경이력") + ".csv", head_cols,
                          [[b.get(c, "") for c in head_cols] for b in body], "변경 이력 파일")
    shown = lines[:LINE_MAX]
    more = [f"(나머지 {len(lines) - len(shown)}줄은 {save})"] if len(lines) > len(shown) else []
    versions = {os.path.basename(r.get(s, "")) for r in recs for s in ("old", "new")}
    preview = "\n".join([f"'{name}'와 이어진 두 판 비교 {len(recs)}건 (판 {len(versions)}개, 오래된 것부터)."] + shown + more
                        + [f"변경 이력 파일: {save} (변경 {len(body)}줄)"])
    out = rt.ok("history", preview, ok_, "preview를 그대로 사용자에게 알리세요 (줄을 직접 세거나 고쳐 쓰지 마세요). 같은 명령을 다시 실행하지 마세요.",
                None if ok_ else ["저장한 변경 이력 파일의 행 수를 확인하지 못함"], count=len(recs), lines=shown,
                truncated=len(lines) > len(shown), output_file=save)
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


# ==============================================================================
# 7. Parser
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
    p = rt.add_command(sub, root, "history", "한 BOM의 지난 두 판 비교(diff) 결과를 시간순으로 모으기 (변경 이력)")
    p.add_argument("--file", required=True, help="그 BOM의 어느 판이든 (.xlsx, .xlsm, .csv)")
    return root


COMMANDS = {"diff": cmd_diff, "where-used": cmd_where_used, "history": cmd_history}


def main(argv: Optional[List[str]] = None) -> None:
    rt.setup_console()
    rt.CURRENT_ARGV = [os.path.abspath(__file__)] + list(sys.argv[1:] if argv is None else argv)
    args = build_cli().parse_args(rt.CURRENT_ARGV[1:])
    rt.run_cli(lambda: COMMANDS[args.cmd](args), hint="확인할 것: 파일 경로, 시트, 열 이름.")


if __name__ == "__main__":
    main()
