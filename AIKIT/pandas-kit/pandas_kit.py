# -*- coding: utf-8 -*-
"""
pandas_kit.py - pandas-kit: table data handled step by step (complete task tool, first target: BOM analysis).

Commands: select, show, join, diff, only, stack, undo, new, export, chart.
    python pandas_kit.py select --file "BOM.xlsx" --range B5:H230 --where "Level=1;Qty>2" --name df1
    python pandas_kit.py select --from df1 --where "Maker=A|B" --name df2
    python pandas_kit.py show --from df2
    python pandas_kit.py diff --old df1 --new df2 --on "Parent,Part Number"

The source files are never written. Every step is recorded in the work folder (steps.json, the original) and its table is
kept as a pickle snapshot (a cache: rebuilt from the record when broken or made by another pandas version). Excel row
numbers (_row) and column letters are kept so results point back to the cells of the source file.
stdout is exactly one JSON line (Agent CLI Protocol, AIKIT/dev/README.md); logs go to stderr.
"""

from __future__ import annotations

import csv
import difflib
import io
import json
import numbers
import os
import pickle
import re
import sys
import time
from contextlib import contextmanager
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pd_runtime as rt  # noqa: E402

TABLE_EXTS = (".xlsx", ".xlsm", ".csv")
SHOW_MAX = 20                      # rows in the JSON of show (R2: the rest in a csv file)
LIST_MAX = 30                      # columns listed in a JSON answer
LOCK_FILE = ".lock"                # 작업 기록을 고치는 동안 잡는 잠금 (같은 작업 폴더의 동시 실행)
LOCK_WAIT, LOCK_STALE = 20.0, 120.0  # 잠금을 기다리는 초, 이보다 오래된 잠금은 끝나지 않은 실행이 남긴 것으로 보고 버림
TABLE_CHOICES = 6                  # select에 --file·--from이 없을 때 [후보]로 보이는 표 수 (최근 것부터)
READ_ONLY_BYTES = 30 * 1024 * 1024  # bigger workbooks are read in openpyxl read-only mode (no hidden-row/merged-cell check)
MARKER = ".created_by_pandas-kit"
STEPS_FILE = "steps.json"
ROW = "_row"                       # Excel row number of the source file (hidden column, kept through every step)
ROW_SHOWN = "엑셀 행"
EMPTY_SHOWN = "(빈 칸)"             # group의 빈 값 묶음 이름 (사용자 결정 2026-11-14, null 대신)
OPS = (">=", "<=", "!=", "!~", "^=", "$=", "=", ">", "<", "~")
_COND = re.compile(r"^\s*(.+?)\s*(>=|<=|!=|!~|\^=|\$=|=|>|<|~)\s*(.*?)\s*$", re.S)
_RANGE = re.compile(r"^([A-Z]{1,3})?(\d+)?(?::([A-Z]{1,3})?(\d+)?)?$")
_LETTERS = re.compile(r"^[A-Z]{1,3}$")
_NAME = re.compile(r"^[^\W\d][\w]*$")
_NUM_TEXT = re.compile(r"^-?\d+(\.\d+)?$")
_NUM_COMMA = re.compile(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$")        # 1,200 (회사 파일에 흔한 천 단위 쉼표)
_NUM_SYMBOL = re.compile(r"^[\u20a9$\u20ac\u00a5\u00a3]\s*-?\d[\d,]*(\.\d+)?$"      # 통화 기호가 앞에
                         r"|^-?\d[\d,]*(\.\d+)?\s*[^\d\s.\-]{1,4}$")     # 단위가 뒤에 (12 kg, 18%). P100, MC-03 같은 코드는 아님
_TEXT_DATE = re.compile(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$")                                       # 2026.01.05
_LEADING_ZERO = re.compile(r"^-?0\d")
_DATE_TEXT = re.compile(r"^\d{4}-\d{1,2}-\d{1,2}( \d{1,2}:\d{2}(:\d{2})?)?$")      # 날짜 열의 조건 값: 연-월-일 (시각은 붙여도 됨)


def pd_mod() -> Any:
    return rt.require_lib("pandas", "pandas", "pandas-kit")


# ==============================================================================
# 1. Small helpers
# ==============================================================================

def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def split_outside_quotes(text: str, sep: str) -> List[str]:
    """Splits on sep outside '...' and "..." quotes."""
    out, cur, q = [], "", ""
    for ch in text:
        if q:
            cur += ch
            if ch == q:
                q = ""
        elif ch in "'\"":
            q = ch
            cur += ch
        elif ch == sep:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def unquote(v: str) -> str:
    v = v.strip()
    return v[1:-1] if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"" else v


def name_list(names: List[Any], limit: int = LIST_MAX) -> str:
    names = [str(n) for n in names]
    return ", ".join(names[:limit]) + (f" 외 {len(names) - limit}개" if len(names) > limit else "")


def is_row_col(c: Any) -> bool:
    return str(c) == ROW or str(c).endswith(" 행")


def data_cols(df: Any) -> List[str]:
    return [c for c in df.columns if not is_row_col(c)]


def kind_of(series: Any) -> str:
    pd = pd_mod()
    if pd.api.types.is_bool_dtype(series):
        return "글자"
    if pd.api.types.is_numeric_dtype(series):
        return "숫자"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "날짜"
    return "빈 열" if series.isna().all() else "글자"


def cell_json(v: Any) -> Any:
    return rt.json_safe(v)


def argv_now() -> List[str]:
    return [str(a) for a in (rt.CURRENT_ARGV or [""])[1:]]


def argv_replacing(option: str, old: str, new: str) -> Optional[List[str]]:
    """The running command with the first `old` inside the value of `option` replaced by `new` (for [후보] commands)."""
    args = argv_now()
    for i, a in enumerate(args[:-1]):
        if a == option and old in args[i + 1]:
            fixed = list(args)
            fixed[i + 1] = args[i + 1].replace(old, new, 1)
            return fixed
    return None


def argv_setting(option: str, value: str) -> List[str]:
    """The running command with option set to value (added when missing)."""
    args = argv_now()
    if option in args and args.index(option) + 1 < len(args):
        i = args.index(option)
        return args[:i + 1] + [value] + args[i + 2:]
    return args + [option, value]


# ==============================================================================
# 2. Work folder: the step record (original) and table snapshots (cache)
# ==============================================================================

def pointer_path() -> str:
    return os.path.join(rt.get_data_dir(), "current.json")


def current_work() -> Optional[str]:
    try:
        with open(pointer_path(), encoding="utf-8") as f:
            path = json.load(f).get("work")
    except (OSError, ValueError):
        return None
    return path if path and os.path.isfile(os.path.join(path, MARKER)) else None


def set_current_work(path: Optional[str]) -> None:
    os.makedirs(rt.get_data_dir(), exist_ok=True)
    text = json.dumps({"work": path}, ensure_ascii=False)
    rt.retry_fs(lambda: open(pointer_path(), "w", encoding="utf-8").write(text))


def new_work_folder(source: str) -> str:
    """<folder of the first source>/<source name>_YYMMDDHHMM (사용자 결정 2026-11-13; _2, _3 ... when made again in the same minute)."""
    base = os.path.splitext(os.path.basename(source))[0]
    stem = os.path.join(os.path.dirname(os.path.abspath(source)), f"{base}_{datetime.now():%y%m%d%H%M}")
    path, n = stem, 2
    while True:                           # os.mkdir가 실패하면 다른 실행이 같은 이름을 먼저 만든 것: 다음 번호 (동시 실행)
        try:
            os.mkdir(path)
            break
        except FileExistsError:
            path, n = f"{stem}_{n}", n + 1
    os.makedirs(os.path.join(path, "tables"))
    with open(os.path.join(path, MARKER), "w", encoding="utf-8") as f:
        f.write(f"{rt.TOOL_NAME} {rt.VERSION}\n")
    with open(os.path.join(path, STEPS_FILE), "w", encoding="utf-8") as f:
        json.dump({"tool": rt.TOOL_NAME, "steps": []}, f, ensure_ascii=False)
    return path


class Work:
    """One work folder: steps.json (the record, original) and tables/s<id>.pkl (snapshots, rebuilt from the record)."""

    def __init__(self, folder: str) -> None:
        if not os.path.isfile(os.path.join(folder, MARKER)):
            raise rt.InputError(f"pandas-kit 작업 폴더가 아닙니다 (도구가 만든 표시 파일이 없음): {folder}")
        self.folder = os.path.abspath(folder)
        try:
            with open(os.path.join(self.folder, STEPS_FILE), encoding="utf-8") as f:
                self.steps: List[Dict[str, Any]] = json.load(f)["steps"]
        except (OSError, ValueError, KeyError) as e:
            raise rt.InputError(f"작업 기록({STEPS_FILE})을 읽지 못했습니다 ({type(e).__name__}): {self.folder}. "
                                f"새 작업은 python {rt.quote_arg(rt.script_path())} new 뒤 select --file로 시작하세요.") from None
        self.warnings: List[str] = []

    # -- record --
    def save(self) -> None:
        text = json.dumps({"tool": rt.TOOL_NAME, "version": rt.VERSION, "steps": self.steps}, ensure_ascii=False, indent=1)
        path = os.path.join(self.folder, STEPS_FILE)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(text)
        rt.retry_fs(lambda: os.replace(tmp, path))

    def names(self) -> Dict[str, int]:
        out: Dict[str, int] = {}
        for s in self.steps:
            out[s["name"]] = s["id"]
        return out

    def current_name(self) -> Optional[str]:
        return self.steps[-1]["name"] if self.steps else None

    def step(self, sid: int) -> Dict[str, Any]:
        return next(s for s in self.steps if s["id"] == sid)

    def resolve(self, name: Optional[str], option: str) -> Tuple[str, int]:
        """Table name -> (name, step id). 표 이름은 늘 적어야 한다: 빠지면 마지막 표로 짐작하지 않고 멈춘다
        (사용자 결정 2026-11-14, 집 PC P7: 마지막 표가 0행이었음). 표가 하나면 [추천], 여럿이면 [후보] (최근 TABLE_CHOICES개, 행 수)."""
        names = self.names()
        if not names:
            raise rt.InputError("아직 만든 표가 없습니다. 먼저 select --file로 파일을 읽으세요: "
                                f"python {rt.quote_arg(rt.script_path())} select --file \"<파일>\"")
        if not name:
            head = f"{option}이(가) 없습니다. 쓸 표를 늘 {option} <표 이름>으로 적습니다 (도구가 짐작하지 않음)"
            tables = sorted(names.items(), key=lambda kv: kv[1])[-TABLE_CHOICES:]
            if len(tables) == 1:
                raise rt.InputError(f"{head}. [추천] {rt.full_command(argv_setting(option, tables[0][0]))}")
            labels = [f"표 {t}" + (f"({self.step(sid).get('rows')}행)" if self.step(sid).get("rows") is not None else "")
                      for t, sid in tables]
            err = rt.ask(labels, [argv_setting(option, t) for t, _ in tables], "쓸 표")
            err.args = (f"{head}. " + err.args[0],) + err.args[1:]
            raise err
        if name in names:
            return name, names[name]
        close = difflib.get_close_matches(name, list(names), n=5, cutoff=0.5) or list(names)[-5:]
        raise rt.ask(close, [argv_setting(option, c) for c in close], f"{option} 표 '{name}'이(가) 없어 쓸 표")

    def next_name(self) -> str:
        used = set(self.names())
        n = 1
        while f"df{n}" in used:
            n += 1
        return f"df{n}"

    # -- lock: 기록 읽기-고치기-쓰기를 한 실행씩 (Fable 검토 2026-11-14 A1: 동시 실행에서 단계가 사라지고 스냅샷이 뒤바뀜) --
    @contextmanager
    def locked(self):
        path = os.path.join(self.folder, LOCK_FILE)
        end = time.time() + LOCK_WAIT
        while True:
            try:
                fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, f"{os.getpid()} {datetime.now().isoformat(timespec='seconds')}".encode())
                os.close(fd)
                break
            except FileExistsError:
                try:
                    if time.time() - os.path.getmtime(path) > LOCK_STALE:
                        os.remove(path)                     # 끝나지 않은 실행이 남긴 잠금
                        sys.stderr.write(f"[WARN] 오래된 잠금을 버렸습니다 ({path})\n")
                        continue
                except OSError:
                    pass
                if time.time() > end:
                    raise rt.InputError("다른 실행이 이 작업 폴더의 기록을 쓰고 있습니다. 그 실행이 끝난 뒤 같은 명령을 다시 실행하세요.") from None
                time.sleep(0.1)
            except OSError:
                if time.time() > end:
                    raise
                time.sleep(0.1)
        try:
            self.reload()                                   # 기다리는 사이 다른 실행이 더한 단계까지 읽음
            yield
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    def reload(self) -> None:
        with open(os.path.join(self.folder, STEPS_FILE), encoding="utf-8") as f:
            self.steps = json.load(f)["steps"]

    # -- snapshots --
    def table_file(self, sid: int) -> str:
        return os.path.join(self.folder, "tables", f"s{sid}.pkl")

    @staticmethod
    def _tag(step: Dict[str, Any]) -> List[Any]:
        """스냅샷이 어느 단계의 것인지 (기록과 다르면 버리고 기록으로 다시 만듦)."""
        return [step["id"], step["name"], step["lineage"], step.get("at")]

    def write_table(self, sid: int, df: Any, meta: Dict[str, Any], step: Optional[Dict[str, Any]] = None) -> bool:
        """Writes the snapshot and reads it back (R1). True when the read-back table has the same shape."""
        pd = pd_mod()
        path = self.table_file(sid)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tag = self._tag(step or self.step(sid))
        data = pickle.dumps({"pandas": pd.__version__, "df": df, "meta": meta, "step": tag}, protocol=pickle.HIGHEST_PROTOCOL)
        rt.retry_fs(lambda: open(path, "wb").write(data))
        try:
            back = self._read_pickle(path)
            return back is not None and back["df"].shape == df.shape and back.get("step") == tag
        except Exception:
            return False

    def _read_pickle(self, path: str) -> Optional[Dict[str, Any]]:
        """Only pickles of this tool's own work folder are read (decision 2026-11-13). None when broken or other pandas."""
        pd = pd_mod()
        tables = os.path.join(self.folder, "tables")
        if os.path.dirname(os.path.abspath(path)) != tables or not os.path.isfile(os.path.join(self.folder, MARKER)):
            return None
        try:
            with open(path, "rb") as f:
                obj = pickle.load(f)
        except Exception:
            return None
        if not isinstance(obj, dict) or obj.get("pandas") != pd.__version__ or "df" not in obj:
            return None
        return obj

    def load(self, sid: int, depth: int = 0) -> Tuple[Any, Dict[str, Any]]:
        """The table of a step: its snapshot, or rebuilt from the record (JIT recovery, reported in recovered)."""
        if depth == 0 and getattr(self, "_check", False):
            self._check_sources_of(sid)
        obj = self._read_pickle(self.table_file(sid)) if os.path.isfile(self.table_file(sid)) else None
        if obj is not None and obj.get("step") == self._tag(self.step(sid)):
            return obj["df"], obj["meta"]
        if depth > 200:
            raise rt.InputError("작업 기록을 다시 실행하다가 너무 깊어졌습니다. new로 새 작업을 시작하세요.")
        step = self.step(sid)
        inputs = {k: self.load(v, depth + 1) for k, v in step.get("inputs", {}).items()}
        df, meta, _info = run_op(step["op"], step["args"], inputs, step["name"])
        self.write_table(sid, df, meta)
        rt.note_recovered(f"표 {step['name']}의 스냅샷이 없거나 깨짐(또는 pandas 버전·기록과 다름) -> 작업 기록으로 다시 만듦")
        return df, meta

    def add(self, op: str, name: str, args: Dict[str, Any], inputs: Dict[str, int], df: Any, meta: Dict[str, Any],
            lineage: str, source: Optional[Dict[str, Any]] = None) -> Tuple[int, bool]:
        with self.locked():
            sid = max((s["id"] for s in self.steps), default=0) + 1
            old = self.names().get(name)
            step = {"id": sid, "op": op, "name": name, "args": args, "inputs": inputs, "lineage": lineage, "rows": len(df),
                    "at": datetime.now().isoformat(timespec="seconds")}
            if source:
                step["source"] = source
            verified = self.write_table(sid, df, meta, step)
            self.steps.append(step)
            self.save()
            try:                                            # 저장한 기록을 다시 읽어 이 단계가 있는지 (R1)
                with open(os.path.join(self.folder, STEPS_FILE), encoding="utf-8") as f:
                    verified = verified and step in json.load(f)["steps"]
            except (OSError, ValueError, KeyError):
                verified = False
        if old is not None:
            before, now = self.source_files(old), self.source_files(sid)
            if before != now:                               # 이름 가림 (Fable 검토 2026-11-14 A3, 사용자 결정: 경고)
                self.warnings.append(f"표 {name}은(는) {', '.join(sorted(before)) or '다른 표'}에서 만든 표였는데, 이제 "
                                     f"{', '.join(sorted(now)) or '다른 표'}의 표입니다 (옛 {name}은 undo로 되돌릴 수 있음).")
            stale = sorted({s["name"] for s in self.steps if old in s.get("inputs", {}).values() and self.names().get(s["name"]) == s["id"]})
            if stale:
                self.warnings.append(f"{name}을(를) 새로 만들었습니다. {', '.join(stale)}은(는) 옛 {name} 기준입니다 (다시 계산하지 않음).")
        return sid, verified

    def check_sources(self) -> None:
        """원본 확인을 켬: 이 실행에서 처음 읽는 표마다 그 표의 조상 원본만 확인한다 (Fable 검토 2026-11-14 A4:
        작업 안의 모든 원본을 매번 확인해 관계없는 표에도 경고가 붙었음)."""
        self._check = True
        self._checked: set = set()

    def ancestors(self, sid: int) -> List[int]:
        out, todo = [], [sid]
        while todo:
            x = todo.pop()
            if x in out:
                continue
            out.append(x)
            todo.extend(self.step(x).get("inputs", {}).values())
        return out

    def source_files(self, sid: Optional[int]) -> set:
        """표가 나온 원본 파일 이름들 (조상 단계를 따라감)."""
        if sid is None or sid not in {s["id"] for s in self.steps}:
            return set()
        return {os.path.basename(self.step(x)["source"]["file"]) for x in self.ancestors(sid) if self.step(x).get("source")}

    def _check_sources_of(self, sid: int) -> None:
        """원본이 읽은 뒤 바뀌었으면: 다시 읽어 값이 같으면(서식만 바뀜, 예: 바뀐 칸 칠하기) 기록만 조용히 갱신,
        다르면 다시 읽는 전체 명령과 함께 경고."""
        for x in self.ancestors(sid):
            s = self.step(x)
            src = s.get("source")
            if not src or x in self._checked:
                continue
            self._checked.add(x)
            try:
                st = os.stat(src["file"])
                changed = (int(st.st_mtime), st.st_size) != (src["mtime"], src["size"])
            except OSError:
                self.warnings.append(f"원본 {os.path.basename(src['file'])}이(가) 없습니다 (표 {s['name']}).")
                continue
            if not changed:
                continue
            try:
                fresh, _meta, _info = run_op(s["op"], s["args"], {}, s["name"])
                self._check = False                          # 아래 load가 다시 확인하지 않게
                same = fresh.equals(self.load(x)[0])
            except Exception:
                same = False
            finally:
                self._check = True
            if same:
                with self.locked():
                    for t in self.steps:
                        if t["id"] == x and t.get("source"):
                            t["source"].update(mtime=int(st.st_mtime), size=st.st_size)
                    self.save()
                sys.stderr.write(f"[INFO] 원본 {os.path.basename(src['file'])}이(가) 저장됐지만 값은 같아 그대로 씁니다 (표 {s['name']})\n")
                continue
            again = rt.full_command(["select", "--file", src["file"]] + _select_extra(s["args"]) + ["--name", s["name"]])
            self.warnings.append(f"원본 {os.path.basename(src['file'])}이(가) 읽은 뒤 바뀌었습니다 (값이 다름). 행 위치가 달라졌을 수 있습니다. "
                                 f"다시 읽으려면: {again}")


def _select_extra(a: Dict[str, Any]) -> List[str]:
    out: List[str] = []
    for opt in ("sheet", "range", "columns", "where"):
        if a.get(opt):
            out += [f"--{opt}", str(a[opt])]
    return out


SWITCHED: List[str] = []         # --work로 지금 작업 폴더가 바뀐 일 (preview 앞에 알림, 사용자 결정 2026-11-14)


def open_work(args: Any, create_for: Optional[str] = None) -> Work:
    """--work <folder> (switch), else the current work folder, else a new one next to create_for (select --file).
    --work는 어느 명령에서나 '그 작업 폴더로 옮겨 가기'이고, 옮겼으면 결과 preview 앞에 알린다 (보기 명령도 같음)."""
    if getattr(args, "work", None):
        work = Work(os.path.abspath(args.work))
        before = current_work()
        if not before or os.path.normcase(os.path.abspath(before)) != os.path.normcase(work.folder):
            SWITCHED.append(f"지금 작업 폴더를 {os.path.basename(work.folder)}(으)로 바꿨습니다"
                            + (f" (전: {os.path.basename(before)})" if before else "")
                            + ". 다음 명령도 이 작업 폴더의 표를 씁니다.")
        set_current_work(work.folder)
        return work
    cur = current_work()
    if cur:
        return Work(cur)
    if create_for:
        folder = new_work_folder(create_for)
        set_current_work(folder)
        sys.stderr.write(f"[INFO] 새 작업 폴더: {folder}\n")
        return Work(folder)
    raise rt.InputError("지금 작업 폴더가 없습니다. 먼저 select --file로 파일을 읽으세요: "
                        f"python {rt.quote_arg(rt.script_path())} select --file \"<파일>\"")


# ==============================================================================
# 3. Reading a source file (xlsx, xlsm, csv): anchor, Excel rows, column letters, values unchanged
# ==============================================================================

def parse_range(text: Optional[str]) -> Optional[Tuple[Optional[int], Optional[int], Optional[int], Optional[int]]]:
    """'B5:H230', 'B:H', 'B5' (start cell) -> (min_col, min_row, max_col, max_row), None parts = open."""
    if not text:
        return None
    t = str(text).strip().upper().replace("$", "")
    m = _RANGE.match(t)
    if not m or not any(m.groups()):
        raise rt.InputError(f"--range '{text}'를 읽을 수 없습니다. 예: B5:H230, B:H, B5 (표의 왼쪽 위 칸).")
    c1, r1, c2, r2 = m.groups()
    one_cell = ":" not in t
    return (col_index(c1) if c1 else None, int(r1) if r1 else None,
            None if one_cell else (col_index(c2) if c2 else None), None if one_cell else (int(r2) if r2 else None))


def resolve_file(path: str) -> str:
    """--file (Excel에 열려 있으면 저장됐는지만 봄). 열린 통합문서를 고르는 일은 _select_needs_input."""
    full = os.path.abspath(str(path).strip().strip('"'))
    if not os.path.isfile(full):
        raise FileNotFoundError(f"파일이 없습니다: {full}")
    if os.path.splitext(full)[1].lower() not in TABLE_EXTS:
        raise rt.InputError(f"읽을 수 있는 파일은 {', '.join(TABLE_EXTS)}입니다: {os.path.basename(full)}")
    for wb in rt.open_workbooks():
        if wb["path"] and os.path.normcase(wb["path"]) == os.path.normcase(full) and not wb["saved"]:
            raise rt.InputError(f"{os.path.basename(full)}이(가) Excel에서 열려 있고 저장되지 않은 변경이 있습니다. "
                                "Excel에서 저장한 뒤 같은 명령을 다시 실행하세요 (도구가 대신 저장하지 않음).")
    return full


def _csv_rows(path: str) -> List[List[Any]]:
    text = rt.read_text_file(path)
    sample = text[:20000]
    try:
        delim = csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
    except csv.Error:
        delim = ","
    return [list(r) for r in csv.reader(io.StringIO(text), delimiter=delim)]


def _empty(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v == "")


def read_grid(path: str, sheet: Optional[str], rng: Optional[Tuple]) -> Dict[str, Any]:
    """The cells of the area to read: {"sheet", "rows": [(excel_row, [values])], "min_col", "max_col", "warnings", ...}."""
    warnings: List[str] = []
    ext = os.path.splitext(path)[1].lower()
    hidden: set = set()
    merged: List[Tuple[int, int, int, int]] = []
    formulas_empty: set = set()
    if ext == ".csv":
        if sheet:
            warnings.append("csv 파일에는 시트가 없어 --sheet를 쓰지 않았습니다.")
        all_rows = _csv_rows(path)
        sheet_name = os.path.basename(path)
        max_row = len(all_rows)
        max_col = max((len(r) for r in all_rows), default=0)
        get = lambda r1, r2, c1, c2: [(r, [(all_rows[r - 1][c - 1] if c - 1 < len(all_rows[r - 1]) else None)
                                          for c in range(c1, c2 + 1)]) for r in range(r1, r2 + 1)]
    else:
        openpyxl = rt.require_lib("openpyxl", "openpyxl", "엑셀 읽기")
        big = os.path.getsize(path) > READ_ONLY_BYTES
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=big)
        except Exception as e:
            raise rt.InputError(f"엑셀 파일을 열지 못했습니다 ({type(e).__name__}): {os.path.basename(path)}") from None
        names = list(wb.sheetnames)
        if sheet and sheet not in names:
            low = {n.strip().lower(): n for n in names}
            if sheet.strip().lower() in low:
                sheet = low[sheet.strip().lower()]
            else:
                raise rt.ask(names, [argv_setting("--sheet", n) for n in names], f"시트 '{sheet}'이(가) 없어 읽을 시트")
        ws = wb[sheet] if sheet else wb.active
        sheet_name = ws.title
        max_row, max_col = int(ws.max_row or 0), int(ws.max_column or 0)
        if big:
            warnings.append("큰 파일이라 숨긴 행·병합 칸·계산 안 된 수식은 확인하지 않았습니다.")
        else:
            hidden = {r for r, d in ws.row_dimensions.items() if d.hidden}
            merged = [(m.min_col, m.min_row, m.max_col, m.max_row) for m in ws.merged_cells.ranges]
            try:
                wf = openpyxl.load_workbook(path, data_only=False)[sheet_name]
                for row in wf.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith("="):
                            formulas_empty.add((c.row, c.column))
            except Exception:
                pass

        def get(r1: int, r2: int, c1: int, c2: int) -> List[Tuple[int, List[Any]]]:
            if r2 < r1 or c2 < c1:
                return []
            return [(r1 + i, list(vals)) for i, vals in
                    enumerate(ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True))]
    c1, r1, c2, r2 = rng or (None, None, None, None)
    explicit_end = rng is not None and (c2 is not None or r2 is not None)
    c1, r1 = c1 or 1, r1 or 1
    c2, r2 = c2 or max_col, r2 or max_row
    rows = get(r1, r2, c1, c2)
    return {"sheet": sheet_name, "rows": rows, "c1": c1, "c2": c2, "r2": r2, "explicit_end": explicit_end,
            "warnings": warnings, "hidden": hidden, "merged": merged, "formulas": formulas_empty}


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_number_text(v: Any) -> bool:
    return isinstance(v, str) and bool(_NUM_TEXT.match(v)) and not _LEADING_ZERO.match(v)


def _is_comma_number(v: Any) -> bool:
    return isinstance(v, str) and bool(_NUM_COMMA.match(v.strip()))


def _text_date(v: Any) -> Optional[date]:
    """'2026.01.05', '2026-1-5' -> date. 연이 앞인 네 자리일 때만 (05/01/2026처럼 뜻이 둘인 것은 글자 그대로)."""
    if not isinstance(v, str):
        return None
    m = _TEXT_DATE.match(v.strip())
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def as_numbers(series: Any) -> Any:
    """글자 열의 칸을 숫자로 (쉼표·공백과 앞뒤 기호를 뺌: '1,200' -> 1200, 'W1,500' -> 1500, '12 kg' -> 12).
    비교·정렬에만 쓰고 표의 값은 바꾸지 않는다."""
    pd = pd_mod()
    text = series.astype("string").str.strip()
    text = text.str.replace(r"[,\s]", "", regex=True)
    text = text.str.replace(r"^[^\d\-+.]+", "", regex=True).str.replace(r"[^\d]+$", "", regex=True)
    return pd.to_numeric(text, errors="coerce")


def _rows_text(rs: List[int], limit: int = 5) -> str:
    return ", ".join(str(r) for r in rs[:limit]) + (" 등" if len(rs) > limit else "")


def column_read_warning(name: str, letter: str, cells: List[Tuple[int, Any]], kind: str) -> Optional[Tuple[str, str, List[int]]]:
    """A column whose cells are not all of one kind -> (분류, '<이름>(<글자>)', 그 칸들의 엑셀 행). 여러 열이면 read_table이 한 줄로 묶는다
    (집 PC 2026-11-13, MBOM: 글자로 저장된 숫자 한 칸이 열 전체를 글자로 만들고 'Scrap %>3'이 글자 비교였음)."""
    if kind == "숫자":
        comma = [r for r, v in cells if _is_comma_number(v)]
        if comma:
            return ("쉼표 숫자", f"{name}({letter})", comma)
        txt = [r for r, v in cells if isinstance(v, str)]
        return ("글자로 저장된 숫자", f"{name}({letter})", txt) if txt and len(txt) < len(cells) else None
    if kind == "날짜":
        txt = [r for r, v in cells if isinstance(v, str)]
        return ("글자 날짜", f"{name}({letter})", txt) if txt else None
    if kind != "글자":
        return None
    sym = [r for r, v in cells if isinstance(v, str) and _NUM_SYMBOL.match(v.strip())]
    if len(sym) >= max(2, len(cells) // 2):
        return ("기호 붙은 숫자", f"{name}({letter})", sym)
    other = [r for r, v in cells if not (_is_number(v) or _is_number_text(v))]
    if other and len(other) <= max(1, len(cells) // 10) and len(other) < len(cells):
        return ("숫자 사이 글자", f"{name}({letter})", other)
    return None


def column_read_warnings(notes: List[Tuple[str, str, List[int]]]) -> List[str]:
    """열마다 한 줄이면 같은 말이 여러 줄 되므로(공정 샘플 9열) 분류별로 한 줄로 묶는다."""
    out: List[str] = []
    for cat in ("글자로 저장된 숫자", "쉼표 숫자", "글자 날짜", "기호 붙은 숫자", "숫자 사이 글자"):
        group = [(label, rows) for c, label, rows in notes if c == cat]
        if not group:
            continue
        cols = ", ".join(label for label, _ in group[:8]) + (" 등" if len(group) > 8 else "")
        rows = sorted({r for _, rs in group for r in rs})
        one = len(group) == 1                                               # 한 열이면 몇 칸인지까지
        where = f"(엑셀 행 {_rows_text(rows)})"
        if cat == "글자로 저장된 숫자":
            n = f"{len(group[0][1])}칸" if one else "칸"
            out.append(f"{cols} 열의 글자로 저장된 숫자 {n}{where}을 숫자로 읽었습니다.")
        elif cat == "쉼표 숫자":
            out.append(f"{cols} 열은 천 단위 쉼표가 있는 숫자라 쉼표를 뺀 숫자로 읽었습니다{where}.")
        elif cat == "글자 날짜":
            out.append(f"{cols} 열은 글자로 적힌 날짜(2026.01.05 모양)라 날짜로 읽었습니다{where}.")
        elif cat == "기호 붙은 숫자":
            out.append(f"{cols} 열은 숫자에 기호가 붙어 있어 글자 열로 읽었습니다{where}. "
                       "숫자 조건(>, <, >=, <=)과 sort는 기호를 뺀 숫자로 비교하지만, group --sum으로 더할 수는 없습니다.")
        else:
            n, ga = (f"칸 {len(group[0][1])}개", "가") if one else ("칸", "이")
            out.append(f"{cols} 열은 숫자 사이에 숫자가 아닌 {n}{where}{ga} 있어 글자 열로 읽었습니다. "
                       "숫자 조건(>, <, >=, <=)과 sort는 숫자로 읽히는 칸끼리 비교합니다.")
    return out


def _two_row_header(grid: Dict[str, Any], header_row: int, header_vals: List[Any], data: List[Tuple[int, List[Any]]],
                    j1: int, j2: int) -> Tuple[int, List[Any], List[Tuple[int, List[Any]]], Optional[Dict[str, Any]]]:
    """머리글이 두 줄인 표(위는 병합한 묶음 이름, 아래가 진짜 이름)면 아래 줄을 머리글로 쓴다
    (집 PC 2026-11-14: 위 줄이 머리글이 되어 열 이름이 '기본', '열_B'가 되고 진짜 머리글이 데이터 첫 행이 됐음).
    조건: 머리글 줄에 두 칸 이상 병합이 있고, 바로 아래 줄이 글자로만 더 많이 차 있고, 그 아래로 데이터가 두 행 이상."""
    if len(data) < 3:
        return header_row, header_vals, data, None
    merged_here = [m for m in grid.get("merged", [])
                   if m[1] == header_row and m[2] > m[0] and m[0] <= grid["c1"] + j2 and m[2] >= grid["c1"] + j1]
    if not merged_here:
        return header_row, header_vals, data, None
    row2, vals2 = data[0]
    top = header_vals[j1:j2 + 1]
    if sum(1 for v in vals2 if not _empty(v)) <= sum(1 for v in top if not _empty(v)):
        return header_row, header_vals, data, None
    if any(not isinstance(v, str) or not v.strip() for v in vals2 if not _empty(v)):
        return header_row, header_vals, data, None
    merged_names = [str(v).strip() for v in top if not _empty(v)]
    names = [v if not _empty(v) else top[i] for i, v in enumerate(vals2)]
    return row2, [None] * j1 + names, data[1:], {"row": row2, "above_row": header_row, "above": merged_names}


def unit_row(data: List[Tuple[int, List[Any]]]) -> Optional[int]:
    """머리글 바로 아래 단위 행(`°C`, `bar`, `wt%` ...)의 자리. 그 행 때문에 숫자 열이 글자가 될 때만 (집 PC 2026-11-14 공정 샘플).
    조건: 데이터가 3행 이상, 첫 행의 찬 칸이 모두 짧은 글자(숫자 모양 아님)이고 둘 이상, 그 행을 빼면 두 열 이상이 숫자 열이 된다."""
    if len(data) < 3:
        return None
    first = data[0][1]
    filled = [j for j, v in enumerate(first) if not _empty(v)]
    if len(filled) < 2 or any(not isinstance(first[j], str) or len(first[j].strip()) > 12
                              or _is_number_text(first[j]) for j in filled):
        return None
    rest = data[1:]
    became = 0
    for j in filled:
        with_row = convert_column([vals[j] for _, vals in data])[2]
        without = convert_column([vals[j] for _, vals in rest])[2]
        if with_row == "글자" and without == "숫자":
            became += 1
    return data[0][0] if became >= 2 and became >= len(filled) // 2 else None


def convert_column(values: List[Any]) -> Tuple[List[Any], str, str]:
    """Values of one column -> (values, dtype, kind). Numbers only when the whole column is numbers (no leading 0);
    otherwise the cells are kept as they are (numbers written as text, a code like 0102 stays 0102)."""
    filled = [v for v in values if not _empty(v)]
    if not filled:
        return [None] * len(values), "object", "빈 열"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in filled):
        if all(isinstance(v, int) or float(v).is_integer() for v in filled) and all(isinstance(v, int) for v in filled):
            return [None if _empty(v) else int(v) for v in values], "Int64", "숫자"
        return [None if _empty(v) else float(v) for v in values], "float64", "숫자"
    if all(_is_number(v) or _is_number_text(v) or _is_comma_number(v) for v in filled) \
            and any(_is_comma_number(v) for v in filled):
        # 천 단위 쉼표가 있는 숫자 (글자로 저장된 숫자와 같은 종류: 숫자 값은 그대로, 쉼표만 뺌)
        nums = [None if _empty(v) else (v if _is_number(v) else float(str(v).replace(",", ""))) for v in values]
        if all(float(v).is_integer() for v in nums if v is not None):
            return [None if v is None else int(v) for v in nums], "Int64", "숫자"
        return [None if v is None else float(v) for v in nums], "float64", "숫자"
    if filled and all(_text_date(v) is not None for v in filled):
        return [None if _empty(v) else _text_date(v) for v in values], "datetime64[ns]", "날짜"
    if all(_is_number(v) or _is_number_text(v) for v in filled):
        # all numbers written as text (csv), or numbers mixed with a few numbers written as text (Excel's green triangle):
        # a number column (read_table warns about the mixed case). A code with a leading 0 (0102) is not a number.
        nums = [None if _empty(v) else (v if _is_number(v) else (float(v) if "." in v else int(v))) for v in values]
        if all(isinstance(v, int) for v in nums if v is not None):
            return nums, "Int64", "숫자"
        return [None if v is None else float(v) for v in nums], "float64", "숫자"
    if all(isinstance(v, (datetime, date)) for v in filled):
        return [None if _empty(v) else v for v in values], "datetime64[ns]", "날짜"

    def text(v: Any) -> Any:
        if _empty(v):
            return None
        if isinstance(v, bool):
            return str(v).upper()
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        if isinstance(v, datetime):
            return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.isoformat(sep=" ")
        if isinstance(v, date):
            return v.isoformat()
        return v if isinstance(v, str) else str(v)
    return [text(v) for v in values], "object", "글자"


def _part_span(rows: List[Tuple[int, List[Any]]], part: List[int]) -> set:
    return {j for i in part for j, v in enumerate(rows[i][1]) if not _empty(v)}


def _pick_block(rows: List[Tuple[int, List[Any]]], filled_n: List[int], explicit_end: bool) -> Tuple[Optional[int], List[int]]:
    """Header row index of the table and the Excel rows of other content above it that has two or more cells in a row.
    Content parts separated by fully empty rows; the table is the biggest part (columns of its header x rows), the first on a tie.
    In each part the header is the first row with two or more filled cells (a one-cell title above it is skipped).
    집 PC 2026-11-13 (MBOM): a title block (공장·Rev·작성일, 2 cells a row) above the real table was taken as the table."""
    parts: List[List[int]] = []
    for i, n in enumerate(filled_n):
        if n:
            if parts and parts[-1][-1] == i - 1:
                parts[-1].append(i)
            else:
                parts.append([i])
    if not parts:
        return None, []
    if explicit_end:                      # an explicit --range: the whole range is the table (empty rows inside are dropped)
        first = [i for i in range(len(filled_n)) if filled_n[i]]
        return next((i for i in first if filled_n[i] >= 2), first[0]), []
    best, best_score = None, -1
    for part in parts:
        head = next((i for i in part if filled_n[i] >= 2), part[0])
        score = (part[-1] - head + 1) * filled_n[head]
        if score > best_score:
            best, best_score = (part, head), score
    part, head = best
    above = [rows[i][0] for p in parts if p[-1] < part[0] for i in p if filled_n[i] >= 1]     # 한 칸짜리 제목 줄도 셈
    above += [rows[i][0] for i in part if i < head and filled_n[i] >= 1]
    _ask_when_split(rows, parts, part, head)
    return head, above


class _WholeTable(Exception):
    """표 가운데 빈 행: read_table이 빈 행까지 한 표(area)로 다시 읽는다."""

    def __init__(self, area: str, label: str):
        super().__init__(area)
        self.area, self.label = area, label


def _ask_when_split(rows: List[Tuple[int, List[Any]]], parts: List[List[int]], part: List[int], head: int) -> None:
    """표 가운데 빈 행이 있으면 표가 갈려 한 덩어리만 읽힌다 (집 PC 2026-11-14: 18·49행이 비어 BOM 111행 중 69행만 읽고
    50행을 머리글로 삼았음). 열이 거의 같은 다른 덩어리가 크면 묻지 않고 빈 행까지 한 표로 읽고 알린다
    (사용자 결정 2026-11-14: 표 가운데 빈 행은 표가 끝난 것이 아님. 전에는 [후보]로 물었음)."""
    mine, my_rows = _part_span(rows, part), part[-1] - head + 1

    def area(idx: List[int], cols: List[int]) -> str:
        return f"{col_letter(cols[0] + 1)}{rows[idx[0]][0]}:{col_letter(cols[-1] + 1)}{rows[idx[-1]][0]}"

    same = [p for p in parts if len(mine & _part_span(rows, p)) >= max(len(mine), len(_part_span(rows, p))) / 2]
    others = sum(len(p) for p in same if p is not part)
    if len(same) < 2 or others < max(3, my_rows / 3):
        return
    all_cols = sorted(set().union(*(_part_span(rows, p) for p in same)))
    whole = list(range(same[0][0], same[-1][-1] + 1))
    raise _WholeTable(area(whole, all_cols), f"{area(whole, all_cols)}, 빈 행 빼고 {sum(len(p) for p in same) - 1}행")


def read_table(path: str, sheet: Optional[str], range_text: Optional[str]) -> Tuple[Any, Dict[str, Any], List[str]]:
    """A table from a source file: header = first row with content, data until a fully empty row (or to the end of an
    explicit --range); a one-cell title row above it is skipped. Column names from the header (empty -> 열_F, repeated -> Qty_2), letters and Excel rows kept."""
    pd = pd_mod()
    grid = read_grid(path, sheet, parse_range(range_text))
    warnings = list(grid["warnings"])
    rows = grid["rows"]
    filled_n = [sum(not _empty(v) for v in vals) for _, vals in rows]
    try:
        head_i, above = _pick_block(rows, filled_n, grid["explicit_end"])
    except _WholeTable as whole:
        df, meta, more = read_table(path, sheet, whole.area)
        more = [w for w in more if "범위 안의 빈 행" not in w]               # 같은 일을 두 번 알리지 않음
        return df, meta, [f"표 안에 빈 행이 있어 빈 행까지 한 표로 읽었습니다 ({whole.label}). "
                          "일부만 읽으려면 --range로 범위를 주세요."] + more
    if above:
        span = str(above[0]) if len(above) == 1 else f"{above[0]}~{above[-1]}"
        warnings.append(f"표 위의 다른 내용 {len(above)}행(엑셀 행 {span})은 표가 아니라고 보고 읽지 않았습니다 "
                        "(빈 행으로 나뉜 내용 중 가장 큰 것을 표로 봄). 그것을 읽으려면 --range로 범위를 주세요.")
    if head_i is None:
        raise rt.InputError(f"{os.path.basename(path)} {grid['sheet']}"
                            + (f" {range_text}" if range_text else "") + "에 내용이 없습니다. --sheet나 --range를 확인하세요.")
    header_row, header_vals = rows[head_i]
    filled_cols = [j for j, v in enumerate(header_vals) if not _empty(v)]
    if grid["explicit_end"]:
        j1, j2 = 0, len(header_vals) - 1
    else:
        j1, j2 = filled_cols[0], filled_cols[-1]
    data: List[Tuple[int, List[Any]]] = []
    skipped_empty = 0
    after_gap = 0
    for r, vals in rows[head_i + 1:]:                       # 표의 끝은 머리글 줄의 칸 안에서 본다 (지금까지와 같게)
        if all(_empty(v) for v in vals[j1:j2 + 1]):
            if grid["explicit_end"]:
                skipped_empty += 1
                continue
            rest = [(rr, vv) for rr, vv in rows if rr > r and any(not _empty(x) for x in vv[j1:j2 + 1])]
            after_gap = len(rest)
            break
        data.append((r, list(vals)))
    if not grid["explicit_end"] and data:
        # 머리글 칸이 비어 있어도 값이 여러 행에 있으면 그 열도 표의 열이다 (집 PC 2026-11-14: 머리글이 두 줄인 표에서 G열이 통째로 빠졌음).
        need = max(2, len(data) // 10)
        extra = [j for j in range(len(header_vals)) if not (j1 <= j <= j2)
                 and sum(1 for _, v in data if j < len(v) and not _empty(v[j])) >= need]
        if extra:
            added = sorted(set(extra) | {j1, j2})
            j1, j2 = added[0], added[-1]
    data = [(r, vals[j1:j2 + 1]) for r, vals in data]
    header_row, header_vals, data, two_row = _two_row_header(grid, header_row, header_vals, data, j1, j2)
    units: Dict[str, str] = {}
    unit_at = unit_row(data)
    c0 = grid["c1"] + j1
    letters = [col_letter(c0 + j) for j in range(j2 - j1 + 1)]
    names: List[str] = []
    renamed: List[str] = []
    for j, v in enumerate(header_vals[j1:j2 + 1]):
        base = f"열_{letters[j]}" if _empty(v) else str(v).strip()
        if is_row_col(base):              # '엑셀 행' 같은 '... 행'은 도구의 행 번호 열 이름 (export한 표를 다시 읽을 때)
            renamed.append(f"'{base}' -> '{base}번호'")
            base = f"{base}번호"
        name, n = base, 2
        while name in names:
            name, n = f"{base}_{n}", n + 1
        names.append(name)
    if unit_at is not None:
        units = {names[j]: str(v).strip() for j, v in enumerate(data[0][1]) if not _empty(v)}
        data = data[1:]
        warnings.append(f"머리글 아래 {unit_at}행({', '.join(list(units.values())[:6])}"
                        + (" 등" if len(units) > 6 else "") + ")은 데이터가 아니라 열 단위로 보고 뺐습니다 "
                        "(그 행을 데이터로 읽으면 숫자 열이 글자 열이 됨). 단위는 결과의 columns에 있습니다.")
    cols: Dict[str, Any] = {ROW: pd.array([r for r, _ in data], dtype="Int64")}
    kinds: Dict[str, str] = {}
    notes: List[Tuple[str, str, List[int]]] = []
    for j, name in enumerate(names):
        vals, dtype, kind = convert_column([vals[j] for _, vals in data])
        kinds[name] = kind
        note = column_read_warning(name, letters[j], [(r, v[j]) for r, v in data if not _empty(v[j])], kind)
        if note:
            notes.append(note)
        cols[name] = pd.Series(vals, dtype=dtype if dtype != "datetime64[ns]" else None)
        if dtype == "datetime64[ns]":
            cols[name] = pd.to_datetime(cols[name])
    warnings.extend(column_read_warnings(notes))
    df = pd.DataFrame(cols)
    last = data[-1][0] if data else header_row
    area = f"{letters[0]}{header_row}:{letters[-1]}{last}"
    if two_row:
        warnings.append(f"머리글이 두 줄이라 아래 줄({two_row['row']}행)을 열 이름으로 썼습니다 "
                        f"(위 줄 {two_row['above_row']}행의 {', '.join(two_row['above'][:5])}"
                        + (" 등" if len(two_row["above"]) > 5 else "") + "은(는) 열 묶음 이름이라 쓰지 않음).")
    no_name = [letters[j] for j, v in enumerate(header_vals[j1:j2 + 1]) if _empty(v)]
    if no_name and not grid["explicit_end"]:
        warnings.append(f"머리글이 비어 있지만 값이 있는 열 {len(no_name)}개({', '.join(no_name)})도 읽었습니다 "
                        f"(이름은 열_{no_name[0]}처럼 붙임).")
    if renamed:
        warnings.append(f"도구가 행 번호에 쓰는 이름이라 열 이름을 바꿨습니다: {', '.join(renamed)}. "
                        "다른 표의 엑셀 행과 짝지으려면 join --on \"엑셀 행=엑셀 행번호\".")
    if after_gap:
        warnings.append(f"빈 행 아래에 내용이 있는 행이 {after_gap}개 더 있습니다 (읽지 않음). 함께 읽으려면 --range로 범위를 주세요.")
    if skipped_empty:
        warnings.append(f"범위 안의 빈 행 {skipped_empty}개는 뺐습니다.")
    data_rows = {r for r, _ in data}
    hid = len(data_rows & grid["hidden"])
    if hid:
        where = sorted(data_rows & grid["hidden"])
        out = argv_now() + ["--where", f"엑셀 행!={';엑셀 행!='.join(str(r) for r in where[:10])}"]
        warnings.append(f"숨긴 행(또는 자동 필터로 가려진 행) {hid}개(엑셀 행 {_rows_text(where)})도 읽었습니다. "
                        f"빼고 읽으려면: {rt.full_command(out)}")
    c_lo, c_hi = c0, c0 + len(letters) - 1
    r_lo, r_hi = header_row, last
    mg = sum(1 for (a, b, c, d) in grid["merged"] if a <= c_hi and c >= c_lo and b <= r_hi and d >= r_lo)
    if mg:
        warnings.append(f"병합한 칸 {mg}곳이 있습니다. 병합 칸의 빈자리는 채우지 않고 빈 값으로 읽었습니다.")
    by_row = dict(data)
    fe = sum(1 for (r, c) in grid["formulas"] if r in by_row and c_lo <= c <= c_hi and _empty(by_row[r][c - c0]))
    if fe:
        warnings.append(f"값이 저장되지 않은 수식 칸 {fe}개는 빈 값으로 읽었습니다 (Excel에서 열어 저장하면 값이 생김).")
    meta = {"units": units,
            "letters": {n: {"file": os.path.basename(path), "sheet": grid["sheet"], "letter": l} for n, l in zip(names, letters)},
            "rows": {ROW: {"file": os.path.basename(path), "path": path, "sheet": grid["sheet"]}},
            "kinds": kinds, "anchor": {"file": path, "sheet": grid["sheet"], "header_row": header_row, "range": area},
            "hidden": sorted(data_rows & grid["hidden"])}
    return df, meta, warnings


# ==============================================================================
# 4. Columns and conditions (--columns, --on, --where)
# ==============================================================================

def resolve_column(name: str, df: Any, meta: Dict[str, Any], option: str) -> str:
    """Column name, case/space-insensitive name, or column letter of the source file. A misspelt name -> [후보].
    --where takes the Excel row number column too (`엑셀 행`), so a row can be left out by its row number."""
    cols = data_cols(df)
    name = unquote(name)
    if option == "--where" and name == ROW_SHOWN and ROW in df.columns:
        return ROW
    if name in cols:
        return name
    key = lambda s: re.sub(r"\s+", "", str(s)).lower()
    same = [c for c in cols if key(c) == key(name)]
    if len(same) == 1:
        return same[0]
    if _LETTERS.match(name.upper()):
        hit = [c for c in cols if (meta.get("letters", {}).get(c) or {}).get("letter") == name.upper()]
        if len(hit) == 1:
            return hit[0]
    close = difflib.get_close_matches(name, cols, n=5, cutoff=0.5)
    if close:
        raise rt.ask(close, [argv_replacing(option, name, c) for c in close], f"{option}의 열 '{name}'이(가) 없어 쓸 열")
    raise rt.InputError(f"{option}의 열 '{name}'이(가) 표에 없습니다. 열: {name_list(cols)}")


def column_list(text: str, df: Any, meta: Dict[str, Any], option: str) -> List[str]:
    """'A,B' (names or letters, ',' between names). ';' between names -> [추천] the fixed command."""
    if ";" in text and "," not in text:
        fixed = ",".join(p.strip() for p in text.split(";") if p.strip())
        raise rt.InputError(f"{option}의 열 이름 사이는 ','로 나눕니다. [추천] {rt.full_command(argv_setting(option, fixed))}")
    names = [p.strip() for p in split_outside_quotes(text, ",") if p.strip()]
    if not names:
        raise rt.InputError(f"{option}가 비어 있습니다. 열 이름을 ','로 나눠 주세요.")
    out = []
    for n in names:
        c = resolve_column(n, df, meta, option)
        if c not in out:
            out.append(c)
    return out


def parse_where(text: str) -> List[List[Tuple[str, str, str]]]:
    """'Level=1;Qty>2' (; = and), 'Maker=A|B' (| = or; a value alone takes the column and operator before it).
    -> [[(col, op, value), ...or...], ...and...]. Conditions joined by ',' -> [추천] with ';'."""
    parts = [p for p in split_outside_quotes(str(text), ";") if p.strip()]
    if len(parts) == 1 and "," in parts[0]:
        pieces = [p for p in split_outside_quotes(parts[0], ",") if p.strip()]
        if len(pieces) > 1 and all(_COND.match(p) for p in pieces):
            raise rt.InputError("--where의 조건 사이는 ';'로 나눕니다 (','가 아님). "
                                f"[추천] {rt.full_command(argv_setting('--where', ';'.join(p.strip() for p in pieces)))}")
    groups = []
    for part in parts:
        alts: List[Tuple[str, str, str]] = []
        for alt in split_outside_quotes(part, "|"):
            m = _COND.match(alt)
            if m:
                col, op, val = m.group(1).strip(), m.group(2), unquote(m.group(3))
                if val == "" and op not in ("=", "!="):        # 'Qty>=' matched every row (집 PC 2026-11-13)
                    raise rt.InputError(f"--where '{alt.strip()}'에 비교할 값이 없습니다. 빈 칸은 '{col}='(빈 칸) 또는 "
                                        f"'{col}!='(빈 칸 아님)으로만 씁니다. 예: '{col}{op}2'.")
                alts.append((col, op, val))
            elif alts and alt.strip():
                alts.append((alts[-1][0], alts[-1][1], unquote(alt)))
            else:
                raise rt.InputError(f"--where '{text}'의 '{alt.strip()}'를 읽을 수 없습니다. 형식: '열=값;열>값' "
                                    f"(그리고는 ;, 또는은 |, 연산자 = != > < >= <= ~(포함) !~(포함 안 함) ^=(로 시작) $=(로 끝)).")
        groups.append(alts)
    if not groups:
        raise rt.InputError("--where가 비어 있습니다.")
    return groups


def _cond_mask(series: Any, op: str, val: str, notes: Optional[List[str]] = None, col: str = "") -> Any:
    pd = pd_mod()
    present = series.notna()
    if val == "" and op in ("=", "!="):           # 'col=' : empty cells
        empty = ~present | (series.astype("string") == "")
        return empty if op == "=" else ~empty
    text = series.astype("string")
    if op in ("~", "!~"):
        hit = text.str.contains(val, regex=False).fillna(False).astype(bool)
        return hit if op == "~" else ~hit
    if op in ("^=", "$="):                        # starts with / ends with (text as written, like ~)
        t = _as_text(series)
        hit = (t.str.startswith(val) if op == "^=" else t.str.endswith(val)).fillna(False).astype(bool)
        return hit
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        try:
            num = float(val.replace(",", ""))
        except ValueError:
            num = None
        if num is None and op in (">", "<", ">=", "<="):
            raise rt.InputError(f"'{col}' 열은 숫자 열이라 '{val}'과(와) 크기를 비교할 수 없습니다. 조건 값을 숫자로 쓰세요 "
                                f"(예: {col}{op}10). 글자로 비교하면 결과가 틀립니다.")
        if num is None and notes is not None:
            notes.append(f"'{col}' 열은 숫자 열인데 조건 값 '{val}'은(는) 숫자가 아니라 글자로 비교했습니다 (같은 칸이 없을 수 있음).")
        if num is not None:
            s = series.astype("Float64")
            res = {"=": s == num, "!=": s != num, ">": s > num, "<": s < num, ">=": s >= num, "<=": s <= num}[op]
            return res.fillna(op == "!=").astype(bool)
    if op in (">", "<", ">=", "<=") and _NUM_TEXT.match(val.replace(",", "")) and not pd.api.types.is_datetime64_any_dtype(series):
        # a text column with numbers: compare as numbers, not as text ("18" > "3" is False as text; 집 PC 2026-11-13)
        num = float(val.replace(",", ""))
        s = as_numbers(series)
        bad = int((present & s.isna()).sum())
        if bad and notes is not None:
            notes.append(f"'{col}' 열은 글자 열이라 숫자로 읽히는 칸끼리 비교했습니다 (숫자가 아닌 칸 {bad}개는 조건 '{col}{op}{val}'에 맞지 않음으로 봄).")
        res = {">": s > num, "<": s < num, ">=": s >= num, "<=": s <= num}[op]
        return res.fillna(False).astype(bool)
    if pd.api.types.is_datetime64_any_dtype(series):
        if not _DATE_TEXT.match(val.strip()):
            raise rt.InputError(f"'{col}' 열은 날짜 열이라 '{val}'을(를) 날짜로 읽을 수 없습니다. 날짜는 2026-12-31처럼 연-월-일로 쓰세요 "
                                "(글자로 비교하면 결과가 틀림).")
        try:
            when = pd.Timestamp(val)
        except (ValueError, TypeError):
            raise rt.InputError(f"'{col}' 열의 조건 값 '{val}'은(는) 없는 날짜입니다. 2026-12-31처럼 연-월-일로 쓰세요.") from None
        if when is not None:
            res = {"=": series == when, "!=": series != when, ">": series > when, "<": series < when,
                   ">=": series >= when, "<=": series <= when}[op]
            return res.fillna(op == "!=").astype(bool)
    res = {"=": text == val, "!=": text != val, ">": text > val, "<": text < val, ">=": text >= val, "<=": text <= val}[op]
    return res.fillna(op == "!=").astype(bool)


def _note_same_number(series: Any, mask: Any, col: str, val: str, notes: List[str]) -> None:
    """'열=125' on a text column compares the text as it is: cells written differently with the same number ('000125', '125.0')
    do not match. Said in warnings with the values there are (MBOM 2026-11-14: 'Part Number=125' was 0 rows without a word)."""
    pd = pd_mod()
    if pd.api.types.is_numeric_dtype(series) or not _NUM_TEXT.match(val.replace(",", "")):
        return
    text = series.astype("string").str.strip()
    same = (pd.to_numeric(text, errors="coerce") == float(val.replace(",", ""))).fillna(False).astype(bool) & ~mask
    if same.any():
        found = sorted(set(text[same].dropna().tolist()))[:5]
        notes.append(f"'{col}' 열은 글자 열이라 '{val}'을(를) 글자 그대로 비교했습니다. 숫자로 같은 값이 다르게 적힌 칸 {int(same.sum())}개는 "
                     f"고르지 않았습니다 (적힌 값: {', '.join(found)}). 그 칸도 고르려면 조건에 적힌 값 그대로 쓰세요 (예: {col}={found[0]}).")


def apply_where(df: Any, meta: Dict[str, Any], text: str) -> Tuple[Any, List[str]]:
    """Rows matching --where, and notes for warnings. 'A!=x|y' (every choice negative) means none of x, y
    (집 PC 2026-11-13: 'Item Type!=RAW|CONS' kept every row as 'not RAW or not CONS')."""
    groups = parse_where(text)
    notes: List[str] = []
    mask = None
    for alts in groups:
        none_of = len(alts) > 1 and all(op in ("!=", "!~") for _, op, _ in alts)
        part = None
        for col, op, val in alts:
            c = resolve_column(col, df, meta, "--where")
            m = _cond_mask(df[c], op, val, notes, c)
            if op == "=":
                _note_same_number(df[c], m, c, val, notes)
            part = m if part is None else ((part & m) if none_of else (part | m))
        mask = part if mask is None else (mask & part)
    return df[mask.values].reset_index(drop=True), notes


# ==============================================================================
# 5. Operations (every step is one of these; the same function runs again to rebuild a snapshot)
# ==============================================================================

def _keep_meta(meta: Dict[str, Any], cols: List[str]) -> Dict[str, Any]:
    out = dict(meta)
    out["letters"] = {c: v for c, v in meta.get("letters", {}).items() if c in cols}
    out["kinds"] = {c: v for c, v in meta.get("kinds", {}).items() if c in cols}
    out["units"] = {c: v for c, v in meta.get("units", {}).items() if c in cols}
    return out


def _side_rows(df: Any, meta: Dict[str, Any], table: str, taken: set) -> Tuple[Any, Dict[str, Any]]:
    """Row columns of one side of a two-table step: _row -> '<table> 행' (other '... 행' names kept, made unique)."""
    rename, rows = {}, {}
    for c in [c for c in df.columns if is_row_col(c)]:
        new = f"{table} 행" if c == ROW else c
        base, n = new, 2
        while new in taken:
            new, n = f"{base}_{n}", n + 1
        taken.add(new)
        rename[c] = new
        rows[new] = meta.get("rows", {}).get(c, {})
    return df.rename(columns=rename), rows


def op_select(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    warnings: List[str] = []
    if args.get("file"):
        df, meta, warnings = read_table(args["file"], args.get("sheet"), args.get("range"))
    else:
        df, meta = inputs["from"]
        df = df.copy()
    if args.get("where"):                  # conditions first: they may use columns that --columns leaves out
        df, notes = apply_where(df, meta, args["where"])
        warnings = warnings + notes
    if args.get("columns"):
        keep = column_list(args["columns"], df, meta, "--columns")
        df = df[[c for c in df.columns if is_row_col(c)] + keep]
        meta = _keep_meta(meta, keep)
    return df, meta, {"warnings": warnings}


def _key_col(name: str, side: Tuple[Any, Dict], option: str) -> str:
    """A key column of one table: a column name or letter, or '엑셀 행' (the table's source Excel row numbers)."""
    df, meta = side
    if unquote(name) == ROW_SHOWN and ROW in df.columns:
        return ROW
    return resolve_column(name, df, meta, option)


def _keys(text: str, left: Tuple[Any, Dict], right: Tuple[Any, Dict]) -> Tuple[List[str], List[str]]:
    """--on 'A,B' (same names in both tables) or 'A=X,B' (left column A = right column X). '엑셀 행' is the source row number,
    so columns a skill computed into a csv with an '엑셀 행번호' column join back to the table: --on "엑셀 행=엑셀 행번호"."""
    if ";" in text and "," not in text:
        fixed = ",".join(p.strip() for p in text.split(";") if p.strip())
        raise rt.InputError(f"--on의 열 이름 사이는 ','로 나눕니다. [추천] {rt.full_command(argv_setting('--on', fixed))}")
    items = [p.strip() for p in split_outside_quotes(text, ",") if p.strip()]
    if not items:
        raise rt.InputError("--on이 비어 있습니다. 키 열 이름을 ','로 나눠 주세요 (이름이 다르면 \"왼쪽 열=오른쪽 열\").")
    lkeys: List[str] = []
    rkeys: List[str] = []
    for i, item in enumerate(items):
        lname, rname = (p.strip() for p in item.split("=", 1)) if "=" in item else (item, item)
        lk = _key_col(lname, left, "--on")
        try:
            rk = _key_col(rname, right, "--on")
        except rt.InputError:
            rcols = data_cols(right[0])
            # 값이 많이 겹치는 열부터 (이름이 전혀 다른 열 Item도): 철자보다 값이 짝을 잘 말해 줌 (MBOM 2026-11-14: Item이 둘째였음)
            lvals = set(_key_frame(left[0], [lk])["k0"].tolist()) - {""}
            shared = {c: len(lvals & set(_key_frame(right[0], [c])["k0"].tolist())) for c in rcols}
            close = difflib.get_close_matches(rname, rcols, n=3, cutoff=0.3)
            close = sorted(rcols, key=lambda c: (-shared[c], close.index(c) if c in close else 99, rcols.index(c)))[:10]
            fixed = [",".join(items[:i] + [f"{lname}={c}"] + items[i + 1:]) for c in close]
            labels = [f"{c} (겹치는 값 {shared[c]}개)" for c in close]
            raise rt.ask(labels, [argv_setting("--on", f) for f in fixed],
                         f"--on의 '{rname}' 열이 오른쪽 표에 없어 짝지을 오른쪽 열 (이름이 다르면 \"왼쪽 열=오른쪽 열\")") from None
        lkeys.append(lk)
        rkeys.append(rk)
    return lkeys, rkeys


def _key_frame(df: Any, keys: List[str]) -> Any:
    """Keys as text for matching; a whole number is the same key whatever the column type (16, 16.0, '16')."""
    pd = pd_mod()
    out = {}
    for i, k in enumerate(keys):
        col = df[k]
        if pd.api.types.is_numeric_dtype(col) and not pd.api.types.is_bool_dtype(col):
            col = col.map(lambda v: None if v is None or v != v else (str(int(v)) if float(v).is_integer() else str(v)))
        out[f"k{i}"] = col.astype("string").fillna("").str.strip()
    return pd.DataFrame(out, index=df.index)


def _key_label(lk: str) -> str:
    return ROW_SHOWN if lk == ROW else lk


def op_join(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    pd = pd_mod()
    (ldf, lmeta), (rdf, rmeta) = inputs["left"], inputs["right"]
    lkeys, rkeys = _keys(args["on"], inputs["left"], inputs["right"])
    lk, rk = _key_frame(ldf, lkeys), _key_frame(rdf, rkeys)          # before the row columns are renamed
    keys = [k for k in rkeys if k != ROW]
    taken: set = set()
    # 왼쪽 키가 엑셀 행 하나면(스킬이 계산한 열을 원래 표에 붙이기) 결과는 원래 표 그대로라 엑셀 행을 그대로 둔다
    # (MBOM 2026-11-14: 'a 행'으로 바뀌어 그 표로 diff하면 엑셀 행과 칠하기 명령이 말없이 빠졌음)
    keep = lkeys == [ROW]
    if keep:
        taken.add(ROW)
        lrows = {ROW: lmeta.get("rows", {}).get(ROW, {})}
    else:
        ldf, lrows = _side_rows(ldf, lmeta, args["left_name"], taken)
    rdf, rrows = _side_rows(rdf, rmeta, args["right_name"], taken)
    rename = {c: f"{c}_{args['right_name']}" for c in data_cols(rdf) if c in data_cols(ldf) and c not in keys}
    rdf = rdf.rename(columns=rename)
    ldf2 = ldf.assign(**{f"__k{i}": lk[f"k{i}"].values for i in range(len(lkeys))})
    rdf2 = rdf.drop(columns=[k for k in keys if k in rdf.columns]).assign(**{f"__k{i}": rk[f"k{i}"].values for i in range(len(rkeys))})
    on = [f"__k{i}" for i in range(len(lkeys))]
    keep_left = bool(args.get("keep_left"))
    out = ldf2.merge(rdf2, on=on, how="left" if keep_left else "inner").drop(columns=on).reset_index(drop=True)
    dup = int(rk.duplicated().sum())
    warnings = [f"{args['right_name']}에 같은 키가 여러 번 나오는 행이 {dup}개 있어 결과 행이 늘었습니다."] if dup else []
    # 짝이 없어 빠진 행 (양쪽에 있는 행만 남기므로): 몇 행이 빠졌는지 알림 (MBOM 2026-11-14: 110행 중 96행이 말없이 빠짐)
    lset, rset = set(map(tuple, lk.itertuples(index=False))), set(map(tuple, rk.itertuples(index=False)))
    lmiss = sum(1 for t in lk.itertuples(index=False) if tuple(t) not in rset)
    rmiss = sum(1 for t in rk.itertuples(index=False) if tuple(t) not in lset)
    if keep_left and lmiss:
        warnings.append(f"{args['left_name']}의 {lmiss}행(전체 {len(lk)}행)은 {args['right_name']}에 짝이 없어 {args['right_name']}의 열이 빈 칸입니다.")
    if keep_left and rmiss:
        warnings.append(f"짝이 없어 빠진 {args['right_name']} 행: {rmiss}행(전체 {len(rk)}행).")
    if not keep_left and (lmiss or rmiss):
        warnings.append(f"짝이 없어 빠진 행: {args['left_name']} {lmiss}행(전체 {len(lk)}행), {args['right_name']} {rmiss}행(전체 {len(rk)}행). "
                        f"한쪽에만 있는 행은 only로 볼 수 있습니다."
                        + (f" {args['left_name']}의 행을 모두 남기려면 {rt.full_command(argv_now() + ['--keep-left'])}" if lmiss else ""))
    letters = {**lmeta.get("letters", {}), **{rename.get(c, c): v for c, v in rmeta.get("letters", {}).items() if c not in keys}}
    meta = {"letters": letters, "rows": {**lrows, **rrows}, "kinds": {}}
    return out, meta, {"warnings": warnings}


def op_only(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    (ldf, lmeta), (rdf, _rmeta) = inputs["left"], inputs["right"]
    lkeys, rkeys = _keys(args["on"], inputs["left"], inputs["right"])
    lk, rk = _key_frame(ldf, lkeys), _key_frame(rdf, rkeys)
    rset = set(map(tuple, rk.itertuples(index=False)))
    mask = [tuple(t) not in rset for t in lk.itertuples(index=False)]
    warnings = []
    ldup, rdup = int(lk.duplicated().sum()), int(rk.duplicated().sum())
    if ldup or rdup:     # 키가 행마다 다르지 않으면, 지운 한 행의 키가 다른 행에 남아 있을 때 '한쪽에만 있는 행'으로 나오지 않음
        warnings.append(f"--on({', '.join(_key_label(k) for k in lkeys)})이 같은 행이 {args['left_name']}에 {ldup}개, {args['right_name']}에 {rdup}개 "
                        "있습니다. 같은 키가 다른 행에 남아 있으면 그 행은 '한쪽에만 있는 행'으로 나오지 않습니다. "
                        "행마다 다른 키가 되게 열을 더 넣으세요.")
    return ldf[mask].reset_index(drop=True), dict(lmeta), {"warnings": warnings}


def op_stack(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    pd = pd_mod()
    frames, letters, rows = [], None, {}
    for i, tname in enumerate(args["tables"]):
        df, meta = inputs[f"t{i}"]
        frames.append(df.assign(**{"표": tname}))
        lt = meta.get("letters", {})
        letters = dict(lt) if letters is None else {c: v for c, v in letters.items() if lt.get(c) == v}
        for c, v in meta.get("rows", {}).items():
            rows.setdefault(c, v if not rows.get(c) or rows[c] == v else {"mixed": True})
    out = pd.concat(frames, ignore_index=True, sort=False)
    out = out[["표"] + [c for c in out.columns if c != "표"]]
    return out, {"letters": letters or {}, "rows": rows, "kinds": {}}, {"warnings": []}


def np_bool() -> type:
    import numpy
    return numpy.bool_


def _same(a: Any, b: Any) -> bool:
    pd = pd_mod()
    na_a, na_b = a is None or (not isinstance(a, str) and pd.isna(a)), b is None or (not isinstance(b, str) and pd.isna(b))
    if na_a or na_b:
        return na_a and na_b
    # numpy int64 (Int64 열)은 int가 아님: 한 판만 소수 수량이 있으면 2(Int64)와 2.0(float)이 달라 보였음 (집 PC 2026-11-13)
    if isinstance(a, numbers.Number) and isinstance(b, numbers.Number) and not isinstance(a, (bool, np_bool()))             and not isinstance(b, (bool, np_bool())):
        return float(a) == float(b)
    return str(cell_json(a)) == str(cell_json(b))


def _dup_examples(df: Any, kf: Any, n: int = 3) -> str:
    """'000125 (엑셀 행 23, 45)': up to n repeated keys with their source rows, so the user sees which rows collide."""
    rows = df[ROW].tolist() if ROW in df.columns else list(range(1, len(df) + 1))
    seen: Dict[tuple, List[Any]] = {}
    for t, r in zip(kf.itertuples(index=False), rows):
        seen.setdefault(tuple(t), []).append(r)
    out = [f"{'/'.join(k)} (엑셀 행 {', '.join(str(x) for x in v[:5])}{' 등' if len(v) > 5 else ''})"
           for k, v in seen.items() if len(v) > 1][:n]
    return "; ".join(out)


def op_diff(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """One row per change: 추가 (only in new), 삭제 (only in old), 바뀜 (one row per changed cell)."""
    pd = pd_mod()
    (odf, ometa), (ndf, nmeta) = inputs["old"], inputs["new"]
    okeys, nkeys = _keys(args["on"], inputs["old"], inputs["new"])
    keys = [_key_label(k) for k in okeys]
    for tname, df, ks in ((args["old_name"], odf, okeys), (args["new_name"], ndf, nkeys)):
        kf = _key_frame(df, ks)
        dup = int(kf.duplicated().sum())
        if dup:
            raise rt.InputError(f"{tname}에 --on({', '.join(_key_label(k) for k in ks)})이 같은 행이 {dup}개 있어 짝을 정할 수 없습니다 "
                                f"(예: {_dup_examples(df, kf)}). 열을 더 넣어 행마다 다르게 해 주세요.")
    on, nn = f"{args['old_name']} 행", f"{args['new_name']} 행"
    orow = odf[ROW] if ROW in odf.columns else pd.Series([None] * len(odf))
    nrow = ndf[ROW] if ROW in ndf.columns else pd.Series([None] * len(ndf))
    okey = {tuple(t): i for i, t in enumerate(_key_frame(odf, okeys).itertuples(index=False))}
    nkey = {tuple(t): i for i, t in enumerate(_key_frame(ndf, nkeys).itertuples(index=False))}
    common = [c for c in data_cols(odf) if c in data_cols(ndf) and c not in okeys and c not in nkeys]
    out: List[Dict[str, Any]] = []
    for k, i in okey.items():
        if k not in nkey:
            out.append({"구분": "삭제", **dict(zip(keys, k)), "열": None, "옛 값": None, "새 값": None, on: orow.iloc[i], nn: None})
    for k, j in nkey.items():
        if k not in okey:
            out.append({"구분": "추가", **dict(zip(keys, k)), "열": None, "옛 값": None, "새 값": None, on: None, nn: nrow.iloc[j]})
    for k, i in okey.items():
        j = nkey.get(k)
        if j is None:
            continue
        for c in common:
            a, b = odf[c].iloc[i], ndf[c].iloc[j]
            if not _same(a, b):
                out.append({"구분": "바뀜", **dict(zip(keys, k)), "열": c, "옛 값": cell_json(a), "새 값": cell_json(b),
                            on: orow.iloc[i], nn: nrow.iloc[j]})
    cols = ["구분"] + keys + ["열", "옛 값", "새 값", on, nn]
    df = pd.DataFrame(out, columns=cols)
    for c in (on, nn):
        df[c] = df[c].astype("Int64")
    for c in ("열", "옛 값", "새 값"):
        df[c] = df[c].astype(object)
    only_cols = sorted((set(data_cols(odf)) ^ set(data_cols(ndf))) - set(keys))
    warnings = [f"한쪽 표에만 있는 열은 비교하지 않았습니다: {name_list(only_cols)}"] if only_cols else []
    meta = {"letters": {}, "rows": {on: ometa.get("rows", {}).get(ROW, {}), nn: nmeta.get("rows", {}).get(ROW, {})},
            "kinds": {}, "diff": {"old_letters": ometa.get("letters", {}), "new_letters": nmeta.get("letters", {}),
                                  "old_rows": ometa.get("rows", {}).get(ROW, {}), "new_rows": nmeta.get("rows", {}).get(ROW, {})}}
    return df, meta, {"warnings": warnings}


def _numeric_cols_of(text: str, df: Any, meta: Dict[str, Any], option: str) -> List[str]:
    cols = column_list(text, df, meta, option)
    nums = _numeric_columns(df)
    bad = [c for c in cols if c not in nums]
    if bad:
        if not nums:
            raise rt.InputError(f"{option}에 쓸 숫자 열이 표에 없습니다 ('{bad[0]}'은(는) 글자 열).")
        cands = nums[:10]
        raise rt.ask(cands, [argv_setting(option, ",".join(n if c == bad[0] else c for c in cols)) for n in cands],
                     f"{option}의 '{bad[0]}'은(는) 숫자 열이 아니어서 대신 쓸 숫자 열")
    return cols


AGGS = (("sum", "합계"), ("mean", "평균"), ("min", "최소"), ("max", "최대"), ("count", "개수"), ("nunique", "종류 수"))
NUMBER_AGGS = ("sum", "mean", "min", "max")          # 숫자 열만 (count, nunique는 아무 열)


def _filled(series: Any) -> Any:
    """칸에 값이 있는지 (빈 칸과 공백만 있는 글자 칸은 빈 칸)."""
    return (series.notna() & (series.astype("string").str.strip() != "").fillna(False)).astype(bool)


def _as_text(series: Any) -> Any:
    """칸을 보이는 글자로: 정수인 숫자는 2.0이 아니라 2 (unique의 값, ^= $= 비교)."""
    pd = pd_mod()
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        num = pd.to_numeric(series, errors="coerce").astype("float64")
        return num.map(lambda v: None if v != v else (str(int(v)) if float(v).is_integer() else repr(float(v)))).astype("string")
    return series.astype("string").str.strip()


def _agg(series: Any, how: str) -> Any:
    if how == "sum":
        return series.sum(min_count=1)
    if how == "count":
        return int(_filled(series).sum())
    if how == "nunique":
        return int(_as_text(series[_filled(series)]).nunique())
    return getattr(series, how)()


def _agg_column(values: Any, source: Any, how: str) -> Any:
    """결과 열의 숫자 모양: 개수·종류 수는 정수, 평균은 소수 6자리, 합계·최소·최대는 모두 정수면 정수."""
    pd = pd_mod()
    if how in ("count", "nunique"):
        return values.astype("Int64")
    num = pd.to_numeric(values).astype("float64")
    if how == "mean":
        return num.round(6)
    num = num.round(9)
    whole = pd.api.types.is_integer_dtype(source) or bool(num.dropna().map(float.is_integer).all())
    return num.astype("Int64") if whole else num


def op_group(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """One row per value of --by: 행 수, <열> 합계/평균/최소/최대/개수/종류 수 for the columns of --sum/--mean/..., 엑셀 행 목록."""
    pd = pd_mod()
    df, meta = inputs["from"]
    by = args["by"]
    aggs = [(how, label, c) for how, label in AGGS for c in (args.get(how) or [])]
    sums = [c for how, _l, c in aggs if how in NUMBER_AGGS]
    warnings: List[str] = []
    work = df.assign(**{b: df[b].astype("string") for b in by})
    out_rows = []
    for key, g in work.groupby(by, sort=False, dropna=False):
        key = key if isinstance(key, tuple) else (key,)
        row: Dict[str, Any] = {b: (EMPTY_SHOWN if v is None or v is pd.NA or (isinstance(v, float) and v != v) else v)
                               for b, v in zip(by, key)}
        row["행 수"] = len(g)
        for how, label, c in aggs:
            row[f"{c} {label}"] = _agg(g[c], how)
        if ROW in g.columns:
            rows = [int(r) for r in g[ROW].dropna().tolist()]
            row["엑셀 행 목록"] = ", ".join(str(r) for r in rows[:10]) + (" 등" if len(rows) > 10 else "")
        out_rows.append(row)
    cols = by + ["행 수"] + [f"{c} {label}" for _h, label, c in aggs] + (["엑셀 행 목록"] if ROW in df.columns else [])
    out = pd.DataFrame(out_rows, columns=cols)
    out["행 수"] = out["행 수"].astype("Int64")
    for how, label, c in aggs:
        out[f"{c} {label}"] = _agg_column(out[f"{c} {label}"], df[c], how)
    unit = next((c for c in data_cols(df) if re.sub(r"[\s/._-]+", "", str(c)).lower() in UNIT_NAMES), None)
    if sums and unit and unit not in by:
        per = df.assign(**{unit: df[unit].astype("string").str.strip()}).groupby(
            [df[b].astype("string") for b in by], dropna=False)[unit].nunique()
        if (per > 1).any():                  # 집 PC 2026-11-13 (MBOM): EA·KG 수량을 한 합계에 더함
            warnings.append(f"'{unit}' 열의 단위가 다른 행을 한 합계에 더한 묶음이 {int((per > 1).sum())}개 있습니다. "
                            f"단위별로 나누려면: {rt.full_command(argv_setting('--by', ','.join(by + [unit])))}")
    letters = {b: v for b, v in meta.get("letters", {}).items() if b in by}
    return out, {"letters": letters, "rows": {}, "kinds": {}}, {"warnings": warnings}


def op_sort(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """Rows ordered by --by (first column first), --desc for large to small. Empty cells last, ties keep the table order.
    A text column whose cells are mostly numbers is ordered as numbers, like --where (집 PC 2026-11-14: 유량 5.02가 10.02보다
    크다고 나왔고 알리지도 않았음). Cells that are not numbers go last, with the rest of the empty cells."""
    pd = pd_mod()
    df, meta = inputs["from"]
    warnings: List[str] = []
    keys: Dict[str, Any] = {}
    for i, c in enumerate(args["by"]):
        col = df[c]
        if pd.api.types.is_numeric_dtype(col) or pd.api.types.is_datetime64_any_dtype(col) or pd.api.types.is_bool_dtype(col):
            continue
        num = as_numbers(col)
        filled = int(col.notna().sum())
        bad = int((col.notna() & num.isna()).sum())
        if filled and bad <= max(1, filled // 10) and bad < filled:
            keys[f"__s{i}"] = num
            warnings.append(f"'{c}' 열은 글자 열이라 숫자로 읽히는 칸끼리 숫자로 정렬했습니다"
                            + (f" (숫자가 아닌 칸 {bad}개는 맨 뒤)." if bad else "."))
    by = [f"__s{i}" if f"__s{i}" in keys else c for i, c in enumerate(args["by"])]
    work = df.assign(**keys) if keys else df
    out = work.sort_values(by, ascending=not args.get("desc"), na_position="last", kind="stable")
    out = out.drop(columns=list(keys)).reset_index(drop=True)
    return out, dict(meta), {"warnings": warnings}


def _row_list(rows: Any) -> str:
    """'엑셀 행 목록' 칸: 앞 10개 (group과 같은 모양)."""
    rows = [int(r) for r in rows]
    return ", ".join(str(r) for r in rows[:10]) + (" 등" if len(rows) > 10 else "")


def _calc_meta() -> Dict[str, Any]:
    """계산한 표(행이 원본 행과 1:1이 아님)의 meta."""
    return {"letters": {}, "rows": {}, "kinds": {}}


DESCRIBE_NUMBER = ("평균", "표준편차", "최소", "25%", "50%", "75%", "최대")
DESCRIBE_COLS = ["열", "종류", "개수", "빈 칸", *DESCRIBE_NUMBER, "종류 수", "가장 많은 값", "가장 많은 값 개수"]


def op_describe(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """열마다 한 행: 숫자 열은 개수·평균·표준편차·최소·25%·50%·75%·최대, 그 밖의 열은 개수·종류 수·가장 많은 값과 그 개수."""
    pd = pd_mod()
    df, meta = inputs["from"]
    cols = args.get("columns") or data_cols(df)
    nums = set(_numeric_columns(df))
    rows, warnings, like_number = [], [], []
    for c in cols:
        s = df[c]
        filled = _filled(s)
        n = int(filled.sum())
        row: Dict[str, Any] = {"열": c, "종류": kind_of(s), "개수": n, "빈 칸": len(s) - n}
        if c in nums:
            v = pd.to_numeric(s[filled], errors="coerce").astype("float64")
            if n:
                row.update({"평균": v.mean(), "표준편차": v.std() if n > 1 else None, "최소": v.min(), "25%": v.quantile(0.25),
                            "50%": v.quantile(0.5), "75%": v.quantile(0.75), "최대": v.max()})
        elif n:
            text = _as_text(s[filled])
            counts = text.value_counts(sort=False)
            top = max(counts.items(), key=lambda kv: kv[1])            # 같은 개수면 먼저 나온 값
            row.update({"종류 수": len(counts), "가장 많은 값": str(top[0]), "가장 많은 값 개수": int(top[1])})
            as_num = pd.to_numeric(text.str.replace(",", "", regex=False), errors="coerce")
            code = bool(text.str.match(_LEADING_ZERO.pattern).fillna(False).any())      # 앞자리 0이 있는 코드 열은 글자가 맞음
            if not code and not pd.api.types.is_datetime64_any_dtype(s) and int(as_num.notna().sum()) >= max(1, n * 9 // 10) and n >= 2:
                like_number.append(c)
        rows.append(row)
    out = pd.DataFrame(rows, columns=DESCRIBE_COLS)
    for c in ("개수", "빈 칸", "종류 수", "가장 많은 값 개수"):
        out[c] = pd.to_numeric(out[c]).astype("Int64")
    for c in DESCRIBE_NUMBER:
        out[c] = pd.to_numeric(out[c]).astype("float64").round(6)
    if like_number:
        warnings.append(f"글자 열인데 숫자로 읽히는 칸이 대부분인 열: {name_list(like_number, 10)}. 글자 열이라 평균·최소·최대는 내지 않았습니다 "
                        "(숫자가 아닌 칸이 섞였거나 글자로 저장된 숫자).")
    info = {"warnings": warnings, "numbers": len([c for c in cols if c in nums]), "others": len([c for c in cols if c not in nums])}
    return out, _calc_meta(), info


def op_unique(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """열마다 서로 다른 값과 그 개수 (많은 것부터, 같은 개수면 표에 먼저 나온 값부터). 빈 칸은 '(빈 칸)'."""
    pd = pd_mod()
    df, meta = inputs["from"]
    has_rows = ROW in df.columns
    rows, kinds = [], []
    for c in args["columns"]:
        text = _as_text(df[c]).where(_filled(df[c]), EMPTY_SHOWN)
        found: Dict[str, List[int]] = {}
        for i, v in enumerate(text.tolist()):
            found.setdefault(str(v), []).append(i)
        kinds.append((c, len(found)))
        for v, idx in sorted(found.items(), key=lambda kv: -len(kv[1])):       # sorted는 같은 개수의 순서를 지킴
            row = {"열": c, "값": v, "개수": len(idx)}
            if has_rows:
                row["엑셀 행 목록"] = _row_list(df[ROW].iloc[idx].dropna().tolist())
            rows.append(row)
    out = pd.DataFrame(rows, columns=["열", "값", "개수"] + (["엑셀 행 목록"] if has_rows else []))
    out["개수"] = pd.to_numeric(out["개수"]).astype("Int64")
    return out, _calc_meta(), {"warnings": [], "kinds": kinds}


# --- calc: 식을 직접 해석 (eval/exec를 쓰지 않음). 허용: 열 이름, [열 이름], 숫자, + - * /, 괄호 ---------------------------

CALC_FUNCS: Dict[str, Tuple[int, int]] = {}            # 함수 이름 -> (인자 수 최소, 최대)
_CALC_STOP = r"\s+\-*/()\[\],"                        # 감싸지 않은 열 이름에 쓸 수 없는 글자
_CALC_TOKEN = re.compile(rf"\s*(?:(\d+(?:\.\d+)?)(?![^{_CALC_STOP}])|\[([^\]]+)\]|([+\-*/(),])|([^{_CALC_STOP}]+))")
CALC_TOKENS_MAX = 200
CALC_HELP = "식에는 열 이름, 숫자, + - * /, 괄호만 씁니다. 공백이나 기호가 든 열 이름은 [열 이름]으로 감쌉니다 (예: 금액=Qty*[Unit Price])."


def _calc_tokens(text: str) -> List[Tuple[str, str]]:
    out, pos, text = [], 0, text.rstrip()
    while pos < len(text):
        m = _CALC_TOKEN.match(text, pos)
        if not m or m.end() == pos:
            raise rt.InputError(f"--expr의 식 '{text}'에서 '{text[pos:].strip()[:20]}'을(를) 읽을 수 없습니다. {CALC_HELP}")
        if m.group(1) is not None:
            out.append(("num", m.group(1)))
        elif m.group(2) is not None:
            out.append(("col", m.group(2).strip()))
        elif m.group(3) is not None:
            out.append(("op", m.group(3)))
        else:
            out.append(("name", m.group(4)))
        pos = m.end()
    if len(out) > CALC_TOKENS_MAX:
        raise rt.InputError(f"--expr의 식이 너무 깁니다 ({len(out)}개 조각, 최대 {CALC_TOKENS_MAX}). 식을 ';'로 나눠 열을 하나씩 만드세요.")
    return out


class _CalcParser:
    """식 -> 나무 (목록: ["num", 2.0], ["col", "Qty"], ["neg", x], ["bin", "+", l, r], ["call", "round", [x, ...]]).
    우선순위: 괄호·함수 > 부호 > * / > + -. 나무는 steps.json에 그대로 적어, 다시 만들 때 식을 다시 읽지 않는다."""

    def __init__(self, text: str) -> None:
        self.text, self.toks, self.i = text, _calc_tokens(text), 0

    def fail(self, why: str) -> rt.InputError:
        return rt.InputError(f"--expr의 식 '{self.text}'을(를) 읽을 수 없습니다 ({why}). {CALC_HELP}")

    def peek(self) -> Tuple[str, str]:
        return self.toks[self.i] if self.i < len(self.toks) else ("end", "")

    def take(self) -> Tuple[str, str]:
        tok = self.peek()
        self.i += 1
        return tok

    def parse(self) -> List[Any]:
        if not self.toks:
            raise self.fail("식이 비어 있음")
        node = self.expr()
        if self.peek()[0] != "end":
            kind, val = self.peek()
            raise self.fail(f"'{val}' 앞에 연산자가 없음" if kind in ("num", "col", "name") or val == "(" else f"'{val}'의 자리가 맞지 않음")
        return node

    def expr(self) -> List[Any]:
        node = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            node = ["bin", op, node, self.term()]
        return node

    def term(self) -> List[Any]:
        node = self.unary()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            node = ["bin", op, node, self.unary()]
        return node

    def unary(self) -> List[Any]:
        if self.peek() == ("op", "-"):
            self.take()
            return ["neg", self.unary()]
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        return self.atom()

    def atom(self) -> List[Any]:
        kind, val = self.take()
        if kind == "num":
            return ["num", float(val)]
        if kind == "col":
            return ["col", val]
        if kind == "name":
            if self.peek() == ("op", "("):
                fn = val.lower()
                if fn not in CALC_FUNCS:
                    raise self.fail(f"'{val}(' 함수는 쓸 수 없음" + (f", 쓸 수 있는 함수: {', '.join(sorted(CALC_FUNCS))}" if CALC_FUNCS else ""))
                self.take()
                args = [self.expr()]
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.expr())
                if self.take() != ("op", ")"):
                    raise self.fail("닫는 괄호 ')'가 없음")
                low, high = CALC_FUNCS[fn]
                if not low <= len(args) <= high:
                    raise self.fail(f"{fn}()에 넣는 값의 수가 맞지 않음")
                return ["call", fn, args]
            return ["col", val]
        if (kind, val) == ("op", "("):
            node = self.expr()
            if self.take() != ("op", ")"):
                raise self.fail("닫는 괄호 ')'가 없음")
            return node
        raise self.fail("식이 끝나지 않음" if kind == "end" else f"'{val}'의 자리가 맞지 않음")


def _calc_cols(node: List[Any]) -> List[str]:
    if node[0] == "col":
        return [node[1]]
    if node[0] == "neg":
        return _calc_cols(node[1])
    if node[0] == "bin":
        return _calc_cols(node[2]) + _calc_cols(node[3])
    if node[0] == "call":
        return [c for a in node[2] for c in _calc_cols(a)]
    return []


def _calc_rename(node: List[Any], names: Dict[str, str]) -> List[Any]:
    """식에 적은 열 이름을 표의 실제 열 이름으로 (대소문자·공백 차이, 열 글자)."""
    if node[0] == "col":
        return ["col", names[node[1]]]
    if node[0] == "neg":
        return ["neg", _calc_rename(node[1], names)]
    if node[0] == "bin":
        return ["bin", node[1], _calc_rename(node[2], names), _calc_rename(node[3], names)]
    if node[0] == "call":
        return ["call", node[1], [_calc_rename(a, names) for a in node[2]]]
    return list(node)


def _calc_number_column(series: Any, col: str, note: Dict[str, Any]) -> Any:
    """식에 쓰는 열을 숫자로. 글자 열은 숫자로 읽히는 칸만 (천 단위 쉼표는 뺌), 숫자가 아닌 칸은 빈 값으로 두고 note에 남김."""
    pd = pd_mod()
    if pd.api.types.is_bool_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        raise rt.InputError(f"--expr의 '{col}' 열은 숫자 열이 아니어서 계산할 수 없습니다 ({kind_of(series)} 열).")
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").astype("float64")
    filled = _filled(series)
    num = pd.to_numeric(series.astype("string").str.strip().str.replace(",", "", regex=False), errors="coerce").astype("float64")
    bad = filled & num.isna()
    if int(filled.sum()) and int(bad.sum()) == int(filled.sum()):
        raise rt.InputError(f"--expr의 '{col}' 열은 글자 열이고 숫자로 읽히는 칸이 없어 계산할 수 없습니다.")
    if bad.any():
        note["bad"] = note["bad"] | bad
        note["bad_cols"].append(col)
    return num


def _calc_value(node: List[Any], df: Any, note: Dict[str, Any]) -> Any:
    """나무 -> 행마다의 값 (float64 Series, 빈 값은 NaN)."""
    pd = pd_mod()
    kind = node[0]
    if kind == "num":
        return pd.Series(float(node[1]), index=df.index, dtype="float64")
    if kind == "col":
        return _calc_number_column(df[node[1]], node[1], note)
    if kind == "neg":
        return -_calc_value(node[1], df, note)
    if kind == "bin":
        left, right = _calc_value(node[2], df, note), _calc_value(node[3], df, note)
        if node[1] == "+":
            return left + right
        if node[1] == "-":
            return left - right
        if node[1] == "*":
            return left * right
        zero = (right == 0) & left.notna()
        note["zero"] = note["zero"] | zero
        return (left / right.where(right != 0)).where(~zero)
    if kind == "call":
        return _calc_call(node[1], [_calc_value(a, df, note) for a in node[2]], node)
    raise rt.InputError(f"작업 기록의 식을 읽을 수 없습니다 ({kind}).")


def _calc_call(fn: str, values: List[Any], node: List[Any]) -> Any:
    raise rt.InputError(f"'{fn}(' 함수는 쓸 수 없습니다.")


def op_calc(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """식으로 계산한 열을 표 끝에 더함 (행과 엑셀 행 번호는 그대로). 계산하지 못한 칸은 빈 값 + warnings."""
    pd = pd_mod()
    df, meta = inputs["from"]
    df = df.copy()
    warnings: List[str] = []
    for e in args["exprs"]:
        none = pd.Series(False, index=df.index)
        note: Dict[str, Any] = {"bad": none, "zero": none, "bad_cols": []}
        val = _calc_value(e["tree"], df, note).astype("float64").round(9)
        val = val.where(~val.isin([float("inf"), float("-inf")]))
        whole = bool(val.notna().any()) and bool(val.dropna().map(float.is_integer).all())
        df[e["name"]] = val.astype("Int64") if whole else val
        rows_of = (lambda mask: f" (엑셀 행 {_rows_text([int(r) for r in df.loc[mask, ROW].dropna().tolist()])})") if ROW in df.columns \
            else (lambda mask: "")
        if note["bad"].any():
            warnings.append(f"'{e['name']}' 열: {', '.join(dict.fromkeys(note['bad_cols']))} 열의 숫자가 아닌 칸 {int(note['bad'].sum())}행은 "
                            f"계산하지 못해 빈 값으로 두었습니다{rows_of(note['bad'])}.")
        if note["zero"].any():
            warnings.append(f"'{e['name']}' 열: 0으로 나눈 {int(note['zero'].sum())}행은 빈 값으로 두었습니다{rows_of(note['zero'])}.")
        empty = int((val.isna() & ~note["bad"] & ~note["zero"]).sum())
        if empty:
            warnings.append(f"'{e['name']}' 열: 식에 쓴 열이 빈 칸인 {empty}행은 빈 값입니다.")
    return df, dict(meta), {"warnings": warnings}


def op_top(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """--by 숫자 열이 큰 순(--asc면 작은 순)으로 앞 --n행. 그 열이 빈 칸인 행은 빼고, 같은 값은 표 순서대로."""
    df, meta = inputs["from"]
    by = args["by"]
    empty = df[by[0]].isna()
    out = df[~empty].sort_values(by, ascending=bool(args.get("asc")), na_position="last", kind="stable").head(int(args["n"]))
    warnings = [f"'{by[0]}' 열이 빈 칸인 {int(empty.sum())}행은 순위에서 뺐습니다."] if empty.any() else []
    return out.reset_index(drop=True), dict(meta), {"warnings": warnings}


def _rows_note(df: Any, mask: Any) -> str:
    """고른 행의 엑셀 행 번호 (lineage에 붙임). 엑셀 행이 없는 표(계산한 표)는 빈 글."""
    if ROW not in df.columns or not bool(mask.any()):
        return ""
    return f": 엑셀 행 {_rows_text([int(r) for r in df.loc[mask, ROW].dropna().tolist()], 10)}"


def op_dropna(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """--columns의 열 중 하나라도 빈 칸인 행을 뺌 (공백만 있는 글자 칸도 빈 칸)."""
    pd = pd_mod()
    df, meta = inputs["from"]
    empty = pd.Series(False, index=df.index)
    for c in args["columns"]:
        empty = empty | ~_filled(df[c])
    return df[~empty].reset_index(drop=True), dict(meta), {"warnings": [], "dropped": int(empty.sum()), "rows": _rows_note(df, empty)}


def _fill_value(series: Any, col: str, text: str) -> Any:
    """--value를 그 열의 종류에 맞는 값으로. 숫자 열에 글자, 날짜 열에 날짜가 아닌 값은 거절 (열의 종류를 바꾸지 않음)."""
    pd = pd_mod()
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        try:
            num = float(text.replace(",", ""))
        except ValueError:
            raise rt.InputError(f"'{col}' 열은 숫자 열이라 빈 칸을 '{text}'(으)로 채울 수 없습니다 (열이 글자 열로 바뀜). --value에 숫자를 주세요.") from None
        return int(num) if num.is_integer() and pd.api.types.is_integer_dtype(series) else num
    if pd.api.types.is_datetime64_any_dtype(series):
        if not _DATE_TEXT.match(text.strip()):
            raise rt.InputError(f"'{col}' 열은 날짜 열이라 '{text}'(으)로 채울 수 없습니다. --value에 2026-12-31처럼 연-월-일로 주세요.")
        try:
            return pd.Timestamp(text)
        except (ValueError, TypeError):
            raise rt.InputError(f"--value '{text}'은(는) 없는 날짜입니다.") from None
    return text


def op_fillna(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """--columns 열의 빈 칸을 --value로, 또는 --down이면 바로 위의 값으로 채움 (병합 칸을 푼 표, 위와 같은 값을 비워 둔 표).
    --down에서 맨 위부터 비어 있는 칸은 채울 값이 없어 그대로 두고 warnings."""
    pd = pd_mod()
    df, meta = inputs["from"]
    df = df.copy()
    warnings: List[str] = []
    filled_n: List[Tuple[str, int]] = []
    for c in args["columns"]:
        empty = ~_filled(df[c])
        if args.get("down"):
            new = df[c].where(~empty).ffill()
            left = empty & ~_filled(new)
            if left.any():
                rows = _rows_note(df, left)
                warnings.append(f"'{c}' 열: 위에 값이 없는 맨 위 빈 칸 {int(left.sum())}개는 그대로 두었습니다" + (f" ({rows[2:]})" if rows else "") + ".")
        else:
            new = df[c].where(~empty, _fill_value(df[c], c, str(args["value"])))
        if pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c]):
            new = pd.to_numeric(new)
            if pd.api.types.is_integer_dtype(df[c]) and bool(new.dropna().map(lambda v: float(v).is_integer()).all()):
                new = new.astype(df[c].dtype)
        filled_n.append((c, int((empty & _filled(new)).sum())))
        df[c] = new
    return df, dict(meta), {"warnings": warnings, "filled": filled_n}


def op_drop_duplicates(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """--columns(생략하면 모든 열)의 값이 모두 같은 행은 하나만 남김 (--keep first: 표에서 먼저 나온 행, last: 마지막 행).
    보이는 글자로 비교하고(2와 2.0은 같음), 빈 칸끼리는 같은 값으로 본다."""
    pd = pd_mod()
    df, meta = inputs["from"]
    cols = args.get("columns") or data_cols(df)
    key = pd.DataFrame({c: _as_text(df[c]).where(_filled(df[c]), EMPTY_SHOWN) for c in cols}, index=df.index)
    dup = key.duplicated(keep=args.get("keep") or "first")
    return df[~dup].reset_index(drop=True), dict(meta), {"warnings": [], "dropped": int(dup.sum()), "rows": _rows_note(df, dup)}


def op_rename(args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    """열 이름 바꾸기 (값과 행, 원본 열 글자는 그대로)."""
    df, meta = inputs["from"]
    names = dict(args["columns"])
    out = dict(meta)
    for part in ("letters", "kinds", "units"):
        out[part] = {names.get(c, c): v for c, v in (meta.get(part) or {}).items()}
    return df.rename(columns=names), out, {"warnings": []}


OP_FUNCS = {"select": op_select, "join": op_join, "only": op_only, "stack": op_stack, "diff": op_diff,
            "group": op_group, "sort": op_sort, "describe": op_describe, "unique": op_unique, "calc": op_calc, "top": op_top,
            "dropna": op_dropna, "fillna": op_fillna, "drop-duplicates": op_drop_duplicates, "rename": op_rename}


def run_op(op: str, args: Dict[str, Any], inputs: Dict[str, Tuple[Any, Dict[str, Any]]], name: str):
    return OP_FUNCS[op](args, inputs, name)


# ==============================================================================
# 6. Answers
# ==============================================================================

def columns_info(df: Any, meta: Dict[str, Any]) -> List[Dict[str, Any]]:
    out = []
    for c in data_cols(df)[:LIST_MAX]:
        lt = meta.get("letters", {}).get(c)
        item = {"name": c, "letter": lt["letter"] if lt else None, "kind": meta.get("kinds", {}).get(c) or kind_of(df[c])}
        if meta.get("units", {}).get(c):
            item["unit"] = meta["units"][c]
        out.append(item)
    return out


def where_text(meta: Dict[str, Any], df: Any) -> str:
    a = meta.get("anchor")
    if a and ROW in df.columns:
        return f"{os.path.basename(a['file'])} {a['sheet']} {a['range']} (머리글 {a['header_row']}행)"
    return "엑셀 위치 없음(계산한 표)" if not any(is_row_col(c) for c in df.columns) else "원본 행 번호는 '... 행' 열"


def show_cmd(name: str) -> Dict[str, Any]:
    return rt.next_command(["show", "--from", name], f"표 {name}의 앞부분 보기")


def table_answer(command: str, work: Work, name: str, df: Any, meta: Dict[str, Any], lineage: str, verified: bool,
                 warnings: List[str], extra_next: Optional[List[Dict[str, Any]]] = None, **fields: Any) -> Dict[str, Any]:
    n = len(df)
    warnings = list(warnings) + work.warnings
    preview = f"{lineage} -> {n}행 {len(data_cols(df))}열"
    then = (f"0건. 조건에 맞는 행이 없습니다(찾은 행 없음). 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요."
            if n == 0 else f"표 {name}을(를) 만들고 다시 읽어 행·열 수를 확인했습니다. preview를 알리고, 사용자가 내용을 보려 하면 next_commands[0]을 실행하세요.")
    out = rt.ok(command, preview, verified, then, None if verified else ["표 스냅샷을 다시 읽어 확인하지 못함"],
                table=name, count=n, columns=columns_info(df, meta), location=where_text(meta, df), lineage=lineage,
                work=work.folder, **fields)
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    nexts = ([show_cmd(name)] if n else []) + (extra_next or [])
    if nexts:
        out["next_commands"] = nexts
    return out


# ==============================================================================
# 7. Commands
# ==============================================================================

def check_name(name: Optional[str], work: Work) -> str:
    if not name:
        return work.next_name()
    if not _NAME.match(name):
        raise rt.InputError(f"--name '{name}'은(는) 쓸 수 없습니다. 글자로 시작하고 글자·숫자·_만 씁니다 (예: df1, 상위).")
    return name


def _select_needs_input(a: Any) -> rt.InputError:
    """--file도 --from도 없는 select: 무엇을 읽을지 도구가 짐작하지 않고 멈춤 (사용자 결정 2026-11-14).
    만든 표에서 고르는 명령과 Excel에 열린(저장된) 통합문서를 읽는 명령을 전체 명령으로 줌."""
    cands, argvs = [], []
    work = os.path.abspath(a.work) if getattr(a, "work", None) else current_work()
    if work and (getattr(a, "where", None) or getattr(a, "columns", None)):
        # 고를 조건·열이 있을 때만 '만든 표에서 고르기'를 권함: 아무것도 없는 select --from <표>는 표를 그대로 복사할 뿐이라
        # 경량 모델이 그것을 실행하고 "열린 엑셀을 읽었다"고 답할 수 있음 (집 PC 2026-11-13 4-2, 0.2.9).
        # 표가 여럿이면 어느 표인지 짐작하지 않고 모두 후보로 (집 PC 2026-11-14 P7: 마지막 표가 0행이었음, 0.2.10)
        try:
            w = Work(work)
            tables = sorted(w.names().items(), key=lambda kv: kv[1])[-TABLE_CHOICES:]
            for t, sid in tables:
                rows = w.step(sid).get("rows")
                cands.append(f"표 {t}" + (f"({rows}행)" if rows is not None else "") + "에서 고르기")
                argvs.append(argv_setting("--from", t))
        except Exception:
            pass
    books = rt.open_workbooks()
    for b in books:
        if b["saved"] and b["path"]:
            cands.append(f"Excel에 열린 {b['name']} 읽기")
            argvs.append(argv_setting("--file", b["path"]))
    unsaved = [b["name"] for b in books if not b["saved"]]
    note = f" (Excel에 저장하지 않은 문서 {', '.join(unsaved)}은(는) 저장해야 읽을 수 있음)" if unsaved else ""
    head = "select에는 --file(파일 읽기)이나 --from(만든 표에서 고르기) 중 하나가 필요합니다" + note + "."
    if not cands:
        return rt.InputError(head + " 사용자에게 읽을 파일 경로를 물어 --file \"<경로>\"로 주세요.")
    if len(cands) == 1:
        return rt.InputError(f"{head} [추천] {rt.full_command(argvs[0])}")
    err = rt.ask(cands, argvs, "읽을 대상")
    err.args = (head + " " + err.args[0],) + err.args[1:]
    return err


def cmd_select(a: Any) -> Dict[str, Any]:
    if not a.file and not a.source:
        raise _select_needs_input(a)
    if a.file:
        path = resolve_file(a.file)
        if a.source:
            raise rt.InputError("--file과 --from은 함께 쓰지 않습니다 (파일을 읽거나, 만든 표에서 고르거나).")
        work = open_work(a, create_for=path)
        args = {"file": path, "sheet": a.sheet, "range": a.range, "columns": a.columns, "where": a.where}
        df, meta, info = op_select(args, {}, "")
        name = check_name(a.name, work)
        st = os.stat(path)
        lineage = f"{name} = {where_text(meta, df)}" + (f"에서 {a.columns} 열" if a.columns else "") + (f", 조건 {a.where}" if a.where else "")
        sid, ok_ = work.add("select", name, args, {}, df, meta, lineage,
                            source={"file": path, "mtime": int(st.st_mtime), "size": st.st_size})
        kept = set(df[ROW].tolist()) if ROW in df.columns else set()
        hidden = [r for r in meta.get("hidden", []) if r in kept]       # 숨긴 행의 엑셀 행 (bom-kit이 결과에 든 숨긴 행을 알림)
        return table_answer("select", work, name, df, meta, lineage, ok_, info["warnings"],
                            anchor=meta.get("anchor"), source=path, **({"hidden_rows": hidden} if hidden else {}))
    if a.sheet or a.range:
        raise rt.InputError("--sheet와 --range는 --file과 함께만 씁니다 (만든 표에서 고를 때는 --columns, --where).")
    work = open_work(a)
    work.check_sources()
    src, sid_in = work.resolve(a.source, "--from")
    inp = work.load(sid_in)
    args = {"columns": a.columns, "where": a.where}
    df, meta, info = op_select(args, {"from": inp}, "")
    name = check_name(a.name, work)
    lineage = f"{name} = {src}에서" + (f" {a.columns} 열" if a.columns else "") + (f" 조건 {a.where}" if a.where else " 그대로")
    _sid, ok_ = work.add("select", name, args, {"from": sid_in}, df, meta, lineage)
    return table_answer("select", work, name, df, meta, lineage, ok_, info["warnings"])


def cmd_show(a: Any) -> Dict[str, Any]:
    pd = pd_mod()
    work = open_work(a)
    work.check_sources()
    name, sid = work.resolve(a.source, "--from")
    df, meta = work.load(sid)
    if a.columns:
        keep = column_list(a.columns, df, meta, "--columns")
        df = df[[c for c in df.columns if is_row_col(c)] + keep]
    n = len(df)
    start = max(1, int(a.start or 1))
    if n and start > n:
        raise rt.InputError(f"--start {start}이(가) 표의 행 수({n}행)보다 큽니다. [추천] {rt.full_command(argv_setting('--start', '1'))}")
    shown = df.iloc[start - 1:start - 1 + SHOW_MAX]
    more = n > start - 1 + SHOW_MAX
    rows = []
    for _, r in shown.iterrows():
        row = {}
        for c in df.columns:
            row[ROW_SHOWN if c == ROW else c] = cell_json(r[c])
        rows.append(row)
    fields: Dict[str, Any] = {"table": name, "count": n, "start": start, "shown": len(rows), "truncated": more or start > 1,
                              "columns": columns_info(df, meta), "location": where_text(meta, df), "rows": rows,
                              "lineage": work.step(sid)["lineage"], "work": work.folder}
    if more or start > 1:
        path = os.path.join(work.folder, "results", f"{name}_전체.csv")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        out = df.rename(columns={ROW: ROW_SHOWN})
        path = rt.save_with_lock_fallback(path, lambda p: out.to_csv(p, index=False, encoding="utf-8-sig"), "전체 결과")
        fields["full_result_file"] = path
    if n == 0:
        return rt.ok("show", f"표 {name}: 0행 (찾은 행 없음)", True,
                     "0건. 표에 행이 없습니다(찾은 행 없음). 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요.", **fields)
    part = f"{start}~{start - 1 + len(rows)}번째 행" if (more or start > 1) else ""
    preview = f"표 {name}: {n}행 {len(data_cols(df))}열, {fields['location']}" + (
        f". {part}만 보여 줌 (전체: full_result_file)" if part else "")
    out = rt.ok("show", preview, True, "preview와 rows(엑셀 행 번호 포함)를 사용자에게 알리세요. 같은 명령을 다시 실행하지 마세요."
                + (" 이어서 보려면 next_commands[0]." if more else ""), **fields)
    if more:
        out["next_commands"] = [rt.next_command(argv_setting("--start", str(start + SHOW_MAX)),
                                                f"표 {name}의 {start + SHOW_MAX}번째 행부터 보기")]
    if work.warnings:
        out["warnings"] = work.warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


def _two(a: Any, work: Work, first: str, second: str) -> Tuple[Tuple[str, int], Tuple[str, int]]:
    return work.resolve(getattr(a, first), f"--{first}"), work.resolve(getattr(a, second), f"--{second}")


def cmd_join(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    work.check_sources()
    (ln, ls), (rn, rs) = _two(a, work, "left", "right")
    args = {"on": a.on, "left_name": ln, "right_name": rn, "keep_left": bool(a.keep_left)}
    df, meta, info = op_join(args, {"left": work.load(ls), "right": work.load(rs)}, "")
    name = check_name(a.name, work)
    lineage = f"{name} = {ln}, {rn}을(를) {a.on} 기준으로 짝 맞춤 " + (f"({ln}의 행은 모두)" if a.keep_left else "(양쪽에 있는 행만)")
    _sid, ok_ = work.add("join", name, args, {"left": ls, "right": rs}, df, meta, lineage)
    return table_answer("join", work, name, df, meta, lineage, ok_, info["warnings"])


def cmd_only(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    work.check_sources()
    (ln, ls), (rn, rs) = _two(a, work, "left", "right")
    args = {"on": a.on, "left_name": ln, "right_name": rn}
    df, meta, info = op_only(args, {"left": work.load(ls), "right": work.load(rs)}, "")
    name = check_name(a.name, work)
    lineage = f"{name} = {ln}에만 있는 행 ({a.on} 기준으로 {rn}과(와) 비교)"
    _sid, ok_ = work.add("only", name, args, {"left": ls, "right": rs}, df, meta, lineage)
    return table_answer("only", work, name, df, meta, lineage, ok_, info["warnings"])


def cmd_stack(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    work.check_sources()
    if ";" in a.tables and "," not in a.tables:
        fixed = ",".join(p.strip() for p in a.tables.split(";") if p.strip())
        raise rt.InputError(f"--tables의 표 이름 사이는 ','로 나눕니다. [추천] {rt.full_command(argv_setting('--tables', fixed))}")
    names = [p.strip() for p in a.tables.split(",") if p.strip()]
    if len(names) < 2:
        raise rt.InputError("--tables에 이을 표를 둘 이상 ','로 주세요 (예: --tables \"df1,df2\").")
    resolved = [work.resolve(n, "--tables") for n in names]
    args = {"tables": [n for n, _ in resolved]}
    df, meta, info = op_stack(args, {f"t{i}": work.load(s) for i, (_, s) in enumerate(resolved)}, "")
    name = check_name(a.name, work)
    lineage = f"{name} = {', '.join(args['tables'])}를 위아래로 이음 ('표' 열에 어느 표의 행인지)"
    _sid, ok_ = work.add("stack", name, args, {f"t{i}": s for i, (_, s) in enumerate(resolved)}, df, meta, lineage)
    return table_answer("stack", work, name, df, meta, lineage, ok_, info["warnings"])


def cmd_group(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    work.check_sources()
    src, sid = work.resolve(a.source, "--from")
    df, meta = work.load(sid)
    args: Dict[str, Any] = {"by": column_list(a.by, df, meta, "--by")}
    for how, _label in AGGS:
        text = getattr(a, how, None)
        if how in NUMBER_AGGS:
            args[how] = _numeric_cols_of(text, df, meta, f"--{how}") if text else []
        elif text:
            args[how] = column_list(text, df, meta, f"--{how}")
    out, m, info = op_group(args, {"from": (df, meta)}, "")
    name = check_name(a.name, work)
    lineage = (f"{name} = {src}을(를) {','.join(args['by'])}별로 묶음 (행 수"
               + "".join(f", {c} {label}" for how, label in AGGS for c in (args.get(how) or [])) + ")")
    _sid, ok_ = work.add("group", name, args, {"from": sid}, out, m, lineage)
    return table_answer("group", work, name, out, m, lineage, ok_, info["warnings"])


def cmd_sort(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    work.check_sources()
    src, sid = work.resolve(a.source, "--from")
    df, meta = work.load(sid)
    args = {"by": column_list(a.by, df, meta, "--by"), "desc": bool(a.desc)}
    out, m, info = op_sort(args, {"from": (df, meta)}, "")
    name = check_name(a.name, work)
    lineage = f"{name} = {src}을(를) {','.join(args['by'])} " + ("큰 것부터" if a.desc else "작은 것부터") + " 정렬"
    _sid, ok_ = work.add("sort", name, args, {"from": sid}, out, m, lineage)
    return table_answer("sort", work, name, out, m, lineage, ok_, info["warnings"])


def _one_table(a: Any, op: str, make_args: Any, lineage_of: Any) -> Dict[str, Any]:
    """표 하나(--from)로 새 표(--name)를 만드는 명령의 공통 순서: 표 찾기 -> 인자 확인 -> 계산 -> 기록·스냅샷 -> 답.
    make_args(df, meta, src) -> 기록에 남길 args, lineage_of(name, src, args, info) -> 한 줄 설명."""
    work = open_work(a)
    work.check_sources()
    src, sid = work.resolve(a.source, "--from")
    df, meta = work.load(sid)
    args = make_args(df, meta, src)
    out, m, info = run_op(op, args, {"from": (df, meta)}, "")
    name = check_name(a.name, work)
    lineage = lineage_of(name, src, args, info)
    _sid, ok_ = work.add(op, name, args, {"from": sid}, out, m, lineage)
    return table_answer(op, work, name, out, m, lineage, ok_, info["warnings"])


def cmd_describe(a: Any) -> Dict[str, Any]:
    return _one_table(
        a, "describe",
        lambda df, meta, src: {"columns": column_list(a.columns, df, meta, "--columns") if a.columns else None},
        lambda name, src, args, info: f"{name} = {src}의 열 {info['numbers'] + info['others']}개 요약 "
                                      f"(숫자 열 {info['numbers']}개, 그 밖의 열 {info['others']}개, 열마다 한 행)")


def cmd_unique(a: Any) -> Dict[str, Any]:
    return _one_table(
        a, "unique",
        lambda df, meta, src: {"columns": column_list(a.columns, df, meta, "--columns")},
        lambda name, src, args, info: f"{name} = {src}의 서로 다른 값 (" + ", ".join(f"{c} {n}가지" for c, n in info["kinds"]) + ")")


def _expr_parts(text: str) -> List[str]:
    """'금액=Qty*단가;합=A+B' -> 식들. ','로 나눠 적었으면 ';'로 고친 전체 명령을 [추천] (round(x, 2) 같은 괄호 안의 ','는 그대로)."""
    parts = [p.strip() for p in str(text).split(";") if p.strip()]
    if len(parts) == 1 and "," in parts[0]:
        pieces, cur, depth = [], "", 0
        for ch in parts[0]:
            depth += ch in "(["
            depth -= ch in ")]"
            if ch == "," and depth == 0:
                pieces.append(cur)
                cur = ""
            else:
                cur += ch
        pieces = [p.strip() for p in pieces + [cur] if p.strip()]
        if len(pieces) > 1 and all("=" in p for p in pieces):
            raise rt.InputError(f"--expr의 식 사이는 ';'로 나눕니다 (','가 아님). [추천] {rt.full_command(argv_setting('--expr', ';'.join(pieces)))}")
    if not parts:
        raise rt.InputError("--expr가 비어 있습니다. 형식: \"새 열 이름=식\" (예: \"금액=Qty*단가\").")
    return parts


def _bracket_fix(text: str, expr: str, cols: List[str]) -> Optional[str]:
    """공백·기호가 든 열 이름을 감싸지 않고 적은 식 -> [열 이름]으로 감싼 식 (그렇게 읽히면). 아니면 None."""
    fixed = expr
    for c in sorted((c for c in cols if re.search(rf"[{_CALC_STOP}]", str(c))), key=lambda c: -len(str(c))):
        fixed = re.sub(rf"(?<!\[){re.escape(str(c))}(?!\])", lambda m: f"[{c}]", fixed)
    if fixed == expr:
        return None
    try:
        tree = _CalcParser(fixed).parse()
    except rt.InputError:
        return None
    return text.replace(expr, fixed, 1) if all(n in cols for n in _calc_cols(tree)) else None


def _calc_column(n: str, df: Any, meta: Dict[str, Any], text: str) -> str:
    """식의 열 이름 -> 표의 열. 없으면 [후보] (고친 식의 열 이름은 필요하면 [열 이름]으로 감쌈)."""
    cols = data_cols(df)
    try:
        return resolve_column(n, df, meta, "--expr")
    except rt.InputError:
        close = difflib.get_close_matches(n, cols, n=5, cutoff=0.5)
        if not close:
            raise
        old = f"[{n}]" if f"[{n}]" in text else n
        wrap = lambda c: f"[{c}]" if re.search(rf"[{_CALC_STOP}]", str(c)) else str(c)
        raise rt.ask(close, [argv_replacing("--expr", old, wrap(c)) for c in close], f"--expr의 열 '{n}'이(가) 없어 쓸 열") from None


def _calc_args(a: Any, df: Any, meta: Dict[str, Any]) -> Dict[str, Any]:
    exprs: List[Dict[str, Any]] = []
    cols = data_cols(df)
    made: List[str] = []
    for part in _expr_parts(a.expr):
        new, eq, expr = part.partition("=")
        new, expr = new.strip().strip("[]").strip(), expr.strip()
        if not eq or not new or not expr or expr.startswith("="):
            raise rt.InputError(f"--expr '{part}'은(는) \"새 열 이름=식\" 모양이 아닙니다 (예: \"금액=Qty*단가\", 여러 개는 ';'로 나눔).")
        if new in cols or new in made or is_row_col(new):
            raise rt.InputError(f"--expr의 새 열 이름 '{new}'은(는) 이미 있는 열이거나 쓸 수 없는 이름입니다 (있는 열의 값은 바꾸지 않음). "
                                "다른 이름을 쓰세요 (예: " + f"{new}2={expr}).")
        try:
            tree = _CalcParser(expr).parse()
            names = {}
            for n in dict.fromkeys(_calc_cols(tree)):
                names[n] = n if n in made else _calc_column(n, df, meta, a.expr)
        except rt.InputError as e:
            fixed = _bracket_fix(a.expr, expr, cols)
            if fixed:
                raise rt.InputError("--expr에서 공백이나 기호가 든 열 이름은 [열 이름]으로 감쌉니다. "
                                    f"[추천] {rt.full_command(argv_setting('--expr', fixed))}") from None
            raise e
        exprs.append({"name": new, "text": expr, "tree": _calc_rename(tree, names)})
        made.append(new)
    return {"exprs": exprs}


def cmd_calc(a: Any) -> Dict[str, Any]:
    return _one_table(
        a, "calc", lambda df, meta, src: _calc_args(a, df, meta),
        lambda name, src, args, info: f"{name} = {src}에 계산한 열을 더함 (" + ", ".join(f"{e['name']}={e['text']}" for e in args["exprs"]) + ")")


def cmd_top(a: Any) -> Dict[str, Any]:
    if a.n < 1:
        raise rt.InputError(f"--n {a.n}은(는) 쓸 수 없습니다. 1 이상의 행 수를 주세요. [추천] {rt.full_command(argv_setting('--n', '10'))}")
    return _one_table(
        a, "top",
        lambda df, meta, src: {"by": _numeric_cols_of(a.by, df, meta, "--by"), "n": int(a.n), "asc": bool(a.asc)},
        lambda name, src, args, info: f"{name} = {src}에서 {','.join(args['by'])} " + ("작은" if args["asc"] else "큰") + f" 순 앞 {args['n']}행")


def cmd_dropna(a: Any) -> Dict[str, Any]:
    return _one_table(
        a, "dropna", lambda df, meta, src: {"columns": column_list(a.columns, df, meta, "--columns")},
        lambda name, src, args, info: f"{name} = {src}에서 {','.join(args['columns'])}이(가) 빈 칸인 행 {info['dropped']}개를 뺌{info['rows']}")


def cmd_fillna(a: Any) -> Dict[str, Any]:
    if a.down and a.value is not None:
        raise rt.InputError("--value(이 값으로 채움)와 --down(위 값으로 채움)은 함께 쓰지 않습니다. 하나만 주세요.")
    if not a.down and a.value is None:
        raise rt.InputError("빈 칸을 무엇으로 채울지 정해 주세요: --value \"<값>\"(이 값으로) 또는 --down(바로 위 칸의 값으로, 병합 칸 정리).")

    def make(df: Any, meta: Dict[str, Any], src: str) -> Dict[str, Any]:
        cols = column_list(a.columns, df, meta, "--columns")
        if not a.down:
            for c in cols:
                _fill_value(df[c], c, str(a.value))                # 열 종류에 맞지 않는 값은 여기서 멈춤
        return {"columns": cols, "down": bool(a.down), "value": None if a.down else str(a.value)}
    return _one_table(
        a, "fillna", make,
        lambda name, src, args, info: f"{name} = {src}의 빈 칸을 " + ("위 값으로" if args["down"] else f"'{args['value']}'(으)로")
                                      + " 채움 (" + ", ".join(f"{c} {n}칸" for c, n in info["filled"]) + ")")


def cmd_drop_duplicates(a: Any) -> Dict[str, Any]:
    return _one_table(
        a, "drop-duplicates",
        lambda df, meta, src: {"columns": column_list(a.columns, df, meta, "--columns") if a.columns else None, "keep": a.keep},
        lambda name, src, args, info: f"{name} = {src}에서 {','.join(args['columns']) if args['columns'] else '모든 열'}이(가) 같은 행은 "
                                      + ("먼저 나온" if args["keep"] == "first" else "마지막") + f" 행만 남김 (뺀 행 {info['dropped']}개{info['rows']})")


def _rename_pairs(a: Any, df: Any, meta: Dict[str, Any]) -> List[List[str]]:
    text = str(a.columns)
    if ";" in text and "," not in text:
        fixed = ",".join(p.strip() for p in text.split(";") if p.strip())
        raise rt.InputError(f"--columns의 이름 사이는 ','로 나눕니다. [추천] {rt.full_command(argv_setting('--columns', fixed))}")
    pairs: List[List[str]] = []
    cols = data_cols(df)
    for part in [p.strip() for p in split_outside_quotes(text, ",") if p.strip()]:
        old, eq, new = part.partition("=")
        new = unquote(new)
        if not eq or not old.strip() or not new:
            raise rt.InputError(f"--columns '{part}'은(는) \"지금 이름=새 이름\" 모양이 아닙니다 (예: --columns \"Qty=수량,UoM=단위\").")
        c = resolve_column(old.strip(), df, meta, "--columns")
        if c in [p[0] for p in pairs]:
            raise rt.InputError(f"--columns에 '{c}' 열이 두 번 나옵니다. 열마다 새 이름을 하나만 주세요.")
        if c != new:
            pairs.append([c, new])
    stay = [c for c in cols if c not in [p[0] for p in pairs]]        # 이름이 그대로인 열 (서로 바꾸기 Qty=Maker,Maker=Qty는 됨)
    news = [p[1] for p in pairs]
    for new in news:
        if is_row_col(new) or new in stay or news.count(new) > 1:
            raise rt.InputError(f"--columns의 새 이름 '{new}'은(는) 이미 있는 열이거나 쓸 수 없는 이름입니다 ('엑셀 행', '... 행'은 도구가 씀). 다른 이름을 주세요.")
    if not pairs:
        raise rt.InputError("--columns에 바꿀 이름이 없습니다 (예: --columns \"Qty=수량,UoM=단위\").")
    return pairs


def cmd_rename(a: Any) -> Dict[str, Any]:
    return _one_table(
        a, "rename", lambda df, meta, src: {"columns": _rename_pairs(a, df, meta)},
        lambda name, src, args, info: f"{name} = {src}의 열 이름을 바꿈 (" + ", ".join(f"{o} -> {n}" for o, n in args["columns"]) + ")")


OFFICE_EXCEL = os.path.join(os.path.dirname(rt.TOOL_DIR), "office-cli", "excel.py")


def _paint_files(work: Work, name: str, df: Any, meta: Dict[str, Any]) -> Tuple[List[str], List[Dict[str, Any]]]:
    """Cells to color per source file (old: 삭제 rows and changed cells, new: 추가 rows and changed cells) as csv files
    (one cell or row range per line, office-cli excel style --file) and the office-cli commands (confirm_commands)."""
    d = meta.get("diff", {})
    files, cmds = [], []
    for side, kinds, col in (("old", ("삭제", "바뀜"), list(meta["rows"])[0]), ("new", ("추가", "바뀜"), list(meta["rows"])[1])):
        rows_meta, letters = d.get(f"{side}_rows") or {}, d.get(f"{side}_letters") or {}
        if not rows_meta.get("path") or not letters:
            continue
        span = sorted((v["letter"] for v in letters.values()), key=col_index)
        cells, nrow, ncell = [], 0, 0
        for _, r in df[df["구분"].isin(kinds)].iterrows():
            xr = r[col]
            if xr is None or (not isinstance(xr, int) and pd_mod().isna(xr)):
                continue
            if r["구분"] == "바뀜":
                lt = letters.get(r["열"])
                if lt:
                    cells.append(f"{lt['letter']}{int(xr)}")
                    ncell += 1
            else:
                cells.append(f"{span[0]}{int(xr)}:{span[-1]}{int(xr)}")
                nrow += 1
        if not cells:
            continue
        base = os.path.splitext(os.path.basename(rows_meta["path"]))[0]
        path = os.path.join(work.folder, "results", f"{name}_칠할칸_{side}_{base}.csv")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        text = "\n".join(cells) + "\n"
        path = rt.save_with_lock_fallback(path, lambda p: open(p, "w", encoding="utf-8-sig").write(text), "칠할 칸 목록")
        files.append(path)
        cmds.append(rt.next_command(["style", "--target", rows_meta["path"], "--sheet", rows_meta.get("sheet") or "",
                                     "--file", path, "--bg", "yellow"],
                                    # 행과 칸을 나눠 셈 (MBOM 2026-11-14: 삭제 3행 + 바뀐 칸 8곳을 '바뀐 칸 11곳'이라 적었음)
                                    f"원본 {os.path.basename(rows_meta['path'])}의 "
                                    + ", ".join(t for t in (f"{'삭제' if side == 'old' else '추가'}된 {nrow}행" if nrow else "",
                                                            f"바뀐 칸 {ncell}곳" if ncell else "") if t)
                                    + "을 노랗게 칠하기 (사용자가 원할 때만, office-cli)",
                                    script=OFFICE_EXCEL))
    return files, cmds


def cmd_diff(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    work.check_sources()
    (on_, os_), (nn_, ns_) = _two(a, work, "old", "new")
    args = {"on": a.on, "old_name": on_, "new_name": nn_}
    df, meta, info = op_diff(args, {"old": work.load(os_), "new": work.load(ns_)}, "")
    name = check_name(a.name, work)
    counts = df["구분"].value_counts().to_dict()
    summary = f"추가 {counts.get('추가', 0)}행, 삭제 {counts.get('삭제', 0)}행, 바뀐 칸 {counts.get('바뀜', 0)}곳"
    lineage = f"{name} = {on_}(옛) -> {nn_}(새) 비교 ({a.on} 기준): {summary}"
    _sid, ok_ = work.add("diff", name, args, {"old": os_, "new": ns_}, df, meta, lineage)
    files, cmds = _paint_files(work, name, df, meta)
    out = table_answer("diff", work, name, df, meta, lineage, ok_, info["warnings"], summary=summary)
    if len(df) == 0:
        out["agent_directive"] = "VERIFIED: 0건. 두 표에서 다른 곳이 없습니다(찾은 차이 없음). 사용자에게 그대로 알리고 다시 실행하지 마세요."
    if files:
        out["paint_files"] = files
        out["confirm_commands"] = cmds
        out["agent_directive"] += " 원본에 칠하기는 사용자가 원할 때만 confirm_commands를 1번 실행하세요."
    return out


def cmd_undo(a: Any) -> Dict[str, Any]:
    work = open_work(a)
    if not work.steps:
        raise rt.InputError("되돌릴 단계가 없습니다 (이 작업 폴더에 만든 표가 없음).")
    with work.locked():
        if not work.steps:
            raise rt.InputError("되돌릴 단계가 없습니다 (이 작업 폴더에 만든 표가 없음).")
        step = work.steps.pop()
        work.save()
        try:
            os.remove(work.table_file(step["id"]))
        except OSError:
            pass
    still = work.names()
    back = still.get(step["name"])
    preview = (f"되돌림: {step['lineage']}. " + (f"{step['name']}은(는) 그 전 표를 다시 가리킵니다. " if back else f"{step['name']}은(는) 없어졌습니다. ")
               + (f"남은 표: {', '.join(still)}" if still else "남은 표 없음"))   # 다음 명령은 --from으로 표를 적음
    verified = step["id"] not in [s["id"] for s in Work(work.folder).steps]
    return rt.ok("undo", preview, verified, "preview를 사용자에게 알리세요. 같은 명령을 다시 실행하지 마세요 (한 번에 한 단계씩 되돌림).",
                 None if verified else ["작업 기록을 다시 읽어 확인하지 못함"], removed=step["name"],
                 tables=sorted(still), work=work.folder)


def cmd_new(a: Any) -> Dict[str, Any]:
    old = current_work()
    set_current_work(None)
    verified = current_work() is None
    preview = "새 작업을 시작합니다. 다음 select --file이 그 파일 옆에 새 작업 폴더를 만듭니다." + (f" (지난 작업 폴더: {old})" if old else "")
    return rt.ok("new", preview, verified, "preview를 사용자에게 알리세요.", None if verified else ["지금 작업 폴더 표시를 지우지 못함"],
                 previous_work=old)


def _export_path(work: Work, name: str, save: Optional[str]) -> str:
    if not save:
        return os.path.join(work.folder, f"{name}.xlsx")
    p = str(save).strip().strip('"')
    if not os.path.splitext(p)[1]:
        p += ".xlsx"
    if os.path.splitext(p)[1].lower() not in (".xlsx", ".csv"):
        raise rt.InputError(f"--save는 .xlsx나 .csv입니다. [추천] {rt.full_command(argv_setting('--save', os.path.splitext(p)[0] + '.xlsx'))}")
    return p if os.path.isabs(p) else os.path.join(work.folder, p)


def cmd_export(a: Any) -> Dict[str, Any]:
    pd = pd_mod()
    work = open_work(a)
    work.check_sources()
    name, sid = work.resolve(a.source, "--from")
    df, meta = work.load(sid)
    path = os.path.abspath(_export_path(work, name, a.save))
    sources = {os.path.normcase(os.path.abspath(s["source"]["file"])) for s in work.steps if s.get("source")}
    if os.path.normcase(path) in sources:
        raise rt.InputError(f"원본 파일({os.path.basename(path)})에는 쓰지 않습니다. 다른 이름을 --save로 주세요.")
    out = df.rename(columns={ROW: ROW_SHOWN})
    for c in out.columns:
        if str(out[c].dtype) in ("Int64", "Float64", "string"):
            out[c] = out[c].astype(object).where(out[c].notna(), None)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path) and not a.overwrite:
        new = rt.timestamped(path)
        rt.note_recovered(f"{os.path.basename(path)}이(가) 이미 있음 -> {os.path.basename(new)}로 저장 (덮어쓰기는 --overwrite)")
        path = new
    if path.lower().endswith(".csv"):
        writer = lambda p: out.to_csv(p, index=False, encoding="utf-8-sig")
    else:
        # 날짜만 있는 열(시각이 모두 0시)은 엑셀에서도 날짜만 보이게 (MBOM 2026-11-14: Eff From이 '2026-01-01 00:00:00'로 보임)
        dates = [c for c in out.columns if pd.api.types.is_datetime64_any_dtype(out[c])]
        date_only = all(((out[c].dropna().dt.normalize() == out[c].dropna())).all() for c in dates)
        fmt = "yyyy-mm-dd" if date_only else "yyyy-mm-dd hh:mm:ss"

        def writer(p: str) -> None:
            with pd.ExcelWriter(p, engine="openpyxl") as xw:
                out.to_excel(xw, index=False, sheet_name=name[:31])
                ws = xw.sheets[name[:31]]
                for j, c in enumerate(out.columns, 1):
                    if c in dates:
                        for (cell,) in ws.iter_rows(min_row=2, max_row=len(out) + 1, min_col=j, max_col=j):
                            cell.number_format = fmt
    path = rt.save_with_lock_fallback(path, writer, "내보낼 파일")
    try:
        if path.lower().endswith(".csv"):
            back = len(pd.read_csv(path, encoding="utf-8-sig", dtype=str))
        else:
            back = len(pd.read_excel(path, dtype=str))
        verified = back == len(df)
    except Exception:
        verified = False
    preview = f"표 {name}({len(df)}행 {len(data_cols(df))}열)을 저장했습니다: {path}"
    return rt.ok("export", preview, verified, "preview를 사용자에게 알리세요. 같은 명령을 다시 실행하지 마세요.",
                 None if verified else ["저장한 파일을 다시 읽어 행 수를 확인하지 못함"], table=name, count=len(df),
                 output_file=path, work=work.folder)


# ==============================================================================
# 7-1. chart: a table drawn as a PNG picture (matplotlib, no window). Not a step: the table is not changed
# ==============================================================================

CHART_KINDS = ("bar", "line", "scatter", "pie", "hist", "box")
CHART_KIND_NAMES = {"bar": "막대", "line": "선", "scatter": "산점", "pie": "원", "hist": "히스토그램", "box": "상자"}
BAR_MAX = 30                  # bars (x values) drawn at most, in table order (the rest: warnings + truncated)
PIE_MAX = 12                  # pie slices: the largest 11 and the rest summed as 기타
UNIT_NAMES = {"uom", "unit", "units", "um", "단위"}   # unit columns (names without spaces, / . _ -, lower case)
KOREAN_FONTS = ("Malgun Gothic", "NanumGothic", "AppleGothic", "Noto Sans CJK KR")   # same order as ml-kit charts


def _numeric_columns(df: Any) -> List[str]:
    pd = pd_mod()
    return [c for c in data_cols(df) if pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c])]


def _chart_y(text: str, df: Any, meta: Dict[str, Any]) -> List[str]:
    cols = column_list(text, df, meta, "--y")
    nums = _numeric_columns(df)
    bad = [c for c in cols if c not in nums]
    if bad:
        if nums:
            raise rt.ask(nums[:10], [argv_setting("--y", n) for n in nums[:10]],
                         f"--y의 '{', '.join(bad)}'은(는) 숫자 열이 아니어서 그릴 숫자 열")
        raise rt.InputError(f"표에 숫자 열이 없어 --y로 그릴 수 없습니다 ('{', '.join(bad)}'은(는) 글자 열).")
    return cols


def _chart_path(work: Work, name: str, kind: str, save: Optional[str], overwrite: bool) -> str:
    if save:
        p = str(save).strip().strip('"')
        if not os.path.splitext(p)[1]:
            p += ".png"
        if os.path.splitext(p)[1].lower() != ".png":
            raise rt.InputError(f"차트는 .png로 저장합니다. [추천] {rt.full_command(argv_setting('--save', os.path.splitext(p)[0] + '.png'))}")
        path = p if os.path.isabs(p) else os.path.join(work.folder, p)
    else:
        path = os.path.join(work.folder, "results", f"{name}_{kind}.png")
    path = os.path.abspath(path)
    if os.path.exists(path) and not overwrite:
        new = rt.timestamped(path)
        rt.note_recovered(f"{os.path.basename(path)}이(가) 이미 있음 -> {os.path.basename(new)}로 저장 (덮어쓰기는 --overwrite)")
        path = new
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path


def cmd_chart(a: Any) -> Dict[str, Any]:
    pd = pd_mod()
    matplotlib = rt.require_lib("matplotlib", "matplotlib", "차트")
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    work = open_work(a)
    work.check_sources()
    name, sid = work.resolve(a.source, "--from")
    df, meta = work.load(sid)
    kind = a.kind
    warnings: List[str] = list(work.warnings)
    ys = _chart_y(a.y, df, meta)
    x = resolve_column(a.x, df, meta, "--x") if a.x else None
    if kind in ("hist", "box") and x:
        warnings.append(f"{CHART_KIND_NAMES[kind]} 그림은 --x를 쓰지 않습니다 (--y 열의 값만 그림).")
        x = None
    if kind in ("pie", "scatter") and not x:
        raise rt.InputError(f"{CHART_KIND_NAMES[kind]} 그림에는 --x가 필요합니다 ("
                            + ("조각 이름 열" if kind == "pie" else "가로축 숫자 열") + f"). 열: {name_list(data_cols(df))}")
    if kind == "pie" and len(ys) != 1:
        raise rt.InputError(f"원 그림은 --y에 숫자 열 하나만 줍니다 (받은 값: {', '.join(ys)}). [추천] "
                            f"{rt.full_command(argv_setting('--y', ys[0]))}")
    if kind == "scatter" and x not in _numeric_columns(df):
        raise rt.InputError(f"산점 그림의 --x는 숫자 열이어야 합니다 ('{x}'은(는) 글자 열). 숫자 열: {name_list(_numeric_columns(df))}")
    # data to draw
    if x:
        label_col = x
        data = df[[x] + [c for c in ys if c != x]].copy()
    else:
        label_col = ROW_SHOWN if ROW in df.columns else "순서"
        data = df[ys].copy()
        data.insert(0, label_col, df[ROW].tolist() if ROW in df.columns else list(range(1, len(df) + 1)))
    missing = int(data[ys].isna().any(axis=1).sum())
    if missing:
        warnings.append(f"--y 값이 빈 행 {missing}개는 그리지 않았습니다.")
        data = data.dropna(subset=ys)
    unit = next((c for c in data_cols(df) if re.sub(r"[\s/._-]+", "", str(c)).lower() in UNIT_NAMES), None)
    if unit and unit not in ys and unit != x:
        units = [u for u in df.loc[data.index, unit].astype("string").str.strip().dropna().unique().tolist() if u]
        if len(units) > 1:                       # 집 PC 2026-11-13: EA·KG·M·ML 수량을 한 막대에 더했음
            summed = kind in ("bar", "pie") and x and data[x].astype("string").duplicated().any()
            warnings.append(f"'{unit}' 열의 단위가 여러 가지({', '.join(units[:6])}{' 등' if len(units) > 6 else ''})라 "
                            f"단위가 다른 값을 함께 {'더해 ' if summed else ''}그렸습니다. 단위 하나만 보려면 먼저 "
                            f"{rt.full_command(['select', '--from', name, '--where', f'{unit}={units[0]}'])} 로 고른 표로 그리세요.")
    truncated = False
    if kind == "line" and x and data[x].astype("string").duplicated().any():   # 사용자 결정 2026-11-14: 합계하지 않고 경고
        warnings.append(f"--x '{x}'의 같은 값이 여러 행이라 선이 표 순서대로 같은 x를 오가며 그려졌습니다. 값마다 합계로 보려면 "
                        f"{rt.full_command(argv_setting('--kind', 'bar'))} 로 막대를 그리세요.")
    if kind in ("bar", "pie") and x and data[x].astype("string").duplicated().any():
        data = data.assign(**{x: data[x].astype("string")}).groupby(x, sort=False, dropna=False)[ys].sum().reset_index()
        warnings.append(f"--x '{x}'의 같은 값이 여러 행이라 값마다 --y를 합계로 그렸습니다.")
    if kind == "bar" and len(data) > BAR_MAX:
        warnings.append(f"막대가 {len(data)}개라 표 순서대로 앞 {BAR_MAX}개만 그렸습니다 (전부는 select --where로 줄인 뒤).")
        data = data.head(BAR_MAX)
        truncated = True
    if kind == "pie":
        if (data[ys[0]] < 0).any():
            raise rt.InputError(f"원 그림은 음수 값을 그릴 수 없습니다 ('{ys[0]}'에 음수가 있음). 막대 그림(--kind bar)을 쓰세요.")
        if len(data) > PIE_MAX:
            data = data.sort_values(ys[0], ascending=False)
            rest = data.iloc[PIE_MAX - 1:][ys[0]].sum()
            data = pd.concat([data.iloc[:PIE_MAX - 1], pd.DataFrame({label_col: ["기타"], ys[0]: [rest]})], ignore_index=True)
            warnings.append(f"조각이 많아 큰 {PIE_MAX - 1}개와 나머지 합계(기타)로 그렸습니다.")
    if len(data) == 0:
        return rt.ok("chart", f"표 {name}에 그릴 행이 없습니다 (0행)", True,
                     "0건. 그릴 행이 없습니다(찾은 행 없음). 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요.",
                     table=name, count=0, work=work.folder)
    title = a.title or f"{name}: {', '.join(ys)}" + (f" ({x}별)" if x and kind in ("bar", "line", "pie") else "")
    have = {f.name for f in font_manager.fontManager.ttflist}
    font = next((f for f in KOREAN_FONTS if f in have), None)
    texts = [title, str(label_col)] + [str(c) for c in ys] + [str(v) for v in data[label_col].tolist()[:BAR_MAX]]
    if not font and any(ord(ch) > 127 for t in texts for ch in t):
        warnings.append("한글 글꼴(맑은 고딕 등)을 찾지 못해 한글이 네모로 보일 수 있습니다.")
    rc = {"axes.unicode_minus": False, **({"font.family": font} if font else {})}
    path = _chart_path(work, name, kind, a.save, a.overwrite)
    import warnings as _w
    with _w.catch_warnings(), matplotlib.rc_context(rc):
        _w.simplefilter("ignore")
        fig, ax = plt.subplots(figsize=(8, 4.8))
        labels = data[label_col].astype("string").fillna("").tolist()
        if kind == "bar":
            n = len(ys)
            width = 0.8 / n
            pos = list(range(len(data)))
            for i, c in enumerate(ys):
                ax.bar([p + (i - (n - 1) / 2) * width for p in pos], data[c].astype(float), width=width, label=c)
            ax.set_xticks(pos)
            ax.set_xticklabels(labels, rotation=45 if len(pos) > 8 else 0, ha="right" if len(pos) > 8 else "center")
        elif kind == "line":
            for c in ys:
                ax.plot(range(len(data)), data[c].astype(float), marker="o", markersize=3, label=c)
            step = max(1, len(data) // 20)
            ax.set_xticks(list(range(0, len(data), step)))
            ax.set_xticklabels(labels[::step], rotation=45 if len(data) > 8 else 0, ha="right" if len(data) > 8 else "center")
        elif kind == "scatter":
            for c in ys:
                ax.scatter(data[x].astype(float), data[c].astype(float), s=14, label=c)
        elif kind == "pie":
            ax.pie(data[ys[0]].astype(float), labels=labels, autopct="%1.1f%%", startangle=90, counterclock=False)
            ax.axis("equal")
        elif kind == "hist":
            for c in ys:
                ax.hist(data[c].astype(float), bins="auto", alpha=0.6 if len(ys) > 1 else 1.0, label=c)
        else:
            ax.boxplot([data[c].astype(float) for c in ys])
            ax.set_xticks(list(range(1, len(ys) + 1)))
            ax.set_xticklabels(ys)
        if kind in ("bar", "line", "scatter"):
            ax.set_xlabel(str(label_col))
            ax.set_ylabel(ys[0] if len(ys) == 1 else "값")
        if kind == "hist":
            ax.set_xlabel(ys[0] if len(ys) == 1 else "값")
            ax.set_ylabel("행 수")
        if kind in ("bar", "line", "scatter", "hist") and len(ys) > 1:
            ax.legend()
        if kind != "pie":
            ax.grid(axis="y", alpha=0.3)
        ax.set_title(title)
        fig.tight_layout()
        path = rt.save_with_lock_fallback(path, lambda p: fig.savefig(p, dpi=120), "차트")
        plt.close(fig)
    try:
        import matplotlib.image as mpimg
        img = mpimg.imread(path)
        verified = os.path.getsize(path) > 0 and img.shape[0] > 0 and img.shape[1] > 0
    except Exception:
        verified = False
    points = len(data)
    what = {"bar": "막대", "pie": "조각", "line": "점", "scatter": "점", "hist": "값", "box": "값"}[kind]
    preview = (f"차트를 저장했습니다: {path} ({CHART_KIND_NAMES[kind]} 그림, 표 {name}, "
               + (f"x={x}, " if x else "") + f"y={', '.join(ys)}, {what} {points}개)")
    out = rt.ok("chart", preview, verified,
                "preview를 사용자에게 알리세요. 같은 명령을 다시 실행하지 마세요. 문서에 넣으려면 office-cli 그림 넣기 명령에 output_file을 씁니다.",
                None if verified else ["저장한 그림을 다시 열어 확인하지 못함"],
                table=name, kind=kind, x=x, y=ys, count=points, truncated=truncated, output_file=path, work=work.folder)
    if warnings:
        out["warnings"] = warnings
        out["agent_directive"] += " warnings도 그대로 알리세요."
    return out


# ==============================================================================
# 8. Parser
# ==============================================================================

def build_cli() -> rt.JsonArgumentParser:
    root, sub = rt.make_parser("pandas-kit: 표 데이터를 단계별로 다룸 (원본은 고치지 않고 단계마다 이름 붙인 표를 만듦)")

    def cmd(name: str, help_text: str) -> rt.JsonArgumentParser:
        p = rt.add_command(sub, root, name, help_text)
        p.add_argument("--work", help="이 작업 폴더로 돌아가기 (지난 작업). 생략하면 지금 작업 폴더")
        return p

    p = cmd("select", "파일(시트·범위)이나 만든 표에서 열과 조건으로 골라 이름 붙인 표 만들기")
    p.add_argument("--file", help="읽을 파일 (.xlsx, .xlsm, .csv). 만든 표에서 고를 때는 --from (둘 다 없으면 도구가 읽을 것을 묻는다)")
    p.add_argument("--sheet", help="시트 이름 (생략하면 파일을 열었을 때 보이는 시트)")
    p.add_argument("--range", help="읽을 범위 (예: B5:H230, B:H, 왼쪽 위 칸 B5). 생략하면 도구가 표를 찾음")
    p.add_argument("--from", dest="source", help="고를 표 이름 (--file이 없으면 꼭 씀)")
    p.add_argument("--columns", help="남길 열 (이름이나 열 글자, ','로 나눔)")
    p.add_argument("--where", help="행 조건 (그리고 ;, 또는 |. 예: \"Level=1;Qty>2\", \"Maker=A|B\")")
    p.add_argument("--name", help="새 표 이름 (생략하면 df1, df2 ...)")

    p = cmd("show", "표의 앞 20행을 엑셀 행 번호와 함께 보기 (넘으면 전체를 csv로)")
    p.add_argument("--from", dest="source", help="볼 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", help="볼 열 (이름이나 열 글자, ','로 나눔)")
    p.add_argument("--start", type=int, help="몇 번째 행부터 볼지 (표 안의 순서, 1부터, 생략하면 1). 이어 보기는 next_commands")

    for name, help_text in (("join", "두 표를 키 열로 짝 맞춰 붙이기 (양쪽에 있는 행만)"),
                            ("only", "왼쪽 표에만 있는 행 (키 열로 비교)")):
        p = cmd(name, help_text)
        p.add_argument("--left", required=True, help="왼쪽 표")
        p.add_argument("--right", required=True, help="오른쪽 표")
        p.add_argument("--on", required=True, help="키 열 (','로 나눔)")
        p.add_argument("--name", help="새 표 이름")
        if name == "join":
            p.add_argument("--keep-left", action="store_true", help="왼쪽 표의 행을 모두 남김 (짝이 없으면 오른쪽 열이 빈 칸)")

    p = cmd("diff", "두 표 비교: 추가·삭제 행과 바뀐 칸 (키 열로 짝)")
    p.add_argument("--old", required=True, help="옛 표")
    p.add_argument("--new", required=True, help="새 표")
    p.add_argument("--on", required=True, help="키 열 (','로 나눔, 행마다 달라야 함)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("stack", "여러 표를 위아래로 잇기 ('표' 열에 어느 표의 행인지)")
    p.add_argument("--tables", required=True, help="이을 표 이름 (','로 나눔)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("group", "열 값별로 묶어 행 수와 합계·평균·최소·최대·개수·종류 수")
    p.add_argument("--from", dest="source", help="묶을 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--by", required=True, help="묶을 열 (','로 나눔)")
    p.add_argument("--sum", help="합계를 낼 숫자 열 (','로 나눔, 생략하면 행 수만)")
    p.add_argument("--mean", help="평균을 낼 숫자 열 (','로 나눔)")
    p.add_argument("--min", help="최소를 볼 숫자 열 (','로 나눔)")
    p.add_argument("--max", help="최대를 볼 숫자 열 (','로 나눔)")
    p.add_argument("--count", help="빈 칸이 아닌 칸 수를 셀 열 (','로 나눔)")
    p.add_argument("--nunique", help="서로 다른 값의 수를 셀 열 (','로 나눔)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("sort", "행 정렬 (작은 것부터, --desc는 큰 것부터)")
    p.add_argument("--from", dest="source", help="정렬할 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--by", required=True, help="정렬 기준 열 (','로 나눔, 앞 열이 먼저)")
    p.add_argument("--desc", action="store_true", help="큰 것부터")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("describe", "열마다 요약: 숫자 열은 개수·평균·표준편차·최소·25%·50%·75%·최대, 그 밖의 열은 개수·종류 수·가장 많은 값")
    p.add_argument("--from", dest="source", help="요약할 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", help="요약할 열 (','로 나눔, 생략하면 모든 열)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("unique", "열의 서로 다른 값과 그 개수 (빈 칸은 '(빈 칸)')")
    p.add_argument("--from", dest="source", help="볼 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", required=True, help="값을 볼 열 (','로 나눔, 열마다 따로 셈)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("calc", "식으로 계산한 열을 더함 (열 이름, 숫자, + - * /, 괄호)")
    p.add_argument("--from", dest="source", help="계산할 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--expr", required=True, help="\"새 열 이름=식\" (여러 개는 ';'로 나눔, 공백·기호가 든 열 이름은 [열 이름]. 예: \"금액=Qty*[Unit Price]\")")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("top", "숫자 열이 큰 순으로 앞 몇 행 (--asc는 작은 순)")
    p.add_argument("--from", dest="source", help="고를 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--by", required=True, help="순위를 매길 숫자 열 (','로 나눔, 앞 열이 먼저)")
    p.add_argument("--n", type=int, default=10, help="남길 행 수 (생략하면 10)")
    p.add_argument("--asc", action="store_true", help="작은 순")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("dropna", "빈 칸이 있는 행 빼기")
    p.add_argument("--from", dest="source", help="표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", required=True, help="빈 칸이면 안 되는 열 (','로 나눔, 하나라도 빈 칸이면 그 행을 뺌)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("fillna", "빈 칸 채우기 (--value 값으로, --down 바로 위 칸의 값으로)")
    p.add_argument("--from", dest="source", help="표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", required=True, help="빈 칸을 채울 열 (','로 나눔)")
    p.add_argument("--value", help="채울 값 (숫자 열에는 숫자, 날짜 열에는 2026-12-31)")
    p.add_argument("--down", action="store_true", help="바로 위 칸의 값으로 채움 (병합 칸을 푼 표, 위와 같은 값을 비워 둔 표)")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("drop-duplicates", "값이 같은 행은 하나만 남기기")
    p.add_argument("--from", dest="source", help="표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", help="같은지 볼 열 (','로 나눔, 생략하면 모든 열)")
    p.add_argument("--keep", choices=("first", "last"), default="first", help="남길 행: first 먼저 나온 행(기본), last 마지막 행")
    p.add_argument("--name", help="새 표 이름")

    p = cmd("rename", "열 이름 바꾸기")
    p.add_argument("--from", dest="source", help="표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--columns", required=True, help="\"지금 이름=새 이름\" (','로 나눔. 예: \"Qty=수량,UoM=단위\")")
    p.add_argument("--name", help="새 표 이름")

    cmd("undo", "마지막 단계 하나 되돌리기")
    rt.add_command(sub, root, "new", "새 작업 시작 (다음 select --file이 새 작업 폴더를 만듦)")

    p = cmd("chart", "표를 차트 그림(.png)으로 저장 (막대·선·산점·원·히스토그램·상자)")
    p.add_argument("--from", dest="source", help="그릴 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--kind", choices=CHART_KINDS, default="bar",
                   help="bar 막대(기본), line 선, scatter 산점, pie 원, hist 히스토그램, box 상자")
    p.add_argument("--x", help="가로축(막대·선) 또는 조각 이름(원) 열, 산점은 숫자 열. 생략하면 엑셀 행 순서 (원·산점은 필요)")
    p.add_argument("--y", required=True, help="그릴 숫자 열 (','로 나눔, 원은 하나)")
    p.add_argument("--title", help="차트 제목 (생략하면 표 이름과 열 이름)")
    p.add_argument("--save", nargs="?", const=None, help="저장할 그림 (.png, 생략하면 작업 폴더 results\\<표>_<종류>.png)")
    p.add_argument("--overwrite", action="store_true", help="같은 이름 그림이 있으면 덮어쓰기 (사용자가 원할 때만)")

    p = cmd("export", "표를 새 파일(.xlsx, .csv)로 저장")
    p.add_argument("--from", dest="source", help="저장할 표 이름 (꼭 씀, 빼면 [후보]로 물음)")
    p.add_argument("--save", nargs="?", const=None, help="저장할 파일 (생략하면 작업 폴더의 <표 이름>.xlsx, 이름만 주면 작업 폴더에)")
    p.add_argument("--overwrite", action="store_true", help="같은 이름 파일이 있으면 덮어쓰기 (사용자가 원할 때만)")
    return root


COMMANDS = {"select": cmd_select, "show": cmd_show, "join": cmd_join, "only": cmd_only, "diff": cmd_diff,
            "stack": cmd_stack, "undo": cmd_undo, "new": cmd_new, "export": cmd_export, "chart": cmd_chart,
            "group": cmd_group, "sort": cmd_sort, "describe": cmd_describe, "unique": cmd_unique, "calc": cmd_calc,
            "top": cmd_top, "dropna": cmd_dropna, "fillna": cmd_fillna, "drop-duplicates": cmd_drop_duplicates, "rename": cmd_rename}


def main(argv: Optional[List[str]] = None) -> None:
    rt.setup_console()
    rt.CURRENT_ARGV = [os.path.abspath(__file__)] + list(sys.argv[1:] if argv is None else argv)
    args = build_cli().parse_args(rt.CURRENT_ARGV[1:])
    rt.run_cli(lambda: _with_switch_notice(args), hint="확인할 것: 표 이름, 열 이름, 파일 경로.")


def _with_switch_notice(args: Any) -> Dict[str, Any]:
    """명령 실행. --work로 지금 작업 폴더가 바뀌었으면 preview(실패면 message) 앞에 그 사실을 붙임."""
    SWITCHED.clear()
    try:
        out = COMMANDS[args.cmd](args)
    except rt.InputError as e:
        if SWITCHED and e.args:
            e.args = (f"{SWITCHED[0]} {e.args[0]}",) + e.args[1:]
        raise
    if SWITCHED:
        out["preview"] = f"{SWITCHED[0]} {out.get('preview', '')}".strip()
    return out


if __name__ == "__main__":
    main()
