# -*- coding: utf-8 -*-
"""
_pd_runtime.py - pandas-kit shared runtime (not an entry point; the agent never runs it).

pandas_kit.py imports the Agent CLI Protocol (one JSON line on stdout, run_cli, JsonArgumentParser), the Windows-safe file
helpers (retry_fs, save_with_lock_fallback, read_text_file), the command quoting (quote_arg) and the open Excel workbook
lookup from here. Copied from ml-kit _ml_runtime.py with the same names and the same behaviour (AIKIT/dev/README.md R1-R10):
tool folders are deployed alone and never import each other.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

VERSION = "0.2.2"
TOOL_NAME = "pandas-kit"
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_ENV = "PANDAS_KIT_DATA_DIR"
ERROR_LOG_MAX = 2 * 1024 * 1024
SKILL_NAME = "pandas-kit"
# The running command (script + arguments) for error logs and the [추천] full commands. Set by run_cli / the parser.
CURRENT_ARGV: Optional[List[str]] = None


def setup_console() -> None:
    """Forces UTF-8 on stdout/stderr (Windows consoles and pipes default to CP949)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and (stream.encoding or "").lower() != "utf-8":
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def json_safe(obj: Any) -> Any:
    """Plain JSON values: NumPy scalars, NaN/NA (-> None), Timestamps (-> ISO text)."""
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [json_safe(v) for v in obj]
    if obj is None or isinstance(obj, (str, bool, int)):
        return obj
    if isinstance(obj, float):
        return obj if obj == obj and obj not in (float("inf"), float("-inf")) else None
    try:
        import numpy as np
        import pandas as pd
    except ImportError:
        return str(obj)
    if obj is pd.NA or obj is pd.NaT:
        return None
    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, np.integer):
        return int(obj)
    if isinstance(obj, np.floating):
        f = float(obj)
        return f if np.isfinite(f) else None
    if isinstance(obj, (pd.Timestamp, datetime)):
        return obj.strftime("%Y-%m-%d") if (obj.hour, obj.minute, obj.second) == (0, 0, 0) else obj.isoformat(sep=" ")
    if isinstance(obj, np.ndarray):
        return [json_safe(v) for v in obj.tolist()]
    return str(obj)


def dumps_json(obj: Any) -> str:
    return json.dumps(json_safe(obj), ensure_ascii=False)


_OUT_DONE: List[bool] = [False]


def emit_json_response(payload: Dict[str, Any]) -> bool:
    """Prints the one JSON line of this command (only the first call writes)."""
    setup_console()
    if _OUT_DONE[0]:
        return False
    sys.stdout.write(dumps_json(payload) + "\n")
    sys.stdout.flush()
    _OUT_DONE[0] = True
    return True


class InputError(ValueError):
    """User/agent input or data problem: retrying with the same arguments will not help (STOP_RETRY).
    extra: keys added to the error JSON (confirm_commands of a [후보]); error_type 'needs_user_decision' for a choice."""

    def __init__(self, message: str = "", extra: Optional[Dict[str, Any]] = None, error_type: Optional[str] = None):
        super().__init__(str(message))
        self.extra = dict(extra or {})
        self.error_type = error_type


class MissingLibraryError(InputError):
    def __init__(self, pip_name: str, feature: str):
        super().__init__(f"{feature}에 필요한 {pip_name}이(가) 설치되어 있지 않습니다.")
        self.pip_name = pip_name
        self.feature = feature


def require_lib(module: str, pip_name: str, feature: str) -> Any:
    """import module or MissingLibraryError (R5: STOP_RETRY with install_command)."""
    import importlib
    try:
        return importlib.import_module(module)
    except ImportError:
        raise MissingLibraryError(pip_name, feature) from None


_STOP_RETRY_ERRORS = (InputError, FileNotFoundError, NotADirectoryError, IsADirectoryError, PermissionError)


def error_payload(exc: BaseException, hint: str = "") -> Dict[str, Any]:
    """The standard error JSON with an agent_directive the LLM can act on (same words as ml-kit)."""
    msg = str(exc) or type(exc).__name__
    if isinstance(exc, KeyboardInterrupt):
        return {"status": "error", "tool": TOOL_NAME, "error_type": "interrupted",
                "message": "사용자 또는 상위 프로세스에 의해 실행이 중단되었습니다.",
                "agent_directive": "STOP_RETRY: 실행이 중단되었습니다. 다시 실행하기 전에 사용자에게 물으세요."}
    if isinstance(exc, MissingLibraryError):
        return {"status": "error", "tool": TOOL_NAME, "error_type": "missing_library", "message": msg,
                "missing_package": exc.pip_name, "install_command": f"python -m pip install {exc.pip_name}",
                "agent_directive": (f"STOP_RETRY: {exc.feature}에 필요한 {exc.pip_name}이(가) 설치되어 있지 않습니다. "
                                    "install_command를 사용자에게 전달하세요. 직접 설치하지 마세요.")}
    extra = {k: v for k, v in getattr(exc, "extra", {}).items() if k not in ("status", "message", "agent_directive", "error_type")}
    if isinstance(exc, InputError) and exc.error_type == "needs_user_decision":
        return {"status": "error", "tool": TOOL_NAME, "error_type": "needs_user_decision", "message": msg, **extra,
                "agent_directive": ("STOP_RETRY: 사용자가 정할 일이 있어 멈췄습니다. message를 그대로 알리고 사용자에게 물으세요. "
                                    "스스로 실행하지 말고, 사용자가 [후보]에서 번호를 고르면 같은 번호의 "
                                    "confirm_commands[번호-1].command를 글자 그대로 1번 실행하세요.")}
    if isinstance(exc, _STOP_RETRY_ERRORS):
        return {"status": "error", "tool": TOOL_NAME, "error_type": "input", "message": msg, **extra,
                "agent_directive": (
                    "STOP_RETRY: 입력이나 데이터 문제로 멈췄습니다. message에 [추천]이 있으면 [추천] 뒤의 전체 명령을 그대로 1번만 실행하세요. "
                    "[후보]가 있으면 다시 실행하지 말고 후보를 사용자에게 보여 주고, 사용자가 고르면 같은 번호의 confirm_commands를 1번 실행하세요. "
                    "둘 다 없으면 다시 실행하지 말고 message를 그대로 알리세요." + (f" {hint}" if hint else ""))}
    return {"status": "error", "tool": TOOL_NAME, "error_type": "internal", "message": f"{type(exc).__name__}: {msg}",
            "agent_directive": "ERROR: 도구 내부 오류입니다. 다시 실행하지 말고 message를 그대로 알리세요."}


_RECOVERED: List[str] = []
_RECOVERY_ATTEMPTS: List[str] = []


def note_recovered(msg: str) -> None:
    """A JIT recovery that changed the result (R3): stderr [WARN] and JSON `recovered`."""
    msg = str(msg)[:200]
    if msg not in _RECOVERED:
        sys.stderr.write(f"[WARN] {msg}\n")
        _RECOVERED.append(msg)


def record_recovery_attempt(msg: str) -> None:
    """A recovery that hit its limit (R4): the '복구 시도:' line of the error log."""
    _RECOVERY_ATTEMPTS.append(str(msg)[:200])


def _with_recovered(payload: Dict[str, Any]) -> Dict[str, Any]:
    return {**payload, "recovered": list(_RECOVERED)} if _RECOVERED and "recovered" not in payload else payload


def get_data_dir() -> str:
    """Tool data folder: PANDAS_KIT_DATA_DIR > <tool folder>/data (the current work folder pointer and the error log)."""
    return os.path.abspath(os.environ.get(DATA_ENV) or os.path.join(TOOL_DIR, "data"))


def error_log_path() -> str:
    return os.path.join(get_data_dir(), "logs", "errors.txt")


# The shell that runs the commands the tools print (next_commands, [추천] full commands) as given.
# Same order as the OpenCode shell tool: pwsh (PowerShell 7) when on PATH, else Windows PowerShell 5.1.
# PANDAS_KIT_SHELL=pwsh|powershell|cmd|bash forces one. Both PowerShells are exact, cmd and bash best effort.
SHELL_ENV = "PANDAS_KIT_SHELL"
SHELLS = ("pwsh", "powershell", "cmd", "bash")
_FOUND_SHELL: List[str] = []
# Characters PowerShell reads as a single quote (doubled inside a single-quoted string)
_PS_SINGLE_QUOTES = "'\u2018\u2019\u201a\u201b"
_PLAIN_PUNCT = "_-.:\\/,=+@!"
_PLAIN_INT = re.compile(r"-?(?:0|[1-9]\d{0,8})")
_PLAIN_OPTION = re.compile(r"--?[A-Za-z][A-Za-z0-9-]*")


def command_shell() -> str:
    """The shell for command strings: PANDAS_KIT_SHELL > pwsh when on PATH > powershell (Windows PowerShell 5.1)."""
    forced = os.environ.get(SHELL_ENV, "").strip().lower()
    if forced in SHELLS:
        return forced
    if not _FOUND_SHELL:
        _FOUND_SHELL.append("pwsh" if shutil.which("pwsh") else "powershell")
    return _FOUND_SHELL[0]


def _net_space(ch: str) -> bool:
    """.NET char.IsWhiteSpace (what makes Windows PowerShell 5.1 wrap an argument in double quotes)."""
    import unicodedata
    return ch in "\t\n\v\f\r\x85\xa0" or unicodedata.category(ch) in ("Zs", "Zl", "Zp")


def _plain_arg(s: str, shell: str) -> bool:
    """A value every shell passes as is without quotes: letters and digits (any script) and _-.:\\/,=+@! only.
    Quoted anyway: shapes PowerShell may turn into a number (007, 1e3, 0x10, 1.50, 2-8), a leading - that is not an
    option (-a.b splits in two), a leading = : + @ (array, splatting) !, a leading, trailing or doubled comma (array)
    and, for bash, a backslash."""
    if not s or s == "--":
        return False
    for ch in s:
        if not (ch.isalnum() if not ch.isascii() else (ch.isalnum() or ch in _PLAIN_PUNCT)):
            return False
    if s[0] in "=:+,@!" or s[-1] == "," or ",," in s:
        return False
    if shell == "bash":
        return "\\" not in s
    if s[0] == "-" and not (_PLAIN_OPTION.fullmatch(s) or _PLAIN_INT.fullmatch(s)):
        return False
    if (s[0].isdigit() or s[0] == "." and s[1:2].isdigit()) and not _PLAIN_INT.fullmatch(s):
        return False
    return True


def _argv_escape(s: str, wrapped: bool, quote: str = '\\"') -> str:
    """Escapes s so the Windows command-line parser (CommandLineToArgvW, Python argv) gets it back: backslashes before
    a " are doubled and the " becomes quote; in a wrapped value the trailing backslashes are doubled too (before the closing ")."""
    out: List[str] = []
    slashes = 0
    for ch in s:
        if ch == "\\":
            slashes += 1
            continue
        if ch == '"':
            out.append("\\" * (2 * slashes) + quote)
        else:
            out.append("\\" * slashes + ch)
        slashes = 0
    out.append("\\" * (2 * slashes if wrapped else slashes))
    return "".join(out)


def _ps_single(s: str) -> str:
    return "'" + "".join(ch * 2 if ch in _PS_SINGLE_QUOTES else ch for ch in s) + "'"


def quote_arg(value: Any, shell: Optional[str] = None) -> str:
    """Quotes one argument of a command string for the shell that runs it (command_shell).

    - A value without shell special characters (_plain_arg) stays as is.
    - PowerShell: single quotes '...' (a ' inside is '') keep $, backtick, {}, @, #, ; and ( literal.
      Windows PowerShell 5.1 hands the value to python wrapped in "..." when it has a space but leaves a " inside as is,
      so " -> \\" (backslashes before it doubled), trailing backslashes of a wrapped value doubled and an empty value
      '""' (5.1 drops an empty argument) are done here. PowerShell 7.3+ escapes by itself, so nothing is done for it.
    - cmd: "..." (a " inside is "", % is cut as "^%" so %NAME% is not expanded). bash: '...' (a ' inside is '\\'')."""
    s = str(value)
    shell = shell or command_shell()
    if _plain_arg(s, shell):
        return s
    if shell == "pwsh":
        return _ps_single(s)
    if shell == "bash":
        return "'" + s.replace("'", "'\\''") + "'"
    if shell == "cmd":
        return '"' + _argv_escape(s, True, '""').replace("%", '"^%"') + '"'
    if not s:
        return _ps_single('""')
    return _ps_single(_argv_escape(s, any(_net_space(ch) for ch in s)))




def write_error_log(payload: Dict[str, Any], trace: str = "", argv: Optional[List[str]] = None) -> None:
    """Appends one failed run to errors.txt (R6). Never raises. Only paths and arguments, never cell values."""
    try:
        import platform
        path = error_log_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            if os.path.exists(path) and os.path.getsize(path) > ERROR_LOG_MAX:
                os.replace(path, path[:-4] + ".old.txt")
        except Exception:
            pass
        if argv is None:
            argv = CURRENT_ARGV if CURRENT_ARGV is not None else [os.path.basename(sys.argv[0] or "")] + sys.argv[1:]
        lines = [f"===== {datetime.now().isoformat(timespec='seconds')} | {TOOL_NAME} {VERSION} | "
                 f"Python {platform.python_version()} | {platform.platform()}",
                 f"명령: {' '.join(quote_arg(a) for a in argv)}",
                 f"결과: {payload.get('error_type')} | {payload.get('message')}"]
        if _RECOVERY_ATTEMPTS:
            lines.append(f"복구 시도: {'; '.join(_RECOVERY_ATTEMPTS)}")
        if trace:
            lines.append("traceback:\n" + trace.rstrip())
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n\n")
    except Exception:
        pass


def run_cli(execute: Callable[[], Dict[str, Any]], hint: str = "") -> None:
    """Runs a command body: exactly one JSON line on stdout, exit 0/1/130. Failures (exit 1) go to errors.txt."""
    global CURRENT_ARGV
    if CURRENT_ARGV is None:
        CURRENT_ARGV = [sys.argv[0]] + sys.argv[1:]
    _RECOVERED.clear()
    _RECOVERY_ATTEMPTS.clear()
    _OUT_DONE[0] = False
    try:
        result = execute()
    except KeyboardInterrupt as e:
        sys.stderr.write("\n[ERROR] 실행 중단 (KeyboardInterrupt)\n")
        emit_json_response(_with_recovered(error_payload(e)))
        sys.exit(130)
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        sys.stderr.write(f"\n[ERROR] {e}\n")
        payload = _with_recovered(error_payload(e, hint))
        write_error_log(payload, trace if payload.get("error_type") == "internal" else "")
        emit_json_response(payload)
        sys.exit(1)
    emit_json_response(_with_recovered(result))


# ------------------------------------------------------------------------------
# Results and full commands (AGENTS.md 사용법 단순성 2: the tool gives whole commands, never values to splice in)
# ------------------------------------------------------------------------------

def script_path() -> str:
    return os.path.join(TOOL_DIR, "pandas_kit.py")


def full_command(args: List[Any], script: Optional[str] = None) -> str:
    """'python "<script>" <args>' with every argument quoted for the shell that runs it (quote_arg)."""
    return " ".join([f"python {quote_arg(script or script_path())}"] + [quote_arg(a) for a in args])


def next_command(args: List[Any], purpose: str, script: Optional[str] = None, needs_input: bool = False,
                 destructive: bool = False) -> Dict[str, Any]:
    item: Dict[str, Any] = {"purpose": purpose, "command": full_command(args, script), "needs_input": needs_input}
    if destructive:
        item["destructive"] = True
    return item


def ask(cands: List[Any], argv_list: List[Optional[List[Any]]], what: str) -> InputError:
    """[후보] 1) a 2) b with the full command of each number as confirm_commands (R8). The caller raises it."""
    cands = [str(c) for c in cands][:10]
    text = "[후보] " + " ".join(f"{i}) {c}" for i, c in enumerate(cands, 1))
    items = [next_command(a, f"{what}: {c}") for c, a in zip(cands, argv_list) if a is not None]
    extra = {"confirm_commands": items} if len(items) == len(cands) else {}
    return InputError(f"{what}을(를) 정해야 합니다. {text} (사용자가 고른 번호의 confirm_commands를 실행)", extra,
                      "needs_user_decision" if extra else None)


def ok(command: str, preview: str, verified: bool, then: str = "", unverified: Optional[List[str]] = None,
       **fields: Any) -> Dict[str, Any]:
    """Success JSON (R1). verified -> 'VERIFIED: ...', else 'UNVERIFIED: ...' with the unverified list."""
    if verified:
        directive = "VERIFIED: " + (then or "preview를 사용자에게 알리세요.")
    else:
        directive = "UNVERIFIED: 확인하지 못한 것이 있습니다 (" + "; ".join(unverified or []) + "). 다시 실행하지 말고 그대로 알리세요. " + then
    payload: Dict[str, Any] = {"status": "ok", "tool": TOOL_NAME, "command": command, "verified": verified,
                               "preview": preview, **fields, "agent_directive": directive.strip()}
    if not verified:
        payload["unverified"] = list(unverified or [])
    if payload.get("warnings") and "warnings" not in directive:
        payload["agent_directive"] += " warnings도 그대로 알리세요."
    return payload


# ------------------------------------------------------------------------------
# Argument parser (R7 help as JSON, argument errors as JSON with [추천] for same-meaning option names)
# ------------------------------------------------------------------------------

# Option names a light LLM invents for an option that exists: [추천] the full command with the real option.
SAME_MEANING: Dict[str, List[str]] = {
    "--target": ["--file"], "--input": ["--file"], "--path": ["--file"], "--data": ["--file"],
    "--table": ["--from"], "--source": ["--from"], "--df": ["--from"],
    "--as": ["--name"], "--key": ["--on"], "--keys": ["--on"], "--by": ["--on"],
    "--cols": ["--columns"], "--column": ["--columns"], "--sheet-name": ["--sheet"],
    "--output": ["--save"], "--out": ["--save"], "--save-as": ["--save"], "--filter": ["--where"], "--query": ["--where"],
    "--base": ["--old"], "--before": ["--old"], "--after": ["--new"],
    "--first": ["--left"], "--second": ["--right"],
    "-f": ["--file"], "-o": ["--save"], "-w": ["--where"], "-n": ["--name"],
}
NO_FUZZY_OPTIONS = {"--overwrite", "--help"}


class _ProbeError(Exception):
    pass


class JsonArgumentParser(argparse.ArgumentParser):
    """argparse that answers help and usage errors with one JSON line (R7, R8). Abbreviations are not read."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        kwargs.setdefault("allow_abbrev", False)
        super().__init__(*args, **kwargs)
        self._probing = False

    def _root(self) -> "JsonArgumentParser":
        return getattr(self, "_root_parser", None) or self

    def _parses(self, args: List[str]) -> bool:
        root = self._root()
        root._probing = True
        subs = [a for a in root._actions if isinstance(a, argparse._SubParsersAction)]
        for s in subs:
            for p in s.choices.values():
                p._probing = True
        try:
            with contextlib.redirect_stderr(io.StringIO()):
                root.parse_args(args)
            return True
        except (_ProbeError, SystemExit):
            return False
        finally:
            root._probing = False
            for s in subs:
                for p in s.choices.values():
                    p._probing = False

    def _hint(self) -> Tuple[str, List[Dict[str, Any]]]:
        """[추천] the full command with same-meaning names replaced (when it parses), else [후보] lookalike option names."""
        import difflib
        given = [str(a) for a in (CURRENT_ARGV or [""] + sys.argv[1:])[1:]]
        if "-h" in given or "--help" in given:
            return "", []
        opts = {o for a in self._actions for o in a.option_strings}
        for act in self._actions:          # an unknown option reported by the root parser: the options of the sub-command used
            if isinstance(act, argparse._SubParsersAction) and given and given[0] in act.choices:
                opts |= {o for a in act.choices[given[0]]._actions for o in a.option_strings}
        fixed, changes, unknown = [], [], []
        for a in given:
            key, eq, val = a.partition("=")
            if re.fullmatch(r"--?[A-Za-z][A-Za-z0-9_-]*", key) and key not in opts:
                dashed = key.replace("_", "-")
                real = dashed if dashed in opts else next((o for o in SAME_MEANING.get(key, []) if o in opts), None)
                if real:
                    changes.append(f"{key} -> {real}")
                    fixed.append(real + eq + val)
                    continue
                unknown.append(key)
            fixed.append(a)
        if changes and not unknown and self._parses(fixed):
            return f" 같은 뜻의 옵션으로 고친 명령 ({', '.join(changes)}). [추천] {full_command(fixed)}", []
        for key in unknown:
            near = [o for o in difflib.get_close_matches(key, sorted(o for o in opts if o.startswith("--")), n=3, cutoff=0.6)
                    if o not in NO_FUZZY_OPTIONS]
            if near:
                items = [next_command([o + a[len(key):] if a == key or a.startswith(key + "=") else a for a in given], f"{key} 대신 {o}")
                         for o in near]
                text = "[후보] " + " ".join(f"{i}) {o}" for i, o in enumerate(near, 1))
                return f" '{key}' 옵션은 없습니다. {text} (뜻이 같은지 사용자에게 확인)", items
        return "", []

    def print_help(self, file: Any = None) -> None:  # type: ignore[override]
        try:
            super().print_help(sys.stderr)
        except Exception as e:
            sys.stderr.write(f"[WARN] 도움말을 만들지 못했습니다 ({type(e).__name__}). SKILL.md를 보세요.\n")
        skill = os.path.join(os.path.dirname(TOOL_DIR), ".opencode", "skills", SKILL_NAME, "SKILL.md")
        emit_json_response({
            "status": "ok", "tool": TOOL_NAME, "command": "help", "verified": True, "skill_file": skill,
            "agent_directive": f"VERIFIED: 도움말을 stderr에 출력했습니다. {SKILL_NAME} 스킬(SKILL.md)의 명령 템플릿을 그대로 사용하세요. "
                               "도구 코드(.py)는 읽지 마세요. 방법을 모르겠으면 사용자에게 물어보세요.",
        })

    def error(self, message: str) -> None:  # type: ignore[override]
        if self._probing:
            raise _ProbeError(message)
        self.print_usage(sys.stderr)
        try:
            hint, items = self._hint()
        except Exception:
            hint, items = "", []
        payload: Dict[str, Any] = {
            "status": "error", "tool": TOOL_NAME, "error_type": "input",
            "message": f"CLI 인자 오류: {message}.{hint}",
            "agent_directive": ("STOP_RETRY: 명령 인자가 잘못되었습니다. message에 [추천]이 있으면 [추천] 뒤의 명령을 그대로 1번만 "
                                "실행하세요. [후보]가 있으면 후보를 사용자에게 보여 주고 고른 번호의 confirm_commands를 1번 실행하세요. "
                                "둘 다 없으면 다시 실행하지 말고 message를 그대로 알리세요 (옵션 이름은 SKILL.md 기준)."),
        }
        if items:
            payload["confirm_commands"] = items
        write_error_log(payload)
        emit_json_response(payload)
        sys.exit(2)


def make_parser(description: str) -> Tuple[JsonArgumentParser, Any]:
    """The root parser and its sub-command action; every sub-parser is a JsonArgumentParser that knows the root."""
    root = JsonArgumentParser(description=description)
    sub = root.add_subparsers(dest="cmd", required=True, parser_class=JsonArgumentParser)
    root._sub = sub        # type: ignore[attr-defined]
    return root, sub


def add_command(sub: Any, root: JsonArgumentParser, name: str, help_text: str) -> JsonArgumentParser:
    p = sub.add_parser(name, help=help_text, description=help_text)
    p._root_parser = root  # type: ignore[attr-defined]
    return p


# ------------------------------------------------------------------------------
# Files (R4): retries for antivirus/Explorer handles, new name when the target is locked, encodings
# ------------------------------------------------------------------------------

def _transient_fs_error(e: OSError) -> bool:
    return isinstance(e, PermissionError) or getattr(e, "winerror", None) in (5, 32, 33)


def retry_fs(fn: Callable[[], Any], attempts: int = 6, delay: float = 0.5) -> Any:
    for i in range(attempts):
        try:
            return fn()
        except OSError as e:
            if not _transient_fs_error(e):
                raise
            if i == attempts - 1:
                record_recovery_attempt(f"파일 작업 {attempts}회 재시도 실패 ({type(e).__name__})")
                raise
            time.sleep(delay * (i + 1))


_STAMP = re.compile(r"_\d{8}_\d{6}$")


def timestamped(path: str) -> str:
    """<name>_<YYYYMMDD_HHMMSS><ext> (an existing stamp is replaced, never doubled)."""
    stem, ext = os.path.splitext(path)
    return f"{_STAMP.sub('', stem)}_{datetime.now():%Y%m%d_%H%M%S}{ext}"


def save_with_lock_fallback(path: str, writer: Callable[[str], None], what: str = "파일") -> str:
    """writer(path); a locked file (open in Excel) is saved under a timestamped name instead (R4). Returns the real path."""
    try:
        retry_fs(lambda: writer(path), attempts=2, delay=0.2)
        return path
    except PermissionError:
        alt = timestamped(path)
        writer(alt)
        note_recovered(f"{what} {os.path.basename(path)} 잠김(Excel 등에서 열림) -> {os.path.basename(alt)}로 저장")
        return alt


def read_text_file(path: str) -> str:
    """Text written by any common Windows tool (UTF-8 +/- BOM, UTF-16 BOM, CP949/EUC-KR)."""
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise InputError(f"텍스트 파일 인코딩 인식 실패 (UTF-8/UTF-16/CP949 아님): {path}")


def open_workbooks() -> List[Dict[str, Any]]:
    """Workbooks open in Excel: [{"path": full path or "", "name", "saved": bool}]. Only the paths are read through COM
    (pandas-kit never edits an open document). [] when Excel is not running or this is not Windows."""
    if os.name != "nt":
        return []
    try:
        import win32com.client  # type: ignore
        app = win32com.client.GetActiveObject("Excel.Application")
        out = []
        for i in range(1, int(app.Workbooks.Count) + 1):
            wb = app.Workbooks(i)
            path = str(wb.Path or "")
            out.append({"path": str(wb.FullName) if path else "", "name": str(wb.Name), "saved": bool(wb.Saved) and bool(path)})
        return out
    except Exception:
        return []
