"""office-cli 공통 런타임 (pywin32 없이 import 가능).

모든 진입점(excel.py, word.py, ppt.py, outlook.py, reader.py, office.py)이 공유하는 기능:
- ACI 응답 계약: stdout 정확히 1줄 JSON, exit 0/1/2/130, agent_directive VERIFIED:/STOP_RETRY:/ERROR:
- 인자 파서(JsonArgumentParser): 인자 오류도 JSON으로 응답 (exit 2)
- 로그: stderr, 태그 [INFO] [PROGRESS] [WARN] [ERROR] [STATUS] (상세 로그는 OFFICE_CLI_VERBOSE=1일 때만)
- 경로: 설정(config.json), data 폴더, 작업 폴더(_work), 입력/출력 경로 해석
- 텍스트/CSV 입출력: 인코딩 자동 판별, 잠금 파일 대비 저장
- 결과 길이 제한, 이름 추천([추천]), 다음 명령(next_commands) 구성
- 오류 로그: 실패한 실행(exit 1, 2)을 <data>/logs/errors.txt에 자동 기록 (write_error_log)
"""
import argparse
import csv
import datetime as _dt
import difflib
import io
import json
import math
import os
import re
import sys
import time
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

TOOL_NAME = "office-cli"
VERSION = "1.4.12"                         # SKILL.md metadata.version과 같아야 함 (테스트로 확인)
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(TOOL_DIR, "config.json")

# 결과 길이 기본 상한 (경량 LLM 컨텍스트 보호)
DEFAULT_MAX_ROWS = 200
DEFAULT_MAX_CHARS = 8000
MAX_CELL_CHARS = 500        # JSON에 담는 셀·찾은 값 하나의 글자 상한 (전체는 --save 파일)
# 작업 폴더 파일 보관 기간 (일)
WORK_RETENTION_DAYS = 7
# 오류 로그가 이보다 커지면 errors.old.txt로 넘기고 새로 시작 (1세대만 보관)
ERROR_LOG_MAX = 2 * 1024 * 1024
# 실제 명령 인자. JsonArgumentParser.parse_args가 채웁니다 (인자 오류 기록에도 실제 명령을 남기려고)
CURRENT_ARGV: Optional[List[str]] = None
# JIT 복구가 한도를 넘겨 실패한 시도 (오류 로그의 '복구 시도:' 줄)
RECOVERY_ATTEMPTS: List[str] = []


# ==============================================================================
# 1. 콘솔과 로그
# ==============================================================================

def setup_console() -> None:
    """stdout/stderr를 UTF-8로 고정합니다 (Windows 콘솔과 파이프의 기본값은 CP949)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and (stream.encoding or "").lower() != "utf-8":
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _verbose() -> bool:
    return os.environ.get("OFFICE_CLI_VERBOSE", "") not in ("", "0", "false", "False")


def log(tag: str, message: str) -> None:
    """stderr 로그. tag는 INFO, PROGRESS, WARN, ERROR, STATUS 중 하나."""
    try:
        sys.stderr.write(f"[{tag}] {message}\n")
        sys.stderr.flush()
    except Exception:
        pass


def debug(message: str) -> None:
    """상세 로그. OFFICE_CLI_VERBOSE=1일 때만 [INFO] 태그로 출력합니다."""
    if _verbose():
        log("INFO", message)


# ==============================================================================
# 2. 응답 계약
# ==============================================================================

class InputError(ValueError):
    """인자/대상/데이터 문제. 같은 인자로 재시도해도 소용없음 (STOP_RETRY).
    candidates를 주면 '후보가 여러 개라 도구가 고를 수 없음'으로 응답합니다 (R8 [후보], 사용자가 고름)."""

    def __init__(self, message: str = "", candidates: Optional[List[str]] = None, no_fallback: bool = False,
                 *, directive: Optional[str] = None, error_type: Optional[str] = None,
                 fields: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.candidates = list(candidates or [])
        # 선택: STOP_RETRY: 뒤에 붙일 전용 안내, 전용 error_type, 응답에 더할 키 (예: next_commands)
        self.directive = directive
        self.error_type = error_type
        self.fields = dict(fields or {})
        # 대상(문서, 앵커) 자체의 문제라 다른 방식으로 대체해도 소용없음 (transfer의 방식 자동 전환을 멈춤)
        self.no_fallback = no_fallback or bool(self.candidates)


MAX_CANDIDATES = 10


def ambiguous(message: str, candidates: Iterable[str]) -> InputError:
    """후보가 둘 이상이라 고를 수 없음. message 끝에 '[후보] a, b, c'(최대 10개)를 붙입니다."""
    cands = [str(c) for c in candidates]
    shown = cands[:MAX_CANDIDATES]
    tail = f" [후보] {', '.join(shown)}" + (f" 외 {len(cands) - len(shown)}개" if len(cands) > len(shown) else "")
    return InputError(message + tail, candidates=shown)


class UserDecisionError(InputError):
    """도구가 대신 정하면 안 되는 일 (예: 없는 시트를 새로 만들지). 재시도하지 않고 사용자에게 묻게 합니다.
    message에는 사실만 적고, 할 일(directive)은 사용자 동의를 조건으로 적습니다. 경량 LLM이 message의 명령문을 그대로 따르기 때문."""

    def __init__(self, message: str, directive: str, candidates: Optional[List[str]] = None):
        cands = [str(c) for c in (candidates or [])][:MAX_CANDIDATES]
        super().__init__(message + (f" [후보] {', '.join(cands)}" if cands else ""), candidates=cands, no_fallback=True)
        self.directive = directive


class OfficeBusyError(RuntimeError):
    """Office가 편집 중이거나 대화상자가 떠 있어 응답하지 않음 (STOP_RETRY, 사용자 조치 필요)."""


# R5: 일반 ValueError/KeyError는 코드 버그일 수 있어 ERROR:(내부 오류, traceback 기록). 입력 문제는 경계에서 InputError로 바꿉니다.
_STOP_RETRY_ERRORS = (InputError, OfficeBusyError, FileNotFoundError, NotADirectoryError, IsADirectoryError,
                      PermissionError)

# COM HRESULT -> 사용자 안내 (pywintypes.com_error 해석용)
_COM_HINTS = {
    0x800AC472: ("busy", "Excel이 셀 편집 중이거나 대화상자가 열려 있습니다. Esc를 누르거나 대화상자를 닫은 뒤 다시 요청하세요."),
    0x8001010A: ("busy", "Office가 다른 작업으로 바빠 응답하지 않습니다. 대화상자가 떠 있는지 확인하세요."),
    0x80010001: ("busy", "Office가 호출을 거부했습니다. 대화상자가 떠 있는지 확인하세요."),
    0x800706BA: ("gone", "Office 프로그램과의 연결이 끊겼습니다 (창이 닫혔거나 응답 없음)."),
    0x800706BE: ("gone", "Office 프로그램과의 연결이 끊겼습니다 (원격 호출 실패)."),
    0x800401E3: ("gone", "실행 중인 Office 창이나 열린 문서를 찾을 수 없습니다."),
}


def _com_error_info(exc: BaseException) -> Optional[Tuple[int, str, str]]:
    """pywintypes.com_error면 (HRESULT, Office 설명, 분류)를 돌려줍니다."""
    if type(exc).__name__ != "com_error":
        return None
    args = getattr(exc, "args", ())
    if not args or not isinstance(args[0], int):
        return None
    hr = args[0] & 0xFFFFFFFF
    desc = str(args[1]) if len(args) > 1 else ""
    excinfo = args[2] if len(args) > 2 else None
    scode = None
    if excinfo and len(excinfo) >= 6:
        if excinfo[2]:
            desc = str(excinfo[2])
        if isinstance(excinfo[5], int):
            scode = excinfo[5] & 0xFFFFFFFF
    for code in (scode, hr):
        if code in _COM_HINTS:
            kind, hint = _COM_HINTS[code]
            return code, f"{hint} (원문: {desc})", kind
    return hr, desc, "other"


def json_safe(obj: Any) -> Any:
    """COM/파이썬 값을 엄격한 JSON 값으로 바꿉니다 (날짜, NaN, 튜플, 바이트 등)."""
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [json_safe(v) for v in obj]
    if isinstance(obj, bool) or obj is None or isinstance(obj, (int, str)):
        return obj
    if isinstance(obj, float):
        return obj if math.isfinite(obj) else None
    if isinstance(obj, (_dt.datetime, _dt.date, _dt.time)):
        return obj.isoformat()
    if isinstance(obj, bytes):
        return obj.decode("utf-8", errors="replace")
    if hasattr(obj, "isoformat"):  # pywintypes.datetime
        try:
            return obj.isoformat()
        except Exception:
            pass
    try:
        f = float(obj)
        return f if math.isfinite(f) else None
    except Exception:
        return str(obj)


def emit_json(payload: Dict[str, Any]) -> None:
    """stdout에 정확히 1줄 JSON을 출력합니다."""
    setup_console()
    sys.stdout.write(json.dumps(json_safe(payload), ensure_ascii=False) + "\n")
    sys.stdout.flush()


def ok(tool: str, command: str, preview: str, directive: str = "", **fields: Any) -> Dict[str, Any]:
    """성공 응답. directive를 생략하면 'VERIFIED: 완료' 문구를 씁니다."""
    if fields.get("verified") is False:
        # R1: 확인하지 못한 결과에 VERIFIED:를 붙이지 않음 (무엇을 확인하지 못했는지 ok_unverified로 적어야 함)
        raise ValueError("ok()에 verified=False를 넘길 수 없습니다. 확인하지 못했으면 ok_unverified를 쓰세요.")
    payload: Dict[str, Any] = {"status": "ok", "tool": tool, "command": command, "preview": preview}
    payload.update(fields)
    if not directive:     # 읽기·조회 명령: 읽은 결과 자체가 확인
        directive = "VERIFIED: 작업이 완료되었습니다. 같은 명령을 다시 실행하거나 결과를 다시 조회하지 마세요."
    elif not directive.startswith("VERIFIED:"):
        # R1: VERIFIED:를 대신 붙이지 않음. 확인하지 못한 결과는 ok_unverified를 써야 함 (개발 중 실수를 잡는 내부 오류)
        raise ValueError(f"ok()의 directive는 'VERIFIED:'로 시작해야 합니다 (확인하지 못했으면 ok_unverified): {directive[:40]}")
    payload.setdefault("verified", True)
    payload["agent_directive"] = directive
    _attach_recovered(payload)
    return payload


RECOVERED: List[str] = []


def note_recovered(msg: str) -> None:
    """JIT 복구로 결과가 달라진 일을 알립니다 (R3): stderr [WARN] + 응답의 recovered 목록."""
    text = str(msg)[:200]
    log("WARN", text)
    if text not in RECOVERED:
        RECOVERED.append(text)


def _attach_recovered(payload: Dict[str, Any]) -> None:
    if RECOVERED:
        payload["recovered"] = list(RECOVERED)


def check_saved_file(path: str, expected_rows: Optional[int] = None, delimiter: Optional[str] = None) -> Optional[str]:
    """저장한 파일을 형식에 맞게 다시 열어 봅니다 (R1). 문제가 없으면 None, 있으면 확인하지 못한 내용.
    delimiter: 이 도구가 쓴 CSV는 ","를 줌 (구분자 추측은 '| ' 같은 칸 때문에 다른 구분자로 읽어 행 수가 틀릴 수 있음)."""
    try:
        size = os.path.getsize(path)
    except OSError:
        return "저장한 파일이 없음"
    if size <= 0:
        return "저장한 파일이 비어 있음 (0바이트)"
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            try:
                if not wb.sheetnames:
                    return "저장한 엑셀 파일에 시트가 없음"
            finally:
                wb.close()
        elif ext == ".docx":
            import docx
            docx.Document(path)
        elif ext in (".pptx", ".pptm"):        # python-pptx는 매크로 포함 pptm도 엶 (docm은 python-docx가 열지 못함)
            import pptx
            pptx.Presentation(path)
        elif ext in (".csv", ".tsv"):
            rows = [r for r in read_csv_rows(path, delimiter) if any(str(v).strip() for v in r)]
            if expected_rows is not None and len(rows) != expected_rows:
                return f"저장한 CSV의 행 수가 다름 (기대 {expected_rows}행, 읽은 {len(rows)}행)"
        elif ext in (".md", ".txt", ".json"):
            read_text_file(path)
        elif ext == ".pdf":
            with open(path, "rb") as f:
                if not f.read(5).startswith(b"%PDF"):
                    return "저장한 PDF 파일의 머리가 PDF 형식이 아님"
        else:
            return f"{ext or '확장자 없는'} 파일은 다시 열어 확인하는 방법이 없음 (크기 {size}바이트만 확인)"
    except ImportError as e:
        return f"다시 열어 볼 라이브러리가 없음 ({e.name or e})"
    except Exception as e:
        return f"저장한 파일을 다시 열지 못함 ({type(e).__name__}: {str(e)[:80]})"
    return None


def ok_file(tool: str, command: str, preview: str, path: str, directive: str,
            expected_rows: Optional[int] = None, delimiter: Optional[str] = None, **fields: Any) -> Dict[str, Any]:
    """저장한 파일 결과 (R1): 형식에 맞게 다시 열리면(CSV는 expected_rows를 주면 행 수까지) VERIFIED, 아니면 UNVERIFIED."""
    problem = check_saved_file(path, expected_rows, delimiter)
    if problem is None:
        return ok(tool, command, preview, directive, verified=True, **fields)
    return ok_unverified(tool, command, preview, [problem], **fields)


def ok_unverified(tool: str, command: str, preview: str, unverified: List[str], **fields: Any) -> Dict[str, Any]:
    """실행은 됐지만 결과를 다시 읽어 확인하지 못한 성공 응답 (R1). LLM은 재시도하지 않고 unverified를 그대로 알립니다."""
    payload: Dict[str, Any] = {"status": "ok", "tool": tool, "command": command, "preview": preview}
    payload.update(fields)
    payload.update(verified=False, unverified=list(unverified))
    payload["agent_directive"] = ("UNVERIFIED: 실행했지만 확인하지 못했습니다: " + "; ".join(unverified)
                                  + ". 다시 실행하지 말고 사용자에게 그대로 알리세요.")
    _attach_recovered(payload)
    return payload


def error_payload(exc: BaseException, tool: str = "", command: str = "") -> Dict[str, Any]:
    """예외를 표준 오류 응답으로 바꿉니다."""
    base = _error_payload(exc, tool, command)
    _attach_recovered(base)
    return base


def _error_payload(exc: BaseException, tool: str = "", command: str = "") -> Dict[str, Any]:
    base: Dict[str, Any] = {"status": "error", "tool": tool, "command": command}
    if isinstance(exc, KeyboardInterrupt):
        base.update(error_type="interrupted",
                    message="사용자 또는 상위 프로세스에 의해 실행이 중단되었습니다.",
                    agent_directive="STOP_RETRY: 실행이 중단되었습니다. 다시 실행하기 전에 사용자에게 확인하세요.")
        return base

    com = _com_error_info(exc)
    if com is not None:
        code, desc, kind = com
        if kind in ("busy", "gone"):
            base.update(error_type="office_busy" if kind == "busy" else "office_unavailable",
                        message=desc, hresult=f"0x{code:08X}",
                        agent_directive="STOP_RETRY: Office가 응답하지 않습니다. message를 사용자에게 그대로 보고하세요.")
        else:
            base.update(error_type="office", message=f"Office 오류: {desc}", hresult=f"0x{code:08X}",
                        agent_directive="STOP_RETRY: Office 작업이 실패했습니다. 인자(범위, 시트, 슬라이드 번호 등)를 확인할 수 있으면 "
                                        "1회만 고쳐서 다시 실행하고, 아니면 message를 사용자에게 보고하세요.")
        return base

    msg = str(exc) or type(exc).__name__
    if isinstance(exc, UserDecisionError):
        base.update(error_type="needs_user_decision", message=msg, agent_directive=f"STOP_RETRY: {exc.directive}")
        if exc.candidates:
            base["candidates"] = exc.candidates
        return base
    if isinstance(exc, InputError) and exc.candidates:
        base.update(error_type="ambiguous", message=msg, candidates=exc.candidates,
                    agent_directive="STOP_RETRY: 후보가 여러 개라 도구가 고를 수 없습니다. 재시도하지 말고 message의 "
                                    "[후보]를 사용자에게 보여주고 고르게 한 뒤, 사용자가 고른 값으로 실행하세요.")
        return base
    if isinstance(exc, InputError) and (exc.directive or exc.error_type or exc.fields):
        base.update(error_type=exc.error_type or "input", message=msg,
                    agent_directive="STOP_RETRY: " + (exc.directive or "인자나 대상 확인에 실패했습니다. message를 사용자에게 그대로 보고하세요."))
        for k, v in exc.fields.items():
            base.setdefault(k, v)
        return base
    if isinstance(exc, _STOP_RETRY_ERRORS):
        base.update(error_type="office_busy" if isinstance(exc, OfficeBusyError) else "input",
                    message=msg,
                    agent_directive="STOP_RETRY: 인자나 대상 확인에 실패했습니다. message에 '[추천]'이 있으면 그 값으로 1회만 "
                                    "다시 실행하고, 없으면 message를 사용자에게 그대로 보고하세요.")
        return base
    base.update(error_type="internal", message=f"{type(exc).__name__}: {msg}",
                agent_directive="ERROR: 예상하지 못한 내부 오류입니다. 재시도하지 말고 message를 사용자에게 보고하세요.")
    return base


def record_recovery_attempt(msg: str) -> None:
    """JIT 복구가 한도를 넘겨 실패했을 때 무엇을 몇 번 시도했는지 남깁니다 (오류 로그의 '복구 시도:' 줄)."""
    try:
        RECOVERY_ATTEMPTS.append(str(msg)[:300])
    except Exception:
        pass


def error_log_path() -> str:
    """<데이터 폴더>/logs/errors.txt (데이터 폴더: OFFICE_CLI_DATA_DIR > config.json data_dir > <도구폴더>/data)"""
    return os.path.join(get_data_dir(), "logs", "errors.txt")


# 오류 로그에 값을 남기지 않는 옵션 (본문·받는 사람·검색어 등 문서 내용과 개인정보). 경로와 그 밖의 인자는 그대로 남깁니다.
REDACT_OPTIONS = {"--body", "--text", "--subject", "--to", "--cc", "--bcc", "--from", "--query", "--find",
                  "--replace", "--value", "--values", "--formula", "--header", "--title", "--cards", "--anchor",
                  "--search", "--new-name", "--map"}
# 위 옵션의 짧은 이름 (모든 office-cli 파서에서 모음, 테스트로 확인): reader/outlook -q = --query
REDACT_SHORT = {"-q": "--query"}
# 도구마다 같은 이름이 내용이 아닌 옵션 (office.py transfer --from excel --to word, ppt.py slide --to 3)
REDACT_KEEP = {"office": {"--from", "--to"}, "ppt": {"--to"}}


_REDACTED = re.compile(r"^(?:-{1,2}[\w-]+=)?<\d+자 생략>$")


def _hidden_option(opt: str, hidden: set) -> bool:
    """내용 옵션이면 True. 줄여 쓴 이름(--que)도 내용 옵션의 앞부분이면 가립니다
    (파서는 줄임을 받지 않지만, 그 인자 오류도 오류 로그에 남으므로)."""
    if opt in hidden:
        return True
    return len(opt) >= 3 and any(h.startswith(opt) for h in hidden)


def redact_argv(argv: Sequence[Any], tool: str = "") -> List[str]:
    """내용 옵션의 값을 '<N자 생략>'으로 바꿉니다 (AGENTS.md: 오류 로그에 본문·개인정보를 남기지 않음).
    '--opt 값', '--opt=값', '-q 값', '-q값', 줄여 쓴 '--que 값', 값이 여러 개인 옵션(--formula a b)을 모두 처리합니다."""
    hidden = REDACT_OPTIONS - REDACT_KEEP.get(tool, set())
    shorts = {k for k, v in REDACT_SHORT.items() if v in hidden}
    out: List[str] = []
    hiding = False
    for a in [str(x) for x in argv]:
        if a.startswith("--"):
            opt, eq, val = a.partition("=")
            secret = _hidden_option(opt, hidden)
            hiding = secret and not eq
            out.append(f"{opt}=<{len(val)}자 생략>" if eq and secret else a)
            continue
        short = a[:2]
        if a.startswith("-") and not a.startswith("--") and short in shorts:
            hiding = len(a) == 2
            if len(a) > 2:                        # -q값, -q=값 (붙여 쓴 값)
                val = a[3:] if a[2] == "=" else a[2:]
                out.append(f"{short}=<{len(val)}자 생략>")
            else:
                out.append(a)
            continue
        if hiding and not (a.startswith("-") and len(a) == 2 and a[1].isalpha()):
            out.append(f"<{len(a)}자 생략>")
            continue
        hiding = False
        out.append(a)
    return out


def write_error_log(tool: str, argv: Optional[Sequence[str]], payload: Dict[str, Any], trace: str = "") -> None:
    """실패한 실행을 <데이터 폴더>/logs/errors.txt에 덧붙입니다 (사람과 개발 에이전트의 사후 분석용).
    기준 구현: archive/doc-classifier/_dc_runtime.py. 경로와 인자만 기록하고 문서 본문은 기록하지 않습니다.
    회전 실패와 기록 실패를 따로 감싸고, 어떤 오류도 삼킵니다 (stdout 1줄 JSON 계약을 깨면 안 됨)."""
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
            argv = CURRENT_ARGV if CURRENT_ARGV is not None else sys.argv[1:]
        script = f"{tool}.py" if tool else os.path.basename(sys.argv[0] or "")
        args = " ".join([script] + [a if _REDACTED.match(a) else quote_arg(a) for a in redact_argv(argv, tool)]).strip()
        lines = [f"===== {_dt.datetime.now().isoformat(timespec='seconds')} | {TOOL_NAME} {VERSION} | "
                 f"Python {platform.python_version()} | {platform.platform()}",
                 f"명령: {args}",
                 f"결과: {payload.get('error_type')} | {payload.get('message')}"]
        if RECOVERY_ATTEMPTS:
            lines.append("복구 시도: " + "; ".join(RECOVERY_ATTEMPTS))
        if trace:
            lines.append("traceback:\n" + trace.rstrip())
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n\n")
    except Exception:
        pass


def run_cli(tool: str, command: str, execute: Callable[[], Dict[str, Any]]) -> None:
    """도구 본문을 실행하고, 어떤 경우에도 stdout 1줄 JSON과 종료 코드(0/1/130)를 보장합니다.
    실패(exit 1)는 <데이터 폴더>/logs/errors.txt에 자동 기록합니다 (사용자 중단 130과 성공은 기록하지 않음)."""
    setup_console()
    try:
        result = execute()
    except KeyboardInterrupt as e:
        log("ERROR", "실행 중단 (KeyboardInterrupt)")
        emit_json(error_payload(e, tool, command))
        sys.exit(130)
    except SystemExit:
        raise
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        log("ERROR", f"{type(e).__name__}: {e}")
        if _verbose():
            sys.stderr.write(trace)
        payload = error_payload(e, tool, command)
        # traceback은 내부 오류(ERROR:)일 때만 남깁니다
        internal = str(payload.get("agent_directive", "")).startswith("ERROR:")
        write_error_log(tool, None, payload, trace if internal else "")
        emit_json(payload)
        sys.exit(1)
    if result.get("status") != "ok":
        write_error_log(tool, None, result)
    emit_json(result)
    sys.exit(0 if result.get("status") == "ok" else 1)


# 경량 LLM이 자주 지어내는 명령·옵션 (회사 PC 오류 로그). 실제 이름으로 안내합니다.
COMMAND_ALIASES = {"excel": {"list": "info", "sheets": "info"}, "word": {"list": "info", "toc": "outline"},
                   "ppt": {"list": "info"}, "office": {"list": "windows", "redo": "undo"}}
OPTION_SYNONYMS = {"--target": ["--app", "--file", "--name"], "--file": ["--target"], "--path": ["--file", "--target"],
                   "--output": ["--save"], "--out": ["--save"], "--sheet-name": ["--sheet", "--name"]}


class JsonArgumentParser(argparse.ArgumentParser):
    """인자 오류를 텍스트 대신 JSON(exit 2)으로 응답하는 ArgumentParser. 서브파서에도 그대로 상속됩니다."""

    tool_name = ""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        # 줄여 쓴 옵션(--que = --query)을 받지 않음: 지어낸 옵션이 다른 옵션으로 해석되는 일과,
        # 오류 로그에서 내용 옵션을 가리지 못하는 일을 막음 (서브파서도 이 클래스로 만들어짐)
        kwargs.setdefault("allow_abbrev", False)
        super().__init__(*args, **kwargs)

    def parse_args(self, args=None, namespace=None):  # type: ignore[override]
        # 최상위 파서만 parse_args를 부릅니다 (서브파서는 parse_known_args). 오류 로그에 남길 실제 인자를 기억합니다.
        global CURRENT_ARGV
        CURRENT_ARGV = [str(a) for a in args] if args is not None else sys.argv[1:]
        return super().parse_args(args, namespace)

    def print_help(self, file=None) -> None:  # type: ignore[override]
        """R7: 도움말 텍스트는 stderr로, stdout에는 JSON 1줄 (argparse가 이어서 exit 0)."""
        super().print_help(sys.stderr)
        emit_json({"status": "ok", "tool": self.tool_name, "command": "help", "verified": True,
                   "skill_file": os.path.join(os.path.dirname(os.path.abspath(__file__)), "SKILL.md"),
                   "agent_directive": "VERIFIED: 도움말을 stderr에 출력했습니다. SKILL.md의 명령 템플릿을 그대로 사용하세요. "
                                      "도구 코드(.py)는 읽지 마세요. 방법을 모르겠으면 사용자에게 물어보세요."})

    def _hint(self, message: str) -> str:
        """흔한 잘못(없는 명령, 다른 명령의 옵션, 따옴표 없는 경로)에 [추천]을 붙입니다 (회사 PC 오류 로그)."""
        import difflib
        top = self if self.prog.count(" ") == 0 else None
        subs = {}
        for parser in ([top] if top else []):
            for act in parser._actions:
                if isinstance(act, argparse._SubParsersAction):
                    subs = act.choices
        m = re.search(r"invalid choice: '([^']+)'", message)
        if m:
            choices = list(subs) or re.findall(r"'([^']+)'", message.split("choose from", 1)[-1])
            alias = COMMAND_ALIASES.get(self.tool_name, {}).get(m.group(1).lower())
            best = alias or next(iter(difflib.get_close_matches(m.group(1), choices, n=1, cutoff=0.6)), None)
            return f" [추천] {best}" if best else f" [후보] {', '.join(choices)}"
        m = re.search(r"unrecognized arguments: (.+)$", message)
        if not m:
            return ""
        cmd = CURRENT_ARGV[0] if CURRENT_ARGV else ""
        sub = subs.get(cmd) if subs else self
        opts = [o for a in (sub._actions if sub else []) for o in a.option_strings if o.startswith("--")]
        extra = m.group(1).split()
        flags = [x for x in extra if x.startswith("--")]
        for f in flags:
            near = [o for o in OPTION_SYNONYMS.get(f, []) if o in opts] or difflib.get_close_matches(f, opts, n=1, cutoff=0.7)
            if near:
                return f" '{cmd}' 명령에는 {f} 옵션이 없습니다. [추천] {near[0]}"
        if not flags:
            return " 띄어쓰기가 있는 경로나 값은 큰따옴표로 감싸세요 (예: --target \"C:\\폴더 이름\\파일.xlsx\")."
        opts = [o for o in opts if o != "--help"]
        return f" '{cmd}' 명령의 옵션: {', '.join(opts)}" if opts else ""

    def error(self, message: str) -> None:  # type: ignore[override]
        self.print_usage(sys.stderr)
        payload = {
            "status": "error",
            "tool": self.tool_name,
            "command": "",
            "error_type": "input",
            "message": f"CLI 인자 오류: {message}." + self._hint(message),
            "agent_directive": "STOP_RETRY: CLI 인자가 잘못되었습니다. message에 [추천]이 있으면 그 부분만 바꿔 1번만 다시 실행하고, "
                               "없으면 SKILL.md의 명령 형식을 확인해 사용자에게 보고하세요.",
        }
        write_error_log(self.tool_name, None, payload)
        emit_json(payload)
        sys.exit(2)


def build_parser(tool: str, description: str):
    """서브커맨드 방식 파서를 만듭니다. 반환: (parser, subparsers)."""
    JsonArgumentParser.tool_name = tool
    parser = JsonArgumentParser(prog=f"{tool}.py", description=description)
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    subparsers.required = True
    return parser, subparsers


# ==============================================================================
# 3. 설정과 폴더
# ==============================================================================

_config_cache: Optional[Dict[str, Any]] = None


def load_config() -> Dict[str, Any]:
    """도구 폴더의 config.json (없으면 빈 설정). 프로세스당 한 번만 읽습니다."""
    global _config_cache
    if _config_cache is None:
        cfg: Dict[str, Any] = {}
        if os.path.exists(CONFIG_FILE):
            try:
                cfg = json.loads(read_text_file(CONFIG_FILE))
                if not isinstance(cfg, dict):
                    cfg = {}
            except Exception as e:
                log("WARN", f"config.json을 읽지 못해 기본 설정을 사용합니다: {e}")
                cfg = {}
        _config_cache = cfg
    return _config_cache


def set_config(key: str, value: Any) -> None:
    """설정 값을 저장합니다 (예: transfer 방식 자동 판별 결과)."""
    cfg = dict(load_config())
    cfg[key] = value
    safe_write_text(CONFIG_FILE, json.dumps(cfg, ensure_ascii=False, indent=2))
    global _config_cache
    _config_cache = cfg


def _dir_setting(env_name: str, cfg_key: str, default: str) -> str:
    path = os.environ.get(env_name) or load_config().get(cfg_key) or default
    path = os.path.abspath(os.path.expandvars(os.path.expanduser(str(path))))
    os.makedirs(path, exist_ok=True)
    return path


def get_data_dir() -> str:
    """사용자 데이터 폴더. 우선순위: OFFICE_CLI_DATA_DIR > config.json data_dir > <도구폴더>/data"""
    return _dir_setting("OFFICE_CLI_DATA_DIR", "data_dir", os.path.join(TOOL_DIR, "data"))


def get_cache_dir() -> str:
    """캐시 폴더 (읽은 문서 텍스트, 색인, 스냅샷). 우선순위: OFFICE_CLI_CACHE_DIR > config.json cache_dir > <도구폴더>/cache"""
    return _dir_setting("OFFICE_CLI_CACHE_DIR", "cache_dir", os.path.join(TOOL_DIR, "cache"))


def get_work_dir() -> str:
    """중간 결과 파일 폴더 (data/_work). 오래된 파일은 자동 정리합니다."""
    path = os.path.join(get_data_dir(), "_work")
    os.makedirs(path, exist_ok=True)
    return path


def cleanup_work_dir(days: int = WORK_RETENTION_DAYS) -> int:
    """작업 폴더에서 보관 기간이 지난 파일을 지웁니다. 지운 개수를 반환 (실패는 무시)."""
    removed = 0
    cutoff = time.time() - days * 86400
    try:
        work = get_work_dir()
        for name in os.listdir(work):
            p = os.path.join(work, name)
            try:
                if os.path.isfile(p) and os.path.getmtime(p) < cutoff:
                    os.remove(p)
                    removed += 1
            except OSError:
                continue
    except OSError:
        pass
    return removed


def timestamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def resolve_input_path(path: str, kind: str = "입력 파일") -> str:
    """입력 파일 경로를 찾습니다. 순서: 절대경로/현재 폴더 -> data 폴더 -> 작업 폴더.
    없으면 data 폴더의 비슷한 파일명을 [추천]으로 붙여 InputError를 냅니다."""
    if not path:
        raise InputError(f"{kind} 경로가 비어 있습니다.")
    raw = os.path.expandvars(os.path.expanduser(str(path).strip().strip('"')))
    candidates = [raw]
    if not os.path.isabs(raw):
        candidates += [os.path.join(get_data_dir(), raw), os.path.join(get_work_dir(), raw)]
    for c in candidates:
        if os.path.isfile(c):
            return os.path.abspath(c)
    names: List[str] = []
    for folder in (get_data_dir(), get_work_dir()):
        try:
            names += [n for n in os.listdir(folder) if os.path.isfile(os.path.join(folder, n))]
        except OSError:
            pass
    raise not_found(kind, os.path.basename(raw), names, extra=f" (찾은 위치: {', '.join(candidates)})")


def resolve_output_path(path: Optional[str], default_stem: str, ext: str, overwrite: bool = False) -> str:
    """출력 파일 경로를 정합니다.
    - 생략하면 작업 폴더에 '<default_stem>_<시각><ext>'로 만듭니다.
    - 파일명만 주면 data 폴더 아래에 만듭니다.
    - 확장자가 없거나 다르면 ext로 맞춥니다.
    - 그 파일이 이미 있으면 덮어쓰지 않고 시각을 붙인 새 이름으로 바꿉니다 (R4, recovered로 알림). overwrite=True면 그대로.
    """
    ext = ext if ext.startswith(".") else "." + ext
    if not path:
        cleanup_work_dir()
        return os.path.join(get_work_dir(), f"{default_stem}_{timestamp()}{ext}")
    raw = os.path.expandvars(os.path.expanduser(str(path).strip().strip('"')))
    base, cur_ext = os.path.splitext(raw)
    if cur_ext.lower() != ext.lower():
        raw = (base if cur_ext else raw) + ext
    if not os.path.isabs(raw) and not os.path.dirname(raw):
        raw = os.path.join(get_data_dir(), raw)
    raw = os.path.abspath(raw)
    os.makedirs(os.path.dirname(raw), exist_ok=True)
    if os.path.exists(raw) and not overwrite:
        alt = _timestamped(raw)
        note_recovered(f"{os.path.basename(raw)} 이미 있음 -> 덮어쓰지 않고 {os.path.basename(alt)}로 저장")
        raw = alt
    return raw


# ==============================================================================
# 4. 파일 입출력
# ==============================================================================

_WIN_PATH = re.compile(r"(?:[A-Za-z]:\\|\\\\)[^\s\"'<>|]*")


def unescape_newlines(text: str) -> str:
    """--text의 글자 '\\n'을 줄바꿈으로. Windows 경로 안의 \\n(C:\\new\\data)은 그대로 둡니다."""
    out, pos = [], 0
    for m in _WIN_PATH.finditer(text):
        path, tail = m.group(0), ""
        cut = path.rfind("\\n")
        # 경로가 파일 이름(확장자)으로 끝났거나, \n 뒤가 비었거나 목록·제목 기호뿐이면 경로가 끝나고 줄이 바뀐 것
        # ("C:\\a\\b.xlsx\\n다음", "C:\\data\\n- 다음 줄")
        if cut > 2 and "\\" not in path[cut + 2:] and (re.fullmatch(r"[-#*>|]*|\d+\.", path[cut + 2:])
                                                         or re.search(r"\.\w{1,5}$", path[:cut])):
            path, tail = path[:cut], "\n" + path[cut + 2:]
        out += [text[pos:m.start()].replace("\\n", "\n"), path, tail]
        pos = m.end()
    out.append(text[pos:].replace("\\n", "\n"))
    return "".join(out)


def read_text_file(path: str) -> str:
    """Windows 도구들이 쓰는 인코딩(UTF-8 +/- BOM, UTF-16 BOM, CP949/EUC-KR)을 자동 판별해 읽습니다.
    줄바꿈은 \\n 하나로 맞춥니다 (Windows 메모장의 CRLF가 Office에 넣을 때 빈 줄로 남지 않게)."""
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16").replace("\r\n", "\n")
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return raw.decode(enc).replace("\r\n", "\n")
        except UnicodeDecodeError:
            continue
    raise InputError(f"텍스트 파일 인코딩 인식 실패 (UTF-8/UTF-16/CP949 아님): {path}")


def read_csv_rows(path: str, delimiter: Optional[str] = None) -> List[List[str]]:
    """CSV/TSV를 2차원 문자열 리스트로 읽습니다 (delimiter를 주지 않으면 구분자 자동 판별)."""
    text = read_text_file(path)
    if delimiter:
        return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]
    sample = text[:4096]
    delimiter = ","
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        if "\t" in sample and "," not in sample:
            delimiter = "\t"
    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


def _transient_fs_error(e: OSError) -> bool:
    """잠깐 기다리면 풀릴 수 있는 오류: 접근 거부, 공유 위반, 잠금 위반 (winerror 5/32/33)."""
    return isinstance(e, PermissionError) or getattr(e, "winerror", None) in (5, 32, 33)


def retry_fs(fn: Callable[[], Any], attempts: int = 6, delay: float = 0.5) -> Any:
    """백신/동기화 프로그램의 일시적 잠금에 대비해 파일 작업을 재시도합니다.
    한도를 넘기면 오류 로그의 '복구 시도:' 줄에 남기고 마지막 예외를 올립니다."""
    last: Optional[BaseException] = None
    for i in range(attempts):
        try:
            return fn()
        except OSError as e:
            if not _transient_fs_error(e):
                raise                       # 없는 파일, 디스크 가득 참 등은 기다려도 소용없음
            last = e
            time.sleep(delay * (i + 1))
    assert last is not None
    name = os.path.basename(str(getattr(last, "filename", "") or ""))
    record_recovery_attempt(f"파일 작업 {name} {attempts}회 재시도 실패 ({type(last).__name__})".replace("  ", " "))
    raise last


def safe_write_text(path: str, text: str, encoding: str = "utf-8") -> None:
    """임시 파일에 쓴 뒤 교체해, 쓰다가 중단돼도 기존 파일이 깨지지 않게 합니다."""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = f"{path}.tmp_{os.getpid()}"
    with open(tmp, "w", encoding=encoding, newline="") as f:
        f.write(text)
    retry_fs(lambda: os.replace(tmp, path))


_STAMP_RE = re.compile(r"_\d{8}_\d{6}$")


def _timestamped(path: str) -> str:
    """'<이름>_<YYYYMMDD_HHMMSS><확장자>'. 이미 시각이 붙은 이름이면 바꿔 붙이고, 같은 이름이 있으면 번호를 더합니다."""
    base, ext = os.path.splitext(path)
    base = _STAMP_RE.sub("", base)
    cand = f"{base}_{timestamp()}{ext}"
    n = 2
    while os.path.exists(cand):
        cand = f"{base}_{timestamp()}_{n}{ext}"
        n += 1
    return cand


def save_with_lock_fallback(path: str, writer: Callable[[str], None], what: str = "파일") -> str:
    """writer(path)로 저장합니다 (R4). 잠겨 있으면 retry_fs로 기다려 보고, 그래도 Excel/Word 등에서 열려 있으면
    시각을 붙인 새 이름으로 저장하고 recovered로 알립니다. 실제 저장 경로를 반환합니다."""
    try:
        retry_fs(lambda: writer(path), attempts=3)
        return path
    except PermissionError:
        alt = _timestamped(path)
        writer(alt)
        note_recovered(f"{os.path.basename(path)} 잠김({what}) -> {os.path.basename(alt)}로 저장")
        return alt


def write_csv_rows(path: str, rows: Iterable[Sequence[Any]], encoding: str = "utf-8-sig") -> str:
    """표를 CSV로 저장합니다 (기본 utf-8-sig: Excel에서 한글이 깨지지 않음). 실제 저장 경로를 반환합니다."""
    data = [["" if v is None else v for v in row] for row in rows]

    def _write(p: str) -> None:
        with open(p, "w", encoding=encoding, newline="") as f:
            csv.writer(f).writerows(data)

    return save_with_lock_fallback(path, _write, "CSV 파일")


def write_text(path: str, text: str) -> str:
    """텍스트(마크다운 등)를 UTF-8로 저장합니다. 실제 저장 경로를 반환합니다."""
    def _write(p: str) -> None:
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
    return save_with_lock_fallback(path, _write, "텍스트 파일")


# ==============================================================================
# 5. LLM 친화 도우미
# ==============================================================================

def suggest(name: str, candidates: Iterable[str], n: int = 3) -> List[str]:
    """오타로 보이는 이름에 대해 비슷한 후보를 찾습니다 (대소문자/공백 무시, 부분 일치 우선)."""
    cands = [str(c) for c in candidates]
    key = str(name).strip().lower().replace(" ", "")
    if not key:
        return []
    norm = {c: c.strip().lower().replace(" ", "") for c in cands}
    exact = [c for c in cands if norm[c] == key]
    if exact:
        return exact[:n]
    partial = [c for c in cands if key in norm[c] or (norm[c] and norm[c] in key)]
    close = difflib.get_close_matches(key, list(norm.values()), n=n, cutoff=0.5)
    by_norm = {v: k for k, v in norm.items()}
    result: List[str] = []
    for c in partial + [by_norm[v] for v in close]:
        if c not in result:
            result.append(c)
    return result[:n]


def find_name(name: str, candidates: Iterable[str]) -> Optional[str]:
    """대소문자/앞뒤 공백을 무시하고 정확히 같은 이름을 찾습니다."""
    key = str(name).strip().lower()
    for c in candidates:
        if str(c).strip().lower() == key:
            return str(c)
    return None


def _same_name_key(s: Any) -> str:
    return re.sub(r"[\s_\-]+", "", str(s)).casefold()


def not_found(kind: str, name: str, candidates: Iterable[str], extra: str = "", vocab: bool = False) -> InputError:
    """'찾을 수 없음' InputError (R8).
    - 표기만 다른(대소문자·공백·_-) 이름이 하나뿐이면 [추천]: 뜻이 같아 LLM이 바로 1회 재시도해도 됩니다.
    - 부분·유사 일치는 뜻이 다를 수 있어(2024 -> 2024_1분기, 견적서.pdf -> 견적서_구버전.pdf) [후보]: 사용자가 고릅니다.
    - vocab=True(도구가 정한 옵션 값: 색 이름, 테마 등): 오타로 보이는 가까운 값 하나도 [추천] (사용자 데이터가 아니라 뜻이 바뀌지 않음).
    - 비슷한 것도 없으면 사용 가능한 이름 일부."""
    cands = [str(c) for c in candidates]
    key = _same_name_key(name)
    same = [c for c in cands if key and _same_name_key(c) == key]
    if vocab and not same:
        close = difflib.get_close_matches(key, [_same_name_key(c) for c in cands], n=2, cutoff=0.75)
        if len(close) == 1:
            same = [c for c in cands if _same_name_key(c) == close[0]][:1]
    if len(same) == 1:
        return InputError(f"{kind} '{name}'을(를) 찾을 수 없습니다.{extra} 표기만 다른 이름이 있습니다. [추천] {same[0]}")
    hints = same or suggest(name, cands)
    if hints:
        return ambiguous(f"{kind} '{name}'을(를) 찾을 수 없습니다.{extra} 비슷한 이름이 있지만 같은 것인지 도구가 정할 수 없습니다. "
                         "사용자에게 고르게 하세요.", hints)
    if cands:
        shown = ", ".join(cands[:15]) + (" ..." if len(cands) > 15 else "")
        tail = f" 사용 가능한 {kind}: {shown}"
    else:
        tail = ""
    return InputError(f"{kind} '{name}'을(를) 찾을 수 없습니다.{extra}{tail}")


def clamp_cell(value: Any, max_chars: int = MAX_CELL_CHARS) -> Tuple[Any, bool]:
    """긴 글 값 하나를 상한으로 자릅니다 (숫자 등 글이 아닌 값은 그대로). (값, 잘렸는지 여부)"""
    if isinstance(value, str) and len(value) > max_chars:
        return value[:max_chars], True
    return value, False


def clamp_rows(rows: List[List[Any]], max_rows: int = DEFAULT_MAX_ROWS,
               max_chars: int = DEFAULT_MAX_CHARS) -> Tuple[List[List[Any]], bool, int]:
    """행 수와 대략적인 글자 수 상한으로 표를 자릅니다. (잘린 표, 행이 잘렸는지, 글이 잘린 칸 수)
    칸마다 MAX_CELL_CHARS에서 자르고, 첫 행도 글자 수 상한을 넘으면 넘는 칸부터 뺍니다 (첫 행이 아주 큰 표도 상한 안)."""
    out: List[List[Any]] = []
    used = 0
    cells_cut = 0
    for row in rows:
        if len(out) >= max_rows:
            return out, True, cells_cut
        cut_row = []
        cut_here = 0
        for v in row:
            v, cut = clamp_cell(v)
            cut_here += cut
            cut_row.append(v)
        size = sum(len(str(v)) + 1 for v in cut_row)
        if out and used + size > max_chars:
            return out, True, cells_cut
        if not out and size > max_chars:          # 첫 행만으로 상한을 넘음: 들어가는 칸까지만
            keep, width = [], 0
            for v in cut_row:
                width += len(str(v)) + 1
                if keep and width > max_chars:
                    break
                keep.append(v)
            cells_cut += cut_here + len(cut_row) - len(keep)
            out.append(keep)
            return out, len(rows) > 1, cells_cut
        cells_cut += cut_here
        out.append(cut_row)
        used += size
    return out, False, cells_cut


def clamp_text(text: str, max_chars: int = DEFAULT_MAX_CHARS) -> Tuple[str, bool]:
    """글자 수 상한으로 텍스트를 자릅니다. (잘린 텍스트, 잘렸는지 여부)"""
    if len(text) <= max_chars:
        return text, False
    return text[:max_chars], True


def quote_arg(value: Any) -> str:
    """명령 문자열에 넣을 인자를 Windows cmd 기준으로 따옴표 처리합니다."""
    s = str(value)
    if s and all(ch not in s for ch in ' \t"&|<>^()%,;='):
        return s
    return '"' + s.replace('"', '\\"') + '"'


def next_command(script: str, args: Sequence[Any], purpose: str) -> Dict[str, str]:
    """next_commands 항목 하나를 만듭니다. script는 도구 폴더 기준 파일명 (예: 'office.py')."""
    parts = [f"python {quote_arg(os.path.join(TOOL_DIR, script))}"] + [quote_arg(a) for a in args]
    return {"purpose": purpose, "command": " ".join(parts)}


# ==============================================================================
# 6. 마크다운 블록 해석 (word.py / ppt.py 공통)
# ==============================================================================

_MD_BOLD = re.compile(r"\*\*(.+?)\*\*|__(.+?)__")


def split_bold(text: str) -> Tuple[str, List[Tuple[int, int]]]:
    """'가 **나** 다' -> ('가 나 다', [(2, 3)]). 굵게 표시 구간(시작, 끝 오프셋)을 돌려줍니다."""
    plain: List[str] = []
    spans: List[Tuple[int, int]] = []
    pos = 0
    length = 0
    for m in _MD_BOLD.finditer(text):
        before = text[pos:m.start()]
        plain.append(before)
        length += len(before)
        inner = m.group(1) if m.group(1) is not None else m.group(2)
        spans.append((length, length + len(inner)))
        plain.append(inner)
        length += len(inner)
        pos = m.end()
    plain.append(text[pos:])
    return "".join(plain), spans


def parse_markdown_table(lines: List[str]) -> List[List[str]]:
    rows: List[List[str]] = []
    for ln in lines:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if cells and all(re.fullmatch(r":?-{3,}:?", c) for c in cells if c):
            continue
        rows.append(cells)
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


def parse_markdown_blocks(text: str) -> List[Tuple[str, Any]]:
    """마크다운을 블록 목록으로 바꿉니다.
    ('heading', (레벨, 텍스트)), ('paragraph', 텍스트), ('bullets', [항목]), ('numbered', [항목]),
    ('table', [[셀]]), ('break', None)  # '---' 구분선
    """
    blocks: List[Tuple[str, Any]] = []
    para: List[str] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    def flush() -> None:
        if para:
            blocks.append(("paragraph", " ".join(s.strip() for s in para)))
            para.clear()

    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            flush()
            i += 1
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            flush()
            blocks.append(("heading", (len(m.group(1)), m.group(2).strip().rstrip("#").strip())))
            i += 1
            continue
        if s.startswith("|"):
            flush()
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i])
                i += 1
            rows = parse_markdown_table(tbl)
            if rows:
                blocks.append(("table", rows))
            continue
        if re.match(r"^[-*+]\s+", s):
            flush()
            items = []
            while i < len(lines) and re.match(r"^\s*[-*+]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*+]\s+", "", lines[i]).strip())
                i += 1
            blocks.append(("bullets", items))
            continue
        if re.match(r"^\d+[.)]\s+", s):
            flush()
            items = []
            while i < len(lines) and re.match(r"^\s*\d+[.)]\s+", lines[i]):
                items.append(re.sub(r"^\s*\d+[.)]\s+", "", lines[i]).strip())
                i += 1
            blocks.append(("numbered", items))
            continue
        if re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", s):
            flush()
            blocks.append(("break", None))
            i += 1
            continue
        para.append(s)
        i += 1
    flush()
    return blocks


def looks_like_markdown(text: str) -> bool:
    return bool(re.search(r"(?m)^\s*(#{1,6}\s|[-*+]\s|\d+[.)]\s|\|.*\|)", text) or "**" in text)
