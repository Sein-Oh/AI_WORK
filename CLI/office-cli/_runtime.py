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
VERSION = "1.4.68"                         # skills\<이름>\SKILL.md metadata.version과 같아야 함 (테스트로 확인)
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))


def engenius_home() -> Optional[str]:
    """도구 파일이 En-Genius 트레이의 임시 폴더(%TEMP%\\engenius-skill\\<id>)에서 실행 중이면 사용자 데이터를 둘 고정 폴더.
    트레이는 그 폴더를 실행이 끝나면 지우므로 도구 폴더 아래(data, cache, config.json)가 남지 않는다.
    배포 스킬(docs/engenius)은 C:\\CLI의 도구를 실행하므로 해당 없음: OpenCode와 같은 data, cache를 쓴다."""
    if not os.environ.get("ENGENIUS_OUT") or "engenius-skill" not in TOOL_DIR.replace("\\", "/").lower().split("/"):
        return None
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(home, "Documents", "En-Genius", "CLI-data", TOOL_NAME)


def _tool_home() -> str:
    """config.json, data, cache의 기본 위치: En-Genius 실행이면 고정 폴더, 아니면 도구 폴더."""
    return engenius_home() or TOOL_DIR


CONFIG_FILE = os.environ.get("OFFICE_CLI_CONFIG") or os.path.join(_tool_home(), "config.json")

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

# 오류 로그 message에서 문서에서 가져온 글(칸 값, 제목, 노트, 이름 목록 등)을 뺀 자리에 넣는 글
DOC_OMITTED = "<문서 내용 생략>"


class InputError(ValueError):
    """인자/대상/데이터 문제. 같은 인자로 재시도해도 소용없음 (STOP_RETRY).
    candidates를 주면 '후보가 여러 개라 도구가 고를 수 없음'으로 응답합니다 (R8 [후보], 사용자가 고름).
    log_message: 오류 로그에 남길 message (문서에서 가져온 글을 뺀 글). message에 문서 글을 넣는 곳은 반드시 줍니다.
    없으면 message를 쓰고, 어느 쪽이든 write_error_log가 가린 인자 값과 LOG_CUT_MARKERS를 한 번 더 뺍니다."""

    def __init__(self, message: str = "", candidates: Optional[List[str]] = None, no_fallback: bool = False,
                 *, directive: Optional[str] = None, error_type: Optional[str] = None,
                 fields: Optional[Dict[str, Any]] = None, log_message: Optional[str] = None):
        super().__init__(message)
        self.log_message = log_message
        self.candidates = list(candidates or [])
        # 선택: STOP_RETRY: 뒤에 붙일 전용 안내, 전용 error_type, 응답에 더할 키 (예: next_commands)
        self.directive = directive
        self.error_type = error_type
        self.fields = dict(fields or {})
        # 대상(문서, 앵커) 자체의 문제라 다른 방식으로 대체해도 소용없음 (transfer의 방식 자동 전환을 멈춤)
        self.no_fallback = no_fallback or bool(self.candidates)


MAX_CANDIDATES = 10


def _omitted_candidates(n: int) -> str:
    return f" [후보 {n}개 생략]" if n else ""


def ambiguous(message: str, candidates: Iterable[str], log_message: Optional[str] = None) -> InputError:
    """후보가 둘 이상이라 고를 수 없음. message 끝에 '[후보] a, b, c'(최대 10개)를 붙입니다.
    후보는 문서에서 가져온 글(제목, 단락, 시트·문서 이름)이라 오류 로그에는 개수만 남깁니다."""
    cands = [str(c) for c in candidates]
    shown = cands[:MAX_CANDIDATES]
    tail = f" [후보] {', '.join(shown)}" + (f" 외 {len(cands) - len(shown)}개" if len(cands) > len(shown) else "")
    return InputError(message + tail, candidates=shown,
                      log_message=(message if log_message is None else log_message) + _omitted_candidates(len(cands)))


class UserDecisionError(InputError):
    """도구가 대신 정하면 안 되는 일 (예: 없는 시트를 새로 만들지). 재시도하지 않고 사용자에게 묻게 합니다.
    message에는 사실만 적고, 할 일(directive)은 사용자 동의를 조건으로 적습니다. 경량 LLM이 message의 명령문을 그대로 따르기 때문."""

    def __init__(self, message: str, directive: str, candidates: Optional[List[str]] = None,
                 log_message: Optional[str] = None):
        cands = [str(c) for c in (candidates or [])][:MAX_CANDIDATES]
        super().__init__(message + (f" [후보] {', '.join(cands)}" if cands else ""), candidates=cands, no_fallback=True,
                         log_message=(message if log_message is None else log_message) + _omitted_candidates(len(cands)))
        self.directive = directive


ELEVATED_HINT = ("이 도구가 관리자 권한으로 실행 중이라, 일반 권한으로 연 Office 문서와 창을 볼 수 없습니다 "
                 "(Windows 권한 수준이 다르면 COM 연결이 막힘). 관리자가 아닌 일반 PowerShell 창에서 OpenCode를 다시 실행하세요.")


def is_elevated() -> bool:
    """Windows에서 관리자 권한(높은 무결성 수준)으로 실행 중인지. 판단하지 못하면 False."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def elevation_hint() -> str:
    """'열린 문서가 없음'을 알릴 때 붙일 안내. 관리자 권한으로 실행 중이면 ELEVATED_HINT, 아니면 빈 글."""
    return ELEVATED_HINT if is_elevated() else ""


class MissingLibraryError(InputError):
    """필요한 파이썬 라이브러리가 설치되어 있지 않음 (R5). LLM은 직접 설치하지 않고 install_command를 사용자에게 전달합니다."""

    def __init__(self, pip_name: str, feature: str):
        super().__init__(f"{feature}에 필요한 {pip_name}이(가) 설치되어 있지 않습니다. 설치 명령: python -m pip install {pip_name}")
        self.pip_name = pip_name
        self.feature = feature


def require_lib(module: str, pip_name: str, feature: str) -> Any:
    """module을 불러옵니다. 없으면 MissingLibraryError (STOP_RETRY, missing_library, install_command)."""
    import importlib
    try:
        return importlib.import_module(module)
    except ImportError:
        raise MissingLibraryError(pip_name, feature) from None


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
    if isinstance(exc, MissingLibraryError):
        base.update(error_type="missing_library", message=msg, missing_package=exc.pip_name,
                    install_command=f"python -m pip install {exc.pip_name}",
                    agent_directive=f"STOP_RETRY: {exc.feature}에 필요한 {exc.pip_name}이(가) 설치되어 있지 않습니다. "
                                    "install_command를 사용자에게 전달하세요. 직접 설치하지 마세요.")
        return base
    if isinstance(exc, UserDecisionError):
        base.update(error_type=exc.error_type or "needs_user_decision", message=msg, agent_directive=f"STOP_RETRY: {exc.directive}")
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
                  "--search", "--new-name", "--map", "--section"}
# 위 옵션의 짧은 이름 (모든 office-cli 파서에서 모음, 테스트로 확인): reader/outlook -q = --query
REDACT_SHORT = {"-q": "--query"}
# 도구마다 같은 이름이 내용이 아닌 옵션 (office.py transfer --from excel --to word, ppt.py slide --to 3)
REDACT_KEEP = {"office": {"--from", "--to"}, "ppt": {"--to"}}
# 도구마다 같은 이름이 내용인 옵션 (excel style --highlight '=완료:lightgreen'은 셀 값, word --highlight는 색 이름이라 남김)
REDACT_EXTRA = {"excel": {"--highlight"}}


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
    hidden = (REDACT_OPTIONS | REDACT_EXTRA.get(tool, set())) - REDACT_KEEP.get(tool, set())
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


def hidden_values(argv: Sequence[Any], tool: str = "") -> List[str]:
    """redact_argv가 가린 인자 값 (원래 글). redact_argv는 인자 하나에 결과 하나를 내므로 나란히 비교합니다."""
    raw = [str(x) for x in argv]
    values: List[str] = []
    for a, r in zip(raw, redact_argv(raw, tool)):
        if a == r:
            continue
        if _REDACTED.match(r) and "=" in r and r.index("=") < r.index("<"):
            opt = r.partition("=")[0]
            val = a[len(opt) + 1:] if a.startswith(opt + "=") else a[len(opt):]     # --opt=값, -q=값, -q값
        else:
            val = a
        values.append(val)
    return values


# 오류 로그에서 잘라 내는 message 부분 (문서·통합 문서에서 가져온 후보 목록과 비슷한 글: 제목, 시트·차트 이름, 본문 일부).
# 1차 방어는 오류를 만드는 곳의 log_message(InputError, UserDecisionError, ambiguous, not_found)이고, 이 표식은 빠뜨린 곳을 위한 대비입니다.
# 이 표식부터 message 끝까지를 '<문서에서 가져온 후보·문구 N자 생략>'으로 바꿉니다. stdout JSON의 message는 그대로입니다.
# 표식을 만드는 곳: ambiguous()·UserDecisionError의 '[후보]', not_found()의 '[추천]'·'사용 가능한 <종류>:'(종류에 공백 가능: '메일 폴더'),
# word 앵커의 '비슷한 문구:', 대괄호를 뺀 '후보:'·'추천:'
LOG_CUT_MARKERS = re.compile(r"\s*(?:\[후보\]|\[추천\]|후보:|추천:|비슷한 문구:|사용 가능한 [^:]{1,30}:)")
# 가린 값이 이보다 길면 앞부분(이 길이 이상)만 남은 미리보기('긴 글...')도 가립니다
LOG_PREFIX_MIN = 10
# 이보다 짧은 가린 값(--value 1)은 따옴표로 둘러싸인 곳만 가립니다 ('Sheet1'의 1을 가리지 않게. 따옴표 없이 쓴 1~2자 값은 남음)
LOG_SHORT_VALUE = 3
_QUOTES = "'\"\u2018\u2019\u201c\u201d"


def _hidden_spans(text: str, vals: Iterable[str]) -> List[Tuple[int, int]]:
    """text에서 가린 값이 나온 구간. 긴 값은 LOG_PREFIX_MIN자 이상의 앞부분도, 짧은 값은 따옴표 안만."""
    spans: List[Tuple[int, int]] = []
    for v in vals:
        if len(v) >= LOG_PREFIX_MIN:
            head = v[:LOG_PREFIX_MIN]
            i = text.find(head)
            while i >= 0:
                k = LOG_PREFIX_MIN
                while k < len(v) and i + k < len(text) and text[i + k] == v[k]:
                    k += 1
                spans.append((i, i + k))
                i = text.find(head, i + 1)
        elif len(v) >= LOG_SHORT_VALUE:
            spans += [(m.start(), m.end()) for m in re.finditer(re.escape(v), text)]
        else:
            spans += [(m.start(), m.end()) for m in re.finditer(re.escape(v), text)
                      if m.start() > 0 and text[m.start() - 1] in _QUOTES and m.end() < len(text) and text[m.end()] in _QUOTES]
    merged: List[Tuple[int, int]] = []
    for x, y in sorted(spans):
        if merged and x <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], y))
        else:
            merged.append((x, y))
    return merged


def mask_hidden_values(text: str, argv: Sequence[Any], tool: str = "") -> str:
    """가린 인자 값을 text에서 redact_argv와 같은 '<N자 생략>'으로 바꿉니다. 오류 로그의 message와 traceback에 씁니다."""
    vals = set()
    for v in hidden_values(argv, tool):
        vals.update(x for x in (v, v.strip(), unescape_newlines(v).strip()) if x)
    out: List[str] = []
    last = 0
    for x, y in _hidden_spans(text, vals):      # 원래 글에서 구간을 한 번에 구함 (넣은 '<N자 생략>'을 다시 바꾸지 않게)
        out += [text[last:x], f"<{y - x}자 생략>"]
        last = y
    return "".join(out) + text[last:]


def log_message(message: Any, argv: Sequence[Any], tool: str = "") -> str:
    """오류 로그에 남길 message: 가린 인자 값과 문서에서 가져온 후보·문구(LOG_CUT_MARKERS)를 뺀 글.
    message에는 오류의 log_message(문서 글을 뺀 글)가 있으면 그것을 줍니다."""
    try:
        msg = mask_hidden_values("" if message is None else str(message), argv, tool)
        m = LOG_CUT_MARKERS.search(msg)
        if m:
            msg = msg[:m.start()] + f" <문서에서 가져온 후보·문구 {len(msg) - m.start()}자 생략>"
        return msg
    except Exception:
        return "<message 생략 (가리기 실패)>"


def log_trace(trace: str, argv: Sequence[Any], tool: str = "") -> str:
    """traceback의 가린 인자 값을 바꿉니다 (표식 자르기는 하지 않음: 코드 줄과 예외 종류는 남김)."""
    try:
        return mask_hidden_values(trace, argv, tool)
    except Exception:
        return "<traceback 생략 (가리기 실패)>"


# 결과 dict(status: error)에 오류 로그용 message를 담는 키. run_cli가 stdout에 내기 전에 뺍니다.
LOG_MESSAGE_KEY = "_log_message"


def write_error_log(tool: str, argv: Optional[Sequence[str]], payload: Dict[str, Any], trace: str = "",
                    log_msg: Optional[str] = None) -> None:
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
                 f"결과: {payload.get('error_type')} | "
                 f"{log_message(payload.get('message') if log_msg is None else log_msg, argv, tool)}"]
        if RECOVERY_ATTEMPTS:
            lines.append("복구 시도: " + "; ".join(RECOVERY_ATTEMPTS))
        if trace:
            lines.append("traceback:\n" + log_trace(trace, argv, tool).rstrip())
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n\n")
    except Exception:
        pass


def _exc_log_message(exc: BaseException, payload: Dict[str, Any]) -> Optional[str]:
    """예외의 log_message (문서 글을 뺀 글). 없으면 None (payload의 message를 씀)."""
    lm = getattr(exc, "log_message", None)
    return None if lm is None or payload.get("error_type") == "internal" else str(lm)


# ------------------------------------------------------------------------------
# R10. 긴 작업은 나눠서 진행 (CONTINUE:). 에이전트의 명령 실행 도구는 120초에 명령을 끊고, 끊기면 결과 JSON이 없어
# LLM이 같은 명령을 반복함. 오래 걸리는 반복(문서 등록 등)은 한 단계를 끝낼 때마다 결과를 저장하고, 다음 단계 전에
# 시계를 봄(check_time). 제한을 넘을 것 같으면 Unfinished로 멈추고 run_cli가 CONTINUE 결과를 냄(같은 명령 = 이어서).
# 이름과 키는 ml-kit(ml_predict.py PART C)과 같음.
# ------------------------------------------------------------------------------
TIME_LIMIT_ENV = "OFFICE_CLI_TIME_LIMIT"     # 명령 하나의 초. 0 = 제한 없음 (사람이 돌리는 러너와 테스트)
DEFAULT_TIME_LIMIT = 100.0                   # OpenCode 120초보다 짧게 (JSON과 저장할 여유)
TIME_RESERVE = 30.0                          # 마지막 단계 뒤(검색, 결과 저장)에 남겨 둘 시간
RUN_START = time.monotonic()                 # run_cli가 명령마다 다시 잼


def time_limit() -> float:
    raw = os.environ.get(TIME_LIMIT_ENV, "").strip()
    if not raw:
        return DEFAULT_TIME_LIMIT
    try:
        v = float(raw)
    except ValueError:
        return DEFAULT_TIME_LIMIT
    return float("inf") if v <= 0 else v


class Unfinished(BaseException):
    """시간 제한 때문에 긴 작업을 멈춤 (끝낸 부분은 저장됨). 중간의 'except Exception'이 삼키지 않도록 BaseException.
    run_cli가 status ok, continue: true, 같은 명령의 next_commands로 바꿉니다."""

    def __init__(self, progress: str):
        super().__init__(progress)
        self.progress = progress


def check_time(next_step_seconds: float, progress: str) -> None:
    """다음 느린 단계(예상 next_step_seconds초)가 제한에서 TIME_RESERVE를 뺀 시각을 넘을 것 같으면 Unfinished.
    부르는 쪽은 끝낸 단계를 모두 저장해 두어야 합니다 (다음 실행이 다시 씀). 실행마다 최소 한 단계는 한 뒤에 부릅니다."""
    limit = time_limit()
    if limit == float("inf"):
        return
    if time.monotonic() - RUN_START + max(next_step_seconds, 0.0) > limit - TIME_RESERVE:
        raise Unfinished(progress)


def unfinished_payload(exc: "Unfinished", tool: str, command: str) -> Dict[str, Any]:
    argv = list(CURRENT_ARGV or [])
    nexts = [next_command(f"{tool}.py", argv, "같은 명령으로 이어서 진행 (끝낸 부분은 저장됨, 묻지 않고 그대로 실행)")] if argv else []
    payload: Dict[str, Any] = {
        "status": "ok", "tool": tool, "command": command, "preview": f"진행 중: {exc.progress}",
        "verified": False, "continue": True, "progress": exc.progress, "next_commands": nexts,
        "agent_directive": ("CONTINUE: 아직 끝나지 않았습니다. 명령 시간 제한 때문에 나눠서 진행하며, 지금까지 한 일은 저장되어 있습니다. "
                            + ("next_commands[0]의 command를" if nexts else "방금 실행한 명령을")
                            + " 바꾸지 말고 그대로 다시 실행하세요 (사용자에게 묻지 않음). 끝나면 VERIFIED 결과가 나옵니다."),
    }
    _attach_recovered(payload)
    return payload


def run_cli(tool: str, command: str, execute: Callable[[], Dict[str, Any]]) -> None:
    """도구 본문을 실행하고, 어떤 경우에도 stdout 1줄 JSON과 종료 코드(0/1/130)를 보장합니다.
    실패(exit 1)는 <데이터 폴더>/logs/errors.txt에 자동 기록합니다 (사용자 중단 130과 성공은 기록하지 않음).
    시간 제한으로 멈춘 긴 작업(Unfinished)은 CONTINUE 결과(exit 0, 오류 로그 없음, R10)."""
    global RUN_START
    RUN_START = time.monotonic()
    setup_console()
    try:
        result = execute()
    except Unfinished as u:
        log("STATUS", f"시간 제한으로 여기서 멈춤 (저장됨, 같은 명령으로 이어서): {u.progress}")
        emit_json(unfinished_payload(u, tool, command))
        sys.exit(0)
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
        write_error_log(tool, None, payload, trace if internal else "", log_msg=_exc_log_message(e, payload))
        emit_json(payload)
        sys.exit(1)
    log_msg = result.pop(LOG_MESSAGE_KEY, None) if isinstance(result, dict) else None
    if result.get("status") != "ok":
        write_error_log(tool, None, result, log_msg=log_msg)
    emit_json(result)
    sys.exit(0 if result.get("status") == "ok" else 1)


# 경량 LLM이 자주 지어내는 명령·옵션 (실사용 오류 로그에서 본 것). 실제 이름으로 안내합니다.
COMMAND_ALIASES = {"excel": {"list": "info", "sheets": "info"}, "word": {"list": "info", "toc": "outline"},
                   "ppt": {"list": "info"}, "office": {"list": "windows", "redo": "undo"},
                   "outlook": {"attachment": "attachments", "attach": "attachments", "download": "attachments"}}
OPTION_SYNONYMS = {"--target": ["--app", "--file", "--name"], "--file": ["--target"], "--path": ["--file", "--target"],
                   "--text": ["--value", "--file"],  # 엑셀 write에 워드식 --text를 쓰면 --target(철자 비슷)이 아니라 --value
                   "--output": ["--save"], "--out": ["--save"], "--sheet-name": ["--sheet", "--name"],
                   "--cell": ["--range"]}
# 다른 파일의 명령이 맞는 경우: [추천]으로 값만 바꾸면 또 틀리므로 맞는 명령을 안내만 함
_WRITE_TEXT = ('글을 문서에 넣으려면 열린 워드는 word.py write --text "내용", PPT는 ppt.py write --text "내용", '
               '엑셀 칸은 excel.py write --range A1 --value "내용"을 쓰세요.')
OTHER_COMMAND = {("office", "write"): "office.py에는 write가 없습니다. " + _WRITE_TEXT,
                 ("office", "transfer", "--text"): "transfer는 엑셀 차트·범위와 문서의 표를 옮기는 명령이라 글은 받지 않습니다. " + _WRITE_TEXT}


# 도구 파일 -> 그 명령을 설명하는 작업별 스킬 (skills\<이름>\SKILL.md). excel.py check·comment는 office-spec에도 있음
SKILL_OF_TOOL = {"excel": "office-excel", "word": "office-doc", "ppt": "office-doc", "office": "office-doc",
                 "outlook": "office-mail", "reader": "office-spec"}


def skill_file(tool: str) -> str:
    """도구의 스킬 파일 경로. 모르는 도구(테스트용 등)는 office-doc."""
    name = SKILL_OF_TOOL.get(tool, "office-doc")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "skills", name, "SKILL.md")


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
                   "skill_file": skill_file(self.tool_name),
                   "agent_directive": f"VERIFIED: 도움말을 stderr에 출력했습니다. {SKILL_OF_TOOL.get(self.tool_name, 'office-doc')} "
                                      "스킬(SKILL.md)의 명령 템플릿을 그대로 사용하세요. "
                                      "도구 코드(.py)는 읽지 마세요. 방법을 모르겠으면 사용자에게 물어보세요."})

    def _hint(self, message: str) -> str:
        """흔한 잘못(없는 명령, 다른 명령의 옵션, 따옴표 없는 경로)에 [추천]을 붙입니다."""
        import difflib
        top = self if self.prog.count(" ") == 0 else None
        subs = {}
        for parser in ([top] if top else []):
            for act in parser._actions:
                if isinstance(act, argparse._SubParsersAction):
                    subs = act.choices
        m = re.search(r"invalid choice: '([^']+)'", message)
        if m:
            other = OTHER_COMMAND.get((self.tool_name, m.group(1).lower()))
            if other:
                return " " + other
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
            other = OTHER_COMMAND.get((self.tool_name, cmd, f))
            if other:
                return " " + other
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
                               f"없으면 {SKILL_OF_TOOL.get(self.tool_name, 'office-doc')} 스킬(SKILL.md)의 명령 형식을 확인해 사용자에게 보고하세요.",
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
    """config.json (없으면 빈 설정). 위치: OFFICE_CLI_CONFIG > En-Genius 고정 폴더 > 도구 폴더. 프로세스당 한 번만 읽습니다."""
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
    os.makedirs(os.path.dirname(CONFIG_FILE) or ".", exist_ok=True)
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
    """사용자 데이터 폴더. 우선순위: OFFICE_CLI_DATA_DIR > config.json data_dir > <도구폴더 또는 En-Genius 고정 폴더>/data"""
    return _dir_setting("OFFICE_CLI_DATA_DIR", "data_dir", os.path.join(_tool_home(), "data"))


def get_cache_dir() -> str:
    """캐시 폴더 (읽은 문서 텍스트, 색인, 스냅샷). 우선순위: OFFICE_CLI_CACHE_DIR > config.json cache_dir > <도구폴더 또는 En-Genius 고정 폴더>/cache"""
    return _dir_setting("OFFICE_CLI_CACHE_DIR", "cache_dir", os.path.join(_tool_home(), "cache"))


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


def resolve_input_path(path: str, kind: str = "입력 파일", exts: Optional[Iterable[str]] = None) -> str:
    """입력 파일 경로를 찾습니다. 순서: 절대경로/현재 폴더 -> data 폴더 -> 작업 폴더.
    없으면 data 폴더의 비슷한 파일명을 [추천]으로 붙여 InputError를 냅니다.
    exts(소문자 확장자, 예: (".csv", ".md"))를 주면 후보 목록에 그 확장자 파일만 보입니다
    (표 파일을 못 찾았는데 그림·문서 파일 이름까지 나열하던 문제)."""
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
            names += [n for n in os.listdir(folder) if os.path.isfile(os.path.join(folder, n))
                      and (exts is None or os.path.splitext(n)[1].lower() in exts)]
        except OSError:
            pass
    raise not_found(kind, os.path.basename(raw), names, extra=f" (찾은 위치: {', '.join(candidates)})")


_BAD_NAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]+')
_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"} | {f"{p}{i}" for p in ("COM", "LPT") for i in range(1, 10)}


def safe_filename(text: Any, limit: int = 0, default: str = "file") -> str:
    """파일 이름에 쓸 수 없는 글자(\\ / : * ? \" < > |, 제어 문자)를 _로, 끝의 점·공백은 빼고,
    Windows 예약 이름(CON, NUL, COM1 등)은 앞에 _를 붙입니다. limit>0이면 그 길이로 자릅니다.
    (차트·시트 이름 '1/4분기 매출'이 기본 파일 이름이 되면 경로가 깨지던 문제)"""
    name = _BAD_NAME_CHARS.sub("_", "" if text is None else str(text))
    if limit > 0:
        name = name[:limit]
    name = name.strip().rstrip(". ")
    if name.split(".")[0].strip().upper() in _RESERVED_NAMES:
        name = "_" + name
    return name or default


def resolve_output_path(path: Optional[str], default_stem: str, ext: str, overwrite: bool = False) -> str:
    """출력 파일 경로를 정합니다.
    - 생략하면 작업 폴더에 '<default_stem>_<시각><ext>'로 만듭니다 (default_stem의 쓸 수 없는 글자는 _로).
    - 파일명만 주면 data 폴더 아래에 만듭니다.
    - 확장자가 없거나 다르면 ext로 맞춥니다.
    - 그 파일이 이미 있으면 덮어쓰지 않고 시각을 붙인 새 이름으로 바꿉니다 (R4, recovered로 알림). overwrite=True면 그대로.
    """
    ext = ext if ext.startswith(".") else "." + ext
    if not path:
        cleanup_work_dir()
        # 기본 이름은 차트·시트·문서 이름에서 오므로 / ? : 같은 글자가 섞일 수 있음 -> 여기 한 곳에서 정리
        auto = os.path.join(get_work_dir(), f"{safe_filename(default_stem, default='output')}_{timestamp()}{ext}")
        if os.path.exists(auto) and not overwrite:
            # 같은 초에 같은 이름으로 저장하면 앞 결과를 덮어쓰던 문제: 번호를 붙인 새 이름으로
            alt = _timestamped(auto)
            note_recovered(f"{os.path.basename(auto)} 이미 있음 (같은 초에 저장) -> {os.path.basename(alt)}로 저장")
            auto = alt
        return auto
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


def _sniff_delimiter(text: str) -> str:
    sample = text[:4096]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return "\t" if "\t" in sample and "," not in sample else ","


def read_csv_rows(path: str, delimiter: Optional[str] = None) -> List[List[str]]:
    """CSV/TSV를 2차원 문자열 리스트로 읽습니다 (delimiter를 주지 않으면 구분자 자동 판별)."""
    text = read_text_file(path)
    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter or _sniff_delimiter(text))]


def check_two_columns(path: str, rows: Sequence[Sequence[Any]], start: int, what: str, cols: str, done: str) -> None:
    """2열 표 파일(칸, 값 / 칸, 메모 / 자리, 값)에서 셋째 열 이후에 비지 않은 칸이 있으면 InputError (쓰기 전에).
    값의 쉼표(구분자)가 열을 나눈 경우라 둘째 열만 쓰면 값 뒷부분이 사라짐 (예: C7,1,500 -> 1).
    rows는 머리글을 뺀 행, start는 rows[0]의 파일 줄 번호. 고친 줄은 둘째 열 이후를 구분자로 다시 이어 큰따옴표로 감쌈.
    오류 로그에는 줄 번호만 (값은 문서 내용)."""
    def cell(v: Any) -> str:
        return "" if v is None else str(v)
    bad = [(n, list(r)) for n, r in enumerate(rows, start) if any(cell(v).strip() for v in list(r)[2:])]
    if not bad:
        return
    ext = os.path.splitext(path)[1].lower()
    lines = ", ".join(str(n) for n, _ in bad[:10]) + (" ..." if len(bad) > 10 else "")
    head = f"{what}의 {lines}번째 줄에 열이 3개 이상이라 아무것도 {done}. 열은 {cols} 두 개입니다."
    if ext not in (".csv", ".tsv"):
        raise InputError(head + " 값에 | 가 있으면 \\| 로 쓰고 파일을 저장한 뒤 다시 실행하세요.", log_message=head)
    try:
        delim = _sniff_delimiter(read_text_file(path))
    except Exception:
        delim = "\t" if ext == ".tsv" else ","
    shown = "쉼표" if delim == "," else "탭" if delim == "\t" else f"'{delim}'"

    def fixed(r: List[Any]) -> str:
        vals = [cell(v) for v in r]
        while len(vals) > 2 and not vals[-1].strip():
            vals.pop()
        buf = io.StringIO()
        csv.writer(buf, delimiter=delim, quoting=csv.QUOTE_MINIMAL, lineterminator="").writerow([vals[0], delim.join(vals[1:])])
        return buf.getvalue()
    hint = f" 값에 {shown}가 있으면 값을 큰따옴표로 감싸세요. 그 줄을 고쳐 파일을 저장한 뒤 다시 실행하세요."
    # [추천]은 쓰지 않음: '[추천] C7,"1,500"'을 경량 LLM이 인자 값으로 읽어 명령에 넣었음 (고친 줄은 파일 내용)
    tail = " 고친 줄: " + " / ".join(f"{n}번째 줄 {fixed(r)}" for n, r in bad[:5])
    raise InputError(head + hint + tail, log_message=head + hint + f" (고친 줄 {DOC_OMITTED})")


UNCOMPUTED_POSITIONS = 100_000   # positions=True일 때 담는 칸 위치 수 상한 (넘으면 멀리 떨어진 칸을 가려내지 않음)
UNCOMPUTED_SHOW = 20       # 계산된 값이 없는 수식 칸을 결과에 담는 개수 상한
_XLSX_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XLSX_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_XLSX_F_TAG = re.compile(rb"<(?:\w+:)?f[\s>/]")
_CELL_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def _cell_rc(ref: str) -> Optional[Tuple[int, int]]:
    m = _CELL_REF.match(ref or "")
    if not m:
        return None
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


def _col_letters(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _array_box(ref: str) -> Optional[Tuple[int, int, int, int]]:
    """배열 수식의 ref("C1:C3" 또는 "C1") -> (행1, 열1, 행2, 열2). 읽지 못하면 None."""
    parts = ref.split(":")
    a = _cell_rc(parts[0])
    b = _cell_rc(parts[-1])
    if a is None or b is None or len(parts) > 2:
        return None
    return min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])


def _full_calc_on_load(wb_root: Any) -> bool:
    """workbook.xml의 <calcPr fullCalcOnLoad="1"|"true">: 엑셀이 열 때 모든 수식을 다시 계산하라는 표시.
    xlsxwriter·openpyxl 같은 프로그램이 쓰며 파일의 <v>(예: 0)는 계산된 값이 아님. 엑셀은 저장할 때 이 표시를 지웁니다."""
    calc = wb_root.find(_XLSX_MAIN + "calcPr")
    return calc is not None and (calc.get("fullCalcOnLoad") or "").strip().lower() in ("1", "true")


def xlsx_full_calc_on_load(path: str) -> bool:
    """파일의 workbook.xml에 fullCalcOnLoad 표시가 있는지. 읽지 못하면 True (수식 칸 검사를 건너뛰지 않게)."""
    import zipfile
    import xml.etree.ElementTree as ET
    try:
        with zipfile.ZipFile(path) as z:
            return _full_calc_on_load(ET.fromstring(z.read("xl/workbook.xml")))
    except Exception:
        return True


def xlsx_uncomputed_formulas(path: str, sheet: Optional[str] = None,
                             area: Optional[Tuple[int, int, int, int]] = None,
                             limit: int = UNCOMPUTED_SHOW, positions: bool = False) -> Dict[str, Any]:
    """계산된 값이 저장되지 않은 수식 칸 (openpyxl 등이 쓰고 엑셀에서 한 번도 열지 않은 파일).
    openpyxl data_only로 읽으면 이런 칸은 빈 칸(None)이라 값이 없는 것처럼 보입니다.
    workbook.xml에 fullCalcOnLoad 표시가 있으면(xlsxwriter는 <v>0</v>을 씀) 모든 수식 칸을 계산되지 않은 칸으로 셉니다.
    sheet: 이 시트만 (None이면 모든 시트), area: (시작 행, 시작 열, 끝 행, 끝 열) 안의 칸만 (끝 행 뒤는 해석하지 않음).
    수식 태그(<f>)가 없는 시트는 XML을 해석하지 않습니다. 엑셀에서 저장한 파일은 값이 있어 0개입니다
    (빈 글 결과(t="str")는 계산된 값).
    반환: {"count": 전체 수, "cells": ["시트!A3", ...] (앞 limit개), "sheets": {시트: 수},
    "bounds": [시작 행, 시작 열, 끝 행, 끝 열] (센 칸 전체를 감싸는 범위, 0개면 없음), "full_calc_on_load": True (표시가 있을 때만)}.
    positions=True(시트 하나를 볼 때): "positions"(센 칸 (행, 열) 목록, UNCOMPUTED_POSITIONS개까지, 넘으면 "positions_truncated")와
    "value_bounds"(값이 있는 칸(계산된 값이 없는 수식 칸 제외)을 감싸는 범위, 없으면 키 없음)도 줍니다 (멀리 떨어진 칸 가려내기용).
    읽지 못하면(MemoryError 포함) {"count": 0, "unknown": True, "error": 예외 이름}: 확인하지 못한 것이라 호출하는 쪽은 UNVERIFIED."""
    import posixpath
    import zipfile
    import xml.etree.ElementTree as ET
    result: Dict[str, Any] = {"count": 0, "cells": [], "sheets": {}}
    bounds: List[int] = []
    vbounds: List[int] = []
    pos: List[Tuple[int, int]] = []
    pos_cut = False

    def grow(box: List[int], r1: int, c1: int, r2: int, c2: int) -> List[int]:
        return [r1, c1, r2, c2] if not box else [min(box[0], r1), min(box[1], c1), max(box[2], r2), max(box[3], c2)]

    def keep_pos(rc: Tuple[int, int]) -> None:
        nonlocal pos_cut
        if len(pos) < UNCOMPUTED_POSITIONS:
            pos.append(rc)
        else:
            pos_cut = True
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            full = _full_calc_on_load(wb)
            if full:
                result["full_calc_on_load"] = True
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            targets = {r.get("Id"): r.get("Target") or "" for r in rels}
            for sh in wb.iter(_XLSX_MAIN + "sheet"):
                title = sh.get("name") or ""
                if sheet is not None and title != sheet:
                    continue
                target = targets.get(sh.get(_XLSX_REL + "id"), "")
                part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
                if part not in names:
                    continue
                data = z.read(part)
                if not _XLSX_F_TAG.search(data):
                    continue
                n = 0
                q = f"'{title}'" if re.search(r"[\s!']", title) else title
                arrays: List[List[Any]] = []        # 배열 수식 [r1, c1, r2, c2, 왼쪽 위 (행, 열), 만난 다른 칸 set]

                def count(rc: Optional[Tuple[int, int]], ref: str) -> None:
                    nonlocal n, bounds
                    n += 1
                    if len(result["cells"]) < limit:
                        result["cells"].append(f"{q}!{ref}")
                    if rc is not None:
                        bounds = grow(bounds, rc[0], rc[1], rc[0], rc[1])
                        if positions:
                            keep_pos(rc)

                def in_area(rc: Optional[Tuple[int, int]]) -> bool:
                    return area is None or rc is None or (area[0] <= rc[0] <= area[2] and area[1] <= rc[1] <= area[3])

                for _, el in ET.iterparse(io.BytesIO(data), events=("end",)):
                    if el.tag == _XLSX_MAIN + "row":
                        r = el.get("r") or ""
                        el.clear()
                        if area is not None and r.isdigit() and int(r) >= area[2]:
                            break                   # 행은 번호 순: 범위 끝 행 뒤는 해석하지 않음
                        continue
                    if el.tag != _XLSX_MAIN + "c":
                        continue
                    f = el.find(_XLSX_MAIN + "f")
                    v = el.find(_XLSX_MAIN + "v")
                    ref = el.get("r") or ""
                    rc = _cell_rc(ref)
                    if f is not None and f.get("t") == "array" and rc is not None:
                        box = _array_box(f.get("ref") or "")
                        if box is not None and box != (rc[0], rc[1], rc[0], rc[1]):
                            arrays.append([box[0], box[1], box[2], box[3], rc, set()])
                    counted_before = n
                    if f is not None and (full or el.get("t") not in ("str", "inlineStr")):
                        if (full or v is None or not (v.text or "")) and in_area(rc):
                            count(rc, ref)
                    elif f is None and rc is not None and arrays:
                        # 배열 수식의 다른 칸: 파일에는 <f> 없이 값만 있음 (계산된 값이 없으면 수식 칸으로 셈, 23차 L1)
                        for a in arrays:
                            if a[0] <= rc[0] <= a[2] and a[1] <= rc[1] <= a[3]:
                                a[5].add(rc)
                                computed = el.get("t") in ("str", "inlineStr") or (v is not None and (v.text or ""))
                                if (full or not computed) and in_area(rc):
                                    count(rc, ref)
                                break
                    if positions and rc is not None and n == counted_before and (
                            (v is not None and (v.text or "")) or el.find(_XLSX_MAIN + "is") is not None):
                        vbounds = grow(vbounds, rc[0], rc[1], rc[0], rc[1])
                    el.clear()
                # 배열 수식의 다른 칸 중 파일에 아예 없는 칸 (openpyxl은 왼쪽 위 칸만 씀): 계산된 값이 없는 수식 칸
                for r1, c1, r2, c2, anchor, seen in arrays:
                    if area is not None:
                        r1, c1, r2, c2 = max(r1, area[0]), max(c1, area[1]), min(r2, area[2]), min(c2, area[3])
                    if r1 > r2 or c1 > c2:
                        continue
                    seen_in = sum(1 for x in seen if r1 <= x[0] <= r2 and c1 <= x[1] <= c2)
                    anchor_in = 1 if (r1 <= anchor[0] <= r2 and c1 <= anchor[1] <= c2) else 0
                    missing = (r2 - r1 + 1) * (c2 - c1 + 1) - seen_in - anchor_in
                    if missing <= 0:
                        continue
                    shown = 0
                    for rr in range(r1, r2 + 1):
                        if shown >= missing or len(result["cells"]) >= limit:
                            break
                        for cc in range(c1, c2 + 1):
                            if (rr, cc) in seen or (rr, cc) == anchor:
                                continue
                            if shown >= missing or len(result["cells"]) >= limit:
                                break
                            result["cells"].append(f"{q}!{_col_letters(cc)}{rr}")
                            shown += 1
                    n += missing
                    bounds = grow(bounds, r1, c1, r2, c2)
                    if positions:
                        if missing > UNCOMPUTED_POSITIONS - len(pos):
                            pos_cut = True
                        else:
                            pos.extend((rr, cc) for rr in range(r1, r2 + 1) for cc in range(c1, c2 + 1)
                                       if (rr, cc) not in seen and (rr, cc) != anchor)
                if n:
                    result["count"] += n
                    result["sheets"][title] = n
    except Exception as e:                      # MemoryError 포함: 확인하지 못함 (0개로 보면 VERIFIED가 되어 버림)
        debug(f"수식 값 확인 실패: {type(e).__name__}")
        return {"count": 0, "cells": [], "sheets": {}, "unknown": True, "error": type(e).__name__}
    if bounds:
        result["bounds"] = bounds
    if positions:
        result["positions"] = pos
        if pos_cut:
            result["positions_truncated"] = True
        if vbounds:
            result["value_bounds"] = vbounds
    return result


def uncomputed_found(info: Optional[Dict[str, Any]]) -> bool:
    """알려야 할 결과인지: 계산되지 않은 수식 칸이 있거나 확인하지 못함(unknown)."""
    return bool(info) and bool(info.get("count") or info.get("unknown"))


def uncomputed_note(info: Dict[str, Any]) -> str:
    """xlsx_uncomputed_formulas 결과를 UNVERIFIED 문구로 (칸 주소는 앞 10개)."""
    if info.get("unknown"):
        return (f"수식 계산 여부를 확인하지 못함 (파일의 시트 XML을 읽지 못함: {info.get('error') or '알 수 없음'}). "
                "수식 칸이 있으면 빈 칸이나 계산되지 않은 값으로 읽혔을 수 있으니 사용자에게 그대로 알리세요")
    cells = ", ".join(info["cells"][:10]) + (" 등" if info["count"] > 10 else "")
    if info.get("full_calc_on_load"):
        return (f"수식 칸 {info['count']}개({cells})는 파일이 '열 때 다시 계산'(fullCalcOnLoad)으로 저장되어 있어 파일에 적힌 값"
                "(빈 칸이나 0)이 계산된 값이 아님 (엑셀에서 연 적 없이 프로그램이 만든 파일). 사용자에게 엑셀에서 파일을 열어 저장한 뒤 "
                "다시 실행하도록 알리세요")
    return (f"수식 칸 {info['count']}개({cells})는 계산된 값이 파일에 저장되어 있지 않아 빈 칸으로 읽힘 "
            "(엑셀에서 연 적 없이 프로그램이 만든 파일이라 값이 계산된 적이 없음). 사용자에게 엑셀에서 파일을 열어 저장한 뒤 "
            "다시 실행하도록 알리세요")


def uncomputed_fields(info: Dict[str, Any]) -> Dict[str, Any]:
    """결과 JSON 키: uncomputed_formulas(앞 UNCOMPUTED_SHOW개), uncomputed_count, 잘렸으면 uncomputed_truncated.
    확인하지 못했으면 uncomputed_check: "unknown"만."""
    if info.get("unknown"):
        return {"uncomputed_check": "unknown"}
    out: Dict[str, Any] = {"uncomputed_formulas": list(info["cells"]), "uncomputed_count": info["count"]}
    if info["count"] > len(info["cells"]):
        out["uncomputed_truncated"] = True
    return out


_PAYLOAD_BASE = ("status", "tool", "command", "preview", "verified", "unverified", "agent_directive", "recovered")


def with_uncomputed(payload: Dict[str, Any], info: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """R1: 계산된 값이 없는 수식 칸을 빈 칸으로 읽은 성공 응답은 VERIFIED가 아님 -> UNVERIFIED로 바꾸고 칸 목록을 붙입니다."""
    if not uncomputed_found(info) or payload.get("status") != "ok":
        return payload
    fields = {k: v for k, v in payload.items() if k not in _PAYLOAD_BASE}
    fields.update(uncomputed_fields(info))
    return ok_unverified(payload["tool"], payload["command"], payload["preview"],
                         [uncomputed_note(info)] + list(payload.get("unverified") or []), **fields)


def _transient_fs_error(e: OSError) -> bool:
    """잠깐 기다리면 풀릴 수 있는 오류: 접근 거부, 공유 위반, 잠금 위반 (winerror 5/32/33)."""
    return isinstance(e, PermissionError) or getattr(e, "winerror", None) in (5, 32, 33)


def retry_fs(fn: Callable[[], Any], attempts: int = 6, delay: float = 0.5, deadline: Optional[float] = None) -> Any:
    """백신/동기화 프로그램의 일시적 잠금에 대비해 파일 작업을 재시도합니다.
    deadline(time.monotonic() 기준 시각)을 주면 그 시각을 넘겨 기다리지 않습니다 (잠금을 쥔 채 오래 기다리지 않게).
    한도를 넘기면 오류 로그의 '복구 시도:' 줄에 남기고 마지막 예외를 올립니다."""
    last: Optional[BaseException] = None
    for i in range(attempts):
        try:
            return fn()
        except OSError as e:
            if not _transient_fs_error(e):
                raise                       # 없는 파일, 디스크 가득 참 등은 기다려도 소용없음
            last = e
            wait = delay * (i + 1)
            if deadline is not None:
                wait = min(wait, deadline - time.monotonic())
                if wait <= 0:
                    attempts = i + 1
                    break
            time.sleep(wait)
    assert last is not None
    name = os.path.basename(str(getattr(last, "filename", "") or ""))
    record_recovery_attempt(f"파일 작업 {name} {attempts}회 재시도 실패 ({type(last).__name__})".replace("  ", " "))
    raise last


def safe_write_text(path: str, text: str, encoding: str = "utf-8", deadline: Optional[float] = None) -> None:
    """임시 파일에 쓴 뒤 교체해, 쓰다가 중단돼도 기존 파일이 깨지지 않게 합니다. deadline은 retry_fs와 같음."""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = f"{path}.tmp_{os.getpid()}"
    with open(tmp, "w", encoding=encoding, newline="") as f:
        f.write(text)
    retry_fs(lambda: os.replace(tmp, path), deadline=deadline)


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


def not_found(kind: str, name: str, candidates: Iterable[str], extra: str = "", vocab: bool = False,
              display: Optional[Callable[[str], str]] = None, log_name: Optional[str] = None,
              log_extra: Optional[str] = None) -> InputError:
    """'찾을 수 없음' InputError (R8).
    - 표기만 다른(대소문자·공백·_-) 이름이 하나뿐이면 [추천]: 뜻이 같아 LLM이 바로 1회 재시도해도 됩니다.
    - 부분·유사 일치는 뜻이 다를 수 있어(2024 -> 2024_1분기, 견적서.pdf -> 견적서_구버전.pdf) [후보]: 사용자가 고릅니다.
    - vocab=True(도구가 정한 옵션 값: 색 이름, 테마 등): 오타로 보이는 가까운 값 하나도 [추천] (사용자 데이터가 아니라 뜻이 바뀌지 않음).
    - 비슷한 것도 없으면 사용 가능한 이름 일부.
    display: 비교는 name·candidates(끝 이름)로 하고 message에 보일 때만 바꾸는 함수 (예: 메일 폴더 끝 이름 -> 전체 경로.
    전체 경로로 비교하면 같은 상위 경로 때문에 관계없는 폴더가 모두 비슷하게 나옴).
    log_name: 오류 로그 message에 name 대신 쓸 글 (예: 메일 계정 이름을 뺀 경로).
    log_extra: 오류 로그 message에 extra 대신 쓸 글 (extra에 계정 이름·문서 글이 들어갈 때 반드시 줌)."""
    show = display or (lambda c: c)
    cands = [str(c) for c in candidates]
    key = _same_name_key(name)
    same = [c for c in cands if key and _same_name_key(c) == key]
    if vocab and not same:
        close = difflib.get_close_matches(key, [_same_name_key(c) for c in cands], n=2, cutoff=0.75)
        if len(close) == 1:
            same = [c for c in cands if _same_name_key(c) == close[0]][:1]
    head = f"{kind} '{show(name)}'을(를) 찾을 수 없습니다.{extra}"
    masked = log_name is not None or log_extra is not None
    log_head = head if not masked else (f"{kind} '{name if log_name is None else log_name}'을(를) 찾을 수 없습니다."
                                        f"{extra if log_extra is None else log_extra}")
    # 오류 로그에는 후보 이름(시트·문서·제목·메일 폴더·파일 이름 목록)을 남기지 않음 (log_message)
    if len(same) == 1:
        return InputError(f"{head} 표기만 다른 이름이 있습니다. [추천] {show(same[0])}",
                          log_message=f"{log_head} 표기만 다른 이름이 있습니다. [추천 생략]")
    hints = same or suggest(name, cands)
    if hints:
        tail = " 비슷한 이름이 있지만 같은 것인지 도구가 정할 수 없습니다. 사용자에게 고르게 하세요."
        return ambiguous(head + tail, [show(c) for c in hints], log_message=log_head + tail)
    if cands:
        shown = ", ".join(show(c) for c in cands[:15]) + (" ..." if len(cands) > 15 else "")
        return InputError(f"{head} 사용 가능한 {kind}: {shown}", log_message=f"{log_head} [사용 가능한 이름 {len(cands)}개 생략]")
    return InputError(head, log_message=log_head if masked else None)


def clamp_cell(value: Any, max_chars: int = MAX_CELL_CHARS) -> Tuple[Any, bool]:
    """긴 글 값 하나를 상한으로 자릅니다 (숫자 등 글이 아닌 값은 그대로). (값, 잘렸는지 여부)"""
    if isinstance(value, str) and len(value) > max_chars:
        return value[:max_chars], True
    return value, False


def clamp_rows(rows: List[List[Any]], max_rows: int = DEFAULT_MAX_ROWS,
               max_chars: int = DEFAULT_MAX_CHARS) -> Tuple[List[List[Any]], bool, int, int]:
    """행 수와 대략적인 글자 수 상한으로 표를 자릅니다. (잘린 표, 행이 잘렸는지, 글이 잘린 칸 수, 뺀 열 수)
    칸마다 MAX_CELL_CHARS에서 자르고, 첫 행도 글자 수 상한을 넘으면 넘는 칸부터 뺍니다 (첫 행이 아주 큰 표도 상한 안).
    '글이 잘린 칸'(긴 값)과 '뺀 열'(첫 행이 상한을 넘어 표시하지 않은 칸)은 따로 셉니다."""
    out: List[List[Any]] = []
    used = 0
    cells_cut = 0
    for row in rows:
        if len(out) >= max_rows:
            return out, True, cells_cut, 0
        cut_row = []
        cut_flags = []
        for v in row:
            v, cut = clamp_cell(v)
            cut_flags.append(cut)
            cut_row.append(v)
        size = sum(len(str(v)) + 1 for v in cut_row)
        if out and used + size > max_chars:
            return out, True, cells_cut, 0
        if not out and size > max_chars:          # 첫 행만으로 상한을 넘음: 들어가는 칸까지만
            keep, width = [], 0
            for v in cut_row:
                width += len(str(v)) + 1
                if keep and width > max_chars:
                    break
                keep.append(v)
            cells_cut += sum(cut_flags[:len(keep)])
            out.append(keep)
            return out, len(rows) > 1, cells_cut, len(cut_row) - len(keep)
        cells_cut += sum(cut_flags)
        out.append(cut_row)
        used += size
    return out, False, cells_cut, 0


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

# '__'는 CommonMark처럼 단어 안에서는 굵게 표시가 아님: 앞에 글자·숫자가 없을 때만 열고 뒤에 없을 때만 닫음
# (report__2024__final.xlsx는 그대로, __init__ 메서드는 굵은 'init'). '**'는 단어 안에서도 굵게.
_MD_BOLD = re.compile(r"\*\*(.+?)\*\*|(?<!\w)__(?=\S)(.+?)(?<=\S)__(?!\w)")
# 번호 목록 기호: 1~3자리 숫자만 ('2026. 11. 3. 회의'의 연도는 목록 번호가 아님)
MD_NUMBERED = re.compile(r"^\s*(\d{1,3})[.)]\s+")


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


def plain_heading(text: str) -> str:
    """제목·슬라이드 제목에 넣을 글: 굵게 표시(** __)를 뺀 글 (제목은 굵기가 섞인 글을 담지 않음. word·ppt 모든 수준 같은 규칙)."""
    return split_bold(text)[0]


_TABLE_PIPE = re.compile(r"(?<!\\)\|")
_FENCE = re.compile(r"^(`{3,}(?!.*`)|~{3,})")    # 백틱 펜스의 정보 글에는 백틱이 없음 (```x```는 인라인 코드)


def parse_markdown_table(lines: List[str]) -> List[List[str]]:
    """'|'로 나눈 표 (앞에 역슬래시가 있는 '\\|'는 칸 안의 글자 '|'). 정렬 줄(머리글 다음 줄은 ':-', '-:' 등 '-' 하나부터, 그 밖은 '---' 이상)은 뺍니다."""
    rows: List[List[str]] = []
    for no, ln in enumerate(lines):
        s = ln.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|") and not s.endswith("\\|"):
            s = s[:-1]
        cells = [c.strip().replace("\\|", "|") for c in _TABLE_PIPE.split(s)]
        dash = r":?-+:?" if no == 1 else r":?-{3,}:?"   # 머리글 다음 줄은 '-' 하나부터 정렬 줄
        if cells and any(cells) and all(re.fullmatch(dash, c) for c in cells if c):
            continue
        rows.append(cells)
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


_YAML_LINE = re.compile(r"^(?:\s|-\s|-$|#|[^\s:#][^:]{0,60}:(?:\s|$))")


def front_matter_lines(text: str) -> int:
    """맨 앞 YAML 머리말(1번째 줄 '---', 그 뒤 'key: 값' 줄들, '---' 또는 '...' 줄로 닫힘)의 줄 수 (닫는 줄 포함). 없으면 0.
    닫는 줄이 없거나 사이에 'key:' 모양이 아닌 줄이 있으면 머리말이 아님 (첫 줄 '---'는 구분선)."""
    lines = text.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines or lines[0].strip() != "---":
        return 0
    keys = 0
    for i, ln in enumerate(lines[1:], 1):
        s = ln.rstrip()
        if s in ("---", "..."):
            return i + 1 if keys else 0
        if not s.strip():
            continue
        if not _YAML_LINE.match(s):
            return 0
        keys += bool(re.match(r"^[^\s:#-][^:]{0,60}:(?:\s|$)", s))
    return 0


def parse_markdown_blocks(text: str) -> List[Tuple[str, Any]]:
    """마크다운을 블록 목록으로 바꿉니다.
    ('heading', (레벨, 텍스트)), ('paragraph', 텍스트), ('bullets', [항목]), ('numbered', [항목]),
    ('table', [[셀]]), ('code', 줄), ('break', None)  # '---' 구분선(thematic break)
    ('front_matter', 줄 수): 맨 앞 YAML 머리말. 넣지 않는 블록 (쓰는 쪽이 warnings에 남김)
    """
    blocks: List[Tuple[str, Any]] = []
    para: List[str] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    skip = front_matter_lines(text)
    if skip:
        blocks.append(("front_matter", skip))
        lines = lines[skip:]

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
        fence = _FENCE.match(s)
        if fence:                                   # 코드 블록: 줄마다 단락 하나, 글자 그대로 (제목·목록으로 읽지 않음)
            flush()
            mark = fence.group(1)
            i += 1
            while i < len(lines):
                ln = lines[i].rstrip()
                i += 1
                if ln.strip().startswith(mark[0] * len(mark)) and not ln.strip().lstrip(mark[0]):
                    break                           # 닫는 펜스 (닫지 않으면 끝까지 코드)
                if ln.strip():
                    blocks.append(("code", ln))     # 굵게(** __) 등 마크다운 기호를 해석하지 않고 그대로
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
        num = MD_NUMBERED.match(s)
        if num:
            flush()
            run = []
            while i < len(lines) and MD_NUMBERED.match(lines[i]):
                run.append(lines[i].strip())
                i += 1
            first = int(num.group(1))
            if blocks and blocks[-1][0] == "numbered" and first == len(blocks[-1][1]) + 1:
                blocks[-1][1].extend(MD_NUMBERED.sub("", x).strip() for x in run)   # 빈 줄로 띄운 같은 목록의 다음 번호
            elif first == 1:
                blocks.append(("numbered", [MD_NUMBERED.sub("", x).strip() for x in run]))
            else:
                # 1이 아닌 번호로 시작하는 목록: 자동 번호는 1부터 다시 매기므로 사용자 번호를 글로 둔 단락으로 넣음
                blocks += [("paragraph", x) for x in run]
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
    return bool(re.search(r"(?m)^\s*(#{1,6}\s|[-*+]\s|\d{1,3}[.)]\s|\|.*\||```|~~~)", text) or "**" in text
                or front_matter_lines(text))


def front_matter_warning(lines: int) -> str:
    return f"맨 앞 YAML 머리말({lines}줄, '---'로 둘러싼 key: 값)은 문서 정보라 넣지 않았습니다."
