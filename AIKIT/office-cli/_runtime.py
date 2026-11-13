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
VERSION = "1.6.7"                         # <AIKIT>\.opencode\skills\<*-kit>\SKILL.md metadata.version과 같아야 함 (테스트로 확인)
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))


def engenius_home() -> Optional[str]:
    """도구 파일이 En-Genius 트레이의 임시 폴더(%TEMP%\\engenius-skill\\<id>)에서 실행 중이면 사용자 데이터를 둘 고정 폴더.
    트레이는 그 폴더를 실행이 끝나면 지우므로 도구 폴더 아래(data, cache, config.json)가 남지 않는다.
    배포 스킬(docs/engenius)은 C:\\AIKIT의 도구를 실행하므로 해당 없음: OpenCode와 같은 data, cache를 쓴다."""
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
DEFAULT_MAX_CHARS = 6000   # 표 글자 수 상한 (8000이면 8열 200행에 JSON 약 16KB, 2026-11-09 검토 5장 4번)
MAX_CELL_CHARS = 500        # JSON에 담는 셀·찾은 값 하나의 글자 상한 (전체는 --save 파일)
# 작업 폴더 파일 보관 기간 (일)
WORK_RETENTION_DAYS = 7
# 오류 로그가 이보다 커지면 errors.old.txt로 넘기고 새로 시작 (1세대만 보관)
ERROR_LOG_MAX = 2 * 1024 * 1024
# 실제 명령 인자. JsonArgumentParser.parse_args가 채웁니다 (인자 오류 기록에도 실제 명령을 남기려고)
CURRENT_ARGV: Optional[List[str]] = None
CURRENT_TOOL = ""          # 실행 중인 도구 (excel, word, ...): retry_command가 전체 명령을 만들 때 씀
# JIT 복구가 한도를 넘겨 실패한 시도 (오류 로그의 '복구 시도:' 줄)
RECOVERY_ATTEMPTS: List[str] = []


# ==============================================================================
# 1. 콘솔과 로그
# ==============================================================================

def setup_console() -> None:
    """stdout/stderr를 UTF-8로 고정합니다 (Windows 콘솔과 파이프의 기본값은 CP949)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            # 이미 UTF-8이어도 errors="replace"로 (PYTHONIOENCODING=utf-8은 strict라 반쪽 이모지(lone surrogate) 하나로
            # UnicodeEncodeError, stdout에 JSON이 없던 것, 2026-11-09 검토 B9)
            if stream is not None and ((stream.encoding or "").lower() != "utf-8" or getattr(stream, "errors", "") != "replace"):
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


def ambiguous(message: str, candidates: Iterable[str], log_message: Optional[str] = None,
              replace: Optional[str] = None, values: Optional[Iterable[str]] = None,
              option: Optional[str] = None) -> InputError:
    """후보가 둘 이상이라 고를 수 없음. message 끝에 '[후보] a, b, c'(최대 10개)를 붙입니다.
    후보는 문서에서 가져온 글(제목, 단락, 시트·문서 이름)이라 오류 로그에는 개수만 남깁니다.
    replace: 실행한 명령에서 바꿀 값(틀린 이름). 후보마다 그 값을 바꾼 전체 명령을 만들 수 있으면 '[후보] 1) a 2) b'와 같은 번호의
    confirm_commands (1.4.79, 사용법 단순성 2번: 사용자가 고르면 LLM이 명령에 끼워 넣지 않고 그 번호의 명령을 실행. 사용자가 고른
    뒤에만 실행하므로 next_commands가 아님, 덮어쓰기·삭제·발송이면 destructive).
    values: 명령에 넣을 값 (보이는 글 candidates와 다를 때, 예: 메일 폴더 전체 경로 대신 끝 이름).
    option: 명령에 없던 옵션에 후보 값을 넣을 때 (replace 없이)."""
    cands = [str(c) for c in candidates]
    shown = cands[:MAX_CANDIDATES]
    more = f" 외 {len(cands) - len(shown)}개" if len(cands) > len(shown) else ""
    vals = [str(v) for v in values][:len(shown)] if values is not None else shown
    if replace and len(vals) == len(shown):
        cmds = [retry_replacing(replace, v, option) for v in vals]
    elif option and len(vals) == len(shown):
        cmds = [retry_command({option: v}) for v in vals]
    else:
        cmds = []
    if cmds and all(c.startswith("python ") for c in cmds):
        tail = " [후보] " + " ".join(f"{i}) {c}" for i, c in enumerate(shown, 1)) + more
        argv = list(CURRENT_ARGV or [])
        risky = "--overwrite" in argv or bool(set(argv) & {"delete", "clear", "send"}) or (argv[:1] and argv[0] in DESTRUCTIVE_COMMANDS)
        nexts = [{"purpose": (f"'{replace}' 대신 '{c}'" if replace else f"{option} {v}") + " (사용자가 고른 경우에만)", "command": cmd,
                  **({"destructive": True} if risky else {})}
                 for c, v, cmd in zip(shown, vals, cmds)]
        return InputError(message + tail, candidates=shown, fields={"confirm_commands": nexts},
                          log_message=(message if log_message is None else log_message) + _omitted_candidates(len(cands)))
    tail = f" [후보] {', '.join(shown)}" + more
    return InputError(message + tail, candidates=shown,
                      log_message=(message if log_message is None else log_message) + _omitted_candidates(len(cands)))


class UserDecisionError(InputError):
    """도구가 대신 정하면 안 되는 일 (예: 없는 시트를 새로 만들지). 재시도하지 않고 사용자에게 묻게 합니다.
    message에는 사실만 적고, 할 일(directive)은 사용자 동의를 조건으로 적습니다. 경량 LLM이 message의 명령문을 그대로 따르기 때문."""

    def __init__(self, message: str, directive: str, candidates: Optional[List[str]] = None,
                 log_message: Optional[str] = None, commands: Optional[List[str]] = None, destructive: bool = False):
        """commands: candidates와 같은 순서의 전체 명령 (모두 python으로 시작하면 '[후보] 1) a'와 같은 번호의 confirm_commands, 1.4.83)."""
        cands = [str(c) for c in (candidates or [])][:MAX_CANDIDATES]
        cmds = list(commands or [])[:len(cands)]
        destructive = destructive or "--overwrite" in (CURRENT_ARGV or [])
        numbered = bool(cands) and len(cmds) == len(cands) and all(str(c).startswith("python ") for c in cmds)
        tail = (" [후보] " + " ".join(f"{i}) {c}" for i, c in enumerate(cands, 1)) if numbered
                else f" [후보] {', '.join(cands)}") if cands else ""
        super().__init__(message + tail, candidates=cands, no_fallback=True,
                         log_message=(message if log_message is None else log_message) + _omitted_candidates(len(cands)),
                         fields={"confirm_commands": [{"purpose": f"{c} (사용자가 고른 경우에만)", "command": cmd,
                                                       **({"destructive": True} if destructive else {})}
                                                      for c, cmd in zip(cands, cmds)]} if numbered else None)
        self.directive = directive


ELEVATED_HINT = ("이 도구가 관리자 권한으로 실행 중이라, 일반 권한으로 연 Office 문서와 창을 볼 수 없습니다 "
                 "(Windows 권한 수준이 다르면 COM 연결이 막힘). 관리자가 아닌 일반 PowerShell 창에서 OpenCode를 다시 실행하세요.")


def is_elevated() -> bool:
    """UAC로 권한을 올린 관리자 창(TokenElevationTypeFull)에서 실행 중인지. UAC가 꺼진 PC는 False. 판단하지 못하면 False."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        from ctypes import wintypes
        # UAC가 꺼진 PC(EnableLUA=0)나 기본 Administrator 계정은 모든 프로그램이 같은 권한으로 떠서
        # Office 문서도 보인다. 권한이 갈리는 것은 UAC 분리 토큰을 올린 경우(TokenElevationTypeFull=2)뿐이다.
        token = wintypes.HANDLE()
        if not ctypes.windll.advapi32.OpenProcessToken(ctypes.windll.kernel32.GetCurrentProcess(), 0x0008,
                                                       ctypes.byref(token)):
            return False
        try:
            kind = ctypes.c_int(0)
            size = wintypes.DWORD(0)
            ok = ctypes.windll.advapi32.GetTokenInformation(token, 18, ctypes.byref(kind),
                                                            ctypes.sizeof(kind), ctypes.byref(size))
            return bool(ok) and kind.value == 2
        finally:
            ctypes.windll.kernel32.CloseHandle(token)
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
            wb = retry_fs(lambda: openpyxl.load_workbook(path, read_only=True))
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
        for k, v in (exc.fields or {}).items():       # 1.4.83: 번호 [후보]의 confirm_commands
            base.setdefault(k, v)
        return base
    if isinstance(exc, InputError) and exc.candidates:
        nexts = exc.fields.get("confirm_commands") if exc.fields else None
        base.update(error_type="ambiguous", message=msg, candidates=exc.candidates,
                    agent_directive=("STOP_RETRY: 후보가 여러 개라 도구가 고를 수 없습니다. 재시도하지 말고 message의 [후보]를 "
                                     "사용자에게 보여주고 고르게 한 뒤, 사용자가 고른 번호와 같은 번호의 confirm_commands(첫째 = "
                                     "confirm_commands[0])를 그대로 1번 실행하세요." if nexts else
                                     "STOP_RETRY: 후보가 여러 개라 도구가 고를 수 없습니다. 재시도하지 말고 message의 "
                                     "[후보]를 사용자에게 보여주고 고르게 한 뒤, 사용자가 고른 값으로 실행하세요."))
        if nexts:
            base["confirm_commands"] = nexts
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
                    agent_directive="STOP_RETRY: 인자나 대상 확인에 실패했습니다. message에 '[추천]'이 있으면 [추천] 뒤의 명령을 "
                                    "그대로 1회만 실행하고, 없으면 message를 사용자에게 그대로 보고하세요.")
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
# 도구마다 같은 이름이 내용이 아닌 옵션: office.py·ppt.py의 --to·--from은 없는 옵션이지만 잘못 쓴 인자 오류도 로그에 남으므로
# 앱 이름·슬라이드 번호(transfer --to word, slide --to 3)는 가리지 않음
REDACT_KEEP = {"office": {"--from", "--to"}, "ppt": {"--to"}}
# 도구마다 같은 이름이 내용인 옵션 (excel style --highlight '=완료:lightgreen'은 셀 값, word --highlight는 색 이름이라 남김)
REDACT_EXTRA = {"excel": {"--highlight"}}
# 오류 로그의 명령·결과에 남은 메일 주소 (outlook --folder 'hong@corp.com\받은 편지함'의 계정 이름, 2026-11-09 검토 B13)
_MAIL_ADDR = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")


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
        args = " ".join([script] + [a if _REDACTED.match(a) else quote_arg(_MAIL_ADDR.sub("<메일 계정>", a))
                                    for a in redact_argv(argv, tool)]).strip()
        lines = [f"===== {_dt.datetime.now().isoformat(timespec='seconds')} | {TOOL_NAME} {VERSION} | "
                 f"Python {platform.python_version()} | {platform.platform()}",
                 f"명령: {args}",
                 f"결과: {payload.get('error_type')} | "
                 f"{_MAIL_ADDR.sub('<메일 계정>', log_message(payload.get('message') if log_msg is None else log_msg, argv, tool))}"]
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


def time_up(reserve: float = TIME_RESERVE) -> bool:
    """명령 시간 제한에서 reserve초를 뺀 시각을 넘었는지. 이어서 할 수 없는 반복(check의 행 대조)이 남은 일을
    '못 함'으로 표시하고 결과 JSON을 내려고 씀 (이어서 할 수 있는 일은 check_time -> CONTINUE)."""
    limit = time_limit()
    return limit != float("inf") and time.monotonic() - RUN_START > limit - reserve


# 도구가 같은 명령의 CONTINUE 횟수를 셈 (1.4.79, AGENTS.md "사용법 단순성" 5: LLM은 세지 않음):
# <데이터 폴더>/continue_state.json {명령 열쇠: {n, progress, t}}. CONTINUE_MAX번을 넘거나 바로 앞과 같은 progress면
# CONTINUE 대신 STOP_RETRY. 명령이 다른 결과(ok, 오류)로 끝나면 그 항목을 지움. 하루 지난 항목은 버림.
CONTINUE_MAX = 300
CONTINUE_STATE_AGE = 24 * 3600.0


def _continue_key(tool: str) -> str:
    import hashlib
    return hashlib.sha1((tool + " " + " ".join(CURRENT_ARGV or [])).encode("utf-8", "replace")).hexdigest()[:16]


def _continue_state(update: Optional[Callable[[Dict[str, Any]], None]] = None) -> Dict[str, Any]:
    """CONTINUE 횟수를 읽고 (update가 있으면) 고쳐 씀. 실패하지 않음: 폴더에 못 쓰면 세지 않음
    (데이터 폴더를 만들 수 없어도: 파일 경로이거나 읽기 전용이면 전에는 traceback만 나고 stdout에 JSON이 없었음, 2026-11-09 검토 B1)."""
    try:
        path = os.path.join(get_data_dir(), "continue_state.json")
    except Exception:
        return {}
    try:
        state = json.loads(read_text_file(path)) if os.path.isfile(path) else {}
        if not isinstance(state, dict):
            state = {}
    except Exception:
        state = {}
    now = time.time()
    state = {k: v for k, v in state.items() if isinstance(v, dict) and now - float(v.get("t", 0)) < CONTINUE_STATE_AGE}
    if update is not None:
        update(state)
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False)
        except Exception:
            pass
    return state


def continue_clear(tool: str) -> None:
    try:
        key = _continue_key(tool)
        if key in _continue_state():
            _continue_state(lambda st: st.pop(key, None))
    except Exception:
        pass


def unfinished_payload(exc: "Unfinished", tool: str, command: str) -> Dict[str, Any]:
    argv = list(CURRENT_ARGV or [])
    key = _continue_key(tool)
    prev = _continue_state().get(key) or {}
    n = int(prev.get("n", 0)) + 1
    if n > CONTINUE_MAX or prev.get("progress") == exc.progress:
        _continue_state(lambda st: st.pop(key, None))
        why = (f"같은 명령을 {n - 1}번 이어서 실행했지만 끝나지 않았습니다" if n > CONTINUE_MAX else
               "바로 앞 실행과 같은 곳에서 멈춰 더 이어서 해도 끝나지 않습니다")
        return {"status": "error", "tool": tool, "command": command, "error_type": "unfinished_limit", "progress": exc.progress,
                "message": f"{why} (진행: {exc.progress}). 지금까지 한 일은 저장되어 있습니다.",
                "agent_directive": "STOP_RETRY: 다시 실행하지 말고 message와 progress를 사용자에게 그대로 알리세요."}
    _continue_state(lambda st: st.__setitem__(key, {"n": n, "progress": exc.progress, "t": time.time()}))
    nexts = [next_command(f"{tool}.py", argv, "같은 명령으로 이어서 진행 (끝낸 부분은 저장됨, 묻지 않고 그대로 실행)")] if argv else []
    payload: Dict[str, Any] = {
        "status": "ok", "tool": tool, "command": command, "preview": f"진행 중: {exc.progress}",
        "verified": False, "continue": True, "progress": exc.progress, "next_commands": nexts, "continue_runs": n,
        "agent_directive": ("CONTINUE: 아직 끝나지 않았습니다. 명령 시간 제한 때문에 나눠서 진행하며, 지금까지 한 일은 저장되어 있습니다. "
                            + ("next_commands[0]의 command를" if nexts else "방금 실행한 명령을")
                            + " 바꾸지 말고 그대로 다시 실행하세요 (사용자에게 묻지 않음, 몇 번째인지 세지 않음: 멈출 때는 도구가 알림). "
                            "끝나면 VERIFIED 결과가 나옵니다."),
    }
    _attach_recovered(payload)
    return payload


def run_cli(tool: str, command: str, execute: Callable[[], Dict[str, Any]]) -> None:
    """도구 본문을 실행하고, 어떤 경우에도 stdout 1줄 JSON과 종료 코드(0/1/130)를 보장합니다.
    실패(exit 1)는 <데이터 폴더>/logs/errors.txt에 자동 기록합니다 (사용자 중단 130과 성공은 기록하지 않음).
    시간 제한으로 멈춘 긴 작업(Unfinished)은 CONTINUE 결과(exit 0, 오류 로그 없음, R10).
    뒷정리(CONTINUE 횟수, 오류 로그, 출력)가 실패해도 가장 바깥에서 한 번 더 막아 JSON 1줄을 냅니다 (2026-11-09 검토 B1)."""
    try:
        _run_cli(tool, command, execute)
    except SystemExit:
        raise
    except BaseException as e:                  # noqa: BLE001  마지막 방어선
        try:
            sys.stdout.write(json.dumps({"status": "error", "tool": tool, "command": command, "error_type": "internal",
                                         "message": f"출력 중 내부 오류: {type(e).__name__}",
                                         "agent_directive": "ERROR: 내부 오류입니다. 다시 실행하지 말고 사용자에게 알리세요."},
                                        ensure_ascii=True) + "\n")
            sys.stdout.flush()
        except Exception:
            pass
        sys.exit(130 if isinstance(e, KeyboardInterrupt) else 1)


def _run_cli(tool: str, command: str, execute: Callable[[], Dict[str, Any]]) -> None:
    global RUN_START
    RUN_START = time.monotonic()
    setup_console()
    try:
        result = execute()
    except Unfinished as u:
        log("STATUS", f"시간 제한으로 여기서 멈춤 (저장됨, 같은 명령으로 이어서): {u.progress}")
        payload = unfinished_payload(u, tool, command)
        if payload["status"] != "ok":
            write_error_log(tool, None, payload)
        emit_json(payload)
        sys.exit(0 if payload["status"] == "ok" else 1)
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
        continue_clear(tool)
        # traceback은 내부 오류(ERROR:)일 때만 남깁니다
        internal = str(payload.get("agent_directive", "")).startswith("ERROR:")
        write_error_log(tool, None, payload, trace if internal else "", log_msg=_exc_log_message(e, payload))
        emit_json(payload)
        sys.exit(1)
    log_msg = result.pop(LOG_MESSAGE_KEY, None) if isinstance(result, dict) else None
    continue_clear(tool)
    if result.get("status") != "ok":
        write_error_log(tool, None, result, log_msg=log_msg)
    emit_json(result)
    sys.exit(0 if result.get("status") == "ok" else 1)


# 경량 LLM이 자주 지어내는 명령·옵션 (실사용 오류 로그에서 본 것). 실제 이름으로 안내합니다.
COMMAND_ALIASES = {"excel": {"list": "info", "sheets": "info"}, "word": {"list": "info", "toc": "outline"},
                   "ppt": {"list": "info"}, "office": {"list": "windows", "redo": "undo"},
                   "outlook": {"attachment": "attachments", "attach": "attachments", "download": "attachments"}}
# 같은 뜻이라 바꿔 실행해도 되는 옵션 (도구, 명령, 지어낸 옵션) -> 바꿀 토큰 ("*" = 모든 도구·명령, [] = 빼기).
# 바꿀 옵션이 그 명령에 있을 때만 씀. 값은 그대로 뒤에 붙고, 고친 전체 명령을 [추천]으로 줌 (Gemma 기록 2026-11-06 감사)
SAME_MEANING = [
    ("*", "*", "--file", ["--target"]), ("*", "*", "--path", ["--target"]), ("*", "*", "--input", ["--target"]),
    ("reader", "*", "--target", ["--file"]), ("reader", "*", "--path", ["--file"]), ("reader", "*", "--input", ["--file"]),
    ("office", "undo", "--target", ["--app"]),
    ("*", "*", "--output", ["--save"]), ("*", "*", "--out", ["--save"]), ("*", "*", "--save-as", ["--save"]),
    ("*", "*", "--sheet-name", ["--sheet"]), ("excel", "*", "--cell", ["--range"]), ("*", "*", "--rows", ["--max-rows"]),
    ("excel", "write", "--text", ["--value"]), ("excel", "read", "--text", ["--search"]),
    ("word", "write", "--append", ["--pos", "end"]), ("word", "write", "--position", ["--pos"]), ("word", "write", "--at", ["--pos"]),
    ("outlook", "*", "--text", ["--body"]), ("outlook", "*", "--max", ["--limit"]),
    # 옵션 하나 = 뜻 하나 (사용법 단순성 4번): --to·--from은 outlook의 받는 사람·보낸 사람만. 앱은 --to-app·--from-app, 위치는 --dest
    ("office", "transfer", "--to", ["--to-app"]), ("office", "transfer", "--from", ["--from-app"]),
    ("ppt", "slide", "--to", ["--dest"]),
    ("*", "*", "--json", []), ("*", "*", "--encoding", []),      # 출력은 늘 JSON, 인코딩은 자동 판별
]
# 한 글자 옵션은 없음 (1.4.88, 사용법 단순성 1번: 한 가지 일에 한 가지 사용법). 받지 않고 긴 이름으로 고친 전체 명령을 [추천]
# (같은 뜻 표). 도구마다 뜻이 달랐음: -n은 outlook --limit, reader --top, excel --name / -r은 excel --range, reader --regex
SHORT_OPTIONS = {
    "excel": {"-a": "--action", "-c": "--columns", "-f": "--file", "-n": "--name", "-r": "--range", "-s": "--sheet", "-t": "--target"},
    "word": {"-a": "--action", "-f": "--file", "-t": "--target"},
    "ppt": {"-a": "--action", "-f": "--file", "-t": "--target"},
    "outlook": {"-n": "--limit", "-q": "--query"},
    "reader": {"-C": "--context", "-d": "--folder", "-f": "--file", "-n": "--top", "-p": "--pages", "-q": "--query", "-r": "--regex"},
}
SAME_MEANING += [(tool, "*", short, [long]) for tool, opts in SHORT_OPTIONS.items() for short, long in opts.items()]
_SHORT_FLAG = re.compile(r"^-[A-Za-z]$")


def _is_flag(arg: str) -> bool:
    """인자 오류에서 옵션으로 볼 것: --이름, 한 글자 -x (없앤 짧은 옵션을 같은 뜻 표로 고치려고. 음수 -5는 값)."""
    return arg.startswith("--") or bool(_SHORT_FLAG.match(arg))


# 철자가 비슷해도 추천하지 않는 옵션 (뜻이 크게 다르거나 되돌리기 어려움: --cell -> --all이 캐시 전부 삭제·전체 답장이었음)
NO_FUZZY_OPTIONS = {"--all", "--overwrite", "--to", "--cc", "--bcc", "--attach", "--target", "--delete", "--send", "--app"}
# 되돌리기 어려운 명령 (메일 발송, 캐시 지우기, 되돌리기): 오타 추천을 [추천](LLM이 묻지 않고 1번 실행)이 아니라
# [후보](사용자에게 물음)로 줌 (1.4.75, 11차 검토: 'sent', 'sned', 'end' -> send)
DESTRUCTIVE_COMMANDS = {"send", "clear", "undo", "delete"}
# 읽기만 하는 명령·동작: 오타가 이것과 가장 가까우면 [추천] (잘못 골라도 문서가 바뀌지 않음, 1.4.83)
def choices_in_message(message: str) -> list:
    """argparse 'invalid choice' 오류의 고를 수 있는 값. Python 3.13.13은 값에 따옴표를 붙이지 않음
    ('choose from list, add'), 그 전 판은 붙임 ("choose from 'list', 'add'"). 둘 다 읽음 (1.4.84)."""
    m = re.search(r"choose from (.+?)\)?\s*$", message)
    if not m:
        return []
    return [c.strip().strip("'\"") for c in m.group(1).split(",") if c.strip().strip("'\"")]


READ_ONLY_CHOICES = {"read", "info", "outline", "list", "search", "windows", "current", "revisions"}
# 다른 파일의 명령이 맞는 경우: [추천]으로 값만 바꾸면 또 틀리므로 맞는 명령을 안내만 함
_WRITE_TEXT = ('글을 문서에 넣으려면 열린 워드는 word.py write --text "내용", PPT는 ppt.py write --text "내용", '
               '엑셀 칸은 excel.py write --range A1 --value "내용"을 쓰세요.')
OTHER_COMMAND = {("office", "write"): "office.py에는 write가 없습니다. " + _WRITE_TEXT,
                 ("office", "transfer", "--text"): "transfer는 엑셀 차트·범위와 문서의 표를 옮기는 명령이라 글은 받지 않습니다. " + _WRITE_TEXT}
# 다른 명령의 옵션으로 찾기를 시킴: (도구, 명령, 옵션) -> (맞는 명령, 맞는 옵션). 바꾼 명령이 읽기만 하는 명령일 때만 [추천]
# (READ_ONLY_CHOICES: 잘못 골라도 바뀌는 것이 없음). 회사 PC Gemma 2026-11-13 O7: reader.py read --search "-20 ℃"가 옵션 목록만 받음
SWAP_COMMAND = {("reader", "read", "--search"): ("search", "--query"), ("reader", "read", "--query"): ("search", "--query")}
# excel comment --cells "C3:OK,C4:NG" (회사 PC Gemma 2026-11-13 O1): 칸마다 다른 메모는 CSV 파일로
OTHER_COMMAND[("excel", "comment", "--cells")] = (
    "comment에는 --cells가 없습니다. 한 칸은 --action add --range C3 --text \"<메모>\", 칸마다 다른 메모 여러 개는 "
    "파일 쓰기 도구로 CSV(열: 칸, 메모)를 C:\\AIKIT\\data에 만든 뒤 --action add --file \"<메모.csv>\"를 쓰세요. "
    "판정·검토 의견은 사용자가 달아 달라고 했을 때만 넣고, 아니면 넣기 전에 사용자에게 보여 주고 \"예\"를 받으세요.")
# 1.4.72에서 없앤 outlook --send (1.4.73): 초안을 만든 뒤 사용자 확인을 받아 send로 보냄. {cmd} = --send를 뺀 같은 명령
_NO_SEND = ("--send는 없습니다. 초안을 만든 뒤 사용자 확인을 받아 send를 실행합니다(결과의 confirm_commands). "
            "[추천] {cmd}")
OTHER_COMMAND.update({("outlook", c, "--send"): _NO_SEND for c in ("new", "reply", "forward")})


# 도구 파일 -> 그 명령을 설명하는 작업별 스킬 (<설치 폴더>\.opencode\skills\<이름>\SKILL.md). excel.py comment는 search-kit에도 있음
SKILL_OF_TOOL = {"excel": "excel-kit", "word": "word-kit", "ppt": "ppt-kit", "office": "word-kit",
                 "outlook": "outlook-kit", "reader": "search-kit"}
# 도구 파일과 다른 스킬에 있는 명령 (표 집계·중복 찾기는 stats-kit)
SKILL_OF_COMMAND = {("excel", "stats"): "stats-kit", ("excel", "duplicates"): "stats-kit",
                    ("excel", "style"): "excel-format-kit", ("excel", "sheet"): "excel-format-kit",
                    ("excel", "range"): "excel-format-kit", ("excel", "chart"): "excel-format-kit"}


def skill_name(tool: str) -> str:
    """지금 실행한 명령을 설명하는 스킬 이름 (도움말·인자 오류 안내). 모르는 도구(테스트용 등)는 word-kit."""
    cmd = CURRENT_ARGV[0] if CURRENT_ARGV and CURRENT_TOOL == tool else ""
    return SKILL_OF_COMMAND.get((tool, cmd)) or SKILL_OF_TOOL.get(tool, "word-kit")


def skill_file(tool: str) -> str:
    """도구(지금 실행한 명령)의 스킬 파일 경로: 도구 폴더의 부모(C:\\AIKIT, OpenCode 프로젝트)의 .opencode\\skills (1.6.7)."""
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".opencode", "skills", skill_name(tool), "SKILL.md")


# 글 값을 받는 옵션: 값이 '-'로 시작해도 값으로 받음 (JsonArgumentParser._join_dash_values)
DASH_VALUE_OPTIONS = {"--query", "--body", "--value", "--text", "--find", "--replace", "--subject", "--title", "--search",
                      "--note", "--anchor", "--callout", "--header", "--where", "--from", "--sender"}


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
        global CURRENT_ARGV, CURRENT_TOOL
        CURRENT_ARGV = [str(a) for a in args] if args is not None else sys.argv[1:]
        CURRENT_TOOL = self.tool_name
        return super().parse_args(self._join_dash_values(list(CURRENT_ARGV)), namespace)

    def _option_names(self) -> set:
        names = set()
        for act in self._actions:
            names.update(act.option_strings)
            if isinstance(act, argparse._SubParsersAction):
                for sub in act.choices.values():
                    names.update(o for a in sub._actions for o in a.option_strings)
        return names

    def _join_dash_values(self, argv: List[str]) -> List[str]:
        """글 값을 받는 옵션 바로 뒤의 '-'로 시작하는 값('--query -40℃', '--body -확인')을 '--query=-40℃'로 붙임.
        argparse는 그 값을 옵션으로 보고 'expected one argument'로 멈췄음 (2026-11-09 검토 B5). 도구의 옵션 이름이면 붙이지 않음."""
        names = self._option_names()
        out: List[str] = []
        i = 0
        while i < len(argv):
            a = argv[i]
            if (a in DASH_VALUE_OPTIONS and a in names and i + 1 < len(argv) and argv[i + 1].startswith("-")
                    and len(argv[i + 1]) > 1 and argv[i + 1].split("=", 1)[0] not in names):
                out.append(f"{a}={argv[i + 1]}")
                i += 2
                continue
            out.append(a)
            i += 1
        return out

    def print_help(self, file=None) -> None:  # type: ignore[override]
        """R7: 도움말 텍스트는 stderr로, stdout에는 JSON 1줄 (argparse가 이어서 exit 0).
        도움말 글을 만들다 실패해도(옵션 설명의 '%' 등) JSON 1줄은 냄 (1.4.78: excel style·replace -h가 traceback만 냈음)."""
        try:
            super().print_help(sys.stderr)
        except Exception as e:                      # 도움말이 깨져도 계약(stdout JSON 1줄)은 지킴
            sys.stderr.write(f"[WARN] 도움말을 만들지 못했습니다 ({type(e).__name__}). SKILL.md를 보세요.\n")
        emit_json({"status": "ok", "tool": self.tool_name, "command": "help", "verified": True,
                   "skill_file": skill_file(self.tool_name),
                   "agent_directive": f"VERIFIED: 도움말을 stderr에 출력했습니다. {skill_name(self.tool_name)} "
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
            choices = list(subs) or choices_in_message(message)
            alias = COMMAND_ALIASES.get(self.tool_name, {}).get(m.group(1).lower())
            if alias and alias not in DESTRUCTIVE_COMMANDS:       # 같은 뜻 표에 있을 때만 [추천] (설계 권장 7번)
                return f" [추천] {retry_replacing(m.group(1), alias)}"
            # 철자만 비슷하면 [후보] (1.4.83: 'sheet --action del' -> delete, 'range --action erase' -> merge를 묻지 않고 실행하게 했음).
            # 가장 가까운 것이 읽기만 하는 명령이면 잘못 골라도 바뀌는 것이 없어 [추천] ('raed' -> read)
            near = ([alias] if alias else []) + [c for c in difflib.get_close_matches(m.group(1), choices, n=3, cutoff=0.6)
                                                 if c != alias]
            if near and not alias and near[0] in READ_ONLY_CHOICES:
                return f" [추천] {retry_replacing(m.group(1), near[0])}"
            warn = (" 되돌리기 어려운 동작이 있으니" if any(c in DESTRUCTIVE_COMMANDS for c in near) else "")
            return (f" 철자만 비슷한 값이라{warn} 무엇을 하려던 것인지 사용자에게 물으세요." if near else "") \
                + self._numbered(m.group(1), near or choices)
        m = re.search(r"unrecognized arguments: (.+)$", message)
        req = re.search(r"the following arguments are required: (.+)$", message)
        if not m and not req:
            return ""
        cmd = CURRENT_ARGV[0] if CURRENT_ARGV else ""
        sub = subs.get(cmd) if subs else self
        opts = [o for a in (sub._actions if sub else []) for o in a.option_strings if o.startswith("--")]
        known = {o for a in (sub._actions if sub else []) for o in a.option_strings}
        self._valued = {o for a in (sub._actions if sub else []) for o in a.option_strings if a.nargs not in (0, "?")}
        argv = list(CURRENT_ARGV or [])
        flags = [x for x in argv[1:] if _is_flag(x) and "=" not in x and x not in known]
        if req:
            # 필수 옵션이 없다는 오류: 지어낸 옵션이 그 필수 옵션과 같은 뜻이면 고친 명령 (reader read --target a.pdf -> --file)
            fixed = self._same_meaning(argv, flags, opts, known) if flags else None
            if fixed is not None:
                return " " + fixed
            return ""
        extra = m.group(1).split()
        flags = [x for x in extra if _is_flag(x)]
        for f in flags:
            key = f.split("=", 1)[0]       # '--query=-20 ℃': 음수로 시작하는 값은 옵션에 붙어 옴
            swap = SWAP_COMMAND.get((self.tool_name, cmd, key))
            new_sub = subs.get(swap[0]) if swap and subs else None
            if swap and new_sub is not None and swap[0] in READ_ONLY_CHOICES:
                new_known = {o for a in new_sub._actions for o in a.option_strings}
                rest, skip = [swap[0]], False
                for a in argv[1:]:
                    if skip:              # 새 명령에 없는 옵션의 값
                        skip = False
                        continue
                    if _is_flag(a) and a.split("=", 1)[0] != key and a.split("=", 1)[0] not in new_known:
                        skip = "=" not in a and a in self._valued
                        continue
                    if a.split("=", 1)[0] == key:
                        rest += [swap[1]] + ([a.split("=", 1)[1]] if "=" in a else [])
                    else:
                        rest.append(a)
                return (f" '{cmd}' 명령에는 {key} 옵션이 없습니다. 파일 안에서 찾는 명령은 {swap[0]} {swap[1]}입니다. "
                        f"[추천] {next_command(f'{self.tool_name}.py', rest, '')['command']}")
            other = OTHER_COMMAND.get((self.tool_name, cmd, f))
            if other:
                if "{cmd}" in other:      # 받지 않는 옵션(과 그 값)을 모두 뺀 같은 전체 명령 ('--send --dry-run' 둘 다, 1.4.75)
                    rest, drop = [], False
                    for a in (CURRENT_ARGV or []):
                        if a in flags:
                            drop = True
                            continue
                        if drop and not a.startswith("-") and a in extra:
                            drop = False
                            continue
                        drop = False
                        rest.append(a)
                    other = other.replace("{cmd}", next_command(f"{self.tool_name}.py", rest, "")["command"])
                return " " + other
        fixed = self._same_meaning(argv, flags, opts, known) if flags else None
        if fixed is not None:
            return " " + fixed
        for f in flags:
            near = [o for o in difflib.get_close_matches(f, opts, n=1, cutoff=0.85)
                    if o not in NO_FUZZY_OPTIONS and o[2:5] == f[2:5]]
            if near:
                # 철자만 비슷한 옵션은 뜻이 다를 수 있어 묻게 함 ([추천]이면 LLM이 묻지 않고 실행)
                return f" '{cmd}' 명령에는 {f} 옵션이 없습니다. 비슷한 옵션이 맞는지 사용자에게 물으세요." + self._numbered(f, near)
        if extra and all(x in RESULT_WORDS for x in extra):
            # 'read ... rows': 스킬의 '행 수는 결과의 rows'를 인자로 읽음 -> 그 글을 뺀 같은 명령 (2026-11-09 검토 B3)
            rest = [a for a in argv if a not in extra]
            return (f" {', '.join(extra)}은(는) 옵션이 아니라 결과 JSON의 키입니다 (행 수는 결과의 rows). "
                    f"[추천] {next_command(f'{self.tool_name}.py', rest, '')['command']}")
        if not flags:
            text_opts = [o for o in ("--body", "--query", "--text", "--value") if o in opts]
            if text_opts and not any(re.search(r"[\\/:]", x) for x in extra):
                # 'reply --id 3 "확인했습니다"': 빠진 것은 따옴표가 아니라 옵션 이름 (검토 B13)
                return (f" 옵션 이름 없이 준 글이 있습니다. 글은 옵션 이름과 함께 주세요: {', '.join(text_opts)} \"<글>\" "
                        f"(예: {text_opts[0]} \"확인했습니다\").")
            return " 띄어쓰기가 있는 경로나 값은 큰따옴표로 감싸세요 (예: --target \"C:\\폴더 이름\\파일.xlsx\")."
        opts = [o for o in opts if o != "--help"]
        return f" '{cmd}' 명령에는 {', '.join(flags)} 옵션이 없습니다. 옵션: {', '.join(opts)}" if opts else ""

    def _numbered(self, old: str, cands: List[str], destructive: bool = False) -> str:
        """' [후보] 1) a 2) b'와 같은 번호의 confirm_commands (old를 후보로 바꾼 전체 명령, 사용법 단순성 2번).
        전체 명령을 만들 수 없으면 번호 없는 '[후보] a, b'."""
        shown = [str(c) for c in cands][:MAX_CANDIDATES]
        cmds = [retry_replacing(old, c) for c in shown]
        if not shown or not all(c.startswith("python ") for c in cmds):
            return f" [후보] {', '.join(shown)}"
        self._confirm = [{"purpose": f"'{old}' 대신 '{c}' (사용자가 고른 경우에만)", "command": cmd,
                          **({"destructive": True} if destructive or c in DESTRUCTIVE_COMMANDS else {})}
                         for c, cmd in zip(shown, cmds)]
        return " [후보] " + " ".join(f"{i}) {c}" for i, c in enumerate(shown, 1)) + " (사용자가 고른 번호의 confirm_commands를 실행)"

    def _same_meaning(self, argv: List[str], flags: List[str], opts: List[str], known: set) -> Optional[str]:
        """지어낸 옵션을 모두 같은 뜻의 옵션으로 바꿀 수 있으면 '... [추천] <고친 전체 명령>', 하나라도 못 바꾸면 None."""
        cmd = argv[0] if argv else ""
        plan: Dict[str, List[str]] = {}
        for f in flags:
            repl = next((r for t, c, o, r in SAME_MEANING if o == f and t in ("*", self.tool_name) and c in ("*", cmd)
                         and (not r or r[0] in opts)), None)
            if repl is None:
                return None
            plan[f] = repl
        out: List[str] = []
        i = 0
        while i < len(argv):
            a = argv[i]
            if a in plan:
                has_value = i + 1 < len(argv) and not _is_flag(argv[i + 1]) and argv[i + 1] not in known
                repl = plan[a]
                if len(repl) == 1 and not has_value and repl[0] in getattr(self, "_valued", set()):
                    # 값이 필요한 옵션으로 바꾸는데 값이 없음 ('read --rows' -> '--max-rows'는 실행하면 인자 오류, 검토 B3)
                    if a not in VALUELESS_DROP:
                        return None
                    i += 1                      # 결과 키를 옵션처럼 쓴 것: 빼고 같은 명령
                    plan[a] = []
                    continue
                if repl:
                    out += repl
                    if has_value and len(repl) == 1:
                        out.append(argv[i + 1])
                    i += 2 if has_value and len(repl) == 1 else 1     # --append처럼 바꾼 토큰이 둘이면 뒤의 글은 그대로 둠
                else:
                    i += 2 if has_value else 1                # 빼는 옵션은 값도 뺌
                continue
            out.append(a)
            i += 1
        what = ", ".join(f"{f} -> {' '.join(r) if r else '(뺌)'}" for f, r in plan.items())
        if any(f in VALUELESS_DROP and not r for f, r in plan.items()):
            what += ". 행 수는 결과의 rows 값"
        return (f"'{cmd}' 명령에는 {', '.join(plan)} 옵션이 없어 같은 뜻으로 고쳤습니다 ({what}). "
                f"[추천] {next_command(f'{self.tool_name}.py', out, '')['command']}")

    def error(self, message: str) -> None:  # type: ignore[override]
        self.print_usage(sys.stderr)
        self._confirm: List[Dict[str, Any]] = []
        nope = not_supported_hint(self.tool_name, list(CURRENT_ARGV or []))
        if nope:
            what, where, kind = nope
            if kind == "user":
                payload = {"status": "error", "tool": self.tool_name, "command": "", "error_type": "not_supported",
                           "message": f"{what}은(는) 도구로 하지 않는 일입니다. 사용자가 직접 하세요: {where}.",
                           "agent_directive": "STOP_RETRY: 도구로 하지 않는 일입니다. 다시 실행하거나 파이썬 스크립트로 대신하지 말고, "
                                              "message의 방법(사용자가 직접 하는 곳)을 사용자에게 그대로 알리세요."}
            else:
                how = "전역 지침 8장 '가. 확인 뒤 하는 일'대로, 바꿀 내용을 보여 주고 사용자의 '예'를 받은 뒤 파이썬 스크립트로 하세요 (저장은 하지 않음)"
                payload = {"status": "error", "tool": self.tool_name, "command": "", "error_type": "not_in_tool",
                           "message": f"{what}은(는) 이 도구에 명령이 없습니다. {how}.",
                           "agent_directive": f"STOP_RETRY: 도구에 명령이 없습니다. 같은 명령을 다시 실행하지 말고 {how}."}
            write_error_log(self.tool_name, None, payload)
            emit_json(payload)
            sys.exit(2)
        hint = self._hint(message)
        payload = {
            "status": "error",
            "tool": self.tool_name,
            "command": "",
            "error_type": "input",
            "message": f"CLI 인자 오류: {message}." + hint,
            "agent_directive": "STOP_RETRY: CLI 인자가 잘못되었습니다. message에 [추천]이 있으면 [추천] 뒤의 명령을 그대로 1번만 실행하고, "
                               "[후보]가 있으면 후보를 사용자에게 보여 주고 고른 번호의 confirm_commands를 1번 실행하고, "
                               f"둘 다 없으면 {skill_name(self.tool_name)} 스킬(SKILL.md)의 명령 형식을 확인해 사용자에게 보고하세요.",
        }
        if self._confirm:
            payload["confirm_commands"] = self._confirm
        # 옵션 이름 없이 준 글(본문·검색어일 수 있음)은 오류 로그의 명령·결과에 남기지 않음 (검토 B13)
        m_un = re.search(r"unrecognized arguments: (.+)$", message)
        unrec = f" {m_un.group(1)} " if m_un else ""
        stray = [a for a in (CURRENT_ARGV or [])[1:] if a and not a.startswith("-") and f" {a} " in unrec]
        argv = [f"<{len(a)}자 생략>" if a in stray else a for a in (CURRENT_ARGV or [])]
        log_msg = payload["message"]
        for x in stray:
            log_msg = log_msg.replace(x, f"<{len(x)}자 생략>")
        write_error_log(self.tool_name, argv if stray else None, payload, log_msg=log_msg if stray else None)
        emit_json(payload)
        sys.exit(2)


# 도구에 명령이 없는 기능 (사용자 결정 2026-11-10, 둘로 나눔 2026-11-12, 전역 지침 8장): 인자 오류가 이 기능을 찾는 명령이면 안내.
# (도구, 낱말들, 기능, 사용자가 할 곳, 종류). 종류: confirm = 8장 가(사용자 확인 뒤 파이썬), user = 8장 나(하지 않음, 사용자가 직접).
# 두 판 비교는 넣지 않음 (나중에 구현할 때 검토, 사용자 결정 2026-11-12).
# 낱말은 명령 이름·옵션 이름·--action/--type 값에서만 찾음 (검색어 같은 값에서는 찾지 않음)
NOT_SUPPORTED = [
    ("word", ("accept", "reject", "수락", "거부"), "변경 내용 수락·거부", "", "confirm"),
    ("excel", ("pivot", "피벗"), "피벗 테이블", "", "confirm"),
    ("excel", ("validation", "유효성"), "데이터 유효성(목록)", "", "confirm"),
    ("excel", ("edit", "modify", "update", "format"), "기존 차트 고치기", "", "confirm"),
    ("", ("print", "printout", "page-setup", "인쇄"), "인쇄", "그 앱의 파일 > 인쇄", "user"),
    ("outlook", ("meeting", "invite", "attendee", "attendees", "회의", "참석자"), "회의 요청(참석자 초대)", "Outlook 일정 > 새 모임 (내 일정만은 outlook.py event)", "user"),
    ("outlook", ("move", "이동", "category", "categories", "분류"), "메일 이동·분류", "Outlook에서 메일을 폴더로 끌어 놓기, 홈 탭 > 분류", "user"),
    ("outlook", ("flag", "플래그", "follow-up"), "메일 플래그", "Outlook 메일 목록의 깃발 아이콘", "user"),
]


def not_supported_hint(tool: str, argv: Sequence[str]) -> Optional[Tuple[str, str, str]]:
    """인자 오류가 난 명령이 도구에 명령이 없는 기능을 찾는지: (기능, 사용자가 할 곳, 종류) 또는 None."""
    words: List[str] = []
    for i, a in enumerate(argv):
        key = str(a).split("=", 1)[0]
        if i == 0 or key.startswith("--"):
            words.append(key.lstrip("-").casefold())
        if i > 0 and str(argv[i - 1]) in ("--action", "--type"):
            words.append(str(a).casefold())
    for t, keys, what, where, kind in NOT_SUPPORTED:
        if t and t != tool:
            continue
        if t == "excel" and what == "기존 차트 고치기" and not (argv and str(argv[0]) == "chart"):
            continue
        if any(w == k or w.startswith(k + "-") or w.endswith("-" + k) for w in words for k in keys):
            return what, where, kind
    return None


# 결과 JSON의 키를 옵션·인자처럼 쓴 것 (스킬의 "행 수는 결과의 rows"): 값 없이 오면 빼고 [추천] (2026-11-09 검토 B3)
VALUELESS_DROP = {"--rows", "--count"}
RESULT_WORDS = {"rows", "count", "values", "cols"}


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


def get_result_dir() -> str:
    """대상 파일이 없는 작업(메일, 저장한 적 없는 문서, 폴더·캐시 검색)의 결과 파일 기본 폴더: C:\\AIKIT\\data (전역 지침 8장).
    우선순위: OFFICE_CLI_RESULT_DIR > En-Genius면 데이터 폴더 > <설치 폴더(도구 폴더의 부모)>\\data."""
    return _dir_setting("OFFICE_CLI_RESULT_DIR", "result_dir",
                        get_data_dir() if engenius_home() else os.path.join(os.path.dirname(TOOL_DIR), "data"))


def output_folder(near: Optional[str] = None) -> str:
    """결과 파일을 둘 폴더: 대상 파일(near)이 있으면 그 파일과 같은 폴더, 없으면 get_result_dir()."""
    if near and os.path.isfile(str(near)):
        return os.path.dirname(os.path.abspath(str(near)))
    return get_result_dir()


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
                if os.path.isfile(p) and os.path.getmtime(p) < cutoff and remove_file(p):
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
    # 전체 경로로 준 파일은 그 폴더에서만 비슷한 이름을 찾음 (다른 폴더의 같은 이름 파일을 [추천]하지 않음)
    for folder in ((os.path.dirname(raw),) if os.path.isabs(raw) else (get_data_dir(), get_work_dir())):
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


def resolve_output_path(path: Optional[str], default_stem: str, ext: str, overwrite: bool = False,
                        near: Optional[str] = None, internal: bool = False) -> str:
    """출력 파일 경로를 정합니다 (전역 지침 8장: 경로를 주면 그대로, 아니면 대상 파일 옆, 대상 파일이 없으면 C:\\AIKIT\\data).
    - 생략하면 output_folder(near)에 '<default_stem>_<시각><ext>'로 만듭니다 (default_stem의 쓸 수 없는 글자는 _로).
      internal=True(도구가 다음 단계에 넘기는 중간 파일)는 작업 폴더(data/_work, 오래되면 지움)에.
    - 파일명만 주거나 상대 경로면 output_folder(near) 기준입니다.
    - 확장자가 없거나 다르면 ext로 맞춥니다.
    - 그 파일이 이미 있으면 덮어쓰지 않고 시각을 붙인 새 이름으로 바꿉니다 (R4, recovered로 알림). overwrite=True면 그대로.
    """
    ext = ext if ext.startswith(".") else "." + ext
    if not path:
        if internal:
            cleanup_work_dir()
        folder = get_work_dir() if internal else output_folder(near)
        os.makedirs(folder, exist_ok=True)
        # 기본 이름은 차트·시트·문서 이름에서 오므로 / ? : 같은 글자가 섞일 수 있음 -> 여기 한 곳에서 정리
        auto = os.path.join(folder, f"{safe_filename(default_stem, default='output')}_{timestamp()}{ext}")
        if os.path.exists(auto) and not overwrite:
            # 같은 초에 같은 이름으로 저장하면 앞 결과를 덮어쓰던 문제: 번호를 붙인 새 이름으로
            alt = _timestamped(auto)
            note_recovered(f"{os.path.basename(auto)} 이미 있음 (같은 초에 저장) -> {os.path.basename(alt)}로 저장")
            auto = alt
        return auto
    raw = os.path.expandvars(os.path.expanduser(str(path).strip().strip('"')))
    base, cur_ext = os.path.splitext(raw)
    if cur_ext.lower() != ext.lower():
        if cur_ext:      # 말없이 확장자를 바꾸던 것 (read --save res.xlsx -> res.csv, 2026-11-09 검토 B13): 결과가 달라진 일이라 recovered
            note_recovered(f"{os.path.basename(raw)}의 {cur_ext} 형식으로는 저장하지 않음 -> {os.path.basename(base + ext)}({ext})로 저장")
        raw = (base if cur_ext else raw) + ext
    if not os.path.isabs(raw):
        raw = os.path.join(output_folder(near), raw)
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


def _transient_fs_error(e: OSError) -> bool:
    """잠깐 기다리면 풀릴 수 있는 오류: 접근 거부, 공유 위반, 잠금 위반 (winerror 5/32/33)."""
    return isinstance(e, PermissionError) or getattr(e, "winerror", None) in (5, 32, 33)


def retry_fs(fn: Callable[[], Any], attempts: int = 6, delay: float = 0.5, deadline: Optional[float] = None,
             record: bool = True) -> Any:
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
            if i + 1 >= attempts:           # 마지막 시도 뒤에는 기다리지 않음
                break
            wait = delay * (i + 1)
            if deadline is not None:
                wait = min(wait, deadline - time.monotonic())
                if wait <= 0:
                    attempts = i + 1
                    break
            time.sleep(wait)
    assert last is not None
    name = os.path.basename(str(getattr(last, "filename", "") or ""))
    if record:
        record_recovery_attempt(f"파일 작업 {name} {attempts}회 재시도 실패 ({type(last).__name__})".replace("  ", " "))
    raise last


def remove_file(path: str) -> bool:
    """파일을 지웁니다 (백신·탐색기가 잠깐 잡고 있으면 retry_fs로 기다렸다 다시, 1.4.81). 없으면 그냥 넘어감.
    지웠으면 True, 한도를 넘겨도 못 지웠으면 False (남은 임시 파일·스냅샷이 결과를 바꾸지 않는 곳에서 씀)."""
    try:
        retry_fs(lambda: os.remove(path), attempts=4, delay=0.2, record=False)   # 최대 약 1.2초, 실패해도 오류 로그에 남기지 않음
        return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


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

# 이름 비교 규칙 (시트·열 머리글·메일 폴더 등 모든 이름 찾기, 열린 문서(COM)와 닫힌 파일(openpyxl) 같음, 1.4.87):
# - 같은 이름: exact_key (앞뒤 공백, 대소문자만 무시) -> 그 이름을 그대로 씀 (find_name)
# - 표기만 다른 이름: name_key (대소문자, 공백, _ - . 무시) -> 하나뿐이면 [추천] (not_found, missing_sheet), 열 머리글은 그대로 씀
# - 그 밖에 비슷한 이름: [후보] (suggest)

def exact_key(s: Any) -> str:
    """같은 이름인지 견주는 값: 앞뒤 공백과 대소문자만 무시."""
    return str(s).strip().casefold()


def name_key(s: Any) -> str:
    """표기만 다른 이름인지 견주는 값: 대소문자, 공백, _ - . 를 무시 ('Sheet.1' = 'sheet 1' = 'Sheet1')."""
    return re.sub(r"[\s_\-.]+", "", str(s)).casefold()


def suggest(name: str, candidates: Iterable[str], n: int = 3) -> List[str]:
    """오타로 보이는 이름에 대해 비슷한 후보를 찾습니다 (name_key로 견줌, 부분 일치 우선)."""
    cands = [str(c) for c in candidates]
    key = name_key(name)
    if not key:
        return []
    norm = {c: name_key(c) for c in cands}
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
    """같은 이름(exact_key: 앞뒤 공백·대소문자만 무시)을 찾습니다. 없으면 None."""
    key = exact_key(name)
    for c in candidates:
        if exact_key(c) == key:
            return str(c)
    return None


def not_found(kind: str, name: str, candidates: Iterable[str], extra: str = "", vocab: bool = False,
              display: Optional[Callable[[str], str]] = None, log_name: Optional[str] = None,
              log_extra: Optional[str] = None, option: Optional[str] = None) -> InputError:
    """'찾을 수 없음' InputError (R8).
    - 표기만 다른(name_key: 대소문자·공백·_-.) 이름이 하나뿐이면 [추천]: 뜻이 같아 LLM이 바로 1회 재시도해도 됩니다.
    - 부분·유사 일치는 뜻이 다를 수 있어(2024 -> 2024_1분기, 견적서.pdf -> 견적서_구버전.pdf) [후보]: 사용자가 고릅니다.
    - vocab=True(도구가 정한 옵션 값: 색 이름, 테마 등): 오타로 보이는 가까운 값 하나도 [추천] (사용자 데이터가 아니라 뜻이 바뀌지 않음).
    - 비슷한 것도 없으면 사용 가능한 이름 일부.
    display: 비교는 name·candidates(끝 이름)로 하고 message에 보일 때만 바꾸는 함수 (예: 메일 폴더 끝 이름 -> 전체 경로.
    전체 경로로 비교하면 같은 상위 경로 때문에 관계없는 폴더가 모두 비슷하게 나옴).
    log_name: 오류 로그 message에 name 대신 쓸 글 (예: 메일 계정 이름을 뺀 경로).
    log_extra: 오류 로그 message에 extra 대신 쓸 글 (extra에 계정 이름·문서 글이 들어갈 때 반드시 줌).
    option: 이름을 준 옵션 (고친 명령에서 그 옵션 값만 바꿈, retry_replacing)."""
    show = display or (lambda c: c)
    cands = [str(c) for c in candidates]
    key = name_key(name)
    same = [c for c in cands if key and name_key(c) == key]
    if vocab and not same:
        close = difflib.get_close_matches(key, [name_key(c) for c in cands], n=2, cutoff=0.75)
        if len(close) == 1:
            same = [c for c in cands if name_key(c) == close[0]][:1]
    head = f"{kind} '{show(name)}'을(를) 찾을 수 없습니다.{extra}"
    masked = log_name is not None or log_extra is not None
    log_head = head if not masked else (f"{kind} '{name if log_name is None else log_name}'을(를) 찾을 수 없습니다."
                                        f"{extra if log_extra is None else log_extra}")
    # 오류 로그에는 후보 이름(시트·문서·제목·메일 폴더·파일 이름 목록)을 남기지 않음 (log_message)
    if len(same) == 1:
        return InputError(f"{head} 표기만 다른 이름이 있습니다. [추천] {retry_replacing(name, same[0], option)}",
                          log_message=f"{log_head} 표기만 다른 이름이 있습니다. [추천 생략]")
    hints = same or suggest(name, cands)
    if hints:
        tail = " 비슷한 이름이 있지만 같은 것인지 도구가 정할 수 없습니다. 사용자에게 고르게 하세요."
        return ambiguous(head + tail, [show(c) for c in hints], log_message=log_head + tail, replace=name, values=hints,
                         option=option)
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


# 도구가 내는 명령(next_commands, confirm_commands, [추천] 전체 명령)을 LLM이 그대로 실행하는 셸.
# OpenCode 셸 도구와 같은 순서: pwsh(PowerShell 7)가 있으면 pwsh, 없으면 Windows PowerShell 5.1.
# OFFICE_CLI_SHELL=pwsh|powershell|cmd|bash로 강제 (1.4.76). PowerShell 두 종류는 정확, cmd·bash는 최선 노력.
SHELL_ENV = "OFFICE_CLI_SHELL"
SHELLS = ("pwsh", "powershell", "cmd", "bash")
_FOUND_SHELL: List[str] = []
# PowerShell이 작은따옴표로 읽는 글자 (안에서는 두 번 써야 글자 하나)
_PS_SINGLE_QUOTES = "'\u2018\u2019\u201a\u201b"
_PLAIN_PUNCT = "_-.:\\/,=+@!"
_PLAIN_INT = re.compile(r"-?(?:0|[1-9]\d{0,8})")
_PLAIN_OPTION = re.compile(r"--?[A-Za-z][A-Za-z0-9-]*")


def command_shell() -> str:
    """명령 문자열을 실행할 셸: OFFICE_CLI_SHELL > pwsh가 PATH에 있으면 pwsh > powershell (Windows PowerShell 5.1)."""
    forced = os.environ.get(SHELL_ENV, "").strip().lower()
    if forced in SHELLS:
        return forced
    if not _FOUND_SHELL:
        import shutil
        _FOUND_SHELL.append("pwsh" if shutil.which("pwsh") else "powershell")
    return _FOUND_SHELL[0]


def _net_space(ch: str) -> bool:
    """.NET char.IsWhiteSpace (Windows PowerShell 5.1이 인자를 큰따옴표로 감쌀지 정하는 기준)."""
    import unicodedata
    return ch in "\t\n\v\f\r\x85\xa0" or unicodedata.category(ch) in ("Zs", "Zl", "Zp")


def _plain_arg(s: str, shell: str) -> bool:
    """따옴표 없이 써도 셸이 글자 그대로 넘기는 값: 영숫자·한글(그 밖의 유니코드 글자)과 _-.:\\/,=+@! 만.
    PowerShell이 숫자로 바꿀 수 있는 모양(007, 1e3, 0x10, 1.50, 2-8), 옵션이 아닌 -로 시작(-a.b는 둘로 나뉨),
    맨 앞 =·:·+·@(배열·스플래팅)·!, 앞뒤·연속 쉼표(배열), bash의 백슬래시는 따옴표로 감쌉니다."""
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
    """Windows 명령줄 해석(CommandLineToArgvW, 파이썬 argv)이 s를 그대로 얻게: " 앞의 백슬래시는 두 배 + quote,
    큰따옴표로 감싸는 값은 끝 백슬래시도 두 배 (닫는 " 앞이라)."""
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
    """명령 문자열에 넣을 인자 하나를 실행 셸(command_shell)에 맞게 따옴표 처리합니다 (1.4.76).

    - 셸 특수문자가 없는 값(_plain_arg)은 그대로.
    - PowerShell: 작은따옴표 '...'(안의 '는 '') 안은 $·백틱·{}·@·#·;·( 모두 글자 그대로.
      Windows PowerShell 5.1은 값을 파이썬에 넘길 때 공백이 있으면 "..."로 감싸기만 하고 안의 "는 그대로 두므로,
      " -> \\"(앞 백슬래시 두 배), 감쌀 값의 끝 백슬래시 두 배, 빈 값은 '""'(5.1은 빈 인자를 버림)로 미리 바꿉니다.
      PowerShell 7.3+는 스스로 올바르게 바꾸므로 미리 바꾸지 않습니다.
    - cmd: "..."(안의 "는 "", %는 "^%"로 끊어 %NAME% 확장을 막음). bash: '...'(안의 '는 '\\'')."""
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


# 읽기·저장 뒤 열린 문서를 바꾸는 이어 하기 명령(confirm_commands, 1.4.75)의 directive 문구
INTO_OTHER_DOC = "사용자가 다른 문서에 넣기를 원하면 confirm_commands에서 맞는 명령을 1번 실행하세요(열린 문서를 바꿈)."


def retry_command(options: Dict[str, Any], drop: Sequence[str] = (), multi: Sequence[str] = ()) -> str:
    """[추천]에 넣을 전체 명령: 지금 실행한 명령에서 options의 옵션 값을 바꾸고(None이면 뺌) drop의 옵션(값 없는 플래그)을 뺌.
    multi: 값을 여러 개 받는 옵션(--formula A B)은 다음 옵션 앞까지 모두 뺌. 사용법 단순성 2번: 값만 주어 LLM이 끼워 넣게 하지 않음.
    실행 중인 명령을 모르면(테스트에서 함수를 직접 부름) 옵션과 값만 돌려줌."""
    argv = list(CURRENT_ARGV or [])
    out: List[str] = []
    i = 0
    while i < len(argv):
        a = argv[i]
        key = a.split("=", 1)[0]
        if key in options:
            i += 1
            if "=" not in a:
                if key in multi:
                    while i < len(argv) and not argv[i].startswith("--"):
                        i += 1
                elif i < len(argv):
                    i += 1
            continue
        if key in drop:
            i += 1
            continue
        out.append(a)
        i += 1
    add: List[str] = []
    for k, v in options.items():
        if v is None:
            continue
        add += [k] + ([] if v is True else [str(x) for x in v] if isinstance(v, (list, tuple)) else [str(v)])   # True: 값 없는 플래그
    if not argv or not CURRENT_TOOL:
        return " ".join(quote_arg(x) for x in add)
    return next_command(f"{CURRENT_TOOL}.py", out + add, "")["command"]


def retry_replacing(old: str, new: str, option: Any = None) -> str:
    """[추천]·[후보]에 넣을 전체 명령: 지금 실행한 명령에서 값 old를 new로 바꿈. 사용법 단순성 2번: 값만 주어 LLM이 끼워 넣게 하지 않음.
    option(옵션 이름 또는 그 묶음, 예: ('--by', '--desc'))이 명령에 있으면 그 옵션 값들 안에서 old인 값을 모두 바꿈
    (1.4.83: '--id 2 --index 2'의 --index만, '--by 금액,부서 --desc 금액'은 둘 다). 거기서 못 찾았거나 option이 없으면 명령 전체에서
    값 전체가 old인 인자만 바꾸고(폴더 이름 '요약'이 든 경로는 그대로), 그것도 없으면 시트 주소의 앞부분('old!A1', 목록 안의 'old!C7'),
    ;·, 목록의 한 값, 경로의 한 부분 중 처음 찾은 한 곳. 못 찾으면 option이 있으면 그 옵션 값으로 넣고(retry_command), 없으면 new만."""
    argv = list(CURRENT_ARGV or [])
    options = [option] if isinstance(option, str) else list(option or [])
    part_sep = re.compile(r"([\\/,])")          # 경로의 한 부분('보관\\회의록'), 쉼표 목록의 한 값('word,ppt')
    list_sep = re.compile(r"(\s*[;,]\s*)")       # 목록의 한 값 전체 ('a.pdf;C:\\폴더\\b.pdf'의 전체 경로, 'A;B'의 B)
    sheet_ref = re.compile(r"^(\s*)(?:'" + re.escape(old.replace("'", "''")) + r"'|" + re.escape(old) + r")\s*!")
    quoted_new = f"'{new.replace(chr(39), chr(39) * 2)}'" if re.search(r"[^\w가-힣]", new) else new

    def split_value(a: str) -> Tuple[str, str, str]:
        key, eq, val = a.partition("=")
        return (key, eq, val) if eq and key.startswith("--") else ("", "", a)

    def put(i: int, text: str) -> None:
        key, eq, _ = split_value(argv[i])
        argv[i] = f"{key}={text}" if eq else text

    def in_list(text: str) -> Optional[str]:
        parts = list_sep.split(text)
        if old not in [p.strip() for p in parts]:
            return None
        return "".join(new if p.strip() == old else p for p in parts)

    def partial(text: str) -> Optional[str]:
        """시트 주소 앞부분(목록 안 포함), 목록의 한 값, 경로의 한 부분."""
        items = list_sep.split(text)
        if any(sheet_ref.match(p) for p in items):
            return "".join(sheet_ref.sub(lambda m: m.group(1) + quoted_new + "!", p, count=1) for p in items)
        fixed = in_list(text)
        if fixed is None and old in part_sep.split(text):
            fixed = "".join(new if p == old else p for p in part_sep.split(text))
        return fixed

    found = False
    own = [i + 1 for i, a in enumerate(argv) if a in options and i + 1 < len(argv)] + \
          [i for i, a in enumerate(argv) if any(a.startswith(o + "=") for o in options)]
    for i in own:                                        # 1) 이름을 준 옵션의 값 안에서 모두 (값 전체, 목록, 시트 주소, 경로 조각)
        text = split_value(argv[i])[2]
        fixed = new if text == old else partial(text)
        if fixed is not None:
            put(i, fixed)
            found = True
    if not found:                                        # 2) 명령 전체에서 값 전체가 old인 인자
        for i, a in enumerate(argv):
            if split_value(a)[2] == old:
                put(i, new)
                found = True
    if not found:                                        # 3) 그 밖의 인자에서 처음 한 곳 (시트 주소, 목록의 한 값, 경로 조각)
        for i in range(len(argv)):
            fixed = partial(split_value(argv[i])[2])
            if fixed is None:
                continue
            put(i, fixed)
            found = True
            break
    if not found:
        return retry_command({options[0]: new}) if options else str(new)
    if not CURRENT_TOOL:
        return str(new)
    return next_command(f"{CURRENT_TOOL}.py", argv, "")["command"]


def next_command(script: str, args: Sequence[Any], purpose: str, destructive: bool = False,
                 needs_input: bool = False) -> Dict[str, Any]:
    """next_commands·confirm_commands 항목 하나를 만듭니다. script는 도구 폴더 기준 파일명 (예: 'office.py').

    destructive=True(덮어쓰기·삭제·메일 발송, confirm_commands에만)이면 "destructive": true를 붙입니다 (1.4.75): 사용자가
    '묻지 말고 끝까지'라고 했어도 LLM이 늘 묻게 하는 표시 (전역 지침 3장 6번)."""
    parts = [f"python {quote_arg(os.path.join(TOOL_DIR, script))}"] + [quote_arg(a) for a in args]
    item: Dict[str, Any] = {"purpose": purpose, "command": " ".join(parts)}
    if destructive:
        item["destructive"] = True
    if needs_input:
        item["needs_input"] = True          # <...> 자리를 사용자 값으로 바꿔야 함 (전역 지침 3장 6번)
    return item


# ==============================================================================
# 6. 미리보기(--dry-run) 공통 (excel.py, word.py, ppt.py의 쓰기 명령)
# ==============================================================================
# 규칙 하나 (사용법 단순성 4번): 문서를 바꾸는 명령은 모두 --dry-run을 받음.
#   미리보기가 있는 동작 -> dry_run_result (키: dry_run, count, planned, truncated, 바뀔 것이 있을 때만 confirm_commands)
#   읽기만 하는 동작(list 등) -> 그대로 실행하고 warnings에 무시했다고 적음 (dry_run_read_only)
#   미리보기가 없는 쓰기 동작 -> 바꾸지 않고 STOP_RETRY (dry_run_unsupported): 되돌릴 수 있으면 [추천] --dry-run을 뺀 같은 명령,
#   덮어쓰기·삭제가 들면(--overwrite, delete, clear) [후보] 1) 미리보기 없이 실행 + 같은 번호의 confirm_commands (사용자가 고른 뒤에만)

DRY_RUN_SHOW = 50            # 미리보기 planned 목록 상한 (넘으면 truncated: true)
DRY_RUN_DIRECTIVE = ("VERIFIED: 미리보기만 했고 문서는 바꾸지 않았습니다. preview를 사용자에게 알리고, 사용자가 실행을 원하면('예') "
                     "confirm_commands[0]을 그대로 실행하세요.")
DRY_RUN_NOTHING = ("VERIFIED: 미리보기만 했고 바뀔 것이 없습니다(0건, count 0). 실행할 명령이 없으니 실행하지 말고, 값을 바꿔 "
                   "다시 미리 보지도 말고 preview를 사용자에게 그대로 알리세요.")
DRY_RUN_DESTRUCTIVE = {"delete", "clear"}      # 미리보기 없이 실행하면 되돌리기 어려운 동작 (그리고 --overwrite)


def dry_run_command(tool: str, purpose: str, destructive: bool = True, overwrite: bool = False) -> Dict[str, Any]:
    """confirm_commands 항목: 지금 실행한 명령에서 --dry-run만 뺀 전체 명령 (사용자가 '예'라고 한 뒤에만 실행).
    overwrite: 실제 실행이 --overwrite 없이는 멈출 때 붙임 (값이 있는 칸 덮어쓰기 등)."""
    argv = [a for a in (CURRENT_ARGV or []) if a != "--dry-run"]
    if overwrite and "--overwrite" not in argv:
        argv.append("--overwrite")
    return next_command(f"{tool}.py", argv, purpose, destructive=destructive or overwrite)


def dry_run_result(tool: str, command: str, preview: str, count: int, planned: Sequence[str],
                   confirm: Any, directive: str = "", **extra: Any) -> Dict[str, Any]:
    """--dry-run 결과 (바꾸지 않음). planned는 앞 DRY_RUN_SHOW개, truncated는 늘 있음.
    count가 0이거나 confirm이 없으면 실행할 명령을 주지 않고 DRY_RUN_NOTHING.
    confirm: confirm_commands 항목 하나 또는 목록 (고를 것이 둘 이상이면 directive로 고르는 법을 적음)."""
    fields: Dict[str, Any] = {"dry_run": True, "count": count, "planned": list(planned[:DRY_RUN_SHOW]),
                              "truncated": len(planned) > DRY_RUN_SHOW}
    fields.update(extra)
    confirms = [confirm] if isinstance(confirm, dict) else list(confirm or [])
    if count and confirms:
        fields["confirm_commands"] = confirms
        return ok(tool, command, preview, directive or DRY_RUN_DIRECTIVE, **fields)
    return ok(tool, command, preview, DRY_RUN_NOTHING, **fields)


def dry_run_read_only(payload: Dict[str, Any], dry_run: bool, what: str) -> Dict[str, Any]:
    """읽기만 하는 동작(list 등)에 --dry-run: 받고 그대로 실행한 결과(payload)의 warnings에 무시했다고 적음."""
    if dry_run and payload.get("status") == "ok":
        payload["warnings"] = list(payload.get("warnings") or []) + [f"{what}은(는) 읽기만 하므로 --dry-run을 무시했습니다"]
    return payload


def dry_run_unsupported(tool: str, what: str) -> InputError:
    """미리보기가 없는 쓰기 동작에 --dry-run: 아무것도 바꾸지 않고 멈춤 (raise로 씀).
    되돌릴 수 있는 변경(office.py undo)이면 [추천] --dry-run을 뺀 같은 전체 명령,
    덮어쓰기·삭제가 들면 [후보] 1) 미리보기 없이 실행과 같은 번호의 confirm_commands (destructive, 사용자가 고른 뒤에만)."""
    argv = list(CURRENT_ARGV or [])
    destructive = "--overwrite" in argv or bool(DRY_RUN_DESTRUCTIVE & set(argv))
    cmd = dry_run_command(tool, f"{what} 미리보기 없이 실행", destructive)["command"]
    head = f"{what}에는 미리보기(--dry-run)가 없어 아무것도 바꾸지 않았습니다."
    if not destructive:
        return InputError(head + " 되돌릴 수 있는 변경(office.py undo)이라 미리보기 없이 실행합니다. [추천] " + cmd)
    return UserDecisionError(head + " 덮어쓰기·삭제가 들어 있어 미리보기 없이 실행할지 사용자 확인이 필요합니다.",
                             "미리보기가 없는 동작입니다. 사용자에게 미리보기 없이 실행할지 묻고, '예'라고 하면 confirm_commands[0]을 "
                             "1번 실행하세요.", candidates=["미리보기 없이 실행"], commands=[cmd], destructive=True)
