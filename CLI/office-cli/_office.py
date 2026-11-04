"""office-cli Office COM 공통 모듈 (excel.py, word.py, ppt.py, office.py가 사용).

- 대상 해석: "active" / 열린 문서 이름 / 파일 경로 -> 실제 문서
- 세션: 사용자가 연 문서는 그대로 사용(닫지 않음), 닫힌 파일은 전용 백그라운드 인스턴스로 열고 정리
- 읽기 세션: 시트/파일을 만들지 않고, 저장하지 않으며, 사용자 화면(활성 시트, 스크롤, 슬라이드)을 움직이지 않음
- 쓰기 세션: 작업 전 스냅샷(되돌리기용), 작업 후 실제로 고친 시트/슬라이드로만 화면 이동(--no-focus로 끔)
- 스냅샷: 원본 문서를 저장하거나 경로를 바꾸지 않음 (Excel/PPT SaveCopyAs, Word 파일 복사 또는 WordOpenXML)
"""
import glob
import hashlib
import json
import os
import re
import shutil
import sys
import time
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, Iterator, List, Optional, Tuple

import _runtime as rt

try:
    import pythoncom
    import win32com.client
    import win32gui
except ImportError:  # 비 Windows 환경(테스트)에서도 import는 되도록
    pythoncom = None
    win32com = None
    win32gui = None

KINDS = ("excel", "word", "ppt")
APP_LABEL = {"excel": "Excel", "word": "Word", "ppt": "PowerPoint"}
PROG_ID = {"excel": "Excel.Application", "word": "Word.Application", "ppt": "PowerPoint.Application"}
FRAME_CLASS = {"excel": ("XLMAIN",), "word": ("OpusApp",), "ppt": ("PPTFrameClass", "PP12FrameClass")}
DOC_CLASS = {"excel": "EXCEL7", "word": "_WwG", "ppt": "paneClassDC"}
EXTENSIONS = {
    "excel": {".xlsx": 51, ".xlsm": 52, ".xlsb": 50, ".xls": 56},
    "word": {".docx": 16, ".docm": 13, ".doc": 0, ".rtf": 6},
    "ppt": {".pptx": 24, ".pptm": 25, ".ppt": 1},
}
SNAPSHOT_MAX_BYTES = 50 * 1024 * 1024   # 이보다 큰 문서는 스냅샷 생략 (경고만)
SNAPSHOT_KEEP = 5                        # 문서별 보관 개수


def require_com() -> None:
    if win32com is None:
        raise rt.InputError("Office 작업에는 pywin32가 필요합니다 (python -m pip install pywin32). "
                            "Office 없이 파일만 읽으려면 reader.py read --file 을 쓰세요.")


class BusyMessageFilter:
    """COM 메시지 필터 (JIT 복구): Office가 재계산·자동 저장 등으로 잠깐 바빠 호출을 거부하면(RPC_E_SERVERCALL_RETRYLATER,
    RPC_E_CALL_REJECTED) 바로 실패하지 않고 BUSY_RETRY_MS 동안 짧게 기다렸다 다시 부릅니다. 넘으면 기존 office_busy 오류."""
    _public_methods_ = ["HandleInComingCall", "RetryRejectedCall", "MessagePending"]
    BUSY_RETRY_MS = 15000
    RETRY_DELAY_MS = 250
    NOTE = "Office가 바빠 호출을 거부함 -> 기다렸다 다시 호출"

    def HandleInComingCall(self, dwCallType, hTaskCaller, dwTickCount, lpInterfaceInfo):
        return 0                                     # SERVERCALL_ISHANDLED

    def RetryRejectedCall(self, hTaskCallee, dwTickCount, dwRejectType):
        if dwRejectType in (1, 2) and dwTickCount < self.BUSY_RETRY_MS:   # SERVERCALL_REJECTED / RETRYLATER
            if self.NOTE not in rt.RECOVERED:
                rt.note_recovered(self.NOTE)
            return self.RETRY_DELAY_MS
        if self.NOTE in rt.RECOVERED:                # 한도까지 기다려도 안 됨: 복구가 아니므로 recovered에서 빼고 오류 로그에 남김
            rt.RECOVERED.remove(self.NOTE)
        rt.record_recovery_attempt(f"Office 호출 거부 {self.BUSY_RETRY_MS // 1000}초 동안 재시도 실패")
        return -1                                    # 포기: 호출한 쪽이 오류를 받음 (office_busy)

    def MessagePending(self, hTaskCallee, dwTickCount, dwPendingType):
        return 2                                     # PENDINGMSG_WAITDEFPROCESS


_FILTER_REGISTERED = [False]


def init_com() -> None:
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        if not _FILTER_REGISTERED[0]:
            try:
                from win32com.server.util import wrap
                filt = BusyMessageFilter()
                filt._com_interfaces_ = [pythoncom.IID_IMessageFilter]
                pythoncom.CoRegisterMessageFilter(wrap(filt))
                _FILTER_REGISTERED[0] = True
            except Exception as e:                   # 필터가 없어도 동작은 같음 (바쁠 때 바로 office_busy)
                rt.debug(f"COM 메시지 필터 등록 실패: {e}")



# ==============================================================================
# 1. 실행 중인 Office 탐색
# ==============================================================================

def _enum_windows(match) -> List[int]:
    """EnumWindows 콜백은 항상 True를 반환합니다 (False 반환 시 pywin32가 예외를 던질 수 있음)."""
    found: List[int] = []

    def _cb(hwnd, _):
        try:
            if match(win32gui.GetClassName(hwnd)):
                found.append(hwnd)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return found


def _enum_children(parent: int, class_name: str) -> List[int]:
    found: List[int] = []

    def _cb(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) == class_name:
                found.append(hwnd)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumChildWindows(parent, _cb, None)
    except Exception:
        pass
    return found


def _dispatch_from_hwnd(hwnd: int) -> Optional[Any]:
    """문서 창 핸들에서 oleacc로 COM 객체를 얻습니다 (여러 Office 프로세스를 구분하기 위함)."""
    try:
        import ctypes
        ptr = ctypes.c_void_p()
        iid = ctypes.create_string_buffer(bytes(pythoncom.IID_IDispatch))
        if ctypes.windll.oleacc.AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, iid, ctypes.byref(ptr)) == 0 and ptr.value:
            return win32com.client.Dispatch(pythoncom.ObjectFromAddress(ptr.value, pythoncom.IID_IDispatch))
    except Exception:
        pass
    return None


def _window_pid(hwnd: int) -> Optional[int]:
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(hwnd)[1]
    except Exception:
        return None


def find_apps(kind: str) -> List[Any]:
    """실행 중인 Office 인스턴스를 프로세스 단위로 중복 없이 모읍니다. 포그라운드 창의 앱이 맨 앞."""
    require_com()
    init_com()
    apps: List[Any] = []
    if win32gui is not None:
        frames = _enum_windows(lambda c: c in FRAME_CLASS[kind])
        try:
            fg = win32gui.GetForegroundWindow()
            frames.sort(key=lambda h: 0 if h == fg else 1)
        except Exception:
            pass
        seen = set()
        for frame in frames:
            pid = _window_pid(frame)
            if pid is not None and pid in seen:
                continue
            for child in _enum_children(frame, DOC_CLASS[kind]):
                disp = _dispatch_from_hwnd(child)
                app = getattr(disp, "Application", None) if disp is not None else None
                if app is not None:
                    apps.append(win32com.client.Dispatch(app))
                    if pid is not None:
                        seen.add(pid)
                    break
    if not apps:
        try:
            apps.append(win32com.client.GetActiveObject(PROG_ID[kind]))
        except Exception:
            pass
    return apps


def _collection(app: Any, kind: str) -> Any:
    return getattr(app, {"excel": "Workbooks", "word": "Documents", "ppt": "Presentations"}[kind])


def _active_doc(app: Any, kind: str) -> Optional[Any]:
    attr = {"excel": "ActiveWorkbook", "word": "ActiveDocument", "ppt": "ActivePresentation"}[kind]
    try:
        doc = getattr(app, attr)
        if doc is not None:
            return doc
    except Exception:
        pass
    try:
        coll = _collection(app, kind)
        if coll.Count > 0:
            return coll(1)
    except Exception:
        pass
    return None


def is_saved(doc: Any) -> bool:
    saved = getattr(doc, "Saved", True)
    return saved != 0 if isinstance(saved, int) and not isinstance(saved, bool) else bool(saved)


def list_open(kind: str) -> List[Dict[str, Any]]:
    """열린 문서 목록 [{name, path, saved, active, app, doc}] (전체 경로 기준으로 중복 제거)."""
    docs: List[Dict[str, Any]] = []
    seen = set()
    for app in find_apps(kind):
        try:
            active = _active_doc(app, kind)
            active_full = str(active.FullName) if active is not None else ""
            coll = _collection(app, kind)
            for i in range(1, coll.Count + 1):
                doc = coll(i)
                full = str(getattr(doc, "FullName", "") or doc.Name)
                if full.lower() in seen:
                    continue
                seen.add(full.lower())
                docs.append({"name": str(doc.Name), "path": full if os.path.isabs(full) else "",
                             "saved": is_saved(doc),
                             "active": full == active_full, "app": app, "doc": doc})
        except Exception as e:
            rt.debug(f"{APP_LABEL[kind]} 문서 목록 수집 예외: {e}")
    return docs


def open_untitled(kind: str, name: str) -> List[Dict[str, Any]]:
    """열린 문서 중 저장한 적 없고 이름이 name인 문서 [{name, app, doc}].
    list_open과 달리 이름으로 중복을 없애지 않습니다 (다른 Excel 창의 Book1 둘을 하나로 보지 않도록)."""
    hits: List[Dict[str, Any]] = []
    key = str(name).strip().lower()
    for app in find_apps(kind):
        try:
            coll = _collection(app, kind)
            for i in range(1, coll.Count + 1):
                doc = coll(i)
                full = str(getattr(doc, "FullName", "") or "")
                if not os.path.isabs(full) and str(doc.Name).strip().lower() == key:
                    hits.append({"name": str(doc.Name), "app": app, "doc": doc})
        except Exception as e:
            rt.debug(f"{APP_LABEL[kind]} 저장 안 된 문서 찾기 예외: {e}")
    return hits


def get_active(kind: str) -> Dict[str, Any]:
    """사용자가 보고 있는 문서. 포그라운드 앱의 활성 문서를 우선합니다."""
    for app in find_apps(kind):
        doc = _active_doc(app, kind)
        if doc is not None:
            full = str(getattr(doc, "FullName", "") or doc.Name)
            return {"name": str(doc.Name), "path": full if os.path.isabs(full) else "", "app": app, "doc": doc}
    raise no_open_doc(kind)


ELEVATED_DIRECTIVE = ("실행하지 않았습니다. 다시 실행하지 말고 message의 관리자 권한 안내를 사용자에게 그대로 전하세요. "
                      "관리자 권한 창에서는 문서를 열어도, 경로를 줘도 사용자가 연 문서에 닿지 않습니다.")


def no_open_doc(kind: str, option: str = "--target", note: str = "") -> "rt.UserDecisionError":
    """열린 문서가 없음. option: 이 명령에서 대상 문서를 주는 옵션 (office.py transfer는 --from-target/--to-target).
    관리자 권한으로 실행 중이면 일반 권한 창의 문서를 볼 수 없으므로 그 안내를 붙입니다 (rt.elevation_hint)."""
    hint = rt.elevation_hint()
    if hint:
        err = rt.UserDecisionError(f"열려 있는 {APP_LABEL[kind]} 문서를 찾지 못했습니다.{note} {hint}", ELEVATED_DIRECTIVE)
        err.error_type = "elevated_process"
        return err
    return rt.UserDecisionError(
        f"열려 있는 {APP_LABEL[kind]} 문서가 없습니다.{note}",
        "실행하지 않았습니다. 사용자에게 문서를 열어 달라고 하거나 파일의 전체 경로를 물어, 같은 명령에 "
        f"{option} \"<전체경로>\"를 붙여 1번만 실행하세요. 경로를 추측하지 마세요.")


def no_open_docs_result(tool: str, command: str, preview: str, directive: str, **fields: Any) -> Dict[str, Any]:
    """열린 문서 목록이 0건인 성공 응답 (info --all, office.py windows). 관리자 권한으로 실행 중이면
    일반 권한 창의 문서를 볼 수 없어 0건을 확인한 것이 아니므로 UNVERIFIED:와 그 안내."""
    hint = rt.elevation_hint()
    if hint:
        return rt.ok_unverified(tool, command, preview, [hint], **fields)
    return rt.ok(tool, command, preview, directive, **fields)


def match_open(kind: str, target: str, docs: Optional[List[Dict[str, Any]]] = None,
               option: str = "--target") -> Optional[Dict[str, Any]]:
    """열린 문서 중 이름/파일명/전체 경로가 target과 같은 것 (대소문자 무시).
    전체 경로가 같으면 그 문서. 이름만 같은 문서가 둘 이상이면(다른 폴더의 같은 파일명, 다른 Excel 창의 Book1 등)
    하나를 고르지 않고 [후보](전체 경로)와 함께 오류를 냅니다.
    option: 이 명령에서 대상 문서를 주는 옵션 (office.py transfer는 --from-target/--to-target)."""
    docs = list_open(kind) if docs is None else docs
    t = target.strip().strip('"')
    t_low = t.lower()
    t_abs = os.path.normcase(os.path.abspath(t)) if (os.sep in t or "/" in t or ":" in t) else None
    if t_abs:
        for d in docs:
            if d["path"] and os.path.normcase(d["path"]) == t_abs:
                return d
    hits = [d for d in docs
            if d["name"].lower() == t_low or (d["path"] and os.path.basename(d["path"]).lower() == t_low)]
    if len(hits) > 1:
        raise rt.ambiguous(
            f"열린 {APP_LABEL[kind]} 문서 중 이름이 '{t}'인 문서가 {len(hits)}개입니다. "
            f"어느 문서인지 사용자에게 확인한 뒤 {option}에 전체 경로를 지정하세요.",
            [d["path"] or f"{d['name']} (저장 안 됨)" for d in hits])
    return hits[0] if hits else None


def _sheet_key(name: str) -> str:
    """대소문자, 공백, _ - . 를 무시한 시트 이름 비교값."""
    return re.sub(r"[\s_\-.]+", "", str(name)).lower()


def missing_sheet(name: str, names: List[str]) -> "rt.InputError":
    """쓰기 대상 시트가 없을 때의 오류. 새 시트는 만들지 않습니다.
    - 표기만 다른 시트(대소문자·공백·_-. 차이)가 하나면 [추천]으로 그 이름 (바로 재시도 가능)
    - 그 밖에는 needs_user_decision: 새 시트를 만들지(또는 비슷한 기존 시트에 쓸지) 사용자가 정함.
      message에는 사실만 적고 --new-sheet나 '만드세요' 같은 명령문을 넣지 않습니다
      (경량 LLM이 message의 명령문이나 [후보]의 '새 시트' 항목을 그대로 골라 묻지 않고 만든 사례: 시나리오 10).
      --new-sheet는 agent_directive에만, 사용자가 만들라고 답한 경우라는 조건과 함께 적습니다."""
    key = _sheet_key(name)
    same = [n for n in names if _sheet_key(n) == key]
    if len(same) == 1:
        return rt.InputError(f"시트 '{name}'이(가) 없습니다. 표기(대소문자·공백·_-.)만 다른 기존 시트가 있습니다. [추천] {same[0]}",
                             no_fallback=True, log_message=f"시트 '{name}'이(가) 없습니다. 표기만 다른 기존 시트가 있습니다. [추천 생략]")
    shown = ", ".join(names[:15]) + (" ..." if len(names) > 15 else "")
    similar = same or rt.suggest(name, names)
    # 오류 로그에는 시트 이름 목록을 남기지 않음 (문서에서 가져온 이름)
    log_msg = f"시트 '{name}'이(가) 없습니다 (있는 시트 {len(names)}개 이름 생략). 새 시트를 만들지 않았고 아무 데도 쓰지 않았습니다."
    if similar:
        return rt.UserDecisionError(
            f"시트 '{name}'이(가) 없습니다 (있는 시트: {shown}). 새 시트를 만들지 않았고 아무 데도 쓰지 않았습니다. "
            "이름이 비슷한 기존 시트에 쓸지, 새 시트를 만들지 사용자에게 물어보세요.",
            "다시 실행하지 마세요. message를 사용자에게 그대로 보여 주고, [후보]의 기존 시트에 쓸지 "
            f"새 시트 '{name}'을(를) 만들지 물어보세요. 사용자가 기존 시트를 고르면 --sheet를 그 이름으로 바꿔 실행하고, "
            "사용자가 새 시트를 만들라고 답한 경우에만 같은 명령에 --new-sheet를 붙여 실행하세요.",
            candidates=similar, log_message=log_msg)
    return rt.UserDecisionError(
        f"시트 '{name}'이(가) 없습니다 (있는 시트: {shown}). 새 시트를 만들지 않았고 아무 데도 쓰지 않았습니다. "
        "새 시트를 만들지 사용자에게 물어보세요.",
        f"다시 실행하지 마세요. message를 사용자에게 그대로 보여 주고, 새 시트 '{name}'을(를) 만들지 물어보세요. "
        "사용자가 만들라고 답한 경우에만 같은 명령에 --new-sheet를 붙여 실행하세요.", log_message=log_msg)


def shape_box(obj: Any) -> Dict[str, float]:
    """COM 도형/차트/그림의 위치와 크기(pt)."""
    return {k.lower(): round(float(getattr(obj, k)), 1) for k in ("Left", "Top", "Width", "Height")}


def outside_area(box: Dict[str, float], width: float, height: float, tol: float = 0.5) -> bool:
    """box가 (0, 0, width, height) 영역 밖으로 나가는지."""
    return (box["left"] < -tol or box["top"] < -tol or box["left"] + box["width"] > width + tol
            or box["top"] + box["height"] > height + tol)


def overlap_ratio(a: Dict[str, float], b: Dict[str, float]) -> float:
    """두 상자가 겹친 넓이 / 작은 상자 넓이 (0~1)."""
    w = min(a["left"] + a["width"], b["left"] + b["width"]) - max(a["left"], b["left"])
    h = min(a["top"] + a["height"], b["top"] + b["height"]) - max(a["top"], b["top"])
    if w <= 0 or h <= 0:
        return 0.0
    smaller = min(a["width"] * a["height"], b["width"] * b["height"])
    return round(w * h / smaller, 2) if smaller > 0 else 0.0


# ==============================================================================
# 2. 스냅샷 (되돌리기)
# ==============================================================================

def _snapshot_dir() -> str:
    path = os.path.join(rt.get_cache_dir(), "snapshots")
    os.makedirs(path, exist_ok=True)
    return path


_CREATED_RE = re.compile(r"<dcterms:created[^>]*>([^<]+)</dcterms:created>")


def doc_identity(path: str) -> Optional[str]:
    """디스크 파일의 문서 식별값: OOXML docProps/core.xml의 만든 시각.
    Office가 저장해도 유지되고, 같은 경로에 새로 만든 파일은 달라집니다. 알 수 없으면 None(비교 생략)."""
    if not path or os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm", ".docx", ".docm", ".pptx", ".pptm"):
        return None
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            m = _CREATED_RE.search(z.read("docProps/core.xml").decode("utf-8", errors="replace"))
        return m.group(1).strip() if m else None
    except Exception:
        return None


def save_snapshot(kind: str, doc: Any, tool: str = "", command: str = "", rotate: bool = True) -> Optional[str]:
    """쓰기 직전 문서 상태를 저장합니다. 원본 문서는 저장하거나 경로를 바꾸지 않습니다. 실패해도 작업은 계속합니다.
    rotate=False: 오래된 지점 정리(SNAPSHOT_KEEP개 유지)를 하지 않음. open_target은 명령이 실제로 바꾼 뒤에만
    정리합니다 (쓰기 전에 멈춘 실행이 진짜 옛 지점을 밀어내지 않도록)."""
    try:
        name = str(doc.Name)
        full = str(getattr(doc, "FullName", "") or "")
        path = full if os.path.isabs(full) else ""
        if path and os.path.exists(path) and os.path.getsize(path) > SNAPSHOT_MAX_BYTES:
            rt.log("WARN", f"문서가 커서({os.path.getsize(path) // 1048576}MB) 되돌리기용 스냅샷을 만들지 않았습니다.")
            return None
        key = hashlib.sha1((path or name).lower().encode("utf-8")).hexdigest()[:10]
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")   # 연속 쓰기에도 이름이 겹치지 않도록 마이크로초까지
        base = os.path.join(_snapshot_dir(), f"{kind}_{key}_{stamp}")
        ext = (os.path.splitext(path)[1].lower() if path else "") or {"excel": ".xlsx", "word": ".docx", "ppt": ".pptx"}[kind]
        fmt = "native"
        data_file = base + ext
        if kind in ("excel", "ppt"):
            doc.SaveCopyAs(os.path.normpath(data_file))
        elif path and os.path.exists(path) and bool(getattr(doc, "Saved", False)):
            shutil.copy2(path, data_file)
        else:
            data_file = base + ".xml"
            fmt = "flat_xml"
            with open(data_file, "w", encoding="utf-8") as f:
                f.write(doc.WordOpenXML)
        if not os.path.exists(data_file):
            return None
        meta = {"app": kind, "name": name, "original_path": path, "snapshot_file": data_file,
                "snapshot_format": fmt, "created_at": stamp, "created_ts": time.time(), "by": f"{tool} {command}".strip(),
                "doc_created": doc_identity(path),
                "untitled": not path}   # 저장한 적 없는 문서: undo가 열린 문서 중 같은 이름의 '저장 안 된' 문서를 찾아 되돌림
        rt.safe_write_text(base + ".json", json.dumps(meta, ensure_ascii=False))
        if rotate:
            _rotate_snapshots(kind, key)
        rt.debug(f"스냅샷 저장: {os.path.basename(data_file)}")
        return data_file
    except Exception as e:
        rt.log("WARN", f"되돌리기용 스냅샷을 만들지 못했습니다 (작업은 계속): {e}")
        return None


def save_barrier(kind: str, doc: Any, reason: str, tool: str = "", command: str = "", rotate: bool = True) -> Optional[str]:
    """되돌릴 지점을 만들지 못한 쓰기를 기록합니다. undo가 이 작업을 건너뛰고 더 오래된 지점으로 되돌려
    그 사이 작업까지 사라지게 하지 않도록 합니다 (기록 실패는 무시). 만든 기록 파일(.json) 경로를 돌려줍니다."""
    try:
        name = str(doc.Name)
        full = str(getattr(doc, "FullName", "") or "")
        path = full if os.path.isabs(full) else ""
        key = hashlib.sha1((path or name).lower().encode("utf-8")).hexdigest()[:10]
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        base = os.path.join(_snapshot_dir(), f"{kind}_{key}_{stamp}")
        meta = {"app": kind, "name": name, "original_path": path, "barrier": True, "reason": reason,
                "created_at": stamp, "created_ts": time.time(), "by": f"{tool} {command}".strip()}
        rt.safe_write_text(base + ".json", json.dumps(meta, ensure_ascii=False))
        if rotate:
            _rotate_snapshots(kind, key)
        return base + ".json"
    except Exception:
        return None


def delete_snapshot(meta_file: str) -> None:
    base = os.path.splitext(meta_file)[0]
    for f in glob.glob(glob.escape(base) + ".*"):
        try:
            os.remove(f)
        except OSError:
            pass


def _rotate_snapshots(kind: str, key: str) -> None:
    metas = sorted(glob.glob(os.path.join(_snapshot_dir(), f"{kind}_{key}_*.json")), reverse=True)  # 이름의 시각 순
    for old in metas[SNAPSHOT_KEEP:]:
        delete_snapshot(old)


def rotate_point(meta_file: str) -> None:
    """이 지점(.json)과 같은 문서의 지점을 최신 SNAPSHOT_KEEP개만 남기고 정리합니다 (실패는 무시)."""
    try:
        kind, key, _ = os.path.basename(meta_file).split("_", 2)
        _rotate_snapshots(kind, key)
    except Exception:
        pass


def list_snapshots(kind: Optional[str] = None) -> List[Dict[str, Any]]:
    """최신순 스냅샷 목록. 각 항목에 meta_file 포함.
    같은 경로에 다른 문서가 새로 만들어졌으면(문서 식별값이 다름) 그 경로의 옛 스냅샷은 지우고 빼 줍니다."""
    pattern = f"{kind}_*.json" if kind else "*.json"
    result = []
    current: Dict[str, Optional[str]] = {}
    stale = 0
    files = glob.glob(os.path.join(_snapshot_dir(), pattern))
    for mf in sorted(files, key=lambda f: os.path.basename(f).split("_", 2)[-1], reverse=True):  # 문서 무관 시각 순
        try:
            meta = json.loads(rt.read_text_file(mf))
            if not meta.get("barrier") and not os.path.exists(meta.get("snapshot_file", "")):
                continue
            orig, recorded = meta.get("original_path", ""), meta.get("doc_created")
            if orig and recorded:
                if orig not in current:
                    current[orig] = doc_identity(orig)
                if current[orig] is not None and current[orig] != recorded:
                    delete_snapshot(mf)
                    stale += 1
                    continue
            meta["meta_file"] = mf
            result.append(meta)
        except Exception:
            continue
    if stale:
        rt.log("INFO", f"같은 경로에 새로 만든 문서가 있어 이전 문서의 되돌리기 지점 {stale}개를 정리했습니다.")
    return result


# ==============================================================================
# 3. 세션
# ==============================================================================

class OfficeTarget:
    """세션이 넘겨주는 대상 문서. 도구는 doc/app을 쓰고, 고친 위치를 set_focus로 알려 줍니다."""

    def __init__(self, kind: str, app: Any, doc: Any, source: str, readonly: bool):
        self.kind = kind
        self.app = app
        self.doc = doc
        self.source = source          # active | open | file | new
        self.readonly = readonly
        self.focus_sheet: Optional[Any] = None
        self.focus_slide: Optional[int] = None
        self.changed = False

    @property
    def name(self) -> str:
        return str(self.doc.Name)

    @property
    def path(self) -> str:
        full = str(getattr(self.doc, "FullName", "") or "")
        return full if os.path.isabs(full) else ""

    def resolved(self) -> Dict[str, Any]:
        """결과 JSON의 target_resolved 값 (LLM이 실제 작업 대상을 알 수 있도록)."""
        info: Dict[str, Any] = {"app": APP_LABEL[self.kind], "name": self.name, "path": self.path, "source": self.source}
        if self.focus_sheet is not None:
            try:
                info["sheet"] = str(self.focus_sheet.Name)
            except Exception:
                pass
        if self.focus_slide is not None:
            info["slide"] = self.focus_slide
        return info

    # --- Excel ---
    def sheet(self, name: Optional[str] = None, create: bool = False, allow_create: bool = False) -> Any:
        """시트를 찾습니다 (대소문자 무시). 없으면 쓰기 세션 + create=True일 때만 만들고, 아니면 [추천]/[후보]와 함께 오류.
        create는 사용자가 새 시트를 명시한 경우(--new-sheet, sheet add)에만 켭니다. 오타 난 이름으로 새 시트를 만들지 않도록.
        allow_create: 이 명령에 --new-sheet가 있어 '새 시트를 만들지' 물을 수 있는 경우(excel write)만 True.
        그 밖의 명령(replace, style, range, chart 등)은 없는 옵션을 안내하지 않고 '찾을 수 없음'([후보]/[추천])."""
        sheets = self.doc.Worksheets
        if not name:
            try:
                ws = self.doc.ActiveSheet
                _ = ws.UsedRange          # 차트 시트에는 UsedRange가 없음 -> 첫 워크시트 사용
                return ws
            except Exception:
                return sheets(1)
        names = [str(sheets(i).Name) for i in range(1, sheets.Count + 1)]
        exact = rt.find_name(name, names)
        if exact is not None:
            return sheets(exact)
        if create and not self.readonly:
            check_sheet_name(name, names)       # 잘못된 이름은 시트를 만들기 전에 (만든 뒤 Name에서 실패하면 빈 시트가 남음)
            ws = sheets.Add(After=sheets(sheets.Count))
            ws.Name = name
            self.changed = True
            return ws
        if self.readonly or not allow_create:
            raise rt.not_found("시트", name, names)
        raise missing_sheet(name, names)

    def set_focus(self, sheet: Any = None, slide: Optional[int] = None, changed: bool = True) -> None:
        """쓰기 후 사용자 화면을 이동할 위치를 기록합니다 (세션 종료 시 적용).
        changed: 이 명령이 문서를 바꿨는지. 바꾼 것 없이 성공한 명령(차트 0개 삭제, 이미 있는 시트 add 등)은 False로 주어
        이번 되돌리기 지점을 지우게 합니다. 문서를 바꾸는 명령은 set_focus를 부르거나 t.changed = True로 표시합니다."""
        if sheet is not None:
            self.focus_sheet = sheet
        if slide is not None:
            self.focus_slide = int(slide)
        if changed:
            self.changed = True

    def apply_focus(self) -> None:
        if self.readonly or self.source not in ("active", "open"):
            return
        try:
            if self.kind == "excel" and self.focus_sheet is not None:
                self.focus_sheet.Activate()
            elif self.kind == "ppt" and self.focus_slide is not None:
                view = self.app.ActiveWindow.View
                if self.app.ActiveWindow.Presentation.FullName == self.doc.FullName:
                    view.GotoSlide(self.focus_slide)
            elif self.kind == "word":
                self.doc.Activate()
        except Exception as e:
            rt.debug(f"화면 이동 생략: {e}")


SHEET_NAME_BAD = re.compile(r"[\\/?*\[\]:]")    # 엑셀 시트 이름에 쓸 수 없는 글자
SHEET_NAME_RESERVED = ("history",)              # 엑셀이 예약한 시트 이름 (변경 내용 추적용, 대소문자 무관)


def check_sheet_name(name: Any, existing: Optional[List[str]] = None) -> None:
    """엑셀 시트 이름 규칙(31자 이하, / \\ ? * [ ] : 없음, 빈 이름 아님, 앞뒤 ' 없음, 예약된 이름 History 아님)을
    문서를 바꾸기 전에 확인합니다. 어기면 InputError와 [추천](쓸 수 없는 글자를 _로 바꾸고 31자로 자른 이름,
    History는 History_, existing과 겹치지 않게). 시트를 만든 뒤 Name에서 실패하면 빈 시트가 남기 때문.
    existing을 주지 않으면 [추천]이 있는 시트와 겹칠 수 있으므로 시트 이름을 아는 곳에서 부릅니다."""
    text = "" if name is None else str(name)
    if not text.strip():
        raise rt.InputError("시트 이름이 비어 있습니다. 아무것도 바꾸지 않았습니다. 시트 이름을 지정하세요.")
    problems = []
    if SHEET_NAME_BAD.search(text):
        problems.append("/ \\ ? * [ ] : 는 쓸 수 없음")
    if len(text) > 31:
        problems.append(f"31자까지 ({len(text)}자)")
    if text.startswith("'") or text.endswith("'"):
        problems.append("앞뒤에 ' 는 쓸 수 없음")
    reserved = text.lower() in SHEET_NAME_RESERVED
    if reserved:
        problems.append("엑셀이 예약한 이름 (대소문자 무관)")
    if not problems:
        return
    rec = SHEET_NAME_BAD.sub("_", text)[:31]
    if reserved:
        rec += "_"
    if rec.startswith("'"):
        rec = "_" + rec[1:]
    if rec.endswith("'"):
        rec = rec[:-1] + "_"
    base, k = rec, 2
    while existing and rt.find_name(rec, existing) is not None:
        tail = f" {k}"
        rec = base[:31 - len(tail)] + tail
        k += 1
    raise rt.InputError(f"시트 이름 '{text}'은(는) 엑셀에서 쓸 수 없습니다 ({', '.join(problems)}). "
                        f"아무것도 바꾸지 않았습니다. [추천] {rec}")


def _new_instance(kind: str) -> Any:
    """전용 백그라운드 인스턴스. Excel/Word는 DispatchEx로 사용자 창과 분리됩니다."""
    if kind == "ppt":   # PowerPoint는 단일 인스턴스 앱: 사용자 창 상태는 건드리지 않음
        return win32com.client.Dispatch(PROG_ID[kind])
    app = win32com.client.DispatchEx(PROG_ID[kind])
    app.Visible = False
    if kind == "excel":
        app.DisplayAlerts = False
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass
    else:
        app.DisplayAlerts = 0
    try:
        # 보이지 않는 인스턴스에서 열기 매크로(Workbook_Open의 MsgBox 등)가 창을 띄우면 끝나지 않으므로 매크로를 끔
        app.AutomationSecurity = 3                   # msoAutomationSecurityForceDisable
    except Exception:
        pass
    return app


# --- 닫힌 파일 열기 감시: 문서 보안(DRM) 인증 창·대화상자로 보이지 않는 인스턴스가 멈추면 그 인스턴스만 끝냄 ---

OPEN_TIMEOUT_ENV = "OFFICE_CLI_OPEN_TIMEOUT"   # 초. 0이면 감시하지 않음
DEFAULT_OPEN_TIMEOUT = 30.0
IMAGE_NAME = {"excel": "EXCEL.EXE", "word": "WINWORD.EXE", "ppt": "POWERPNT.EXE"}
KIND_SCRIPT = {"excel": "excel.py", "word": "word.py", "ppt": "ppt.py"}


def open_timeout() -> float:
    raw = os.environ.get(OPEN_TIMEOUT_ENV, "").strip()
    try:
        v = float(raw) if raw else DEFAULT_OPEN_TIMEOUT
    except ValueError:
        v = DEFAULT_OPEN_TIMEOUT
    return v if v > 0 else 0.0


def _office_pids(kind: str) -> set:
    """실행 중인 그 Office 프로그램의 프로세스 번호들 (읽지 못하면 빈 집합)."""
    try:
        import subprocess
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {IMAGE_NAME[kind]}", "/FO", "CSV", "/NH"],
                             capture_output=True, timeout=10, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        pids = set()
        for line in out.stdout.decode("mbcs" if os.name == "nt" else "utf-8", errors="replace").splitlines():
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) > 1 and parts[1].isdigit():
                pids.add(int(parts[1]))
        return pids
    except Exception:
        return set()


def _instance_pid(kind: str, app: Any, before: set) -> Optional[int]:
    """도구가 띄운 인스턴스의 프로세스 번호. Excel은 창 핸들, 그 밖은 띄우기 전후 프로세스 목록의 차이(하나일 때만)."""
    if kind == "excel":
        try:
            pid = _window_pid(int(app.Hwnd))
            if pid:
                return pid
        except Exception:
            pass
    new = _office_pids(kind) - before
    return next(iter(new)) if len(new) == 1 else None


def _kill_pid(pid: int) -> bool:
    """그 프로세스 하나만 끝냅니다 (사용자의 다른 Office 창은 그대로)."""
    try:
        import signal
        os.kill(pid, signal.SIGTERM)            # Windows: TerminateProcess
        return True
    except Exception:
        return False


def _user_windows(kind: str, pid: int) -> Optional[bool]:
    """도구가 띄운 그 프로세스에 사용자가 볼 수 있는 문서 창(프레임 창)이 있는지 COM 없이 창 목록으로 확인합니다.
    도구가 연 문서는 창 없이(Visible=False, WithWindow=False) 열므로, 보이는 프레임 창이 있으면 사용자가 그 프로세스에서 연 문서
    (PowerPoint는 한 프로세스라 사용자가 감시 중에 연 발표 파일도 같은 프로세스). True 있음, False 없음, None 확인하지 못함."""
    if win32gui is None:
        return None
    try:
        for hwnd in _enum_windows(lambda c: c in FRAME_CLASS[kind]):
            owner = _window_pid(hwnd)
            if owner is None:
                return None
            if owner == pid and win32gui.IsWindowVisible(hwnd):
                return True
        return False
    except Exception:
        return None


def _drm_or_dialog_error(kind: str, path: str, tool: str, waited: float, killed: bool = True,
                         pid: Optional[int] = None) -> rt.InputError:
    name = os.path.basename(path)

    def same(a: str) -> bool:
        try:
            return bool(a.strip()) and os.path.normcase(os.path.abspath(os.path.expandvars(a.strip().strip('"')))) \
                == os.path.normcase(os.path.abspath(path))
        except Exception:
            return False
    new_argv = [name if same(a) else a for a in (rt.CURRENT_ARGV or [])]   # 경로 대신 문서 이름 (열린 문서로)
    nexts = [rt.next_command(f"{tool}.py" if tool else KIND_SCRIPT[kind], new_argv,
                             "사용자가 문서를 Office에서 직접 열어 둔 뒤 같은 작업을 열린 문서로 실행")] if new_argv else []
    what = (f"백그라운드 {APP_LABEL[kind]}만 끝냈습니다 (사용자의 다른 창은 그대로)" if killed else
            f"{APP_LABEL[kind]}에 사용자가 연 문서가 있을 수 있어 {APP_LABEL[kind]}을(를) 끝내지 않았습니다 "
            f"(그대로 둔 프로세스: {IMAGE_NAME[kind]}" + (f" PID {pid}" if pid else ", 이미 떠 있던 것") +
            ". 열기 창이 떠 있으면 사용자가 닫으면 됨)")
    return rt.InputError(
        f"'{name}'이(가) {waited:g}초 안에 열리지 않아 {what}. "
        "문서 보안 인증 창 등으로 열리지 않음. 사용자가 문서를 직접 열어 두면 열린 문서로 읽음. "
        f"열기가 오래 걸린 경우일 수도 있습니다({OPEN_TIMEOUT_ENV}으로 늘릴 수 있음).",
        directive=("문서가 열리지 않아 멈췄습니다. 다시 실행하지 말고 message를 사용자에게 알리세요. 사용자가 문서를 Office에서 직접 "
                   "열어 두었다고 하면 next_commands[0]을 1번 실행하세요."),
        error_type="drm_or_dialog", fields={"next_commands": nexts}, no_fallback=True)


def _abandon(err: rt.InputError, kind: str, tool: str) -> None:
    """프로그램을 끝내지 못해 멈춘 COM 호출이 돌아오지 않을 때: 오류 JSON 1줄을 내고 이 도구 프로세스만 끝냅니다 (감시 스레드에서)."""
    name = (tool or KIND_SCRIPT[kind]).replace(".py", "")
    argv = list(rt.CURRENT_ARGV or [])
    payload = rt.error_payload(err, name, argv[0] if argv else "")
    try:
        rt.log("ERROR", str(err))
        rt.write_error_log(name, None, payload)
        rt.emit_json(payload)
        sys.stdout.flush()
        sys.stderr.flush()
    finally:
        os._exit(1)


class _OpenWatch:
    """닫힌 파일을 여는 동안 시간을 잽니다. 시간이 지나면 도구가 띄운 인스턴스 프로세스만 끝냅니다(그러면 멈춘 COM 호출이 오류로 돌아옴).
    Word·PowerPoint는 그 프로세스에 사용자가 연 문서 창이 없다고 확인될 때만 끝냄(PowerPoint는 한 프로세스라 감시 중에 사용자가 연
    발표 파일도 같은 프로세스). 끝낼 프로세스를 모르거나(사용자가 이미 띄운 PowerPoint 등) 확인하지 못하면 끝내지 않고
    drm_or_dialog 오류만 내고 도구를 끝냄(_abandon)."""

    def __init__(self, kind: str, pid: Optional[int], timeout: float, path: str = "", tool: str = ""):
        import threading
        self.kind, self.pid, self.timeout, self.path, self.tool = kind, pid, timeout, path, tool
        self.fired = False
        self._timer = threading.Timer(timeout, self._fire) if timeout > 0 and (pid or path) else None
        if self._timer is not None:
            self._timer.daemon = True

    def _fire(self) -> None:
        safe = bool(self.pid) and (self.kind == "excel" or _user_windows(self.kind, self.pid) is False)
        if safe:
            self.fired = _kill_pid(self.pid)
            if self.fired:
                return
        if self.path:
            _abandon(_drm_or_dialog_error(self.kind, self.path, self.tool, self.timeout, killed=False, pid=self.pid),
                     self.kind, self.tool)

    def __enter__(self) -> "_OpenWatch":
        if self._timer is not None:
            self._timer.start()
        return self

    def __exit__(self, *exc: Any) -> None:
        if self._timer is not None:
            self._timer.cancel()


def _ppt_running() -> bool:
    try:
        win32com.client.GetActiveObject(PROG_ID["ppt"])
        return True
    except Exception:
        return False


# 열기 암호가 걸린 파일: 보이지 않는 인스턴스에서 암호 창이 뜨면 끝나지 않으므로 틀린 암호를 넘겨 바로 오류가 나게 함
_NO_PASSWORD = "office-cli-no-password"


def _word_encrypted(path: str) -> Optional[bool]:
    """Word 파일에 열기 암호가 있는지 파일 구조로 판별합니다. 판별하지 못하면 None (DRM 파일 등)."""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
        if head[:2] == b"PK":                       # 암호 없는 docx/docm (암호를 걸면 OLE 파일이 됨)
            return False
        if head != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            return None
        import olefile
        with olefile.OleFileIO(path) as ole:
            if ole.exists("EncryptedPackage"):      # 암호 docx/docm
                return True
            with ole.openstream("WordDocument") as s:
                return bool(s.read(12)[11] & 0x01)  # doc: FIB fEncrypted (0x000A의 비트 0x0100)
    except Exception:
        return None


def _password_error(path: str) -> rt.InputError:
    return rt.InputError(f"열기 암호가 걸린 문서라 열 수 없습니다: {os.path.basename(path)}. 사용자가 Office에서 직접 열어 둔 뒤 "
                         "같은 요청을 하면 열린 문서로 작업합니다.", no_fallback=True)


def _open_file(kind: str, app: Any, path: str, readonly: bool) -> Any:
    encrypted: Optional[bool] = False
    try:
        if kind == "excel":
            return app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=readonly, Password=_NO_PASSWORD,
                                      IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
        if kind == "word":
            # Word 2007은 ReadOnly와 PasswordDocument를 함께 주면 암호 없는 파일도 열지 못하므로 암호가 없다고 확인된 파일에는 넘기지 않음
            encrypted = _word_encrypted(path)
            password = {} if encrypted is False else {"PasswordDocument": _NO_PASSWORD}
            return app.Documents.Open(path, ConfirmConversions=False, ReadOnly=readonly, AddToRecentFiles=False,
                                      Visible=False, **password)
        return app.Presentations.Open(path, -1 if readonly else 0, 0, 0)   # ReadOnly, Untitled, WithWindow=False
    except Exception as e:
        info = rt._com_error_info(e)
        text = (info[1] if info else str(e)).lower()
        if encrypted is True or "암호" in text or "password" in text:    # Word 2007의 오류 문구에는 '암호'가 없음
            raise _password_error(path)
        raise


def _opened_readonly(doc: Any) -> bool:
    try:
        return bool(doc.ReadOnly)                     # Excel/Word True/False, PowerPoint -1/0
    except Exception:
        return False


def word_save_as(doc: Any, path: str, fmt: Optional[int] = None) -> None:
    """Word SaveAs2(2010+)를 쓰고, 없는 버전(Word 2007)은 SaveAs로 저장합니다."""
    try:
        save = doc.SaveAs2
    except AttributeError:
        save = doc.SaveAs
    save(path, FileFormat=fmt) if fmt is not None else save(path)


def _create_file(kind: str, app: Any, path: str) -> Any:
    fmt = EXTENSIONS[kind].get(os.path.splitext(path)[1].lower())
    if kind == "excel":
        doc = app.Workbooks.Add()
        doc.SaveAs(path, FileFormat=fmt) if fmt is not None else doc.SaveAs(path)
    elif kind == "word":
        doc = app.Documents.Add(Visible=False)
        word_save_as(doc, path, fmt)
    else:
        doc = app.Presentations.Add(0)
        doc.SaveAs(path, fmt) if fmt is not None else doc.SaveAs(path)
    return doc


def _file_stamp(path: str) -> Optional[Tuple[float, int]]:
    """디스크 파일의 (수정 시각, 크기). 없거나 읽지 못하면 None."""
    try:
        st = os.stat(path)
        return (st.st_mtime, st.st_size)
    except (OSError, TypeError, ValueError):
        return None


def resolve_target(kind: str, target: str = "active", readonly: bool = True, create: bool = False,
                   option: str = "--target") -> Dict[str, Any]:
    """대상 문서를 바꾸지 않고 찾습니다 (open_target의 찾기 단계, 앱을 띄우거나 파일을 열지 않음).
    {"source": "active"|"open", "app", "doc"} 또는 {"source": "file", "path", "exists"}.
    대상을 찾지 못한 오류(오타 [추천], [후보], 확장자 등)는 no_fallback: 방식을 바꿔도 같은 대상이라 transfer가 다음 방식으로
    넘어가지 않고 원래 오류 그대로 멈춥니다 (클립보드를 쓰는 relay 등이 대상 오타로 실행되지 않게)."""
    target = (target or "active").strip()
    try:
        if target.lower() == "active":
            try:
                info = get_active(kind)
            except rt.UserDecisionError:
                if option == "--target":
                    raise
                raise no_open_doc(kind, option) from None        # 이 명령에 없는 --target을 안내하지 않음
            return {"source": "active", "app": info["app"], "doc": info["doc"]}
        docs = list_open(kind)
        hit = match_open(kind, target, docs, option)
        if hit is not None:
            return {"source": "open", "app": hit["app"], "doc": hit["doc"]}
        # 전체 경로인지는 %USERPROFILE%·~ 를 푼 뒤에 판단 (%USERPROFILE%\Desktop\새.xlsx도 전체 경로)
        expanded = os.path.expanduser(os.path.expandvars(target.strip('"')))
        is_abs = os.path.isabs(expanded)
        path = os.path.abspath(expanded)
        exists = os.path.isfile(path)
        if not exists and not is_abs and os.path.isfile(os.path.join(rt.get_data_dir(), target)):
            path, exists = os.path.join(rt.get_data_dir(), target), True
        if not exists and (readonly or not create or not is_abs):
            # 파일 이름만 주면 새로 만들지 않음: 열린 문서 이름의 오타로 현재 폴더에 새 파일이 생기고 거기에 쓰는 것을 막음
            names = [d["name"] for d in docs]
            tail = "" if readonly or not create else f" 새 파일은 사용자에게 전체 경로를 물어 {option}으로 줄 때만 만듭니다."
            hint = "" if names else rt.elevation_hint()      # 관리자 권한이면 열린 문서가 안 보여 0개일 수 있음
            err = rt.not_found(f"{APP_LABEL[kind]} 문서", target, names,
                               extra=" (열린 문서 이름도 아니고 파일도 없음)." + tail + (f" {hint}" if hint else ""))
            # 파일 이름만 준 경우에만 열린 문서를 [후보]로 (전체 경로가 없는 것은 열린 문서 이름의 오타가 아님)
            if names and not is_abs and not err.candidates and "[추천]" not in str(err):
                err = rt.ambiguous(f"{APP_LABEL[kind]} 문서 '{target}'을(를) 찾을 수 없습니다 (열린 문서 이름도 아니고 파일도 없음)."
                                   f"{tail} 열린 문서 중에서 사용자에게 고르게 하세요.", names)
            raise err
        ext = os.path.splitext(path)[1].lower()
        if ext not in EXTENSIONS[kind]:
            raise rt.InputError(f"{APP_LABEL[kind]} 파일 확장자가 아닙니다: {ext} (지원: {', '.join(EXTENSIONS[kind])})")
        return {"source": "file", "path": path, "exists": exists}
    except rt.InputError as e:
        e.no_fallback = True
        raise


@contextmanager
def open_target(kind: str, target: str = "active", readonly: bool = True, snapshot: Optional[bool] = None,
                focus: bool = True, create: bool = False, tool: str = "", command: str = "",
                option: str = "--target") -> Iterator[OfficeTarget]:
    """대상 문서를 열어 OfficeTarget을 넘겨줍니다.

    target: "active" | 열린 문서 이름 | 파일 경로
    readonly: 읽기 세션 (저장/생성/화면 이동 없음)
    snapshot: 쓰기 전 스냅샷 (기본: 쓰기 세션이면 True)
    focus: 쓰기 후 고친 위치로 화면 이동 (--no-focus면 False)
    create: 파일이 없을 때 새로 만들지 (쓰기 세션 전용)
    option: 오류 안내에 쓸 대상 옵션 이름 (기본 --target, office.py transfer는 --from-target/--to-target)
    """
    require_com()
    init_com()
    if snapshot is None:
        snapshot = not readonly
    target = (target or "active").strip()

    app = doc = None
    owns_app = owns_doc = False
    restore_alerts = None
    t: Optional[OfficeTarget] = None
    try:
        found = resolve_target(kind, target, readonly, create, option)
        if found["source"] in ("active", "open"):
            t = OfficeTarget(kind, found["app"], found["doc"], found["source"], readonly)
        else:
            path, exists = found["path"], found["exists"]
            ppt_was_running = _ppt_running() if kind == "ppt" else False
            timeout = open_timeout()
            # 감시할 프로세스 번호: Excel은 창 핸들로, Word·PowerPoint는 띄우기 전후 목록 차이로 (사용자가 띄운 PowerPoint는 감시 안 함)
            before = _office_pids(kind) if timeout and kind != "excel" and not ppt_was_running else set()
            app = _new_instance(kind)
            owns_app = kind != "ppt" or not ppt_was_running
            if kind == "ppt" and owns_app:
                # 사용자가 이미 띄운 PowerPoint의 DisplayAlerts는 바꾸지 않음 (열기 감시가 도구를 끝내면 되돌릴 수 없음)
                try:
                    restore_alerts = app.DisplayAlerts
                    app.DisplayAlerts = 1  # ppAlertsNone
                except Exception:
                    restore_alerts = None
            rt.log("PROGRESS", f"{APP_LABEL[kind]} 파일 {'여는' if exists else '만드는'} 중 (백그라운드): {os.path.basename(path)}")
            pid = _instance_pid(kind, app, before) if timeout and owns_app else None
            with _OpenWatch(kind, pid, timeout, path, tool) as watch:
                try:
                    doc = _open_file(kind, app, path, readonly) if exists else _create_file(kind, app, path)
                except BaseException:
                    if watch.fired:          # 감시 시간이 지나 그 인스턴스를 끝냄 -> 멈춘 COM 호출이 오류로 돌아옴
                        raise _drm_or_dialog_error(kind, path, tool, timeout) from None
                    raise
            if watch.fired:
                raise _drm_or_dialog_error(kind, path, tool, timeout)
            owns_doc = True
            t = OfficeTarget(kind, app, doc, "file" if exists else "new", readonly)
            if exists and not readonly and _opened_readonly(doc):
                # 다른 사람·프로그램이 열고 있어 읽기 전용으로 열림: 고친 내용을 저장할 수 없으므로 쓰기 전에 멈춤
                raise rt.InputError(f"'{os.path.basename(path)}'을(를) 다른 사용자나 프로그램이 열고 있어 읽기 전용으로만 열립니다. "
                                    "아무것도 바꾸지 않았습니다. 그 파일을 닫은 뒤 다시 실행하거나, 이 PC의 Office에서 열어 둔 뒤 요청하세요.",
                                    no_fallback=True)

        point: Optional[str] = None       # 이 실행이 만든 되돌리기 지점(.json)
        was_saved = is_saved(t.doc)
        disk_before = _file_stamp(t.path) if owns_doc else None
        if snapshot and not readonly:
            # 정리(rotate)는 명령이 바꾼 뒤에만: 쓰기 전에 멈춘 실행이 진짜 옛 지점을 밀어내지 않도록
            if t.source == "new":
                point = save_barrier(kind, t.doc, "새로 만든 파일이라 되돌릴 이전 상태가 없음", tool, command, rotate=False)
            else:
                data_file = save_snapshot(kind, t.doc, tool, command, rotate=False)
                if data_file is None:
                    point = save_barrier(kind, t.doc, "되돌리기용 스냅샷을 만들지 못함 (DRM, 크기 초과 등)", tool, command,
                                         rotate=False)
                else:
                    point = os.path.splitext(data_file)[0] + ".json"

        def unchanged(e: Optional[BaseException]) -> bool:
            """이 실행이 문서를 바꾸지 않았는지 (그러면 이번 지점은 지움: 남기면 다음 undo가 바뀐 것 없는 이 지점으로
            되돌려 실제로 바꾼 작업이 그대로 남음). 오류로 멈춘 실행(e)과 바꾼 것 없이 성공한 실행(e None: 0곳 바꿈,
            차트 0개 삭제, 이미 있는 시트 add 등) 모두 같은 기준.
            - 도구가 열거나 만든 파일(owns_doc): 바뀐 것만 저장하므로 디스크 파일이 그대로면 바뀐 것 없음
            - 열린 문서: 바꾼 표시(t.changed)가 없고 저장 상태가 그대로일 때만. 원래 저장 안 된 문서가 오류로 멈췄으면
              입력 오류일 때만 (그 밖의 오류는 중간까지 바꿨는지 알 수 없어 지점을 남김)"""
            if owns_doc and disk_before is not None:
                return _file_stamp(t.path) == disk_before
            if t.changed or is_saved(t.doc) != was_saved:
                return False
            return e is None or was_saved or isinstance(e, rt.InputError)

        def drop_or_keep(e: Optional[BaseException]) -> None:
            if not point:
                return
            if unchanged(e):
                delete_snapshot(point)
                rt.debug("문서를 바꾸지 않아 이번 되돌리기 지점을 지웠습니다.")
            else:
                rotate_point(point)

        try:
            yield t
            if not readonly and focus:
                t.apply_focus()
            # 바꾼 것(t.changed)이 있을 때만 저장: 휘발성 수식(NOW, RAND 등)이 있으면 열기만 해도 저장 안 된 상태라
            # 바꾼 것 없이 성공한 실행이 사용자 파일을 다시 저장하고 바뀐 것 없는 되돌리기 지점을 남겼음
            if owns_doc and not readonly and t.changed and not is_saved(t.doc):
                rt.log("PROGRESS", "변경사항 저장 중")
                t.doc.Save()
        except BaseException as e:
            drop_or_keep(e)
            raise
        drop_or_keep(None)
    finally:
        own_path = ""
        if owns_doc and t is not None:
            try:
                own_path = t.path
            except Exception:
                own_path = ""
            try:
                if kind == "word":
                    t.doc.Close(SaveChanges=0)
                elif kind == "excel":
                    t.doc.Close(SaveChanges=False)
                else:
                    t.doc.Close()
            except Exception:
                pass
        if app is not None:
            quit_ = False
            if owns_app:
                # Word·PowerPoint: 자기 문서를 닫은 뒤 남은 문서가 없을 때만 끝냄 (사용자가 그 숨은 인스턴스에서 연 문서의
                # 변경을 버리지 않게). Excel은 도구 전용 새 인스턴스라 늘 끝냄
                if kind == "excel":
                    quit_ = True
                else:
                    others = _other_docs_left(kind, app, own_path)
                    quit_ = others == 0
                    if others is None:
                        rt.log("INFO", f"{APP_LABEL[kind]}의 열린 문서 수를 읽지 못해 {APP_LABEL[kind]}을(를) 끝내지 않고 그대로 둡니다.")
                    elif others:
                        rt.log("INFO", f"{APP_LABEL[kind]}에 다른 문서 {others}개가 열려 있어 {APP_LABEL[kind]}을(를) 끝내지 않고 "
                                       "그대로 둡니다.")
            if restore_alerts is not None and not quit_:
                try:
                    app.DisplayAlerts = restore_alerts
                except Exception:
                    pass
            if quit_:
                try:
                    if kind == "word":
                        app.Quit(SaveChanges=0)      # 남은 문서가 없음을 확인한 뒤라 버릴 변경이 없음
                    else:
                        app.Quit()
                except Exception:
                    pass
        t = None
        app = doc = None
        import gc
        gc.collect()


QUIT_COUNT_TRIES = 5          # 남은 문서 수 읽기 재시도 (Office가 바빠 COM 호출이 실패할 때)
QUIT_COUNT_DELAY = 0.3


def _other_docs_left(kind: str, app: Any, own_path: str) -> Optional[int]:
    """도구가 띄운 Word·PowerPoint에 남은 문서 중 이 실행의 문서(own_path)가 아닌 것의 수. 끝까지 못 읽으면 None.
    자기 문서를 닫다가 실패해 남아 있어도(닫기 오류는 삼킴) 그 문서뿐이면 0이라 끌 수 있음."""
    for i in range(QUIT_COUNT_TRIES):
        try:
            docs = app.Documents if kind == "word" else app.Presentations
            n = int(docs.Count)
            if n == 0:
                return 0
            if not own_path:
                return n
            own = os.path.normcase(os.path.abspath(own_path))
            others = 0
            for k in range(1, n + 1):
                full = str(getattr(docs.Item(k), "FullName", "") or "")
                if not full or os.path.normcase(os.path.abspath(full)) != own:
                    others += 1
            return others
        except Exception:
            if i + 1 < QUIT_COUNT_TRIES:
                time.sleep(QUIT_COUNT_DELAY)
    return None


@contextmanager
def excel_batch(app: Any) -> Iterator[None]:
    """Excel 대량 쓰기 동안 화면 갱신/자동 계산/이벤트를 끄고, 끝나면 원래대로 되돌립니다."""
    saved = {}
    for attr, value in (("ScreenUpdating", False), ("EnableEvents", False), ("Calculation", -4135)):  # xlCalculationManual
        try:
            saved[attr] = getattr(app, attr)
            setattr(app, attr, value)
        except Exception:
            pass
    try:
        yield
    finally:
        for attr in ("Calculation", "EnableEvents", "ScreenUpdating"):
            if attr in saved:
                try:
                    setattr(app, attr, saved[attr])
                except Exception:
                    pass


def range_to_matrix(values: Any) -> List[List[Any]]:
    """Range.Value/Formula 결과를 항상 2차원 리스트로 (단일 셀은 스칼라, 여러 셀은 한 행이어도 튜플의 튜플)."""
    if not isinstance(values, (tuple, list)):
        return [[values]]
    return [list(row) if isinstance(row, (tuple, list)) else [row] for row in values]


def _office12_pdf_addin() -> bool:
    """Office 2007은 'PDF/XPS로 저장' 추가 기능(EXP_PDF.DLL)이 있어야 PDF로 내보낼 수 있습니다 (SP2 이상은 기본 포함)."""
    for env in ("CommonProgramFiles", "CommonProgramFiles(x86)", "CommonProgramW6432"):
        base = os.environ.get(env)
        if base and os.path.exists(os.path.join(base, "Microsoft Shared", "OFFICE12", "EXP_PDF.DLL")):
            return True
    return False


def office_caps(app: Any) -> Dict[str, Any]:
    """Office 버전과 PDF 내보내기 가능 여부. pdf_export: True | False | None(알 수 없음)."""
    try:
        version = str(app.Version)
        major = int(float(version))
    except Exception:
        return {"version": None, "major": None, "pdf_export": None}
    pdf = True if major >= 14 else (_office12_pdf_addin() if major == 12 else False)
    return {"version": version, "major": major, "pdf_export": pdf}


def require_pdf_export(app: Any, alt_ext: str) -> None:
    """PDF 내보내기 기능이 없는 Office면 시도하지 않고 입력 오류로 안내합니다 (버전을 모르면 시도)."""
    caps = office_caps(app)
    if caps["pdf_export"] is False:
        hint = " (Office 2007은 'PDF/XPS로 저장' 추가 기능 필요)" if caps["major"] == 12 else ""
        raise rt.InputError(f"이 PC의 Office {caps['version']}에는 PDF 내보내기 기능이 없습니다{hint}. "
                            f"--save에 {alt_ext}를 쓰세요.")


def running_office_caps() -> Dict[str, Dict[str, Any]]:
    """실행 중인 Office 앱별 버전과 PDF 내보내기 가능 여부 (앱을 새로 띄우지 않음)."""
    caps: Dict[str, Dict[str, Any]] = {}
    for kind in KINDS:
        try:
            apps = find_apps(kind)
            if apps:
                caps[APP_LABEL[kind]] = office_caps(apps[0])
        except Exception as e:
            rt.debug(f"{APP_LABEL[kind]} 버전 확인 실패: {e}")
    return caps


def open_documents_summary() -> List[Dict[str, Any]]:
    """세 앱의 열린 문서를 한 목록으로 (office.py windows, *.py info --all 용)."""
    result = []
    for kind in KINDS:
        try:
            for d in list_open(kind):
                result.append({"app": APP_LABEL[kind], "name": d["name"], "path": d["path"],
                               "saved": d["saved"], "active": d["active"]})
        except rt.InputError:
            raise
        except Exception as e:
            rt.debug(f"{APP_LABEL[kind]} 목록 실패: {e}")
    return result


# ==============================================================================
# 4. 색상과 테마 (Excel/Word/PPT 공통)
# ==============================================================================

# 이름 -> #RRGGBB
COLOR_NAMES = {
    "black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "green": "#008000", "blue": "#0000FF",
    "yellow": "#FFFF00", "orange": "#FFA500", "purple": "#800080", "gray": "#808080", "grey": "#808080",
    "lightgray": "#D3D3D3", "lightblue": "#ADD8E6", "lightgreen": "#90EE90", "lightyellow": "#FFFFE0",
    "pink": "#FFC0CB", "navy": "#1B365D", "darkblue": "#00008B", "darkred": "#8B0000", "skyblue": "#87CEEB",
    "빨강": "#FF0000", "빨간색": "#FF0000", "파랑": "#0000FF", "파란색": "#0000FF", "초록": "#008000", "초록색": "#008000",
    "노랑": "#FFFF00", "노란색": "#FFFF00", "주황": "#FFA500", "주황색": "#FFA500", "회색": "#808080", "검정": "#000000",
    "검은색": "#000000", "흰색": "#FFFFFF", "하늘색": "#87CEEB", "분홍": "#FFC0CB", "보라": "#800080",
}


def parse_color(value: Any) -> Optional[int]:
    """'#RRGGBB', 'RRGGBB', 색 이름(영문/한글) -> Office COM용 BGR 정수. 검정(0)도 올바르게 처리합니다.
    값이 없으면 None, 해석할 수 없으면 InputError."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        raise rt.InputError(f"색상 값이 잘못되었습니다: {value}")
    if isinstance(value, int):
        return value
    s = str(value).strip()
    hexval = COLOR_NAMES.get(s.lower(), s)
    h = hexval.lstrip("#")
    if len(h) == 6:
        try:
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            return (b << 16) | (g << 8) | r
        except ValueError:
            pass
    raise rt.not_found("색상", s, list(COLOR_NAMES), extra=" ('#RRGGBB' 형식도 가능)", vocab=True)


THEMES: Dict[str, Dict[str, Any]] = {
    "navy": {"label": "네이비", "primary": "#1B365D", "secondary": "#4A777A", "accent": "#D9381E",
             "header_bg": "#1B365D", "header_fg": "#FFFFFF", "zebra": "#F4F7FB", "summary_bg": "#E8EEF5",
             "light": "#E8EEF5", "number_format": "#,##0"},
    "slate": {"label": "슬레이트", "primary": "#2B2D42", "secondary": "#8D99AE", "accent": "#EF233C",
              "header_bg": "#2B2D42", "header_fg": "#FFFFFF", "zebra": "#F8F9FA", "summary_bg": "#EDF2F4",
              "light": "#EDF2F4", "number_format": "#,##0"},
    "teal": {"label": "틸", "primary": "#004D40", "secondary": "#00897B", "accent": "#FF6F00",
             "header_bg": "#004D40", "header_fg": "#FFFFFF", "zebra": "#F0FDF4", "summary_bg": "#E0F2F1",
             "light": "#E0F2F1", "number_format": "#,##0"},
    "emerald": {"label": "에메랄드", "primary": "#1E4620", "secondary": "#3E7B43", "accent": "#E65100",
                "header_bg": "#1E4620", "header_fg": "#FFFFFF", "zebra": "#F2F8F2", "summary_bg": "#E5EFE5",
                "light": "#E5EFE5", "number_format": "#,##0"},
}


def get_theme(name: Optional[str]) -> Optional[Dict[str, Any]]:
    """테마 프리셋 (navy, slate, teal, emerald). 없으면 None, 모르는 이름이면 [추천]과 함께 InputError."""
    if not name:
        return None
    key = str(name).strip().lower()
    if key in THEMES:
        return dict(THEMES[key], name=key)
    raise rt.not_found("테마", str(name), list(THEMES), vocab=True)
