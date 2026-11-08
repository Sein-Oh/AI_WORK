"""실기 테스트 러너(run_win_tests.py)의 판단 규칙 (표준 라이브러리만, pywin32 없이 import 가능).

러너와 정리 스크립트(CLEANUP_SCRIPT)가 같이 쓰고, tests/test_win_runner.py가 Windows 없이 검증합니다.
- 어떤 Office 프로세스를 끝내도 되는지 (사용자 문서를 가진 프로세스는 절대 끝내지 않음)
- COM 오류가 '문서/항목 없음'(재시도해도 소용없음)인지
- tasklist 출력에서 PID 목록 읽기
"""
import csv
import io

# 러너가 화면에 열어 두고 모든 그룹이 쓰는 테스트 문서
TEST_DOCS = {"excel": "A.xlsx", "word": "A.docx", "ppt": "A.pptx"}

# 문서가 없는 Office 창의 제목 (Office 16)
EMPTY_TITLES = ("", "Excel", "PowerPoint")

DISP_E_BADINDEX = 0x8002000B          # Workbooks("A.xlsx") 같은 이름 조회 실패 (문서가 닫힘)
MK_E_UNAVAILABLE = 0x800401E3         # GetActiveObject 실패 (그 Office 앱이 아예 실행 중이 아님)
MISSING_CODES = (DISP_E_BADINDEX, MK_E_UNAVAILABLE)
_MASK = 0xFFFFFFFF


def _com_codes(exc):
    if type(exc).__name__ != "com_error":
        return []
    args = getattr(exc, "args", ()) or ()
    codes = []
    if args and isinstance(args[0], int):
        codes.append(args[0] & _MASK)
    if len(args) > 2 and args[2] and len(args[2]) >= 6 and isinstance(args[2][5], int):
        codes.append(args[2][5] & _MASK)
    return codes


def is_missing_doc_error(exc):
    """pywintypes.com_error가 '그 이름의 항목 없음'(DISP_E_BADINDEX)인지 (예: 닫힌 A.xlsx를 Workbooks("A.xlsx")로 찾음).
    Office가 바쁜 것이 아니므로 재시도해도 소용없습니다 (러너의 com_retry가 바로 실패시킴)."""
    return DISP_E_BADINDEX in _com_codes(exc)


def is_doc_gone_error(exc):
    """문서가 닫혔거나 그 Office 앱이 아예 실행 중이 아님(MK_E_UNAVAILABLE). 테스트 문서를 다시 열지 판단할 때 씁니다."""
    return any(c in MISSING_CODES for c in _com_codes(exc))


def parse_tasklist_pids(text, image):
    """`tasklist /FI "IMAGENAME eq X" /FO CSV /NH` 출력에서 PID 목록을 읽습니다."""
    pids = []
    for row in csv.reader(io.StringIO(text or "")):
        if len(row) >= 2 and row[0].strip().lower() == image.lower():
            try:
                pids.append(int(row[1].strip()))
            except ValueError:
                continue
    return pids


def window_shows_document(title, visible):
    """보이는 창이고 제목이 빈 Office 창 제목이 아니면 문서를 가진 창으로 봅니다 (예: 'A.xlsx - Excel')."""
    return bool(visible) and str(title or "").strip() not in EMPTY_TITLES


def protected_pids(windows, doc_pids=(), protect_names=()):
    """끝내면 안 되는 프로세스.
    - doc_pids: COM으로 확인한 '문서가 남아 있는' 앱의 PID
    - 문서를 보여 주는 창(보이는 창 + 제목 있음)을 가진 PID
    - 창 제목에 테스트 문서 이름(A.xlsx 등)이 들어 있는 PID (숨은 창 포함)"""
    result = {int(p) for p in doc_pids if p is not None}
    for w in windows:
        title = str(w.get("title") or "")
        if window_shows_document(title, w.get("visible")) or any(title_has_doc(title, n) for n in protect_names):
            result.add(w["pid"])
    return result


def title_has_doc(title, name):
    """창 제목에 문서 이름이 들어 있는지 ('A.xlsx - Excel', 'C:\\t\\A.xlsx [읽기 전용]'). 'Data.xlsx'는 A.xlsx가 아님."""
    t, n = str(title or "").lower(), str(name or "").lower()
    if not n:
        return False
    i = t.find(n)
    while i >= 0:
        if i == 0 or not (t[i - 1].isalnum() or t[i - 1] in "_-."):
            return True
        i = t.find(n, i + 1)
    return False


def may_end_process(pid, protected, allowed_pids=None, run_start=None, started=None):
    """이 프로세스를 끝내도 되는지와 이유. (bool, 이유)
    - 보호 PID(문서를 가진 프로세스)는 절대 끝내지 않음
    - allowed_pids를 주면(예: A5가 직접 띄운 두 번째 Excel) 그 PID만
    - 아니면 이 실행이 시작된 뒤 생긴 프로세스만 (시작 시각을 모르면 끝내지 않음)"""
    if pid is None:
        return False, "PID를 알 수 없음"
    if pid in protected:
        return False, "문서가 열려 있는 프로세스 (보호)"
    if allowed_pids is not None:
        if pid in allowed_pids:
            return True, "이 단계가 띄운 프로세스"
        return False, "이 단계가 띄운 프로세스가 아님"
    if run_start is None or started is None:
        return False, "시작 시각을 알 수 없음"
    if started < run_start:
        return False, "테스트 실행 전부터 있던 프로세스"
    return True, "테스트 실행 중 생긴 빈 프로세스"


def select_close_targets(windows, protected, allowed_pids=None, run_start=None, starts=None):
    """WM_CLOSE를 보낼 창을 고릅니다. windows: [{hwnd, pid, cls, title, visible}].
    후보는 숨은 창이거나 제목이 빈 창(문서 없음)뿐이고, 프로세스 단위로 may_end_process를 통과해야 합니다.
    반환: (보낼 창 목록, [(건너뛴 후보 창, 이유)])"""
    starts = starts or {}
    targets, skipped = [], []
    for w in windows:
        title = str(w.get("title") or "").strip()
        if w.get("visible") and title not in EMPTY_TITLES:
            continue                                  # 문서를 보여 주는 창은 후보가 아님
        ok, why = may_end_process(w.get("pid"), protected, allowed_pids, run_start, starts.get(w.get("pid")))
        (targets.append(w) if ok else skipped.append((w, why)))
    return targets, skipped


def may_quit_app(pid, doc_count, protected, allowed_pids=None, run_start=None, started=None):
    """COM Quit을 불러도 되는지. 문서가 0개이고, PID를 알면 may_end_process 규칙을 따릅니다.
    PID를 모르면(Word 등 창 핸들을 못 얻은 경우) 특정 PID만 허용한 단계(allowed_pids)에서는 끝내지 않고,
    전체 정리에서는 문서 0개만 확인합니다 (Quit은 문서가 없을 때만 정상 종료되므로)."""
    if doc_count:
        return False, f"문서 {doc_count}개가 열려 있음"
    if pid is None:
        if allowed_pids is not None:
            return False, "PID를 알 수 없어 이 단계가 띄운 프로세스인지 확인 불가"
        return True, "문서 0개 (PID 모름)"
    return may_end_process(pid, protected, allowed_pids, run_start, started)
