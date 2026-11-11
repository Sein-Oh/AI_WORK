"""excel.py - 엑셀 작업

명령:
    read     내용 읽기 (범위, 검색, --save로 CSV 저장, --formulas로 수식 칸 목록)
    info     통합문서 요약 (시트, 사용 범위, 차트, 선택 영역). --all: 열린 통합문서 전체 목록
    write    값/표/수식 쓰기 (--value, --values, --file, --formula, --header, --append)
    replace  찾아 바꾸기 (수식 셀은 보존)
    style    서식 (테마, 배경/글꼴/정렬/숫자서식/테두리, 조건부 강조)
    export   CSV / XLSX / PDF로 저장 (원본 창은 그대로)
    sheet    시트 목록/추가/이름변경/복사/삭제/선택
    range    너비/높이/자동맞춤/병합/틀고정/숨김/지우기/복사/이동/테두리
    chart    차트 만들기/목록/삭제
    comment  칸 메모 목록/달기/지우기
    check    표의 값을 등록된 규격서와 대조 (판정과 근거, 칸 메모 파일)

--target: active(기본, 지금 보고 있는 통합문서) | 열린 통합문서 이름 | 파일 경로
닫힌 xlsx/xlsm 파일 읽기는 Excel을 띄우지 않고 openpyxl로 읽습니다.

예:
    python excel.py write --range A1:B101 --header "x,sin(x)" --formula "=(ROW()-2)*4*PI()/99" "=SIN(A2)"
    python excel.py chart --range A1:B101 --type line --title "사인 2주기"
    python excel.py write --file "결과.csv" --sheet "검색결과" --new-sheet
    python excel.py read --range A1:D20
"""
import datetime as _dt
import json
import os
import re
import unicodedata
import sys
import zipfile
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _office as off  # noqa: E402
import _runtime as rt  # noqa: E402
import _markdown as md  # noqa: E402
import _xlsx_formula as xf  # noqa: E402

# 직접 실행(__main__)해도 내부 모듈(_spec_check)의 'from excel import ...'가 이 모듈을 쓰게 함 (두 번 읽지 않음)
sys.modules.setdefault("excel", sys.modules[__name__])

TOOL = "excel"

XL_CHART = {"column": 51, "bar": 57, "line": 4, "line_markers": 65, "pie": 5, "area": 1,
            "scatter": -4169, "scatter_lines": 75, "combo": 51}   # -4169 표식만, 75 선만(표식 없음)
XL_LEGEND = {"bottom": -4107, "right": -4152, "top": -4160, "left": -4131}
XL_ALIGN = {"left": -4131, "center": -4108, "right": -4152}
XL_WEIGHT = {"hairline": 1, "thin": 2, "medium": -4138, "thick": 4}
XL_OPERATOR = {">": 5, ">=": 7, "<": 6, "<=": 8, "=": 3, "<>": 4, "!=": 4}
XL_A1, XL_R1C1 = 1, -4150    # ReferenceStyle (ConvertFormula)


# ==============================================================================
# 1. 주소와 값 도우미
# ==============================================================================

def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


XL_MAX_ROW = 1048576
XL_MAX_COL = 16384          # XFD


def _check_row(row: int, ref: str) -> int:
    if not 1 <= row <= XL_MAX_ROW:
        raise rt.InputError(f"행 번호가 엑셀 범위(1~{XL_MAX_ROW}) 밖입니다: '{ref}' (예: A1, B12)")
    return row


def parse_cell(ref: str) -> Tuple[int, int]:
    """'B12' -> (12, 2). 0행, 1048576행 초과, XFD(16384열) 초과는 InputError."""
    m = _CELL_RE.match(ref.strip())
    if not m:
        raise rt.InputError(f"셀 주소 형식이 아닙니다: '{ref}' (예: A1, B12)")
    col = col_number(m.group(1))
    if col > XL_MAX_COL:
        raise rt.InputError(f"열이 엑셀 범위(A~XFD) 밖입니다: '{ref}' (예: A1, B12)")
    return _check_row(int(m.group(2)), ref), col


def parse_area(cells: str, max_row: Optional[int], max_col: Optional[int]) -> Tuple[int, int, Optional[int], Optional[int]]:
    """'A1:C5', 'B:B', 'B:D', '5:7', 'A1' -> (r1, c1, r2, c2). 열/행 전체는 데이터 끝까지 (None = 끝까지)."""
    parts = cells.replace("$", "").strip().split(":")
    a, b = parts[0], parts[-1]
    if re.fullmatch(r"[A-Za-z]{1,3}", a) and re.fullmatch(r"[A-Za-z]{1,3}", b):
        ca, cb = sorted((col_number(a), col_number(b)))
        if cb > XL_MAX_COL:
            raise rt.InputError(f"열이 엑셀 범위(A~XFD) 밖입니다: '{cells}' (예: B:B, A1:C5)")
        return 1, ca, max_row, cb
    if a.isdigit() and b.isdigit():
        ra, rb = sorted((_check_row(int(a), cells), _check_row(int(b), cells)))
        return ra, 1, rb, max_col
    r1, c1 = parse_cell(a)
    r2, c2 = parse_cell(b)
    # 거꾸로 준 범위(C5:A1)는 Excel처럼 A1:C5로 (닫힌 파일에서도 같게)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def split_sheet(address: str) -> Tuple[Optional[str], str]:
    """"'My Sheet'!A1:B2" -> ('My Sheet', 'A1:B2'), 'A1' -> (None, 'A1')"""
    if "!" not in address:
        return None, address.strip()
    sheet, cells = address.rsplit("!", 1)
    sheet = sheet.strip()
    if len(sheet) >= 2 and sheet[0] == "'" and sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, cells.strip()


_CELL_ADDR = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$")


def is_cell_address(text: str) -> bool:
    """'A1', '$B$2', 'A1:C5' 꼴인지 (열·행 전체 'B:B', '3:3'은 아님)."""
    return bool(_CELL_ADDR.match((text or "").strip()))


def addr(r1: int, c1: int, r2: Optional[int] = None, c2: Optional[int] = None) -> str:
    a = f"{col_letter(c1)}{r1}"
    if r2 is None or (r2 == r1 and c2 == c1):
        return a
    return f"{a}:{col_letter(c2)}{r2}"


def out_value(v: Any) -> Any:
    """Excel 값 -> JSON/CSV용 값 (정수형 실수는 정수로, 날짜는 ISO 문자열로, 기간([h]:mm 칸)은 '30:00:00')."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, _dt.timedelta):          # '1 day, 6:00:00'으로 나오던 것 (2026-11-09 검토 B13)
        total = int(round(v.total_seconds()))
        sign, total = ("-", -total) if total < 0 else ("", total)
        return f"{sign}{total // 3600}:{total % 3600 // 60:02d}:{total % 60:02d}"
    if isinstance(v, float):
        return int(v) if v.is_integer() and abs(v) < 1e15 else v
    if hasattr(v, "isoformat") and hasattr(v, "year"):
        try:
            if getattr(v, "hour", 0) or getattr(v, "minute", 0) or getattr(v, "second", 0):
                return v.strftime("%Y-%m-%d %H:%M:%S")
            return v.strftime("%Y-%m-%d")
        except Exception:
            return str(v)
    return v


_NUM_RE = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")


def coerce(s: Any) -> Any:
    """CSV 문자열 -> Excel 값. 숫자는 숫자로, 앞자리 0이 있는 코드(001)와 15자리를 넘는 숫자는 문자 그대로."""
    if not isinstance(s, str):
        return s
    t = s.strip()
    if t == "":
        return None
    if _NUM_RE.match(t):
        digits = t.lstrip("+-").replace(",", "")
        if len(digits) > 1 and digits[0] == "0" and not digits.startswith("0."):
            return t
        if len(digits.replace(".", "")) > 15:       # Excel 숫자 정밀도(15자리)를 넘으면 값이 바뀜
            return t
        num = float(digits) * (-1 if t.startswith("-") else 1)
        return int(num) if num.is_integer() and "." not in digits else num
    return s


# Excel이 COM으로 받은 글자를 직접 입력한 것처럼 다시 해석해 값이 바뀌는 모양:
# 앞자리 0 코드(01234 -> 1234), 15자리 초과 숫자, 월-일/분수처럼 보이는 글(1-2, 3/4 -> 날짜), 괄호 숫자((100) -> -100), 지수(1E5)
_REINTERPRETED = re.compile(r"^(?:0\d+(?:\.\d+)?|[+-]?\d{16,}(?:\.\d+)?|\d{1,2}\s*[-/]\s*\d{1,2}|\(\s*[\d.,]+\s*\)|\d+(?:\.\d+)?[eE][+-]?\d+)$")


def xl_text(v: Any) -> Any:
    """셀에 쓸 값. 글자인데 Excel이 숫자/날짜로 바꿔 버릴 모양이면 앞에 '를 붙여 글자 그대로 둡니다 ('는 셀 값에 남지 않음).
    + - @ = 로 시작하는 글도 (Excel이 수식으로 계산: '+82-10-1234-5678' -> -6840, '- 검토 필요' -> #NAME?, 2026-11-09 검토 A8).
    수식은 --formula와 --cells의 == 경로로만 씁니다."""
    if isinstance(v, str) and (_REINTERPRETED.match(v.strip()) or v.lstrip()[:1] in ("+", "-", "@", "=")):
        return "'" + v
    return v


def rect(rows: List[List[Any]]) -> List[List[Any]]:
    width = max((len(r) for r in rows), default=0)
    return [list(r) + [None] * (width - len(r)) for r in rows]


TABLE_FILE_EXTS = (".csv", ".tsv", ".md", ".txt")  # load_table_file이 읽는 표 파일 (못 찾을 때 후보 목록도 이것만)


def load_table_file(path: str, raw: bool = False) -> List[List[Any]]:
    """CSV/TSV, 마크다운 표(.md의 첫 번째 표), 일반 텍스트(한 줄 = 한 행)를 표로 읽습니다.
    raw=True면 칸 값을 글자 그대로 (숫자로 바꾸지 않음: 메모 파일)."""
    ext = os.path.splitext(path)[1].lower()
    conv = (lambda c: c) if raw else coerce
    if ext in (".csv", ".tsv"):
        return [[conv(c) for c in row] for row in rt.read_csv_rows(path)]
    text = rt.read_text_file(path)
    # word/ppt와 같은 마크다운 해석: 첫 번째 표만 (코드 블록 안의 '|' 줄, 두 번째 표 제외), 정렬 줄 제외, '\|'는 칸 안의 글자 '|'
    table = next((payload for kind, payload in md.parse_markdown_blocks(text) if kind == "table"), None)
    if table:
        return [[conv(c) for c in row] for row in table]
    return [[ln.rstrip()] for ln in text.splitlines() if ln.strip()]


def parse_values_arg(raw: str) -> List[List[Any]]:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise rt.InputError(f"--values는 JSON 2차원 배열이어야 합니다 (예: [[\"이름\",\"값\"],[\"가\",1]]): {e}")
    if not isinstance(data, list) or not data:
        raise rt.InputError("--values는 비어 있지 않은 JSON 배열이어야 합니다.")
    return [list(r) if isinstance(r, list) else [r] for r in data]


# ==============================================================================
# 2. 대상 열기
# ==============================================================================

def is_closed_xlsx(target: str) -> Optional[str]:
    """--target이 열려 있지 않은 xlsx/xlsm 파일 경로면 그 경로 (openpyxl로 읽기), 아니면 None."""
    if not target or target.strip().lower() == "active":
        return None
    raw = os.path.abspath(os.path.expandvars(target.strip().strip('"')))
    if not os.path.isfile(raw):
        cand = os.path.join(rt.get_data_dir(), target.strip().strip('"'))
        if not os.path.isfile(cand):
            return None
        raw = cand
    if os.path.splitext(raw)[1].lower() not in (".xlsx", ".xlsm") or not zipfile.is_zipfile(raw):
        return None   # .xls, DRM 암호화 파일 등은 Excel(COM)로 읽음
    try:
        if off.win32com is not None and off.match_open("excel", raw) is not None:
            return None   # 열려 있으면 화면의 최신(미저장 포함) 내용을 읽음
    except rt.InputError:
        return None   # 같은 이름의 열린 문서가 여럿: 파일을 읽지 않고 COM 경로에서 [후보] 오류로 안내
    except Exception:
        pass
    return raw


def closed_csv(target: str) -> Optional[str]:
    """--target이 CSV/TSV 파일이면 그 경로 (Excel 없이 글자로 읽음), 아니면 None."""
    if not target or target.strip().lower() == "active":
        return None
    raw = os.path.abspath(os.path.expandvars(target.strip().strip('"')))
    if not os.path.isfile(raw):
        cand = os.path.join(rt.get_data_dir(), target.strip().strip('"'))
        raw = cand if os.path.isfile(cand) else ""
    return raw if raw and os.path.splitext(raw)[1].lower() in (".csv", ".tsv") else None


def _read_matrix_csv(path: str, address: Optional[str]) -> Tuple[str, int, int, List[List[Any]], List[str]]:
    """CSV를 엑셀 시트처럼 읽습니다 (인코딩 자동 판별, 숫자는 숫자로). --range가 있으면 그 범위만."""
    rows = [[coerce(v) for v in r] for r in rt.read_csv_rows(path)]
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    name = os.path.splitext(os.path.basename(path))[0]
    r1, c1 = 1, 1
    if address:
        width = max((len(r) for r in rows), default=0)
        r1, c1, r2, c2 = parse_area(split_sheet(address)[1], len(rows), width)
        rows = [(r + [None] * (width - len(r)))[c1 - 1:c2] for r in rows[r1 - 1:r2]]
    return name, r1, c1, rows, [name]


def write_session(args, command: str):
    """쓰기 명령의 세션. --dry-run이면 읽기 세션 (저장·새 파일·되돌리기 지점·화면 이동 없음).
    없는 파일을 새로 만드는 것은 write뿐 (style·replace·range·sheet·chart는 경로 오타면 빈 새 파일을 만들고 VERIFIED였음, 2026-11-09 검토 B8)."""
    dry = bool(getattr(args, "dry_run", False))
    return off.open_target("excel", args.target, readonly=dry, focus=not args.no_focus,
                           create=not dry and command == "write", tool=TOOL, command=command)


def target_range(t: off.OfficeTarget, ws: Any, address: Optional[str]) -> Any:
    """주소 해석: 'A1:B2', 'Sheet!A1', 생략 시 UsedRange."""
    if not address:
        return ws.UsedRange
    sheet, cells = split_sheet(address)
    if sheet:
        ws = t.sheet(sheet)
    try:
        return ws.Range(cells)
    except Exception:
        raise rt.InputError(f"범위 주소가 잘못되었습니다: '{address}' (예: A1, A1:D10, B:B, 'Sheet1'!A1)")


def range_info(rng: Any) -> Dict[str, Any]:
    return {"sheet": str(rng.Worksheet.Name), "range": str(rng.Address).replace("$", ""),
            "rows": int(rng.Rows.Count), "cols": int(rng.Columns.Count)}


# ==============================================================================
# 3. read / info
# ==============================================================================

def _closed_bounds(ws: Any, path: Optional[str] = None) -> Tuple[int, int, int, int]:
    """닫힌 시트의 데이터 범위 (첫 행, 첫 열, 끝 행, 끝 열). 빈 시트는 (1, 1, 0, 0).
    크기 정보(<dimension>)가 없는 파일(다른 프로그램이 만든 xlsx)은 칸을 훑어 실제 범위를 셉니다
    (B:B가 B:D로, 5:7이 끝 열 없이 읽히지 않게). 크기 정보가 실제보다 작은 파일(A1:B5인데 11행까지 있음)은 시트 XML에서
    내용이 있는 마지막 칸까지 넓힙니다 (전에는 앞 5행만 읽고 VERIFIED, 2026-11-09 검토 A3)."""
    if ws.max_row is not None and ws.max_column is not None:
        r1, c1, r2, c2 = int(ws.min_row), int(ws.min_column), int(ws.max_row), int(ws.max_column)
        if path and path.lower().endswith((".xlsx", ".xlsm")):
            ext = xf.xlsx_value_extent(path, ws.title)
            if ext and (ext[0] > r2 or ext[1] > c2):
                rt.debug(f"시트 크기 정보가 실제보다 작음: {r2}행 {c2}열 -> {ext[0]}행 {ext[1]}열")
                r2, c2 = max(r2, ext[0]), max(c2, ext[1])
                if r1 > r2 or c1 > c2 or r1 < 1:
                    r1, c1 = 1, 1
        if (r2 - r1 + 1) * (c2 - c1 + 1) > CLOSED_MAX_CELLS:
            # 떠돌이 먼 칸(XFD1048576의 공백 등)으로 크기가 커진 파일: 값이 있는 칸만으로 다시 셈 (전에는 MemoryError, 검토 B7)
            return _scan_bounds(ws)
        return r1, c1, r2, c2
    return _scan_bounds(ws)


def _scan_bounds(ws: Any) -> Tuple[int, int, int, int]:
    """칸을 훑어 값이 있는(공백만인 글 제외) 범위. 빈 시트는 (1, 1, 0, 0). 읽기 전용 시트는 크기 정보를 지우고 행마다 있는 칸만 훑음."""
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()
    r1 = c1 = None
    r2 = c2 = 0
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if not row:
            continue
        cols = [j for j, v in enumerate(row, start=1) if v is not None and not (isinstance(v, str) and not v.strip())]
        if not cols:
            continue
        r1 = i if r1 is None else r1
        r2 = i
        c1 = cols[0] if c1 is None else min(c1, cols[0])
        c2 = max(c2, cols[-1])
    if r1 is None or c1 is None:
        return 1, 1, 0, 0
    return r1, c1, r2, c2


CLOSED_MAX_CELLS = 2_000_000   # 닫힌 xlsx를 준 범위만큼 채워 읽을 때 칸 수 상한 (A1:XFD1048576 -> 메모리 부족 방지)


def _last_value_cell(ws: Any, r1: int, c1: int, r2: int, c2: int) -> Tuple[int, int]:
    """범위 안에서 값이 있는 마지막 행과 마지막 열 (값만 봄: 서식만 있는 칸은 빈 칸). 없으면 (r1, c1).
    읽기 전용 openpyxl은 파일에 없는 빈 행을 같은 튜플로 주므로 한 번 빈 것으로 본 튜플은 다시 훑지 않습니다."""
    last_r, last_c = r1, c1
    seen_empty: List[Any] = []
    for i, row in enumerate(ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True), start=r1):
        if any(row is e for e in seen_empty):
            continue
        cols = [j for j, v in enumerate(row, start=c1) if v is not None and v != ""]
        if not cols:
            if len(seen_empty) < 4:
                seen_empty.append(row)
            continue
        last_r, last_c = i, max(last_c, cols[-1])
    return last_r, last_c


_SHEET_SCANS: Dict[Tuple[str, str], Dict[str, Any]] = {}   # (경로, 시트) -> 이번 읽기에서 훑은 시트 전체의 계산되지 않은 수식 칸


FAR_GAP = 1000   # 값이 있는 끝에서 이보다 멀리(행 또는 열) 떨어진 계산된 값이 없는 수식 칸은 읽는 범위에 넣지 않음
_FAR: Dict[Tuple[str, str], Dict[str, Any]] = {}   # (경로, 시트) -> 범위에 넣지 않은 먼 칸 {"positions", "data_end"}


def _box(cells: List[Tuple[int, int]]) -> Tuple[int, int, int, int]:
    return (min(r for r, _ in cells), min(c for _, c in cells), max(r for r, _ in cells), max(c for _, c in cells))


def _uncomputed_extent(path: str, title: str) -> Optional[Tuple[int, int, int, int]]:
    """시트 전체에서 계산된 값이 없는 수식 칸을 감싸는 범위 (없으면 None). 데이터로 범위를 정할 때(범위 생략, B:B, 큰 범위 줄이기)
    끝쪽 행·열을 자르기 전에 확인합니다: 그런 칸은 빈 칸(None)으로 읽혀 합계 행 같은 수식 행·열이 잘려 나가기 때문.
    값이 있는 끝에서 FAR_GAP보다 멀리 떨어진 칸(데이터 A1:C10과 Z200000)은 넣지 않고 _FAR에 남깁니다: 넣으면 그 칸까지
    빈 행을 채워 읽고, 큰 범위의 [추천]이 그 칸을 빼는 범위가 되기 때문."""
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return None
    info = _SHEET_SCANS.get((path, title))
    if info is None:
        info = _SHEET_SCANS[(path, title)] = xf.xlsx_uncomputed_formulas(path, title, positions=True)
    b = info.get("bounds")
    if not b:
        return None
    vb = info.get("value_bounds")
    pos = info.get("positions") or []
    if vb and pos and not info.get("positions_truncated"):
        lim_r, lim_c = vb[2] + FAR_GAP, vb[3] + FAR_GAP
        far = [p for p in pos if p[0] > lim_r or p[1] > lim_c]
        if far:
            near = [p for p in pos if p[0] <= lim_r and p[1] <= lim_c]
            _FAR[(path, title)] = {"positions": far, "data_end": (vb[2], vb[3])}
            return _box(near) if near else None
    return (b[0], b[1], b[2], b[3])


def _data_end(path: str, ws: Any) -> Tuple[int, int]:
    """닫힌 시트의 데이터 끝 (행, 열): 파일에 적힌 크기(서식만 있는 칸 포함)에 가까운 계산되지 않은 수식 칸까지.
    먼 칸이 있으면(파일 크기가 그 칸까지라) 값이 있는 끝과 가까운 수식 칸까지만."""
    _, _, mr2, mc2 = _closed_bounds(ws, path)
    ext = _uncomputed_extent(path, ws.title)
    far = _FAR.get((path, ws.title))
    if far:
        mr2, mc2 = far["data_end"]
    if ext:
        mr2, mc2 = max(mr2, ext[2]), max(mc2, ext[3])
    return mr2, mc2


def far_uncomputed(path: str, sheet: str, asked: Tuple[int, int, int, int],
                   read: Tuple[int, int, int, int]) -> Optional[Dict[str, Any]]:
    """읽은 범위에 넣지 않은 먼 계산되지 않은 수식 칸 중 요청한 범위(asked) 안이고 읽은 범위(read) 밖인 칸. 없으면 None.
    {"count", "cells"(앞 rt.UNCOMPUTED_SHOW개, 시트!주소), "box"(감싸는 주소)}."""
    far = _FAR.get((path, sheet))
    if not far:
        return None
    hits = [p for p in far["positions"]
            if asked[0] <= p[0] <= asked[2] and asked[1] <= p[1] <= asked[3]
            and not (read[0] <= p[0] <= read[2] and read[1] <= p[1] <= read[3])]
    if not hits:
        return None
    hits.sort()
    q = f"'{sheet}'" if re.search(r"[\s!']", sheet) else sheet
    b = _box(hits)
    return {"count": len(hits), "cells": [f"{q}!{addr(r, c)}" for r, c in hits[:xf.UNCOMPUTED_SHOW]],
            "box": f"{q}!{addr(b[0], b[1], b[2], b[3])}"}


def _clip_closed_range(path: str, ws: Any, address: str, r1: int, c1: int, r2: int, c2: int) -> Tuple[int, int]:
    """준 범위가 CLOSED_MAX_CELLS칸을 넘으면 데이터가 있는 끝(행·열)까지로 줄입니다 (빈 칸만 빠져 값은 같음).
    파일에 적힌 크기(<dimension>)가 서식만 있는 먼 칸(Z300000 등)까지라 그것으로도 넘으면 값을 훑어 실제 끝을 찾습니다
    (기록된 크기만 보고 멈추지 않음). 줄여도 넘으면 입력 오류와 [추천] 범위를 냅니다. 줄인 (끝 행, 끝 열)을 돌려줍니다."""
    mr2, mc2 = _data_end(path, ws)    # 계산된 값이 없는 수식 칸(빈 칸으로 읽힘)까지는 자르지 않음 (먼 칸 제외)
    ext = _uncomputed_extent(path, ws.title)
    nr2, nc2 = min(r2, max(mr2, r1)), min(c2, max(mc2, c1))
    if (nr2 - r1 + 1) * (nc2 - c1 + 1) > CLOSED_MAX_CELLS:
        nr2, nc2 = _last_value_cell(ws, r1, c1, nr2, nc2)
        if ext and ext[0] <= r2 and ext[2] >= r1 and ext[1] <= c2 and ext[3] >= c1:
            nr2, nc2 = max(nr2, min(ext[2], r2)), max(nc2, min(ext[3], c2))
    width = nc2 - c1 + 1
    if (nr2 - r1 + 1) * width > CLOSED_MAX_CELLS:
        rec = addr(r1, c1, r1 + CLOSED_MAX_CELLS // width - 1, nc2)
        raise rt.InputError(f"범위 '{address}'가 너무 큽니다 (데이터가 있는 부분만도 {nr2 - r1 + 1}행 x {width}열, "
                            f"한 번에 {CLOSED_MAX_CELLS:,}칸까지). 범위를 나눠 읽으세요. [추천] {rt.retry_command({'--range': rec})}")
    rt.note_recovered(f"범위 '{address}'가 {CLOSED_MAX_CELLS:,}칸을 넘음 -> 데이터가 있는 {addr(r1, c1, nr2, nc2)}까지만 읽음")
    return nr2, nc2


def worksheet_names(wb: Any) -> List[str]:
    """openpyxl 통합문서의 워크시트 이름 (차트 시트 제외)."""
    return [ws.title for ws in wb.worksheets]


def chart_sheet_note(wb: Any, name: str) -> str:
    """준 이름이 차트 시트면 그 설명 (not_found의 extra)."""
    charts = [c.title for c in getattr(wb, "chartsheets", [])]
    hit = rt.find_name(name, charts) if charts else None
    return f" ('{hit}'은(는) 차트 시트라 칸이 없음, 워크시트를 고르세요)" if hit else ""


TIME_CUT: Dict[str, Any] = {}   # 닫힌 파일 읽기가 명령 시간 제한 전에 멈춘 곳 (cmd_read가 UNVERIFIED로 알림, partial_ok: read만)


def _read_matrix_closed(path: str, sheet: Optional[str], address: Optional[str]) -> Tuple[str, int, int, List[List[Any]], List[str]]:
    """닫힌 파일을 openpyxl로 읽습니다. (시트명, 시작행, 시작열, 값, 전체 시트명)"""
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "Office 없이 닫힌 엑셀 파일 읽기")
    _SHEET_SCANS.clear()      # 이번 읽기에서 훑은 것만 closed_uncomputed가 다시 씀
    _FAR.clear()
    wb = _load_closed(openpyxl, path)
    try:
        return _read_matrix_open_wb(wb, path, sheet, address)
    except (KeyError, SyntaxError, zipfile.BadZipFile) as e:     # ParseError는 SyntaxError (깨진 시트 XML, 2026-11-09 검토 B13)
        raise rt.InputError(f"{os.path.basename(path)}을(를) 읽을 수 없습니다: 손상되었거나 엑셀 파일 구조가 아닙니다 ({type(e).__name__}). "
                            "엑셀에서 열어 다시 저장한 뒤 시도하세요.") from None
    finally:
        wb.close()


def _load_closed(openpyxl: Any, path: str) -> Any:
    """닫힌 xlsx 열기. xlsx가 아닌 zip·깨진 파일은 내부 오류가 아니라 STOP_RETRY."""
    try:
        return rt.retry_fs(lambda: openpyxl.load_workbook(path, read_only=True, data_only=True))
    except (KeyError, SyntaxError, zipfile.BadZipFile, ValueError, TypeError) as e:
        raise rt.InputError(f"{os.path.basename(path)}을(를) 열 수 없습니다: 엑셀 파일(.xlsx)이 아니거나 손상되었습니다 ({type(e).__name__}).") from None


def _read_matrix_open_wb(wb: Any, path: str, sheet: Optional[str], address: Optional[str]) -> Tuple[str, int, int, List[List[Any]], List[str]]:
    whole = False
    ext: Optional[Tuple[int, int, int, int]] = None
    names = worksheet_names(wb)            # 차트 시트는 칸이 없어 뺌 (2026-11-09 검토 B6: 'Chartsheet'에 max_row가 없어 ERROR)
    if not names:
        raise rt.InputError(f"{os.path.basename(path)}에 칸이 있는 워크시트가 없습니다 (차트 시트만 있음).")
    if sheet:
        found = rt.find_name(sheet, names)
        if found is None:
            raise rt.not_found("시트", sheet, names, option="--sheet", extra=chart_sheet_note(wb, sheet))
        ws = wb[found]
    else:
        act = wb.active
        ws = act if act is not None and getattr(act, "title", None) in names else wb[names[0]]
    # 데이터 범위(_closed_bounds, 크기 정보 없는 파일은 시트 전체를 훑음)는 범위를 생략했거나
    # 열·행 전체(B:B, 5:7)일 때만, 그리고 Sheet!A1의 시트로 바꾼 뒤에 셉니다 (작은 범위 읽기마다 전체를 훑지 않게)
    if address:
        sh, cells = split_sheet(address)
        if sh:
            found = rt.find_name(sh, names)
            if found is None:
                raise rt.not_found("시트", sh, names, option="--sheet", extra=chart_sheet_note(wb, sh))
            ws = wb[found]
        r1, c1, r2, c2 = parse_area(cells, None, None)
        if r2 is None or c2 is None:
            whole = True
            mr2, mc2 = _data_end(path, ws)   # 계산된 값이 없는 수식 칸(빈 칸으로 읽힘)의 행·열도 데이터 범위에 넣음
            r1, c1, r2, c2 = parse_area(cells, mr2, mc2)
            ext = _uncomputed_extent(path, ws.title)
    else:
        # 열린 문서의 UsedRange처럼 데이터가 시작하는 칸부터 (A1부터 읽으면 rows가 열린 문서와 달랐음)
        r1, c1, r2, c2 = _closed_bounds(ws, path)
        ext = _uncomputed_extent(path, ws.title)
        if (path, ws.title) in _FAR:          # 먼 칸 때문에 커진 파일 크기 대신 값이 있는 끝까지
            r2, c2 = _FAR[(path, ws.title)]["data_end"]
        if ext:   # 계산된 값이 없는 수식 칸은 빈 칸으로 읽혀도 자르지 않고 남김 (합계 행 등이 있다는 것을 보이게)
            r1, c1, r2, c2 = (ext if r2 < r1 or c2 < c1 else
                              (min(r1, ext[0]), min(c1, ext[1]), max(r2, ext[2]), max(c2, ext[3])))
    if address and r2 >= r1 and c2 >= c1 and (r2 - r1 + 1) * (c2 - c1 + 1) > CLOSED_MAX_CELLS:
        r2, c2 = _clip_closed_range(path, ws, address, r1, c1, r2, c2)
    if r2 < r1 or c2 < c1:
        rows = []
    else:
        rows = []
        for k, r in enumerate(ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True)):
            if k % 2000 == 1999 and rt.time_up():
                # 명령 시간 제한(120초) 전에 멈춤: JSON 없이 끊기지 않게 (2026-11-09 검토 5장 5번). read는 읽은 데까지 UNVERIFIED,
                # 결과를 그대로 쓰는 명령(export, check, transfer)은 일부만 쓰면 안 되므로 STOP_RETRY
                if not TIME_CUT.get("partial_ok"):
                    raise rt.InputError(f"파일이 커서 명령 시간 제한 안에 다 읽지 못했습니다 ({k}행까지 읽음, {addr(r1, c1, r2, c2)} 중). "
                                        "--range로 범위를 나눠 다시 실행하세요 (예: 앞 절반, 뒤 절반).")
                TIME_CUT["rows"] = k
                TIME_CUT["where"] = addr(r1, c1, r1 + k - 1, c2)
                break
            rows.append(list(r))
        width = c2 - c1 + 1
        rows = [r + [None] * (width - len(r)) for r in rows]
        if address:
            # 읽기 전용 openpyxl은 데이터 끝 뒤의 빈 행을 주지 않음: 준 범위(A1:C5)만큼 채워 열린 문서와 같은 rows로
            rows += [[None] * width for _ in range(r2 - r1 + 1 - len(rows))]
    if not address or whole:   # 끝쪽 빈 행 제거 (계산된 값이 없는 수식 칸이 있는 행까지는 남김). B:B도 (서식만 있는 끝 행까지 세던 것, 검토 B13)
        keep = ext[2] if ext else 0
        while rows and r1 + len(rows) - 1 > keep and all(v is None for v in rows[-1]):
            rows.pop()
    return ws.title, r1, c1, rows, names


def closed_uncomputed(path: str, sheet: str, r0: int, c0: int, matrix: List[List[Any]]) -> Optional[Dict[str, Any]]:
    """닫힌 xlsx에서 읽은 범위 안의 수식 칸 중 계산된 값이 없는 것. 없으면 None.
    빈 칸(None)이 있거나 파일에 fullCalcOnLoad 표시가 있을 때만(그때는 수식 칸의 값(0 등)이 계산된 값이 아님) 확인합니다.
    확인하지 못하면 unknown 정보(UNVERIFIED)."""
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return None
    if not any(v is None for row in matrix for v in row) and not xf.xlsx_full_calc_on_load(path):
        return None
    width = max((len(r) for r in matrix), default=0)
    area = (r0, c0, r0 + len(matrix) - 1, c0 + max(width, 1) - 1)
    info = _SHEET_SCANS.get((path, sheet))     # 범위를 정할 때 시트 전체를 이미 훑었고 그 칸이 모두 읽은 범위 안이면 다시 쓰기
    b = (info or {}).get("bounds")
    if info is None or info.get("unknown") or (b and not (area[0] <= b[0] and area[1] <= b[1] and b[2] <= area[2] and b[3] <= area[3])):
        info = xf.xlsx_uncomputed_formulas(path, sheet, area)
    return info if xf.uncomputed_found(info) else None


def _search_matrix(matrix: List[List[Any]], r0: int, c0: int, text: str, limit: int = 50,
                   formulas: Optional[Dict[Tuple[int, int], str]] = None) -> Tuple[List[Dict[str, Any]], int, int]:
    """찾은 칸 (앞 limit개), 전체 수, 값이 잘린 칸 수. 값은 칸마다 MAX_CELL_CHARS에서 자릅니다 (전체는 --save CSV).
    formulas((행 위치, 열 위치) -> 수식, --formulas)를 주면 수식 글에서도 찾고 수식 칸은 formula를 붙입니다."""
    hits: List[Dict[str, Any]] = []
    total = cut_n = 0
    key = text.lower()
    formulas = formulas or {}
    # 표시 형식 숫자('1,000', '25%', '1500원')는 숫자 칸의 값(1000, 0.25)과도 맞춤 (원시 값만 찾아 0건이던 것, 2026-11-09 검토 B13)
    want = _where_num(text) if re.search(r"[,%]|\d\s*[^\W\d_]+$", text.strip()) else None
    for i, row in enumerate(matrix):
        for j, v in enumerate(row):
            f = formulas.get((i, j))
            if v is None and f is None:
                continue
            num_hit = want is not None and _is_num(v) and abs(float(v) - want) <= 1e-9 * max(1.0, abs(want))
            if num_hit or key in str(out_value(v)).lower() or (f is not None and key in f.lower()):
                total += 1
                if len(hits) < limit:
                    value, cut = rt.clamp_cell(out_value(v))
                    hit = {"cell": addr(r0 + i, c0 + j), "value": value}
                    if f is not None:
                        hit["formula"], f_cut = rt.clamp_cell(f)
                        cut = cut or f_cut
                    if cut:
                        hit["value_truncated"] = True
                        cut_n += 1
                    hits.append(hit)
    return hits, total, cut_n


# --- read --formulas: 수식 칸 목록 ----------------------------------------------------------------------

FORMULA_SHOW = 200          # JSON formulas에 담는 수식 칸 수 (넘으면 formulas_truncated, 전체는 full_result_file CSV)
ARRAY_CHECK_MAX = 2000      # 열린 문서에서 배열 수식인지 칸마다 확인하는 수식 칸 수 상한 (COM 호출 수 제한)


def _formula_value(v: Any) -> Any:
    """수식 칸의 값: 계산된 값이 없으면(닫힌 파일의 계산되지 않은 수식) null."""
    return None if v is None else out_value(v)


def _closed_formulas(path: str, sheet: str, r0: int, c0: int, matrix: List[List[Any]]) -> List[Dict[str, Any]]:
    """닫힌 xlsx/xlsm에서 읽은 범위의 수식 칸 (openpyxl data_only=False로 그 범위만 한 번 더 읽음).
    값은 이미 읽은 계산된 값(data_only, 없으면 null). 공유 수식은 openpyxl이 칸마다 펼쳐 줍니다.
    배열 수식은 파일에 왼쪽 위 칸에만 있어 그 칸에 한 번(array, array_range). 수식 칸은 data_type 'f'만
    ('=로 시작하는 글자 상수는 수식이 아님)."""
    nrows = len(matrix)
    ncols = max((len(r) for r in matrix), default=0)
    if not nrows or not ncols:
        return []
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "Office 없이 닫힌 엑셀 파일의 수식 읽기")
    from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
    out: List[Dict[str, Any]] = []
    wb = rt.retry_fs(lambda: openpyxl.load_workbook(path, read_only=True, data_only=False))
    try:
        ws = wb[sheet]
        for i, row in enumerate(ws.iter_rows(min_row=r0, max_row=r0 + nrows - 1, min_col=c0, max_col=c0 + ncols - 1)):
            if i % 2000 == 1999 and rt.time_up():
                TIME_CUT["formulas"] = i                   # 수식을 앞 i행에서만 찾음 (시간 제한)
                break
            for j, cell in enumerate(row):
                v = cell.value
                extra: Dict[str, Any] = {}
                if isinstance(v, ArrayFormula):
                    text = str(v.text or "")
                    extra = {"array": True, "array_range": str(v.ref or "")}
                elif isinstance(v, DataTableFormula):     # 데이터 표(What-If): 엑셀 화면의 {=TABLE(행 입력, 열 입력)}
                    text = f"=TABLE({v.r1 or ''},{v.r2 or ''})"
                    extra = {"array": True, "array_range": str(v.ref or ""), "data_table": True}
                elif isinstance(v, str) and v.startswith("=") and getattr(cell, "data_type", None) == "f":
                    text = v
                else:
                    continue
                if not text.startswith("="):
                    text = "=" + text
                raw = matrix[i][j] if i < len(matrix) and j < len(matrix[i]) else None
                out.append({"cell": addr(r0 + i, c0 + j), "formula": text, "value": _formula_value(raw), **extra})
    finally:
        wb.close()
    return out


def _com_formulas(rng: Any, matrix: List[List[Any]], warnings: List[str]) -> List[Dict[str, Any]]:
    """열린 문서(COM)의 수식 칸. HasFormula가 False(수식 없음)면 수식을 읽지 않습니다.
    수식 글은 범위 전체를 한 번에: Formula2(동적 배열을 그대로, 엑셀 365)가 없으면 Formula (영어 함수 이름, A1 주소).
    2차원 결과에서 상수는 값, 수식은 '='로 시작하는 글이라 칸마다 부르지 않습니다. 다만 값도 같은 글('=로 시작하는
    글자 상수일 수 있음)이면 그 칸만 HasFormula로 확인합니다. 배열 수식은 범위의 HasArray가 False가 아닐 때만
    칸마다 확인해(ARRAY_CHECK_MAX개까지) 배열마다 한 번(배열 안에서 처음 만난 칸, 보통 왼쪽 위) 넣습니다."""
    try:
        has = rng.HasFormula            # True(모두) / False(없음) / None(섞임)
    except Exception:
        has = None
    if has is False:
        return []
    try:
        raw = rng.Formula2
    except Exception:
        raw = rng.Formula
    fm = off.range_to_matrix(raw)
    r0, c0 = int(rng.Row), int(rng.Column)
    ws = rng.Worksheet
    cands = [(i, j, f) for i, row in enumerate(fm) for j, f in enumerate(row) if isinstance(f, str) and f.startswith("=")]
    if not cands:
        return []
    try:
        arrays = rng.HasArray           # False면 배열 수식 없음 (칸마다 확인하지 않음)
    except Exception:
        arrays = None
    out: List[Dict[str, Any]] = []
    seen_arrays = set()
    checked = 0
    for i, j, f in cands:
        v = matrix[i][j] if i < len(matrix) and j < len(matrix[i]) else None
        cell = None
        if isinstance(v, str) and v == f:
            cell = ws.Cells(r0 + i, c0 + j)
            if not cell.HasFormula:
                continue
        entry: Dict[str, Any] = {"cell": addr(r0 + i, c0 + j), "formula": f, "value": _formula_value(v)}
        if arrays is not False:
            if checked < ARRAY_CHECK_MAX:
                checked += 1
                cell = cell or ws.Cells(r0 + i, c0 + j)
                if cell.HasArray:
                    area = str(cell.CurrentArray.Address).replace("$", "")
                    if area in seen_arrays:
                        continue
                    seen_arrays.add(area)
                    entry.update(array=True, array_range=area)
            elif checked == ARRAY_CHECK_MAX:
                checked += 1
                warnings.append(f"배열 수식인지는 앞 {ARRAY_CHECK_MAX}개 수식 칸만 확인함 (그 뒤 배열 수식은 칸마다 나올 수 있음)")
        out.append(entry)
    return out


def _formula_csv(path: str, entries: List[Dict[str, Any]]) -> str:
    rows = [["cell", "formula", "value", "array"]] + [
        [e["cell"], e["formula"], "" if e["value"] is None else e["value"], e.get("array_range", "") if e.get("array") else ""]
        for e in entries]
    return rt.write_csv_rows(path, rows)


_DEFAULT_DIRECTIVE = "VERIFIED: 작업이 완료되었습니다."


def _with_formulas(payload: Dict[str, Any], info: Dict[str, Any]) -> Dict[str, Any]:
    """read 결과에 수식 칸 목록을 붙입니다 (--formulas). formulas(앞 FORMULA_SHOW개), formula_count, formulas_truncated.
    잘렸거나 --save면 전체 목록 CSV(full_result_file)를 쓰고, --save면 다시 읽어 행 수를 확인합니다 (R1)."""
    if payload.get("status") != "ok":
        return payload
    entries: List[Dict[str, Any]] = info["entries"]
    n = len(entries)
    shown = []
    for e in entries[:FORMULA_SHOW]:
        e = dict(e)
        e["formula"], f_cut = rt.clamp_cell(e["formula"])
        e["value"], v_cut = rt.clamp_cell(e["value"])
        if f_cut or v_cut:
            e["value_truncated"] = True
        shown.append(e)
    fields = {k: v for k, v in payload.items() if k not in xf._PAYLOAD_BASE}
    fields.update(formulas=shown, formula_count=n, formulas_truncated=n > FORMULA_SHOW)
    if info.get("warnings"):
        fields["warnings"] = list(fields.get("warnings") or []) + info["warnings"]
    problems: List[str] = []
    full = None
    if n and payload.get("output_file"):      # --save: 저장한 값 CSV 옆에 전체 수식 목록
        base = os.path.splitext(payload["output_file"])[0]
        full = _formula_csv(rt.resolve_output_path(base + "_formulas.csv", "formulas", ".csv"), entries)
        problem = rt.check_saved_file(full, expected_rows=n + 1, delimiter=",")
        if problem:
            problems.append(f"수식 목록 CSV: {problem}")
    elif n > FORMULA_SHOW:
        out = os.path.join(rt.get_data_dir(), "results",
                           f"excel_formulas_{rt.safe_filename(payload.get('sheet'), 40)}_{rt.timestamp()}.csv")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        full = _formula_csv(out, entries)
    if full:
        fields["full_result_file"] = full
    where = f"{payload.get('sheet')}!{payload.get('range')}"
    if info.get("csv"):
        sentence = "수식 0개: CSV 파일은 수식을 저장하지 않아 값만 있습니다. 사용자에게 그대로 알리세요."
        preview_tail = " / 수식 없음 (CSV)"
    elif not n:
        sentence = f"수식 0개: {where}에는 수식이 없습니다(값만 있음). 사용자에게 그대로 알리세요."
        preview_tail = " / 수식 없음"
    else:
        sentence = f"수식 칸 {n}개는 formulas(cell, formula, value)에 있습니다."
        if n > FORMULA_SHOW:
            sentence += f" 앞 {FORMULA_SHOW}개만 담았고 전체는 full_result_file에 있습니다."
        elif full:
            sentence += " 전체 수식 목록은 full_result_file에도 저장했습니다."
        preview_tail = f" / 수식 {n}개"
    preview = str(payload.get("preview", "")) + preview_tail
    if payload.get("verified") is False or problems:
        return rt.ok_unverified(payload["tool"], payload["command"], preview,
                                list(payload.get("unverified") or []) + problems, **fields)
    directive = str(payload.get("agent_directive", ""))
    if directive.startswith(_DEFAULT_DIRECTIVE):
        directive = "VERIFIED: 읽기 완료. " + sentence + " 같은 명령을 다시 실행하거나 결과를 다시 조회하지 마세요."
    else:
        directive = directive + " " + sentence
    return rt.ok(payload["tool"], payload["command"], preview, directive, **fields)


def _no_formula_reader(target: str) -> Optional[str]:
    """Office(COM) 없이 --formulas를 읽을 수 없는 닫힌 파일(.xls, .xlsb, zip이 아닌 xlsx 등)이면 그 경로, 아니면 None."""
    if off.win32com is not None or not target or target.strip().lower() == "active":
        return None
    raw = os.path.abspath(os.path.expandvars(target.strip().strip('"')))
    if os.path.isfile(raw) and os.path.splitext(raw)[1].lower() in off.EXTENSIONS["excel"]:
        return raw
    return None


def cmd_read(args) -> Dict[str, Any]:
    """read. 닫힌 큰 파일을 읽다 시간 제한 전에 멈췄으면(TIME_CUT) 결과를 UNVERIFIED로 바꾸고 나눠 읽는 방법을 알림."""
    TIME_CUT.clear()
    TIME_CUT["partial_ok"] = True
    try:
        res = _cmd_read(args)
    finally:
        TIME_CUT.pop("partial_ok", None)
    if not TIME_CUT or res.get("status") != "ok":
        return res
    why = []
    if "rows" in TIME_CUT:
        why.append(f"파일이 커서 명령 시간 제한 전에 멈춰 앞 {TIME_CUT['rows']}행({TIME_CUT['where']})만 읽음: 결과(건수·찾은 칸)는 그 범위 안의 것")
    if "formulas" in TIME_CUT:
        why.append(f"수식 칸은 앞 {TIME_CUT['formulas']}행에서만 찾음 (시간 제한)")
    why.append("--range로 범위를 나눠 다시 읽거나 엑셀에서 열어 읽으세요")
    fields = {k: v for k, v in res.items() if k not in ("status", "tool", "command", "preview", "verified", "unverified",
                                                        "agent_directive", "recovered")}
    fields["time_limited"] = dict(TIME_CUT)
    return rt.ok_unverified(TOOL, "read", res.get("preview", "") + " (시간 제한으로 일부만)", why + list(res.get("unverified") or []), **fields)


def _cmd_read(args) -> Dict[str, Any]:
    found: Dict[str, Any] = {}
    payload = _read(args, found)
    if args.formulas:
        payload = _with_formulas(payload, found["formulas"])
    return _with_far(xf.with_uncomputed(payload, found.get("uncomputed")), found.get("far"), args.target)


def _far_for_read(path: str, sheet: str, address: Optional[str], r0: int, c0: int,
                  matrix: List[List[Any]]) -> Optional[Dict[str, Any]]:
    """읽은 범위에 넣지 않은 먼 계산되지 않은 수식 칸 중 요청한 범위 안의 것 (범위 생략이면 시트 전체)."""
    if (path, sheet) not in _FAR:
        return None
    asked = (1, 1, XL_MAX_ROW, XL_MAX_COL)
    if address:
        r1, c1, r2, c2 = parse_area(split_sheet(address)[1], XL_MAX_ROW, XL_MAX_COL)
        asked = (r1, c1, r2 if r2 is not None else XL_MAX_ROW, c2 if c2 is not None else XL_MAX_COL)
    width = max((len(r) for r in matrix), default=0)
    read = (r0, c0, r0 + len(matrix) - 1, c0 + width - 1) if matrix and width else (0, 0, -1, -1)
    return far_uncomputed(path, sheet, asked, read)


def _with_far(payload: Dict[str, Any], far: Optional[Dict[str, Any]], target: str) -> Dict[str, Any]:
    """R1: 읽지 않은 먼 계산되지 않은 수식 칸이 있으면 UNVERIFIED와 그 칸을 따로 읽는 명령."""
    if not far or payload.get("status") != "ok":
        return payload
    fields = {k: v for k, v in payload.items() if k not in xf._PAYLOAD_BASE}
    cmd = rt.next_command("excel.py", ["read", "--target", target, "--range", far["box"]],
                          "데이터에서 멀리 떨어진 계산된 값이 없는 수식 칸만 따로 읽기")
    fields["next_commands"] = list(fields.get("next_commands") or []) + [cmd]
    fields.update(far_uncomputed_formulas=far["cells"], far_uncomputed_count=far["count"])
    shown = ", ".join(far["cells"][:5]) + (" 등" if far["count"] > 5 else "")
    note = (f"데이터에서 {FAR_GAP}행(열) 넘게 떨어진 계산된 값이 없는 수식 칸 {far['count']}개({shown})는 빈 행을 채우지 않으려고 "
            "읽은 범위에 넣지 않음. 사용자에게 그대로 알리고, 그 칸이 필요하면 next_commands의 '따로 읽기' 명령을 쓰세요")
    return rt.ok_unverified(payload["tool"], payload["command"], payload["preview"],
                            list(payload.get("unverified") or []) + [note], **fields)


def _read_formulas(fn: Any, *a: Any) -> List[Dict[str, Any]]:
    """수식 읽기 실패는 내부 오류로: 예외 글에 수식이 섞여 오류 로그에 남지 않게 예외 종류만 남기고 앞 예외도 잇지 않습니다."""
    try:
        return fn(*a)
    except rt.InputError:
        raise
    except Exception as e:
        raise RuntimeError(f"수식 읽기 실패 ({type(e).__name__})") from None


def _read(args, found: Dict[str, Any]) -> Dict[str, Any]:
    csv_path = closed_csv(args.target)
    closed = None if csv_path else is_closed_xlsx(args.target)
    finfo: Dict[str, Any] = {"entries": [], "warnings": []}
    found["formulas"] = finfo
    conds = None
    if args.where is not None:          # 식 오류는 문서를 읽기 전에
        if args.search or args.formulas:
            raise rt.InputError("--where는 --search, --formulas와 함께 쓸 수 없습니다. 하나만 쓰세요.")
        conds = parse_where(args.where)
    if args.formulas and not csv_path and not closed:
        no_reader = _no_formula_reader(args.target)
        if no_reader:
            raise rt.InputError(f"'{os.path.basename(no_reader)}'의 수식은 Office 없이 읽을 수 없습니다 "
                                "(Office 없이 수식을 읽는 것은 .xlsx/.xlsm만). 엑셀에서 이 파일을 연 뒤 같은 명령을 다시 실행하세요 "
                                "(열린 문서는 엑셀에서 수식을 읽음).")
    if csv_path:
        # CSV도 같은 명령으로 (행 수 세기, 머리글 확인): 확장자 오류를 내면 LLM이 pandas 코드를 직접 짬
        sheet, r0, c0, matrix, names = _read_matrix_csv(csv_path, args.range)
        source = {"app": "Excel", "name": os.path.basename(csv_path), "path": csv_path, "source": "file(csv)"}
        finfo["csv"] = True
    elif closed:
        sheet, r0, c0, matrix, names = _read_matrix_closed(closed, args.sheet, args.range)
        found["uncomputed"] = closed_uncomputed(closed, sheet, r0, c0, matrix)
        found["far"] = _far_for_read(closed, sheet, args.range, r0, c0, matrix)
        if args.formulas:
            finfo["entries"] = _read_formulas(_closed_formulas, closed, sheet, r0, c0, matrix)
        source = {"app": "Excel", "name": os.path.basename(closed), "path": closed, "source": "file(openpyxl)"}
    else:
        with off.open_target("excel", args.target, readonly=True) as t:
            ws = t.sheet(args.sheet)
            rng = target_range(t, ws, args.range)
            matrix = off.range_to_matrix(rng.Value)
            sheet, r0, c0 = str(rng.Worksheet.Name), int(rng.Row), int(rng.Column)
            if args.formulas:
                finfo["entries"] = _read_formulas(_com_formulas, rng, matrix, finfo["warnings"])
            source = t.resolved()
    nrows = len(matrix)
    ncols = max((len(r) for r in matrix), default=0)
    full_addr = addr(r0, c0, r0 + max(nrows, 1) - 1, c0 + max(ncols, 1) - 1)
    result: Dict[str, Any] = {"target_resolved": source, "sheet": sheet, "range": full_addr, "rows": nrows, "cols": ncols}

    save_cmd = rt.next_command("excel.py", ["read", "--target", args.target, "--range", full_addr, "--sheet", sheet, "--save"]
                               + (["--formulas"] if args.formulas else []) + (["--where", args.where] if conds else []),
                               "전체를 CSV로 저장 (잘린 값까지 전체)")
    if conds is not None:
        return _read_where(args, conds, matrix, sheet, r0, c0, full_addr, result, save_cmd)
    if args.search:
        fmap = None
        if args.formulas:     # 수식 글에서도 찾기 (수식 칸 위치 -> 수식)
            fmap = {}
            for e in finfo["entries"]:
                rr, cc = parse_cell(e["cell"])
                fmap[(rr - r0, cc - c0)] = e["formula"]
        hits, total, cut_n = _search_matrix(matrix, r0, c0, args.search, formulas=fmap)
        truncated = total > len(hits) or cut_n > 0
        result.update(matches=hits, total_matches=total, count=total, shown=len(hits), truncated=truncated)
        preview = f"'{args.search}': {sheet}!{full_addr}에서 {total}개 셀 일치" + (
            f" (앞 {len(hits)}개 표시)" if total > len(hits) else "")
        if not total:
            return rt.ok(TOOL, "read", f"0건: {preview}",
                         f"VERIFIED: 0건. {sheet}!{full_addr}에 '{args.search}'이(가) 들어간 칸이 없습니다. "
                         "사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요.", **result)
        if truncated:
            reasons = ([f"찾은 칸 {total}개 중 앞 {len(hits)}개만 표시"] if total > len(hits) else []) + (
                [f"{rt.MAX_CELL_CHARS}자보다 긴 값 {cut_n}개는 앞부분만 (value_truncated)"] if cut_n else [])
            result.update(truncated_reason=", ".join(reasons), next_commands=[save_cmd])
            return rt.ok(TOOL, "read", preview, "VERIFIED: 찾기 완료. 일부만 표시했습니다(truncated_reason). "
                         "전체 값이 필요하면 next_commands[0]으로 CSV에 저장하세요.", **result)
        return rt.ok(TOOL, "read", preview, "", **result)

    values = [[out_value(v) for v in row] for row in matrix]
    if args.save:
        out = rt.resolve_output_path(None if args.save is True else args.save, f"{sheet}_{full_addr.replace(':', '-')}", ".csv")
        out = rt.write_csv_rows(out, values)
        # 열린 문서를 바꾸는 이어 하기는 confirm_commands (1.4.75): next_commands는 묻지 않고 실행해도 되는 읽기·미리보기만
        result.update(output_file=out, confirm_commands=[
            rt.next_command("word.py", ["write", "--file", out], "저장한 표를 열린 워드 문서에 넣기 (사용자가 원할 때만)"),
            rt.next_command("ppt.py", ["write", "--file", out], "저장한 표를 열린 PPT에 넣기 (사용자가 원할 때만)")])
        # R1: 저장한 CSV를 다시 읽어 행 수 비교 (check_saved_file은 빈 행을 빼고 세므로 기대값도 빈 행 제외)
        filled = sum(1 for r in values if any(v is not None and str(v).strip() for v in r))
        return rt.ok_file(TOOL, "read", f"{sheet}!{full_addr} ({nrows}행 x {ncols}열)를 CSV로 저장: {out}", out,
                          "VERIFIED: 저장한 CSV를 다시 열어 행 수를 확인했습니다. " + rt.INTO_OTHER_DOC,
                          expected_rows=filled, delimiter=",", **result)

    if not any(v is not None and str(v).strip() for row in values for v in row):
        # 빈 결과 명시 (ACI 설계 권장 2번, 1.4.81): 빈 배열만 주면 LLM이 표를 지어내거나 범위를 바꿔 다시 읽음
        result.update(values=values[:5], count=0, truncated=False)
        return rt.ok(TOOL, "read", f"0건: {sheet}!{full_addr}에 값이 있는 칸이 없습니다 ({nrows}행 x {ncols}열 모두 빈 칸)",
                     f"VERIFIED: 0건. {sheet}!{full_addr}의 칸이 모두 비어 있습니다 (찾은 항목 없음). 사용자에게 그대로 알리고 "
                     "범위를 바꿔 다시 읽지 마세요.", **result)
    shown, rows_cut, cells_cut, cols_dropped = rt.clamp_rows(values, args.max_rows)
    truncated = rows_cut or cells_cut > 0 or cols_dropped > 0
    result.update(values=shown, truncated=truncated)
    nexts = []
    if rows_cut:
        next_r = r0 + len(shown)
        last_r = r0 + nrows - 1
        rest = f"{sheet}!" if " " not in sheet else f"'{sheet}'!"
        rest += addr(next_r, c0, last_r, c0 + ncols - 1)
        nexts.append(rt.next_command("excel.py", ["read", "--target", args.target, "--range", rest], "나머지 행 이어서 읽기"))
    reasons = []
    if cells_cut:
        reasons.append(f"{rt.MAX_CELL_CHARS}자보다 긴 값 {cells_cut}개는 앞부분만 표시")
    if cols_dropped:
        reasons.append(f"첫 행이 표시 상한을 넘어 뒤쪽 {cols_dropped}열은 표시하지 않음")
    if reasons:
        result["truncated_reason"] = ", ".join(reasons)
    if truncated:
        nexts.append(save_cmd)
        result["next_commands"] = nexts
    preview = f"{sheet}!{full_addr}: {nrows}행 x {ncols}열" + (f" (앞 {len(shown)}행만 표시)" if rows_cut else "") + (
        f" (긴 값 {cells_cut}칸 잘림)" if cells_cut else "") + (f" (뒤쪽 {cols_dropped}열 생략)" if cols_dropped else "")
    return rt.ok(TOOL, "read", preview, "VERIFIED: 읽기 완료. 일부만 표시되었으면 next_commands로 이어서 읽거나 CSV로 저장하세요."
                 if truncated else "", **result)


# --- read --where: 조건에 맞는 행만 ---

_WHERE_HEAD = re.compile(r"\s*([^;|]+?)\s*(>=|<=|!=|!~|>|<|=|~)[ \t]*", re.S)
# 고쳐서 [추천]할 수 있는 표기 (뜻이 하나뿐인 것만, 따옴표 밖에서만)
_WHERE_FIXES = (("==", "="), ("=>", ">="), ("=<", "<="), ("<>", "!="), ("=!", "!="))
_WHERE_AND = re.compile(r"\s+(?:and|AND|And|그리고)\s+|\s*&&\s*")
_WHERE_OR = re.compile(r"\s+(?:or|OR|Or|또는)\s+|\s*\|\|\s*")
_WHERE_CONTAINS = re.compile(r"\s+(?:contains|CONTAINS|포함)\s+")
# 조건을 쉼표로 나눈 것 ('금액 >= 100, 부서 = 영업'): 쉼표 뒤가 '<열> <연산자>'이면 ; 로 (전에는 값 '100, 부서 = 영업' 하나로 읽어 0건)
_WHERE_COMMA = re.compile(r",\s*(?=[^,;|=<>!~'\"]+?\s*(?:>=|<=|!=|=|>|<|~))")
_WHERE_QUOTED = re.compile(r"(\"[^\"]*\"|'[^']*')")
_WHERE_NUM = re.compile(r"^[+-]?(?:\d+\.?\d*|\.\d+)$")
WHERE_FORM = ("'<열> <연산자> <값>' (연산자 >= <= > < = != ~ !~, ~는 포함, !~는 포함하지 않음. 그리고는 ;, 또는은 | (;가 먼저 묶임), "
              "값에 ; | 가 있으면 작은따옴표로 감쌈. 예: 금액 >= 100; 부서 = 영업 | 부서 = 개발)")

Cond = Tuple[str, str, str]


def _where_parts(text: str) -> Optional[List[List[Cond]]]:
    """'<열> <연산자> <값>'을 ;(그리고)와 |(또는)로 이은 식 -> 또는 묶음 [[조건, ...], ...]. 따옴표로 감싼 값 안의 ; | 는 나누지 않음."""
    groups: List[List[Cond]] = [[]]
    pos, n = 0, len(text)
    while pos < n:
        if not text[pos:].strip():
            break
        m = _WHERE_HEAD.match(text, pos)
        if not m or not m.group(1).strip():
            return None
        col, op = m.group(1).strip(), m.group(2)
        pos = m.end()
        q = text[pos] if pos < n and text[pos] in "\"'" else ""
        close = re.compile(re.escape(q) + r"\s*(?=[;|]|$)").search(text, pos + 1) if q else None
        if close:
            val, pos = text[pos + 1:close.start()], close.end()
        else:
            nxt = re.search(r"[;|]", text[pos:])
            stop = pos + nxt.start() if nxt else n
            val, pos = text[pos:stop].strip(), stop
            if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
                val = val[1:-1]
        if op in (">", "<", ">=", "<=", "~", "!~") and val == "":
            return None
        groups[-1].append((col, op, val))
        if pos < n:
            sep = text[pos]
            pos += 1
            if sep == "|":
                if not groups[-1]:
                    return None
                groups.append([])
    groups = [g for g in groups if g]
    return groups or None


def _where_value_text(v: str) -> str:
    return f"'{v}'" if re.search(r"[;|]", v) or v != v.strip() else v


def where_text(groups: List[List[Cond]]) -> str:
    """조건 묶음을 --where 식으로 (값에 ; | 가 있으면 작은따옴표)."""
    return " | ".join("; ".join(f"{c} {o} {_where_value_text(v)}" for c, o, v in g) for g in groups)


def _outside_quotes(text: str, fn: Any) -> str:
    """따옴표로 감싼 부분은 그대로 두고 나머지에만 fn(글)을 적용."""
    return "".join(part if i % 2 else fn(part) for i, part in enumerate(_WHERE_QUOTED.split(text)))


def parse_where(text: str) -> List[List[Cond]]:
    """--where '<열> <연산자> <값>; ... | ...' -> 또는 묶음 [[(열, 연산자, 값), ...], ...] (묶음 안은 모두 맞아야 함).
    고칠 수 있는 표기(==, =>, <>, AND, OR)는 [추천]으로 고친 식."""
    raw = str(text).strip()

    def fix(part: str) -> str:
        part = _WHERE_CONTAINS.sub(" ~ ", _WHERE_OR.sub(" | ", _WHERE_AND.sub("; ", _WHERE_COMMA.sub("; ", part))))
        for a, b in _WHERE_FIXES:
            part = part.replace(a, b)
        return part
    fixed = _outside_quotes(raw, fix)
    groups = _where_parts(fixed)
    if groups is None:
        raise rt.InputError(f"--where 식 '{raw}'을(를) 읽을 수 없습니다. 형식: {WHERE_FORM}")
    if fixed != raw:
        raise rt.InputError(f"--where 식 '{raw}'에 쓸 수 없는 표기가 있습니다 (같음은 =, 같지 않음은 !=, 그리고는 ; (쉼표 아님), 또는은 |). "
                            f"형식: {WHERE_FORM}. [추천] {rt.retry_command({'--where': where_text(groups)})}")
    return groups


_WHERE_LOOSE_NUM = re.compile(r"^([+-]?(?:\d[\d,]*\.?\d*|\.\d+))\s*(%|[^\W\d_]{1,4})?$")
_WHERE_DATE = re.compile(r"^(\d{4})\s*(?:[-/.]|년)\s*(\d{1,2})\s*(?:[-/.]|월)\s*(\d{1,2})\s*일?\.?"
                         r"(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$")
_LEADING_ZERO = re.compile(r"^[+-]?0\d")


def _where_num(v: Any) -> Optional[float]:
    """숫자 칸, 또는 숫자로 읽히는 글 ('1,500', '50%' -> 0.5, '1500원'처럼 숫자 뒤 단위 글자)."""
    if isinstance(v, bool):
        return None
    if _is_num(v):
        return float(v)
    if isinstance(v, str):
        m = _WHERE_LOOSE_NUM.match(v.strip())
        if m and re.search(r"\d", m.group(1)):
            try:
                n = float(m.group(1).replace(",", ""))
            except ValueError:
                return None
            return n / 100 if m.group(2) == "%" else n
    return None


def _where_date(v: Any) -> Optional[Tuple[_dt.datetime, bool]]:
    """날짜 칸, 또는 날짜로 읽히는 글 (2026-1-10, 2026/01/10, 2026.1.10, 2026년 1월 10일, 시각 붙어도 됨) -> (일시, 날짜만인지)."""
    if isinstance(v, bool) or v is None:
        return None
    if _is_date(v):
        d = v if isinstance(v, _dt.datetime) else _dt.datetime(v.year, v.month, v.day)
        d = d.replace(tzinfo=None)
        return d, not (d.hour or d.minute or d.second)
    if isinstance(v, str):
        m = _WHERE_DATE.match(v.strip())
        if m:
            try:
                parts = [int(x) for x in m.groups(default="0")]
                return _dt.datetime(*parts), m.group(4) is None
            except ValueError:
                return None
    return None


def _where_text(v: Any) -> str:
    return "" if v is None else str(out_value(v)).strip()


def _where_kind(val: str) -> str:
    """조건 값의 종류: num(숫자, %, 단위 꼬리), date(느슨한 날짜), text."""
    s = val.strip()
    if _where_date(s) is not None:
        return "date"
    if _where_num(s) is not None:
        return "num"
    return "text"


def _where_match(cell: Any, op: str, val: str) -> Optional[bool]:
    """숫자끼리는 숫자로(조건 값의 %·단위 꼬리 해석), 날짜끼리는 날짜로(느슨한 날짜 표기), 그 밖은 앞뒤 공백을 뺀 글자로(대소문자 무시).
    ~는 포함. = 와 != 에서 앞자리 0이 있는 글('010')은 글자 그대로 비교. 비교할 수 없으면(크기 비교인데 칸이 숫자·날짜가 아님) None."""
    if op == "~":
        return val.strip().casefold() in _where_text(cell).casefold()
    if op == "!~":                      # 포함하지 않음 (1.4.99, 검토 4장 2번: 전에는 '상태 !'를 열 이름으로 읽음)
        return val.strip().casefold() not in _where_text(cell).casefold()
    kind = _where_kind(val)
    a: Any
    b: Any
    if op in ("=", "!=") and (_LEADING_ZERO.match(val.strip()) or (isinstance(cell, str) and _LEADING_ZERO.match(cell.strip()))):
        a, b = _where_text(cell).casefold(), val.strip().casefold()
    elif kind == "date" and _where_date(cell) is not None:
        (cd, _), (vd, date_only) = _where_date(cell), _where_date(val)
        a, b = (cd.date(), vd.date()) if date_only else (cd, vd)
    elif op in ("=", "!=") and isinstance(cell, str):
        cs, vs = cell.strip().replace(",", ""), val.strip().replace(",", "")
        if _WHERE_NUM.match(cs) and _WHERE_NUM.match(vs):     # 글 칸끼리 같음은 둘 다 숫자 모양일 때만 숫자로 ('1A'와 '1B'는 다름)
            a, b = float(cs), float(vs)
        else:
            a, b = _where_text(cell).casefold(), val.strip().casefold()
    elif kind == "num" and _where_num(cell) is not None and not _is_date(cell):
        a, b = _where_num(cell), _where_num(val)
    elif op not in ("=", "!="):
        if kind != "text" or cell is None or _where_text(cell) == "" or _is_num(cell) or _is_date(cell):
            return None                 # 크기 비교인데 칸을 조건 값과 같은 종류로 읽을 수 없음 (빈 칸, 글자, 숫자 칸에 글 조건)
        a, b = _where_text(cell).casefold(), val.strip().casefold()
    else:
        a, b = _where_text(cell).casefold(), val.strip().casefold()
    return {"=": a == b, "!=": a != b, ">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b}[op]


def _where_unreadable(matrix: List[List[Any]], groups: List[List[Tuple[int, str, str]]], names: Dict[int, str]) -> Optional[str]:
    """크기 비교 조건인데 그 열의 값이 있는 칸을 하나도 비교할 수 없으면 그 이유 (0건 VERIFIED로 답하지 않게)."""
    for g in groups:
        for j, op, val in g:
            if op not in (">", "<", ">=", "<="):
                continue
            filled = [r[j] for r in matrix[1:] if j < len(r) and r[j] is not None and _where_text(r[j]) != ""]
            if filled and all(_where_match(c, op, val) is None for c in filled):
                kind = {"num": "숫자", "date": "날짜", "text": "글자"}[_where_kind(val)]
                sample = _where_text(filled[0])[:30]
                return (f"조건 '{names[j]} {op} {val}'을(를) 비교할 수 없습니다: 조건 값은 {kind}로 읽었는데 열 '{names[j]}'의 값(예: {sample})은 "
                        f"{kind}로 읽을 수 없습니다. 숫자는 1500, 1,500, 50%, 1500원처럼, 날짜는 2026-01-10처럼 쓰세요.")
    return None


def _where_header_row(matrix: List[List[Any]], cols: List[str], scan: int = 10) -> int:
    """조건의 열 이름이 첫 행에 없으면 위 scan행 안에서 모든 조건 열 이름이 있는 행 (0이면 첫 행 그대로). 열 문자(B)는 이름이 아님."""
    def names(row: List[Any]) -> set:
        return {rt.name_key(_where_text(v)) for v in row if _where_text(v)}
    keys = [rt.name_key(c) for c in cols if not re.fullmatch(r"[A-Za-z]{1,3}", c.strip())]
    if not keys or not matrix or all(k in names(matrix[0]) for k in keys):
        return 0
    for i in range(1, min(scan, len(matrix) - 1)):
        if all(k in names(matrix[i]) for k in keys):
            return i
    return 0


def _read_where(args, conds: List[List[Cond]], matrix: List[List[Any]], sheet: str, r0: int, c0: int,
                full_addr: str, result: Dict[str, Any], save_cmd: Dict[str, str]) -> Dict[str, Any]:
    """첫 행을 머리글로 보고 조건에 맞는 행만 돌려줍니다 (또는 묶음 중 하나라도 묶음 안의 조건이 모두 맞으면).
    values[0]은 머리글, row_numbers는 values와 같은 순서의 시트 행 번호."""
    ncols = max((len(r) for r in matrix), default=0)
    hi = _where_header_row(matrix, [col for g in conds for col, _, _ in g])
    if hi:
        # 제목 행 아래의 머리글 (검토 4장 2번: 제목을 머리글로 보고 열 이름 [후보]를 내던 것). 그 위 행은 결과에 넣지 않음
        matrix, r0 = matrix[hi:], r0 + hi
        result["header_row_found"] = r0
        full_addr = addr(r0, c0, r0 + len(matrix) - 1, c0 + max(ncols, 1) - 1)
    header = matrix[0] if matrix else []
    groups = [[(resolve_column(col, header if matrix else None, c0, ncols, "조건 열")[0], op, val) for col, op, val in g]
              for g in conds]

    def cell(i: int, j: int) -> Any:
        return matrix[i][j] if j < len(matrix[i]) else None
    names = {j: str(out_value(header[j])) if j < len(header) and _where_text(header[j]) else col
             for g, cg in zip(groups, conds) for (j, _, _), (col, _, _) in zip(g, cg)}
    bad = _where_unreadable(matrix, groups, names)
    if bad:
        raise rt.InputError(bad + f" 형식: {WHERE_FORM}")
    hits = [i for i in range(1, len(matrix))
            if any(all(_where_match(cell(i, j), op, val) for j, op, val in g) for g in groups)]
    values = [[out_value(v) for v in matrix[i]] for i in [0] + hits] if matrix else []
    numbers = [r0 + i for i in [0] + hits] if matrix else []
    cond_text = where_text(conds)
    result.update(where=cond_text, header_row=r0, count=len(hits), data_rows=max(len(matrix) - 1, 0))
    where = f"{sheet}!{full_addr}"
    if not hits:      # R2 0건
        return rt.ok(TOOL, "read", f"0건: {where}에 조건({cond_text})에 맞는 행이 없습니다",
                     f"VERIFIED: 0건. {where}에 조건({cond_text})에 맞는 행이 없습니다. 사용자에게 그대로 알리고 조건을 바꿔 다시 "
                     "실행하지 마세요.", values=values, row_numbers=numbers, truncated=False, **result)
    if args.save:
        out = rt.resolve_output_path(None if args.save is True else args.save, f"{sheet}_{full_addr.replace(':', '-')}_조건", ".csv")
        out = rt.write_csv_rows(out, values)
        # 열린 문서를 바꾸는 이어 하기는 confirm_commands (1.4.75): next_commands는 묻지 않고 실행해도 되는 읽기·미리보기만
        result.update(output_file=out, confirm_commands=[
            rt.next_command("word.py", ["write", "--file", out], "저장한 표를 열린 워드 문서에 넣기 (사용자가 원할 때만)"),
            rt.next_command("ppt.py", ["write", "--file", out], "저장한 표를 열린 PPT에 넣기 (사용자가 원할 때만)")])
        filled = sum(1 for r in values if any(v is not None and str(v).strip() for v in r))
        return rt.ok_file(TOOL, "read", f"{where}에서 조건({cond_text})에 맞는 {len(hits)}행을 머리글과 함께 CSV로 저장: {out}", out,
                          "VERIFIED: 저장한 CSV를 다시 열어 행 수를 확인했습니다. " + rt.INTO_OTHER_DOC,
                          expected_rows=filled, delimiter=",", **result)
    shown, rows_cut, cells_cut, cols_dropped = rt.clamp_rows(values, args.max_rows)
    truncated = rows_cut or cells_cut > 0 or cols_dropped > 0
    result.update(values=shown, row_numbers=numbers[:len(shown)], truncated=truncated)
    reasons = ([f"맞는 행 {len(hits)}개 중 앞 {len(shown) - 1}개만 표시"] if rows_cut else []) + (
        [f"{rt.MAX_CELL_CHARS}자보다 긴 값 {cells_cut}개는 앞부분만 표시"] if cells_cut else []) + (
        [f"첫 행이 표시 상한을 넘어 뒤쪽 {cols_dropped}열은 표시하지 않음"] if cols_dropped else [])
    preview = f"{where}: 조건({cond_text})에 맞는 행 {len(hits)}개 (데이터 {len(matrix) - 1}행 중)" + (
        f" (앞 {len(shown) - 1}개만 표시)" if rows_cut else "")
    if truncated:
        result.update(truncated_reason=", ".join(reasons), next_commands=[save_cmd])
        return rt.ok(TOOL, "read", preview, "VERIFIED: 거르기 완료. 일부만 표시했습니다(truncated_reason). 전체가 필요하면 "
                     "next_commands[0]으로 CSV에 저장하세요.", **result)
    return rt.ok(TOOL, "read", preview, "", **result)



# ==============================================================================
# stats: 표 집계 (2026-11-09 검토 4장 1번, 1.4.98)
# ==============================================================================

STATS_GROUPS_MAX = 50        # 결과 JSON에 담는 묶음 수 (넘으면 --save CSV)
STATS_TOP_VALUES = 10        # 글자 열의 많이 나온 값 개수


def _load_table(args) -> Tuple[str, int, int, List[List[Any]], Dict[str, Any], Optional[Dict[str, Any]]]:
    """read와 같은 대상(열린 문서, 닫힌 xlsx, CSV)에서 표를 읽음. (시트, 시작 행, 시작 열, 값, 대상 정보, 계산 안 된 수식 칸)."""
    csv_path = closed_csv(args.target)
    closed = None if csv_path else is_closed_xlsx(args.target)
    if csv_path:
        sheet, r0, c0, matrix, _ = _read_matrix_csv(csv_path, args.range)
        return sheet, r0, c0, matrix, {"app": "Excel", "name": os.path.basename(csv_path), "path": csv_path, "source": "file(csv)"}, None
    if closed:
        sheet, r0, c0, matrix, _ = _read_matrix_closed(closed, args.sheet, args.range)
        return (sheet, r0, c0, matrix, {"app": "Excel", "name": os.path.basename(closed), "path": closed, "source": "file(openpyxl)"},
                closed_uncomputed(closed, sheet, r0, c0, matrix))
    with off.open_target("excel", args.target, readonly=True) as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range)
        matrix = off.range_to_matrix(rng.Value)
        return str(rng.Worksheet.Name), int(rng.Row), int(rng.Column), matrix, t.resolved(), None


def _stat_num(v: Any) -> Optional[float]:
    if isinstance(v, bool) or v is None:
        return None
    if _is_num(v):
        return float(v)
    if isinstance(v, str) and _LEADING_ZERO.match(v.strip()):
        return None                     # 앞자리 0이 있는 글('0101' 품번)은 숫자가 아니라 코드
    return _where_num(v) if isinstance(v, str) and v.strip() else None


def _fmt(x: float) -> str:
    if abs(x) >= 1e15 or (x != 0 and abs(x) < 1e-4):
        return f"{x:.6g}"
    r = round(x, 4)
    return f"{int(r):,}" if float(r).is_integer() else f"{r:,.4f}".rstrip("0").rstrip(".")


def _column_stats(values: List[Any]) -> Dict[str, Any]:
    """한 열의 집계: 숫자가 많으면 건수·합계·평균·최소·최대, 글자가 많으면 건수·고유값 수·많이 나온 값."""
    filled = [v for v in values if v is not None and _where_text(v) != ""]
    nums = [n for n in (_stat_num(v) for v in filled) if n is not None]
    if nums and len(nums) * 2 >= len(filled):
        out: Dict[str, Any] = {"kind": "number", "count": len(nums), "sum": sum(nums), "mean": sum(nums) / len(nums),
                               "min": min(nums), "max": max(nums)}
        if len(nums) < len(filled):
            out["not_number"] = len(filled) - len(nums)      # 숫자로 읽지 못해 뺀 칸 수
        return out
    counts: Dict[str, int] = {}
    for v in filled:
        k = _where_text(v)
        counts[k] = counts.get(k, 0) + 1
    top = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))[:STATS_TOP_VALUES]
    return {"kind": "text", "count": len(filled), "unique": len(counts), "top": [{"value": k, "count": n} for k, n in top]}


def _stats_text(name: str, st: Dict[str, Any]) -> str:
    if st["kind"] == "number":
        return (f"{name}: 합계 {_fmt(st['sum'])}, 평균 {_fmt(st['mean'])}, 최소 {_fmt(st['min'])}, 최대 {_fmt(st['max'])} "
                f"({st['count']}개)")
    tops = ", ".join(f"{t['value']} {t['count']}" for t in st["top"][:3])
    return f"{name}: {st['count']}개, 고유값 {st['unique']}개 (많은 값: {tops})"


def cmd_stats(args) -> Dict[str, Any]:
    """첫 행을 머리글로 보고 열마다 합계·평균·최소·최대·건수(숫자 열), 건수·고유값(글자 열). --by로 묶음별, --where로 거른 행만,
    --top N은 묶음을 첫 집계 열의 합계가 큰 순으로 N개. 닫힌 xlsx·CSV는 Office 없이. 값은 도구가 계산해 preview에 글로."""
    conds = parse_where(args.where) if args.where else None
    if args.top is not None and args.top < 1:
        raise rt.InputError(f"--top은 1 이상입니다 (받은 값 {args.top}). [추천] {rt.retry_command({'--top': 5})}")
    sheet, r0, c0, matrix, source, unc = _load_table(args)
    if len(matrix) < 2:
        raise rt.InputError(f"{sheet}에 머리글 아래 데이터 행이 없습니다. --range로 표 범위를 주세요 (첫 행은 머리글).")
    header = matrix[0]
    ncols = max(len(r) for r in matrix)
    names = [_where_text(header[j]) if j < len(header) and _where_text(header[j]) else f"{col_letter(c0 + j)}열" for j in range(ncols)]

    def cell(i: int, j: int) -> Any:
        return matrix[i][j] if j < len(matrix[i]) else None
    rows = list(range(1, len(matrix)))
    warnings: List[str] = []
    if conds is not None:
        groups_w = [[(resolve_column(col, header, c0, ncols, "조건 열", option="--where")[0], op, val) for col, op, val in g] for g in conds]
        bad = _where_unreadable(matrix, groups_w, {j: names[j] for g in groups_w for j, _, _ in g})
        if bad:
            raise rt.InputError(bad + f" 형식: {WHERE_FORM}")
        rows = [i for i in rows if any(all(_where_match(cell(i, j), op, val) for j, op, val in g) for g in groups_w)]
    by = [resolve_column(b.strip(), header, c0, ncols, "묶을 열", option="--by")[0] for b in args.by.split(",") if b.strip()] if args.by else []
    if args.columns:
        cols = [resolve_column(c.strip(), header, c0, ncols, "집계할 열", option="--columns")[0] for c in args.columns.split(",") if c.strip()]
    else:   # 생략: 숫자가 많은 열 전부 (묶는 열 제외)
        cols = [j for j in range(ncols) if j not in by
                and _column_stats([cell(i, j) for i in range(1, len(matrix))]).get("kind") == "number"]   # 거르기 전 전체 행으로 정함
        if not cols:
            raise rt.InputError("숫자 열을 찾지 못했습니다. 집계할 열을 --columns로 주세요 (예: --columns 금액,수량).")
    if args.top is not None and not by:
        warnings.append("--top은 --by(묶음)와 함께 쓰는 옵션이라 쓰지 않음")
    buckets: Dict[Tuple[str, ...], List[int]] = {}
    for i in rows:
        key = tuple(_where_text(cell(i, j)) for j in by)
        buckets.setdefault(key, []).append(i)
    result_groups = []
    for key, idx in buckets.items():
        g = {"rows": len(idx), "stats": {names[j]: _column_stats([cell(i, j) for i in idx]) for j in cols}}
        if by:
            g = {"by": {names[j]: k for j, k in zip(by, key)}, **g}
        result_groups.append(g)
    first = names[cols[0]]
    if by:
        def sort_key(g: Dict[str, Any]) -> Any:
            st = g["stats"][first]
            return (-(st.get("sum") or 0.0), -g["rows"]) if args.top is not None else tuple(g["by"].values())
        result_groups.sort(key=sort_key)
        if args.top is not None:
            result_groups = result_groups[:args.top]
    where = f"{sheet}!{addr(r0, c0, r0 + len(matrix) - 1, c0 + ncols - 1)}"
    cond_text = f", 조건 {where_text(conds)}" if conds else ""
    fields: Dict[str, Any] = {"target_resolved": source, "sheet": sheet, "range": where.split("!")[-1], "data_rows": len(matrix) - 1,
                              "rows_used": len(rows), "columns": [names[j] for j in cols], "group_count": len(result_groups)}
    if conds:
        fields["where"] = where_text(conds)
    if by:
        fields["by"] = [names[j] for j in by]
    if warnings:
        fields["warnings"] = warnings
    if not rows:
        return rt.ok(TOOL, "stats", f"0건: {where}에 조건({where_text(conds)})에 맞는 행이 없습니다",
                     f"VERIFIED: 0건. 조건({where_text(conds)})에 맞는 행이 없어 집계할 것이 없습니다. 사용자에게 그대로 알리세요.",
                     groups=[], **fields)
    shown = result_groups[:STATS_GROUPS_MAX]
    fields.update(groups=shown, shown=len(shown), truncated=len(result_groups) > len(shown))
    if args.save or fields["truncated"]:
        head = [names[j] for j in by] + ["행 수"]
        for j in cols:
            st0 = result_groups[0]["stats"][names[j]]
            head += ([f"{names[j]} 합계", f"{names[j]} 평균", f"{names[j]} 최소", f"{names[j]} 최대", f"{names[j]} 건수"]
                     if st0["kind"] == "number" else [f"{names[j]} 건수", f"{names[j]} 고유값"])
        table = [head]
        for g in result_groups:
            line = list((g.get("by") or {}).values()) + [g["rows"]]
            for j in cols:
                st = g["stats"][names[j]]
                line += ([st.get("sum"), st.get("mean"), st.get("min"), st.get("max"), st["count"]] if st["kind"] == "number"
                         else [st["count"], st.get("unique")])
            table.append(line)
        out = rt.resolve_output_path(None if args.save in (None, True) else args.save, f"{sheet}_집계", ".csv")
        out = rt.write_csv_rows(out, table)
        fields["output_file" if args.save else "full_result_file"] = out
    if by:
        lines = [", ".join(f"{k}" for k in g["by"].values()) + f" ({g['rows']}행): "
                 + "; ".join(_stats_text(n, st) for n, st in g["stats"].items()) for g in shown[:5]]
        preview = (f"{where}{cond_text}: {', '.join(fields['by'])}별 {len(result_groups)}묶음"
                   + (f" (합계 상위 {args.top})" if args.top else "") + ". " + " / ".join(lines)
                   + (f" / ... 외 {len(shown) - 5}묶음" if len(shown) > 5 else ""))
    else:
        preview = f"{where}{cond_text} ({len(rows)}행): " + "; ".join(_stats_text(n, st) for n, st in shown[0]["stats"].items())
    directive = ("VERIFIED: 도구가 읽은 값으로 집계했습니다. preview의 숫자를 그대로 알리세요 (직접 더하거나 다시 계산하지 마세요)."
                 + (" 묶음이 많아 일부만 담았습니다: 전체는 full_result_file." if fields["truncated"] else ""))
    payload = rt.ok(TOOL, "stats", preview, directive, **fields)
    return xf.with_uncomputed(payload, unc) if unc else payload


XL_SAVE_FORMATS = {".xlsx": 51, ".xlsm": 52}


def cmd_save(args) -> Dict[str, Any]:
    """열린 통합문서를 저장 (Ctrl+S). --save <경로>면 다른 이름으로 저장 (그 뒤로는 새 파일이 열린 문서).
    같은 이름 파일은 --overwrite가 있을 때만 덮어씀. 닫힌 파일은 이미 저장되어 있어 할 일이 없음."""
    new_path = None
    if args.save:
        raw = os.path.abspath(os.path.expandvars(os.path.expanduser(str(args.save).strip().strip('"'))))
        ext = os.path.splitext(raw)[1].lower()
        if ext not in XL_SAVE_FORMATS:
            fixed = os.path.splitext(raw)[0] + ".xlsx"
            raise rt.InputError(f"--save는 .xlsx 또는 .xlsm 파일입니다 (받은 값 '{os.path.basename(raw)}'). CSV·PDF 사본은 excel.py export. "
                                f"[추천] {rt.retry_command({'--save': fixed})}")
        new_path = raw
    with off.open_target("excel", args.target, readonly=True, tool=TOOL, command="save") as t:
        if t.source not in ("active", "open"):
            raise rt.InputError(f"'{t.name}'은(는) 열려 있지 않은 파일이라 저장할 것이 없습니다 (파일은 이미 저장되어 있음). "
                                "다른 이름의 사본이 필요하면 excel.py export를 쓰세요.")
        wb = t.doc
        old = t.path
        if new_path is None:
            if bool(getattr(wb, "ReadOnly", False)):
                raise rt.InputError(f"'{t.name}'은(는) 읽기 전용으로 열려 있어 그대로 저장할 수 없습니다 (Excel이 확인 창을 띄움). 사용자에게 "
                                    "다른 이름으로 저장할 전체 경로를 묻고 --save \"<전체경로.xlsx>\"로 다시 실행하세요.")
            if not old:
                raise rt.InputError(f"'{t.name}'은(는) 한 번도 저장하지 않은 새 통합문서라 저장할 위치가 없습니다. 사용자에게 저장할 "
                                    "전체 경로를 묻고 --save \"<전체경로.xlsx>\"로 다시 실행하세요.")
            was_saved = bool(wb.Saved)
            if not was_saved:
                wb.Save()
            ok = bool(wb.Saved)
            resolved = t.resolved()
            preview = f"'{t.name}'을(를) 저장했습니다" if not was_saved else f"'{t.name}'은(는) 바뀐 것이 없어 이미 저장된 상태입니다"
            if not ok:
                return rt.ok_unverified(TOOL, "save", preview, ["저장 뒤에도 '저장 안 됨' 상태 (Excel이 편집 중이거나 읽기 전용)"],
                                        target_resolved=resolved, output_file=old)
            return rt.ok(TOOL, "save", preview, "VERIFIED: 저장 상태를 다시 확인했습니다. 같은 명령을 다시 실행하지 마세요.",
                         verified=True, target_resolved=resolved, output_file=old, unchanged=was_saved)
        if os.path.normcase(new_path) == os.path.normcase(old or ""):
            raise rt.InputError(f"--save 경로가 지금 파일과 같습니다. 그대로 저장은 --save 없이: [추천] {rt.retry_command({'--save': None})}")
        if os.path.exists(new_path) and not args.overwrite:
            raise rt.UserDecisionError(
                f"'{new_path}' 파일이 이미 있습니다. 다른 이름으로 저장하면 그 파일이 바뀝니다.",
                "저장하지 않았습니다. 사용자에게 덮어쓸지 묻고, 덮어쓴다고 하면 같은 번호의 confirm_commands를 1번 실행하세요. "
                "아니면 사용자에게 다른 경로를 받으세요.",
                candidates=["같은 이름 파일을 덮어써 저장"], commands=[rt.retry_command({"--overwrite": True})], destructive=True)
        ext = os.path.splitext(new_path)[1].lower()
        try:
            macros = bool(wb.HasVBProject)
        except Exception:
            macros = False
        if macros and ext == ".xlsx":
            raise rt.InputError("이 통합문서에는 매크로가 있어 .xlsx로 저장하면 매크로가 없어집니다. "
                                f"[추천] {rt.retry_command({'--save': os.path.splitext(new_path)[0] + '.xlsm'})}")
        os.makedirs(os.path.dirname(new_path), exist_ok=True)
        prev = t.app.DisplayAlerts
        t.app.DisplayAlerts = False                     # '바꾸시겠습니까?' 창 (덮어쓰기는 위에서 확인함)
        try:
            wb.SaveAs(new_path, XL_SAVE_FORMATS[ext])
        finally:
            t.app.DisplayAlerts = prev
        got = str(getattr(wb, "FullName", "") or "")
        resolved = t.resolved()
    preview = f"'{os.path.basename(new_path)}'(으)로 저장했습니다. 지금 열린 문서는 이 새 파일입니다" + (f" (원래 파일 '{os.path.basename(old)}'은 그대로)" if old else "")
    if os.path.normcase(got) != os.path.normcase(new_path) or not os.path.isfile(new_path):
        return rt.ok_unverified(TOOL, "save", preview, [f"저장 뒤 파일·열린 문서 경로를 확인하지 못함 ({got or '경로 없음'})"],
                                target_resolved=resolved, output_file=new_path)
    return rt.ok(TOOL, "save", preview, "VERIFIED: 새 파일이 생기고 열린 문서가 그 파일로 바뀐 것을 확인했습니다. 같은 명령을 다시 실행하지 마세요.",
                 verified=True, target_resolved=resolved, output_file=new_path, previous_file=old)


DUP_GROUPS_MAX = 30       # JSON에 담는 중복 묶음 수 (넘으면 전체는 CSV)


def _dup_key(v: Any) -> str:
    """중복 비교 값: 앞뒤 공백·전각/반각·대소문자 무시, 숫자는 1 = 1.0 (품번의 앞자리 0은 그대로 다른 값)."""
    t = unicodedata.normalize("NFKC", _where_text(v)).casefold()
    if re.fullmatch(r"-?\d+\.0+", t):
        t = t.split(".")[0]
    return t


def cmd_duplicates(args) -> Dict[str, Any]:
    """첫 행을 머리글로 보고 --columns(생략하면 모든 열)의 값이 같은 행 묶음을 찾음 (읽기만). 닫힌 xlsx·CSV는 Office 없이.
    엑셀 행 번호와 값을 알려 주고 지우지는 않음 (지우기는 사용자가 행을 확인한 뒤 range --action delete)."""
    sheet, r0, c0, matrix, source, unc = _load_table(args)
    if len(matrix) < 2:
        raise rt.InputError(f"{sheet}에 머리글 아래 데이터 행이 없습니다. --range로 표 범위를 주세요 (첫 행은 머리글).")
    header = matrix[0]
    ncols = max(len(r) for r in matrix)
    names = [_where_text(header[j]) if j < len(header) and _where_text(header[j]) else f"{col_letter(c0 + j)}열" for j in range(ncols)]
    cols = ([resolve_column(c.strip(), header, c0, ncols, "비교할 열", option="--columns")[0] for c in args.columns.split(",") if c.strip()]
            if args.columns else list(range(ncols)))

    def cell(i: int, j: int) -> Any:
        return matrix[i][j] if j < len(matrix[i]) else None
    buckets: Dict[Tuple[str, ...], List[int]] = {}
    for i in range(1, len(matrix)):
        key = tuple(_dup_key(cell(i, j)) for j in cols)
        if all(k == "" for k in key):
            continue                                    # 빈 행은 중복으로 보지 않음
        buckets.setdefault(key, []).append(i)
    groups = [{"values": {names[j]: out_value(cell(idx[0], j)) for j in cols}, "rows": [r0 + i for i in idx], "count": len(idx)}
              for idx in buckets.values() if len(idx) > 1]
    groups.sort(key=lambda g: (-g["count"], g["rows"][0]))
    where = f"{sheet}!{addr(r0, c0, r0 + len(matrix) - 1, c0 + ncols - 1)}"
    by = ", ".join(names[j] for j in cols) if args.columns else "모든 열"
    extra_rows = sum(g["count"] - 1 for g in groups)
    fields: Dict[str, Any] = {"target_resolved": source, "sheet": sheet, "range": where.split("!")[-1], "data_rows": len(matrix) - 1,
                              "columns": [names[j] for j in cols], "group_count": len(groups), "duplicate_rows": extra_rows}
    if not groups:
        payload = rt.ok(TOOL, "duplicates", f"0건: {where}에서 {by} 값이 같은 행이 없습니다 ({len(matrix) - 1}행 비교)",
                        f"VERIFIED: 0건. {by} 값이 같은 행이 없습니다. 사용자에게 그대로 알리세요.", groups=[], **fields)
        return xf.with_uncomputed(payload, unc) if unc else payload
    shown = groups[:DUP_GROUPS_MAX]
    fields.update(groups=shown, shown=len(shown), truncated=len(groups) > len(shown))
    if args.save or fields["truncated"]:
        table = [["묶음", "엑셀 행", "같은 행 수"] + [names[j] for j in cols]]
        for n, g in enumerate(groups, 1):
            for r in g["rows"]:
                table.append([n, r, g["count"]] + [out_value(cell(r - r0, j)) for j in cols])
        out = rt.resolve_output_path(None if args.save in (None, True) else args.save, f"{sheet}_중복", ".csv")
        out = rt.write_csv_rows(out, table)
        fields["output_file" if args.save else "full_result_file"] = out
    lines = [", ".join(str(v) for v in g["values"].values()) + f": {', '.join(map(str, g['rows'][:6]))}행" for g in shown[:5]]
    preview = (f"{where}에서 {by} 값이 같은 묶음 {len(groups)}개 (남기고 지울 수 있는 행 {extra_rows}개). "
               + " / ".join(lines) + (f" / ... 외 {len(groups) - 5}묶음" if len(groups) > 5 else ""))
    directive = ("VERIFIED: 중복을 찾았습니다(읽기만, 지우지 않음). preview의 묶음과 엑셀 행 번호를 그대로 알리세요. "
                 "지우기는 사용자가 지울 행을 말한 뒤에만 합니다." + (" 묶음이 많아 일부만 담았습니다: 전체는 full_result_file." if fields["truncated"] else ""))
    payload = rt.ok(TOOL, "duplicates", preview, directive, **fields)
    return xf.with_uncomputed(payload, unc) if unc else payload


XL_CHART_NAMES = {51: "column", 57: "bar", 4: "line", 65: "line_markers", 5: "pie", 1: "area",
                  -4169: "scatter", 75: "scatter_lines", 74: "scatter_lines_markers"}


def _chart_detail(obj: Any, used_box: Optional[Dict[str, float]], warnings: List[str]) -> Dict[str, Any]:
    """차트 종류·계열 수·범례·위치와, 데이터 셀을 가리는지."""
    ch = obj.Chart
    code = int(ch.ChartType)
    d: Dict[str, Any] = {"chart_type": XL_CHART_NAMES.get(code, str(code)), "chart_type_code": code,
                         "series": int(ch.SeriesCollection().Count), "legend": bool(ch.HasLegend), **off.shape_box(obj)}
    if used_box is not None and off.overlap_ratio(d, used_box) > 0:
        d["covers_data"] = True
        warnings.append(f"차트 '{obj.Name}'이(가) 데이터 셀을 가립니다")
    return d


def _sheet_summary(ws: Any, detail: bool = False) -> Dict[str, Any]:
    used = ws.UsedRange
    info: Dict[str, Any] = {"name": str(ws.Name), "used_range": str(used.Address).replace("$", ""),
                            "rows": int(used.Rows.Count), "cols": int(used.Columns.Count)}
    charts = []
    warnings: List[str] = []
    used_box = None
    if detail and int(used.Rows.Count) * int(used.Columns.Count) > 1:
        try:
            used_box = off.shape_box(used)
        except Exception:
            pass
    try:
        objs = ws.ChartObjects()
        for i in range(1, objs.Count + 1):
            ch = objs.Item(i)
            title = ""
            try:
                if ch.Chart.HasTitle:
                    title = str(ch.Chart.ChartTitle.Text)
            except Exception:
                pass
            item = {"name": str(ch.Name), "title": title}
            if detail:
                try:
                    item.update(_chart_detail(ch, used_box, warnings))
                except Exception as e:
                    rt.debug(f"차트 상세 생략: {e}")
            charts.append(item)
    except Exception:
        pass
    info["charts"] = charts
    if detail:
        info["warnings"] = warnings
    try:
        info["shapes"] = int(ws.Shapes.Count) - len(charts)
    except Exception:
        pass
    return info


def _closed_info(path: str) -> Dict[str, Any]:
    """닫힌 xlsx의 시트 목록과 데이터 범위 (Office 없이, 2026-11-09 검토 4장 1번: info가 늘 Excel을 띄웠음). 차트 시트는 kind 'chart'."""
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "Office 없이 닫힌 엑셀 파일의 시트 목록")
    wb = _load_closed(openpyxl, path)
    try:
        sheets = []
        for ws in wb.worksheets:
            r1, c1, r2, c2 = _closed_bounds(ws, path)
            used = addr(r1, c1, r2, c2) if r2 >= r1 and c2 >= c1 else ""
            sheets.append({"name": ws.title, "used_range": used, "rows": max(0, r2 - r1 + 1), "cols": max(0, c2 - c1 + 1),
                           "hidden": getattr(ws, "sheet_state", "visible") != "visible"})
        sheets += [{"name": c.title, "kind": "chart"} for c in getattr(wb, "chartsheets", [])]
        active = getattr(wb.active, "title", None)
    finally:
        wb.close()
    source = {"app": "Excel", "name": os.path.basename(path), "path": path, "source": "file(openpyxl)"}
    names = ", ".join(f"{x['name']}({x.get('used_range') or '차트'})" for x in sheets[:10])
    return rt.ok(TOOL, "info", f"{os.path.basename(path)}: 시트 {len(sheets)}개 - {names}" + (" 등" if len(sheets) > 10 else ""),
                 "VERIFIED: 파일을 읽어 시트 목록을 확인했습니다 (Office 없이).", target_resolved=source, sheets=sheets,
                 active_sheet=active)


def cmd_info(args) -> Dict[str, Any]:
    if args.all:
        books = [{"name": d["name"], "path": d["path"], "saved": d["saved"], "active": d["active"]}
                 for d in off.list_open("excel")]
        if not books:
            return off.no_open_docs_result(TOOL, "info", "0건: 열린 통합문서가 없습니다",
                         "VERIFIED: 0건. 열린 Excel 통합문서가 없습니다. 사용자에게 그대로 알리세요.", workbooks=books, count=0)
        return rt.ok(TOOL, "info", f"열린 통합문서 {len(books)}개", workbooks=books, count=len(books))
    closed = None if closed_csv(args.target) else is_closed_xlsx(args.target)
    if closed and not args.detail:
        return _closed_info(closed)
    with off.open_target("excel", args.target, readonly=True) as t:
        wb = t.doc
        sheets = [_sheet_summary(wb.Worksheets(i), args.detail) for i in range(1, wb.Worksheets.Count + 1)]
        result: Dict[str, Any] = {"target_resolved": t.resolved(), "sheets": sheets,
                                  "saved": off.is_saved(wb)}
        if args.detail:
            result["layout_warnings"] = sum(len(s["warnings"]) for s in sheets)
        try:
            result["active_sheet"] = str(wb.ActiveSheet.Name)
            if t.source == "active":
                result["selection"] = str(t.app.Selection.Address).replace("$", "")
        except Exception:
            pass
    preview = f"{result['target_resolved']['name']}: 시트 {len(sheets)}개 (" + ", ".join(
        f"{s['name']} {s['used_range']}" + (f" 차트{len(s['charts'])}" if s["charts"] else "") for s in sheets[:8]) + ")"
    if args.detail:
        preview += f", 배치 경고 {result['layout_warnings']}건" if result["layout_warnings"] else ", 배치 경고 없음"
    if result.get("selection"):
        # 사용자가 마지막으로 클릭한(선택한) 칸: "지금 선택한 칸이 어디야", "여기에 써줘"에 답할 수 있게 preview에도 (2026-11-11)
        preview += f". 보고 있는 시트 '{result['active_sheet']}', 선택한 칸 {result['selection']}"
    return rt.ok(TOOL, "info", preview, **result)


# ==============================================================================
# 4. write
# ==============================================================================

def _last_used_row(ws: Any, col: int) -> int:
    """해당 열 기준 마지막 사용 행 (없으면 0). End(xlUp)은 필터·숨긴 행을 건너뛰어 숨은 데이터 위를 끝으로 잡으므로
    (--append가 숨은 행을 덮어씀, 2026-11-09 검토 C11) 사용 범위 끝까지 값을 한 번 더 읽어 확인합니다."""
    try:
        cell = ws.Cells(ws.Rows.Count, col).End(-4162)   # xlUp
        last = 0 if (cell.Row == 1 and _cell_empty(ws.Cells(1, col))) else int(cell.Row)
    except Exception:
        used = ws.UsedRange
        return int(used.Row + used.Rows.Count - 1)
    try:
        used = ws.UsedRange
        bottom = int(used.Row) + int(used.Rows.Count) - 1
        if bottom > last:
            vals = off.range_to_matrix(ws.Range(ws.Cells(last + 1, col), ws.Cells(bottom, col)).Value)
            for i in range(len(vals) - 1, -1, -1):
                v = vals[i][0] if vals[i] else None
                if v is not None and str(v) != "":
                    return last + 1 + i
    except Exception:
        pass
    return last


def _cell_empty(cell: Any) -> bool:
    """값도 수식도 없는 칸. 빈 글("")을 내는 수식(=IF(A1="","",...))은 값이 비어 보여도 빈 칸이 아님."""
    v = cell.Value
    if v is not None and str(v) != "":
        return False
    try:
        f = cell.Formula
    except Exception:
        return True
    return not (isinstance(f, str) and f != "")


def _write_sheet(t: Any, name: Optional[str], new: bool) -> Optional[Any]:
    """쓸 시트. 없는 이름은 --new-sheet가 있을 때만 새로 만듭니다 (오타로 새 시트가 생기지 않도록).
    --new-sheet면 이름만 확인하고 None을 돌려줍니다: 시트는 모든 확인이 끝난 뒤 쓰기 직전에 만듭니다
    (확인에서 멈추면 새 시트만 남고 되돌리기 지점은 지워져, 다시 실행하면 '이미 있음'으로 막히던 문제)."""
    if new:
        if not name:
            raise rt.InputError("--new-sheet에는 새 시트 이름이 필요합니다. --sheet 이름 또는 --range '이름'!A1로 지정하세요.")
        names = [str(t.doc.Worksheets(i).Name) for i in range(1, t.doc.Worksheets.Count + 1)]
        existing = rt.find_name(name, names)
        if existing is None:
            off.check_sheet_name(name, names)        # 잘못된 이름(/ ? * [ ] : 31자 초과)은 무엇이든 바꾸기 전에
        if existing is not None:
            raise rt.InputError(f"시트 '{existing}'이(가) 이미 있어 새로 만들지 않았습니다. 기존 시트에 쓰려면 --new-sheet를 빼고, "
                                "새 시트는 다른 이름을 지정하세요.")
        return None
    return t.sheet(name, allow_create=True)


def _rebase_formula(app: Any, ws: Any, formula: str, from_row: int, from_col: int, to_row: int, to_col: int) -> str:
    """from 칸 기준으로 쓴 수식을 to 칸에 둘 때의 수식 (상대 참조만 옮기고 $ 참조는 그대로).
    Excel ConvertFormula로 A1 -> R1C1(from 기준) -> A1(to 기준). 바꾸지 못하면 쓰지 않고 STOP_RETRY."""
    try:
        r1c1 = app.ConvertFormula(Formula=formula, FromReferenceStyle=XL_A1, ToReferenceStyle=XL_R1C1,
                                  RelativeTo=ws.Cells(from_row, from_col))
        out = app.ConvertFormula(Formula=r1c1, FromReferenceStyle=XL_R1C1, ToReferenceStyle=XL_A1,
                                 RelativeTo=ws.Cells(to_row, to_col))
        if not isinstance(out, str) or not out.startswith("="):
            raise ValueError(repr(out))
        return out
    except Exception as e:
        rt.debug(f"수식 변환 실패: {e}")
        raise rt.InputError(
            f"--append로 쓸 행이 {from_row}행에서 {to_row}행으로 옮겨졌는데 수식 {formula!r}의 상대 참조를 그 행 기준으로 바꾸지 못해 "
            f"쓰지 않았습니다. --append를 빼고 {to_row}행 기준 수식과 --range {addr(to_row, to_col)}로 다시 실행하세요 "
            f"(수식을 고쳐야 하므로 [추천] 명령은 없음. 사용자에게 알리세요).")


# ==============================================================================
# 3-1. 쓰기 전 확인: 시트 보호, 쓰기 실패 되돌리기 (모든 쓰기 명령 공통)
# ==============================================================================

# 제한된 보기로 열린 문서도 실제 엑셀에서는 모든 칸이 잠긴 시트로 보여 여기로 옴 (회사 PC 2026-11-11 L11): 둘을 구분하지 않고 함께 안내
PROTECT_HINT = ("문서 위쪽 노란 줄에 [편집 사용]이 있으면 누르고, 없으면 검토 > 시트 보호 해제를 사용자가 직접 해야 합니다 "
                "(도구는 암호를 풀지 않음)")
# 보호된 시트에서도 엑셀이 허용할 수 있는 일: 시트 보호 대화상자의 허용 항목 (Worksheet.Protection.Allow*)
_PROTECT_ALLOW = {"format_cells": "AllowFormattingCells", "format_columns": "AllowFormattingColumns",
                  "format_rows": "AllowFormattingRows", "insert_rows": "AllowInsertingRows",
                  "insert_columns": "AllowInsertingColumns", "delete_rows": "AllowDeletingRows",
                  "delete_columns": "AllowDeletingColumns", "sort": "AllowSorting"}
_PROTECT_WHAT = {"values": "값을 쓸", "format_cells": "서식을 바꿀", "format_columns": "열 너비·숨김을 바꿀",
                 "format_rows": "행 높이·숨김을 바꿀", "insert_rows": "행을 넣을", "insert_columns": "열을 넣을",
                 "delete_rows": "행을 지울", "delete_columns": "열을 지울", "sort": "정렬할",
                 "objects": "메모·차트를 넣거나 지울", "structure": "병합·이동·복사할"}
PROTECT_SCAN_CELLS = 2000      # 잠기지 않은 칸 [후보]를 찾을 때 훑는 칸 수 상한


def _is_protected(ws: Any) -> bool:
    try:
        return bool(ws.ProtectContents)
    except Exception:
        return False          # 읽지 못하면 막지 않음 (쓰기가 실패하면 그때 COM 오류로 알림)


def _locked_cells(rng: Any, limit: int = 500) -> List[str]:
    """범위에서 잠긴 칸 주소 (Locked: True 모두, False 없음, None 섞임 -> 칸마다)."""
    try:
        lk = rng.Locked
    except Exception:
        return []
    r0, c0 = int(rng.Row), int(rng.Column)
    nr, nc = int(rng.Rows.Count), int(rng.Columns.Count)
    if lk is False or lk == 0:
        return []
    if lk is True or lk == -1 or lk == 1:
        return [addr(r0 + i, c0 + j) for i in range(min(nr, 50)) for j in range(nc)][:limit]
    out = []
    for i in range(nr):
        for j in range(nc):
            try:
                if rng.Cells(i + 1, j + 1).Locked:
                    out.append(addr(r0 + i, c0 + j))
            except Exception:
                pass
            if len(out) >= limit or (i * nc + j) > PROTECT_SCAN_CELLS:
                return out
    return out


def _unlocked_cells(ws: Any, limit: int = rt.MAX_CANDIDATES) -> List[str]:
    """보호된 시트에서 써도 되는(잠기지 않은) 칸 (사용 범위 안, 앞에서부터)."""
    try:
        used = ws.UsedRange
        r0, c0 = int(used.Row), int(used.Column)
        nr, nc = int(used.Rows.Count), int(used.Columns.Count)
    except Exception:
        return []
    out: List[str] = []
    for i in range(nr):
        for j in range(nc):
            if i * nc + j >= PROTECT_SCAN_CELLS or len(out) >= limit:
                return out
            try:
                if not ws.Cells(r0 + i, c0 + j).Locked:
                    out.append(addr(r0 + i, c0 + j))
            except Exception:
                pass
    return out


def check_protection(ws: Any, ranges: Sequence[Any], need: str = "values") -> None:
    """쓰기 전에 시트 보호를 확인합니다 (excel 쓰기 명령 공통). 막히면 아무것도 바꾸지 않고 STOP_RETRY sheet_protected.
    need: values(값·수식: 잠기지 않은 칸은 씀), objects(메모·차트), _PROTECT_ALLOW의 일(보호 대화상자에서 허용했으면 함),
    structure(병합·이동 등 늘 막힘). 값 쓰기가 잠긴 칸에 막히면 잠기지 않은 칸을 [후보]로 (사용자가 그 칸에 쓰라고 할 수 있음)."""
    if not _is_protected(ws):
        return
    locked: List[str] = []
    if need == "objects":
        try:
            blocked = bool(ws.ProtectDrawingObjects)
        except Exception:
            blocked = False
    elif need in ("values", "sort"):
        for rng in ranges:
            locked += _locked_cells(rng)
        blocked = bool(locked)
        if need == "sort" and not blocked:
            try:
                blocked = not bool(getattr(ws.Protection, _PROTECT_ALLOW["sort"]))
            except Exception:
                blocked = True
    elif need in _PROTECT_ALLOW:
        try:
            blocked = not bool(getattr(ws.Protection, _PROTECT_ALLOW[need]))
        except Exception:
            blocked = True
    else:
        blocked = True
    if not blocked:
        return
    name = str(ws.Name)
    where = f" (잠긴 칸 {len(locked)}개: {', '.join(locked[:5])}{' ...' if len(locked) > 5 else ''})" if locked else ""
    msg = f"시트 '{name}'이(가) 보호되어 있어 {_PROTECT_WHAT.get(need, '바꿀')} 수 없습니다{where}. {PROTECT_HINT}."
    cands = _unlocked_cells(ws) if need == "values" else []
    if cands:
        msg += " 이 시트에서 잠기지 않아 쓸 수 있는 칸:"
    err = rt.UserDecisionError(
        msg,
        "시트가 보호되어 있어 아무것도 바꾸지 않았습니다. message를 사용자에게 그대로 알리세요. 시트 보호를 풀지 말고(암호를 묻거나 "
        "추측하지 않음) 사용자가 [편집 사용]을 눌렀거나 보호를 풀었다고 하면 같은 명령을 1번 다시 실행하세요."
        + (" 사용자가 [후보]의 잠기지 않은 칸에 쓰라고 하면 그 칸으로 실행하세요." if cands else ""),
        candidates=cands)
    err.error_type = "sheet_protected"
    raise err


def _fixed_formula(formula: str) -> str:
    """기계적으로 고칠 수 있는 수식 표기: 전각 글자, 둥근 따옴표, 따옴표 밖의 ';' 인수 구분 -> ','."""
    import unicodedata
    f = unicodedata.normalize("NFKC", formula).replace("“", '"').replace("”", '"')
    out, quoted = [], False
    for ch in f:
        if ch == '"':
            quoted = not quoted
        out.append("," if ch == ";" and not quoted else ch)
    return "".join(out)


def formula_rejected(e: BaseException, ref: str, formula: str) -> rt.InputError:
    """엑셀이 받지 않은 수식(.Formula 대입 COM 오류) -> 어느 수식이 왜 실패했는지와 고칠 규칙 (STOP_RETRY).
    기계적으로 고칠 수 있으면(';' 구분, 전각) 그 수식으로 바꾼 전체 명령을 [추천]. 오류 로그에는 수식 글을 남기지 않음."""
    rules = "엑셀 수식은 영어 함수 이름(SUM, IF, VLOOKUP), 인수 구분은 쉼표(,), 글자는 큰따옴표(\"글자\")로 씁니다"
    head = f"{ref}에 넣을 수식 {formula!r}을(를) 엑셀이 받지 않았습니다 ({type(e).__name__}). 아무것도 바꾸지 않았습니다. {rules}."
    log_head = f"{ref}에 넣을 수식을 엑셀이 받지 않았습니다 ({type(e).__name__}, 수식 {rt.DOC_OMITTED}). {rules}."
    fixed = _fixed_formula(formula)
    argv = list(rt.CURRENT_ARGV or [])
    if fixed != formula and argv and any(formula in a for a in argv):
        new_argv = [a.replace(formula, fixed) for a in argv]
        cmd = rt.next_command("excel.py", new_argv, "")["command"]
        return rt.InputError(head + f" 고친 수식 {fixed!r}로 다시 실행하세요. [추천] {cmd}",
                             log_message=log_head + " [추천 명령 생략: 수식]", no_fallback=True)
    return rt.InputError(head + " 수식을 고쳐 다시 실행하세요.", log_message=log_head, no_fallback=True)


def _restore_matrix(rng: Any) -> Tuple[List[List[Any]], List[List[Any]], List[List[Any]]]:
    """되돌리기용으로 칸의 지금 상태를 읽습니다 -> (다시 넣을 값, 비교용 Formula, 비교용 Value2).
    수식 칸만 수식 글로, 상수 칸은 Value2로 (Formula 글을 다시 넣으면 Excel이 글자 007·1-2·긴 숫자를 숫자·날짜로,
    글자 '=abc를 수식으로 바꿈). 글자 상수는 앞에 '를 붙여 글자 그대로 (엑셀의 글자 표시, 값에 남지 않음)."""
    forms = off.range_to_matrix(rng.Formula)
    vals = off.range_to_matrix(rng.Value2)
    has = rng.HasFormula                                    # True 모두 수식, False 없음, None 섞임 (글자 '=abc도 Formula는 =abc)
    out: List[List[Any]] = []
    for i, row in enumerate(forms):
        line: List[Any] = []
        for j, f in enumerate(row):
            v = vals[i][j] if i < len(vals) and j < len(vals[i]) else None
            if isinstance(f, str) and f.startswith("=") and (has is True or (has is None and rng.Cells(i + 1, j + 1).HasFormula)):
                line.append(f)                              # 수식
            elif isinstance(f, str) and f.startswith("#") and not isinstance(v, str):
                line.append(f)                              # 오류 상수 (#N/A 등: Value2는 오류 번호)
            elif isinstance(v, str):
                line.append("'" + v if v != "" else "")
            else:
                line.append(v)
        out.append(line)
    return out, forms, vals


def _restore(rng: Any, saved: Tuple[List[List[Any]], List[List[Any]], List[List[Any]]]) -> bool:
    """_restore_matrix로 읽어 둔 상태로 되돌리고 다시 읽어 비교합니다. 같으면 True."""
    data, forms, vals = saved
    if len(data) == 1 and len(data[0]) == 1:
        rng.Formula = data[0][0]
    else:
        rng.Formula = tuple(tuple(row) for row in data)
    _, forms2, vals2 = _restore_matrix(rng)

    def same(a: Any, b: Any) -> bool:
        if a in (None, "") and b in (None, ""):
            return True
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
            return abs(float(a) - float(b)) <= 1e-9 * max(1.0, abs(float(a)))
        return str(a) == str(b) if isinstance(a, str) or isinstance(b, str) else a == b
    return all(same(x, y) for a, b in ((forms, forms2), (vals, vals2))
               for ra, rb in zip(a, b) for x, y in zip(ra, rb)) and len(forms) == len(forms2)


# 동적 배열 함수: Range.Formula로 쓰면 Office 2021에서 암시적 교차(@)로 들어가 펼쳐지지 않음 -> Formula2 (2026-11-09 검토 A9)
DYNAMIC_ARRAY = re.compile(r"\b(?:UNIQUE|FILTER|SORT|SORTBY|SEQUENCE|RANDARRAY|XLOOKUP|XMATCH|TEXTSPLIT|VSTACK|HSTACK|TOCOL|TOROW|"
                           r"TAKE|DROP|CHOOSECOLS|CHOOSEROWS|EXPAND|WRAPROWS|WRAPCOLS|LET|LAMBDA)\s*\(", re.IGNORECASE)


def _set_formula2(rng: Any, formula: Any) -> bool:
    """Formula2로 씀 (Office 2021 이상). 없는 버전(2019)이면 False (호출하는 쪽이 Formula로)."""
    try:
        if not hasattr(rng, "Formula2"):
            return False
        rng.Formula2 = formula
        return True
    except AttributeError:
        return False


def apply_writes(t: Any, items: Sequence[Tuple[Any, int, int, int, int, str, Any]], new_sheet: Any = None) -> None:
    """값·수식을 씁니다 [(시트, r1, c1, r2, c2, value|values|formula, 값)]. 엑셀이 하나라도 받지 않으면(잘못된 수식 등)
    이미 쓴 칸을 원래대로 되돌리고, 이 명령이 만든 새 시트(new_sheet)는 지운 뒤 STOP_RETRY (반쯤 쓴 채로 남기지 않음)."""
    try:
        saved_before = bool(t.doc.Saved)
    except Exception:
        saved_before = None
    done: List[Tuple[Any, Any]] = []
    failure: Optional[rt.InputError] = None
    with off.excel_batch(t.app):
        for ws, a, b, c, d, kind, payload in items:
            rng = ws.Range(ws.Cells(a, b), ws.Cells(c, d))
            try:
                prior = _restore_matrix(rng)    # 되돌리기용: 수식 칸은 수식 글, 상수 칸은 Value2(글자는 ')
            except Exception:
                prior = None
            try:
                if kind == "formula":
                    if not (DYNAMIC_ARRAY.search(str(payload)) and _set_formula2(rng, payload)):
                        rng.Formula = payload
                else:
                    rng.Value = payload
            except Exception as e:
                info = rt._com_error_info(e)
                if info is not None and info[2] in ("busy", "gone"):
                    raise                       # 엑셀이 바쁘거나 꺼짐: 수식 문제가 아님
                text = payload if kind == "formula" else next(
                    (str(v) for row in (payload if kind == "values" else ((payload,),)) for v in row
                     if isinstance(v, str) and v.startswith("=")), None)
                ref = f"{ws.Name}!{addr(a, b, c, d)}"
                failure = formula_rejected(e, ref, text) if text else rt.InputError(
                    f"{ref}에 값을 쓰지 못했습니다 ({type(e).__name__}). 아무것도 바꾸지 않았습니다. 시트 보호·편집 중인 칸·병합된 칸인지 확인하세요.",
                    no_fallback=True)
                break
            done.append((rng, prior))
            t.changed = True
        if failure is not None:
            problems = []
            for rng, prior in reversed(done):
                try:
                    if prior is None:
                        rng.ClearContents()
                    elif not _restore(rng, prior):
                        problems.append(str(rng.Address).replace("$", "") + "(다시 읽은 값이 원래와 다름)")
                except Exception:
                    problems.append(str(rng.Address).replace("$", ""))
            if new_sheet is not None:
                prev = t.app.DisplayAlerts
                try:
                    t.app.DisplayAlerts = False
                    new_sheet.Delete()
                except Exception:
                    problems.append(f"새 시트 {new_sheet.Name}")
                finally:
                    t.app.DisplayAlerts = prev
            if problems:
                failure = rt.InputError(str(failure).replace("아무것도 바꾸지 않았습니다.", "")
                                        + f" 먼저 쓴 칸을 되돌리지 못한 곳: {', '.join(problems[:5])} (office.py undo로 되돌리세요).",
                                        log_message=failure.log_message, no_fallback=True)
            else:
                t.changed = False
                if saved_before is not None:
                    try:
                        t.doc.Saved = saved_before      # 되돌렸으니 저장 상태도 그대로 (되돌리기 지점도 남기지 않음)
                    except Exception:
                        pass
    if failure is not None:
        raise failure


# --- write --cells: 흩어진 칸 여러 개에 한 번에 ---

CELLS_MAX = 500                 # write --cells 한 번에 쓰는 칸 수 상한
CELLS_SHOWN = 30                # 결과 JSON에 담는 칸 목록 상한
_CELLS_VALUE_HEADER = ("값", "value", "내용", "입력값")


def _value_end(text: str, i: int, seps: str) -> int:
    """text[i:]가 '칸=값' 항목의 값일 때 값이 끝나는 자리 (다음 구분자 또는 글 끝).
    1. 값이 따옴표로 시작하고 같은 따옴표로 닫힌 뒤 바로(공백만 두고) 구분자나 글 끝이 오면 그 묶음 안의 구분자는 나누지 않음 ('a;b').
       닫히지 않은 따옴표('007은 엑셀의 글자 표시)는 묶음이 아님: 뒤 항목을 삼키지 않음.
    2. 값이 =로 시작(수식)하면 큰따옴표와 괄호 안의 구분자는 나누지 않음 (=IF(A1="a;b",1,2), =SUM(A1;A2)).
       따옴표·괄호가 끝까지 닫히지 않으면 수식이 아닌 것처럼 다음 구분자에서 나눔.
    3. 그 밖에는 다음 구분자."""
    n = len(text)
    j = i
    while j < n and text[j] in " \t":
        j += 1
    plain = next((k for k in range(i, n) if text[k] in seps), n)
    if j < n and text[j] in "'\"":
        q = text[j]
        k = text.find(q, j + 1)
        while k != -1:
            m = k + 1
            while m < n and text[m] in " \t":
                m += 1
            if m == n or text[m] in seps:
                return m
            k = text.find(q, k + 1)
        return plain
    if j < n and text[j] == "=":
        depth, quoted = 0, False
        for k in range(j + 1, n):
            ch = text[k]
            if ch == '"':
                quoted = not quoted
            elif quoted:
                continue
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth = max(0, depth - 1)
            elif ch in seps and depth == 0:
                return k
        return n if depth == 0 and not quoted else plain
    return plain


def _split_outside_quotes(text: str, seps: str) -> List[str]:
    """'칸=값' 항목들을 seps의 글자로 나눕니다. 값 안에서 나누지 않는 경우는 _value_end.
    칸 쪽('요약'!C3의 시트 따옴표)은 묶음으로 보지 않음."""
    parts: List[str] = []
    i, n = 0, len(text)
    while i <= n:
        eq = next((k for k in range(i, n) if text[k] == "=" or text[k] in seps), n)
        end = _value_end(text, eq + 1, seps) if eq < n and text[eq] == "=" else eq
        parts.append(text[i:end])
        i = end + 1
    return parts


def _unquote(v: str) -> str:
    v = v.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
        return v[1:-1]
    return v


def _is_cells_file(text: str) -> bool:
    t = text.strip().strip('"')
    return "=" not in t and os.path.splitext(t)[1].lower() in TABLE_FILE_EXTS


def _cells_items(spec: str) -> Tuple[List[Tuple[str, str, str]], bool]:
    """--cells 값 -> [(어디(오류 표시용), 칸, 값)], 파일인지. 글: 'C6=ft; 계산기!C11=20' (; 또는 줄바꿈),
    파일(.csv/.tsv/.md/.txt): 열 칸, 값 (첫 행이 칸/값 머리글이면 건너뜀). 값은 글자 그대로 받고 쓸 때 숫자·수식을 가립니다."""
    if _is_cells_file(spec):
        path = rt.resolve_input_path(spec.strip().strip('"'), "칸값 파일", exts=TABLE_FILE_EXTS)
        rows = load_table_file(path, raw=True)
        start = 1
        if rows and str(rows[0][0] if rows[0] else "").strip().lower() in _MEMO_HEADER:
            rows, start = rows[1:], 2
        rt.check_two_columns(path, rows, start, "칸값 파일", "칸, 값", "쓰지 않았습니다")
        items = [(f"{n}번째 줄", str(r[0] if r else "").strip(), "" if len(r) < 2 or r[1] is None else str(r[1]))
                 for n, r in enumerate(rows, start) if any(str(v or "").strip() for v in r)]
        return items, True
    items = []
    for k, part in enumerate(_split_outside_quotes(rt.unescape_newlines(spec), ";\n"), 1):
        if not part.strip():
            continue
        if "=" not in part:
            items.append((f"{k}번째 항목", part.strip(), None))      # 값 없음 -> 오류로 모음
            continue
        ref, val = part.split("=", 1)
        items.append((f"{k}번째 항목", ref.strip(), _unquote(val)))
    return items, False


# 'C6=ft,C7=20'처럼 쉼표로 나눈 항목: 값 안에 ',<칸>=' (시트 붙은 칸 포함)가 있으면 칸을 쉼표로 나눈 것
_COMMA_ITEM = re.compile(r"(?:\s*[,|]\s*|\s+)(?=(?:(?:'[^']+'|[\w가-힣]+)!)?\$?[A-Za-z]{1,3}\$?\d+\s*=)")   # 쉼표·|·공백


def rt_fully_quoted(val: str) -> bool:
    """값 전체가 한 쌍의 따옴표로 묶였는지 ('a,b' 예, 'ft',C7='x' 아니오: 첫 닫는 따옴표가 끝이어야 함)."""
    return len(val) >= 2 and val[0] in "'\"" and val.find(val[0], 1) == len(val) - 1


def _refuse_comma_items(args) -> None:
    """칸 구분은 ; 입니다. 'C6=ft,C7=20'을 그대로 쓰면 C6에 'ft,C7=20'이 들어가므로 쓰기 전에 멈추고 ;로 바꾼 전체 명령을 [추천]
    (값에 쉼표가 들어가는 일은 흔해 쉼표를 구분자로 추측하지 않음, 회사 PC 2026-11-06 O2). 수식(=로 시작)·따옴표로 묶은 값은 제외."""
    def comma_split(part: str) -> bool:
        if "=" not in part:
            return False
        val = part.split("=", 1)[1].strip()
        if val.startswith("=") or rt_fully_quoted(val):
            return False
        m = _COMMA_ITEM.search(val)
        # 첫 ',<칸>=' 앞에 '='가 있으면 값 자체가 'R1=10k,R2=4.7k' 같은 글 (나누지 않음, 1.4.77 검토)
        return bool(m) and "=" not in val[:m.start()]

    parts = [p.strip() for p in _split_outside_quotes(rt.unescape_newlines(args.cells), ";\n") if p.strip()]
    hit = [p.split("=", 1)[0].strip() for p in parts if comma_split(p)]
    if not hit:
        return
    fixed = "; ".join(_COMMA_ITEM.sub("; ", p) if comma_split(p) else p for p in parts)
    cmd = _cells_command(type(args)(**{**vars(args), "cells": fixed}), drop_dry=False)["command"]
    raise rt.InputError(f"--cells의 칸은 ; 로 나눕니다. 쉼표로 나눈 것 같아 쓰지 않았습니다 ({', '.join(hit[:3])}의 값에 '칸=' 이 들어 있음). "
                        f"[추천] {cmd}")


def _cells_command(args, extra: Sequence[str] = (), drop_dry: bool = True) -> Dict[str, Any]:
    argv = ["write", "--target", args.target, "--cells", args.cells]
    if args.sheet:
        argv += ["--sheet", args.sheet]
    for flag, on in (("--overwrite", args.overwrite), ("--no-focus", args.no_focus), ("--dry-run", args.dry_run and not drop_dry)):
        if on:
            argv.append(flag)
    argv += [x for x in extra if x not in argv]
    return rt.next_command("excel.py", argv, "칸 여러 개 쓰기 실행" + (" (값이 있는 칸 덮어쓰기)" if "--overwrite" in argv else ""),
                           destructive="--overwrite" in argv)


def _cells_plan(args) -> Tuple[List[Tuple[str, int, int, int, int, str, Any]], List[str]]:
    """쓸 칸 [(시트 이름 또는 "", r1, c1, r2, c2, 종류 value|formula, 값)]. 하나라도 잘못되면 하나도 쓰지 않음."""
    items, from_file = _cells_items(args.cells)
    if not items:
        raise rt.InputError("--cells에 쓸 칸이 없습니다. 예: --cells \"C6=ft; C7=20\" (많으면 칸값.csv: 열 칸, 값).")
    if not from_file:
        _refuse_comma_items(args)
    bad: List[str] = []
    plan: List[Tuple[str, int, int, int, int, str, Any]] = []
    seen: Dict[Tuple[str, int, int], Any] = {}
    warnings: List[str] = []
    for where, ref, val in items:
        sheet, part = split_sheet(ref)
        sheet = (sheet or args.sheet or "").strip()
        try:
            if val is None:
                raise rt.InputError("'칸=값' 모양이 아닙니다 (예: C6=ft)")
            r1, c1, r2, c2 = parse_area(part, None, None) if part else (None, None, None, None)
            if r1 is None or r2 is None or c2 is None:
                raise rt.InputError("칸 주소가 아닙니다 (예: C6, 계산기!C11)")
        except rt.InputError as e:
            bad.append(f"{where} '{ref}': {str(e).rstrip('.')}")
            continue
        text = val.strip()
        if text.startswith("=") and len(text) > 1:
            kind, payload = "formula", text
        else:
            kind, payload = "value", xl_text(coerce(val))
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                key = (sheet.lower(), r, c)
                if key in seen:
                    if seen[key] != (kind, payload):
                        bad.append(f"{where} '{ref}': 같은 칸({addr(r, c)})에 다른 값이 두 번 있습니다")
                    else:
                        warnings.append(f"{addr(r, c)}가 두 번 있어 한 번만 씀")
                seen[key] = (kind, payload)
        plan.append((sheet, r1, c1, r2, c2, kind, payload))
    if bad:
        src = "칸값 파일" if from_file else "--cells"
        head = f"{src} {len(items)}개 중 {len(bad)}개가 잘못되어 하나도 쓰지 않았습니다 (모양: 칸=값, 칸끼리 ; 로 구분). 고친 뒤 다시 실행하세요. "
        raise rt.InputError(head + " / ".join(bad[:10]),
                            log_message=head + f"잘못된 항목 {len(bad)}개 ({rt.DOC_OMITTED})")
    if len(seen) > CELLS_MAX:
        raise rt.InputError(f"한 번에 {CELLS_MAX}칸까지 쓸 수 있습니다 (받은 칸 {len(seen)}개). 나눠 실행하세요.")
    # 같은 값이 두 번이면 한 번만
    uniq: List[Tuple[str, int, int, int, int, str, Any]] = []
    for p in plan:
        if p not in uniq:
            uniq.append(p)
    return uniq, warnings[:10]


def _cell_same(cell_rng: Any, kind: str, payload: Any) -> bool:
    """칸에 이미 같은 값(수식)이 있는지 (다시 실행해도 덮어쓰기 확인을 묻지 않게)."""
    norm = lambda f: re.sub(r"\s+", "", str(f)).upper()  # noqa: E731
    try:
        if kind == "formula":
            return norm(cell_rng.Formula) == norm(payload)
        vals = off.range_to_matrix(cell_rng.Value)
        forms = off.range_to_matrix(cell_rng.Formula)
    except Exception:
        return False
    for i, row in enumerate(vals):
        for j, v in enumerate(row):
            f = forms[i][j] if i < len(forms) and j < len(forms[i]) else None
            if isinstance(f, str) and f.startswith("="):
                return False
            if not _same_cell(payload, v) or (payload in (None, "") and v not in (None, "")):
                return False
            if isinstance(payload, str) and isinstance(v, (int, float)) and not isinstance(v, bool):
                return False
    return True


def write_cells(args) -> Dict[str, Any]:
    """excel.py write --cells: 흩어진 칸에 값·수식을 한 번에. 덮어쓸 칸은 한 번에 묻고(--overwrite 하나), 되돌리기 지점 하나, 다시 읽어 확인."""
    others = [f for f, v in (("--value", args.value is not None), ("--values", args.values), ("--file", args.file),
                             ("--formula", args.formula), ("--header", args.header), ("--append", args.append),
                             ("--new-sheet", args.new_sheet), ("--range", args.range)) if v]
    if others:
        raise rt.InputError(f"--cells는 {', '.join(others)}와 함께 쓰지 않습니다. 칸마다 '칸=값'으로 적으세요 (예: --cells \"C6=ft; C7=20\").")
    plan, warnings = _cells_plan(args)
    with off.open_target("excel", args.target, readonly=args.dry_run, focus=not args.no_focus,
                         create=not args.dry_run, tool=TOOL, command="write") as t:
        sheets: Dict[str, Any] = {}
        ops: List[Tuple[Any, int, int, int, int, str, Any]] = []
        for name, r1, c1, r2, c2, kind, payload in plan:
            if name.lower() not in sheets:
                sheets[name.lower()] = t.sheet(name or None)       # 없는 시트는 [추천]/[후보] (만들지 않음)
            ops.append((sheets[name.lower()], r1, c1, r2, c2, kind, payload))
        for key, ws in sheets.items():
            check_protection(ws, [ws.Range(ws.Cells(o[1], o[2]), ws.Cells(o[3], o[4])) for o in ops if o[0] is ws], "values")
        conflicts: List[Tuple[str, Any]] = []
        same: List[str] = []
        todo: List[Tuple[Any, int, int, int, int, str, Any]] = []
        planned: List[str] = []
        for ws, r1, c1, r2, c2, kind, payload in ops:
            rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
            ref = f"{ws.Name}!{addr(r1, c1, r2, c2)}"
            shown = payload if kind == "formula" else out_value(payload[1:] if isinstance(payload, str) and payload.startswith("'") else payload)
            if _cell_same(rng, kind, payload):
                same.append(ref)
                planned.append(f"그대로(같은 값): {ref}")
                continue
            filled = _filled_cells(rng)
            if filled:
                conflicts += [(f"{ws.Name}!{c}", v) for c, v in filled]
                planned.append(f"덮어쓰기: {ref} ({_short(filled[0][1])} -> {_short(shown)})")
            else:
                planned.append(f"변경: {ref} (빈 칸 -> {_short(shown)})")
            todo.append((ws, r1, c1, r2, c2, kind, payload))
        if args.dry_run:
            extra: Dict[str, Any] = {"target_resolved": t.resolved()}
            if conflicts:
                extra["overwrite_cells"] = len(conflicts)
            if same:
                extra["unchanged"] = len(same)
            if warnings:
                extra["warnings"] = warnings
            preview = f"미리보기: 칸 {len(todo)}개" + (f" (값이 있어 덮어쓸 칸 {len(conflicts)}개: "
                                                     + _filled_preview(conflicts) + ")" if conflicts else "")
            if not todo:                # 바뀔 것 없음: 실행할 명령을 주지 않음
                preview = f"미리보기: 바뀔 것 없음 (0칸, 이미 같은 값 {len(same)}칸) (바꾸지 않음)"
            return rt.dry_run_result(TOOL, "write", preview, len(todo), planned,
                                     _cells_command(args, ["--overwrite"] if conflicts else []), **extra)
        if conflicts and not args.overwrite:
            head = f"값이 있는 칸 {len(conflicts)}개가 덮어써집니다: "
            raise rt.UserDecisionError(
                head + _filled_preview(conflicts),
                "쓰지 않았습니다(빈 칸도 쓰지 않음). 덮어쓸 칸과 기존 값을 사용자에게 한 번에 보여 주고, 사용자가 덮어써도 된다고 하면 "
                "같은 명령에 --overwrite를 한 번만 붙여 다시 실행하세요 (미리 보려면 --dry-run).",
                log_message=head + _filled_preview(conflicts, log=True))
        if todo:
            apply_writes(t, todo)          # 엑셀이 받지 않는 수식이 있으면 쓴 칸을 되돌리고 STOP_RETRY
        problems: List[str] = []
        for ws, r1, c1, r2, c2, kind, payload in todo:          # R1: 다시 읽어 확인 (시트마다)
            problems += [f"{ws.Name}!{p}" for p in _verify_written(ws, [(r1, c1, r2, c2, kind, payload)])]
        if todo:
            t.set_focus(sheet=todo[0][0])
        resolved = t.resolved()
        # 시트 이름은 세션 안에서: 닫힌 파일은 세션이 끝나면 문서를 닫아 ws.Name이 OLE 0x800a01a8 (회사 PC 2026-11-06 O2)
        written = [f"{ws.Name}!{addr(r1, c1, r2, c2)}" for ws, r1, c1, r2, c2, _, _ in todo]
    fields = {"target_resolved": resolved, "count": len(todo), "cells": written[:CELLS_SHOWN],
              "truncated": len(written) > CELLS_SHOWN, "applied": [p for p in planned if not p.startswith("그대로")][:CELLS_SHOWN]}
    if same:
        fields["unchanged"] = len(same)
    if warnings:
        fields["warnings"] = warnings
    if conflicts:
        fields["overwritten"] = len(conflicts)
    if not todo:
        return rt.ok(TOOL, "write", f"칸 {len(same)}개가 이미 같은 값이라 바꾸지 않았습니다", "VERIFIED: 모든 칸이 이미 같은 값입니다(바뀐 것 없음). 다시 실행하지 마세요.",
                     verified=True, **fields)
    preview = f"칸 {len(todo)}개에 썼습니다: {', '.join(written[:5])}" + (" ..." if len(written) > 5 else "") \
        + (f" (덮어쓴 칸 {len(conflicts)}개)" if conflicts else "") + (f" (이미 같은 값 {len(same)}칸은 그대로)" if same else "")
    if problems:
        return rt.ok_unverified(TOOL, "write", preview, problems[:10], **fields)
    return rt.ok(TOOL, "write", preview, "VERIFIED: 쓴 칸을 모두 다시 읽어 확인했습니다. 잘못 썼으면 office.py undo로 되돌릴 수 있습니다. "
                 "같은 명령을 다시 실행하지 마세요.", verified=True, **fields)


def cmd_write(args) -> Dict[str, Any]:
    if args.cells:
        return write_cells(args)
    if args.dry_run:
        raise rt.dry_run_unsupported(TOOL, "write --value·--values·--file·--formula (미리보기는 --cells에만 있음)")
    sources = [x for x in (args.value is not None, args.values, args.file, args.formula) if x]
    if not sources:
        raise rt.InputError("쓸 내용이 없습니다. --value, --values, --file, --formula, --cells 중 하나를 지정하세요.")
    if len(sources) > 1:
        raise rt.InputError("--value, --values, --file, --formula는 하나만 지정하세요 (--header는 함께 사용 가능).")
    table: Optional[List[List[Any]]] = None
    if args.values:
        table = parse_values_arg(args.values)
    elif args.file:
        table = load_table_file(rt.resolve_input_path(args.file, "입력 표 파일", exts=TABLE_FILE_EXTS))
        if not table:
            raise rt.InputError(f"입력 파일에 데이터가 없습니다: {args.file}")
        head = [str(v or "").strip().lower() for v in table[0][:3]]
        if len(head) == 2 and head[0] in _MEMO_HEADER and head[1] in _CELLS_VALUE_HEADER and not args.range:
            # 머리글이 '칸, 값'이면 표가 아니라 칸마다 쓸 값 목록 (표로 쓰면 A1부터 머리글까지 그대로 들어감)
            argv = ["write", "--target", args.target, "--cells", args.file] + (["--sheet", args.sheet] if args.sheet else [])
            raise rt.InputError("이 파일은 머리글이 '칸, 값'인 칸별 값 목록입니다. 표가 아니라 칸마다 쓰려면 --cells를 씁니다. "
                                f"[추천] {rt.next_command('excel.py', argv, '')['command']}")
    header = [h.strip() for h in args.header.split(",")] if args.header else None
    start = args.range or "A1"
    sh, cells = split_sheet(start)
    parts = cells.replace("$", "").split(":")
    r1, c1 = parse_cell(parts[0])
    r2c2 = parse_cell(parts[1]) if len(parts) > 1 else None
    if r2c2 is not None:     # 뒤집힌 범위(C5:A1)는 Excel처럼 A1:C5로 ('-3행 x -1열'로 보고하고 확인하지 않던 것, 2026-11-09 검토 B13)
        r1, c1, r2c2 = min(r1, r2c2[0]), min(c1, r2c2[1]), (max(r1, r2c2[0]), max(c1, r2c2[1]))

    with write_session(args, "write") as t:
        # 시트·칸 확인과 쓸 내용 계획을 모두 마친 뒤에만 문서를 바꿈 (새 시트도 쓰기 직전에 만듦)
        ws = _write_sheet(t, sh or args.sheet, args.new_sheet)   # --new-sheet면 None (아직 만들지 않음, 빈 시트로 봄)
        if args.append and not args.range and ws is not None:
            # 범위 없이 --append: 표가 A열에서 시작하지 않으면(B3:C5) A6에 머리글과 함께 쓰던 것 -> 사용 범위의 첫 칸 기준 (검토 B13)
            try:
                used = ws.UsedRange
                if int(used.Rows.Count) * int(used.Columns.Count) > 1 or not _cell_empty(used):
                    r1, c1 = int(used.Row), int(used.Column)
            except Exception:
                pass
        authored_row = r1 + (1 if header else 0)                 # 사용자가 준 수식이 가리키는 첫 수식 행 (--append 전)

        if args.append and ws is not None:
            if table:
                body_w = len(rect(table)[0])
            elif r2c2:
                body_w = r2c2[1] - c1 + 1
            elif args.formula:
                body_w = len(args.formula)
            else:
                body_w = 1
            # 수식·값도 쓸 열 수만큼 (1열만 보면 다른 열의 데이터 위에 씀)
            width = max(len(header) if header else 0, body_w)
            # 쓸 열 전체에서 가장 아래 행 (한 열만 보면 그 열이 비어 있는 행에 다른 열 데이터를 덮어씀)
            last = max(_last_used_row(ws, c) for c in range(c1, c1 + width))
            if last >= r1:
                # 같은 머리글은 반복하지 않음: 표의 머리글 행(시작 행)과 비교 (--header도 같음)
                existing = [str(out_value(v)) for v in off.range_to_matrix(
                    ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c1 + width - 1)).Value)[0]]
                if header and existing[:len(header)] == [str(h) for h in header]:
                    header = None
                if table and existing[:len(table[0])] == [str(out_value(x)) for x in table[0]]:
                    table = table[1:]
                    if not table:
                        raise rt.InputError("추가할 데이터 행이 없습니다 (입력 파일에 헤더만 있음).")
                if r2c2:
                    r2c2 = (r2c2[0] - r1 + last + 1, r2c2[1])
                r1 = last + 1

        # 1) plan every block and check it BEFORE writing anything (a failed check leaves the sheet untouched)
        ops: List[Tuple[int, int, int, int, str, Any]] = []      # (r1, c1, r2, c2, kind, payload)
        top = r1
        if header:
            ops.append((r1, c1, r1, c1 + len(header) - 1, "values", (tuple(xl_text(h) for h in header),)))
            r1 += 1
        if args.formula:
            formulas = [f if f.startswith("=") else "=" + f for f in args.formula]
            if all(_is_number(f[1:]) for f in formulas):
                nums = [float(f[1:].replace(",", "")) for f in formulas]
                one_col = r2c2 is not None and r2c2[1] == c1
                rows = [[x] for x in nums] if one_col else [nums]
                opt, val = ("--value", formulas[0][1:]) if len(nums) == 1 else ("--values", json.dumps(rows))
                rec = rt.retry_command({"--formula": None, opt: val}, multi=("--formula",))
                raise rt.InputError(f"--formula에 숫자만 있습니다. 값은 {opt}로 쓰세요. [추천] {rec}")
            if r2c2 is None:
                if len(formulas) != 1:
                    # 옆 열(왼쪽) 표의 마지막 행까지 (전에는 100행을 지어내 데이터 2행 표에 수식 200개를 썼음, 2026-11-09 검토 B4)
                    data_end = _last_used_row(ws, c1 - 1) if c1 > 1 else 0
                    first = top + 1 if header else top
                    head = ("수식이 여러 개면 --range로 범위를 지정하세요"
                            + (f" (머리글이 있으면 머리글 행부터, 수식은 {top + 1}행 기준으로)" if header else ""))
                    if data_end < first:
                        raise rt.InputError(head + ". 옆 열에 데이터가 없어 몇 행까지 쓸지 정할 수 없습니다. 사용자에게 마지막 행을 물어 "
                                            f"--range {addr(top, c1)}:<끝 열><끝 행> 꼴로 다시 실행하세요.")
                    rec = addr(top, c1, data_end, c1 + len(formulas) - 1)
                    raise rt.InputError(head + f". 옆 열 데이터가 {data_end}행까지라 그 범위로. [추천] {rt.retry_command({'--range': rec})}")
                r2, c2 = r1, c1
            else:
                r2, c2 = r2c2
            if r2 < r1:
                raise rt.InputError("--range 범위가 헤더 행만 포함합니다. 데이터 행까지 포함하세요.")
            ncols = c2 - c1 + 1
            if r1 != authored_row:
                # --append로 수식 행이 옮겨짐: 상대 참조를 실제 행 기준으로 (그대로 쓰면 원래 행을 계속 가리킴)
                formulas = [_rebase_formula(t.app, ws, f, authored_row, c1 + j, r1, c1 + j) for j, f in enumerate(formulas)]
            if len(formulas) == 1:
                ops.append((r1, c1, r2, c2, "formula", formulas[0]))          # 상대참조 자동 조정
            elif len(formulas) == ncols:
                ops += [(r1, c1 + j, r2, c1 + j, "formula", f) for j, f in enumerate(formulas)]
            else:
                raise rt.InputError(f"수식 {len(formulas)}개와 범위의 열 수 {ncols}가 다릅니다. 열마다 수식 1개씩 주세요. "
                                    f"[추천] {rt.retry_command({'--range': addr(top, c1, r2, c1 + len(formulas) - 1)})}")
        elif args.value is not None:
            v = xl_text(coerce(args.value))
            ops.append((r1, c1, *(r2c2 or (r1, c1)), "value", v))
        else:
            data = rect(table)
            nrows, ncols = len(data), len(data[0])
            ops.append((r1, c1, r1 + nrows - 1, c1 + ncols - 1, "values", tuple(tuple(xl_text(v) for v in r) for r in data)))

        if ws is not None:          # 보호된 시트: 아무것도 바꾸기 전에 (잠긴 칸이면 STOP_RETRY, 잠기지 않은 칸 [후보])
            check_protection(ws, [ws.Range(ws.Cells(a, b), ws.Cells(c, d)) for (a, b, c, d, _, _) in ops], "values")

        # 2) never overwrite existing values without the user's consent (--append도 쓸 칸을 확인)
        if not args.overwrite and ws is not None:
            filled = [fc for (a, b, c, d, _, _) in ops for fc in _filled_cells(ws.Range(ws.Cells(a, b), ws.Cells(c, d)))]
            if filled:
                area = addr(ops[0][0], ops[0][1], ops[-1][2], max(o[3] for o in ops))
                head = f"{area}에 값이 있는 칸 {len(filled)}개가 덮어써집니다: "
                raise rt.UserDecisionError(
                    f"{ws.Name}!" + head + _filled_preview(filled),
                    "쓰지 않았습니다. 덮어쓸 칸을 사용자에게 보여 주고, 사용자가 덮어써도 된다고 하면 --overwrite를 붙여 다시 실행하세요. "
                    "아니면 비어 있는 다른 --range를 쓰거나 --append로 아래에 이어 쓰세요.",
                    log_message=head + _filled_preview(filled, log=True))

        # 3) write (새 시트는 여기서 처음 만듦). 엑셀이 받지 않는 수식이 있으면 쓴 칸과 새 시트를 되돌리고 STOP_RETRY
        new_ws = None
        if ws is None:
            ws = new_ws = t.sheet(sh or args.sheet, create=True, allow_create=True)
        apply_writes(t, [(ws, *op) for op in ops], new_sheet=new_ws)
        written: List[str] = [addr(a, b, c, d) for a, b, c, d, _, _ in ops]

        problems = _verify_written(ws, ops)
        first = parse_cell(written[0].split(":")[0])
        last_ref = written[-1].split(":")[-1]
        last = parse_cell(last_ref)
        full = addr(first[0], first[1], last[0], last[1])
        if not args.no_autofit and (table is not None or header):
            try:
                ws.Range(ws.Cells(first[0], first[1]), ws.Cells(last[0], last[1])).Columns.AutoFit()
            except Exception:
                pass
        t.set_focus(sheet=ws)
        sheet_name = str(ws.Name)
        resolved = t.resolved()
        if args.target.lower() != "active" and t.source in ("file", "new"):
            tgt = t.path
        else:
            tgt = args.target

    sheet_ref = sheet_name if " " not in sheet_name else f"'{sheet_name}'"
    result = {"target_resolved": resolved, "sheet": sheet_name, "range_written": full,
              "rows": last[0] - first[0] + 1, "cols": last[1] - first[1] + 1}
    if table is not None or args.formula:
        # 차트 만들기는 문서를 바꿈: confirm_commands (1.4.75)
        result["confirm_commands"] = [rt.next_command("excel.py", ["chart", "--target", tgt, "--range", f"{sheet_ref}!{full}",
                                                                   "--type", "line"],
                                                      "방금 쓴 범위로 차트 만들기, 사용자가 원할 때만 (제목은 --title \"제목\"을 붙임, --type은 원하는 종류로)")]
    preview = f"{sheet_name}!{full}에 썼습니다 ({result['rows']}행 x {result['cols']}열)"
    if problems:
        return rt.ok_unverified(TOOL, "write", preview, problems, **result)
    return rt.ok(TOOL, "write", preview, "VERIFIED: 쓴 칸을 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요."
                 + (" 사용자가 차트를 원하면 confirm_commands[0]을 1번 실행하세요." if "confirm_commands" in result else ""),
                 verified=True, **result)


# ==============================================================================
# 5. replace
# ==============================================================================

def _load_map(args) -> Dict[str, str]:
    if args.map:
        raw = args.map.strip()
        if not raw.startswith("{"):
            raw = rt.read_text_file(rt.resolve_input_path(raw, "치환 목록 파일"))
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as e:
            raise rt.InputError(f"--map은 JSON 객체여야 합니다 (예: {{\"초안\": \"최종\"}}): {e}")
        if not isinstance(data, dict) or not data:
            raise rt.InputError("--map은 비어 있지 않은 JSON 객체여야 합니다.")
        mapping = {str(k): ("" if v is None else str(v)) for k, v in data.items()}
    elif args.find is not None:
        mapping = {args.find: args.replace if args.replace is not None else args.find}
    else:
        raise rt.InputError("--find 또는 --map을 지정하세요.")
    if "" in mapping:
        # 빈 글을 찾으면 str.replace가 모든 글자 사이에 바꿀 글을 끼워 넣어 값이 있는 칸을 모두 망가뜨림
        raise rt.InputError("찾을 글(--find 또는 --map의 키)이 비어 있어 바꾸지 않았습니다. 빈 칸을 채우려면 excel.py write를 쓰세요 "
                            "(예: excel.py write --range B2:B10 --value 0). 바꿀 글이 있으면 찾을 글을 정확히 지정하세요.")
    return mapping


def cmd_replace(args) -> Dict[str, Any]:
    mapping = _load_map(args)
    fmt = _format_spec(args)
    if not args.map and args.replace is None and not fmt:
        # 바꿀 글도 서식도 없으면 아무것도 바뀌지 않는데 'N개 셀 처리'로 보고되던 문제
        raise rt.InputError("--replace(바꿀 글)나 서식 옵션(--bg, --color, --bold 등)이 없어 바꿀 것이 없습니다. "
                            "사용자에게 무엇으로 바꿀지 확인하세요. 찾기만 하려면 excel.py read --search를 씁니다.")
    if args.all_sheets and args.range and split_sheet(args.range)[0]:
        raise rt.InputError("--all-sheets와 시트 이름이 든 --range는 함께 쓸 수 없습니다 (같은 시트를 시트 수만큼 바꾸게 됨). "
                            "한 시트만 바꾸려면 --all-sheets를 빼고, 모든 시트를 바꾸려면 --range에서 시트 이름을 빼세요.")
    replaced: List[str] = []           # 실제로 바꾼 칸 (값이 바뀌었거나 서식을 적용)
    same_cells = 0                     # 찾았지만 바꿀 글이 같아 바뀐 것이 없는 칸 (서식도 없음)
    skipped_formulas: List[str] = []
    written: List[Tuple[Any, int, int, Any, str]] = []        # (시트, 행, 열, 쓴 값, 주소): 다시 읽어 확인 (R1)
    with write_session(args, "replace") as t:
        sheets = [t.doc.Worksheets(i) for i in range(1, t.doc.Worksheets.Count + 1)] if args.all_sheets else [t.sheet(args.sheet)]
        ranges = [target_range(t, ws, args.range) for ws in sheets]     # 주소·시트 확인을 바꾸기 전에 모두
        plans = []
        for rng in ranges:
            ws = rng.Worksheet               # --range 'S2!A1'은 읽기와 쓰기 모두 그 시트
            values = off.range_to_matrix(rng.Value)
            formulas = off.range_to_matrix(rng.Formula)
            r0, c0 = int(rng.Row), int(rng.Column)
            hits = []
            for i, row in enumerate(values):
                for j, v in enumerate(row):
                    if v is None:
                        continue
                    s = str(out_value(v))
                    new = None
                    for find, rep in mapping.items():
                        if (args.whole and s == find) or (not args.whole and find in s):
                            new = rep if args.whole else (new or s).replace(find, rep)
                    if new is not None:
                        hits.append((i, j, v, s, new))
            # 보호된 시트: 바꿀 칸이 잠겼으면(값) / 서식이 허용되지 않으면 아무것도 바꾸기 전에 멈춤
            value_cells = [ws.Cells(r0 + i, c0 + j) for i, j, v, s_, new in hits
                           if new != s_ and not (isinstance(formulas[i][j], str) and formulas[i][j].startswith("="))]
            if value_cells:
                check_protection(ws, value_cells, "values")
            if fmt and hits:
                check_protection(ws, [ws.Cells(r0 + i, c0 + j) for i, j, *_ in hits], "format_cells")
            plans.append((rng, ws, formulas, r0, c0, hits))
        if args.dry_run:
            return _replace_dry(args, fmt, plans, t.resolved())
        with off.excel_batch(t.app):
            for rng, ws, formulas, r0, c0, hits in plans:
                for i, j, v, s, new in hits:
                    cell_ref = f"{ws.Name}!{addr(r0 + i, c0 + j)}"
                    f = formulas[i][j]
                    cell = ws.Cells(r0 + i, c0 + j)
                    did = False
                    is_formula = isinstance(f, str) and f.startswith("=")
                    if is_formula:
                        if new != s:
                            skipped_formulas.append(cell_ref)      # 수식은 보존, 서식만
                    elif new != s:
                        val = xl_text(coerce(new) if not isinstance(v, str) else new)
                        cell.Value = val
                        written.append((ws, r0 + i, c0 + j, val, cell_ref))
                        did = True
                    if fmt:
                        _apply_format(cell, fmt)
                        did = True
                    if did:
                        replaced.append(cell_ref)
                    elif not is_formula:
                        # 찾는 글과 바꿀 글이 같고(--map {"A": "A"} 포함) 서식도 없으면 바뀐 것 없음:
                        # 바꾼 칸으로 세지 않음 (되돌리기 지점도 남기지 않음)
                        same_cells += 1
                if replaced:
                    t.set_focus(sheet=ws)
        problems = [ref for ws, r, c, val, ref in written if not _same_cell(val, ws.Cells(r, c).Value)]
        resolved = t.resolved()
    result = {"target_resolved": resolved, "changed_cells": len(replaced), "cells": replaced[:30]}
    if skipped_formulas:
        result["formula_cells_kept"] = skipped_formulas[:30]
    if same_cells:
        result["unchanged_same_text"] = same_cells
    if not replaced and (same_cells or skipped_formulas):
        what = []
        if same_cells:
            what.append(f"{same_cells}칸은 바꿀 글이 찾는 글과 같음")
        if skipped_formulas:
            what.append(f"수식 칸 {len(skipped_formulas)}개는 내용 유지")
        return rt.ok(TOOL, "replace", f"찾은 칸은 있지만 바꾼 칸이 없습니다 ({', '.join(what)}, 0칸 바뀜)",
                     f"VERIFIED: 바꾼 칸 0개 ({', '.join(what)}). 바뀐 것이 없으니 다시 실행하지 말고 사용자에게 알리세요.", **result)
    if not replaced:
        return rt.ok(TOOL, "replace", "일치하는 셀이 없습니다",
                     "VERIFIED: 일치하는 셀이 없습니다. 검색어를 임의로 바꿔 재시도하지 말고 사용자에게 알리세요.", **result)
    preview = f"{len(replaced)}개 셀 처리" + (f" (수식 셀 {len(skipped_formulas)}개는 내용 유지"
                                             + (", 서식만 적용)" if fmt else ")") if skipped_formulas else "")
    if problems:
        return rt.ok_unverified(TOOL, "replace", preview, [f"다시 읽은 값이 다른 칸 {len(problems)}개: {', '.join(problems[:5])}"], **result)
    return rt.ok(TOOL, "replace", preview, "VERIFIED: 바꾼 칸을 다시 읽어 확인했습니다.", verified=True, **result)


def _replace_dry(args, fmt: Dict[str, Any], plans: List[Any], resolved: Dict[str, Any]) -> Dict[str, Any]:
    """replace --dry-run: 바꿀 칸과 바뀔 글만 보여 줌 (바꾸지 않음, 되돌리기 지점 없음). 세는 규칙은 실제 실행과 같음."""
    planned: List[str] = []
    kept_cells: List[str] = []
    same = 0
    for rng, ws, formulas, r0, c0, hits in plans:
        for i, j, v, s, new in hits:
            ref = f"{ws.Name}!{addr(r0 + i, c0 + j)}"
            f = formulas[i][j]
            if isinstance(f, str) and f.startswith("="):
                if new != s:
                    kept_cells.append(ref)
                if fmt:
                    planned.append(f"서식만(수식 유지): {ref}")
            elif new != s:
                planned.append(f"변경: {ref} ({_short(s)} -> {_short(new)})" + (" + 서식" if fmt else ""))
            elif fmt:
                planned.append(f"서식만: {ref}")
            else:
                same += 1
    n = len(planned)
    extra: Dict[str, Any] = {}
    notes: List[str] = []
    kept = len(kept_cells)
    if kept:
        extra["formula_cells_kept"] = kept_cells[:30]          # 실제 실행과 같은 형식 (칸 주소 목록, 1.4.83)
        notes.append(f"수식 칸 {kept}개는 내용 유지")
    if same:
        extra["unchanged_same_text"] = same
        notes.append(f"{same}칸은 바꿀 글이 찾는 글과 같음")
    note = f", {', '.join(notes)}" if notes else ""
    if n:
        preview = f"미리보기: 칸 {n}개가 바뀝니다{note}: " + ", ".join(planned[:3]) + (" ..." if n > 3 else "") + " (바꾸지 않음)"
    else:
        preview = f"미리보기: 바뀔 것 없음 (0칸{note}) (바꾸지 않음)"
    return rt.dry_run_result(TOOL, "replace", preview, n, planned, rt.dry_run_command(TOOL, "찾아 바꾸기 실행"), target_resolved=resolved, **extra)


# ==============================================================================
# 6. style
# ==============================================================================

def _format_spec(args) -> Dict[str, Any]:
    spec: Dict[str, Any] = {}
    if getattr(args, "bg", None):
        spec["bg"] = off.parse_color(args.bg)
    if getattr(args, "color", None):
        spec["color"] = off.parse_color(args.color)
    for key in ("bold", "italic"):
        if getattr(args, key, None) is not None:
            spec[key] = getattr(args, key)
    if getattr(args, "size", None):
        spec["size"] = float(args.size)
    if getattr(args, "align", None):
        spec["align"] = XL_ALIGN[args.align]
    if getattr(args, "number_format", None):
        spec["number_format"] = args.number_format
    if getattr(args, "border", None):
        spec["border"] = args.border
    if getattr(args, "wrap", None):
        spec["wrap"] = True
    return spec


def _apply_format(rng: Any, spec: Dict[str, Any]) -> None:
    if "bg" in spec:
        rng.Interior.Color = spec["bg"]
    if "color" in spec:
        rng.Font.Color = spec["color"]
    if "bold" in spec:
        rng.Font.Bold = spec["bold"]
    if "italic" in spec:
        rng.Font.Italic = spec["italic"]
    if "size" in spec:
        rng.Font.Size = spec["size"]
    if "align" in spec:
        rng.HorizontalAlignment = spec["align"]
    if "number_format" in spec:
        rng.NumberFormat = spec["number_format"]
    if "wrap" in spec:
        rng.WrapText = True
    if "border" in spec:
        _apply_border(rng, spec["border"])


def _apply_border(rng: Any, weight: str, box: bool = False) -> None:
    if weight == "none":
        for b in (7, 8, 9, 10, 11, 12):
            rng.Borders(b).LineStyle = -4142   # xlNone
        return
    w = XL_WEIGHT[weight]
    for b in ((7, 8, 9, 10) if box else (7, 8, 9, 10, 11, 12)):   # 좌/상/하/우(+내부 세로/가로)
        try:
            border = rng.Borders(b)
            border.LineStyle = 1      # 실선 먼저, 두께는 나중 (순서가 바뀌면 두께가 초기화될 수 있음)
            border.Weight = w
        except Exception:
            pass


def _verify_border(rng: Any, weight: str, box: bool = False) -> List[str]:
    """R1: 테두리 선(좌/상/하/우, 격자면 내부 세로/가로까지)의 선 모양과 굵기를 다시 읽어 비교합니다."""
    edges = (7, 8, 9, 10) if box and weight != "none" else (7, 8, 9, 10, 11, 12)
    bad = []
    for b in edges:
        if b in (11, 12) and weight != "none" and (int(rng.Rows.Count) if b == 12 else int(rng.Columns.Count)) < 2:
            continue                                     # 한 줄·한 칸 범위에는 내부 선이 없음
        try:
            border = rng.Borders(b)
            style, w = border.LineStyle, border.Weight if weight != "none" else None
        except Exception:
            style, w = "?", None
        if weight == "none":
            if style not in (-4142, None):
                bad.append(f"{b}={style!r}")
        elif style != 1 or w is None or int(w) != XL_WEIGHT[weight]:
            bad.append(f"{b}={style!r}/{w!r}")
    if bad:
        return [f"테두리를 다시 읽은 값이 다름 (선 번호=모양/굵기: {', '.join(bad)})"]
    return []


def _apply_theme(ws: Any, rng: Any, theme: Dict[str, Any], zebra: bool = True) -> Dict[str, Any]:
    """표 범위에 테마 적용: 헤더, 줄무늬, 숫자 서식, 테두리, 합계 행."""
    r0, c0 = int(rng.Row), int(rng.Column)
    nrows, ncols = int(rng.Rows.Count), int(rng.Columns.Count)
    values = off.range_to_matrix(rng.Value)
    header = ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + ncols - 1))
    header.Interior.Color = off.parse_color(theme["header_bg"])
    header.Font.Color = off.parse_color(theme["header_fg"])
    header.Font.Bold = True
    header.HorizontalAlignment = XL_ALIGN["center"]

    summary = nrows >= 3 and any(k in str(values[-1][0] or "") for k in ("합계", "총계", "소계", "평균", "Total"))
    data_last = r0 + nrows - (2 if summary else 1)
    stripes = 0
    if zebra and data_last > r0 + 1:
        color = off.parse_color(theme["zebra"])
        block: List[str] = []
        for r in range(r0 + 2, data_last + 1, 2):   # 여러 행을 한 번에 칠해 COM 호출 수 축소 (주소는 255자 이내)
            part = addr(r, c0, r, c0 + ncols - 1)
            if block and len(",".join(block + [part])) > 240:
                ws.Range(",".join(block)).Interior.Color = color
                block = []
            block.append(part)
            stripes += 1
        if block:
            ws.Range(",".join(block)).Interior.Color = color

    numeric_cols = []
    for j in range(ncols):
        col_vals = [values[i][j] for i in range(1, len(values)) if values[i][j] is not None]
        if col_vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in col_vals):
            numeric_cols.append(j)
            ws.Range(ws.Cells(r0 + 1, c0 + j), ws.Cells(r0 + nrows - 1, c0 + j)).NumberFormat = theme["number_format"]

    _apply_border(rng, "thin")
    header.Borders(9).Weight = XL_WEIGHT["medium"]
    if summary:
        srow = ws.Range(ws.Cells(r0 + nrows - 1, c0), ws.Cells(r0 + nrows - 1, c0 + ncols - 1))
        srow.Interior.Color = off.parse_color(theme["summary_bg"])
        srow.Font.Bold = True
        srow.Borders(9).LineStyle = -4119      # 이중선
    try:
        rng.Columns.AutoFit()
    except Exception:
        pass
    return {"theme": theme["name"], "zebra_rows": stripes, "number_columns": [col_letter(c0 + j) for j in numeric_cols],
            "summary_row": summary}


def _parse_highlight(rule: str) -> Tuple[str, str, int]:
    """'>=100:yellow' -> ('>=', '100', BGR). '=완료:green', '포함:지연:red'(텍스트 포함)."""
    if ":" not in rule:
        raise rt.InputError(f"--highlight 형식: '조건:색' (예: '>=100:yellow', '=완료:lightgreen', '포함:지연:pink'). 입력: {rule}")
    cond, color = rule.rsplit(":", 1)
    cond = cond.strip()
    if cond.startswith("포함:"):
        return "contains", cond[3:], off.parse_color(color)
    for op in (">=", "<=", "<>", "!=", ">", "<", "="):
        if cond.startswith(op):
            return op, cond[len(op):].strip(), off.parse_color(color)
    return "contains", cond, off.parse_color(color)


STYLE_AREA_CHARS = 240     # Excel의 Range("B2,D5,...") 주소 글자 한도(255) 아래로 나눠 적용


def _style_areas(address: Optional[str]) -> Tuple[Optional[str], List[str]]:
    """--range "B2;D5;F7"(칸 여러 개는 ;, 사용법 단순성 3번) -> (시트 접두, 칸 목록). 칸 하나면 목록이 하나.
    ','로 이으면 받지 않고 ';'로 고친 전체 명령을 [추천] (1.4.91, 행마다 다른 칸을 칠할 때 칸마다 명령을 따로 실행하던 것)."""
    if not address:
        return None, []
    sheet, cells = split_sheet(address)
    prefix = (address[:len(address) - len(cells)]) if sheet else ""
    if "," in cells:
        fixed = prefix + ";".join(p.strip() for p in cells.split(",") if p.strip())
        raise rt.InputError(f"--range의 칸 여러 개는 ';'로 나눕니다: '{address}'. [추천] {rt.retry_command({'--range': fixed})}",
                            no_fallback=True)
    parts = [p.strip() for p in cells.split(";") if p.strip()]
    return prefix, parts


def _style_chunks(prefix: str, parts: List[str]) -> List[str]:
    """칸 목록 -> Range 주소 묶음 ('Sheet1'!B2,D5 ...), 묶음마다 STYLE_AREA_CHARS 글자 이하."""
    out: List[List[str]] = []
    for p in parts:
        if out and len(",".join(out[-1] + [p])) <= STYLE_AREA_CHARS:
            out[-1].append(p)
        else:
            out.append([p])
    return [prefix + ",".join(c) for c in out]


def cmd_style(args) -> Dict[str, Any]:
    if args.dry_run:
        raise rt.dry_run_unsupported(TOOL, "style")
    fmt = _format_spec(args)
    theme = off.get_theme(args.theme)
    if not fmt and not theme and not args.highlight and not args.clear:
        raise rt.InputError("적용할 서식이 없습니다. --theme, --bg, --color, --bold, --number-format, --border, --highlight 등을 지정하세요.")
    # 모든 입력 확인은 문서를 열기 전에 (서식을 지운 뒤 --highlight 오류로 멈추면 지운 서식이 되돌리기 지점 없이 남았음)
    rule = _parse_highlight(args.highlight) if args.highlight else None
    prefix, parts = _style_areas(args.range)
    if len(parts) > 1 and theme:
        raise rt.InputError("--theme(표 모양)은 범위 하나에만 씁니다. 떨어진 칸 여러 개(--range \"B2;D5\")에는 --bg, --color, --bold 같은 "
                            "서식을 씁니다.", no_fallback=True)
    if len(parts) > 1:
        return _style_many(args, fmt, rule, prefix, parts)
    with write_session(args, "style") as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range)
        ws = rng.Worksheet
        check_protection(ws, [rng], "format_cells")      # 보호된 시트에서 서식이 허용되지 않으면 바꾸기 전에 멈춤
        applied: Dict[str, Any] = {}
        with off.excel_batch(t.app):
            if args.clear:
                rng.ClearFormats()
                rng.FormatConditions.Delete()
                applied["cleared"] = True
            if theme:
                applied.update(_apply_theme(ws, rng, theme, zebra=not args.no_zebra))
            if fmt:
                _apply_format(rng, fmt)
                applied["format"] = sorted(fmt)
            if rule is not None:
                op, val, color = rule
                if op == "contains":
                    top_left = str(rng.Cells(1, 1).Address).replace("$", "")
                    # SEARCH는 ~ ? *를 와일드카드로 읽음: ~로 글자 그대로. 수식 글자 안의 "는 ""로
                    lit = re.sub(r"([~?*])", r"~\1", val).replace('"', '""')
                    fc = rng.FormatConditions.Add(Type=2, Formula1=f'=ISNUMBER(SEARCH("{lit}",{top_left}))')
                else:
                    is_num = bool(re.fullmatch(r"[+-]?\d+(\.\d+)?", val))
                    formula = val if is_num else '="' + val.replace('"', '""') + '"'
                    fc = rng.FormatConditions.Add(Type=1, Operator=XL_OPERATOR[op], Formula1=f"={formula}" if is_num else formula)
                fc.Interior.Color = color
                applied["highlight"] = args.highlight
        problems = _verify_style(ws, rng, fmt, theme, bool(args.highlight))
        t.set_focus(sheet=ws)
        info = range_info(rng)
        resolved = t.resolved()
    preview = f"{info['sheet']}!{info['range']}에 서식 적용: {', '.join(map(str, applied))}"
    if problems:
        return rt.ok_unverified(TOOL, "style", preview, problems, target_resolved=resolved, applied=applied, **info)
    return rt.ok(TOOL, "style", preview, "VERIFIED: 적용한 서식을 다시 읽어 확인했습니다.", verified=True,
                 target_resolved=resolved, applied=applied, **info)


def _apply_rule(rng: Any, rule: Tuple[str, str, Any]) -> None:
    """--highlight 조건부 강조 하나 (cmd_style과 같은 식)."""
    op, val, color = rule
    if op == "contains":
        top_left = str(rng.Cells(1, 1).Address).replace("$", "")
        lit = re.sub(r"([~?*])", r"~\1", val).replace('"', '""')
        fc = rng.FormatConditions.Add(Type=2, Formula1=f'=ISNUMBER(SEARCH("{lit}",{top_left}))')
    else:
        is_num = bool(re.fullmatch(r"[+-]?\d+(\.\d+)?", val))
        formula = val if is_num else '="' + val.replace('"', '""') + '"'
        fc = rng.FormatConditions.Add(Type=1, Operator=XL_OPERATOR[op], Formula1=f"={formula}" if is_num else formula)
    fc.Interior.Color = color


def _style_many(args: Any, fmt: Dict[str, Any], rule: Optional[Tuple[str, str, Any]], prefix: str,
                parts: List[str]) -> Dict[str, Any]:
    """--range "B2;D5;F7": 떨어진 칸 여러 개에 같은 서식을 명령 한 번으로 (1.4.91). 되돌리기 지점 하나, 칸마다 다시 읽어 확인."""
    with write_session(args, "style") as t:
        ws = t.sheet(args.sheet)
        rngs = [target_range(t, ws, a) for a in _style_chunks(prefix, parts)]
        ws = rngs[0].Worksheet
        check_protection(ws, rngs, "format_cells")
        applied: Dict[str, Any] = {}
        with off.excel_batch(t.app):
            for rng in rngs:
                if args.clear:
                    rng.ClearFormats()
                    rng.FormatConditions.Delete()
                if fmt:
                    _apply_format(rng, fmt)
                if rule is not None:
                    _apply_rule(rng, rule)
        if args.clear:
            applied["cleared"] = True
        if fmt:
            applied["format"] = sorted(fmt)
        if rule is not None:
            applied["highlight"] = args.highlight
        problems = [p for rng in rngs for p in _verify_style(ws, rng, fmt, None, bool(args.highlight))]
        t.set_focus(sheet=ws)
        sheet = str(ws.Name)
        resolved = t.resolved()
    shown = ";".join(parts[:20]) + (f" 외 {len(parts) - 20}곳" if len(parts) > 20 else "")
    preview = f"{sheet}!{shown} ({len(parts)}곳)에 서식 적용: {', '.join(map(str, applied))}"
    info = {"sheet": sheet, "range": ";".join(parts), "areas": len(parts)}
    if problems:
        return rt.ok_unverified(TOOL, "style", preview, problems[:10], target_resolved=resolved, applied=applied, **info)
    return rt.ok(TOOL, "style", preview, f"VERIFIED: {len(parts)}곳에 적용한 서식을 다시 읽어 확인했습니다.", verified=True,
                 target_resolved=resolved, applied=applied, **info)


def _verify_style(ws: Any, rng: Any, fmt: Dict[str, Any], theme: Optional[Dict[str, Any]], highlight: bool) -> List[str]:
    """R1: 범위에 준 서식(배경·글자색·굵게·기울임·크기·숫자 서식·자동 줄바꿈)과 테마 머리글 색, 조건부 강조를 다시 읽어 비교합니다."""
    problems: List[str] = []
    checks = [("bg", lambda: rng.Interior.Color), ("color", lambda: rng.Font.Color), ("bold", lambda: rng.Font.Bold),
              ("italic", lambda: rng.Font.Italic), ("size", lambda: rng.Font.Size), ("number_format", lambda: rng.NumberFormat)]
    try:
        for key, read in checks:
            if key in fmt:
                got = read()
                same = (float(got) == float(fmt[key])) if key in ("bg", "color", "size") and got is not None else bool(got) == fmt[key] \
                    if key in ("bold", "italic") else got == fmt[key]
                if not same:
                    problems.append(f"서식 '{key}'이(가) {got!r} (준 값 {fmt[key]!r})")
        if theme:
            r0, c0, ncols = int(rng.Row), int(rng.Column), int(rng.Columns.Count)
            got = ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + ncols - 1)).Interior.Color
            if got is None or float(got) != float(off.parse_color(theme["header_bg"])):
                problems.append(f"테마 머리글 배경색이 {got!r}")
        if fmt.get("wrap") and not bool(rng.WrapText):     # 여러 칸 중 일부만 켜졌으면 None
            problems.append(f"서식 'wrap'(자동 줄바꿈)이 {rng.WrapText!r}")
        # 가로 맞춤·테두리도 다시 읽음 (--align·--border만 주면 아무것도 확인하지 않고 VERIFIED였음, 2026-11-09 검토 B13)
        if "align" in fmt:
            got = rng.HorizontalAlignment
            if got is None or int(got) != int(fmt["align"]):
                problems.append(f"서식 'align'(가로 맞춤)이 {got!r} (준 값 {fmt['align']!r})")
        if "border" in fmt:
            problems += _verify_border(rng, fmt["border"])
        if highlight and int(rng.FormatConditions.Count) < 1:
            problems.append("조건부 강조가 없음")
    except Exception as e:
        problems.append(f"서식을 다시 읽지 못함: {type(e).__name__}")
    return problems


# ==============================================================================
# 7. export
# ==============================================================================

def _nonblank_rows(matrix: List[List[Any]]) -> int:
    """다시 읽은 CSV와 비교할 행 수 (check_saved_file처럼 빈 행은 세지 않음)."""
    return sum(1 for r in matrix if any(str(out_value(v)).strip() for v in r))


def _check_encodable(matrix: List[List[Any]], encoding: str, r0: int, cols: List[int]) -> None:
    """CSV 인코딩으로 쓸 수 없는 글자(cp949의 이모지 등)가 있으면 쓰기 전에 STOP_RETRY (내부 오류로 끝나지 않게).
    cols: matrix 열마다의 원래 열 번호. 오류 로그에는 칸 주소만 남깁니다(문서 글 없음)."""
    if encoding.startswith("utf-8"):
        return
    bad = []
    for i, row in enumerate(matrix):
        for j, v in enumerate(row):
            text = out_value(v)
            if isinstance(text, str):
                try:
                    text.encode(encoding)
                except UnicodeEncodeError:
                    bad.append(addr(r0 + i, cols[j]))
    if bad:
        more = f" 외 {len(bad) - 5}개" if len(bad) > 5 else ""
        raise rt.InputError(f"--encoding {encoding}로 쓸 수 없는 글자(이모지·일부 기호)가 있는 칸: {', '.join(bad[:5])}{more}. "
                            f"값을 바꾸지 않고 저장하려면 UTF-8로 저장하세요. [추천] {rt.retry_command({'--encoding': 'utf-8-sig'})}")


def cmd_export(args) -> Dict[str, Any]:
    ext = os.path.splitext(args.save)[1].lower() if args.save else ".csv"
    if ext not in (".csv", ".xlsx", ".pdf"):
        raise rt.InputError(f"--save 확장자는 .csv, .xlsx, .pdf 중 하나여야 합니다: {args.save}")
    closed = is_closed_xlsx(args.target)
    if ext == ".csv" and closed and not args.columns:
        sheet, r0, c0, matrix, _ = _read_matrix_closed(closed, args.sheet, args.range)
        _check_encodable(matrix, args.encoding, r0, [c0 + j for j in range(len(matrix[0]) if matrix else 0)])
        out = rt.resolve_output_path(args.save, sheet, ".csv")
        out = rt.write_csv_rows(out, [[out_value(v) for v in r] for r in matrix], args.encoding)
        far = _far_for_read(closed, sheet, args.range, r0, c0, matrix)
        return _with_far(xf.with_uncomputed(rt.ok_file(TOOL, "export", f"{sheet} ({len(matrix)}행)를 CSV로 저장: {out}", out,
                                                       "VERIFIED: 저장한 파일을 다시 읽어 행 수까지 확인했습니다.", delimiter=",",
                                                       expected_rows=_nonblank_rows(matrix), output_file=out,
                                                       rows=len(matrix)),
                                            closed_uncomputed(closed, sheet, r0, c0, matrix)), far, args.target)

    with off.open_target("excel", args.target, readonly=True) as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range)
        ws = rng.Worksheet
        stem = f"{t.name.rsplit('.', 1)[0]}_{ws.Name}"
        out = rt.resolve_output_path(args.save, stem, ext)
        expected_rows = None
        if ext == ".csv":
            matrix = off.range_to_matrix(rng.Value)
            c0 = int(rng.Column)
            cols = [c0 + j for j in range(len(matrix[0]) if matrix else 0)]
            if args.columns:
                idx = [col_number(c.strip()) - c0 for c in args.columns.split(",") if c.strip()]
                bad = [args.columns.split(",")[k] for k, i in enumerate(idx) if i < 0 or i >= len(matrix[0])]
                if bad:
                    raise rt.InputError(f"--columns의 열이 범위 밖입니다: {bad}")
                matrix = [[row[i] for i in idx] for row in matrix]
                cols = [c0 + i for i in idx]
            _check_encodable(matrix, args.encoding, int(rng.Row), cols)
            out = rt.write_csv_rows(out, [[out_value(v) for v in r] for r in matrix], args.encoding)
            expected_rows = _nonblank_rows(matrix)
            detail = f"{len(matrix)}행"
        elif ext == ".pdf":
            off.require_pdf_export(t.app, ".csv 또는 .xlsx")
            rng.ExportAsFixedFormat(0, os.path.normpath(out))   # xlTypePDF
            detail = "PDF"
        else:
            app = t.app
            prev_alerts = app.DisplayAlerts
            new_wb = None
            try:
                app.DisplayAlerts = False
                new_wb = app.Workbooks.Add()
                try:
                    new_wb.Windows(1).Visible = False
                except Exception:
                    pass
                dest = new_wb.Worksheets(1)
                rng.Copy(Destination=dest.Range("A1"))
                dest.Name = str(ws.Name)[:31]
                dest.Columns.AutoFit()
                try:
                    new_wb.Windows(1).Visible = True
                except Exception:
                    pass
                new_wb.SaveAs(os.path.normpath(out), FileFormat=51)
            finally:
                if new_wb is not None:
                    try:
                        new_wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                app.DisplayAlerts = prev_alerts
            detail = "XLSX"
        info = range_info(rng)
        resolved = t.resolved()
    return rt.ok_file(TOOL, "export", f"{info['sheet']}!{info['range']}를 {detail}로 저장: {out}", out,
                      "VERIFIED: 저장한 파일을 확인했습니다. 원본 통합문서는 그대로입니다.", delimiter=",",
                      expected_rows=expected_rows, target_resolved=resolved, output_file=out, **info)


# ==============================================================================
# 8. sheet / range
# ==============================================================================

def cmd_sheet(args) -> Dict[str, Any]:
    if args.action == "list":
        with off.open_target("excel", args.target, readonly=True) as t:
            names = [str(t.doc.Worksheets(i).Name) for i in range(1, t.doc.Worksheets.Count + 1)]
            active = str(t.doc.ActiveSheet.Name)
            resolved = t.resolved()
        return rt.dry_run_read_only(rt.ok(TOOL, "sheet", f"시트 {len(names)}개: {', '.join(names)}", target_resolved=resolved,
                                          sheets=names, active_sheet=active), args.dry_run, "sheet list")
    if not args.name:
        raise rt.InputError(f"sheet {args.action}에는 --name이 필요합니다.")
    if args.dry_run and args.action != "delete":
        raise rt.dry_run_unsupported(TOOL, f"sheet --action {args.action}")
    with write_session(args, "sheet") as t:
        sheets = t.doc.Worksheets
        names = [str(sheets(i).Name) for i in range(1, sheets.Count + 1)]
        existing = rt.find_name(args.name, names)
        if args.action == "add":
            if existing:
                ws = sheets(existing)
                msg = f"시트 '{existing}'이(가) 이미 있어 그대로 사용합니다"
            else:
                off.check_sheet_name(args.name, names)
                ws = t.sheet(args.name, create=True)
                msg = f"시트 '{args.name}'을(를) 추가했습니다"
            t.set_focus(sheet=ws, changed=not existing)     # 이미 있으면 바뀐 것 없음 (되돌리기 지점을 남기지 않음)
        else:
            if existing is None:
                raise rt.not_found("시트", args.name, names, option="--name")
            ws = sheets(existing)
            if args.action == "select":
                if t.source in ("file", "new"):
                    # 닫힌 파일은 백그라운드로 열었다 저장 없이 닫혀 남는 것이 없음 ('선택했습니다' VERIFIED였음, 2026-11-09 검토 B13)
                    raise rt.InputError(f"sheet --action select는 화면에 열린 통합문서에서만 씁니다 ({os.path.basename(str(args.target))}는 열려 있지 "
                                        "않음). 시트 내용을 보려면 excel.py read --sheet로 읽으세요.")
                t.set_focus(sheet=ws, changed=False)
                msg = f"시트 '{existing}'을(를) 선택했습니다"
            elif args.action == "rename":
                if not args.new_name:
                    raise rt.InputError("rename에는 --new-name이 필요합니다.")
                if rt.find_name(args.new_name, names):
                    raise rt.InputError(f"'{args.new_name}' 이름의 시트가 이미 있습니다.")
                off.check_sheet_name(args.new_name, names)
                ws.Name = args.new_name
                t.set_focus(sheet=ws)
                msg = f"시트 '{existing}' -> '{args.new_name}'"
            elif args.action == "copy":
                new_name = args.new_name or f"{existing[:28]} 복사"     # 기본 이름도 31자 이내
                if rt.find_name(new_name, names):
                    raise rt.InputError(f"'{new_name}' 이름의 시트가 이미 있습니다.")
                off.check_sheet_name(new_name, names)      # 복사한 뒤 이름에서 실패하면 복사본이 남음
                ws.Copy(After=sheets(sheets.Count))
                copied = sheets(sheets.Count)
                copied.Name = new_name
                t.set_focus(sheet=copied)
                msg = f"시트 '{existing}'을(를) '{new_name}'(으)로 복사했습니다"
            else:   # delete
                if sheets.Count == 1:
                    raise rt.InputError("마지막 남은 시트는 삭제할 수 없습니다.")
                if args.dry_run:
                    return _sheet_delete_dry(t, ws, existing)
                prev = t.app.DisplayAlerts
                t.app.DisplayAlerts = False
                try:
                    t.changed = True
                    ws.Delete()
                finally:
                    t.app.DisplayAlerts = prev
                msg = f"시트 '{existing}'을(를) 삭제했습니다 (되돌리기: office.py undo)"
        names = [str(sheets(i).Name) for i in range(1, sheets.Count + 1)]
        resolved = t.resolved()
    # R1: 다시 읽은 시트 목록으로 결과 확인
    want = {"add": (existing or args.name, True), "rename": (args.new_name, True), "delete": (existing, False),
            "copy": (args.new_name or f"{(existing or '')[:28]} 복사", True), "select": (existing, True)}[args.action]
    if (rt.find_name(want[0], names) is not None) != want[1]:
        return rt.ok_unverified(TOOL, "sheet", msg, [f"다시 읽은 시트 목록에 '{want[0]}'이(가) " + ("없음" if want[1] else "남아 있음")],
                                target_resolved=resolved, sheets=names)
    return rt.ok(TOOL, "sheet", msg, "VERIFIED: 시트 목록을 다시 읽어 확인했습니다.", verified=True, target_resolved=resolved, sheets=names)


def _sheet_delete_dry(t: Any, ws: Any, name: str) -> Dict[str, Any]:
    """sheet --action delete --dry-run: 지울 시트의 사용 범위·차트 수와 #REF!가 될 다른 시트의 수식 칸 (바꾸지 않음)."""
    box = _used_box(ws)
    used = addr(*box) if box else ""
    charts = len(_chart_objects(ws))
    refs, skipped = broken_refs(t.doc, ws, "rows", 1, XL_MAX_ROW)    # 시트 전체 = 모든 행
    parts = [f"사용 범위 {used}" if used else "빈 시트"] + ([f"차트 {charts}개"] if charts else [])
    preview = f"미리보기: 시트 '{name}'을(를) 삭제합니다 ({', '.join(parts)})"
    if refs:
        preview += f". 다른 시트의 수식 {len(refs)}개가 이 시트를 참조해 #REF! 오류가 됩니다: " + ", ".join(refs[:5]) + (" ..." if len(refs) > 5 else "")
    extra: Dict[str, Any] = {"sheet": name, "used_range": used, "charts": charts, "broken_refs": len(refs),
                             "broken_ref_cells": refs[:REF_SHOW]}
    if len(refs) > REF_SHOW:
        extra["broken_ref_cells_truncated"] = True
    if skipped:
        extra["warnings"] = [f"칸이 너무 많아 이 시트를 참조하는 수식을 확인하지 못한 시트: {', '.join(skipped[:10])}"]
    return rt.dry_run_result(TOOL, "sheet", preview + " (바꾸지 않음)", 1, [f"삭제: 시트 '{name}'"],
                      rt.dry_run_command(TOOL, f"시트 '{name}' 삭제 실행"), target_resolved=t.resolved(), **extra)


def _same_cell(want: Any, got: Any) -> bool:
    """쓴 값과 다시 읽은 값이 같은지. 앞의 '(글자로 두기)는 값에 남지 않고, Excel이 날짜 등으로 바꾼 글자는 비어 있지만 않으면 인정."""
    if isinstance(want, str) and want.startswith("'"):
        want = want[1:]
    if want is None or want == "":
        return got is None or got == ""
    if isinstance(want, (int, float)) and not isinstance(want, bool):
        try:
            return abs(float(got) - float(want)) <= 1e-9 * max(1.0, abs(float(want)))
        except (TypeError, ValueError):
            return False
    if isinstance(got, str):
        return got == str(want)
    if got is None or got == "" or off.xl_error(got):
        return False
    if isinstance(want, str):
        # 글을 썼는데 글이 아닌 값: Excel이 바꾼 것. 날짜 모양 글이 날짜가 되고 true/false가 논리값이 되는 것만 인정
        # (전에는 비어 있지만 않으면 인정해 '+82-10-...'이 -6840이 되어도 VERIFIED, 2026-11-09 검토 A8)
        w = want.strip()
        if isinstance(got, bool):
            return w.lower() in ("true", "false")
        if isinstance(got, (int, float)):
            # 숫자 모양 글('3', '1,500', '12.5%')을 Excel이 같은 값의 숫자로 바꾼 것은 인정 (워드 표 -> 엑셀, 회사 PC 2026-11-11 I23).
            # 앞자리 0이 사라진 것('0012' -> 12)은 값이 바뀐 것이라 인정하지 않음
            m = _NUM_LIKE.match(w)
            if m and not re.match(r"^[+-]?0\d", w):
                num = float(m.group(1).replace(",", "")) / (100.0 if m.group(2) else 1.0)
                if abs(float(got) - num) <= 1e-9 * max(1.0, abs(num)):
                    return True
        return bool(_DATE_LIKE.match(w)) and (hasattr(got, "year") or isinstance(got, (int, float)))
    return True


_NUM_LIKE = re.compile(r"^([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*(%)?$")


# Excel이 날짜·시각으로 바꾸는 글 (2026-01-05, 2024-01, 2026/1/5, 10:30)
_DATE_LIKE = re.compile(r"^\d{4}\s*[-/.]\s*\d{1,2}(?:\s*[-/.]\s*\d{1,2})?(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$|^\d{1,2}:\d{2}(?::\d{2})?$")


def _verify_written(ws: Any, ops: List[Tuple[int, int, int, int, str, Any]]) -> List[str]:
    """R1: 쓴 칸을 다시 읽어 비교합니다. 수식은 첫 칸의 수식(대소문자·공백 무시), 값은 칸마다."""
    problems: List[str] = []
    norm = lambda f: re.sub(r"\s+", "", str(f)).upper()  # noqa: E731
    for a, b, c, d, kind, payload in ops:
        try:
            if kind == "formula":
                got = ws.Cells(a, b).Formula
                if norm(got).replace("@", "") != norm(payload).replace("@", ""):
                    problems.append(f"{addr(a, b)} 수식이 {got!r} (쓴 수식 {payload!r})")
                    continue
                # 수식 글이 같아도 값이 오류면 확인한 것이 아님 (함수 이름 오타 '=합계(B2:B3)' -> #NAME?, 2026-11-09 검토 A9)
                vals = off.range_to_matrix(ws.Range(ws.Cells(a, b), ws.Cells(c, d)).Value)
                errs = [(addr(a + i, b + j), off.xl_error(v)) for i, row in enumerate(vals) for j, v in enumerate(row) if off.xl_error(v)]
                if errs:
                    problems.append(f"수식 칸 {len(errs)}개의 값이 오류: " + ", ".join(f"{c_} {e}" for c_, e in errs[:5])
                                    + " (함수 이름·참조·인수 확인)")
                continue
            got_rows = off.range_to_matrix(ws.Range(ws.Cells(a, b), ws.Cells(c, d)).Value)
            want_rows = payload if kind == "values" else [[payload] * (d - b + 1) for _ in range(c - a + 1)]
            bad = [addr(a + i, b + j) for i, row in enumerate(want_rows) for j, w in enumerate(row)
                   if not _same_cell(w, got_rows[i][j] if i < len(got_rows) and j < len(got_rows[i]) else None)]
            if bad:
                problems.append(f"다시 읽은 값이 다른 칸 {len(bad)}개: {', '.join(bad[:5])}")
        except Exception as e:
            problems.append(f"{addr(a, b, c, d)}를 다시 읽지 못함: {type(e).__name__}")
    return problems


def _is_number(text: str) -> bool:
    try:
        float(text.strip().replace(",", ""))
        return True
    except ValueError:
        return False


def _filled_cells(rng: Any) -> List[Tuple[str, Any]]:
    """범위에서 값이나 수식이 있는 칸 [(주소, 값 또는 수식)] (문서 순서).
    빈 글("")을 내는 수식도 칸을 채운 것으로 봅니다 (값만 보면 확인 없이 덮어썼음)."""
    r0, c0 = int(rng.Row), int(rng.Column)
    try:
        forms = off.range_to_matrix(rng.Formula)
    except Exception:
        forms = []
    out = []
    for i, row in enumerate(off.range_to_matrix(rng.Value)):
        for j, v in enumerate(row):
            if v is not None and str(v) != "":
                out.append((addr(r0 + i, c0 + j), v))
                continue
            f = forms[i][j] if i < len(forms) and j < len(forms[i]) else None
            if isinstance(f, str) and f != "":
                out.append((addr(r0 + i, c0 + j), f))
    return out


def _r1c1_text(v: Any) -> str:
    return "" if v is None else str(v)


def _axis(args) -> str:
    """--axis 생략 시 주소로 판단: '5:7', '5'처럼 행 번호만 있으면 rows (cols로 하면 시트의 모든 열이 숨겨짐)."""
    if args.axis:
        return args.axis
    cells = split_sheet(args.range or "")[1].replace("$", "")
    return "rows" if re.fullmatch(r"\d+(?::\d+)?", cells) else "cols"


def _inside(cell: Tuple[int, int], box: Tuple[int, int, int, int]) -> bool:
    return box[0] <= cell[0] <= box[2] and box[1] <= cell[1] <= box[3]


def _short(v: Any, n: int = 15) -> str:
    s = str(out_value(v)).replace(",", " ")
    return s if len(s) <= n else s[:n] + "..."


def _filled_preview(filled: List[Tuple[str, Any]], log: bool = False) -> str:
    """덮어쓰거나 지워질 칸 미리보기 'A1=값, B1=값 ...' (최대 5칸). log=True면 오류 로그용: 칸 주소만 (값은 문서 내용)."""
    more = " ..." if len(filled) > 5 else ""
    if log:
        return ", ".join(c for c, _ in filled[:5]) + more + f" (값 {rt.DOC_OMITTED})"
    return ", ".join(f"{c}={_short(v)}" for c, v in filled[:5]) + more


def _formula_cells(rng: Any) -> List[Tuple[str, str]]:
    """범위 안의 수식 칸 [(주소, 수식)] (사용 범위 밖은 비어 있으므로 사용 범위와 겹친 부분만 읽음)."""
    try:
        used = rng.Application.Intersect(rng, rng.Worksheet.UsedRange)
    except Exception:
        used = rng
    if used is None:
        return []
    out: List[Tuple[str, str]] = []
    for area in _areas(used):
        r1, c1 = int(area.Row), int(area.Column)
        for i, row in enumerate(off.range_to_matrix(area.Formula)):
            for j, f in enumerate(row):
                if isinstance(f, str) and f.startswith("="):
                    out.append((addr(r1 + i, c1 + j), f))
    return out


def _areas(rng: Any) -> List[Any]:
    try:
        n = int(rng.Areas.Count)
        return [rng.Areas(i) for i in range(1, n + 1)]
    except Exception:
        return [rng]


def _range_noop(t: Any, ws: Any, rng: Any, a: str, args: Any) -> str:
    """range 명령이 바꿀 것이 없는지 COM 속성으로 싸게 확인합니다. 바꿀 것이 없으면 그 이유, 아니면(모르면) ""."""
    try:
        # hide/unhide는 확인하지 않음: 일부만 숨긴 범위에서 Hidden이 섞임(None)이 아니라 False를 돌려줄 수 있어
        # 실제 바뀌는 unhide를 건너뛸 위험이 있음 (바꾼 것으로 보면 되돌리기 지점이 하나 더 남을 뿐)
        if a == "unmerge":
            if rng.MergeCells is False:                # 섞여 있으면 None
                return "병합된 칸이 없음"
        elif a == "merge":
            if rng.MergeCells is True and str(rng.Cells(1, 1).MergeArea.Address) == str(rng.Address):
                return "이미 이 범위로 병합되어 있음"
        elif a == "unfreeze":
            ws.Activate()
            if not t.app.ActiveWindow.FreezePanes:
                return "고정된 틀이 없음"
        elif a in ("width", "height") and args.size is not None:
            got = rng.ColumnWidth if a == "width" else rng.RowHeight     # 섞여 있으면 None
            if got is not None and float(got) == float(args.size):
                return f"이미 {'너비' if a == 'width' else '높이'}가 {args.size}"
        elif a == "values":
            if not _formula_cells(rng):
                return "수식 칸이 없음"
        elif a == "clear":
            area = rng if rng is not None else ws.UsedRange
            used = t.app.Intersect(area, ws.UsedRange)           # 값·서식이 있는 칸은 UsedRange 안에 있음
            # 서식·전체 지우기는 확인하지 않음: 사용 범위 밖 칸에도 데이터 유효성·메모가 있을 수 있음
            if used is None:
                if args.type == "contents":
                    return "빈 범위 (사용 범위 밖)"
            elif args.type == "contents" and "," not in str(used.Address) and not _filled_cells(used):
                return "값이 있는 칸이 없음"
    except Exception as e:
        rt.debug(f"바뀌는 것이 있는지 확인 생략: {e}")
    return ""


def _check_range_protection(ws: Any, rng: Any, a: str, args: Any) -> None:
    """range 동작별 시트 보호 확인 (바꾸기 전에). freeze/unfreeze는 창 설정이라 보호와 상관없음."""
    if a in ("freeze", "unfreeze", "copy"):
        return
    rows = _axis(args) == "rows"
    need = {"autofit": "format_rows" if rows else "format_columns", "width": "format_columns", "height": "format_rows",
            "hide": "format_rows" if rows else "format_columns", "unhide": "format_rows" if rows else "format_columns",
            "border": "format_cells", "merge": "structure", "unmerge": "structure", "move": "structure"}.get(a)
    target = rng if rng is not None else ws.UsedRange
    if a == "clear":
        if args.type in ("all", "contents"):
            check_protection(ws, [target], "values")
        if args.type in ("all", "formats"):
            check_protection(ws, [target], "format_cells")
        return
    if need:
        check_protection(ws, [target], need)


def _freeze_cell_check(cell: str) -> None:
    """range --action freeze의 칸: 칸 주소만. 행·열 전체('3:3', 'C:C')는 그 행·열의 첫 칸으로 [추천]."""
    body = split_sheet(cell)[1]
    if is_cell_address(body):
        return
    m = re.fullmatch(r"\$?(\d+):\$?(\d+)", body.strip()) or re.fullmatch(r"\$?([A-Za-z]{1,3}):\$?[A-Za-z]{1,3}", body.strip())
    rec = ""
    if m:
        first = m.group(1)
        rec = " [추천] " + rt.retry_command({"--range": f"A{first}" if first.isdigit() else f"{first.upper()}1"})
    raise rt.InputError(f"freeze의 --range '{cell}'은(는) 칸 주소가 아닙니다: 그 칸의 위 행과 왼쪽 열을 고정합니다 (예: A2는 1행, B3은 1~2행과 A열). "
                        f"아무것도 바꾸지 않았습니다.{rec}")


def _range_dest(t: Any, ws: Any, args) -> Tuple[Any, str, Any]:
    """copy/move 목적지 (시트, 칸 주소, 범위). 붙여 넣을 시트가 보호되어 있으면 멈춤 (move의 원본 시트는 따로 확인)."""
    if not args.dest:
        raise rt.InputError(f"range {args.action}에는 --dest가 필요합니다 (예: --dest Sheet2!A1).")
    dsheet, dcell = split_sheet(args.dest)
    if not is_cell_address(dcell):
        # '--dest 요약'(시트 이름만) -> ERROR였음 (2026-11-09 검토 B12). 그런 시트가 있으면 그 시트 A1
        names = [str(t.doc.Worksheets(i).Name) for i in range(1, t.doc.Worksheets.Count + 1)] if not dsheet else []
        hit = rt.find_name(dcell, names) if names else None
        rec = f" [추천] {rt.retry_command({'--dest': f'{hit}!A1'})}" if hit else ""
        raise rt.InputError(f"--dest '{args.dest}'은(는) 칸 주소가 아닙니다 (예: --dest A20, --dest 'Sheet2'!A1). 아무것도 바꾸지 않았습니다.{rec}")
    dws = t.sheet(dsheet) if dsheet else ws
    dest = dws.Range(dcell)
    check_protection(dws, [dest], "structure")
    return dws, dcell, dest


def _dest_conflicts(rng: Any, ws: Any, dws: Any, dcell: str) -> Tuple[List[Tuple[str, Any]], Tuple[int, int, int, int]]:
    """copy/move로 덮어써질 대상의 값이 있는 칸 [(주소, 값)]과 (대상 행, 열, 행 수, 열 수). 원본과 겹치는 칸은 옮겨지는 값이라 뺌."""
    dr, dc = parse_cell(dcell.replace("$", "").split(":")[0])
    nr, nc = int(rng.Rows.Count), int(rng.Columns.Count)
    block = dws.Range(dws.Cells(dr, dc), dws.Cells(dr + nr - 1, dc + nc - 1))
    src = (int(rng.Row), int(rng.Column), int(rng.Row) + nr - 1, int(rng.Column) + nc - 1)
    same_sheet = str(dws.Name) == str(ws.Name)
    filled = [(c, v) for c, v in _filled_cells(block) if not (same_sheet and _inside(parse_cell(c), src))]
    return filled, (dr, dc, nr, nc)


def _range_dry(t: Any, ws: Any, rng: Any, a: str, args, noop: str) -> Dict[str, Any]:
    """range --action clear|move|values --dry-run: 지울 값·옮길 칸과 덮어써질 칸·값으로 바뀔 수식만 보여 줌 (바꾸지 않음, 되돌리기 지점 없음)."""
    where = f"{ws.Name}!{str(rng.Address).replace('$', '')}"
    info: Dict[str, Any] = {"sheet": str(ws.Name), "range": str(rng.Address).replace("$", "")}
    if a == "values":
        forms = _formula_cells(rng)
        if not forms:
            return rt.dry_run_result(TOOL, "range", f"미리보기: {where}에는 수식 칸이 없어 바뀔 것 없음 (바꾸지 않음)", 0, [], None,
                                     target_resolved=t.resolved(), **info)
        planned = [f"값으로: {ws.Name}!{c} ({_short(f)})" for c, f in forms[:rt.DRY_RUN_SHOW + 1]]
        info["formula_cells"] = len(forms)
        return rt.dry_run_result(TOOL, "range", f"미리보기: {where}의 수식 칸 {len(forms)}개를 지금 값으로 바꿉니다 "
                                 "(수식은 없어지고 값은 그대로, 바꾸지 않음)", len(forms), planned,
                                 rt.dry_run_command(TOOL, f"{where} 수식 {len(forms)}개를 값으로 바꾸기"),
                                 target_resolved=t.resolved(), **info)
    if a == "clear":
        info["type"] = args.type
        if noop:
            return rt.dry_run_result(TOOL, "range", f"미리보기: {where}에는 바뀔 것 없음 ({noop}) (바꾸지 않음)", 0, [], None, target_resolved=t.resolved(), **info)
        filled: List[Tuple[str, Any]] = []
        if args.type != "formats":
            used = t.app.Intersect(rng, ws.UsedRange)        # 값이 있는 칸은 사용 범위 안에만 있음
            filled = _filled_cells(used) if used is not None else []
        planned = [f"지움: {ws.Name}!{c} ({_short(v)})" for c, v in filled[:rt.DRY_RUN_SHOW + 1]]
        if args.type == "formats":
            count = int(rng.Rows.Count) * int(rng.Columns.Count)
            preview = f"미리보기: {where}의 서식을 지웁니다 (칸 {count}개, 값은 그대로)"
            planned = [f"서식 지움: {where}"]
        elif filled:
            count = len(filled)
            preview = f"미리보기: {where}을(를) 지우면 값이 있는 칸 {count}개가 지워집니다: {_filled_preview(filled)}"
            preview += " (서식도 지움)" if args.type == "all" else ""
        else:
            count = 1
            preview = f"미리보기: {where}에는 값이 있는 칸이 없습니다. 서식만 지워집니다"
            planned = [f"서식 지움: {where}"]
        info["filled_cells"] = len(filled)
        return rt.dry_run_result(TOOL, "range", preview + " (바꾸지 않음)", count, planned,
                          rt.dry_run_command(TOOL, f"{where} 지우기 실행"), target_resolved=t.resolved(), **info)
    # move
    dws, dcell, _ = _range_dest(t, ws, args)
    filled, (dr, dc, nr, nc) = _dest_conflicts(rng, ws, dws, dcell)
    to = f"{dws.Name}!{addr(dr, dc, dr + nr - 1, dc + nc - 1)}"
    info["dest"] = to
    if str(dws.Name) == str(ws.Name) and (dr, dc) == (int(rng.Row), int(rng.Column)):
        return rt.dry_run_result(TOOL, "range", f"미리보기: {where}은(는) 원본과 같은 자리라 바뀔 것 없음 (바꾸지 않음)", 0, [], None, target_resolved=t.resolved(), **info)
    planned = [f"옮김: {where} -> {to}"] + [f"덮어쓰기: {dws.Name}!{c} ({_short(v)})" for c, v in filled[:rt.DRY_RUN_SHOW]]
    preview = f"미리보기: {where}(칸 {nr * nc}개)을(를) {to}(으)로 옮깁니다"
    if filled:
        preview += f". 대상에서 값이 있는 칸 {len(filled)}개가 덮어써집니다: {_filled_preview(filled)}"
        info["overwrite_cells"] = len(filled)
    return rt.dry_run_result(TOOL, "range", preview + " (바꾸지 않음)", nr * nc, planned,
                      rt.dry_run_command(TOOL, f"{where} 옮기기 실행" + (" (값이 있는 대상 칸 덮어쓰기)" if filled else ""), overwrite=bool(filled)),
                      target_resolved=t.resolved(), **info)


def cmd_range(args) -> Dict[str, Any]:
    a = args.action
    if a in ("insert", "delete"):
        return range_lines(args)
    if args.dry_run and a not in ("clear", "move", "values"):
        raise rt.dry_run_unsupported(TOOL, f"range --action {a}")
    if a == "sort":
        return range_sort(args)
    if a in ("merge", "unmerge", "width", "height", "hide", "unhide", "copy", "move", "clear", "values") and not args.range:
        raise rt.InputError(f"range {a}에는 --range가 필요합니다." + (
            " 시트 전체를 지우려면 사용자에게 확인한 뒤 excel.py info로 사용 범위를 보고 그 범위를 --range로 지정하세요."
            if a == "clear" else ""))
    with write_session(args, "range") as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range) if args.range or a not in ("freeze", "unfreeze") else None
        if rng is not None:
            ws = rng.Worksheet
        detail = ""
        # 이미 그 상태라 바뀌는 것이 없으면(숨긴 것 없는 unhide, 고정 없는 unfreeze, 병합 없는 unmerge, 빈 범위 clear 등)
        # 바꾸지 않고 바뀐 것 없음으로 표시 (되돌리기 지점을 남기지 않음). 알 수 없으면 바꾼 것으로 봄
        noop = _range_noop(t, ws, rng, a, args)
        if not noop:
            _check_range_protection(ws, rng, a, args)         # 읽기만 함: 미리보기도 실제 실행과 같이 막힘을 알림
        if args.dry_run:
            return _range_dry(t, ws, rng, a, args, noop)
        if noop:
            detail = f"(바뀐 것 없음: {noop})"
        elif a == "autofit":
            rng.Columns.AutoFit()
            if _axis(args) == "rows":
                rng.Rows.AutoFit()
        elif a in ("width", "height"):
            if args.size is None:
                raise rt.InputError(f"range {a}에는 --size가 필요합니다 (예: --size 20).")
            if a == "width":
                rng.ColumnWidth = float(args.size)
            else:
                rng.RowHeight = float(args.size)
        elif a in ("merge", "unmerge"):
            if a == "merge" and not args.overwrite:
                top_left = addr(int(rng.Row), int(rng.Column))
                lost = [(c, v) for c, v in _filled_cells(rng) if c != top_left]
                if lost:          # 병합하면 왼쪽 위 칸만 남음 (왼쪽 위가 비고 다른 칸에 값 하나만 있어도 지워짐)
                    head = f"병합하면 왼쪽 위 칸({top_left})만 남고 값 {len(lost)}개가 지워집니다: "
                    raise rt.UserDecisionError(
                        head + _filled_preview(lost),
                        "병합하지 않았습니다. 지워질 값을 사용자에게 보여 주고, 사용자가 지워도 된다고 하면 --overwrite를 붙여 다시 실행하세요.",
                        log_message=head + _filled_preview(lost, log=True))
            prev = t.app.DisplayAlerts
            t.app.DisplayAlerts = False        # '왼쪽 위 값만 유지' 확인창 방지
            try:
                rng.Merge() if a == "merge" else rng.UnMerge()
            finally:
                t.app.DisplayAlerts = prev
        elif a in ("freeze", "unfreeze"):
            cell = args.range or "A2"               # 기본: 1행(머리글) 고정
            if a == "freeze":
                _freeze_cell_check(cell)            # 기존 고정을 풀기 전에 (전에는 '3:3'에서 이미 풀고 ERROR, 검토 B12)
            ws.Activate()
            win = t.app.ActiveWindow
            win.FreezePanes = False
            if a == "freeze":
                win.ScrollRow = 1
                win.ScrollColumn = 1
                ws.Range(split_sheet(cell)[1].split(":")[0]).Select()
                win.FreezePanes = True
                detail = f"{cell} 위/왼쪽 고정"
        elif a in ("hide", "unhide"):
            target = rng.EntireRow if _axis(args) == "rows" else rng.EntireColumn
            target.Hidden = a == "hide"
        elif a == "clear":
            rng = rng if rng is not None else ws.UsedRange
            {"contents": rng.ClearContents, "formats": rng.ClearFormats}.get(args.type, rng.Clear)()
            detail = args.type
        elif a in ("copy", "move"):
            dws, dcell, dest = _range_dest(t, ws, args)
            if not args.overwrite:
                filled, (dr, dc, nr, nc) = _dest_conflicts(rng, ws, dws, dcell)
                if filled:
                    head = f"{addr(dr, dc, dr + nr - 1, dc + nc - 1)}에 값이 있는 칸 {len(filled)}개가 덮어써집니다: "
                    raise rt.UserDecisionError(
                        f"대상 {dws.Name}!" + head + _filled_preview(filled),
                        "옮기지 않았습니다. 덮어쓸 칸을 사용자에게 보여 주고, 사용자가 덮어써도 된다고 하면 --overwrite를 붙여 "
                        "다시 실행하세요. 아니면 비어 있는 다른 --dest를 쓰세요.",
                        log_message="대상 " + head + _filled_preview(filled, log=True))
            src_vals = off.range_to_matrix(rng.Value)
            src_forms = None
            if a == "copy":
                # 복사는 상대 참조 수식의 값이 달라지므로 R1C1 수식으로 비교 (=A1*2를 옮기면 =C1*2지만 R1C1은 같음)
                try:
                    src_forms = off.range_to_matrix(rng.FormulaR1C1)
                except Exception:
                    src_forms = None
            dr, dc = parse_cell(dcell.replace("$", "").split(":")[0])
            dest_block = (dws, dr, dc, dr + len(src_vals) - 1, dc + len(src_vals[0]) - 1)
            if str(dws.Name) == str(ws.Name) and (dr, dc) == (int(rng.Row), int(rng.Column)):
                noop = "원본과 같은 자리"          # 같은 자리로 복사·이동은 바뀐 것 없음
                detail = f"-> {dws.Name}!{dcell} (바뀐 것 없음: {noop})"
            else:
                if a == "copy":
                    rng.Copy(Destination=dest)
                else:
                    rng.Cut(Destination=dest)           # 겹쳐도 데이터 손실 없음, 수식 참조 갱신
                detail = f"-> {dws.Name}!{dcell}"
            ws = dws
        elif a == "border":
            _apply_border(rng if rng is not None else ws.UsedRange, args.weight, box=args.box)
            detail = args.weight + (" (외곽만)" if args.box else "")
        elif a == "values":
            # 수식을 지금 값으로 (Value2: 날짜·통화를 바꾸지 않고 그대로). 영역마다 따로 (여러 영역 범위는 한 번에 못 씀)
            # 글자 결과는 앞에 '를 붙여 그대로 (붙이지 않으면 '01234'·'1-2'가 숫자·날짜로 바뀜), 오류 값(#N/A 등)은 숫자 코드가
            # 들어가지 않게 그 칸만 수식 그대로 둠 (실제 Excel의 Value2는 오류를 음수 정수로 줌)
            forms = _formula_cells(rng)
            before_vals, kept_errors = [], []
            for area in _areas(rng):
                vals = off.range_to_matrix(area.Value2)
                fml = off.range_to_matrix(area.Formula)
                r1_, c1_ = int(area.Row), int(area.Column)
                out_rows = []
                for i, row in enumerate(vals):
                    line = []
                    for j, v in enumerate(row):
                        f = fml[i][j] if i < len(fml) and j < len(fml[i]) else None
                        if off.xl_error(v) and isinstance(f, str) and f.startswith("="):
                            kept_errors.append(addr(r1_ + i, c1_ + j))
                            line.append(f)
                        elif isinstance(v, str):
                            line.append("'" + v)
                        else:
                            line.append(v)
                    out_rows.append(tuple(line))
                area.Value2 = tuple(out_rows)
                before_vals.append((area, vals))
            detail = f"(수식 {len(forms) - len(kept_errors)}개를 값으로)"
            if kept_errors:
                values_warn = f"오류 값 칸 {len(kept_errors)}개는 수식 그대로 둠: {', '.join(kept_errors[:5])}"
        # R1: 결과를 다시 읽어 확인 (autofit은 Excel이 정하는 너비라 비교할 값이 없음)
        problems: List[str] = []
        try:
            if a in ("copy", "move"):
                d, r1_, c1_, r2_, c2_ = dest_block
                dest_rng = d.Range(d.Cells(r1_, c1_), d.Cells(r2_, c2_))
                if src_forms is not None:
                    got = off.range_to_matrix(dest_rng.FormulaR1C1)
                    bad = [addr(r1_ + i, c1_ + j) for i, row in enumerate(src_forms) for j, v in enumerate(row)
                           if _r1c1_text(v) != _r1c1_text(got[i][j] if i < len(got) and j < len(got[i]) else None)]
                else:
                    got = off.range_to_matrix(dest_rng.Value)
                    bad = [addr(r1_ + i, c1_ + j) for i, row in enumerate(src_vals) for j, v in enumerate(row)
                           if not _same_cell(v, got[i][j] if i < len(got) and j < len(got[i]) else None)]
                if bad:
                    problems.append(f"대상에서 원본과 다른 칸 {len(bad)}개: {', '.join(bad[:5])}")
            elif a == "clear" and args.type != "formats":
                left = _filled_cells(rng)
                if left:
                    problems.append(f"지운 범위에 값이 남은 칸 {len(left)}개: {', '.join(c for c, _ in left[:5])}")
            elif a in ("merge", "unmerge") and bool(rng.MergeCells) != (a == "merge"):
                problems.append(f"병합 상태가 {rng.MergeCells!r}")
            elif a in ("width", "height"):
                got = rng.ColumnWidth if a == "width" else rng.RowHeight
                if got is None or abs(float(got) - float(args.size)) > 0.5:
                    problems.append(f"{'너비' if a == 'width' else '높이'}가 {got!r} (준 값 {args.size})")
            elif a in ("hide", "unhide"):
                tgt = rng.EntireRow if _axis(args) == "rows" else rng.EntireColumn
                if bool(tgt.Hidden) != (a == "hide"):
                    problems.append(f"숨김 상태가 {tgt.Hidden!r}")
            elif a in ("freeze", "unfreeze") and bool(t.app.ActiveWindow.FreezePanes) != (a == "freeze"):
                problems.append("틀 고정 상태가 다름")
            elif a == "border":
                problems += _verify_border(rng if rng is not None else ws.UsedRange, args.weight, args.box)
            elif a == "values" and not noop:
                left = [x for x in _formula_cells(rng) if x[0] not in set(kept_errors)]
                if left:
                    problems.append(f"수식이 남은 칸 {len(left)}개: {', '.join(c for c, _ in left[:5])}")
                for area, vals in before_vals:
                    got = off.range_to_matrix(area.Value2)
                    r1_, c1_ = int(area.Row), int(area.Column)
                    bad = [addr(r1_ + i, c1_ + j) for i, row in enumerate(vals) for j, v in enumerate(row)
                           if addr(r1_ + i, c1_ + j) not in kept_errors
                           and not _same_cell(v, got[i][j] if i < len(got) and j < len(got[i]) else None)]
                    if bad:
                        problems.append(f"값이 바뀐 칸 {len(bad)}개: {', '.join(bad[:5])}")
        except Exception as e:
            problems.append(f"결과를 다시 읽지 못함: {type(e).__name__}")
        t.set_focus(sheet=ws, changed=not noop)
        info = range_info(rng) if rng is not None else {"sheet": str(ws.Name)}
        if noop:
            info["unchanged"] = True
        if a == "values" and not noop and kept_errors:
            info["warnings"] = [values_warn]
        resolved = t.resolved()
    where = f"{info['sheet']}!{info.get('range', '')}".rstrip("!")
    preview = f"{a} 완료: {where} {detail}".strip()
    if a == "autofit" and not problems and not noop:
        problems = ["자동 맞춤 너비는 Excel이 정해 비교할 값이 없음 (적용만 함)"]
    if problems:
        return rt.ok_unverified(TOOL, "range", preview, problems, target_resolved=resolved, **info)
    return rt.ok(TOOL, "range", preview, "VERIFIED: 결과를 다시 읽어 확인했습니다.", verified=True, target_resolved=resolved, **info)


# ==============================================================================
# 8b. range insert/delete (행·열 넣기·지우기), sort (정렬)
# ==============================================================================

LINES_SCAN = 1000        # 넣기·지우기 확인용 기준 칸을 찾을 때 훑는 행(열은 200) 수
SORT_RANGE_KEYS = 3      # Range.Sort가 받는 기준 수 (Key1~Key3). 넘으면 Worksheet.Sort.SortFields
SORT_MAX_KEYS = 64       # 엑셀 정렬 기준 상한 (정렬 대화상자의 수준 64개)
XL_ASC, XL_DESC = 1, 2
XL_NO = 2     # Range.Sort Header: 머리글은 정렬 범위에서 빼고 늘 xlNo (xlYes·xlGuess에 기대지 않음)


def _lines_spec(args) -> Tuple[Optional[str], str, int, int]:
    """--range를 (시트, 'rows'|'cols', 시작, 끝)으로. '5:7'·'5'는 행, 'B:C'·'B'는 열.
    A5:C7 같은 칸 범위는 --axis가 있어야 함 (없으면 [후보]: 행 3개인지 열 3개인지 도구가 정할 수 없음)."""
    if not args.range:
        raise rt.InputError(f"range {args.action}에는 --range가 필요합니다 (행은 --range 5:7, 열은 --range B:C).")
    sheet, cells = split_sheet(args.range)
    c = cells.replace("$", "").replace(" ", "").upper()
    m_rows = re.fullmatch(r"(\d+)(?::(\d+))?", c)
    m_cols = re.fullmatch(r"([A-Z]{1,3})(?::([A-Z]{1,3}))?", c)
    if m_rows:
        a, b = sorted((_check_row(int(m_rows.group(1)), cells), _check_row(int(m_rows.group(2) or m_rows.group(1)), cells)))
        given = "rows"
    elif m_cols:
        a, b = sorted((col_number(m_cols.group(1)), col_number(m_cols.group(2) or m_cols.group(1))))
        if b > XL_MAX_COL:
            raise rt.InputError(f"열이 엑셀 범위(A~XFD) 밖입니다: '{cells}' (예: B:C)")
        given = "cols"
    else:
        r1, c1, r2, c2 = parse_area(c, None, None)
        if args.axis == "rows":
            return sheet, "rows", r1, r2
        if args.axis == "cols":
            return sheet, "cols", c1, c2
        raise rt.ambiguous(f"--range {cells}이(가) 행인지 열인지 정할 수 없습니다. 행이면 --range {r1}:{r2}, "
                           f"열이면 --range {col_letter(c1)}:{col_letter(c2)}로 쓰거나 --axis를 붙이세요.",
                           [f"--axis rows ({r1}~{r2}행)", f"--axis cols ({col_letter(c1)}~{col_letter(c2)}열)"],
                           values=["rows", "cols"], option="--axis")
    if args.axis and args.axis != given:
        raise rt.InputError(f"--range {cells}은(는) {'행' if given == 'rows' else '열'} 주소인데 --axis {args.axis}를 주었습니다. "
                            f"--axis를 빼고 다시 실행하세요. [추천] {rt.retry_command({'--axis': None})}")
    return sheet, given, a, b


def _lines_text(axis: str, a: int, b: int) -> str:
    if axis == "rows":
        return f"{a}행" if a == b else f"{a}~{b}행"
    return f"{col_letter(a)}열" if a == b else f"{col_letter(a)}~{col_letter(b)}열"


def _lines_addr(axis: str, a: int, b: int) -> str:
    return f"{a}:{b}" if axis == "rows" else f"{col_letter(a)}:{col_letter(b)}"


def _used_box(ws: Any) -> Optional[Tuple[int, int, int, int]]:
    """사용 범위 (r1, c1, r2, c2). 빈 시트(A1 하나뿐이고 비었음)는 None."""
    used = ws.UsedRange
    r1, c1 = int(used.Row), int(used.Column)
    r2, c2 = r1 + int(used.Rows.Count) - 1, c1 + int(used.Columns.Count) - 1
    if (r1, c1, r2, c2) == (1, 1, 1, 1) and not _filled_cells(ws.Cells(1, 1)):
        return None
    return r1, c1, r2, c2


def _first_constant(ws: Any, r1: int, c1: int, r2: int, c2: int) -> Optional[Tuple[int, int, Any]]:
    """범위에서 수식이 아닌 값이 있는 첫 칸 (행, 열, 값). 수식 칸은 옮기면 값이 바뀔 수 있어(=ROW()) 기준으로 쓰지 않음."""
    if r1 > r2 or c1 > c2:
        return None
    rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    vals = off.range_to_matrix(rng.Value)
    try:
        forms = off.range_to_matrix(rng.Formula)
    except Exception:
        forms = []
    for i, row in enumerate(vals):
        for j, v in enumerate(row):
            if v is None or str(v) == "":
                continue
            f = forms[i][j] if i < len(forms) and j < len(forms[i]) else None
            if isinstance(f, str) and f.startswith("="):
                continue
            return r1 + i, c1 + j, v
    return None


def _ref_cell(ws: Any, axis: str, start: int, box: Tuple[int, int, int, int]) -> Optional[Tuple[int, int, Any]]:
    """start행(열)부터 사용 범위 끝까지에서 확인 기준 칸 (앞쪽 LINES_SCAN행 / 200열만 훑음)."""
    ur1, uc1, ur2, uc2 = box
    if axis == "rows":
        s = max(start, ur1)
        return _first_constant(ws, s, uc1, min(ur2, s + LINES_SCAN - 1), uc2)
    s = max(start, uc1)
    return _first_constant(ws, ur1, s, min(ur2, ur1 + LINES_SCAN - 1), min(uc2, s + 199))


def _lines_command(args, axis: str, a: int, b: int, overwrite: bool) -> Dict[str, Any]:
    argv = ["range", "--target", args.target, "--action", args.action, "--range", args.range]
    if args.sheet:
        argv += ["--sheet", args.sheet]
    if args.axis:
        argv += ["--axis", args.axis]
    if overwrite:
        argv.append("--overwrite")
    what = "넣기" if args.action == "insert" else "지우기"
    return rt.next_command("excel.py", argv, f"{_lines_text(axis, a, b)} {what} 실행" + (" (값이 있는 칸도 지움)" if overwrite else ""),
                           destructive=args.action == "delete" or overwrite)


REF_SCAN_CELLS = 2_000_000   # 지울 칸을 참조하는 수식을 찾을 때 훑는 칸 수 상한 (통합문서 전체)
REF_SHOW = 20                # 결과에 담는 #REF! 될 수식 칸 주소 수

_FORMULA_STRING = re.compile(r'"(?:[^"]|"")*"')
_FORMULA_REF = re.compile(
    r"(?:'((?:[^']|'')+)'!|([^\s'!(),;:=+\-*/&^<>{}\"\[\]]+)!)?"          # 시트 이름 (따옴표 있거나 없거나)
    r"(?<![A-Za-z0-9_.$])"
    r"(?:\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?"            # A1, A1:B5
    r"|\$?([A-Z]{1,3}):\$?([A-Z]{1,3})"                                    # A:C
    r"|\$?(\d+):\$?(\d+))"                                                # 5:7
    r"(?![A-Za-z0-9_(!])")


def formula_refs(formula: str) -> List[Tuple[Optional[str], int, int, int, int]]:
    """수식의 A1 참조 [(시트 이름 또는 None, 행1, 열1, 행2, 열2)]. 따옴표 글은 빼고, 다른 통합문서([책]) 참조와
    이름 정의·INDIRECT 같은 간접 참조는 셀 수 없어 넣지 않습니다. 행·열 전체 참조는 시트 끝까지."""
    text = _FORMULA_STRING.sub('""', formula)
    out: List[Tuple[Optional[str], int, int, int, int]] = []
    for m in _FORMULA_REF.finditer(text):
        sheet = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
        if (sheet and ("[" in sheet or "]" in sheet)) or (m.start() > 0 and text[m.start() - 1] == "]"):
            continue                          # 다른 통합문서 ([책.xlsx]Sheet1!A1)
        if m.group(3):
            r1, c1 = int(m.group(4)), col_number(m.group(3))
            r2, c2 = (int(m.group(6)), col_number(m.group(5))) if m.group(5) else (r1, c1)
        elif m.group(7):
            r1, r2, c1, c2 = 1, 1048576, col_number(m.group(7)), col_number(m.group(8))
        else:
            r1, r2, c1, c2 = int(m.group(9)), int(m.group(10)), 1, 16384
        r1, r2 = sorted((r1, r2))
        c1, c2 = sorted((c1, c2))
        if r1 < 1 or r2 > 1048576 or c1 < 1 or c2 > 16384:
            continue
        out.append((sheet, r1, c1, r2, c2))
    return out


def broken_refs(wb: Any, ws: Any, axis: str, n1: int, n2: int) -> Tuple[List[str], List[str]]:
    """행·열을 지우면 #REF!가 될 수식 칸 (다른 칸의 수식이 지울 행·열 안에만 있는 칸이나 범위를 참조).
    범위 일부만 지워지면 범위가 줄어들 뿐 #REF!가 되지 않아 세지 않습니다. 지울 행·열 안의 수식 칸은 함께 사라져 빼고,
    같은 통합문서의 다른 시트에서 이 시트를 이름으로 참조하는 수식도 봅니다.
    반환: (["시트!주소", ...] 문서 순서, 훑지 못한 시트 이름). 수식 글은 돌려주지 않습니다(문서 내용)."""
    target = str(ws.Name)
    tkey = target.casefold()
    hits: List[str] = []
    skipped: List[str] = []
    budget = REF_SCAN_CELLS
    try:
        sheets = [wb.Worksheets(i) for i in range(1, int(wb.Worksheets.Count) + 1)]
    except Exception:
        sheets = [ws]
    for sh in sheets:
        name = str(sh.Name)
        same = name.casefold() == tkey
        try:
            used = sh.UsedRange
            r0, c0 = int(used.Row), int(used.Column)
            cells = int(used.Rows.Count) * int(used.Columns.Count)
            if cells > budget:
                skipped.append(name)
                continue
            budget -= cells
            forms = off.range_to_matrix(used.Formula)
        except Exception:
            skipped.append(name)
            continue
        for i, row in enumerate(forms):
            r = r0 + i
            for j, f in enumerate(row):
                if not (isinstance(f, str) and f.startswith("=")):
                    continue
                c = c0 + j
                if same and (n1 <= (r if axis == "rows" else c) <= n2):
                    continue                  # 지울 행·열 안의 수식은 함께 사라짐
                for sheet, ra, ca, rb, cb in formula_refs(f):
                    if (sheet.casefold() if sheet else name.casefold()) != tkey:
                        continue
                    lo, hi = (ra, rb) if axis == "rows" else (ca, cb)
                    if n1 <= lo and hi <= n2:
                        hits.append(f"{name}!{addr(r, c)}")
                        break
    return hits, skipped


def range_lines(args) -> Dict[str, Any]:
    """range --action insert|delete: 행·열 넣기·지우기. 지울 행·열에 값이 있으면 --overwrite가 있어야 지움 (값은 보이지 않음)."""
    a = args.action
    sheet, axis, n1, n2 = _lines_spec(args)
    n = n2 - n1 + 1
    what = _lines_text(axis, n1, n2)
    with write_session(args, "range") as t:
        ws = t.sheet(sheet) if sheet else t.sheet(args.sheet)
        block = ws.Range(_lines_addr(axis, n1, n2))
        lines = block.EntireRow if axis == "rows" else block.EntireColumn
        check_protection(ws, [block], f"{a}_{'rows' if axis == 'rows' else 'columns'}")   # 미리보기도 실제와 같은 검사 (1.4.83)
        box = _used_box(ws)
        filled: List[Tuple[str, Any]] = []
        if a == "delete" and box is not None:
            used = t.app.Intersect(block, ws.UsedRange)
            if used is not None:
                filled = _filled_cells(used)
        cells_text = ", ".join(c for c, _ in filled[:5]) + (" ..." if len(filled) > 5 else "")
        refs: List[str] = []
        ref_skipped: List[str] = []
        if a == "delete":
            refs, ref_skipped = broken_refs(t.doc, ws, axis, n1, n2)
        refs_text = ", ".join(refs[:5]) + (" ..." if len(refs) > 5 else "")
        refs_note = f"다른 칸의 수식 {len(refs)}개가 지울 칸을 참조해 #REF! 오류가 됩니다: {refs_text}" if refs else ""
        info: Dict[str, Any] = {"sheet": str(ws.Name), "range": _lines_addr(axis, n1, n2), "axis": axis, "count": n}
        if a == "delete":
            info["filled_cells"] = len(filled)
            info["broken_refs"] = len(refs)
            info["broken_ref_cells"] = refs[:REF_SHOW]
            if len(refs) > REF_SHOW:
                info["broken_ref_cells_truncated"] = True
            if ref_skipped:
                info["warnings"] = [f"칸이 너무 많아 지울 칸을 참조하는 수식을 확인하지 못한 시트: {', '.join(ref_skipped[:10])}"]
        if args.dry_run:
            # 바꾸지 않음: 되돌리기 지점도 남지 않음 (t.changed 없음)
            info["filled_addresses"] = [c for c, _ in filled[:20]]
            if a == "insert":
                shift = "아래 행이" if axis == "rows" else "오른쪽 열이"
                preview = f"미리보기: {what} 앞에 {n}{'행' if axis == 'rows' else '열'}을 넣으면 {shift} {n}칸씩 밀립니다 (바꾸지 않음)"
            elif filled:
                preview = f"미리보기: {what}을(를) 지우면 값이 있는 칸 {len(filled)}개가 함께 지워집니다: {cells_text}"
                preview += (f". {refs_note}" if refs else "") + " (바꾸지 않음)"
            elif refs:
                preview = f"미리보기: {what}에는 값이 있는 칸이 없지만 {refs_note} (바꾸지 않음)"
            else:
                preview = f"미리보기: {what}에는 값이 있는 칸이 없습니다. 지우면 {'아래 행이 올라' if axis == 'rows' else '오른쪽 열이 당겨'}옵니다 (바꾸지 않음)"
            count = info.pop("count")
            planned = [f"{'넣기' if a == 'insert' else '지우기'}: {ws.Name}!{_lines_addr(axis, n1, n2)} ({what})"]
            return rt.dry_run_result(TOOL, "range", preview, count, planned,
                                     _lines_command(args, axis, n1, n2, overwrite=bool(filled or refs)),
                                     target_resolved=t.resolved(), **info)
        if (filled or refs) and not args.overwrite:
            head = f"{ws.Name}!{_lines_addr(axis, n1, n2)}({what})을(를) 지우면 "
            parts = [f"값이 있는 칸 {len(filled)}개가 함께 지워집니다: {cells_text} (값은 보이지 않음)"] if filled else []
            log_parts = [f"값이 있는 칸 {len(filled)}개 ({cells_text})"] if filled else []
            if refs:
                parts.append(refs_note)
                log_parts.append(f"#REF! 될 수식 칸 {len(refs)}개 ({refs_text})")
            raise rt.UserDecisionError(
                head + ". ".join(parts),
                "지우지 않았습니다. " + ("지워질 칸 수와 위치" if filled else "") + (", " if filled and refs else "")
                + ("#REF! 오류가 될 수식 칸 수와 위치" if refs else "") + "를 사용자에게 알리고, 사용자가 지워도 된다고 하면 "
                "--overwrite를 붙여 다시 실행하세요 (미리 보려면 --dry-run).",
                log_message=f"{_lines_addr(axis, n1, n2)}({what}) 지우기: " + ", ".join(log_parts))
        # 확인 기준: 옮겨질 칸 하나 (넣기는 넣는 자리부터, 지우기는 지운 자리 뒤부터)
        ref = _ref_cell(ws, axis, n1 if a == "insert" else n2 + 1, box) if box is not None else None
        if a == "insert":
            lines.Insert()
        else:
            lines.Delete()
        t.set_focus(sheet=ws)
        # R1: 사용 범위 끝과 옮겨진 칸을 다시 읽어 확인
        problems: List[str] = []
        moved: Optional[Dict[str, str]] = None
        try:
            after = _used_box(ws)
            idx = 2 if axis == "rows" else 3
            if box is not None:
                last = box[idx]
                if a == "insert":
                    want = last + n if last >= n1 else last
                    got = after[idx] if after else 0
                    if got != want:
                        problems.append(f"사용 범위 끝이 {got}{'행' if axis == 'rows' else '열'} (기대 {want})")
                else:
                    got = after[idx] if after else 0
                    if last > n2 and got != last - n:
                        problems.append(f"사용 범위 끝이 {got}{'행' if axis == 'rows' else '열'} (기대 {last - n})")
                    elif last <= n2 and got >= n1:
                        left = _filled_cells(ws.Range(_lines_addr(axis, n1, n2)))
                        if left:
                            problems.append(f"지운 자리에 값이 남은 칸 {len(left)}개")
            if ref is not None:
                rr, rc, v = ref
                d = n if a == "insert" else -n
                nr, nc = (rr + d, rc) if axis == "rows" else (rr, rc + d)
                moved = {"from": addr(rr, rc), "to": addr(nr, nc)}
                if not _same_cell(v, ws.Cells(nr, nc).Value):
                    problems.append(f"{addr(rr, rc)}의 값이 {addr(nr, nc)}로 옮겨지지 않음")
        except Exception as e:
            problems.append(f"결과를 다시 읽지 못함: {type(e).__name__}")
        if moved:
            info["moved"] = moved
        resolved = t.resolved()
    if a == "insert":
        preview = f"{info['sheet']}!{what} 앞에 {n}{'행' if axis == 'rows' else '열'} 넣음 ({'아래 행' if axis == 'rows' else '오른쪽 열'}은 {n}칸씩 밀림)"
    else:
        preview = f"{info['sheet']}!{what} 지움" + (f" (값이 있던 칸 {len(filled)}개 포함)" if filled else "")
        if refs:
            preview += f". 수식 {len(refs)}개가 #REF! 오류가 됨: {refs_text}"
    if problems:
        return rt.ok_unverified(TOOL, "range", preview, problems, target_resolved=resolved, **info)
    return rt.ok(TOOL, "range", preview, "VERIFIED: 사용 범위와 옮겨진 칸을 다시 읽어 확인했습니다. 잘못이면 office.py undo로 되돌립니다.",
                 verified=True, target_resolved=resolved, **info)


# --- 정렬과 --where가 함께 쓰는 열 찾기 ---

_LETTERS = re.compile(r"^[A-Za-z]{1,3}$")


def resolve_column(key: str, header: Optional[List[Any]], c0: int, ncols: int, what: str = "열",
                   option: Any = None) -> Tuple[int, bool]:
    """--by/--where의 열 -> (범위 안 열 위치 0부터, 머리글 이름으로 찾았는지).
    머리글(header)이 있으면 이름(같은 이름, 없으면 rt.name_key로 표기만 다른 이름)을 먼저 찾고, 없으면 열 문자(B, AA). 찾지 못하면 [후보] 머리글 이름."""
    k = str(key).strip()
    names = [("" if v is None else str(out_value(v))).strip() for v in (header or [])]
    if header is not None:
        exact = [i for i, nm in enumerate(names) if nm and rt.exact_key(nm) == rt.exact_key(k)]
        if not exact:
            nk = rt.name_key(k)
            exact = [i for i, nm in enumerate(names) if nm and rt.name_key(nm) == nk]
        if len(exact) == 1:
            return exact[0], True
        if len(exact) > 1:
            raise rt.ambiguous(f"머리글에 '{k}' 열이 여러 개입니다. 열 문자로 고르세요.",
                               [f"{col_letter(c0 + i)} ({names[i]})" for i in exact],
                               log_message=f"머리글에 같은 이름의 {what}이 {len(exact)}개",
                               replace=k, values=[col_letter(c0 + i) for i in exact], option=option)
    if _LETTERS.match(k):
        idx = col_number(k) - c0
        if 0 <= idx < ncols:
            return idx, False
        raise rt.InputError(f"{what} {k.upper()}은(는) 범위({col_letter(c0)}~{col_letter(c0 + ncols - 1)}열) 밖입니다.")
    shown = [nm for nm in names if nm]
    if header is None or not shown:
        raise rt.InputError(f"{what} '{k}'을(를) 찾을 수 없습니다. 머리글 행이 없으면 열 문자(예: B)를 쓰세요.")
    near = rt.suggest(k, shown, 5)
    raise rt.ambiguous(f"{what} '{k}'을(를) 머리글에서 찾을 수 없습니다. 머리글 이름이나 열 문자(예: B)를 쓰세요.", near or shown,
                       log_message=f"{what} 이름을 머리글에서 찾을 수 없음", replace=k, option=option)


def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_date(v: Any) -> bool:
    return isinstance(v, (_dt.datetime, _dt.date)) or (hasattr(v, "year") and hasattr(v, "isoformat"))


def _looks_header(matrix: List[List[Any]], first_bold: bool) -> bool:
    """첫 행이 머리글인지 (--header auto, --by가 열 문자일 때): 첫 행 칸이 모두 글자이고,
    굵게이거나 아래 행에 숫자·날짜가 있는 열이 있으면 머리글."""
    if len(matrix) < 2:
        return False
    first = matrix[0]
    if not first or any(not isinstance(v, str) or not v.strip() for v in first):
        return False
    if first_bold:
        return True
    for j in range(len(first)):
        if any(j < len(r) and (_is_num(r[j]) or _is_date(r[j]) or isinstance(r[j], bool)) for r in matrix[1:]):
            return True
    return False


def _date_num(v: Any) -> float:
    d = v if isinstance(v, _dt.datetime) else _dt.datetime(v.year, v.month, v.day)
    d = d.replace(tzinfo=None)
    return (d - _dt.datetime(1899, 12, 30)).total_seconds() / 86400.0


def _sort_class(v: Any) -> Optional[Tuple[int, Any]]:
    """Excel 정렬 순서 (오름차순): 숫자·날짜 < 글자(대소문자 무시) < 논리값. 빈 칸은 None (오름·내림 모두 맨 뒤)."""
    if v is None or (isinstance(v, str) and v == ""):
        return None
    if isinstance(v, bool):
        return (2, v)
    if _is_num(v):
        return (0, float(v))
    if _is_date(v):
        try:
            return (0, _date_num(v))
        except Exception:
            return (3, str(v))
    if isinstance(v, str):
        return (1, re.sub(r"[-']", "", v).casefold())       # Excel은 하이픈·작은따옴표를 무시하고 비교
    return (3, str(v))


def _sort_violation(rows: List[List[Any]], keys: List[Tuple[int, bool]]) -> Optional[int]:
    """정렬 순서가 맞지 않는 첫 행 위치 (0부터, 다음 행과 비교). 모두 맞으면 None."""
    for i in range(len(rows) - 1):
        for k, desc in keys:
            a = _sort_class(rows[i][k] if k < len(rows[i]) else None)
            b = _sort_class(rows[i + 1][k] if k < len(rows[i + 1]) else None)
            if a is None and b is None:
                continue
            if a is None:                  # 빈 칸 뒤에 값
                return i
            if b is None:
                break
            if a == b:
                continue
            if (a < b) != (not desc):
                return i
            break
    return None


def _row_key(row: List[Any]) -> str:
    return json.dumps([out_value(v) for v in row], ensure_ascii=False, default=str)


def _adjacent_filled(ws: Any, r1: int, r2: int, col: int) -> bool:
    if col < 1 or col > XL_MAX_COL:
        return False
    return bool(_filled_cells(ws.Range(ws.Cells(r1, col), ws.Cells(r2, col))))


def _sort_fields(ws: Any, rng: Any, top: int, c1: int, r2: int, order: List[Tuple[int, bool]]) -> None:
    """기준 4개 이상: Worksheet.Sort.SortFields에 기준 열(데이터 행)을 차례로 넣고 Apply. rng는 머리글을 뺀 데이터 행.
    시트에 남는 정렬 기준은 지웁니다."""
    srt = ws.Sort
    srt.SortFields.Clear()
    for idx, desc in order:
        col = ws.Range(ws.Cells(top, c1 + idx), ws.Cells(r2, c1 + idx))
        srt.SortFields.Add(Key=col, SortOn=0, Order=XL_DESC if desc else XL_ASC, DataOption=0)   # xlSortOnValues, 보통 비교
    srt.SetRange(rng)
    srt.Header = XL_NO
    srt.MatchCase = False
    srt.Orientation = 1                     # xlTopToBottom (행 단위)
    try:
        srt.Apply()
    finally:
        try:
            srt.SortFields.Clear()
        except Exception:
            pass


def range_sort(args) -> Dict[str, Any]:
    """range --action sort: --by 열(이름 또는 열 문자, 쉼표로 여럿)로 정렬. 머리글 행은 도구가 정하고 결과에 알림.
    기준 3개까지는 Range.Sort, 4개부터는 Worksheet.Sort.SortFields."""
    if not args.range:
        raise rt.InputError("range sort에는 --range가 필요합니다 (예: --range A1:E100, 머리글 행 포함).")
    if not args.by:
        raise rt.InputError("range sort에는 --by가 필요합니다 (열 이름 또는 열 문자, 예: --by \"금액\", --by \"부서,금액\").")
    keys = [k.strip() for k in str(args.by).split(",") if k.strip()]
    if not keys:
        raise rt.InputError("--by가 비어 있습니다 (예: --by \"금액\").")
    if len(keys) > SORT_MAX_KEYS:
        raise rt.InputError(f"정렬 기준(--by)은 {SORT_MAX_KEYS}개까지입니다 (받은 값: {len(keys)}개).")
    with write_session(args, "range") as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range)
        ws = rng.Worksheet
        if "," in str(rng.Address):
            raise rt.InputError("정렬 범위는 한 영역이어야 합니다 (예: A1:E100).")
        used = t.app.Intersect(rng, ws.UsedRange)        # 열 전체(A:E)는 사용 범위까지만
        if used is None:
            raise rt.InputError(f"{args.range}에 정렬할 값이 없습니다 (사용 범위 밖).")
        rng = used
        check_protection(ws, [rng], "sort")              # 보호된 시트: 잠긴 칸이 있거나 정렬이 허용되지 않으면 멈춤
        r1, c1 = int(rng.Row), int(rng.Column)
        nr, nc = int(rng.Rows.Count), int(rng.Columns.Count)
        r2, c2 = r1 + nr - 1, c1 + nc - 1
        where = f"{ws.Name}!{addr(r1, c1, r2, c2)}"
        if rng.MergeCells is not False:        # 병합 칸이 있으면 True(전부) 또는 None(일부)
            raise rt.InputError(f"{where}에 병합된 칸이 있어 정렬할 수 없습니다. 병합을 풀어도 되는지 사용자에게 묻고, 원하면 "
                                f"excel.py range --action unmerge --range {addr(r1, c1, r2, c2)} 뒤에 다시 정렬하세요.")
        # 표를 가로지르는 범위: 바로 옆 열의 같은 행에 값이 있으면 정렬한 열만 행이 바뀌어 표가 어긋남
        lo, hi = c1, c2
        while lo > 1 and hi - lo < 200 and _adjacent_filled(ws, r1, r2, lo - 1):
            lo -= 1
        while hi < XL_MAX_COL and hi - lo < 200 and _adjacent_filled(ws, r1, r2, hi + 1):
            hi += 1
        if (lo, hi) != (c1, c2):
            full = addr(r1, lo, r2, hi)
            raise rt.UserDecisionError(
                f"정렬 범위 {where} 옆 열({col_letter(lo)}~{col_letter(hi)}열)에도 같은 행의 값이 있습니다. 이 범위만 정렬하면 "
                f"표의 행이 어긋납니다. 표 전체는 {full}입니다.",
                f"정렬하지 않았습니다. 사용자에게 표 전체({full})를 정렬할지 묻고, 원하면 --range {full}로 다시 실행하세요.")
        matrix = off.range_to_matrix(rng.Value)
        first = matrix[0] if matrix else []
        header_given = args.header
        header_row = None if header_given == "no" else first
        cols: List[Tuple[int, bool]] = []
        by_name = False
        for k in keys:
            idx, named = resolve_column(k, header_row, c1, nc, "정렬 기준 열", option=("--by", "--desc"))
            by_name = by_name or named
            cols.append((idx, named))
        if header_given == "yes" or by_name:
            has_header = True
            header_how = "--header yes" if header_given == "yes" else "--by가 머리글 이름"
        elif header_given == "no":
            has_header, header_how = False, "--header no"
        else:
            try:
                bold = ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2)).Font.Bold is True
            except Exception:
                bold = False
            has_header = _looks_header(matrix, bold)
            header_how = "자동 판단 (" + ("첫 행이 글자 머리글" if has_header else "첫 행도 데이터") + ")"
        # --desc: 없으면 모두 오름차순, 값 없이 주면 모두 내림차순, 이름을 주면 그 기준만 내림차순
        desc_keys: Optional[set] = None
        if args.desc is not None and args.desc != "*":
            desc_keys = set()
            for d in [x.strip() for x in str(args.desc).split(",") if x.strip()]:
                hit = [i for i, k in enumerate(keys) if rt.name_key(k) == rt.name_key(d)]
                if not hit:
                    raise rt.ambiguous(f"--desc '{d}'이(가) --by에 없습니다. --by에 쓴 기준 중에서 고르세요.", keys,
                                       log_message="--desc 기준이 --by에 없음", replace=d)
                desc_keys.update(hit)
        order = [(idx, args.desc == "*" or (desc_keys is not None and i in desc_keys)) for i, (idx, _) in enumerate(cols)]
        data = matrix[1:] if has_header else matrix
        info: Dict[str, Any] = {"sheet": str(ws.Name), "range": addr(r1, c1, r2, c2), "rows": len(data), "header": has_header,
                                "header_decided_by": header_how,
                                "keys": [{"column": col_letter(c1 + idx), "name": str(out_value(first[idx])) if has_header else None,
                                          "order": "desc" if desc else "asc"} for idx, desc in order]}
        if len(data) < 2 or _sort_violation(data, order) is None:
            t.set_focus(sheet=ws, changed=False)        # 이미 정렬됨: 바꾸지 않음 (되돌리기 지점을 남기지 않음)
            info["unchanged"] = True
            resolved = t.resolved()
            return rt.ok(TOOL, "range", f"{where}: 이미 정렬되어 있어 바꾸지 않았습니다 ({len(data)}행)",
                         "VERIFIED: 정렬 순서를 다시 읽어 확인했고 바뀐 것이 없습니다.", target_resolved=resolved, **info)
        # 머리글 행을 빼고 데이터 행만 정렬 (Header=xlNo): Header=xlYes를 줘도 실제 엑셀이 머리글을 데이터에 섞은 일
        # (회사 PC 2026-11-10 L1, 통화 서식 열). 머리글을 범위에 넣지 않으면 엑셀의 머리글 판단과 무관
        top = r1 + 1 if has_header else r1
        data_rng = ws.Range(ws.Cells(top, c1), ws.Cells(r2, c2))
        if len(order) <= SORT_RANGE_KEYS:
            kw: Dict[str, Any] = {}
            for i, (idx, desc) in enumerate(order, start=1):
                kw[f"Key{i}"] = ws.Cells(top, c1 + idx)
                kw[f"Order{i}"] = XL_DESC if desc else XL_ASC
            data_rng.Sort(Header=XL_NO, MatchCase=False, Orientation=1, **kw)
        else:
            _sort_fields(ws, data_rng, top, c1, r2, order)
        t.set_focus(sheet=ws)
        # R1: 다시 읽어 행 수·머리글·행 묶음이 그대로이고 기준 열이 정렬됐는지
        problems: List[str] = []
        try:
            got = off.range_to_matrix(ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value)
            got_data = got[1:] if has_header else got
            if len(got) != len(matrix):
                problems.append(f"행 수가 {len(got)} (정렬 전 {len(matrix)})")
            elif has_header and _row_key(got[0]) != _row_key(first):
                problems.append("머리글 행이 바뀜")
            elif sorted(map(_row_key, got_data)) != sorted(map(_row_key, data)):
                problems.append("정렬 전과 행 내용이 다름 (행 번호를 쓰는 수식 등)")
            bad = _sort_violation(got_data, order)
            if bad is not None:
                problems.append(f"{r1 + (1 if has_header else 0) + bad}행 근처의 순서가 기준과 다름 (Excel 비교 규칙 차이일 수 있음)")
        except Exception as e:
            problems.append(f"결과를 다시 읽지 못함: {type(e).__name__}")
        resolved = t.resolved()
    keys_text = ", ".join(f"{k['name'] or k['column']} {'내림차순' if k['order'] == 'desc' else '오름차순'}" for k in info["keys"])
    preview = f"{where} 정렬: {keys_text} ({len(data)}행, 머리글 {'있음' if has_header else '없음'}: {header_how})"
    if problems:
        return rt.ok_unverified(TOOL, "range", preview, problems, target_resolved=resolved, **info)
    return rt.ok(TOOL, "range", preview, "VERIFIED: 다시 읽어 정렬 순서와 행 수를 확인했습니다. 잘못이면 office.py undo로 되돌립니다.",
                 verified=True, target_resolved=resolved, **info)


# ==============================================================================
# 9. chart
# ==============================================================================

def _open_receivers() -> List[str]:
    """차트를 받을 앱 중 문서가 열린 앱 (next_commands의 --to). 둘 다 열렸거나 둘 다 없거나 확인하지 못하면 word,ppt
    (열린 앱만 주면 열리지 않은 앱 때문에 transfer가 멈추지 않음. 확인하지 못해도 transfer가 넣기 전에 다시 확인함)."""
    try:
        found = [k for k in ("word", "ppt") if off.list_open(k)]
    except Exception as e:
        rt.debug(f"열린 워드/PPT 확인 생략: {e}")
        found = []
    return found or ["word", "ppt"]


def _chart_objects(ws: Any) -> List[Any]:
    objs = ws.ChartObjects()
    return [objs.Item(i) for i in range(1, objs.Count + 1)]


def _chart_title(obj: Any) -> str:
    try:
        return str(obj.Chart.ChartTitle.Text) if obj.Chart.HasTitle else ""
    except Exception:
        return ""


_QUOTED_CHART = re.compile(r"^'((?:[^']|'')+)'!(.+)$", re.S)
_INDEX_KEY = re.compile(r"^#?(\d+)$")


def _chart_matches(ws: Any, key: str) -> Tuple[List[Tuple[Any, Any, int]], List[Tuple[Any, Any, int]], List[str]]:
    """시트 하나에서 이름이 같은 차트, 제목이 같은 차트, 보이는 이름 목록 (대소문자 무시). 항목은 (시트, 차트, 번호)."""
    by_name: List[Tuple[Any, Any, int]] = []
    by_title: List[Tuple[Any, Any, int]] = []
    names: List[str] = []
    k = rt.exact_key(key)
    for idx, obj in enumerate(_chart_objects(ws), 1):
        nm, title = str(obj.Name), _chart_title(obj)
        names += [x for x in (nm, title) if x]
        if rt.exact_key(nm) == k:
            by_name.append((ws, obj, idx))
        elif title and rt.exact_key(title) == k:
            by_title.append((ws, obj, idx))
    return by_name, by_title, names


def _chart_by_index(ws: Any, key: str) -> List[Tuple[Any, Any, int]]:
    """'3' 또는 '#3' = 그 시트의 3번째 차트."""
    m = _INDEX_KEY.match(key.strip())
    if not m:
        return []
    objs = _chart_objects(ws)
    i = int(m.group(1))
    return [(ws, objs[i - 1], i)] if 1 <= i <= len(objs) else []


def _chart_labels(hits: List[Tuple[Any, Any, int]]) -> List[str]:
    """[후보]에 줄 '시트!이름'. 그 시트에 이름이 같은 차트가 하나라도 더 있으면(후보가 아닌 차트 포함) '시트!#번호',
    시트·차트 이름에 '!'나 작은따옴표가 있으면 모두 "'시트'!이름" (그대로 다시 주면 그 차트 하나로 찾아짐)."""
    count: Dict[Tuple[str, str], int] = {}
    done: set = set()
    for ws, _, _ in hits:
        if str(ws.Name) in done:
            continue
        done.add(str(ws.Name))
        for obj in _chart_objects(ws):                  # 후보만이 아니라 그 시트의 모든 차트로 셈
            k = (str(ws.Name), str(obj.Name).strip().lower())
            count[k] = count.get(k, 0) + 1
    parts = [(str(ws.Name), f"#{idx}" if count[(str(ws.Name), str(obj.Name).strip().lower())] > 1 else str(obj.Name))
             for ws, obj, idx in hits]
    quote = any("!" in w or "!" in c or "'" in w for w, c in parts)
    return [("'" + w.replace("'", "''") + "'" if quote else w) + "!" + c for w, c in parts]


def find_chart(t: off.OfficeTarget, name: Optional[str], sheet: Optional[str] = None) -> Tuple[Any, Any]:
    """차트 이름/제목(대소문자 무시) 또는 번호('3', '#3')로 차트를 찾습니다. (시트, 차트개체)
    시트를 생략하면 모든 시트에서 찾고, '시트!이름'과 "'시트'!이름"(작은따옴표 안의 ''는 ') 형식도 받습니다.
    - 따옴표 형식은 그 시트에서만 찾습니다.
    - 따옴표 없는 '시트!이름'은 '!' 자리마다 나눠 보고 앞부분이 있는 시트이고 그 시트에서 찾아질 때만 씁니다.
      글자 그대로를 차트 이름으로도 찾아, 둘 다 찾아지면 하나를 고르지 않고 [후보] 오류를 냅니다.
    이름이 같은 차트를 제목이 같은 차트보다 우선하고, 그래도 둘 이상이면(한 시트 안에서 이름이 같은 경우 포함)
    하나를 고르지 않고 [후보] 오류를 냅니다 (엉뚱한 차트를 옮기거나 지우지 않게)."""
    wb = t.doc
    all_sheets = [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
    sheet_names = [str(ws.Name) for ws in all_sheets]
    scope = [t.sheet(sheet)] if sheet else all_sheets
    sheet_key = str(sheet).strip().strip("'").lower() if sheet else None
    where = f"시트 '{scope[0].Name}'" if sheet else "통합 문서"
    if name is None:
        for ws in scope:
            objs = _chart_objects(ws)
            if objs:
                return ws, objs[0]
        raise rt.InputError(f"{where}에 차트가 없습니다. excel.py chart --range ... 로 먼저 만드세요.")

    def sheet_by_name(n: str) -> Optional[Any]:
        return next((ws for ws in all_sheets if rt.exact_key(ws.Name) == rt.exact_key(n)), None)

    def mismatch(left: str) -> rt.InputError:
        return rt.InputError(f"차트 '{name}'의 시트 '{left}'와 --sheet '{sheet}'가 다릅니다. 아무것도 바꾸지 않았습니다. "
                             "--sheet를 빼거나 같은 시트로 맞추세요.")

    text = str(name).strip()
    by_name: List[Tuple[Any, Any, int]] = []
    by_title: List[Tuple[Any, Any, int]] = []
    by_index: List[Tuple[Any, Any, int]] = []
    all_names: List[str] = []
    m = _QUOTED_CHART.match(text)
    if m:
        left, right = m.group(1).replace("''", "'"), m.group(2).strip()
        ws = sheet_by_name(left)
        if ws is None:
            raise rt.not_found("시트", left, sheet_names)
        if sheet_key and sheet_key != left.lower():
            raise mismatch(left)
        by_name, by_title, all_names = _chart_matches(ws, right)
        by_index = _chart_by_index(ws, right)
        where = f"시트 '{ws.Name}'"
    else:
        contradicted = ""
        if "!" in text:
            for pos in [i for i, ch in enumerate(text) if ch == "!"]:
                left, right = text[:pos].strip().strip("'"), text[pos + 1:].strip()
                ws = sheet_by_name(left)
                if not right or ws is None:
                    continue
                if sheet_key and sheet_key != left.lower():
                    contradicted = contradicted or str(ws.Name)
                    continue
                n_hits, t_hits, _ = _chart_matches(ws, right)       # 찾아지는 나눔만 씀
                by_name += n_hits
                by_title += t_hits
                by_index += _chart_by_index(ws, right)
        else:
            by_index = _chart_by_index(scope[0], text)
        for ws in scope:                                            # 글자 그대로를 차트 이름으로
            n_hits, t_hits, names = _chart_matches(ws, text)
            by_name += n_hits
            by_title += t_hits
            all_names += names
        if not (by_name or by_title or by_index) and contradicted:
            raise mismatch(contradicted)
    seen: set = set()
    hits = []
    for h in (by_name or by_title or by_index):
        k = (str(h[0].Name), h[2])
        if k not in seen:
            seen.add(k)
            hits.append(h)
    if len(hits) > 1:
        labels = _chart_labels(hits)
        raise rt.ambiguous(
            f"'{name}'에 해당하는 차트가 {len(hits)}개입니다 (이름 또는 제목이 같음). 하나를 고르지 않았고 아무것도 바꾸지 않았습니다. "
            "어느 차트인지 사용자에게 확인한 뒤 차트 이름 자리에 후보를 글자 그대로 주세요.", labels, replace=str(name), option="--name")
    if hits:
        return hits[0][0], hits[0][1]
    if not all_names:
        raise rt.InputError(f"{where}에 차트가 없습니다. excel.py chart --range ... 로 먼저 만드세요.")
    raise rt.not_found("차트", str(name), sorted(set(all_names)), extra=f" ({where})" if where != "통합 문서" else "", option="--name")


def _chart_delete_dry(t: Any, args) -> Dict[str, Any]:
    """chart --action delete --dry-run: 지울 차트 이름만 보여 줌 (바꾸지 않음). 보호 확인은 읽기만 해서 실제 실행과 같이 함."""
    if args.all:
        ws = t.sheet(args.sheet)
        names = [str(o.Name) for o in _chart_objects(ws)]
        if names:
            check_protection(ws, [], "objects")
            preview = f"미리보기: 시트 '{ws.Name}'의 차트 {len(names)}개를 삭제합니다: " + ", ".join(names[:5]) + (" ..." if len(names) > 5 else "")
        else:
            preview = f"미리보기: 시트 '{ws.Name}'에 차트가 없어 바뀔 것 없음 (0개)"
    else:
        ws, obj = find_chart(t, args.name, args.sheet)
        check_protection(ws, [], "objects")
        names = [str(obj.Name)]
        preview = f"미리보기: 시트 '{ws.Name}'의 차트 '{names[0]}'을(를) 삭제합니다 (1개)"
    return rt.dry_run_result(TOOL, "chart", preview + " (바꾸지 않음)", len(names), [f"삭제: {ws.Name}!{n}" for n in names],
                      rt.dry_run_command(TOOL, "차트 삭제 실행"), target_resolved=t.resolved(), sheet=str(ws.Name))


def _chart_list(args) -> Dict[str, Any]:
    with off.open_target("excel", args.target, readonly=True) as t:
        wb = t.doc
        charts = [{"sheet": str(wb.Worksheets(i).Name), "name": str(o.Name), "title": _chart_title(o)}
                  for i in range(1, wb.Worksheets.Count + 1) for o in _chart_objects(wb.Worksheets(i))]
        resolved = t.resolved()
    if not charts:
        return rt.ok(TOOL, "chart", f"0건: {resolved['name']}에 차트가 없습니다",
                     f"VERIFIED: 0건. {resolved['name']}에 차트가 없습니다. 사용자에게 그대로 알리세요.",
                     target_resolved=resolved, charts=charts, count=0)
    return rt.ok(TOOL, "chart", f"차트 {len(charts)}개", target_resolved=resolved, charts=charts, count=len(charts))


def cmd_chart(args) -> Dict[str, Any]:
    if args.action == "list":
        return rt.dry_run_read_only(_chart_list(args), args.dry_run, "chart list")
    if args.dry_run and args.action != "delete":
        raise rt.dry_run_unsupported(TOOL, f"chart --action {args.action}")
    if args.action == "delete":
        if not args.name and not args.all:
            raise rt.InputError("삭제할 차트를 --name(이름 또는 제목)으로 지정하거나, 시트의 모든 차트를 지우려면 --all을 쓰세요.")
        with write_session(args, "chart") as t:
            if args.dry_run:
                return _chart_delete_dry(t, args)
            if args.all:
                ws = t.sheet(args.sheet)
                n = len(_chart_objects(ws))
                if n:
                    check_protection(ws, [], "objects")
                    ws.ChartObjects().Delete()
                msg = f"시트 '{ws.Name}'의 차트 {n}개를 삭제했습니다"
            else:
                ws, obj = find_chart(t, args.name, args.sheet)
                check_protection(ws, [], "objects")
                gone = str(obj.Name)
                before = sum(1 for o in _chart_objects(ws) if str(o.Name) == gone)   # 같은 이름이 여럿일 수 있음
                msg = f"차트 '{gone}'을(를) 삭제했습니다"
                obj.Delete()
            left = [str(o.Name) for o in _chart_objects(ws)]            # R1: 차트 목록을 다시 읽음
            t.set_focus(sheet=ws, changed=not args.all or bool(n))      # 0개 삭제는 바뀐 것 없음
            resolved = t.resolved()
        if (args.all and left) or (not args.all and left.count(gone) >= before):
            return rt.ok_unverified(TOOL, "chart", msg, [f"다시 읽은 차트 목록에 남아 있음: {', '.join(left[:5])}"], target_resolved=resolved)
        return rt.ok(TOOL, "chart", msg, "VERIFIED: 차트 목록을 다시 읽어 확인했습니다.", verified=True, target_resolved=resolved)

    # create
    theme = off.get_theme(args.theme)
    with write_session(args, "chart") as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range)
        ws = rng.Worksheet
        if args.range is None and int(rng.Rows.Count) * int(rng.Columns.Count) <= 1:
            raise rt.InputError("차트로 만들 데이터가 없습니다. --range로 데이터 범위를 지정하세요.")
        check_protection(ws, [], "objects")              # 보호된 시트에서 개체(차트) 편집이 막혀 있으면 만들기 전에
        values = off.range_to_matrix(rng.Value)
        chart_type = args.type
        # 첫 열이 숫자(x값)인 꺾은선은 분산형(선)으로 그려야 x축이 숫자 축이 됨 (예: 사인 그래프)
        if chart_type == "line" and len(values) > 2 and len(values[0]) >= 2:
            first_col = [r[0] for r in values[1:] if r[0] is not None]
            if first_col and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in first_col):
                chart_type = "scatter_lines"
        if args.at:
            if not is_cell_address(args.at):
                raise rt.InputError(f"--at '{args.at}'은(는) 칸 주소가 아닙니다. 차트는 데이터가 있는 시트에 만들며 --at에는 그 시트의 칸만 씁니다 "
                                    "(예: --at E1). 아무것도 바꾸지 않았습니다.")
            anchor = ws.Range(args.at)
        else:
            # 기본 위치: 데이터 범위 오른쪽, 단 시트의 다른 데이터도 가리지 않도록 사용 범위 오른쪽 끝 다음 열
            last_col = int(rng.Column) + int(rng.Columns.Count) - 1
            try:
                used = ws.UsedRange
                last_col = max(last_col, int(used.Column) + int(used.Columns.Count) - 1)
            except Exception:
                pass
            anchor = ws.Cells(int(rng.Row), last_col + 2)
        obj = ws.ChartObjects().Add(anchor.Left, anchor.Top, args.width, args.height)
        chart = obj.Chart
        # 종류 -> 데이터 -> 종류: 데이터 앞에 정해야 분산형이 첫 열을 x값으로 쓰고(뒤에만 정하면 x가 계열이 됨),
        # 데이터 뒤에 다시 정해야 세부 종류가 유지됨 (SetSourceData가 분산형을 선+표식으로 바꿀 수 있음)
        chart.ChartType = XL_CHART[chart_type]
        chart.SetSourceData(rng)
        chart.ChartType = XL_CHART[chart_type]
        if args.title:
            chart.HasTitle = True
            chart.ChartTitle.Text = args.title
        series = chart.SeriesCollection()
        if chart_type == "combo" and series.Count >= 2:
            series(1).ChartType = 51
            series(2).ChartType = 65
            series(2).AxisGroup = 2
        if args.labels:
            chart.ApplyDataLabels()
        if args.legend == "none" or (args.legend is None and series.Count == 1):
            chart.HasLegend = False
        else:
            chart.HasLegend = True
            chart.Legend.Position = XL_LEGEND[args.legend or "bottom"]
        if theme:
            colors = [off.parse_color(theme[k]) for k in ("primary", "secondary", "accent")]
            for i in range(1, min(series.Count, 3) + 1):
                s = series(i)
                if chart_type in ("line", "line_markers", "scatter_lines", "scatter") or (chart_type == "combo" and i == 2):
                    s.Format.Line.ForeColor.RGB = colors[i - 1]
                else:
                    s.Format.Fill.ForeColor.RGB = colors[i - 1]
        existing = [str(o.Name) for o in _chart_objects(ws) if o.Name != obj.Name]
        wanted = (args.name or args.title or "").strip()
        if wanted:
            final = wanted
            k = 2
            while rt.find_name(final, existing):
                final = f"{wanted} {k}"
                k += 1
            obj.Name = final
        chart_name = str(obj.Name)
        problems = []                                                   # R1: 차트 목록과 제목을 다시 읽음
        names = [str(o.Name) for o in _chart_objects(ws)]
        if chart_name not in names:
            problems.append(f"다시 읽은 차트 목록에 '{chart_name}'이(가) 없음")
        elif args.title and _chart_title(obj) != args.title:
            problems.append(f"차트 제목이 '{_chart_title(obj)}' (준 제목 '{args.title}')")
        t.set_focus(sheet=ws)
        info = range_info(rng)
        resolved = t.resolved()
        tgt = t.path if t.source in ("file", "new") else args.target
        sheet_name = str(ws.Name)            # 세션 안에서 (닫힌 파일은 세션이 끝나면 문서를 닫음)
    preview = f"차트 '{chart_name}' 생성 ({chart_type}, 데이터 {info['sheet']}!{info['range']})"
    to = _open_receivers()
    fields = dict(target_resolved=resolved, chart_name=chart_name, chart_type=chart_type, data_range=info,
                  confirm_commands=[rt.next_command("office.py", ["transfer", "--from-app", "excel", "--from-target", tgt,
                                                               "--sheet", sheet_name, "--chart", chart_name,
                                                               "--to-app", ",".join(to)],
                                                 f"이 차트를 열린 {'와 '.join(off.APP_LABEL[k] for k in to)}에 그림으로 넣기"
                                                 + (" (한쪽만이면 --to-app word 또는 --to-app ppt)" if len(to) > 1 else ""))])
    if problems:
        return rt.ok_unverified(TOOL, "chart", preview, problems, **fields)
    return rt.ok(TOOL, "chart", preview, "VERIFIED: 차트를 다시 읽어 확인했습니다. 사용자가 워드/PPT로 옮기기를 원하면 confirm_commands[0]을 1번 실행하세요(열린 문서를 바꿈).",
                 verified=True, **fields)


# ==============================================================================
# 10. comment (칸 메모: 이전 방식 메모, Range.Comment / AddComment)
# ==============================================================================

COMMENT_LIST_MAX = 100        # 목록 한 번에 보여 줄 메모 수 (넘으면 --start로 이어 보기)
COMMENT_TEXT_MAX = 1000       # 목록에 담는 메모 하나의 글자 수
COMMENT_MAX_CHARS = 32767     # 엑셀 메모 글자 수 상한
COMMENT_MAX_CELLS = 500       # 한 번에 달 수 있는 메모 수
_MEMO_HEADER = ("칸", "셀", "cell", "주소", "위치", "range", "address")


def _note_text(c: Any) -> str:
    """메모 글 (엑셀 줄바꿈은 \\n, 끝 공백 제거)."""
    try:
        raw = c.Text()
    except Exception:
        raw = ""
    return str(raw or "").replace("\r\n", "\n").replace("\r", "\n").rstrip()


def _same_note(a: str, b: str) -> bool:
    return re.sub(r"\s+", " ", a).strip() == re.sub(r"\s+", " ", b).strip()


def _note_has(old: str, memo: str, mode: str) -> bool:
    """이미 그 메모가 달려 있어 바꿀 것이 없는지: 기존 메모가 새 메모와 같거나, --append인데 기존 메모의 끝 줄들이 새 메모와 같음
    (같은 명령을 다시 실행해도 같은 메모가 두 번 붙지 않게)."""
    if _same_note(old, memo):
        return True
    if mode != "append":
        return False
    want = [ln for ln in memo.split("\n") if ln.strip()]
    have = [ln for ln in old.split("\n") if ln.strip()]
    return bool(want) and len(have) >= len(want) and all(_same_note(a, b) for a, b in zip(have[-len(want):], want))


def _sheet_notes(ws: Any, box: Optional[Tuple[int, int, int, int]] = None) -> List[Tuple[int, int, Any]]:
    """시트의 메모 [(행, 열, Comment)] (행, 열 순서). box(r1, c1, r2, c2)를 주면 그 안만.
    칸마다 Range.Comment를 묻지 않고 Worksheet.Comments를 훑음 (큰 범위도 메모 수만큼)."""
    out = []
    coll = ws.Comments
    for i in range(1, int(coll.Count) + 1):
        c = coll.Item(i)
        cell = c.Parent
        r, col = int(cell.Row), int(cell.Column)
        if box is None or _inside((r, col), box):
            out.append((r, col, c))
    return sorted(out, key=lambda x: (x[0], x[1]))


def _range_box(rng: Any) -> Tuple[int, int, int, int]:
    r1, c1 = int(rng.Row), int(rng.Column)
    return r1, c1, r1 + int(rng.Rows.Count) - 1, c1 + int(rng.Columns.Count) - 1


def _memo_rows(path: str) -> List[Tuple[int, str, str]]:
    """메모 파일(.csv/.tsv/.md 표): 칸, 메모. [(파일 줄 번호, 칸, 메모)]. 첫 행이 머리글이면 건너뜀. 값은 글자 그대로 (숫자로 바꾸지 않음)."""
    rows = load_table_file(path, raw=True)
    start = 1
    if rows and str(rows[0][0] if rows[0] else "").strip().lower() in _MEMO_HEADER:
        rows, start = rows[1:], 2
    rt.check_two_columns(path, rows, start, "메모 파일", "칸, 메모", "달지 않았습니다")
    return [(n, str(r[0] if r else "").strip(), str(r[1] if len(r) > 1 and r[1] is not None else "").strip())
            for n, r in enumerate(rows, start) if any(str(v or "").strip() for v in r)]


def _comment_command(args, extra: Sequence[str] = ()) -> Dict[str, Any]:
    """--dry-run을 뺀 같은 comment 명령 (confirm_commands: 사용자 확인 뒤에만)."""
    argv = ["comment", "--target", args.target, "--action", args.action]
    for flag, val in (("--sheet", args.sheet), ("--range", args.range), ("--file", args.file), ("--text", args.text)):
        if val is not None:
            argv += [flag, val]
    for flag, on in (("--overwrite", args.overwrite), ("--append", args.append), ("--no-focus", args.no_focus)):
        if on:
            argv.append(flag)
    argv += list(extra)
    what = {"add": "메모 달기", "delete": "메모 지우기"}[args.action]
    how = {"--overwrite": " (기존 메모를 새 메모로 바꿈)", "--append": " (기존 메모 아래 줄에 덧붙임)"}
    return rt.next_command("excel.py", argv, what + " 실행" + "".join(how.get(x, "") for x in extra),
                           destructive=args.action == "delete" or "--overwrite" in argv)


def _comment_plan(args, t: Any) -> Tuple[List[Tuple[Any, int, int, str]], List[str]]:
    """달 메모 [(시트, 행, 열, 메모)]와 주의 사항. 무엇이든 바꾸기 전에 모든 줄을 확인 (하나라도 틀리면 하나도 달지 않음)."""
    warnings: List[str] = []
    if args.file:
        if args.text is not None:
            raise rt.InputError("--text와 --file은 하나만 지정하세요 (여러 메모는 --file \"메모.csv\": 칸, 메모).")
        if args.range:
            raise rt.InputError("--file의 메모는 파일의 칸 열에 달 칸을 적습니다. --range를 빼고 다시 실행하세요.")
        rows = _memo_rows(rt.resolve_input_path(args.file, "메모 파일", exts=TABLE_FILE_EXTS))
        if not rows:
            raise rt.InputError("메모 파일에 내용이 없습니다. 열: 칸, 메모 (예: B5,규격 3.1과 다름).")
    else:
        memo = rt.unescape_newlines(args.text or "").strip()
        if not args.range or not memo:
            raise rt.InputError("메모를 달 칸과 내용을 지정하세요: --range B5 --text \"메모\" "
                                "(여러 개는 --file \"메모.csv\": 칸, 메모).")
        rows = [(0, args.range, memo)]
    bad: List[Tuple[int, str]] = []
    cells: List[Tuple[str, int, int, str]] = []
    for n, ref, memo in rows:
        where = f"{n}번째 줄" if n else "--range"
        sheet, part = split_sheet(ref)
        try:
            if not memo:
                raise rt.InputError("메모 내용이 없습니다")
            if len(memo) > COMMENT_MAX_CHARS:
                raise rt.InputError(f"메모가 {COMMENT_MAX_CHARS}자(엑셀 상한)보다 깁니다")
            r1, c1, r2, c2 = parse_area(part, None, None) if part else (None, None, None, None)
            if r1 is None or r2 is None or c2 is None:
                raise rt.InputError("칸 주소가 아닙니다 (예: B5, B2:B9)")
            if n and (r1, c1) != (r2, c2):
                raise rt.InputError("칸 하나만 적습니다 (예: B5)")
            if (r2 - r1 + 1) * (c2 - c1 + 1) > COMMENT_MAX_CELLS:
                raise rt.InputError(f"한 번에 {COMMENT_MAX_CELLS}칸까지 메모를 달 수 있습니다")
        except rt.InputError as e:
            bad.append((n, f"{where} '{ref}': {str(e).rstrip('.')}"))
            continue
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cells.append((sheet or args.sheet or "", r, c, memo))
    if bad:
        if not args.file:
            raise rt.InputError(bad[0][1].split(": ", 1)[1] + ". 예: --range B5 --text \"메모\"")
        head = f"메모 파일 {len(rows)}줄 중 {len(bad)}줄이 잘못되어 하나도 달지 않았습니다 (열: 칸, 메모). 고친 뒤 다시 실행하세요. "
        raise rt.InputError(head + " / ".join(b for _, b in bad[:10]),
                            log_message=head + f"잘못된 줄: {', '.join(str(n) for n, _ in bad[:10])} ({rt.DOC_OMITTED})")
    if len(cells) > COMMENT_MAX_CELLS:
        raise rt.InputError(f"한 번에 {COMMENT_MAX_CELLS}개까지 메모를 달 수 있습니다 (받은 칸 {len(cells)}개). 파일을 나눠 실행하세요.")
    # 시트 찾기 (없는 시트는 [추천]/[후보]) + 같은 칸이 여러 줄이면 한 메모로 합침
    plan: Dict[Tuple[str, int, int], List[Any]] = {}
    for name, r, c, memo in cells:
        ws = t.sheet(name or None)
        key = (str(ws.Name), r, c)
        if key in plan:
            plan[key][3] += "\n" + memo
            warnings.append(f"{key[0]}!{addr(r, c)}에 메모가 여러 줄이라 한 메모로 합침")
        else:
            plan[key] = [ws, r, c, memo]
    return [tuple(v) for v in plan.values()], warnings  # type: ignore[misc]


def _comment_list(args) -> Dict[str, Any]:
    with off.open_target("excel", args.target, readonly=True) as t:
        ws = t.sheet(args.sheet)
        rng = target_range(t, ws, args.range) if args.range else None
        if rng is not None:
            ws = rng.Worksheet
        notes = _sheet_notes(ws, _range_box(rng) if rng is not None else None)
        total = len(notes)
        first = max(1, args.start)
        items: List[Dict[str, Any]] = []
        used = 0
        for no, (r, c, note) in enumerate(notes[first - 1:first - 1 + COMMENT_LIST_MAX], first):
            text, cut = rt.clamp_text(_note_text(note), COMMENT_TEXT_MAX)
            item = {"no": no, "cell": addr(r, c), "author": str(getattr(note, "Author", "") or ""), "text": text}
            if cut:
                item["text_truncated"] = True
            size = len(text) + 30
            if items and used + size > rt.DEFAULT_MAX_CHARS:
                break
            items.append(item)
            used += size
        sheet = str(ws.Name)
        resolved = t.resolved()
        where = f"{sheet}!{str(rng.Address).replace('$', '')}" if rng is not None else sheet   # 세션 안에서
    shown_end = first - 1 + len(items)
    truncated = shown_end < total
    result: Dict[str, Any] = {"target_resolved": resolved, "sheet": sheet, "count": total, "shown": len(items),
                              "items": items, "truncated": truncated}
    if not total:
        return rt.ok(TOOL, "comment", f"{where}: 메모 없음 (count 0)",
                     f"VERIFIED: 0건. {where}에는 메모가 없습니다. 사용자에게 그대로 알리세요.", **result)
    if truncated:
        argv = ["comment", "--target", args.target, "--sheet", sheet] + (["--range", args.range] if args.range else [])
        result["next_commands"] = [rt.next_command("excel.py", argv + ["--start", str(shown_end + 1)],
                                                   f"{shown_end + 1}번째 메모부터 이어서 보기")]
    return rt.ok(TOOL, "comment", f"{where}: 메모 {len(items)}개 표시 (전체 {total}개)",
                 "VERIFIED: 읽기 완료." + (" 일부만 표시되었으면 next_commands로 이어서 보세요." if truncated else ""), **result)


def _comment_add(args, t: Any) -> Dict[str, Any]:
    plan, warnings = _comment_plan(args, t)
    for ws in {id(p[0]): p[0] for p in plan}.values():
        check_protection(ws, [], "objects")              # 보호된 시트에서 메모(개체) 편집이 막혀 있으면 달기 전에
    existing = [(ws, r, c, ws.Cells(r, c).Comment) for ws, r, c, _ in plan]
    old = {(str(ws.Name), r, c): _note_text(n) for ws, r, c, n in existing if n is not None}
    mode = "overwrite" if args.overwrite else "append" if args.append else ""
    same = {(str(ws.Name), r, c) for ws, r, c, memo in plan
            if (str(ws.Name), r, c) in old and _note_has(old[(str(ws.Name), r, c)], memo, mode)}
    taken = [(ws, r, c, n) for ws, r, c, n in existing if n is not None and (str(ws.Name), r, c) not in same]
    planned = []
    for ws, r, c, _ in plan:
        ref = f"{ws.Name}!{addr(r, c)}"
        if (str(ws.Name), r, c) in same:
            planned.append(f"그대로(같은 메모가 이미 있음): {ref}")
        elif (str(ws.Name), r, c) not in old:
            planned.append(f"메모 추가: {ref}")
        elif mode:
            planned.append(f"메모 {'바꾸기' if mode == 'overwrite' else '덧붙이기'}: {ref}")
        else:
            planned.append(f"이미 메모가 있음 (바꾸기 --overwrite, 덧붙이기 --append 중 선택): {ref}")
    taken_refs = [f"{ws.Name}!{addr(r, c)}" for ws, r, c, _ in taken]
    taken_text = ", ".join(f"{ref}='{_short(old[(str(ws.Name), r, c)], 30)}'"
                           for ref, (ws, r, c, _) in zip(taken_refs[:5], taken[:5])) + (" ..." if len(taken) > 5 else "")
    if args.dry_run:
        choose = bool(taken and not mode)          # 이미 메모가 있는 칸: 바꿀지(0) 덧붙일지(1) 사용자가 고름
        nexts = [_comment_command(args, ["--overwrite"]), _comment_command(args, ["--append"])] if choose \
            else [_comment_command(args)]
        extra: Dict[str, Any] = {"target_resolved": t.resolved()}
        if taken:
            extra["existing_notes"] = len(taken)
        if warnings:
            extra["warnings"] = warnings[:10]
        directive = ("VERIFIED: 미리보기만 했고 문서는 바꾸지 않았습니다. preview를 사용자에게 알리고, 이미 메모가 있는 칸을 "
                     "바꿀지(confirm_commands[0]) 덧붙일지(confirm_commands[1]) 사용자가 고르면 그 명령을 실행하세요.") if choose else ""
        preview = f"미리보기: 메모 {len(plan)}개" + (f" (이미 메모가 있는 칸 {len(taken)}개: {taken_text})" if taken else "")
        return rt.dry_run_result(TOOL, "comment", preview + " (바꾸지 않음)", len(plan), planned, nexts, directive, **extra)
    if taken and not mode:
        head = f"이미 메모가 있는 칸 {len(taken)}개: "
        raise rt.UserDecisionError(
            head + taken_text,
            "메모를 달지 않았습니다. 이미 메모가 있는 칸과 그 메모를 사용자에게 보여 주고, 사용자가 원하면 새 메모로 바꾸기는 "
            "--overwrite, 기존 메모 아래에 덧붙이기는 --append를 붙여 다시 실행하세요 (미리 보려면 --dry-run).",
            log_message=head + ", ".join(taken_refs[:5]) + (" ..." if len(taken) > 5 else "") + f" (메모 {rt.DOC_OMITTED})")
    expected: Dict[Tuple[str, int, int], str] = {}
    done = at = 0
    try:
        for at, (ws, r, c, memo) in enumerate(plan):
            key = (str(ws.Name), r, c)
            if key in same:
                continue
            cell = ws.Cells(r, c)
            if key in old:
                want = memo if mode == "overwrite" else (old[key] + "\n" + memo if old[key] else memo)
                cell.Comment.Text(Text=want)
            else:
                want = memo
                cell.AddComment(memo)
            expected[key] = want
            done += 1
            t.changed = True
    except Exception as e:
        rt.debug(f"메모 달기 실패: {e}")
        ref = f"{plan[at][0].Name}!{addr(plan[at][1], plan[at][2])}"
        raise rt.InputError(f"{ref}에 메모를 달지 못했습니다 (시트 보호, 편집 중인 칸 등, {type(e).__name__}). "
                            + (f"앞의 {done}개는 달렸습니다(office.py undo로 되돌릴 수 있음). " if done else "")
                            + "시트 보호를 풀거나 칸 편집을 끝낸 뒤 다시 실행하세요.", no_fallback=True)
    # R1: 단 메모를 다시 읽어 확인
    problems: List[str] = []
    for (sheet, r, c), want in expected.items():
        try:
            got = t.sheet(sheet).Cells(r, c).Comment
            if got is None or not _same_note(_note_text(got), want):
                problems.append(f"{sheet}!{addr(r, c)}의 메모가 쓴 내용과 다름")
        except Exception as e:
            problems.append(f"{sheet}!{addr(r, c)}의 메모를 다시 읽지 못함 ({type(e).__name__})")
    t.set_focus(sheet=plan[0][0], changed=bool(expected))     # 모두 그대로면 되돌리기 지점을 남기지 않음
    resolved = t.resolved()
    fields = {"target_resolved": resolved, "applied": [p for p in planned if not p.startswith("그대로")][:50],
              "count": len(plan) - len(same)}
    if same:
        fields["unchanged"] = len(same)
    if warnings:
        fields["warnings"] = warnings[:10]
    preview = f"메모 {len(plan) - len(same)}개를 달았습니다" \
        + (f" (기존 메모 {'바꿈' if mode == 'overwrite' else '덧붙임'} {len(taken)}개)" if taken else "") \
        + (f" (같은 메모가 이미 있는 {len(same)}칸은 그대로)" if same else "")
    if problems:
        return rt.ok_unverified(TOOL, "comment", preview, problems[:10], **fields)
    return rt.ok(TOOL, "comment", preview,
                 "VERIFIED: 메모를 달고 다시 읽어 확인했습니다. 잘못 달았으면 office.py undo로 되돌릴 수 있습니다.", verified=True, **fields)


def _comment_delete(args, t: Any) -> Dict[str, Any]:
    ws = t.sheet(args.sheet)
    rng = target_range(t, ws, args.range)
    ws = rng.Worksheet
    box = _range_box(rng)
    notes = _sheet_notes(ws, box)
    sheet = str(ws.Name)
    where = f"{sheet}!{str(rng.Address).replace('$', '')}"
    planned = [f"메모 삭제: {sheet}!{addr(r, c)}" for r, c, _ in notes]
    if notes:
        check_protection(ws, [], "objects")
    if args.dry_run:
        preview = f"미리보기: {where}의 메모 {len(notes)}개를 지웁니다" if notes else f"미리보기: {where}에는 지울 메모가 없어 바뀔 것 없음"
        return rt.dry_run_result(TOOL, "comment", preview + " (바꾸지 않음)", len(notes), planned, _comment_command(args),
                                 target_resolved=t.resolved())
    if not notes:
        t.set_focus(sheet=ws, changed=False)       # 바뀐 것 없음: 이번 되돌리기 지점은 지움
        return rt.ok(TOOL, "comment", f"{where}: 지울 메모 없음 (count 0)",
                     f"VERIFIED: 0건. {where}에는 메모가 없어 바꾼 것이 없습니다. 사용자에게 그대로 알리세요.",
                     target_resolved=t.resolved(), count=0)
    for _, _, note in notes:
        note.Delete()
    t.set_focus(sheet=ws)
    left = _sheet_notes(ws, box)                       # R1: 다시 읽어 확인
    resolved = t.resolved()
    fields = {"target_resolved": resolved, "applied": planned[:50], "count": len(notes)}
    preview = f"{where}의 메모 {len(notes)}개를 지웠습니다"
    if left:
        return rt.ok_unverified(TOOL, "comment", preview,
                                [f"다시 읽은 범위에 메모가 남음: {', '.join(addr(r, c) for r, c, _ in left[:5])}"], **fields)
    return rt.ok(TOOL, "comment", preview,
                 "VERIFIED: 메모를 지우고 다시 읽어 확인했습니다. 잘못 지웠으면 office.py undo로 되돌릴 수 있습니다.", verified=True, **fields)


def cmd_comment(args) -> Dict[str, Any]:
    if args.overwrite and args.append:
        raise rt.InputError("--overwrite(기존 메모 바꾸기)와 --append(기존 메모에 덧붙이기)는 하나만 지정하세요.")
    # --action add를 빠뜨리고 메모만 준 경우: 목록을 보여 주고 '완료'라고 하면 LLM이 메모를 단 것으로 오해함
    if args.action == "list" and (args.text is not None or args.file):
        args.action = "add"
    if args.action == "list":
        return rt.dry_run_read_only(_comment_list(args), args.dry_run, "comment list")
    if args.action == "delete":
        if not args.range:
            raise rt.InputError("지울 메모의 범위를 --range로 지정하세요 (예: --range B2:D20). 먼저 --dry-run으로 지울 메모를 볼 수 있습니다.")
        if args.text is not None or args.file:
            raise rt.InputError("--action delete에는 --text, --file을 쓰지 않습니다.")
    with off.open_target("excel", args.target, readonly=args.dry_run, focus=not args.no_focus,
                         tool=TOOL, command="comment") as t:
        return _comment_add(args, t) if args.action == "add" else _comment_delete(args, t)


# ==============================================================================
# 9. check(규격서 대조): 내부 모듈 _spec_check.py
# (위의 주소·읽기 도우미를 쓰므로 여기서 불러옴)
# ==============================================================================

import _spec_check  # noqa: E402


def cmd_check(args) -> Dict[str, Any]:
    return _spec_check.cmd_check(args)


# ==============================================================================
# 10. 진입점
# ==============================================================================

def build_cli():
    parser, sub = rt.build_parser(TOOL, "엑셀 작업")

    def common(p, write: bool = False):
        p.add_argument("--target", default="active", help="active(기본) | 열린 통합문서 이름 | 파일 경로")
        p.add_argument("--sheet", help="시트 이름 (생략 시 활성 시트)")
        if write:
            p.add_argument("--no-focus", action="store_true", help="작업 후 화면을 고친 시트로 옮기지 않음")
            p.add_argument("--dry-run", action="store_true", help="바꾸지 않고 바뀔 것만 미리 보기 (미리보기가 없는 동작은 바꾸지 않고 멈춤, 목록은 그대로 읽음)")

    def fmt_args(p):
        p.add_argument("--bg", help="배경색 (yellow, #FFF2CC, 노랑 등)")
        p.add_argument("--color", help="글자색")
        p.add_argument("--bold", action="store_true", default=None, help="굵게")
        p.add_argument("--italic", action="store_true", default=None, help="기울임")
        p.add_argument("--size", type=float, help="글자 크기(pt)")
        p.add_argument("--align", choices=list(XL_ALIGN), help="가로 정렬")
        p.add_argument("--number-format", help="숫자 서식 (예: #,##0  0.0%%  yyyy-mm-dd)")
        p.add_argument("--border", choices=list(XL_WEIGHT) + ["none"], help="테두리 (격자)")
        p.add_argument("--wrap", action="store_true", default=None, help="자동 줄바꿈")

    p = sub.add_parser("read", help="내용 읽기")
    common(p)
    p.add_argument("--range", help="읽을 범위 (생략 시 사용 범위 전체)")
    p.add_argument("--search", help="이 글자가 들어간 셀 찾기")
    p.add_argument("--max-rows", type=int, default=rt.DEFAULT_MAX_ROWS, help=f"표시할 최대 행 수 (기본 {rt.DEFAULT_MAX_ROWS})")
    p.add_argument("--save", nargs="?", const=True, help="CSV로 저장 (경로 생략 시 작업 폴더)")
    p.add_argument("--formulas", action="store_true",
                   help="수식 칸 목록도 함께 (formulas: cell, formula, value). --search는 수식 글에서도 찾음")
    p.add_argument("--where", help="조건에 맞는 행만 (첫 행은 머리글): '금액 >= 100; 부서 = 영업'. 연산자 >= <= > < = != ~(포함)")

    p = sub.add_parser("stats", help="표 집계: 합계·평균·최소·최대·건수 (--by 묶음별, --where 거른 행만)")
    common(p)
    p.add_argument("--range", help="표 범위 (생략 시 사용 범위 전체, 첫 행은 머리글)")
    p.add_argument("--columns", help="집계할 열 (쉼표로, 머리글 이름 또는 열 문자. 생략 시 숫자 열 전부)")
    p.add_argument("--by", help="묶을 열 (쉼표로): 묶음마다 집계 (예: --by 부서)")
    p.add_argument("--where", help="조건에 맞는 행만 (read --where와 같은 식)")
    p.add_argument("--top", type=int, help="--by 묶음을 첫 집계 열의 합계가 큰 순으로 N개만")
    p.add_argument("--save", nargs="?", const=True, help="집계 결과를 CSV로 저장")

    p = sub.add_parser("save", help="열린 통합문서 저장 (--save <경로>면 다른 이름으로 저장)")
    p.add_argument("--target", default="active", help="active(기본) | 열린 통합문서 이름")
    p.add_argument("--save", help="다른 이름으로 저장할 전체 경로 (.xlsx/.xlsm)")
    p.add_argument("--overwrite", action="store_true", help="같은 이름 파일이 있어도 저장 (사용자가 원할 때만)")

    p = sub.add_parser("duplicates", help="중복 행 찾기 (읽기만): --columns 값이 같은 행 묶음과 엑셀 행 번호")
    common(p)
    p.add_argument("--range", help="표 범위 (첫 행 머리글, 생략 시 사용 범위)")
    p.add_argument("--columns", help="비교할 열 (쉼표로, 생략 시 모든 열이 같은 행)")
    p.add_argument("--save", nargs="?", const=True, help="중복 행 목록을 CSV로 저장")

    p = sub.add_parser("info", help="통합문서 요약")
    common(p)
    p.add_argument("--all", action="store_true", help="열린 통합문서 전체 목록")
    p.add_argument("--detail", action="store_true", help="차트 종류/계열 수/위치와 배치 경고(데이터를 가리는 차트)")

    p = sub.add_parser("write", help="값/표/수식 쓰기")
    common(p, write=True)
    p.add_argument("--range", help="시작 셀 또는 범위 (기본 A1). 수식은 범위 전체에 채움")
    p.add_argument("--value", help="값 하나 (범위를 주면 범위 전체에)")
    p.add_argument("--values", help='JSON 2차원 배열 (예: [["이름","값"],["가",1]])')
    p.add_argument("--file", help="표 파일 (.csv, .md 표, .txt)")
    p.add_argument("--formula", nargs="+", help="수식. 1개면 범위 전체, 여러 개면 열마다 (상대참조 자동 조정)")
    p.add_argument("--header", help="첫 행에 쓸 머리글 (쉼표 구분). 데이터/수식은 그 아래부터")
    p.add_argument("--append", action="store_true", help="기존 데이터 아래에 이어 쓰기 (같은 머리글은 생략)")
    p.add_argument("--overwrite", action="store_true", help="값이 있는 칸을 덮어쓰기 (사용자가 확인했을 때만)")
    p.add_argument("--new-sheet", action="store_true",
                   help="--sheet(또는 --range '이름'!A1)의 시트를 새로 만들어 쓰기. 없으면 없는 시트는 오류")
    p.add_argument("--no-autofit", action="store_true", help="열 너비 자동 맞춤 안 함")
    p.add_argument("--cells", help="흩어진 칸 여러 개: \"C6=ft; C7=20; 계산기!C11==A1*2\" 또는 칸값 파일(.csv, 열: 칸, 값). =로 시작하는 값은 수식")

    p = sub.add_parser("replace", help="찾아 바꾸기")
    common(p, write=True)
    p.add_argument("--find", help="찾을 글자")
    p.add_argument("--replace", help="바꿀 글자 (생략 시 서식만 적용)")
    p.add_argument("--map", help='여러 개 한 번에: JSON 객체 또는 JSON 파일 (예: {"초안":"최종"})')
    p.add_argument("--range", help="범위 (생략 시 사용 범위)")
    p.add_argument("--whole", action="store_true", help="셀 전체가 일치할 때만")
    p.add_argument("--all-sheets", action="store_true", help="모든 시트")
    fmt_args(p)

    p = sub.add_parser("style", help="서식")
    common(p, write=True)
    p.add_argument("--range", help="범위 (생략 시 사용 범위)")
    p.add_argument("--theme", help=f"표 테마: {', '.join(off.THEMES)}")
    p.add_argument("--no-zebra", action="store_true", help="테마 적용 시 줄무늬 없음")
    p.add_argument("--highlight", help="조건부 강조: '>=100:yellow', '=완료:lightgreen', '포함:지연:pink'")
    p.add_argument("--clear", action="store_true", help="기존 서식 지우기")
    fmt_args(p)

    p = sub.add_parser("export", help="CSV/XLSX/PDF로 저장")
    common(p)
    p.add_argument("--save", help="저장 경로 (.csv/.xlsx/.pdf, 생략 시 작업 폴더에 CSV)")
    p.add_argument("--range", help="범위 (생략 시 사용 범위)")
    p.add_argument("--columns", help="CSV로 뽑을 열만 (예: A,C,E)")
    p.add_argument("--encoding", default="utf-8-sig", choices=["utf-8-sig", "cp949", "utf-8"], help="CSV 인코딩")

    p = sub.add_parser("sheet", help="시트 관리")
    common(p, write=True)
    p.add_argument("--action", required=True, choices=["list", "add", "rename", "copy", "delete", "select"])
    p.add_argument("--name", help="대상 시트 이름")
    p.add_argument("--new-name", help="새 이름 (rename/copy)")

    p = sub.add_parser("range", help="범위 조작")
    common(p, write=True)
    p.add_argument("--action", required=True,
                   choices=["autofit", "width", "height", "merge", "unmerge", "freeze", "unfreeze",
                            "hide", "unhide", "clear", "copy", "move", "border", "insert", "delete", "sort", "values"],
                   help="values: 수식을 지금 값으로 바꿈 (--dry-run으로 먼저 확인)")
    p.add_argument("--range", help="대상 범위 (freeze는 기준 셀, 기본 A2. insert/delete는 행 5:7 또는 열 B:C)")
    p.add_argument("--dest", help="copy/move 목적지 (예: B2, 'Sheet2'!A1)")
    p.add_argument("--size", type=float, help="width/height 값")
    p.add_argument("--axis", choices=["rows", "cols"],
                   help="hide/unhide/autofit/insert/delete 대상 (생략: '5:7'처럼 행 주소면 rows. insert/delete에 A5:C7을 주면 필요)")
    p.add_argument("--type", choices=["all", "contents", "formats"], default="all", help="clear 종류")
    p.add_argument("--weight", choices=list(XL_WEIGHT) + ["none"], default="thin", help="border 굵기")
    p.add_argument("--box", action="store_true", help="border: 외곽선만")
    p.add_argument("--by", help="sort 기준 열: 머리글 이름 또는 열 문자, 여러 개는 쉼표 (예: \"부서,금액\", 64개까지)")
    p.add_argument("--desc", nargs="?", const="*",
                   help="sort 내림차순: 값 없이 쓰면 모든 기준, 기준 이름을 주면 그것만 (예: --desc \"금액\")")
    p.add_argument("--header", choices=["auto", "yes", "no"], default="auto",
                   help="sort: 첫 행이 머리글인지 (기본 auto: --by가 머리글 이름이면 yes, 열 문자면 도구가 판단)")
    p.add_argument("--overwrite", action="store_true",
                   help="merge: 왼쪽 위 밖의 값을 지우고 병합 / copy·move: 값이 있는 대상 칸을 덮어쓰기 / "
                        "delete: 값이 있는 행·열도 지우기 (사용자가 원할 때만)")

    p = sub.add_parser("chart", help="차트")
    common(p, write=True)
    p.add_argument("--action", default="create", choices=["create", "list", "delete"])
    p.add_argument("--range", help="데이터 범위 (첫 행 머리글, 생략 시 사용 범위)")
    p.add_argument("--type", default="column", choices=list(XL_CHART), help="차트 종류 (기본 column)")
    p.add_argument("--title", help="차트 제목")
    p.add_argument("--name", help="차트 이름 (생략 시 제목). delete 대상 지정에도 사용")
    p.add_argument("--at", help="차트 왼쪽 위 셀 (생략 시 데이터 오른쪽)")
    p.add_argument("--width", type=float, default=480, help="너비(pt)")
    p.add_argument("--height", type=float, default=288, help="높이(pt)")
    p.add_argument("--labels", action="store_true", help="값 표시")
    p.add_argument("--legend", choices=list(XL_LEGEND) + ["none"], help="범례 위치 (계열 1개면 기본 없음)")
    p.add_argument("--theme", help=f"색 테마: {', '.join(off.THEMES)}")
    p.add_argument("--all", action="store_true", help="delete: 시트의 모든 차트")

    p = sub.add_parser("comment", help="칸 메모 (목록 / 달기 / 지우기)")
    common(p, write=True)
    p.add_argument("--action", choices=["list", "add", "delete"], default="list",
                   help="list(기본): 메모 목록, add: 메모 달기 (--text나 --file이 있으면 add), delete: --range 안의 메모 지우기")
    p.add_argument("--range", help="list: 이 범위의 메모만, add: 메모를 달 칸 (예: B5), delete: 메모를 지울 범위")
    p.add_argument("--text", help="add: 메모 내용. 줄바꿈은 \\n")
    p.add_argument("--file", help="add: 여러 메모 파일 (.csv/.md 표, 열: 칸, 메모. 예: B5,규격과 다름)")
    p.add_argument("--overwrite", action="store_true", help="add: 이미 메모가 있는 칸은 새 메모로 바꾸기 (사용자가 확인했을 때만)")
    p.add_argument("--append", action="store_true", help="add: 이미 메모가 있는 칸은 기존 메모 아래 줄에 덧붙이기")
    p.add_argument("--start", type=int, default=1, help="list: 이 번호의 메모부터 보기")

    p = sub.add_parser("check", help="표의 값을 등록된 규격서와 한 번에 대조 (일치/불일치/근거없음/판단필요, 메모 파일)")
    common(p)
    p.add_argument("--range", required=True, help="대조할 행 범위 (예: B2:D50). 바로 위 행은 머리글(값 열의 단위)로 읽음")
    p.add_argument("--item-col", required=True, help="항목 열 (예: B). 이 글로 규격서를 찾음")
    p.add_argument("--value-col", required=True, help="값 열 (예: D)")
    p.add_argument("--section-col", help="장절 열 (예: E). 행마다 그 장절 안에서만 찾음")
    p.add_argument("--unit-col", help="단위 열 (예: E). 생략하면 값 열 바로 오른쪽 열의 머리글이 '단위'일 때 그 열을 씀")
    p.add_argument("--spec", required=True, help="규격서 파일 또는 폴더 (등록 안 된 문서는 먼저 등록)")
    p.add_argument("--section", help="모든 행을 이 장절 안에서만 찾음 (예: 3.2)")
    p.add_argument("--header-row", type=int, help="값 열 단위를 읽을 머리글 행 번호 (기본: --range 바로 위 행, 0 = 머리글 없음)")
    p.add_argument("--save", help="전체 결과를 xlsx로 저장 (예: 대조결과.xlsx)")

    return parser


HANDLERS = {"read": cmd_read, "info": cmd_info, "write": cmd_write, "replace": cmd_replace, "style": cmd_style,
            "export": cmd_export, "sheet": cmd_sheet, "range": cmd_range, "chart": cmd_chart,
            "comment": cmd_comment, "check": cmd_check, "stats": cmd_stats,
            "duplicates": cmd_duplicates, "save": cmd_save}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
