"""excel.py compare의 내부 모듈 (진입점 아님, 스킬에 적지 않음): 표 두 판을 행 키로 짝지어 비교 (BOM 개정 비교).

excel.py가 불러 씁니다 (excel.py compare). 주소 도우미는 excel.py의 것을 씁니다.
"""
import json
import os
import re
import zipfile
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import _office as off
import _runtime as rt
import _xlsx_formula as xf
from excel import TOOL, COMMENT_MAX_CELLS, COMMENT_MAX_CHARS, _LETTERS, addr, col_letter, col_number, out_value  # noqa: F401

# ==============================================================================
# 9-2. 두 판 비교 (compare): 표 두 판을 행 키로 짝지어 비교. BOM(부품 목록) 개정 비교에 맞춤
# 머리글 행·열 역할·짝지을 키는 도구가 정함 (사용법 단순성 5번: LLM은 고르지 않음). Excel 없이 파일만 읽음.
# ==============================================================================

COMPARE_EXTS = (".xlsx", ".xlsm", ".csv", ".tsv")
COMPARE_SHOWN = 20             # 구분마다 JSON에 담는 항목 수 (넘으면 truncated, 전체는 결과 xlsx)
COMPARE_HEADER_SCAN = 30       # 머리글 행을 찾을 때 보는 위쪽 행 수 (제목 행 아래의 머리글)
COMPARE_TEXT = 60              # JSON 항목에 담는 값 하나의 글자 수
COMPARE_REF_RANGE = 2000       # 위치 범위 하나(R1-R3)를 펼치는 최대 개수 (R1-R99999 같은 오타가 메모리를 쓰지 않게)
COMPARE_GUESS_PAIRS = 300_000  # 대체 추정에서 이름을 견주는 쌍 수 상한 (넘으면 건너뛰고 warnings)
COMPARE_NAME_SIM = 0.6         # 같은 위치·상위에서 대체로 볼 품명 유사도
COMPARE_FLAT_SIM = 0.85        # 위치·상위 정보가 없는 표에서 대체로 볼 품명 유사도
COMPARE_CANDIDATES = 10

# 열 역할 (역할, 보일 이름, 머리글 같은 뜻 목록). 머리글은 _cmp_header_key로 견줌 (괄호 안 단위, 공백, 기호 무시, 영문 소문자)
COMPARE_ROLES = (
    ("part", "품번", ("품번", "부품번호", "품목번호", "자재번호", "자재코드", "품목코드", "부품코드", "품번코드", "partno", "partnumber",
                     "pn", "itemno", "itemnumber", "itemcode", "materialno", "materialcode", "material", "partcode", "code", "코드")),
    ("parent", "상위품번", ("상위품번", "모품번", "상위품목", "상위품목코드", "상위코드", "모품목", "모품목코드", "상위자재", "parent", "parentpart",
                         "parentpartno", "parentpn", "parentitem", "parentcode", "assembly", "상위")),
    ("level", "레벨", ("레벨", "level", "lvl", "lv", "단계", "bomlevel", "bom레벨", "차수")),
    ("qty", "수량", ("수량", "소요량", "소요수량", "qty", "quantity", "usage", "사용량", "개수", "소요")),
    ("unit", "단위", ("단위", "unit", "uom", "units")),
    ("name", "품명", ("품명", "품목명", "부품명", "자재명", "명칭", "이름", "description", "desc", "name", "partname", "itemname")),
    ("spec", "규격", ("규격", "사양", "spec", "specification", "specs", "value")),
    ("maker", "메이커", ("메이커", "제조사", "제조업체", "제작사", "maker", "manufacturer", "mfr", "mfg", "vendor", "공급사")),
    ("ref", "위치", ("위치", "위치번호", "부품위치", "실장위치", "refdes", "ref", "reference", "referencedesignator",
                    "referencedesignators", "designator", "designators", "location", "loc")),
    ("note", "비고", ("비고", "remark", "remarks", "note", "notes", "memo", "comment", "comments")),
)
_CMP_LABEL = {r: label for r, label, _ in COMPARE_ROLES}
_CMP_CODE_ROLES = ("part", "parent")           # 코드: 대소문자·공백 무시, 앞자리 0 유지 (숫자로 바꾸지 않음)
_CMP_NOCASE_ROLES = ("unit", "maker")          # 대소문자만 무시 (EA = ea)
_CMP_SPLIT_REF = re.compile(r"[,;/\s]+")
_CMP_REF_RANGE = re.compile(r"^([A-Z]+)(\d+)[-~]([A-Z]*)(\d+)$")
_CMP_REF_ONE = re.compile(r"^([A-Z]*)(\d+)$")
_CMP_NUM = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?$")


def _cmp_nfkc(v: Any) -> str:
    import unicodedata
    return unicodedata.normalize("NFKC", str(v))


def _cmp_header_key(v: Any) -> str:
    """머리글 비교값: 전각 -> 반각, 괄호 안 단위 빼기, 영문 소문자, 글자·숫자만 ('Part No.' = 'partno', 'Q'ty' = 'qty', '수량(EA)' = '수량')."""
    if v is None:
        return ""
    s = _cmp_nfkc(v).lower()
    bare = re.sub(r"[\(\[][^\)\]]*[\)\]]", "", s)
    key = re.sub(r"[^0-9a-z가-힣]", "", bare)
    return key or re.sub(r"[^0-9a-z가-힣]", "", s)


_CMP_ROLE_OF = {_cmp_header_key(w): role for role, _, words in COMPARE_ROLES for w in words}


def _cmp_roles(row: Sequence[Any]) -> Dict[str, int]:
    """머리글 행의 열 역할 {역할: 열 번호(0부터)}. 같은 역할이 여러 열이면 왼쪽 열 (나머지는 이름으로 견주는 열)."""
    roles: Dict[str, int] = {}
    for j, v in enumerate(row):
        if v is None or isinstance(v, (int, float)):
            continue
        role = _CMP_ROLE_OF.get(_cmp_header_key(v))
        if role and role not in roles:
            roles[role] = j
    return roles


def _cmp_num_text(x: float) -> str:
    if x == 0:
        return "0"
    if float(x).is_integer() and abs(x) < 1e15:
        return str(int(x))
    return format(x, ".12g")


def _cmp_number(v: Any) -> Optional[float]:
    """수량 비교값: 1 = 1.0 = "1" = "1,000"(1000). 숫자가 아니면 None."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _cmp_nfkc(v).strip()
    if not s or not _CMP_NUM.match(s) or not re.search(r"\d", s):
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _cmp_value(v: Any) -> str:
    """속성 비교값: 앞뒤 공백·여러 칸 공백, 전각·반각을 무시. 숫자는 1 = 1.0 = "1" (앞자리 0이 있는 글 "007"은 글자 그대로)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return _cmp_num_text(float(v))
    s = re.sub(r"\s+", " ", _cmp_nfkc(v)).strip()
    if s and re.search(r"\d", s) and _CMP_NUM.match(s) and not re.match(r"^[+-]?0\d", s):
        n = _cmp_number(s)
        if n is not None:
            return _cmp_num_text(n)
    return s


def _cmp_code(v: Any) -> str:
    """코드 비교값 (품번, 상위품번): 전각·반각, 대소문자, 공백 무시. 앞자리 0은 그대로 (숫자로 저장된 1234.0은 1234)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        v = int(v)
    return re.sub(r"\s+", "", _cmp_nfkc(v)).upper()


def _cmp_ref_key(d: str) -> Tuple[str, int, str]:
    m = re.match(r"^([A-Z]*)(\d+)(.*)$", d)
    return (m.group(1), int(m.group(2)), m.group(3)) if m else (d, -1, "")


def _cmp_refs(v: Any) -> frozenset:
    """위치(Ref Des) 목록을 집합으로: "R1-R3, R5" = "R1,R2,R3,R5" = "r5 r1~r3". 범위는 같은 머리글자일 때만 펼침."""
    if v is None:
        return frozenset()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = _cmp_nfkc(v).upper()
    s = re.sub(r"\s*([-~])\s*", r"\1", s)
    out = set()
    for tok in _CMP_SPLIT_REF.split(s):
        tok = tok.strip()
        if not tok:
            continue
        m = _CMP_REF_RANGE.match(tok)
        if m and (not m.group(3) or m.group(3) == m.group(1)):
            a, b = m.group(2), m.group(4)
            lo, hi = int(a), int(b)
            if lo <= hi and hi - lo < COMPARE_REF_RANGE:
                width = len(a) if a.startswith("0") and len(a) == len(b) else 0
                out.update(f"{m.group(1)}{str(i).zfill(width)}" for i in range(lo, hi + 1))
                continue
        out.add(tok)
    return frozenset(out)


def _cmp_ref_text(refs: Iterable[str]) -> str:
    return ",".join(sorted(refs, key=_cmp_ref_key))


def _cmp_show(v: Any) -> Any:
    """JSON에 담는 원래 값 (긴 글은 자름)."""
    v = out_value(v)
    if isinstance(v, str) and len(v) > COMPARE_TEXT:
        return v[:COMPARE_TEXT] + "..."
    return v


def _cmp_level(v: Any) -> Optional[int]:
    """레벨 칸: 1, "2", ".3", "..3", "**"(별 개수), "1.2.3"(점으로 나눈 단계 수)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = _cmp_nfkc(v).strip()
    if not s:
        return None
    if re.fullmatch(r"\d+(?:\.0+)?", s):
        return int(float(s))
    m = re.fullmatch(r"[.\-*·\s]*(\d+)", s)
    if m:
        return int(m.group(1))
    if re.fullmatch(r"[.*·+]+", s):
        return len(s)
    if re.fullmatch(r"\d+(?:\.\d+)+", s):
        return s.count(".") + 1
    return None


def _cmp_find_header(rows: List[List[Any]]) -> Tuple[int, Dict[str, int]]:
    """머리글 행(0부터)과 열 역할. 위쪽 COMPARE_HEADER_SCAN행 중 역할을 가장 많이 알아본 행 (품번 열이 있으면 더함).
    알아본 열이 없으면 글자 칸이 둘 이상인 첫 행."""
    best_i, best_roles, best_score = -1, {}, 0.0
    for i, row in enumerate(rows[:COMPARE_HEADER_SCAN]):
        roles = _cmp_roles(row)
        score = len(roles) + (0.5 if "part" in roles else 0)
        if score > best_score:
            best_i, best_roles, best_score = i, roles, score
    if best_i >= 0 and best_score >= 1.5:
        return best_i, best_roles
    for i, row in enumerate(rows[:COMPARE_HEADER_SCAN]):
        filled = [v for v in row if v is not None and str(v).strip()]
        if len(filled) >= 2 and all(isinstance(v, str) for v in filled):
            return i, {}
    for i, row in enumerate(rows):
        if any(v is not None and str(v).strip() for v in row):
            return i, {}
    return 0, {}


def _cmp_cell(v: Any) -> Any:
    """읽은 값을 JSON에 담을 수 있는 값으로 (날짜는 글자, 빈 글은 None). 캐시에도 이 값을 씀."""
    if v is None:
        return None
    if isinstance(v, str):
        return v if v.strip() else None
    if isinstance(v, (bool, int, float)):
        return v
    return out_value(v)


def _cmp_cache_path(path: str, sheet: Optional[str]) -> str:
    import hashlib
    st = os.stat(path)
    key = f"{os.path.abspath(path)}|{st.st_mtime_ns}|{st.st_size}|{sheet or ''}|{rt.VERSION}"
    return os.path.join(rt.get_cache_dir(), "compare", hashlib.sha1(key.encode("utf-8")).hexdigest()[:24] + ".json")


def _cmp_pick_sheet(wb: Any, sheet: Optional[str], prefer: Optional[str], option: str) -> Tuple[Any, str]:
    """읽을 시트: 준 시트 -> 이전 판과 같은 이름(새 판) -> 시트 하나 -> 열 역할을 가장 많이 알아본 시트(같으면 활성 시트)."""
    names = [ws.title for ws in wb.worksheets]     # 차트 시트 제외 (2026-11-09 검토 B6)
    if not names:
        raise rt.InputError(f"칸이 있는 워크시트가 없습니다 (차트 시트만 있음, {option}).")
    if sheet:
        found = rt.find_name(sheet, names)
        if found is None:
            charts = [c.title for c in getattr(wb, "chartsheets", [])]
            note = " (차트 시트라 칸이 없음)" if charts and rt.find_name(sheet, charts) else ""
            raise rt.not_found("시트", sheet, names, extra=f" ({option}){note}", option=option)
        return wb[found], f"{option}로 준 시트"
    if prefer and rt.find_name(prefer, names):
        return wb[rt.find_name(prefer, names)], "이전 판과 같은 이름의 시트"
    if len(names) == 1:
        return wb[names[0]], "시트가 하나뿐"
    active = getattr(wb.active, "title", None) if getattr(wb.active, "title", None) in names else None
    best: Tuple[float, int] = (-1.0, 0)
    pick = names[0]
    for name in names:
        top = [list(r) for r in wb[name].iter_rows(min_row=1, max_row=COMPARE_HEADER_SCAN, values_only=True)]
        _, roles = _cmp_find_header(top)
        score = (len(roles) + (0.5 if "part" in roles else 0), 1 if name == active else 0)
        if score > best:
            best, pick = score, name
    return wb[pick], ("머리글로 표를 알아본 시트" if best[0] > 0 else "활성 시트")


def _cmp_read_file(path: str, sheet: Optional[str], prefer: Optional[str], option: str) -> Dict[str, Any]:
    """파일 하나를 읽음: {sheet, sheet_reason, rows(원래 행 전체, 값만), warnings}. xlsx는 openpyxl(값, 읽기 전용), CSV는 글자 그대로."""
    ext = os.path.splitext(path)[1].lower()
    warnings: List[str] = []
    if ext in (".csv", ".tsv"):
        if sheet:
            warnings.append(f"{os.path.basename(path)}는 CSV라 시트가 없어 {option}를 쓰지 않음")
        rows = [[_cmp_cell(v) for v in r] for r in rt.read_csv_rows(path)]
        return {"sheet": os.path.splitext(os.path.basename(path))[0], "sheet_reason": "CSV", "rows": rows, "warnings": warnings}
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "엑셀 파일 두 판 비교")
    wb = rt.retry_fs(lambda: openpyxl.load_workbook(path, read_only=True, data_only=True))
    try:
        ws, reason = _cmp_pick_sheet(wb, sheet, prefer, option)
        ws.reset_dimensions()          # 크기 정보가 틀린 파일(다른 프로그램이 만든 xlsx)도 끝까지 읽음
        rows = [[_cmp_cell(v) for v in r] for r in ws.iter_rows(min_row=1, min_col=1, values_only=True)]
        title = ws.title
    finally:
        wb.close()
    info = xf.xlsx_uncomputed_formulas(path, title)
    return {"sheet": title, "sheet_reason": reason, "rows": rows, "warnings": warnings,
            "uncomputed": info if xf.uncomputed_found(info) else None}


def _cmp_load(path: str, sheet: Optional[str], prefer: Optional[str], option: str) -> Dict[str, Any]:
    """캐시(이전 실행이 시간 제한으로 멈추기 전에 저장한 것)가 있으면 그것을, 없으면 파일을 읽음."""
    cache = _cmp_cache_path(path, sheet if sheet else (f"prefer:{prefer}" if prefer else None))
    try:
        with open(cache, encoding="utf-8") as f:
            data = json.load(f)
        data["cache"] = cache
        data["from_cache"] = True
        return data
    except (OSError, ValueError):
        pass
    data = _cmp_read_file(path, sheet, prefer, option)
    data["cache"] = cache
    return data


def _cmp_save_cache(data: Dict[str, Any]) -> None:
    keep = {k: v for k, v in data.items() if k not in ("cache", "from_cache")}
    os.makedirs(os.path.dirname(data["cache"]), exist_ok=True)
    rt.safe_write_text(data["cache"], json.dumps(keep, ensure_ascii=False))


def _cmp_drop_cache(*tables: Dict[str, Any]) -> None:
    for t in tables:
        if t.get("cache"):
            rt.remove_file(t["cache"])


def _cmp_table(data: Dict[str, Any], path: str) -> Dict[str, Any]:
    """읽은 행에서 머리글을 찾아 표로: header(보일 이름), keys(머리글 비교값), roles, body[(행 번호, 값)], 레벨로 만든 상위·경로."""
    rows = data["rows"]
    hi, roles = _cmp_find_header(rows)
    head = list(rows[hi]) if rows else []
    width = max([len(head)] + [len(r) for r in rows[hi + 1:]]) if rows else 0
    while width and all((r[width - 1] if len(r) >= width else None) is None for r in [head] + rows[hi + 1:]):
        width -= 1
    head = (head + [None] * width)[:width]
    header = [str(v).strip() if v is not None and str(v).strip() else f"{col_letter(j + 1)}열" for j, v in enumerate(head)]
    body = []
    for i, r in enumerate(rows[hi + 1:], start=hi + 2):
        r = (list(r) + [None] * width)[:width]
        if any(v is not None for v in r):
            body.append((i, r))
    t = {"path": path, "name": os.path.basename(path), "sheet": data["sheet"], "sheet_reason": data.get("sheet_reason", ""),
         "header_row": hi + 1, "header": header, "keys": [_cmp_header_key(v) for v in head], "roles": roles, "body": body,
         "csv": os.path.splitext(path)[1].lower() in (".csv", ".tsv")}
    if "level" in roles:
        lv, pc = roles["level"], roles.get("part")
        stack: List[Tuple[int, str]] = []
        parents, paths = [], []
        for _, r in body:
            level = _cmp_level(r[lv])
            code = _cmp_code(r[pc]) if pc is not None else ""
            if level is None:
                parents.append("")
                paths.append("")
                continue
            while stack and stack[-1][0] >= level:
                stack.pop()
            parents.append(stack[-1][1] if stack else "")
            paths.append(">".join(c for _, c in stack))
            stack.append((level, code))
        t["lv_parent"], t["lv_path"] = parents, paths
    return t


def _cmp_pair_columns(old: Dict[str, Any], new: Dict[str, Any]) -> Tuple[List[Tuple[int, int, str, Optional[str]]], List[str], List[str]]:
    """두 판의 열 짝: [(이전 열, 새 열, 보일 이름, 역할)]. 역할이 같으면 머리글이 달라도 짝, 나머지는 머리글 비교값이 같은 열.
    짝이 없는 열은 (이전 판에만, 새 판에만) 머리글 목록."""
    pairs: List[Tuple[int, int, str, Optional[str]]] = []
    used_o, used_n = set(), set()
    for role, _label, _ in COMPARE_ROLES:
        if role in old["roles"] and role in new["roles"]:
            o, n = old["roles"][role], new["roles"][role]
            pairs.append((o, n, old["header"][o], role))
            used_o.add(o)
            used_n.add(n)
    for o, key in enumerate(old["keys"]):
        if o in used_o:
            continue
        n = next((j for j, k in enumerate(new["keys"]) if j not in used_n and k and k == key), None)
        if n is None and not key:
            n = o if o < len(new["keys"]) and o not in used_n and not new["keys"][o] else None   # 머리글이 빈 열은 같은 자리끼리
        if n is not None:
            pairs.append((o, n, old["header"][o], None))
            used_o.add(o)
            used_n.add(n)
    only_old = [old["header"][o] for o in range(len(old["header"])) if o not in used_o]
    only_new = [new["header"][n] for n in range(len(new["header"])) if n not in used_n]
    return pairs, only_old, only_new


# 짝지을 키의 한 부분: ("col", 이전 열, 새 열, 보일 이름, 코드인지) / ("lv_parent",) 레벨로 만든 상위 품번 / ("lv_path",) 레벨로 만든 경로
KeyPart = Tuple[Any, ...]


def _cmp_key_values(t: Dict[str, Any], spec: Sequence[KeyPart], side: int) -> List[str]:
    out = []
    for i, (_, r) in enumerate(t["body"]):
        parts = []
        for p in spec:
            if p[0] == "col":
                v = r[p[side]]
                parts.append(_cmp_code(v) if p[4] else (_cmp_ref_text(_cmp_refs(v)) if p[5] else _cmp_value(v)))
            else:
                parts.append(t[p[0]][i])
        out.append("\x1f".join(parts))
    return out


def _cmp_dups(keys: Sequence[str]) -> int:
    from collections import Counter
    return sum(n for n in Counter(keys).values() if n > 1)


def _cmp_spec_label(spec: Sequence[KeyPart]) -> str:
    names = {"lv_parent": "상위품번(레벨로 찾음)", "lv_path": "상위 경로(레벨로 찾음)"}
    return "+".join(p[3] if p[0] == "col" else names[p[0]] for p in spec)


def _cmp_col_part(old: Dict[str, Any], new: Dict[str, Any], role: str) -> KeyPart:
    return ("col", old["roles"][role], new["roles"][role], old["header"][old["roles"][role]], role in _CMP_CODE_ROLES, role == "ref")


def _cmp_key_arg(args, old: Dict[str, Any], new: Dict[str, Any], pairs: List[Tuple[int, int, str, Optional[str]]]) -> List[KeyPart]:
    """--key "상위품번,품번": 이전 판 머리글 이름(쉼표로 여럿). 열 문자(B)를 주면 그 열의 머리글 이름으로 고친 전체 명령을 [추천]."""
    names = [s.strip() for s in str(args.key).split(",") if s.strip()]
    if not names:
        raise rt.InputError("--key에는 짝지을 열의 머리글 이름을 적습니다 (여럿이면 쉼표, 예: --key \"상위품번,품번\").")
    spec: List[KeyPart] = []
    for name in names:
        o = next((j for j, h in enumerate(old["header"]) if h.strip().lower() == name.lower()), None)
        if o is None:
            o = next((j for j, k in enumerate(old["keys"]) if k and k == _cmp_header_key(name)), None)
        if o is None and _LETTERS.match(name) and col_number(name.upper()) <= len(old["header"]):
            fixed = ",".join(old["header"][col_number(n.upper()) - 1] if n == name else n for n in names)
            raise rt.InputError(f"--key에는 열 문자가 아니라 머리글 이름을 적습니다 ({name}열의 머리글: {old['header'][col_number(name.upper()) - 1]}). "
                                f"[추천] {rt.retry_replacing(args.key, fixed, '--key')}")
        if o is None:
            raise rt.not_found("키 열", name, old["header"], extra=f" (이전 판 {old['name']}의 머리글 {old['header_row']}행)")
        pair = next((p for p in pairs if p[0] == o), None)
        if pair is None:
            raise rt.InputError(f"키 열 '{old['header'][o]}'이(가) 새 판 {new['name']}에 없습니다 (새 판 머리글: "
                                f"{', '.join(new['header'][:15])}). 두 판에 모두 있는 열을 --key에 적으세요.")
        role = pair[3]
        spec.append(("col", pair[0], pair[1], old["header"][o], role in _CMP_CODE_ROLES, role == "ref"))
    return spec


def _cmp_choose_key(args, old: Dict[str, Any], new: Dict[str, Any],
                    pairs: List[Tuple[int, int, str, Optional[str]]]) -> Tuple[List[KeyPart], str, bool]:
    """짝지을 키 (부분 목록, 고른 까닭, 같은 키의 행을 나온 순서대로 짝지을지).
    품번이 행마다 다르면 품번, 같은 품번이 여러 행(여러 단계 BOM)이면 상위품번 열 -> 레벨로 찾은 상위 -> 레벨 경로 -> 위치를 더한 키 중
    두 판 모두에서 행마다 다른 첫 키. 그래도 없으면 [후보](번호와 같은 번호의 confirm_commands: --key를 넣은 전체 명령)."""
    if args.key:
        spec = _cmp_key_arg(args, old, new, pairs)
        return spec, "--key로 준 열", True
    both = [r for r in ("part", "parent", "level", "ref") if r in old["roles"] and r in new["roles"]]
    if "part" in both:
        part = _cmp_col_part(old, new, "part")
        options: List[Tuple[List[KeyPart], str]] = [([part], "품번이 행마다 다름")]
        if "parent" in both:
            options.append(([_cmp_col_part(old, new, "parent"), part], "같은 품번이 여러 행이라 상위품번을 더함"))
        if "level" in both:
            options.append(([("lv_parent",), part], "같은 품번이 여러 행이라 레벨로 찾은 상위품번을 더함"))
            options.append(([("lv_path",), part], "같은 품번이 여러 행이라 레벨로 찾은 상위 경로를 더함"))
        if "ref" in both:
            options.append(([part, _cmp_col_part(old, new, "ref")], "같은 품번이 여러 행이라 위치를 더함"))
            if "parent" in both:
                options.append(([_cmp_col_part(old, new, "parent"), part, _cmp_col_part(old, new, "ref")],
                                "같은 품번이 여러 행이라 상위품번과 위치를 더함"))
        for spec, why in options:
            if not _cmp_dups(_cmp_key_values(old, spec, 1)) and not _cmp_dups(_cmp_key_values(new, spec, 2)):
                return spec, why, False
        # 열로 적을 수 있는 키만 후보 (레벨로 찾은 상위는 --key로 적을 수 없음). 남는 같은 키 행이 적은 순서
        cands = []
        for spec, _ in options:
            if any(p[0] != "col" for p in spec):
                continue
            n = _cmp_dups(_cmp_key_values(old, spec, 1)) + _cmp_dups(_cmp_key_values(new, spec, 2))
            cands.append((n, _cmp_spec_label(spec), ",".join(p[3] for p in spec)))
        head = (f"같은 품번({old['header'][old['roles']['part']]})이 여러 행에 있어 두 판의 행을 짝지을 키를 도구가 정할 수 없습니다 "
                f"(상위품번·레벨·위치를 더해도 같은 키의 행이 남음). 사용자에게 키를 고르게 하세요. 고른 키가 같은 행은 나온 순서대로 짝짓습니다.")
    else:
        cands = []
        for o, n, label, _ in pairs:
            spec = [("col", o, n, label, True, False)]
            ko, kn = _cmp_key_values(old, spec, 1), _cmp_key_values(new, spec, 2)
            if sum(1 for k in ko if k) < max(1, len(ko) // 2):
                continue                                     # 반 넘게 빈 열은 키가 아님
            cands.append((_cmp_dups(ko) + _cmp_dups(kn), label, label))
        if not cands:
            raise rt.InputError(f"두 판에 같은 머리글의 열이 없어 비교할 수 없습니다 (이전 판 머리글 {old['header_row']}행: "
                                f"{', '.join(old['header'][:15])} / 새 판 {new['header_row']}행: {', '.join(new['header'][:15])}).")
        head = ("품번 열(품번, Part No, 자재코드 등)을 알아보지 못해 두 판의 행을 짝지을 키 열을 도구가 정할 수 없습니다. "
                "사용자에게 키 열을 고르게 하세요.")
    cands.sort(key=lambda c: c[0])
    cands = cands[:COMPARE_CANDIDATES]
    shown = [f"{label} (같은 키 {n}행)" if n else label for n, label, _ in cands]
    raise rt.ambiguous(head, shown, values=[v for _, _, v in cands], option="--key",
                       log_message=head.split(" (", 1)[0] + ".")


def _cmp_index(keys: List[str], ordered: bool) -> Tuple[Dict[str, int], int]:
    """키 -> 행 순번. ordered면 같은 키의 n번째 행에 '#n'을 붙여 나온 순서대로 짝지음. (색인, 같은 키 행 수)."""
    idx: Dict[str, int] = {}
    seen: Dict[str, int] = {}
    dup = 0
    for i, k in enumerate(keys):
        if not ordered:
            idx[k] = i
            continue
        n = seen.get(k, 0) + 1
        seen[k] = n
        if n > 1:
            dup += 1
        idx[k if n == 1 else f"{k}\x1e{n}"] = i
    return idx, dup


def _cmp_key_text(t: Dict[str, Any], spec: Sequence[KeyPart], side: int, i: int) -> str:
    """사람이 읽을 키: 'A-100 / B-200' (상위 / 품번)."""
    _, r = t["body"][i]
    parts = []
    for p in spec:
        if p[0] == "col":
            v = out_value(r[p[side]])
            parts.append(str(v) if v != "" else "(빈 칸)")
        else:
            parts.append(t[p[0]][i].replace(">", " > ") or "(최상위)")
    return " / ".join(parts)


def _cmp_ctx(t: Dict[str, Any], i: int) -> Tuple[str, frozenset, str]:
    """대체 추정에 쓰는 행 정보: (상위품번, 위치 집합, 품명 비교값)."""
    _, r = t["body"][i]
    roles = t["roles"]
    parent = _cmp_code(r[roles["parent"]]) if "parent" in roles else (t["lv_parent"][i] if "lv_parent" in t else "")
    refs = _cmp_refs(r[roles["ref"]]) if "ref" in roles else frozenset()
    name_col = roles.get("name", roles.get("spec"))
    name = _cmp_value(r[name_col]).lower() if name_col is not None else ""
    return parent, refs, name


def _cmp_guess(old: Dict[str, Any], new: Dict[str, Any], removed: List[int], added: List[int],
               warnings: List[str]) -> List[Tuple[int, int, str]]:
    """대체 추정: 삭제된 행과 추가된 행 중 위치가 같거나(반 이상 겹침) 상위품번이 같고 품명이 비슷한 짝 (품번 변경 추정).
    위치·상위 정보가 없는 표는 품명이 아주 비슷한 짝만. 한 행은 한 짝에만. 추정이라 합치지 않고 따로 보고. [(이전, 새, 근거)]."""
    import difflib
    if not removed or not added:
        return []
    octx = {i: _cmp_ctx(old, i) for i in removed}
    nctx = {j: _cmp_ctx(new, j) for j in added}
    by_parent: Dict[str, List[int]] = {}
    by_ref: Dict[str, List[int]] = {}
    for j, (parent, refs, _) in nctx.items():
        if parent:
            by_parent.setdefault(parent, []).append(j)
        for d in refs:
            by_ref.setdefault(d, []).append(j)
    flat = not by_parent and not by_ref and not any(p or r for p, r, _ in octx.values())
    if flat and len(removed) * len(added) > COMPARE_GUESS_PAIRS:
        warnings.append(f"삭제 {len(removed)}행 x 추가 {len(added)}행이 많아 대체 추정(품번 변경 추정)을 하지 않음")
        return []
    found: List[Tuple[float, int, int, str]] = []
    pairs = 0
    for i, (parent, refs, name) in octx.items():
        cand = set(added) if flat else set(by_parent.get(parent, []) if parent else [])
        for d in refs:
            cand.update(by_ref.get(d, []))
        for j in cand:
            pairs += 1
            if pairs > COMPARE_GUESS_PAIRS:
                warnings.append(f"대체 추정 비교가 {COMPARE_GUESS_PAIRS}쌍을 넘어 일부만 추정함")
                break
            nparent, nrefs, nname = nctx[j]
            sm = difflib.SequenceMatcher(None, name, nname) if name and nname else None
            sim = sm.ratio() if sm and sm.real_quick_ratio() >= 0.5 and sm.quick_ratio() >= 0.5 else (0.0 if sm else None)
            ref_same = bool(refs and nrefs) and len(refs & nrefs) * 2 >= max(len(refs), len(nrefs))
            parent_same = bool(parent) and parent == nparent
            why = []
            if ref_same and (sim is None or sim >= COMPARE_NAME_SIM - 0.1):
                why.append(f"같은 위치 {_cmp_ref_text(refs & nrefs)[:40]}")
            if parent_same and sim is not None and sim >= COMPARE_NAME_SIM:
                why.append(f"같은 상위품번 {parent}")
            if flat and sim is not None and sim >= COMPARE_FLAT_SIM:
                why.append("위치·상위 정보 없음")
            if not why:
                continue
            if sim is not None:
                why.append(f"품명 유사도 {sim:.2f}")
            found.append(((sim or 0.0) + (1.0 if ref_same else 0.0) + (0.5 if parent_same else 0.0), i, j, ", ".join(why)))
        else:
            continue
        break
    found.sort(key=lambda x: -x[0])
    used_o, used_n = set(), set()
    out = []
    for _, i, j, why in found:
        if i in used_o or j in used_n:
            continue
        used_o.add(i)
        used_n.add(j)
        out.append((i, j, why))
    return out


def _cmp_row_view(t: Dict[str, Any], spec: Sequence[KeyPart], side: int, i: int, which: str) -> Dict[str, Any]:
    """추가·삭제 항목: 키, 행 번호, 알아본 역할 열 값."""
    rown, r = t["body"][i]
    item: Dict[str, Any] = {"key": _cmp_key_text(t, spec, side, i), which: rown}
    for role in ("part", "name", "spec", "qty", "ref"):
        if role in t["roles"]:
            item[role] = _cmp_show(r[t["roles"][role]])
    return item


def _cmp_xlsx(path: str, sheets: List[Tuple[str, List[Any], List[List[Any]], List[int]]]) -> str:
    """결과 xlsx (시트마다 머리글 굵게, 첫 행 고정). 큰 표도 빨리 쓰도록 write_only. 실제 저장 경로."""
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "비교 결과를 xlsx로 저장")
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    bold = Font(bold=True)

    def safe(v: Any) -> Any:
        v = out_value(v) if not isinstance(v, str) else v
        return ILLEGAL_CHARACTERS_RE.sub("", v)[:32000] if isinstance(v, str) else v

    def write(p: str) -> None:
        wb = openpyxl.Workbook(write_only=True)
        for title, header, rows, widths in sheets:
            ws = wb.create_sheet(title)
            for k, w in enumerate(widths, start=1):
                ws.column_dimensions[get_column_letter(k)].width = w
            ws.freeze_panes = "A2"
            cells = []
            for v in header:
                c = WriteOnlyCell(ws, value=safe(v))
                c.font = bold
                cells.append(c)
            ws.append(cells)
            for row in rows:
                out = []
                for v in row:
                    v = safe(v)
                    if isinstance(v, str) and v.startswith("="):
                        c = WriteOnlyCell(ws, value=v)
                        c.data_type = "s"            # 표의 글이 '='로 시작해도 수식이 아닌 글로
                        out.append(c)
                    else:
                        out.append(v)
                ws.append(out)
        wb.save(p)

    return rt.save_with_lock_fallback(path, write, "엑셀 파일")


def _cmp_check_xlsx(path: str, expected: Dict[str, int]) -> Optional[str]:
    """저장한 결과 xlsx를 다시 열어 시트마다 행 수(머리글 포함)를 확인 (R1). 문제가 없으면 None."""
    problem = rt.check_saved_file(path)
    if problem:
        return problem
    import openpyxl
    wb = rt.retry_fs(lambda: openpyxl.load_workbook(path, read_only=True))
    try:
        for title, n in expected.items():
            if title not in wb.sheetnames:
                return f"결과 파일에 '{title}' 시트가 없음"
            got = sum(1 for _ in wb[title].iter_rows(values_only=True))
            if got != n:
                return f"결과 파일 '{title}' 시트의 행 수가 다름 (기대 {n}, 읽은 {got})"
    finally:
        wb.close()
    return None


def cmd_compare(args) -> Dict[str, Any]:
    import time
    old_path = rt.resolve_input_path(args.target, "이전 판 파일", exts=COMPARE_EXTS)
    new_path = rt.resolve_input_path(args.target2, "새 판 파일", exts=COMPARE_EXTS)
    for p, opt in ((old_path, "--target"), (new_path, "--target2")):
        ext = os.path.splitext(p)[1].lower()
        if ext not in COMPARE_EXTS:
            raise rt.InputError(f"compare는 .xlsx, .xlsm, .csv, .tsv 파일을 읽습니다 ({opt}: {os.path.basename(p)}). "
                                ".xls 등은 엑셀에서 .xlsx로 저장한 뒤 그 파일을 주세요.")
        if ext in (".xlsx", ".xlsm") and not zipfile.is_zipfile(p):
            raise rt.InputError(f"{os.path.basename(p)}를 엑셀 파일로 읽을 수 없습니다 (암호·문서 보안(DRM)이 걸렸거나 손상됨). "
                                "엑셀에서 열어 보안이 없는 .xlsx로 저장한 뒤 그 파일을 주세요.")
    if os.path.normcase(os.path.abspath(old_path)) == os.path.normcase(os.path.abspath(new_path)) and \
            (args.sheet or "") == (args.sheet2 or ""):
        raise rt.InputError("--target과 --target2가 같은 파일입니다. 이전 판과 새 판 파일을 따로 주세요.")
    warnings: List[str] = []
    for p in (old_path, new_path):
        try:
            if off.win32com is not None and off.match_open("excel", p) is not None:
                warnings.append(f"{os.path.basename(p)}는 엑셀에 열려 있어 디스크에 저장된 내용으로 비교함 (저장하지 않은 변경은 빠짐)")
        except Exception:
            pass
    t0 = time.monotonic()
    old_data = _cmp_load(old_path, args.sheet, None, "--sheet")
    if not old_data.get("from_cache"):
        # 새 판 읽기가 명령 시간 제한을 넘을 것 같으면 읽은 이전 판을 저장하고 CONTINUE (같은 명령 = 저장한 것을 다시 씀, R10)
        spent = time.monotonic() - t0
        ratio = os.path.getsize(new_path) / max(1, os.path.getsize(old_path))
        try:
            rt.check_time(spent * max(0.2, ratio) * 1.2 + 5, f"이전 판 {os.path.basename(old_path)}을 읽음 (1/2), 새 판 읽기 남음")
        except rt.Unfinished:
            _cmp_save_cache(old_data)
            raise
    new_data = _cmp_load(new_path, args.sheet2, None if args.sheet2 else old_data["sheet"], "--sheet2")
    warnings += old_data.get("warnings", []) + new_data.get("warnings", [])
    old, new = _cmp_table(old_data, old_path), _cmp_table(new_data, new_path)
    if not old["body"] and not new["body"]:
        _cmp_drop_cache(old_data, new_data)
        raise rt.InputError(f"두 파일 모두 머리글 아래에 데이터 행이 없습니다 (이전 판 {old['name']} '{old['sheet']}', 새 판 {new['name']} '{new['sheet']}'). "
                            "비교할 시트를 --sheet, --sheet2로 주세요.")
    pairs, only_old, only_new = _cmp_pair_columns(old, new)
    spec, key_why, ordered = _cmp_choose_key(args, old, new, pairs)
    # 품번이 빈 행(소계, 구분 줄)은 짝짓지 않음
    part_parts = [p for p in spec if p[0] == "col"]
    first = next((p for p in part_parts if p[1] == old["roles"].get("part")), part_parts[0] if part_parts else None)
    skip_o = {i for i, (_, r) in enumerate(old["body"]) if first and r[first[1]] is None}
    skip_n = {i for i, (_, r) in enumerate(new["body"]) if first and r[first[2]] is None}
    if skip_o or skip_n:
        warnings.append(f"키 열({first[3]})이 빈 행은 비교하지 않음 (이전 판 {len(skip_o)}행, 새 판 {len(skip_n)}행)")
    ko, kn = _cmp_key_values(old, spec, 1), _cmp_key_values(new, spec, 2)
    oidx, odup = _cmp_index([k for i, k in enumerate(ko) if i not in skip_o], ordered)
    nidx, ndup = _cmp_index([k for i, k in enumerate(kn) if i not in skip_n], ordered)
    orows = [i for i in range(len(old["body"])) if i not in skip_o]
    nrows = [i for i in range(len(new["body"])) if i not in skip_n]
    oidx = {k: orows[v] for k, v in oidx.items()}
    nidx = {k: nrows[v] for k, v in nidx.items()}
    if odup or ndup:
        warnings.append(f"키({_cmp_spec_label(spec)})가 같은 행은 나온 순서대로 짝지음 (이전 판 {odup}행, 새 판 {ndup}행)")
    roles_pair = {p[3]: p for p in pairs if p[3]}
    qty = roles_pair.get("qty")
    ref = roles_pair.get("ref")
    key_cols = {(p[1], p[2]) for p in spec if p[0] == "col"}
    attrs = [p for p in pairs if (p[0], p[1]) not in key_cols and p[3] not in ("qty", "ref")]
    removed = [oidx[k] for k in oidx if k not in nidx]
    added = [nidx[k] for k in nidx if k not in oidx]
    removed.sort()
    added.sort()
    qty_changed: List[Dict[str, Any]] = []
    ref_changed: List[Dict[str, Any]] = []
    attr_changed: List[Dict[str, Any]] = []
    detail: List[List[Any]] = []          # 결과 xlsx '변경' 시트: 키, 구분, 열, 이전 값, 새 값, 차이, 이전 행, 새 행
    memo: Dict[Tuple[int, int], List[str]] = {}   # 새 판 (행, 열) -> 메모 글

    def note(rown: int, col: int, text: str) -> None:
        memo.setdefault((rown, col), []).append(text)

    changed_rows = 0
    for k, i in oidx.items():
        j = nidx.get(k)
        if j is None:
            continue
        (orow, ro), (nrow, rn) = old["body"][i], new["body"][j]
        hit = False
        key_text = None
        if qty:
            a, b = ro[qty[0]], rn[qty[1]]
            na, nb = _cmp_number(a), _cmp_number(b)
            diff = None
            if na is not None and nb is not None:
                same = abs(na - nb) <= 1e-9 * max(1.0, abs(na), abs(nb))
                diff = _cmp_num_text(nb - na) if not same else None
            else:
                same = _cmp_value(a) == _cmp_value(b)
            if not same:
                key_text = key_text or _cmp_key_text(new, spec, 2, j)
                d = (("+" if not diff.startswith("-") else "") + diff) if diff else ""
                qty_changed.append({"key": key_text, "old": _cmp_show(a), "new": _cmp_show(b), "diff": d,
                                    "row_old": orow, "row_new": nrow})
                detail.append([key_text, "수량 변경", qty[2], out_value(a), out_value(b), d, orow, nrow])
                note(nrow, qty[1], f"비교: 수량 {out_value(a)} -> {out_value(b)}" + (f" ({d})" if d else ""))
                hit = True
        if ref:
            sa, sb = _cmp_refs(ro[ref[0]]), _cmp_refs(rn[ref[1]])
            if sa != sb:
                key_text = key_text or _cmp_key_text(new, spec, 2, j)
                plus, minus = _cmp_ref_text(sb - sa), _cmp_ref_text(sa - sb)
                ref_changed.append({"key": key_text, "added": plus[:200], "removed": minus[:200],
                                    "old": _cmp_show(ro[ref[0]]), "new": _cmp_show(rn[ref[1]]), "row_old": orow, "row_new": nrow})
                what = "; ".join(x for x in ((f"추가 {plus}" if plus else ""), (f"삭제 {minus}" if minus else "")) if x)
                detail.append([key_text, "위치 변경", ref[2], out_value(ro[ref[0]]), out_value(rn[ref[1]]), what, orow, nrow])
                note(nrow, ref[1], f"비교: 위치 {what}")
                hit = True
        cols = []
        for o, n, label, role in attrs:
            a, b = ro[o], rn[n]
            if role in _CMP_CODE_ROLES:
                same = _cmp_code(a) == _cmp_code(b)
            elif role in _CMP_NOCASE_ROLES:
                same = _cmp_value(a).upper() == _cmp_value(b).upper()
            else:
                same = _cmp_value(a) == _cmp_value(b)
            if same:
                continue
            key_text = key_text or _cmp_key_text(new, spec, 2, j)
            cols.append({"column": label, "old": _cmp_show(a), "new": _cmp_show(b)})
            detail.append([key_text, "속성 변경", label, out_value(a), out_value(b), "", orow, nrow])
            note(nrow, n, f"비교: 이전 {label} {out_value(a) if out_value(a) != '' else '(빈 칸)'}")
        if cols:
            attr_changed.append({"key": key_text, "columns": cols[:10], "row_old": orow, "row_new": nrow,
                                 **({"columns_truncated": True} if len(cols) > 10 else {})})
            hit = True
        changed_rows += hit
    guesses: List[Tuple[int, int, str]] = []
    if rt.time_up(reserve=20):
        warnings.append("명령 시간 제한이 가까워 대체 추정(품번 변경 추정)을 하지 않음")
    else:
        guesses = _cmp_guess(old, new, removed, added, warnings)
    gold, gnew = {i for i, _, _ in guesses}, {j for _, j, _ in guesses}
    removed = [i for i in removed if i not in gold]
    added = [j for j in added if j not in gnew]
    part_col_new = new["roles"].get("part", first[2] if first else 0)
    for j in added:
        note(new["body"][j][0], part_col_new, "비교: 이전 판에 없는 행 (추가)")
    guess_view = []
    for i, j, why in guesses:
        o_item = _cmp_row_view(old, spec, 1, i, "row_old")
        n_item = _cmp_row_view(new, spec, 2, j, "row_new")
        guess_view.append({"old_key": o_item["key"], "new_key": n_item["key"], "row_old": o_item["row_old"],
                           "row_new": n_item["row_new"], "old_name": o_item.get("name", ""), "new_name": n_item.get("name", ""),
                           "reason": why, "guess": True})
        note(new["body"][j][0], part_col_new, f"비교(추정, 확인 필요): 이전 판의 {o_item['key']} 대신일 수 있음 ({why})")
    counts = {"추가": len(added), "삭제": len(removed), "수량 변경": len(qty_changed), "위치 변경": len(ref_changed),
              "속성 변경": len(attr_changed), "대체 추정": len(guesses)}
    total = len(added) + len(removed) + changed_rows + len(guesses)
    lists = {"added": [_cmp_row_view(new, spec, 2, j, "row_new") for j in added[:COMPARE_SHOWN]],
             "removed": [_cmp_row_view(old, spec, 1, i, "row_old") for i in removed[:COMPARE_SHOWN]],
             "qty_changed": qty_changed[:COMPARE_SHOWN], "ref_changed": ref_changed[:COMPARE_SHOWN],
             "attr_changed": attr_changed[:COMPARE_SHOWN], "replaced_guess": guess_view[:COMPARE_SHOWN]}
    full_n = {"added": len(added), "removed": len(removed), "qty_changed": len(qty_changed), "ref_changed": len(ref_changed),
              "attr_changed": len(attr_changed), "replaced_guess": len(guesses)}
    truncated = any(full_n[k] > len(v) for k, v in lists.items())
    columns = []
    for o, n, label, role in pairs:
        if role:
            columns.append({"role": _CMP_LABEL[role], "old": f"{col_letter(o + 1)} ({old['header'][o]})", "new": f"{col_letter(n + 1)} ({new['header'][n]})"})
    if only_old:
        warnings.append(f"이전 판에만 있는 열은 비교하지 않음: {', '.join(only_old[:10])}")
    if only_new:
        warnings.append(f"새 판에만 있는 열은 비교하지 않음: {', '.join(only_new[:10])}")
    if not qty:
        warnings.append("수량 열을 알아보지 못해 수량 변경을 따로 세지 않음 (같은 머리글의 열은 속성 변경으로 비교)")
    problems: List[str] = []
    for t, d in ((old, old_data), (new, new_data)):
        if d.get("uncomputed"):
            problems.append(f"{t['name']}: " + xf.uncomputed_note(d["uncomputed"]))
    result: Dict[str, Any] = {
        "old_file": old_path, "new_file": new_path,
        "old_sheet": old["sheet"], "new_sheet": new["sheet"],
        "header_row": {"old": old["header_row"], "new": new["header_row"]},
        "rows": {"old": len(old["body"]), "new": len(new["body"])},
        "key_columns": _cmp_spec_label(spec), "key_reason": key_why, "columns": columns,
        "count": total, "counts": counts, **lists,
        "shown": sum(len(v) for v in lists.values()), "truncated": truncated,
    }
    role_text = ", ".join(f"{c['role']} {c['new'].split(' ', 1)[0]}" for c in columns) or "없음 (머리글 이름이 같은 열끼리 비교)"
    head = f"{old['name']}({len(old['body'])}행) -> {new['name']}({len(new['body'])}행)"
    if args.save and os.path.splitext(str(args.save).strip().strip('"'))[1].lower() not in ("", ".xlsx"):
        warnings.append(f"compare는 .xlsx로만 저장합니다: --save의 {os.path.splitext(str(args.save))[1]} 대신 .xlsx로 저장")
    out = None
    if total or args.save:
        out = rt.resolve_output_path(args.save or None, f"compare_{os.path.splitext(new['name'])[0]}", ".xlsx",
                                     overwrite=bool(args.overwrite))
        summary = [["이전 판", f"{old_path} / 시트 {old['sheet']} / 머리글 {old['header_row']}행 / {len(old['body'])}행"],
                   ["새 판", f"{new_path} / 시트 {new['sheet']} / 머리글 {new['header_row']}행 / {len(new['body'])}행"],
                   ["짝짓기 키", f"{_cmp_spec_label(spec)} ({key_why})"]]
        summary += [[f"열: {c['role']}", f"이전 {c['old']} / 새 {c['new']}"] for c in columns]
        summary += [[k, v] for k, v in counts.items()]
        summary += [["주의", w] for w in warnings]
        ohead = ["이전 행"] + old["header"]
        nhead = ["새 행"] + new["header"]
        add_rows = [[new["body"][j][0]] + list(new["body"][j][1]) for j in added]
        del_rows = [[old["body"][i][0]] + list(old["body"][i][1]) for i in removed]
        guess_rows = [[g["old_key"], g["new_key"], g["row_old"], g["row_new"], g["old_name"], g["new_name"], g["reason"], "추정(확인 필요)"]
                      for g in guess_view]
        for i, j, why in guesses[COMPARE_SHOWN:]:
            o_item, n_item = _cmp_row_view(old, spec, 1, i, "row_old"), _cmp_row_view(new, spec, 2, j, "row_new")
            guess_rows.append([o_item["key"], n_item["key"], o_item["row_old"], n_item["row_new"], o_item.get("name", ""),
                               n_item.get("name", ""), why, "추정(확인 필요)"])
        sheets = [("요약", ["항목", "값"], summary, [16, 90]),
                  ("추가", nhead, add_rows, [8] + [14] * len(new["header"])),
                  ("삭제", ohead, del_rows, [8] + [14] * len(old["header"])),
                  ("변경", ["키", "구분", "열", "이전 값", "새 값", "차이", "이전 행", "새 행"], detail, [24, 10, 14, 24, 24, 20, 8, 8]),
                  ("대체 추정", ["이전 키", "새 키", "이전 행", "새 행", "이전 품명", "새 품명", "근거", "판단"], guess_rows,
                   [22, 22, 8, 8, 24, 24, 40, 14])]
        out = _cmp_xlsx(out, sheets)
        problem = _cmp_check_xlsx(out, {s[0]: len(s[2]) + 1 for s in sheets})
        if problem:
            problems.append(f"결과 파일: {problem}")
        result["output_file"] = out
        if truncated:
            result["full_result_file"] = out
    nexts: List[Dict[str, Any]] = []
    if memo and not new["csv"]:
        sheet = new["sheet"]
        sheet_ref = sheet if re.fullmatch(r"[\w가-힣]+", sheet) else "'" + sheet.replace("'", "''") + "'"
        cells = sorted(memo.items())
        if len(cells) > COMMENT_MAX_CELLS:
            warnings.append(f"바뀐 칸이 {len(cells)}개라 메모 파일에는 앞 {COMMENT_MAX_CELLS}칸만 넣음 (전체는 결과 파일)")
            cells = cells[:COMMENT_MAX_CELLS]
        memo_rows = [["칸", "메모"]] + [[f"{sheet_ref}!{addr(r, c + 1)}", " / ".join(v)[:COMMENT_MAX_CHARS]] for (r, c), v in cells]
        memo_file = rt.write_csv_rows(rt.resolve_output_path(None, f"compare_memo_{sheet}", ".csv"), memo_rows)
        memo_problem = rt.check_saved_file(memo_file, len(memo_rows), delimiter=",")
        if memo_problem:
            problems.append(f"메모 파일: {memo_problem}")
        result["memo_file"] = memo_file
        nexts.append(rt.next_command("excel.py", ["comment", "--target", new_path, "--action", "add", "--file", memo_file, "--dry-run"],
                                     "새 판의 바뀐 칸에 비교 메모 달기 미리보기, 변경 없음 (실제로 다는 명령은 그 결과의 confirm_commands, 사용자 확인 뒤에만)"))
        result["next_commands"] = nexts
    if warnings:
        result["warnings"] = warnings
    _cmp_drop_cache(old_data, new_data)
    say_warn = " warnings도 그대로 알리세요." if warnings else ""
    if not total:
        preview = f"{head}: 바뀐 것 없음 (0건, 짝짓기 키: {_cmp_spec_label(spec)}). 인식한 열: {role_text}"
        if problems:
            return rt.ok_unverified(TOOL, "compare", preview, problems, **result)
        return rt.ok(TOOL, "compare", preview,
                     f"VERIFIED: 0건. 바뀐 것 없음: 두 판을 키({_cmp_spec_label(spec)})로 짝지어 비교했고 추가·삭제·변경이 없습니다. "
                     f"사용자에게 '바뀐 것 없음'과 preview를 그대로 알리고 다시 실행하지 마세요.{say_warn}", verified=True, **result)
    preview = (f"{head}: 추가 {counts['추가']}, 삭제 {counts['삭제']}, 수량 변경 {counts['수량 변경']}, 위치 변경 {counts['위치 변경']}, "
               f"속성 변경 {counts['속성 변경']}, 대체 추정 {counts['대체 추정']} (짝짓기 키: {_cmp_spec_label(spec)}). 인식한 열: {role_text}")
    action = ("preview를 알리고 구분별 목록(added, removed, qty_changed, ref_changed, attr_changed)을 key와 함께 사용자에게 보여 주세요."
              + (" replaced_guess는 품번이 바뀐 것으로 보이는 추정(확인 필요)이라고 알리세요." if guesses else "")
              + (" 목록은 구분마다 앞 20개만이라 전체는 full_result_file이라고 알리세요." if truncated else "")
              + say_warn
              + (" 사용자가 새 판의 바뀐 칸에 메모를 원하면 next_commands[0](미리보기, 변경 없음)을 실행해 보여 주고, 사용자가 '예'라고 한 뒤에만 "
                 "그 결과의 confirm_commands[0]을 실행하세요." if nexts else "")
              + " 같은 compare는 다시 실행하지 마세요.")
    if problems:
        payload = rt.ok_unverified(TOOL, "compare", preview, problems, **result)
        payload["agent_directive"] = "UNVERIFIED: 실행했지만 확인하지 못했습니다: " + "; ".join(problems) + ". " + action
        return payload
    return rt.ok(TOOL, "compare", preview, "VERIFIED: 두 판을 비교하고 결과 파일을 다시 열어 확인했습니다. " + action,
                 verified=True, **result)
