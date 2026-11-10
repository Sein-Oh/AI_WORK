"""office-cli 내부 모듈 (진입점 아님): 닫힌 xlsx에서 계산된 값이 저장되지 않은 수식 칸 찾기.

openpyxl 등이 쓰고 엑셀에서 한 번도 열지 않은 파일은 수식 칸의 값이 비어 있어, 읽은 결과를 VERIFIED로 줄 수 없습니다 (R1).
excel.py, reader.py, office.py가 씁니다.
"""
import io
import re
from typing import Any, Dict, List, Optional, Tuple

import _runtime as rt

UNCOMPUTED_POSITIONS = 100_000   # positions=True일 때 담는 칸 위치 수 상한 (넘으면 멀리 떨어진 칸을 가려내지 않음)
UNCOMPUTED_SHOW = 20       # 계산된 값이 없는 수식 칸을 결과에 담는 개수 상한
_XLSX_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XLSX_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_XLSX_F_TAG = re.compile(rb"<(?:\w+:)?f[\s>/]")
_CELL_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


_XLSX_CELL_OPEN = re.compile(rb"<(?:\w+:)?c\s[^>]*?\br=\"([A-Za-z]{1,3})(\d+)\"[^>]*?(/?)>")


def _sheet_part(z: Any, title: str) -> Optional[str]:
    """통합문서 zip에서 시트 이름 -> 시트 XML 경로 (없으면 None)."""
    import posixpath
    import xml.etree.ElementTree as ET
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") or "" for r in rels}
    for sh in wb.iter(_XLSX_MAIN + "sheet"):
        if (sh.get("name") or "") == title:
            target = targets.get(sh.get(_XLSX_REL + "id"), "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            return part if part in set(z.namelist()) else None
    return None


def xlsx_value_extent(path: str, sheet: str) -> Optional[Tuple[int, int]]:
    """시트 XML에서 내용이 있는 칸(<c ...>...</c>, 서식만 있는 <c .../>는 뺌)의 끝 (행, 열). 파일에 적힌 크기(<dimension>)가
    실제보다 작은 파일(Excel이 아닌 프로그램이 만든 xlsx)도 끝까지 읽기 위해 씁니다. 칸이 없으면 (0, 0), 읽지 못하면 None.
    XML을 조각으로 읽어 정규식으로 훑습니다 (openpyxl로 한 번 더 읽는 것보다 빠름)."""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            part = _sheet_part(z, sheet)
            if part is None:
                return None
            r2 = c2 = 0
            tail = b""
            with z.open(part) as f:
                while True:
                    chunk = f.read(4 << 20)
                    buf = tail + chunk
                    cut = len(buf) if not chunk else max(buf.rfind(b"<"), 0)
                    for m in _XLSX_CELL_OPEN.finditer(buf, 0, cut):
                        if m.group(3):
                            continue
                        r = int(m.group(2))
                        col = 0
                        for ch in m.group(1).upper():
                            col = col * 26 + ch - 64
                        r2, c2 = max(r2, r), max(c2, col)
                    tail = buf[cut:]
                    if not chunk:
                        break
            return r2, c2
    except Exception as e:                      # noqa: BLE001  확인 못 하면 파일에 적힌 크기를 씀
        rt.debug(f"시트 끝 확인 실패: {type(e).__name__}")
        return None


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
        rt.debug(f"수식 값 확인 실패: {type(e).__name__}")
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
    """xlsx_uncomputed_formulas 결과를 UNVERIFIED 문구로 (칸 주소는 앞 3개)."""
    if info.get("unknown"):
        return (f"수식 계산 여부를 확인하지 못함 (파일의 시트 XML을 읽지 못함: {info.get('error') or '알 수 없음'}). "
                "수식 칸이 있으면 빈 칸이나 계산되지 않은 값으로 읽혔을 수 있으니 사용자에게 그대로 알리세요")
    # 칸 목록은 결과의 uncomputed_formulas에 한 번만: 문구에는 앞 3개 (unverified·agent_directive에 같은 목록이 세 번 들어가던 것,
    # 2026-11-09 검토 5장 4번)
    cells = ", ".join(info["cells"][:3]) + (" 등, 목록은 uncomputed_formulas" if info["count"] > 3 else "")
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
    return rt.ok_unverified(payload["tool"], payload["command"], payload["preview"],
                         [uncomputed_note(info)] + list(payload.get("unverified") or []), **fields)
