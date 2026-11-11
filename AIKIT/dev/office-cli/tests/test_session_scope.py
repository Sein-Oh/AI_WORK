"""Office 세션(open_target 등 with 블록) 밖에서 COM 객체를 읽지 않는지 코드로 검사합니다.

닫힌 파일은 세션이 끝나면 도구가 문서를 닫고 백그라운드 Office를 끝냅니다. 그 뒤에 세션 안에서 받은 시트·범위 객체의
속성(ws.Name, rng.Address 등)을 읽으면 OLE 0x800a01a8(개체 필요)로 실패합니다. 저장은 이미 끝나 실제로는 바뀌었는데
도구는 실패를 보고합니다 (회사 PC 2026-11-06 O2: excel write --cells, 같은 모양이 chart·comment list에도 있었음).
가짜 COM은 닫힌 뒤의 객체를 흉내 내지 않으므로 코드 모양으로 막습니다.
"""
import ast
import os

import pytest

TOOL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "office-cli")
FILES = ("excel.py", "word.py", "ppt.py", "office.py", "outlook.py", "reader.py", "_spec_check.py")
SESSION_CALLS = ("open_target", "write_session")
# COM 객체에서 읽는 속성 (파이썬 객체의 속성 이름과 겹치지 않는 것)
COM_ATTRS = {"Name", "Value", "Value2", "Formula", "Range", "Cells", "FullName", "Count", "Text", "Address", "Row",
             "Column", "Paragraphs", "Tables", "Slides", "Shapes", "Worksheets", "Sheets", "Content", "Documents",
             "Worksheet", "Parent", "Path", "Hwnd", "Columns", "Rows", "UsedRange", "Comments", "ChartObjects",
             "Presentations", "Workbooks", "Saved"}


def _bound_names(node):
    names = set()
    for n in ast.walk(node):
        if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Store):
            names.add(n.id)
    return names


def _violations(path):
    tree = ast.parse(open(path, encoding="utf-8").read())
    found = []

    def scan(stmts, fn):
        for i, st in enumerate(stmts):
            if isinstance(st, ast.With) and any(any(c in ast.unparse(it.context_expr) for c in SESSION_CALLS)
                                                for it in st.items):
                bound = _bound_names(st)
                for later in stmts[i + 1:]:
                    for n in ast.walk(later):
                        if isinstance(n, ast.Attribute) and n.attr in COM_ATTRS:
                            base = n.value
                            while isinstance(base, (ast.Attribute, ast.Subscript, ast.Call)):
                                base = base.func if isinstance(base, ast.Call) else base.value
                            if isinstance(base, ast.Name) and base.id in bound:
                                found.append(f"{os.path.basename(path)}:{n.lineno} {fn}: {ast.unparse(n)}")
            for field in ("body", "orelse", "finalbody"):
                sub = getattr(st, field, None)
                if isinstance(sub, list) and sub and isinstance(sub[0], ast.stmt):
                    scan(sub, fn)
            for h in getattr(st, "handlers", []) or []:
                scan(h.body, fn)

    for fn in ast.walk(tree):
        if isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
            scan(fn.body, fn.name)
    return found


@pytest.mark.parametrize("name", FILES)
def test_no_com_access_after_session(name):
    assert _violations(os.path.join(TOOL, name)) == []


def test_checker_finds_the_old_pattern(tmp_path):
    """검사기가 회사 PC에서 실패한 모양(세션 뒤 ws.Name)을 실제로 잡는지."""
    bad = tmp_path / "bad.py"
    bad.write_text(
        "def write_cells(args):\n"
        "    with off.open_target('excel', args.target) as t:\n"
        "        todo = [(t.sheet(None), 1)]\n"
        "        for ws, _ in todo:\n"
        "            ws.Range('A1').Value = 1\n"
        "    written = [f'{ws.Name}!A1' for ws, _ in todo]\n"
        "    return written\n", encoding="utf-8")
    assert _violations(str(bad)) == ["bad.py:6 write_cells: ws.Name"]
