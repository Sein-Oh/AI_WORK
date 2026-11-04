"""excel.py 읽기: read(열린 문서·닫힌 xlsx·CSV), info, --search, --where, --formulas, 출력 상한과 저장 확인."""
import datetime as dt
import json
import os
import re
import types
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
import office
import reader
from fake_excel import FakeExcelWorkbook, FakeRange, FakeXlApp
from fake_office import make_win32com
from office_helpers import error_log_text, formula_xlsx, run_excel, run_main, set_cached_value, TABLE, ws_of


def _strip_dimension(src, dst):
    import re
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = re.sub(rb"<dimension[^>]*/>", b"", data)
            zout.writestr(item, data)


def _grid_xlsx(path, rows, cols):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "데이터"
    for r in range(1, rows + 1):
        ws.append([f"r{r}c{c}" for c in range(1, cols + 1)])
    wb.save(path)


def _patch(path, sheet_old=None, sheet_new=None, calc=None):
    """xlsx 고치기: 시트 XML의 sheet_old -> sheet_new, calc가 있으면 workbook.xml의 fullCalcOnLoad="1"을 calc로."""
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    for n in parts:
        if sheet_old and n.startswith("xl/worksheets/"):
            parts[n] = parts[n].replace(sheet_old.encode(), sheet_new.encode())
    if calc is not None:
        parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(b' fullCalcOnLoad="1"', calc.encode())
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    return str(path)


def _total_xlsx(path):
    """머리글, 3행, 그리고 =SUM 합계 행."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출"
    ws.append(["월", "매출"])
    for m, v in [("1월", 100), ("2월", 200), ("3월", 300)]:
        ws.append([m, v])
    ws.append([None, "=SUM(B2:B4)"])
    wb.save(path)
    return str(path)


def _nodim_xlsx(path):
    """크기 정보(<dimension>)가 없는 파일 (write_only): C열은 머리글 없이 수식만."""
    import openpyxl
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("S")
    ws.append(["a", "b"])
    for i in range(2, 4):
        ws.append([i, i * 2, f"=A{i}+B{i}"])
    ws.append([None, None, "=SUM(C2:C3)"])
    wb.save(path)
    assert b"<dimension" not in zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml")
    return str(path)


def _boom(*a, **k):
    raise MemoryError()


def _patch_sheet(path, subs):
    """시트 XML 고치기 (엑셀에서 저장한 파일 흉내: 계산된 값 넣기, 공유 수식). fullCalcOnLoad 표시도 지움."""
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    for n in parts:
        if n.startswith("xl/worksheets/"):
            text = parts[n].decode("utf-8")
            for old, new in subs:
                assert re.search(old, text), (old, text[:400])
                text = re.sub(old, new, text)
            parts[n] = text.encode("utf-8")
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(b' fullCalcOnLoad="1"', b"")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    return str(path)


def _mixed_xlsx(path):
    """보통 수식(A4), 공유 수식(B1:B3), 배열 수식(C1:C3), '='로 시작하는 글자 상수(D1). 엑셀에서 저장한 것처럼 값이 있음."""
    import openpyxl
    from openpyxl.worksheet.formula import ArrayFormula
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "계산"
    for r, v in enumerate((100, 200, 300), start=1):
        ws.cell(r, 1, v)
        ws.cell(r, 2, f"=A{r}*2")
    ws["A4"] = "=SUM(A1:A3)"
    ws["C1"] = ArrayFormula("C1:C3", "=A1:A3*10")
    ws["C2"], ws["C3"] = 2000, 3000           # 배열의 다른 칸은 파일에 값만 있음
    ws["D1"] = "=글자"
    ws["D1"].data_type = "s"                  # 글자 상수 (수식 아님)
    wb.save(path)
    return _patch_sheet(path, [
        (r'<c r="B1"><f>A1\*2</f><v></v></c>', '<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>200</v></c>'),
        (r'<c r="B2"><f>A2\*2</f><v></v></c>', '<c r="B2"><f t="shared" si="0"/><v>400</v></c>'),
        (r'<c r="B3"><f>A3\*2</f><v></v></c>', '<c r="B3"><f t="shared" si="0"/><v>600</v></c>'),
        (r'(<c r="C1"><f t="array" ref="C1:C3">A1:A3\*10</f>)<v></v>', r'\1<v>1000</v>'),
        (r'(<c r="A4"><f>SUM\(A1:A3\)</f>)<v></v>', r'\1<v>600</v>'),
    ])


def _many_xlsx(path, n=250):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "많음"
    for r in range(1, n + 1):
        ws.cell(r, 1, r)
        ws.cell(r, 2, f"=A{r}+0")
    wb.save(path)
    return _patch_sheet(path, [(r"</f><v></v>", "</f><v>1</v>")])


@pytest.fixture
def xl_formulas(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("Sheet1",))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    ws = wb.Worksheets("Sheet1")
    ws.set(1, 1, 10)
    ws.set(2, 1, 20)
    ws.formulas[(3, 1)], ws.values[(3, 1)] = "=SUM(A1:A2)", 30
    ws.set(1, 2, "'=글자")                       # '로 둔 글자 상수: Formula도 값도 "=글자"
    ws.arrays.append((1, 3, 2, 3))               # C1:C2 배열 수식 (칸마다 같은 수식)
    for r, v in ((1, 20), (2, 40)):
        ws.formulas[(r, 3)], ws.values[(r, 3)] = "=A1:A2*2", v
    return ws


def _table_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "목록"
    for row in TABLE:
        ws.append(list(row))
    ws.append(["마", "개발", None])
    ws["D1"], ws["D2"], ws["D3"] = "날짜", dt.datetime(2026, 10, 1), dt.datetime(2026, 11, 2)
    wb.save(path)
    return str(path)


# ---------------------------------------------------------------------------
# read: 값·검색·잘림, 없는 시트, 닫힌 xlsx
# ---------------------------------------------------------------------------


def test_read_values_search_and_truncation(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    for r in range(1, 31):
        ws.set(r, 1, f"행{r}")
        ws.set(r, 2, float(r))
    code, out = run_excel(capsys, "read", "--range", "A1:B30", "--max-rows", "10")
    assert code == 0 and out["truncated"] and len(out["values"]) == 10
    assert out["values"][1] == ["행2", 2]                    # 2.0 -> 2
    assert "A11:B30" in out["next_commands"][0]["command"]
    code, out = run_excel(capsys, "read", "--search", "행2")
    assert out["total_matches"] == 11 and out["matches"][0]["cell"] == "A2"


def test_read_unknown_sheet_suggests(xl, capsys):
    code, out = run_excel(capsys, "read", "--sheet", "요약표")
    assert code == 1 and "[후보] 요약" in out["message"] and "[추천]" not in out["message"]   # 부분 일치는 추측하지 않음


def test_read_closed_xlsx_with_openpyxl(tmp_path, capsys, monkeypatch):
    import openpyxl
    p = tmp_path / "매출.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "매출"
    wb.active.append(["지역", "금액"])
    wb.active.append(["서울", 1200])
    wb.save(p)
    fake, calls = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    code, out = run_excel(capsys, "read", "--target", str(p))
    assert code == 0 and out["values"] == [["지역", "금액"], ["서울", 1200]]
    assert out["target_resolved"]["source"] == "file(openpyxl)"
    assert not any(c[0] == "DispatchEx" for c in calls)          # Excel을 띄우지 않음
    code, out = run_excel(capsys, "read", "--target", str(p), "--range", "B:B")       # 열 전체 (열린 문서와 같게)
    assert code == 0 and out["values"] == [["금액"], [1200]]
    code, out = run_excel(capsys, "read", "--target", str(p), "--range", "2:2")
    assert code == 0 and out["values"] == [["서울", 1200]]


# ---------------------------------------------------------------------------
# 닫힌 xlsx·CSV 읽기는 열린 문서와 같은 결과 (거꾸로 준 범위 포함)
# ---------------------------------------------------------------------------


def test_closed_xlsx_reads_match_open_document(tmp_path, capsys, monkeypatch):
    import openpyxl
    monkeypatch.setattr(off, "win32com", None)
    src, nodim = tmp_path / "s.xlsx", tmp_path / "nodim.xlsx"
    wb = openpyxl.Workbook()
    for r in range(3, 6):
        for c in range(3, 6):
            wb.active.cell(r, c, r * 10 + c)                   # C3:E5
    wb.save(str(src))
    _strip_dimension(src, nodim)
    for path in (src, nodim):
        code, out = run_main(excel, capsys, "read", "--target", str(path))
        assert code == 0 and out["range"] == "C3:E5" and out["rows"] == 3 and out["values"][0] == [33, 34, 35]
        code, out = run_main(excel, capsys, "read", "--target", str(path), "--range", "E5:C3")       # 거꾸로 준 범위
        assert out["range"] == "C3:E5" and out["values"][-1] == [53, 54, 55]
        code, out = run_main(excel, capsys, "read", "--target", str(path), "--range", "D:D")
        assert out["cols"] == 1 and out["values"][2:] == [[34], [44], [54]]
        code, out = run_main(excel, capsys, "read", "--target", str(path), "--range", "4:4")
        assert out["values"] == [[None, None, 43, 44, 45]] or out["values"] == [["", "", 43, 44, 45]]


def test_csv_reversed_range(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(off, "win32com", None)
    p = tmp_path / "t.csv"
    p.write_text("a,b\n1,2\n3,4\n", encoding="utf-8")
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "B3:A2")
    assert code == 0 and out["values"] == [[1, 2], [3, 4]]


# ---------------------------------------------------------------------------
# 엑셀 범위 밖 칸 주소
# ---------------------------------------------------------------------------


def test_parse_area_rejects_row_zero(tmp_path, capsys):
    with pytest.raises(rt.InputError):
        excel.parse_area("0:3", None, None)
    assert excel.parse_cell("XFD1048576") == (1048576, 16384)
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    p = tmp_path / "x.xlsx"
    wb.save(str(p))
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A0:B2")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")


# ---------------------------------------------------------------------------
# 도구가 쓴 CSV는 ','로 다시 읽음 ('| ' 같은 칸 때문에 구분자를 잘못 추측하지 않게)
# ---------------------------------------------------------------------------


def test_csv_with_pipe_cells_is_verified(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = " | "                       # 구분자 추측이 '|'를 골라 빈 행으로 읽던 칸
    src = tmp_path / "pipe.xlsx"
    wb.save(str(src))
    code, out = run_main(excel, capsys, "read", "--target", str(src), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    p = rt.write_csv_rows(str(tmp_path / "w.csv"), [[" | "]])
    assert rt.check_saved_file(p, 1, delimiter=",") is None


# ---------------------------------------------------------------------------
# 닫힌 xlsx의 명시 범위는 끝쪽 빈 행까지 (열린 문서와 같은 rows)
# ---------------------------------------------------------------------------


def test_closed_xlsx_explicit_range_keeps_trailing_empty_rows(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["이름", "값", "비고"])
    ws.append(["가", 1, ""])
    p = tmp_path / "x.xlsx"
    wb.save(str(p))
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A1:C5")
    assert code == 0 and out["rows"] == 5 and out["range"] == "A1:C5"
    code, out = run_main(excel, capsys, "read", "--target", str(p))
    assert code == 0 and out["rows"] == 2


# ---------------------------------------------------------------------------
# 닫힌 xlsx의 큰 범위는 데이터 끝까지만 채워 읽음 (상한을 넘으면 [추천] 범위)
# ---------------------------------------------------------------------------


def test_huge_explicit_range_is_clipped_to_data(tmp_path, capsys, monkeypatch):
    p = tmp_path / "작은.xlsx"
    _grid_xlsx(p, 3, 2)
    monkeypatch.setattr(excel, "CLOSED_MAX_CELLS", 1000, raising=False)
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A1:Z2000")
    assert code == 0 and out["rows"] == 3 and out["cols"] == 2, out
    assert any("데이터가 있는" in r for r in out.get("recovered", [])), out


def test_whole_sheet_range_does_not_run_out_of_memory(tmp_path, capsys):
    p = tmp_path / "전체.xlsx"
    _grid_xlsx(p, 3, 2)
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A1:XFD1048576")
    assert code == 0 and out["rows"] == 3 and out["range"] == "A1:B3", out


def test_range_too_big_even_after_clip_recommends_narrower(tmp_path, capsys, monkeypatch):
    p = tmp_path / "큰.xlsx"
    _grid_xlsx(p, 6, 5)
    monkeypatch.setattr(excel, "CLOSED_MAX_CELLS", 10, raising=False)
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A1:Z100")
    assert code == 1 and "[추천] --range A1:E2" in out["message"], out


# ---------------------------------------------------------------------------
# 닫힌 xlsx에 서식만 있는 먼 칸(Z300000)이 있어도 큰 범위 읽기는 실제 데이터 끝까지
# ---------------------------------------------------------------------------


def test_stray_formatted_cell_does_not_block_clip(tmp_path, capsys):
    import openpyxl
    from openpyxl.styles import PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "데이터"
    for r in range(1, 11):
        ws.append([f"a{r}", r, r * 2])
    ws["Z300000"].fill = PatternFill("solid", fgColor="FFFF00")          # 값 없이 서식만
    p = tmp_path / "stray.xlsx"
    wb.save(p)
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A1:XFD1048576")
    assert code == 0 and out["range"] == "A1:C10" and out["rows"] == 10, out
    assert any("A1:C10" in r for r in out.get("recovered", [])), out


# ---------------------------------------------------------------------------
# 닫힌 xlsx
# ---------------------------------------------------------------------------


def test_closed_plain_shared_array_and_constant(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = run_excel(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    got = {e["cell"]: e for e in out["formulas"]}
    assert set(got) == {"B1", "B2", "B3", "C1", "A4"}           # D1(글자 상수)과 배열의 C2·C3은 없음
    assert got["A4"] == {"cell": "A4", "formula": "=SUM(A1:A3)", "value": 600}
    assert [got[c]["formula"] for c in ("B1", "B2", "B3")] == ["=A1*2", "=A2*2", "=A3*2"]   # 공유 수식을 칸마다 펼침
    assert [got[c]["value"] for c in ("B1", "B2", "B3")] == [200, 400, 600]
    assert got["C1"] == {"cell": "C1", "formula": "=A1:A3*10", "value": 1000, "array": True, "array_range": "C1:C3"}
    assert out["formula_count"] == 5 and out["formulas_truncated"] is False and "full_result_file" not in out
    assert "수식 칸 5개" in out["agent_directive"] and "수식 5개" in out["preview"]
    assert out["values"][0] == [100, 200, 1000, "=글자"]      # 값은 그대로


def test_closed_without_flag_unchanged(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = run_excel(capsys, "read", "--target", p)
    assert code == 0 and not {"formulas", "formula_count", "formulas_truncated"} & set(out)


def test_closed_uncomputed_value_null_and_unverified(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출"
    ws.append(["월", "매출"])
    for m, v in [("1월", 100), ("2월", 200), ("3월", 300)]:
        ws.append([m, v])
    ws.append([None, "=SUM(B2:B4)"])
    p = str(tmp_path / "t.xlsx")
    wb.save(p)
    code, out = run_excel(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_formulas"] == ["매출!B5"] and out["uncomputed_count"] == 1     # 1.4.31/1.4.32 그대로
    assert out["formulas"] == [{"cell": "B5", "formula": "=SUM(B2:B4)", "value": None}]
    assert out["formula_count"] == 1 and out["rows"] == 5


def test_closed_zero_formulas(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.append(["이름", "값"])
    wb.active.append(["가", 1])
    p = str(tmp_path / "z.xlsx")
    wb.save(p)
    code, out = run_excel(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and out["formula_count"] == 0 and out["formulas"] == [] and out["formulas_truncated"] is False
    assert out["agent_directive"].startswith("VERIFIED:") and "수식 0개" in out["agent_directive"]
    assert "수식이 없습니다" in out["agent_directive"] and "수식 없음" in out["preview"]


def test_cap_truncated_full_result_file(tmp_path, capsys, isolated_dirs):
    p = _many_xlsx(tmp_path / "n.xlsx")
    code, out = run_excel(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["formula_count"] == 250 and len(out["formulas"]) == 200 and out["formulas_truncated"] is True
    assert "full_result_file" in out["agent_directive"]
    full = out["full_result_file"]
    assert full.startswith(str(isolated_dirs["data"])) and os.path.isfile(full)
    rows = [r for r in rt.read_csv_rows(full) if any(r)]
    assert rows[0] == ["cell", "formula", "value", "array"] and len(rows) == 251
    assert rows[-1][:2] == ["B250", "=A250+0"]


def test_cap_with_save_writes_full_list_beside_output(tmp_path, capsys):
    p = _many_xlsx(tmp_path / "n.xlsx")
    save = str(tmp_path / "out" / "값.csv")
    code, out = run_excel(capsys, "read", "--target", p, "--formulas", "--save", save)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["output_file"] == save and out["full_result_file"] == str(tmp_path / "out" / "값_formulas.csv")
    assert len(out["formulas"]) == 200 and out["formulas_truncated"] is True and out["formula_count"] == 250
    rows = [r for r in rt.read_csv_rows(out["full_result_file"]) if any(r)]
    assert len(rows) == 251 and rows[1] == ["B1", "=A1+0", "1", ""]


def test_save_small_list_also_saved(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = run_excel(capsys, "read", "--target", p, "--formulas", "--save", str(tmp_path / "v.csv"))
    assert code == 0 and out["verified"] is True and out["formulas_truncated"] is False
    rows = [r for r in rt.read_csv_rows(out["full_result_file"]) if any(r)]
    assert ["C1", "=A1:A3*10", "1000", "C1:C3"] in rows and len(rows) == 6


def test_search_matches_formula_text(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = run_excel(capsys, "read", "--target", p, "--search", "SUM(")
    assert code == 0 and out["count"] == 0                                 # 값만 찾으면 없음 (그대로)
    code, out = run_excel(capsys, "read", "--target", p, "--search", "SUM(", "--formulas")
    assert code == 0 and out["count"] == 1
    assert out["matches"] == [{"cell": "A4", "value": 600, "formula": "=SUM(A1:A3)"}]
    assert out["formula_count"] == 5


# ---------------------------------------------------------------------------
# CSV, .xls
# ---------------------------------------------------------------------------


def test_csv_has_no_formulas(tmp_path, capsys):
    p = tmp_path / "a.csv"
    p.write_text("이름,값\n가,=1+1\n", encoding="utf-8")
    code, out = run_excel(capsys, "read", "--target", str(p), "--formulas")
    assert code == 0 and out["formula_count"] == 0 and out["formulas"] == []
    assert out["agent_directive"].startswith("VERIFIED:") and "CSV 파일은 수식을 저장하지 않아" in out["agent_directive"]


def test_closed_xls_without_office_is_stop_retry(tmp_path, capsys, monkeypatch, isolated_dirs):
    monkeypatch.setattr(off, "win32com", None)
    p = tmp_path / "old.xls"
    p.write_bytes(b"\xd0\xcf\x11\xe0" + b"\0" * 100)
    code, out = run_excel(capsys, "read", "--target", str(p), "--formulas")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "엑셀에서 이 파일을 연 뒤" in out["message"] and "old.xls" in out["message"]


# ---------------------------------------------------------------------------
# read --where
# ---------------------------------------------------------------------------


def test_where_open_document(xl_table, capsys):
    code, out = run_main(excel, capsys, "read", "--where", "금액 >= 200; 부서 = 영업")
    assert code == 0 and out["count"] == 2, out
    assert out["values"] == [["이름", "부서", "금액"], ["가", "영업", 300], ["다", "영업", 200]], out
    assert out["row_numbers"] == [1, 2, 4], out


@pytest.mark.parametrize("expr, rows", [
    ("금액 > 200", ["가", "라"]), ("금액 < 200", ["나"]), ("금액 <= 200", ["나", "다"]), ("금액 = 1,000", ["라"]),
    ("부서 != 영업", ["나", "라"]), ("부서 ~ 기", ["라"]), ("부서 = '  영업 '", ["가", "다"]), ("B = 개발", ["나"]),
    ("이름 >= 다", ["다", "라"]),
])
def test_where_operators(xl_table, capsys, expr, rows):
    code, out = run_main(excel, capsys, "read", "--where", expr)
    assert code == 0 and [r[0] for r in out["values"][1:]] == rows, out


def test_where_zero_is_r2(xl_table, capsys):
    code, out = run_main(excel, capsys, "read", "--where", "금액 > 5000")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out
    assert out["values"] == [["이름", "부서", "금액"]], out


def test_where_unknown_column_and_bad_expression(xl_table, capsys):
    code, out = run_main(excel, capsys, "read", "--where", "매출 > 1")
    assert code == 1 and "[후보] 이름, 부서, 금액" in out["message"], out
    code, out = run_main(excel, capsys, "read", "--where", "금액 == 100")
    assert code == 1 and out["message"].endswith("[추천] 금액 = 100"), out
    code, out = run_main(excel, capsys, "read", "--where", "금액 >= 100 AND 부서 <> 영업")
    assert code == 1 and out["message"].endswith("[추천] 금액 >= 100; 부서 != 영업"), out
    code, out = run_main(excel, capsys, "read", "--where", "금액 100")
    assert code == 1 and "[추천]" not in out["message"] and "예:" in out["message"], out
    code, out = run_main(excel, capsys, "read", "--where", "금액 > 1", "--search", "가")
    assert code == 1, out


@pytest.mark.parametrize("expr, rows", [
    ("부서 = 개발 | 부서 = 기획", ["나", "라"]),                         # 또는
    ("부서 = 영업; 금액 > 250 | 부서 = 기획", ["가", "라"]),              # ;가 먼저 묶임
    ("금액 > 5000 | 이름 = 나", ["나"]),
    ("부서 = 영업 | 금액 >= 200", ["가", "다", "라"]),                   # 두 묶음에 모두 맞는 행은 한 번
])
def test_where_or(xl_table, capsys, expr, rows):
    """1.4.45: 또는(|)."""
    code, out = run_main(excel, capsys, "read", "--where", expr)
    assert code == 0 and [r[0] for r in out["values"][1:]] == rows and out["count"] == len(rows), out


def test_where_quoted_value_with_separators(xl_table, capsys):
    xl_table.values[(3, 2)] = "개발;품질|지원"
    code, out = run_main(excel, capsys, "read", "--where", "부서 = '개발;품질|지원'")
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["나"], out
    assert out["where"] == "부서 = '개발;품질|지원'", out
    code, out = run_main(excel, capsys, "read", "--where", '부서 ~ "품질|지" | 이름 = 라')
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["나", "라"], out
    xl_table.values[(4, 2)] = "R and D"
    code, out = run_main(excel, capsys, "read", "--where", "부서 = 'R and D'")         # 따옴표 안 AND는 고치지 않음
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["다"], out


def test_where_or_spellings_recommend(xl_table, capsys):
    code, out = run_main(excel, capsys, "read", "--where", "부서 = 개발 OR 부서 = 기획")
    assert code == 1 and out["message"].endswith("[추천] 부서 = 개발 | 부서 = 기획"), out
    code, out = run_main(excel, capsys, "read", "--where", "부서 = 개발 || 금액 > 500")
    assert code == 1 and out["message"].endswith("[추천] 부서 = 개발 | 금액 > 500"), out
    code, out = run_main(excel, capsys, "read", "--where", "| 부서 = 개발")
    assert code == 1 and "[추천]" not in out["message"], out


def test_where_truncation_and_save(xl_table, capsys, tmp_path):
    code, out = run_main(excel, capsys, "read", "--where", "금액 > 0", "--max-rows", "2")
    assert code == 0 and out["truncated"] is True and out["count"] == 4 and len(out["values"]) == 2, out
    assert out["row_numbers"] == [1, 2] and "--where" in out["next_commands"][0]["command"], out
    dest = str(tmp_path / "거른.csv")
    code, out = run_main(excel, capsys, "read", "--where", "부서 = 영업", "--save", dest)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    rows = rt.read_csv_rows(out["output_file"])
    assert rows == [["이름", "부서", "금액"], ["가", "영업", "300"], ["다", "영업", "200"]], rows


def test_where_closed_xlsx_and_csv(tmp_path, capsys):
    p = _table_xlsx(tmp_path / "목록.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p, "--where", "금액 >= 200")
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["가", "다", "라"] and out["row_numbers"] == [1, 2, 4, 5], out
    code, out = run_main(excel, capsys, "read", "--target", p, "--where", "날짜 >= 2026-11-01")
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["나"], out
    code, out = run_main(excel, capsys, "read", "--target", p, "--where", "금액 = ")       # 빈 칸
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["마"], out
    csv_path = tmp_path / "목록.csv"
    csv_path.write_text("이름,금액\n가,10\n나,20\n", encoding="utf-8")
    code, out = run_main(excel, capsys, "read", "--target", str(csv_path), "--where", "금액 > 15")
    assert code == 0 and out["values"] == [["이름", "금액"], ["나", 20]] and out["row_numbers"] == [1, 3], out


# ---------------------------------------------------------------------------
# excel/word/ppt read --save 저장 파일을 다시 열어 확인 (빈 행은 세지 않음)
# ---------------------------------------------------------------------------


def test_excel_read_save_counts_non_blank_rows(tmp_path, capsys, monkeypatch):
    import openpyxl
    p = tmp_path / "표.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A4"] = "머리", 1, 3                  # 3행은 빈 행
    wb.save(str(p))
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 저장한 CSV를 다시 열어"), out
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--save", str(tmp_path / "o2.csv"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:")


# ---------------------------------------------------------------------------
# 출력 크기 상한 (긴 칸, 찾은 값, 슬라이드 글, 워드 단락, 노트)
# ---------------------------------------------------------------------------


def test_excel_read_and_search_cap_long_values(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"], ws["A2"] = "머리", "찾기 " + "다" * 3000, 1
    p = tmp_path / "long.xlsx"
    wb.save(str(p))
    code, out = run_main(excel, capsys, "read", "--target", str(p))
    assert code == 0 and out["truncated"] is True and len(out["values"][0][1]) == rt.MAX_CELL_CHARS
    assert out["next_commands"][-1]["command"].endswith("--save")
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--search", "찾기")
    hit = out["matches"][0]
    assert code == 0 and hit["value_truncated"] and len(hit["value"]) == rt.MAX_CELL_CHARS and out["truncated"] is True
    assert "--save" in out["next_commands"][0]["command"]
    code, out = run_main(excel, capsys, "read", "--target", str(p), "--range", "A1:A2")
    assert code == 0 and out["truncated"] is False and "next_commands" not in out


# ---------------------------------------------------------------------------
# 첫 행이 표시 상한을 넘을 때 '긴 값'과 '뺀 열'을 따로 셈
# ---------------------------------------------------------------------------


def test_preview_separates_long_values_and_dropped_columns(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["가" * 600 for _ in range(40)])
    ws.append([1] * 40)
    p = tmp_path / "넓은.xlsx"
    wb.save(p)
    code, out = run_main(excel, capsys, "read", "--target", str(p))
    assert code == 0 and out["truncated"], out
    assert "긴 값 15칸 잘림" in out["preview"] and "뒤쪽 25열 생략" in out["preview"], out["preview"]
    assert "25열" in out["truncated_reason"]


# ---------------------------------------------------------------------------
# 계산된 값이 없는 수식 칸
# ---------------------------------------------------------------------------


def test_detect_helper(tmp_path):
    p = formula_xlsx(tmp_path / "f.xlsx", n=25)
    info = rt.xlsx_uncomputed_formulas(p)
    assert info["count"] == 25 and info["cells"][0] == "합계!B4" and len(info["cells"]) == rt.UNCOMPUTED_SHOW
    assert rt.xlsx_uncomputed_formulas(p, area=(1, 1, 3, 2))["count"] == 0        # 범위 밖
    assert rt.xlsx_uncomputed_formulas(p, sheet="없음")["count"] == 0
    fields = rt.uncomputed_fields(info)
    assert fields["uncomputed_count"] == 25 and fields["uncomputed_truncated"] is True
    assert rt.xlsx_uncomputed_formulas(formula_xlsx(tmp_path / "c.xlsx", cached=True))["count"] == 0


def _array_xlsx(path, cached_anchor=False):
    """A1:A3 값, C1:C3 배열 수식 (openpyxl은 왼쪽 위 C1만 파일에 씀, fullCalcOnLoad 표시)."""
    import openpyxl
    from openpyxl.worksheet.formula import ArrayFormula
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "배열"
    for r in range(1, 4):
        ws.cell(r, 1, r)
    ws["C1"] = ArrayFormula("C1:C3", "=A1:A3*2")
    wb.save(path)
    if cached_anchor:                              # 왼쪽 위만 값이 있고 표시도 없음: C2, C3은 값이 없음
        set_cached_value(str(path), "<v></v>", "<v>2</v>")
    return str(path)


def test_array_formula_counts_every_cell(tmp_path, capsys):
    """23차 L1 (1.4.43): 배열 수식은 ref 범위의 모든 칸이 수식 칸."""
    info = rt.xlsx_uncomputed_formulas(_array_xlsx(tmp_path / "a.xlsx"))
    assert info["count"] == 3 and info["cells"] == ["배열!C1", "배열!C2", "배열!C3"] and info["bounds"] == [1, 3, 3, 3], info
    p = _array_xlsx(tmp_path / "b.xlsx", cached_anchor=True)
    info = rt.xlsx_uncomputed_formulas(p)
    assert info["count"] == 2 and info["cells"] == ["배열!C2", "배열!C3"], info
    assert rt.xlsx_uncomputed_formulas(p, area=(1, 1, 2, 3))["count"] == 1         # 범위 안의 C2만
    assert rt.xlsx_uncomputed_formulas(p, area=(1, 1, 3, 2))["count"] == 0         # C열 밖
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 2, out


def test_array_formula_with_values_is_computed(tmp_path):
    p = tmp_path / "m.xlsx"
    assert rt.xlsx_uncomputed_formulas(_mixed_xlsx(p))["count"] == 0               # 엑셀 저장: 다른 칸에 값이 있음


def _far_xlsx(path):
    """값 A1:C10과 멀리 떨어진 계산되지 않은 수식 칸 Z200000 하나 (openpyxl 저장, 파일 크기 정보는 A1:Z200000)."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 11):
        for c in range(1, 4):
            ws.cell(r, c, f"h{c}" if r == 1 else r * c)
    ws["Z200000"] = "=A2*2"
    wb.save(path)
    return str(path)


def test_far_uncomputed_cell_is_not_padded(tmp_path, capsys):
    """24차 LOW-4 (1.4.44): 멀리 떨어진 계산되지 않은 수식 칸까지 빈 행을 채우지 않고 따로 읽는 명령을 줌."""
    p = _far_xlsx(tmp_path / "far.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["range"] == "A1:C10" and out["rows"] == 10, {k: out[k] for k in out if k != "values"}
    assert out["agent_directive"].startswith("UNVERIFIED:") and out["far_uncomputed_formulas"] == ["S!Z200000"], out
    cmd = out["next_commands"][-1]["command"]
    assert "--range S!Z200000" in cmd, cmd
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "S!Z200000")    # 그 명령으로 읽으면 그 칸
    assert code == 0 and out["uncomputed_formulas"] == ["S!Z200000"], out


def test_far_uncomputed_cell_whole_columns(tmp_path, capsys):
    p = _far_xlsx(tmp_path / "far.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "A:Z")       # 전에는 '너무 큽니다' + 그 칸을 뺀 [추천]
    assert code == 0 and out["range"] == "A1:Z10" and out["far_uncomputed_count"] == 1, out
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "A:C")       # 그 칸은 요청한 열 밖
    assert code == 0 and out["range"] == "A1:C10" and out["agent_directive"].startswith("VERIFIED:"), out
    assert "far_uncomputed_formulas" not in out
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "A1:Z300000")
    assert code == 0 and out["rows"] == 10 and out["far_uncomputed_count"] == 1, out


def test_far_uncomputed_cell_export(tmp_path, capsys):
    p = _far_xlsx(tmp_path / "far.xlsx")
    code, out = run_main(excel, capsys, "export", "--target", p, "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["rows"] == 10 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["far_uncomputed_formulas"] == ["S!Z200000"]


def test_near_uncomputed_row_still_kept(tmp_path, capsys):
    p = formula_xlsx(tmp_path / "f.xlsx")                                       # 값 바로 아래 합계 행 (가까운 칸)
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["rows"] == 4 and "far_uncomputed_formulas" not in out, out


def test_excel_read_closed_is_unverified(tmp_path, capsys):
    p = formula_xlsx(tmp_path / "f.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_formulas"] == ["합계!B4"] and out["uncomputed_count"] == 1
    assert "계산된 값" in out["agent_directive"] and "엑셀에서 파일을 열어 저장" in out["agent_directive"]
    assert out["values"][3] == ["합계1", ""]
    # 수식 칸이 없는 범위, 계산된 값이 있는 파일은 그대로 VERIFIED
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "A1:B3")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and "uncomputed_formulas" not in out
    code, out = run_main(excel, capsys, "read", "--target", formula_xlsx(tmp_path / "c.xlsx", cached=True))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["values"][3] == ["합계1", 30]


def test_excel_read_search_save_and_export(tmp_path, capsys):
    p = formula_xlsx(tmp_path / "f.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p, "--search", "30")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out   # 0건이라 단정하지 않음
    code, out = run_main(excel, capsys, "read", "--target", p, "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 1
    code, out = run_main(excel, capsys, "export", "--target", p, "--save", str(tmp_path / "e.csv"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 1


# ---------------------------------------------------------------------------
# 끝쪽 수식 행·열을 자르지 않음
# ---------------------------------------------------------------------------


def test_total_row_kept_and_unverified(tmp_path, capsys):
    p = _total_xlsx(tmp_path / "t.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["rows"] == 5 and out["range"] == "A1:B5" and out["values"][4] == ["", ""]
    assert out["uncomputed_formulas"] == ["매출!B5"] and out["uncomputed_count"] == 1


def test_all_formula_sheet(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.append(["이름", "합계"])
    for i in range(2, 6):
        ws.append([f"=Data!A{i}", f"=SUM(Data!B{i}:D{i})"])
    d = wb.create_sheet("Data")
    for i in range(1, 6):
        d.append([f"n{i}", 1, 2, 3])
    p = str(tmp_path / "a.xlsx")
    wb.save(p)
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["rows"] == 5 and out["uncomputed_count"] == 8


def test_search_save_export_transfer(tmp_path, capsys):
    p = _total_xlsx(tmp_path / "t.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p, "--search", "600")
    assert code == 0 and out["count"] == 0 and out["range"] == "A1:B5" and out["agent_directive"].startswith("UNVERIFIED:"), out
    code, out = run_main(excel, capsys, "read", "--target", p, "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["rows"] == 5 and out["agent_directive"].startswith("UNVERIFIED:"), out
    code, out = run_main(excel, capsys, "export", "--target", p, "--save", str(tmp_path / "e.csv"))
    assert code == 0 and out["rows"] == 5 and out["agent_directive"].startswith("UNVERIFIED:"), out
    rows, sheet, info = office._excel_range_rows(types.SimpleNamespace(from_target=p, sheet=None, range=None))
    assert len(rows) == 5 and info["cells"] == ["매출!B5"]


def test_no_dimension_keeps_formula_column(tmp_path, capsys):
    p = _nodim_xlsx(tmp_path / "n.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["range"] == "A1:C4" and out["cols"] == 3 and out["rows"] == 4, out
    assert out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 3
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "C:C")
    assert code == 0 and out["range"] == "C1:C4" and out["uncomputed_count"] == 3, out
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "A1:XFD1048576")    # 큰 범위 줄이기
    assert code == 0 and out["range"] == "A1:C4" and out["uncomputed_count"] == 3, out


# ---------------------------------------------------------------------------
# fullCalcOnLoad
# ---------------------------------------------------------------------------


def test_full_calc_on_load_cached_zero(tmp_path, capsys):
    p = _patch(_total_xlsx(tmp_path / "x.xlsx"), "<v></v>", "<v>0</v>")      # xlsxwriter 모양
    assert b'fullCalcOnLoad="1"' in zipfile.ZipFile(p).read("xl/workbook.xml")
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["values"][4] == ["", 0] and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_formulas"] == ["매출!B5"] and "fullCalcOnLoad" in out["agent_directive"]
    code, out = run_main(excel, capsys, "read", "--target", p, "--range", "B5")          # 빈 칸이 없는 범위도
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 1, out
    code, out = run_main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_formulas"] == ["매출!B5"], out


def test_true_flag_and_excel_saved(tmp_path, capsys):
    t = _patch(_total_xlsx(tmp_path / "t.xlsx"), "<v></v>", "<v>600</v>", ' fullCalcOnLoad="true"')
    assert rt.xlsx_uncomputed_formulas(t)["count"] == 1 and rt.xlsx_full_calc_on_load(t)
    e = _patch(_total_xlsx(tmp_path / "e.xlsx"), "<v></v>", "<v>600</v>", "")         # 엑셀이 저장하면 표시가 지워짐
    assert rt.xlsx_uncomputed_formulas(e)["count"] == 0 and not rt.xlsx_full_calc_on_load(e)
    code, out = run_main(excel, capsys, "read", "--target", e)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["values"][4] == ["", 600], out


# ---------------------------------------------------------------------------
# 확인 실패는 확인 못함
# ---------------------------------------------------------------------------


def test_scan_failure_is_unknown(tmp_path, capsys, monkeypatch):
    import xml.etree.ElementTree as ET
    p = _total_xlsx(tmp_path / "t.xlsx")
    monkeypatch.setattr(ET, "iterparse", _boom)
    info = rt.xlsx_uncomputed_formulas(p)
    assert info["unknown"] is True and info["error"] == "MemoryError" and rt.uncomputed_found(info)
    assert rt.uncomputed_fields(info) == {"uncomputed_check": "unknown"}
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_check"] == "unknown" and "수식 계산 여부를 확인하지 못함" in out["agent_directive"]
    code, out = run_main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_check"] == "unknown", out


def test_area_stops_after_last_row(tmp_path):
    p = _total_xlsx(tmp_path / "t.xlsx")
    # 범위 끝 행(3) 뒤의 XML이 깨져 있어도 범위 안만 확인하고 끝남
    _patch(p, '<row r="4"', '<row r="4"><broken')
    info = rt.xlsx_uncomputed_formulas(p, "매출", (1, 1, 3, 2))
    assert info["count"] == 0 and "unknown" not in info
    assert rt.xlsx_uncomputed_formulas(p, "매출")["unknown"] is True


# ---------------------------------------------------------------------------
# 열린 문서 (가짜 COM)
# ---------------------------------------------------------------------------


def test_com_bulk_formula_hasformula_and_array(xl_formulas, capsys):
    code, out = run_excel(capsys, "read", "--range", "A1:C3", "--formulas")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["formulas"] == [
        {"cell": "C1", "formula": "=A1:A2*2", "value": 20, "array": True, "array_range": "C1:C2"},
        {"cell": "A3", "formula": "=SUM(A1:A2)", "value": 30}]
    assert out["formula_count"] == 2
    reads = [x for x in xl_formulas.log if x[0] == "FormulaRead"]
    assert reads == [("FormulaRead", "$A$1:$C$3")]                         # 범위 전체를 한 번에
    has = [x[1] for x in xl_formulas.log if x[0] == "HasFormula"]
    assert has == ["$A$1:$C$3", "$B$1"]                                     # 값과 같은 글인 칸만 따로 확인


def test_com_no_formulas_skips_formula_read(xl_formulas, capsys):
    code, out = run_excel(capsys, "read", "--range", "A1:A2", "--formulas")
    assert code == 0 and out["formula_count"] == 0 and "수식이 없습니다" in out["agent_directive"]
    assert not [x for x in xl_formulas.log if x[0] == "FormulaRead"]


def test_com_no_arrays_skips_cell_checks(xl_formulas, capsys):
    xl_formulas.arrays.clear()
    code, out = run_excel(capsys, "read", "--range", "A3", "--formulas")
    assert code == 0 and out["formulas"] == [{"cell": "A3", "formula": "=SUM(A1:A2)", "value": 30}]
    assert [x[1] for x in xl_formulas.log if x[0] == "HasArray"] == ["$A$3"]


def test_com_formula2_preferred(xl_formulas, capsys, monkeypatch):
    def formula2(self):
        base = self.Formula
        return tuple(tuple("=SORT(A1:A2)" if f == "=SUM(A1:A2)" else f for f in row) for row in base)
    monkeypatch.setattr(FakeRange, "Formula2", property(formula2), raising=False)
    code, out = run_excel(capsys, "read", "--range", "A1:C3", "--formulas")
    assert code == 0 and {"cell": "A3", "formula": "=SORT(A1:A2)", "value": 30} in out["formulas"]


def test_com_formula2_error_falls_back_to_formula(xl_formulas, capsys, monkeypatch):
    def formula2(self):
        raise Exception("Formula2 없음 (엑셀 2016)")
    monkeypatch.setattr(FakeRange, "Formula2", property(formula2), raising=False)
    code, out = run_excel(capsys, "read", "--range", "A1:C3", "--formulas")
    assert code == 0 and {"cell": "A3", "formula": "=SUM(A1:A2)", "value": 30} in out["formulas"]


# ---------------------------------------------------------------------------
# 오류 로그
# ---------------------------------------------------------------------------


def test_error_log_has_no_formula_text(tmp_path, capsys, monkeypatch, isolated_dirs):
    p = _mixed_xlsx(tmp_path / "m.xlsx")

    def boom(*a):
        raise ValueError("=SECRET_FORMULA(B2)")
    monkeypatch.setattr(excel, "_closed_formulas", boom)
    code, out = run_excel(capsys, "read", "--target", p, "--formulas")
    assert code == 1 and out["agent_directive"].startswith("ERROR:"), out
    assert "SECRET_FORMULA" not in json.dumps(out, ensure_ascii=False)
    log = error_log_text(isolated_dirs)
    assert "수식 읽기 실패 (ValueError)" in log and "SECRET_FORMULA" not in log


def test_error_log_search_formula_text_hidden(tmp_path, capsys, isolated_dirs):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = run_excel(capsys, "read", "--target", p, "--sheet", "없는시트", "--formulas", "--search", "=SUM(A1:A3)")
    assert code == 1
    log = error_log_text(isolated_dirs)
    assert log and "SUM(A1:A3)" not in log and "--formulas" in log


def test_success_writes_no_log(tmp_path, capsys, isolated_dirs):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, _ = run_excel(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and error_log_text(isolated_dirs) == ""


# ---------------------------------------------------------------------------
# 0건 결과
# ---------------------------------------------------------------------------


def test_excel_info_all_zero(xl, capsys):
    code, out = run_main(excel, capsys, "info", "--all")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


# ---------------------------------------------------------------------------
# excel read --search count와 0건
# ---------------------------------------------------------------------------


def test_excel_search_count_and_zero(tmp_path, capsys):
    p = formula_xlsx(tmp_path / "c.xlsx", cached=True)
    code, out = run_main(excel, capsys, "read", "--target", p, "--search", "없는값")
    assert code == 0 and out["count"] == 0 and out["shown"] == 0
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "조건을 바꿔 다시 실행하지 마세요" in out["agent_directive"]
    code, out = run_main(excel, capsys, "read", "--target", p, "--search", "a")
    assert code == 0 and out["count"] == out["total_matches"] == 1 and out["shown"] == 1 and out["truncated"] is False


# ---------------------------------------------------------------------------
# 주소 해석과 표 파일 읽기 도우미
# ---------------------------------------------------------------------------


def test_coerce_and_address_helpers():
    assert excel.coerce("1,200") == 1200 and excel.coerce("3.5") == 3.5
    assert excel.coerce("001") == "001" and excel.coerce("0.5") == 0.5 and excel.coerce("") is None
    assert excel.split_sheet("'My Sheet'!A1:B2") == ("My Sheet", "A1:B2")
    assert excel.addr(2, 1, 101, 2) == "A2:B101" and excel.col_letter(28) == "AB"


def test_load_table_file_markdown_and_csv(tmp_path):
    md = tmp_path / "t.md"
    md.write_text("# 제목\n\n| 공정 | 수율 |\n|---|---|\n| 가열 | 98.2 |\n", encoding="utf-8")
    assert excel.load_table_file(str(md)) == [["공정", "수율"], ["가열", 98.2]]
    csv = tmp_path / "t.csv"
    csv.write_bytes("코드,값\n001,7\n".encode("cp949"))
    assert excel.load_table_file(str(csv)) == [["코드", "값"], ["001", 7]]


# ---------------------------------------------------------------------------
# 데이터 손상 방지
# ---------------------------------------------------------------------------


def test_read_closed_xlsx_without_dimension_reads_all_rows(tmp_path, capsys, monkeypatch):
    import re as _re
    import zipfile
    import openpyxl
    src, dst = tmp_path / "src.xlsx", tmp_path / "nodim.xlsx"
    wb = openpyxl.Workbook()
    for r in [["a", "b", "c"], [1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]]:
        wb.active.append(r)
    wb.save(str(src))
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:     # 다른 프로그램처럼 <dimension> 없이
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = _re.sub(rb"<dimension[^>]*/>", b"", data)
            zout.writestr(item, data)
    monkeypatch.setattr(off, "win32com", None)
    code, out = run_excel(capsys, "read", "--target", str(dst))
    assert code == 0 and out["values"][-1] == [10, 11, 12] and len(out["values"]) == 5


def test_read_csv_target_without_excel(capsys, monkeypatch, tmp_path):
    """CSV도 excel.py read로 읽음 (2026-10-29 Gemma: 확장자 오류 뒤 pandas 코드를 직접 짬). CP949, 숫자 변환, --range."""
    p = tmp_path / "공정.csv"
    p.write_bytes("반응온도(°C),생산효율(%)\r\n200.5,61\r\n210,63\r\n".encode("cp949"))
    monkeypatch.setattr(off, "win32com", None)
    code, out = run_excel(capsys, "read", "--target", str(p))
    assert code == 0 and out["rows"] == 3 and out["values"][1] == [200.5, 61]
    assert out["target_resolved"]["source"] == "file(csv)"
    code, out = run_excel(capsys, "read", "--target", str(p), "--range", "B2:B3")
    assert code == 0 and out["values"] == [[61], [63]] and out["range"] == "B2:B3"
    code, out = run_excel(capsys, "read", "--target", str(p), "--range", "A:A")
    assert code == 0 and out["values"] == [["반응온도(°C)"], [200.5], [210]]


# ---------------------------------------------------------------------------
# excel.py check: 표의 값을 규격서와 한 번에 대조 (B1, 1.4.57)
# ---------------------------------------------------------------------------

SPEC_TEXT = """정격 전압은 5 V ± 0.25 V로 한다.

소비 전류는 최대 3 A로 한다.

절연 저항은 100 MΩ 이상일 것.

내전압은 1.5 kV 1분 이상 견딜 것.

동작 온도는 -10 ~ 50 ℃로 한다.

대기 전력은 0.5 W 이하, 1 W 이하(시험 모드)로 한다.

표시등은 녹색으로 한다.
"""


def _check_book(path, rows, header=("항목", "비고", "측정값 [V]")):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "시험"
    ws["B1"], ws["C1"], ws["D1"] = header
    for i, (item, val) in enumerate(rows, 2):
        ws.cell(i, 2, item)
        ws.cell(i, 4, val)
    wb.save(path)


@pytest.fixture
def check_env(tmp_path, monkeypatch):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    spec = tmp_path / "규격" / "규격서.txt"
    spec.parent.mkdir()
    spec.write_text(SPEC_TEXT, encoding="utf-8")
    book = tmp_path / "시험.xlsx"
    return spec, book


def test_check_verdicts_memo_and_next_command(check_env, capsys):
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.1), ("소비 전류", "3.5 A"), ("절연 저항", "150 MΩ"), ("내전압", "1.5 kV"),
                       ("동작 온도", 60), ("무게", 1.2), ("정격 젼압", 5.0), ("대기 전력", "0.8 W"), ("표시등", "녹색")])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D10", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec.parent))
    assert code == 0 and out["status"] == "ok", out
    got = {r["item"]: (r["verdict"], r["cell"]) for r in out["results"]}
    assert got["정격 전압"] == ("일치", "시험!D2")
    assert got["소비 전류"][0] == "불일치"                       # 칸 단위(A)가 머리글 단위(V)보다 먼저
    assert got["절연 저항"][0] == "일치"
    assert got["내전압"][0] == "일치"                             # 1분(시간)은 단위가 달라 후보에서 빠짐
    assert got["동작 온도"][0] == "판단필요"                       # 머리글 단위 V와 규격 ℃가 다름
    assert got["무게"][0] == "근거없음"
    assert got["정격 젼압"][0] == "판단필요"                       # 한 글자 다른 낱말(fuzzy)
    assert got["대기 전력"][0] == "판단필요"                       # 수치 후보 둘
    assert got["표시등"][0] == "판단필요"                         # 수치 조건 없음
    assert out["counts"] == {"일치": 3, "불일치": 1, "근거없음": 1, "판단필요": 4}
    assert out["agent_directive"].startswith("UNVERIFIED:") and out["not_decided"] == 4
    assert set(out["needs_review"]) == {"시험!D6", "시험!D8", "시험!D9", "시험!D10"}
    r0 = out["results"][0]
    assert "규격서.txt" in r0["evidence"] and "5 V" in r0["spec_text"] and len(r0["spec_text"]) <= 120
    memo = open(out["memo_file"], encoding="utf-8-sig").read().splitlines()
    assert memo[0] == "칸,메모" and memo[1].startswith("시험!D2,일치") and " | 규격서.txt" in memo[1]
    cmd = out["next_commands"][0]["command"]
    assert "excel.py" in cmd and "comment" in cmd and "--dry-run" in cmd and out["memo_file"] in cmd


def test_check_all_decided_is_verified_and_save(check_env, capsys, tmp_path):
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.0), ("소비 전류", "2 A")])
    save = tmp_path / "대조결과.xlsx"
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D3", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec), "--save", str(save))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["counts"]["일치"] == 2 and out["output_file"] == str(save)
    import openpyxl
    rows = list(openpyxl.load_workbook(save, read_only=True).active.iter_rows(values_only=True))
    assert rows[0][:4] == ("칸", "항목", "값", "판정") and len(rows) == 3


def test_check_section_column_limits_search(check_env, capsys, tmp_path):
    spec, book = check_env
    spec.write_text("1. 전원\n정격 전압은 5 V로 한다.\n\n2. 시험 조건\n정격 전압은 12 V로 한다.\n", encoding="utf-8")
    _check_book(book, [("정격 전압", 12)])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D2", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec), "--section", "시험 조건")
    assert code == 0 and out["results"][0]["verdict"] == "일치", out


def test_check_empty_items_is_zero(check_env, capsys):
    spec, book = check_env
    _check_book(book, [])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D5", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec))
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_check_bad_column_is_input_error(check_env, capsys):
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.0)])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D2", "--item-col", "항목",
                          "--value-col", "D", "--spec", str(spec))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
