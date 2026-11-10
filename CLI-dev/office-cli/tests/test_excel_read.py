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
import _spec_check
import _xlsx_formula as xf
import excel
import office
import reader
from fake_excel import FakeExcelWorkbook, FakeRange, FakeXlApp
from fake_office import make_win32com
from office_helpers import option_value, recommended_argv, split_command, error_log_text, formula_xlsx, run_excel, run_main, set_cached_value, TABLE, ws_of


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
    assert code == 1 and "[후보] 1) 요약" in out["message"] and "[추천]" not in out["message"]   # 부분 일치는 추측하지 않음


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
    assert code == 1, out
    argv = recommended_argv(out["message"], "excel")             # 1.4.79: [추천]은 전체 명령
    assert option_value(argv, "--range") == "A1:E2" and "A1:Z100" not in argv and option_value(argv, "--target") == str(p), argv
    code, out = run_main(excel, capsys, *argv)
    assert code == 0 and out["range"] == "A1:E2", out


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


@pytest.mark.parametrize("expr, fixed", [
    ("금액 >= 200, 부서 = 영업", "금액 >= 200; 부서 = 영업"),       # Gemma 실패 유형 감사: 쉼표로 나눈 조건이 값 하나로 읽혀 0건이었음
    ("금액 >= 1,000, 부서 = 영업", "금액 >= 1,000; 부서 = 영업"),
])
def test_where_comma_between_conditions_gets_fixed_expression(xl_table, capsys, expr, fixed):
    code, out = run_main(excel, capsys, "read", "--where", expr)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    _check_where_recommend(capsys, out, expr, fixed)


def _check_where_recommend(capsys, out, expr, fixed):
    """[추천] 전체 명령의 --where가 고친 식이고 옛 식은 없으며, 그대로 다시 실행하면 성공 (1.4.79: [추천]은 전체 명령)."""
    argv = recommended_argv(out["message"], "excel")
    assert argv[0] == "read" and option_value(argv, "--where") == fixed and expr not in argv, argv
    code, again = run_main(excel, capsys, *argv)
    assert code == 0 and again["where"] == fixed, again


@pytest.mark.parametrize("expr, rows", [("금액 = 1,000", ["라"]), ("부서 = '영업, 기획'", [])])
def test_where_comma_inside_value_is_kept(xl_table, capsys, expr, rows):
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
    assert code == 1, out
    _check_where_recommend(capsys, out, "금액 == 100", "금액 = 100")
    code, out = run_main(excel, capsys, "read", "--where", "금액 >= 100 AND 부서 <> 영업")
    assert code == 1, out
    _check_where_recommend(capsys, out, "금액 >= 100 AND 부서 <> 영업", "금액 >= 100; 부서 != 영업")
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
    assert code == 1, out
    _check_where_recommend(capsys, out, "부서 = 개발 OR 부서 = 기획", "부서 = 개발 | 부서 = 기획")    # 1.4.79: 전체 명령
    code, out = run_main(excel, capsys, "read", "--where", "부서 = 개발 || 금액 > 500")
    assert code == 1, out
    _check_where_recommend(capsys, out, "부서 = 개발 || 금액 > 500", "부서 = 개발 | 금액 > 500")
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
    m = re.search(r"긴 값 (\d+)칸 잘림.*뒤쪽 (\d+)열 생략", out["preview"])       # 상한(1.4.97: 6000자) 안에 든 칸은 잘리고, 나머지 열은 생략
    assert m and int(m.group(1)) + int(m.group(2)) == 40 and int(m.group(1)) >= 5, out["preview"]
    assert f"{m.group(2)}열" in out["truncated_reason"]


# ---------------------------------------------------------------------------
# 계산된 값이 없는 수식 칸
# ---------------------------------------------------------------------------


def test_detect_helper(tmp_path):
    p = formula_xlsx(tmp_path / "f.xlsx", n=25)
    info = xf.xlsx_uncomputed_formulas(p)
    assert info["count"] == 25 and info["cells"][0] == "합계!B4" and len(info["cells"]) == xf.UNCOMPUTED_SHOW
    assert xf.xlsx_uncomputed_formulas(p, area=(1, 1, 3, 2))["count"] == 0        # 범위 밖
    assert xf.xlsx_uncomputed_formulas(p, sheet="없음")["count"] == 0
    fields = xf.uncomputed_fields(info)
    assert fields["uncomputed_count"] == 25 and fields["uncomputed_truncated"] is True
    assert xf.xlsx_uncomputed_formulas(formula_xlsx(tmp_path / "c.xlsx", cached=True))["count"] == 0


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
    info = xf.xlsx_uncomputed_formulas(_array_xlsx(tmp_path / "a.xlsx"))
    assert info["count"] == 3 and info["cells"] == ["배열!C1", "배열!C2", "배열!C3"] and info["bounds"] == [1, 3, 3, 3], info
    p = _array_xlsx(tmp_path / "b.xlsx", cached_anchor=True)
    info = xf.xlsx_uncomputed_formulas(p)
    assert info["count"] == 2 and info["cells"] == ["배열!C2", "배열!C3"], info
    assert xf.xlsx_uncomputed_formulas(p, area=(1, 1, 2, 3))["count"] == 1         # 범위 안의 C2만
    assert xf.xlsx_uncomputed_formulas(p, area=(1, 1, 3, 2))["count"] == 0         # C열 밖
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 2, out


def test_array_formula_with_values_is_computed(tmp_path):
    p = tmp_path / "m.xlsx"
    assert xf.xlsx_uncomputed_formulas(_mixed_xlsx(p))["count"] == 0               # 엑셀 저장: 다른 칸에 값이 있음


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
    assert "--range S!Z200000" in cmd, cmd     # 낱말 가운데의 !는 따옴표 없이
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
    assert xf.xlsx_uncomputed_formulas(t)["count"] == 1 and xf.xlsx_full_calc_on_load(t)
    e = _patch(_total_xlsx(tmp_path / "e.xlsx"), "<v></v>", "<v>600</v>", "")         # 엑셀이 저장하면 표시가 지워짐
    assert xf.xlsx_uncomputed_formulas(e)["count"] == 0 and not xf.xlsx_full_calc_on_load(e)
    code, out = run_main(excel, capsys, "read", "--target", e)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["values"][4] == ["", 600], out


# ---------------------------------------------------------------------------
# 확인 실패는 확인 못함
# ---------------------------------------------------------------------------


def test_scan_failure_is_unknown(tmp_path, capsys, monkeypatch):
    import xml.etree.ElementTree as ET
    p = _total_xlsx(tmp_path / "t.xlsx")
    monkeypatch.setattr(ET, "iterparse", _boom)
    info = xf.xlsx_uncomputed_formulas(p)
    assert info["unknown"] is True and info["error"] == "MemoryError" and xf.uncomputed_found(info)
    assert xf.uncomputed_fields(info) == {"uncomputed_check": "unknown"}
    code, out = run_main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_check"] == "unknown" and "수식 계산 여부를 확인하지 못함" in out["agent_directive"]
    code, out = run_main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_check"] == "unknown", out


def test_area_stops_after_last_row(tmp_path):
    p = _total_xlsx(tmp_path / "t.xlsx")
    # 범위 끝 행(3) 뒤의 XML이 깨져 있어도 범위 안만 확인하고 끝남
    _patch(p, '<row r="4"', '<row r="4"><broken')
    info = xf.xlsx_uncomputed_formulas(p, "매출", (1, 1, 3, 2))
    assert info["count"] == 0 and "unknown" not in info
    assert xf.xlsx_uncomputed_formulas(p, "매출")["unknown"] is True


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
    # 4차 검토(1.4.67): 메모에 '자동 판정(참고)'과 읽은 조건·근거 문장 앞부분을 함께 적어 사람이 확인
    assert memo[0] == "칸,메모" and memo[1].startswith("시험!D2,자동 판정(참고): 일치 - 읽은 조건: 4.75~5.25 V - 근거: 규격서.txt"), memo[1]
    assert "정격 전압은 5 V" in memo[1]
    assert any("자동 판정(참고): 근거없음 (규격서에서 항목을 찾지 못함) - 읽은 조건: 없음 - 근거: 없음" in m for m in memo), memo
    cmd = out["next_commands"][0]["command"]
    assert "excel.py" in cmd and "comment" in cmd and "--dry-run" in cmd and out["memo_file"] in cmd
    # 5차 검토(1.4.68): 안내에 메모 미리보기 행동(채점 O1: 미리보기 -> 사용자 확인 뒤 미리보기 결과의 confirm_commands, 1.4.73)
    d = out["agent_directive"]
    assert "next_commands[0](미리보기, 변경 없음)을 실행해 보여 주고" in d and "'예'라고 한 뒤에만" in d and "confirm_commands[0]" in d, d
    assert "다시 실행하지 말고 사용자에게 그대로 알리세요" not in d and len(out["next_commands"]) == 1


def test_check_remaining_specs_reruns_same_command(check_env, capsys, monkeypatch):
    # 5차 검토(1.4.68): 규격서 등록이 남으면 같은 명령 다시 대조만 (메모 미리보기·'다시 실행하지 말고'와 부딪히지 않게)
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.0)])
    real = reader.load_spec_docs

    def partial(path, max_files=10):
        docs, info = real(path)
        info["remaining"] = 2
        return docs, info
    monkeypatch.setattr(reader, "load_spec_docs", partial)
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D2", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec.parent))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    cmds = [c["command"] for c in out["next_commands"]]
    assert len(cmds) == 1 and " check " in cmds[0] and "--dry-run" not in cmds[0], cmds
    d = out["agent_directive"]
    assert "next_commands[0](같은 명령)을 그대로 다시 실행" in d and "다시 실행하지 말고" not in d, d


def test_check_time_limit_marks_rest_needs_review(check_env, capsys, monkeypatch):
    # 6차 검토 (1.4.69): 명령 시간 제한이 다가오면 남은 행은 대조하지 않고 판단필요 + UNVERIFIED (JSON은 늘 냄)
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.1), ("소비 전류", "3.5 A"), ("절연 저항", "150 MΩ")])
    monkeypatch.setattr(rt, "time_up", lambda reserve=0: True)
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D4", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec.parent))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    got = [(r["cell"], r["verdict"]) for r in out["results"]]
    assert got[0] == ("시험!D2", "일치") and got[1:] == [("시험!D3", "판단필요"), ("시험!D4", "판단필요")], got
    # 7차 검토 (1.4.70): 대조하지 못한 행은 reason '대조하지 못함', match_mode null, 메모에서 빼고 남은 범위의 전체 check 명령
    r1 = out["results"][1]
    assert r1["reason"].startswith("대조하지 못함") and r1["match_mode"] is None, r1
    memo = open(out["memo_file"], encoding="utf-8-sig").read().splitlines()
    assert len(memo) == 2 and memo[1].startswith("시험!D2,"), memo
    cmds = [c["command"] for c in out["next_commands"]]
    assert len(cmds) == 2 and "comment" in cmds[0] and " check " in cmds[1], cmds
    assert "--range B3:D4" in cmds[1] and "--header-row 1" in cmds[1] and "--item-col B" in cmds[1], cmds[1]
    d = out["agent_directive"]
    assert "명령 시간 제한으로 2행(시험!D3부터)은 대조하지 못함" in d and "확인을 받은 뒤 next_commands[1]을 그대로 실행" in d, d
    # 남은 범위 명령: 머리글(값 열 단위)은 원래 범위 위 행에서 읽음
    monkeypatch.setattr(rt, "time_up", lambda reserve=0: False)
    code, out = run_excel(capsys, *split_command(cmds[1])[2:])
    assert code == 0 and [r["cell"] for r in out["results"]] == ["시험!D3", "시험!D4"], out
    assert all(r["reason"] != "대조하지 못함 (명령 시간 제한)" for r in out["results"]), out


def test_check_rest_rows_without_header_pass_header_row_0(check_env, capsys, monkeypatch):
    # 8차 검토 (1.4.71): 원래 범위가 1행부터(머리글 없음)면 남은 행 명령에 --header-row 0 (머리글 없음).
    # 전에는 0이 조용히 무시되어 남은 범위 바로 위 데이터 행('2 A')을 값 열 단위 머리글로 읽어 판정이 달라졌음
    import openpyxl
    spec, book = check_env
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "시험"
    for i, (item, val) in enumerate([("소비 전류", "2 A"), ("소비 전류", 2.5), ("정격 전압", 5.1)], 1):
        ws.cell(i, 2, item)
        ws.cell(i, 4, val)
    wb.save(book)
    base = ["check", "--target", str(book), "--item-col", "B", "--value-col", "D", "--spec", str(spec.parent)]
    code, full = run_excel(capsys, *base, "--range", "B1:D3")
    assert code == 0, full
    want = [(r["cell"], r["verdict"], r["reason"]) for r in full["results"][1:]]
    monkeypatch.setattr(rt, "time_up", lambda reserve=0: True)
    code, out = run_excel(capsys, *base, "--range", "B1:D3")
    cmds = [c["command"] for c in out["next_commands"]]
    assert "--range B2:D3" in cmds[1] and "--header-row 0" in cmds[1], cmds[1]
    monkeypatch.setattr(rt, "time_up", lambda reserve=0: False)
    code, rest = run_excel(capsys, *split_command(cmds[1])[2:])
    assert code == 0 and [(r["cell"], r["verdict"], r["reason"]) for r in rest["results"]] == want, (rest, want)
    # 0은 머리글 없음, 범위 첫 행 이상은 인자 오류
    code, bad = run_excel(capsys, *base, "--range", "B2:D3", "--header-row", "2")
    assert code == 1 and bad["agent_directive"].startswith("STOP_RETRY:") and "없으면 0" in bad["message"], bad


def test_check_long_serial_number_is_fast(tmp_path):
    # 6차 검토 [높음]: 규격 문장의 긴 숫자열(시리얼 5500자)에서 조건 읽기가 제곱 시간이던 문제
    import time
    p = tmp_path / "s.txt"
    p.write_text("시리얼 " + "1234567890" * 550 + " 정격 전압\n", encoding="utf-8")
    docs = [{"file": str(p), "units": reader.parse_txt(str(p))[0]}]
    t = time.monotonic()
    r = _spec_check._judge_row(docs, "정격 전압", "5", "", "값(V)")
    assert time.monotonic() - t < 3.0 and r["verdict"] in ("판단필요", "근거없음"), r


def test_check_all_decided_is_still_unverified_and_save(check_env, capsys, tmp_path):
    # 4차 검토(1.4.67): 자연어 규격 해석은 휴리스틱이라 모든 행을 판정해도 VERIFIED가 아님 (참고 판정, 사용자 확인)
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.0), ("소비 전류", "2 A")])
    save = tmp_path / "대조결과.xlsx"
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D3", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec), "--save", str(save))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["verified"] is False, out
    assert "자동 판정은 참고용" in out["agent_directive"] and "condition" in out["agent_directive"]
    assert out["unverified"][0] == _spec_check.CHECK_ADVISORY and out["not_decided"] == 0
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


def test_check_does_not_mix_items_of_one_sentence(check_env, capsys):
    # 같은 문장의 다른 항목 조건을 섞어 VERIFIED였음: 두께 '5~100 mm', 외경 '5 mm 이하' (1.4.63)
    spec, book = check_env
    spec.write_text("3. 치수\n두께 5 mm 이상, 길이 100 mm 이하로 한다.\n외경 3 mm, 폭 5 mm 이하로 한다.\n", encoding="utf-8")
    _check_book(book, [("두께", 150), ("외경", 4), ("폭", 4)], header=("항목", "비고", "값(mm)"))
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D4", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec))
    got = {r["item"]: (r["verdict"], r.get("condition")) for r in out["results"]}
    assert got == {"두께": ("일치", "5 mm 이상"), "외경": ("불일치", "3 mm"), "폭": ("일치", "5 mm 이하")}, out


@pytest.mark.parametrize("sentence,value,verdict", [
    ("정격 전압은 5 V이며, 허용 범위는 4.75 V 이상 5.25 V 이하로 한다.", 5.1, "일치"),     # 1.4.64 회귀: '5 V'만 남아 불일치
    ("정격 전압은 5 V이며, 허용 범위는 4.75 V 이상 5.25 V 이하로 한다.", 5.3, "불일치"),
    ("정격 전압 5 V, 허용 범위 4.75~5.25 V", 5.1, "일치"),
    ("정격 전압: 5 V (허용 범위: 4.75 V ~ 5.25 V)", 5.1, "일치"),
    ("정격 전압 5 V, 시험 전압 4.75~5.25 V", 5.1, "판단필요"),       # 안전장치: 기준 값과 다르지만 같은 문장의 범위 안
    ("정격 전압 5 V, 시험 전압 4.75~5.25 V", 5.0, "일치"),
    ("정격 전압 5 V, 시험 전압 4.75~5.25 V", 6.0, "불일치"),
])
def test_check_limit_clause_belongs_to_item(check_env, capsys, sentence, value, verdict):
    spec, book = check_env
    spec.write_text(sentence + "\n", encoding="utf-8")
    _check_book(book, [("정격 전압", value)])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D2", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec))
    assert code == 0 and out["results"][0]["verdict"] == verdict, out


def _load_check_corpus():
    rows = []
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "check_corpus.tsv"), encoding="utf-8") as f:
        for n, line in enumerate(f, 1):
            line = line.rstrip("\n")
            if line.strip() and not line.startswith("#"):
                spec, item, value, verdict, *note = line.split("\t")
                rows.append(pytest.param(spec, item, value, verdict, id=f"L{n}"))
    return rows


CHECK_CORPUS = _load_check_corpus()


def test_check_corpus_size_and_spread():
    # 표준 판정 말뭉치 (1.4.66): 규격 문장 + 측정값 + 기대 판정. 앞으로의 check 회귀를 막는 기준
    verdicts = [p.values[3] for p in CHECK_CORPUS]
    assert len(verdicts) >= 60 and set(verdicts) == set(_spec_check.VERDICTS)


@pytest.mark.parametrize("spec,item,value,verdict", CHECK_CORPUS)
def test_check_corpus(tmp_path, spec, item, value, verdict):
    # 틀린 VERIFIED(일치·불일치) 금지가 최우선: 조건을 하나로 합치는 규칙이 확실하지 않으면 판단필요
    import reader
    path = tmp_path / "규격서.txt"
    path.write_text(spec + "\n", encoding="utf-8")
    docs = [{"file": str(path), "units": reader.parse_txt(str(path))[0]}]
    r = _spec_check._judge_row(docs, item, value, "", "값")
    assert r["verdict"] == verdict, r


@pytest.mark.parametrize("header,value,verdict", [
    (("항목", "비고", "값"), 30, "판단필요"),              # 칸·머리글에 단위 없음: 규격 단위(cm)로 추정하지 않음
    (("항목", "비고", "값(mm)"), 30, "일치"),              # 머리글 단위로 바꿔 비교 (30 mm <= 5 cm)
    (("항목", "비고", "값(mm)"), 60, "불일치"),
    (("항목", "비고", "값"), "30 mm", "일치"),             # 칸 값의 단위
    (("항목", "비고", "값(cm)"), 4, "일치"),               # 규격과 같은 단위의 머리글
])
def test_check_cell_without_unit_is_not_guessed(check_env, capsys, header, value, verdict):
    spec, book = check_env
    spec.write_text("두께는 5 cm 이하로 한다.\n", encoding="utf-8")
    _check_book(book, [("두께", value)], header=header)
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D2", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec))
    r = out["results"][0]
    assert r["verdict"] == verdict, out
    if verdict == "판단필요":
        assert "칸 단위 없음" in r["reason"] and "머리글" in r["reason"] and out["agent_directive"].startswith("UNVERIFIED:")


def _unit_book(path, head_e, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "시험"
    ws["B1"], ws["D1"], ws["E1"] = "항목", "값", head_e
    for i, (item, val, unit) in enumerate(rows, 2):
        ws.cell(i, 2, item)
        ws.cell(i, 4, val)
        ws.cell(i, 5, unit)
    wb.save(path)


ROWS_UNIT = [("두께", 30, "mm"), ("폭", 7, "cm"), ("길이", 2, None)]


@pytest.mark.parametrize("head_e,extra,verdicts,unit_col", [
    ("단위", [], ["일치", "불일치", "판단필요"], "E (머리글이 '단위'라 자동으로 씀)"),   # 값 열 오른쪽의 '단위' 열
    ("Unit", [], ["일치", "불일치", "판단필요"], "E (머리글이 '단위'라 자동으로 씀)"),
    ("비고", [], ["판단필요", "판단필요", "판단필요"], None),                        # 머리글이 '단위'가 아니면 쓰지 않음
    ("비고", ["--unit-col", "E"], ["일치", "불일치", "판단필요"], "E"),              # 직접 지정하면 머리글과 상관없이
])
def test_check_reads_unit_column(check_env, capsys, head_e, extra, verdicts, unit_col):
    """회사 PC 2026-11-06 O1: 항목·값·단위 3열 표의 값 칸에 단위가 없어 모든 행이 판단필요였다. 단위 열을 행마다 읽는다."""
    spec, book = check_env
    spec.write_text("두께는 5 cm 이하로 한다.\n\n폭은 5 cm 이하로 한다.\n\n길이는 5 cm 이하로 한다.\n", encoding="utf-8")
    _unit_book(book, head_e, ROWS_UNIT)
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D4", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec), *extra)
    assert [r["verdict"] for r in out["results"]] == verdicts, out
    assert out.get("unit_col") == unit_col, out


def test_check_unit_col_must_differ_from_value_col(check_env, capsys):
    spec, book = check_env
    _unit_book(book, "단위", ROWS_UNIT)
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D4", "--item-col", "B",
                          "--value-col", "D", "--unit-col", "D", "--spec", str(spec))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--unit-col" in out["message"], out


def test_check_empty_items_is_zero(check_env, capsys):
    spec, book = check_env
    _check_book(book, [])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D5", "--item-col", "B",
                          "--value-col", "D", "--spec", str(spec))
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert "0건(찾은 항목 없음)" in out["agent_directive"] and "next_commands" not in out


def test_check_bad_column_is_input_error(check_env, capsys):
    spec, book = check_env
    _check_book(book, [("정격 전압", 5.0)])
    code, out = run_excel(capsys, "check", "--target", str(book), "--range", "B2:D2", "--item-col", "항목",
                          "--value-col", "D", "--spec", str(spec))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out


def test_check_layer_b_item_marks_reason(tmp_path):
    # 7차 검토 (1.4.70): 숫자가 든 항목 글을 표기가 다른 곳(층 B)에서만 찾으면 reason에 표시, match_mode exact_mixed
    p = tmp_path / "s.txt"
    p.write_text("SOI-1 전압은 5 V 이하\n\nSOI 2 전압은 3 V 이하\n", encoding="utf-8")
    docs = [{"file": str(p), "units": reader.parse_txt(str(p))[0]}]
    r = _spec_check._judge_row(docs, "SOI 1 전압", "4", "", "값(V)")
    assert r["verdict"] == "일치" and r["match_mode"] == "exact_mixed" and _spec_check.LAYER_B_NOTE in r["reason"], r
    r = _spec_check._judge_row(docs, "SOI 2 전압", "4", "", "값(V)")       # 글자 그대로(층 A)는 그대로
    assert r["verdict"] == "불일치" and r["match_mode"] == "exact" and _spec_check.LAYER_B_NOTE not in r["reason"], r


def test_sheet_candidate_has_numbered_confirm_command(xl, capsys):
    """1.4.79: 비슷한 이름 [후보]에 번호를 붙이고 같은 번호의 confirm_commands가 이름을 바꾼 전체 명령 (사용자가 고른 뒤에만 실행)."""
    code, out = run_excel(capsys, "read", "--sheet", "요약표")
    assert code == 1 and "[후보] 1) 요약" in out["message"] and "next_commands" not in out, out
    assert "confirm_commands" in out["agent_directive"]
    argv = split_command(out["confirm_commands"][0]["command"])
    assert argv[0] == "python" and argv[2:] == ["read", "--sheet", "요약"], argv


def test_read_empty_range_is_explicit_zero(xl_table, capsys):
    """1.4.81 빈 결과 명시: 값이 없는 범위는 count 0과 '찾은 항목 없음'."""
    code, out = run_main(excel, capsys, "read", "--range", "H10:I12")
    assert code == 0 and out["count"] == 0 and out["preview"].startswith("0건:"), out
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "찾은 항목 없음" in out["agent_directive"], out


# 1.4.87: 이름 비교 규칙 하나 (rt.exact_key, rt.name_key). 열린 문서(COM)와 닫힌 파일(openpyxl)이 같은 [추천]을 냄

def _two_sheet_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.active["A1"] = "a"
    wb.create_sheet("요약")
    wb.save(path)
    return str(path)


def test_name_key_is_one_rule():
    assert rt.name_key("Sheet.1") == rt.name_key(" sheet_1 ") == rt.name_key("SHEET-1") == rt.name_key("Sheet1")
    assert rt.find_name(" SHEET1 ", ["Sheet1"]) == "Sheet1" and rt.find_name("Sheet.1", ["Sheet1"]) is None
    assert rt.suggest("Sheet.1", ["Sheet1", "요약"]) == ["Sheet1"]


@pytest.mark.parametrize("argv", [
    ["read", "--sheet", "Sheet.1"],                                     # 열린 문서 읽기 (not_found)
    ["write", "--sheet", "Sheet.1", "--range", "A1", "--value", "x"],  # 열린 문서 쓰기 (missing_sheet)
], ids=["open-read", "open-write"])
def test_sheet_spelling_same_recommendation_open(xl, capsys, argv):
    code, out = run_main(excel, capsys, *argv)
    assert code == 1 and "표기" in out["message"], out
    assert option_value(recommended_argv(out["message"], "excel"), "--sheet") == "Sheet1"


def test_sheet_spelling_same_recommendation_closed(tmp_path, capsys):
    path = _two_sheet_xlsx(tmp_path / "닫힌.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", path, "--sheet", "Sheet.1")
    assert code == 1 and "표기" in out["message"], out
    assert option_value(recommended_argv(out["message"], "excel"), "--sheet") == "Sheet1"


def test_closed_xlsx_load_retries_transient_lock(tmp_path, capsys, monkeypatch):
    """닫힌 xlsx를 여는 openpyxl.load_workbook도 retry_fs (백신·동기화 프로그램의 잠깐 잠금, JIT 복구 표)."""
    import openpyxl
    path = _two_sheet_xlsx(tmp_path / "잠김.xlsx")
    real = openpyxl.load_workbook
    calls = []

    def flaky(*a, **k):
        calls.append(1)
        if len(calls) == 1:
            raise PermissionError(13, "잠김", path)
        return real(*a, **k)
    monkeypatch.setattr(openpyxl, "load_workbook", flaky)
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)
    code, out = run_main(excel, capsys, "read", "--target", path, "--sheet", "Sheet1")
    assert code == 0 and out["values"][0][0] == "a" and len(calls) >= 2, out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A1·A2: --where의 %·단위·느슨한 날짜, 앞자리 0, 비교할 수 없는 조건은 0건이 아니라 STOP_RETRY
# ---------------------------------------------------------------------------


def _where_review_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["날짜", "비율", "코드", "금액", "메모"])
    for d, r, c, m, memo in [(dt.datetime(2026, 1, 5), 0.5, "010", 1000, "a"), (dt.datetime(2026, 1, 15), 0.75, "10", 1500, "b"),
                             (dt.datetime(2026, 1, 25), 0.2, "0010", 2000, "c")]:
        ws.append([d, r, c, m, memo])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "0%"
    wb.save(path)
    return str(path)


@pytest.mark.parametrize("expr, rows", [
    ("비율 >= 50%", [2, 3]), ("비율 < 50%", [4]), ("비율 = 75%", [3]),
    ("날짜 >= 2026-1-10", [3, 4]), ("날짜 > 2026/01/10", [3, 4]), ("날짜 < 2026.1.10", [2]), ("날짜 >= 2026년 1월 15일", [3, 4]),
    ("금액 >= 1500원", [3, 4]), ("금액 <= 1,500", [2, 3]),
    ("코드 = 010", [2]), ("코드 = 10", [3]), ("코드 != 010", [3, 4]),
])
def test_where_percent_units_loose_dates_and_leading_zero(tmp_path, capsys, expr, rows):
    p = _where_review_xlsx(tmp_path / "w.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p, "--where", expr)
    assert code == 0 and out["row_numbers"][1:] == rows, out


@pytest.mark.parametrize("expr", ["날짜 >= 지난달", "금액 > 많음", "메모 > 5"])
def test_where_uncomparable_condition_is_not_zero_verified(tmp_path, capsys, expr):
    p = _where_review_xlsx(tmp_path / "w.xlsx")
    code, out = run_main(excel, capsys, "read", "--target", p, "--where", expr)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "비교할 수 없습니다" in out["message"], out


def test_where_text_columns_compare_as_text_and_text_codes_stay_distinct(tmp_path, capsys):
    csv_path = tmp_path / "c.csv"
    csv_path.write_text("품번,비율,수량\n1A,50%,10개\n1B,20%,3개\n", encoding="utf-8")
    code, out = run_main(excel, capsys, "read", "--target", str(csv_path), "--where", "품번 = 1A")
    assert code == 0 and out["row_numbers"][1:] == [2], out
    code, out = run_main(excel, capsys, "read", "--target", str(csv_path), "--where", "비율 >= 30%; 수량 > 5")
    assert code == 0 and out["row_numbers"][1:] == [2], out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A3: 시트 XML의 <dimension>이 실제보다 작아도 끝까지 읽음 (read, --where, --search, export)
# ---------------------------------------------------------------------------


def _small_dimension_xlsx(tmp_path, ref):
    import openpyxl
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["이름", "값"])
    for i in range(1, 11):
        ws.append([f"n{i}", i * 10])
    wb.save(src)
    dst = tmp_path / "small.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = re.sub(rb'<dimension ref="[^"]*"/>', f'<dimension ref="{ref}"/>'.encode(), data)
            zout.writestr(item, data)
    return str(dst)


@pytest.mark.parametrize("ref", ["A1:B5", "A1"])
def test_small_dimension_reads_every_row(tmp_path, capsys, ref):
    p = _small_dimension_xlsx(tmp_path, ref)
    code, out = run_excel(capsys, "read", "--target", p)
    assert code == 0 and out["rows"] == 11 and out["values"][-1] == ["n10", 100], out
    code, out = run_excel(capsys, "read", "--target", p, "--where", "값 >= 50")
    assert code == 0 and out["count"] == 6, out
    code, out = run_excel(capsys, "read", "--target", p, "--search", "n9")
    assert code == 0 and out["count"] == 1, out
    code, out = run_excel(capsys, "read", "--target", p, "--range", "B:B")
    assert code == 0 and out["rows"] == 11, out
    code, out = run_excel(capsys, "export", "--target", p, "--save", str(tmp_path / "o.csv"))
    assert code == 0 and len(rt.read_csv_rows(out["output_file"])) == 11, out


def test_reader_reads_rows_past_small_dimension(tmp_path, capsys):
    p = _small_dimension_xlsx(tmp_path, "A1:B5")
    units, _ = reader.parse_xlsx(p)
    assert any("n10" in u["text"] for u in units) and len(units) == 11, units


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B6: 차트 시트가 있는 통합문서 (활성 시트·--sheet·compare)
# ---------------------------------------------------------------------------


def _chartsheet_xlsx(path, active_chart=True):
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "데이터"
    for r in [["품번", "값"], ["A1", 1], ["A2", 2]]:
        ws.append(r)
    cs = wb.create_chartsheet("차트")
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
    cs.add_chart(ch)
    if active_chart:
        wb.active = 1
    wb.save(path)
    return str(path)


def test_chartsheet_is_skipped_or_refused(tmp_path, capsys):
    p = _chartsheet_xlsx(tmp_path / "c.xlsx")
    code, out = run_excel(capsys, "read", "--target", p)
    assert code == 0 and out["values"][0] == ["품번", "값"], out
    code, out = run_excel(capsys, "read", "--target", p, "--sheet", "차트")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "차트 시트" in out["message"], out
    q = _chartsheet_xlsx(tmp_path / "d.xlsx")
    code, out = run_excel(capsys, "compare", "--target", p, "--target2", q)
    assert code == 0 and out["status"] == "ok", out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B7: XFD1048576에 떠돌이 칸(공백)이 있어도 read·compare가 데이터 범위만 (MemoryError, 140초 넘게 멈춤)
# ---------------------------------------------------------------------------


def test_stray_far_cell_does_not_blow_up_read_or_compare(tmp_path, capsys):
    import time
    import openpyxl
    paths = []
    for name, extra in (("old.xlsx", 1), ("new.xlsx", 2)):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["품번", "수량"])
        for i in range(1, 11):
            ws.append([f"P{i}", i * extra])
        ws["XFD1048576"] = " "
        p = tmp_path / name
        wb.save(p)
        paths.append(str(p))
    t0 = time.time()
    code, out = run_excel(capsys, "read", "--target", paths[0])
    assert code == 0 and out["rows"] == 11 and len(out["values"][0]) == 2, out
    code, out = run_excel(capsys, "compare", "--target", paths[0], "--target2", paths[1], "--key", "품번")
    assert code == 0 and out["status"] == "ok", out
    assert time.time() - t0 < 30


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B13 (엑셀 읽기): --save 확장자, 표시 형식 숫자 검색, B:B 행 수, 기간 칸, 깨진 파일, compare 자기 자신, CSV check
# ---------------------------------------------------------------------------


def test_b13_excel_read_small_fixes(tmp_path, capsys):
    import openpyxl
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["항목", "금액", "비율", "시간"])
    ws.append(["가", 1000, 0.25, dt.timedelta(days=1, hours=6)])
    ws.append(["나", 20, 0.5, dt.timedelta(hours=2, minutes=30)])
    for r in range(4, 40):
        ws.cell(r, 2).number_format = "0.00"          # 서식만 있는 끝 행
    wb.save(p)
    code, out = run_excel(capsys, "read", "--target", str(p), "--save", str(tmp_path / "res.xlsx"))
    assert code == 0 and out["output_file"].endswith("res.csv") and any(".xlsx" in r for r in out["recovered"]), out
    for q, cell in (("1,000", "B2"), ("25%", "C2")):
        code, out = run_excel(capsys, "read", "--target", str(p), "--search", q)
        assert code == 0 and out["count"] == 1 and out["matches"][0]["cell"] == cell, (q, out)
    code, out = run_excel(capsys, "read", "--target", str(p), "--range", "B:B")
    assert code == 0 and out["rows"] == 3, out
    code, out = run_excel(capsys, "read", "--target", str(p))
    assert out["values"][1][3] == "30:00:00" and out["values"][2][3] == "2:30:00", out
    z = tmp_path / "z.xlsx"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("hello.txt", "not excel")
    code, out = run_excel(capsys, "read", "--target", str(z))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    code, out = run_excel(capsys, "compare", "--target", str(p), "--target2", str(p), "--sheet", "Sheet")
    assert code == 1 and "같은 파일" in out["message"], out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 5장 5번: 큰 닫힌 파일을 읽다 시간 제한이 오면 JSON 없이 끊기지 않음 (read는 UNVERIFIED, export는 STOP_RETRY)
# ---------------------------------------------------------------------------


def test_big_closed_file_stops_before_time_limit(tmp_path, capsys, monkeypatch):
    import openpyxl
    p = tmp_path / "큰.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["번호", "값"])
    for i in range(1, 6001):
        ws.append([i, i * 2])
    wb.save(p)
    monkeypatch.setattr(rt, "time_up", lambda reserve=rt.TIME_RESERVE: True)
    code, out = run_excel(capsys, "read", "--target", str(p), "--where", "값 >= 0")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["time_limited"]["rows"] == 1999, out
    assert "--range" in out["agent_directive"] and out["count"] == 1998, out
    code, out = run_excel(capsys, "export", "--target", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--range" in out["message"], out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 4장 1번 (1.4.98): stats 표 집계 (닫힌 xlsx·CSV는 Office 없이, 열린 문서는 COM)
# ---------------------------------------------------------------------------


def _sales_csv(tmp_path):
    p = tmp_path / "매출.csv"
    p.write_text("부서,품번,금액,수량\n영업,0101,1000,2\n영업,0102,3000,1\n개발,0103,500,4\n개발,0104,1500,1\n기획,0105,200,3\n",
                 encoding="utf-8")
    return str(p)


def test_stats_totals_groups_where_and_top(tmp_path, capsys):
    p = _sales_csv(tmp_path)
    code, out = run_excel(capsys, "stats", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["columns"] == ["금액", "수량"], out    # 품번(앞자리 0)은 코드
    st = out["groups"][0]["stats"]["금액"]
    assert (st["sum"], st["mean"], st["min"], st["max"], st["count"]) == (6200, 1240, 200, 3000, 5), st
    assert "금액: 합계 6,200, 평균 1,240, 최소 200, 최대 3,000 (5개)" in out["preview"], out["preview"]
    code, out = run_excel(capsys, "stats", "--target", p, "--by", "부서", "--columns", "금액")
    assert [g["by"]["부서"] for g in out["groups"]] == ["개발", "기획", "영업"] and out["groups"][2]["stats"]["금액"]["sum"] == 4000, out
    code, out = run_excel(capsys, "stats", "--target", p, "--by", "부서", "--columns", "금액", "--top", "1")
    assert out["group_count"] == 1 and out["groups"][0]["by"]["부서"] == "영업", out
    code, out = run_excel(capsys, "stats", "--target", p, "--where", "금액 >= 1000", "--columns", "수량")
    assert out["rows_used"] == 3 and out["groups"][0]["stats"]["수량"]["sum"] == 4, out
    code, out = run_excel(capsys, "stats", "--target", p, "--columns", "부서")
    st = out["groups"][0]["stats"]["부서"]
    assert st["kind"] == "text" and st["unique"] == 3 and st["top"][0] == {"value": "개발", "count": 2}, st
    code, out = run_excel(capsys, "stats", "--target", p, "--where", "금액 > 99999")
    assert code == 0 and out["groups"] == [] and "0건" in out["agent_directive"], out
    code, out = run_excel(capsys, "stats", "--target", p, "--by", "부서", "--save", str(tmp_path / "집계.csv"))
    rows = rt.read_csv_rows(out["output_file"])
    assert rows[0][:3] == ["부서", "행 수", "금액 합계"] and len(rows) == 4, rows


def test_stats_closed_xlsx_and_open_document(tmp_path, capsys, xl_table):
    p = _table_xlsx(tmp_path / "목록.xlsx")
    code, out = run_excel(capsys, "stats", "--target", p, "--columns", "금액")
    assert code == 0 and out["groups"][0]["stats"]["금액"]["count"] >= 4, out
    code, out = run_main(excel, capsys, "stats", "--by", "부서", "--columns", "금액")      # 열린 문서 (가짜 COM)
    assert code == 0 and {g["by"]["부서"]: g["stats"]["금액"]["sum"] for g in out["groups"]} == {"영업": 500, "개발": 100, "기획": 1000}, out


def test_info_of_closed_xlsx_needs_no_office(tmp_path, capsys, monkeypatch):
    """2026-11-09 검토 4장 1번: 닫힌 파일의 시트 목록 (전에는 info가 늘 Excel을 띄움)."""
    p = _chartsheet_xlsx(tmp_path / "c.xlsx")
    monkeypatch.setattr(excel.off, "open_target", lambda *a, **k: (_ for _ in ()).throw(AssertionError("Office를 열면 안 됨")))
    code, out = run_excel(capsys, "info", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["sheets"][0] == {"name": "데이터", "used_range": "A1:B3", "rows": 3, "cols": 2, "hidden": False}, out
    assert out["sheets"][1] == {"name": "차트", "kind": "chart"} and out["active_sheet"] == "차트", out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 4장 2번 (1.4.99): --where 제목 행 아래의 머리글, 포함하지 않음 !~
# ---------------------------------------------------------------------------


def test_where_finds_header_below_title_and_not_contains(tmp_path, capsys):
    import openpyxl
    p = tmp_path / "기록.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "부품 검사 기록 (2026년 11월)"
    ws.append([])
    ws.append(["품번", "상태", "수량"])
    for r in [("P1", "완료", 3), ("P2", "검토 중", 5), ("P3", "완료(재검)", 1)]:
        ws.append(list(r))
    wb.save(p)
    code, out = run_excel(capsys, "read", "--target", str(p), "--where", "수량 >= 2")
    assert code == 0 and out["header_row_found"] == 3 and out["values"][0] == ["품번", "상태", "수량"], out
    assert out["row_numbers"] == [3, 4, 5] and out["count"] == 2, out
    code, out = run_excel(capsys, "read", "--target", str(p), "--where", "상태 !~ 완료")
    assert code == 0 and out["count"] == 1 and out["values"][1][0] == "P2", out
    code, out = run_excel(capsys, "read", "--target", str(p), "--where", "상태 ~ 완료; 수량 < 2")
    assert code == 0 and [r[0] for r in out["values"][1:]] == ["P3"], out


# ---------------------------------------------------------------------------
# 1.6.0 (검토 4장 4번): 중복 찾기 (읽기만)
# ---------------------------------------------------------------------------


def test_duplicates_by_columns_and_whole_row(tmp_path, capsys):
    p = tmp_path / "부품.csv"
    p.write_text("품번,규격,수량\n0102,M3,10\n0102 ,m3,5\n102,M3,1\n0105,M4,2\n0105,M4,2\n,,\n", encoding="utf-8-sig")
    code, out = run_excel(capsys, "duplicates", "--target", str(p), "--columns", "품번,규격")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["group_count"] == 2, out
    assert out["groups"][0]["rows"] == [2, 3] and out["groups"][1]["rows"] == [5, 6], out      # 공백·대소문자 무시, 앞자리 0은 다른 값
    assert out["duplicate_rows"] == 2 and "2, 3행" in out["preview"], out
    code, out = run_excel(capsys, "duplicates", "--target", str(p))                           # 모든 열이 같은 행
    assert code == 0 and [g["rows"] for g in out["groups"]] == [[5, 6]], out
    code, out = run_excel(capsys, "duplicates", "--target", str(p), "--columns", "수량,품번")
    assert code == 0 and [g["rows"] for g in out["groups"]] == [[5, 6]], out


def test_duplicates_zero_and_open_document(tmp_path, capsys, xl_table):
    p = tmp_path / "목록.csv"
    p.write_text("이름,값\n가,1\n나,2\n", encoding="utf-8-sig")
    code, out = run_excel(capsys, "duplicates", "--target", str(p))
    assert code == 0 and out["group_count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out
    code, out = run_main(excel, capsys, "duplicates", "--columns", "부서")                    # 열린 문서 (가짜 COM)
    assert code == 0 and out["group_count"] >= 1, out
