"""1.4.36 회귀 테스트: excel.py read --formulas (엑셀 수식 읽기).

- 닫힌 xlsx: openpyxl data_only=False로 읽은 범위만 다시 읽어 수식 칸(보통·공유·배열)을 찾고, 값은 계산된 값(없으면 null)
- 계산되지 않은 수식(openpyxl이 쓴 파일)은 value null + 기존 UNVERIFIED(uncomputed_formulas) 그대로
- 수식 0개, 200개 상한과 full_result_file(--save면 저장한 CSV 옆), CSV, Office 없이 닫힌 .xls 거절
- 열린 문서(가짜 COM): 범위 전체 Formula 한 번, HasFormula, 배열 수식 한 번, Formula2 우선과 Formula 대체
- 오류 로그에 수식 글이 남지 않음
"""
import json
import os
import re
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
from fake_excel import FakeExcelWorkbook, FakeRange, FakeXlApp
from fake_office import make_win32com


def _main(capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        excel.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


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


# 닫힌 xlsx ------------------------------------------------------------------------------------------------------

def test_closed_plain_shared_array_and_constant(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = _main(capsys, "read", "--target", p, "--formulas")
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
    code, out = _main(capsys, "read", "--target", p)
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
    code, out = _main(capsys, "read", "--target", p, "--formulas")
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
    code, out = _main(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and out["formula_count"] == 0 and out["formulas"] == [] and out["formulas_truncated"] is False
    assert out["agent_directive"].startswith("VERIFIED:") and "수식 0개" in out["agent_directive"]
    assert "수식이 없습니다" in out["agent_directive"] and "수식 없음" in out["preview"]


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


def test_cap_truncated_full_result_file(tmp_path, capsys, isolated_dirs):
    p = _many_xlsx(tmp_path / "n.xlsx")
    code, out = _main(capsys, "read", "--target", p, "--formulas")
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
    code, out = _main(capsys, "read", "--target", p, "--formulas", "--save", save)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["output_file"] == save and out["full_result_file"] == str(tmp_path / "out" / "값_formulas.csv")
    assert len(out["formulas"]) == 200 and out["formulas_truncated"] is True and out["formula_count"] == 250
    rows = [r for r in rt.read_csv_rows(out["full_result_file"]) if any(r)]
    assert len(rows) == 251 and rows[1] == ["B1", "=A1+0", "1", ""]


def test_save_small_list_also_saved(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = _main(capsys, "read", "--target", p, "--formulas", "--save", str(tmp_path / "v.csv"))
    assert code == 0 and out["verified"] is True and out["formulas_truncated"] is False
    rows = [r for r in rt.read_csv_rows(out["full_result_file"]) if any(r)]
    assert ["C1", "=A1:A3*10", "1000", "C1:C3"] in rows and len(rows) == 6


def test_search_matches_formula_text(tmp_path, capsys):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = _main(capsys, "read", "--target", p, "--search", "SUM(")
    assert code == 0 and out["count"] == 0                                 # 값만 찾으면 없음 (그대로)
    code, out = _main(capsys, "read", "--target", p, "--search", "SUM(", "--formulas")
    assert code == 0 and out["count"] == 1
    assert out["matches"] == [{"cell": "A4", "value": 600, "formula": "=SUM(A1:A3)"}]
    assert out["formula_count"] == 5


# CSV, .xls ------------------------------------------------------------------------------------------------------

def test_csv_has_no_formulas(tmp_path, capsys):
    p = tmp_path / "a.csv"
    p.write_text("이름,값\n가,=1+1\n", encoding="utf-8")
    code, out = _main(capsys, "read", "--target", str(p), "--formulas")
    assert code == 0 and out["formula_count"] == 0 and out["formulas"] == []
    assert out["agent_directive"].startswith("VERIFIED:") and "CSV 파일은 수식을 저장하지 않아" in out["agent_directive"]


def test_closed_xls_without_office_is_stop_retry(tmp_path, capsys, monkeypatch, isolated_dirs):
    monkeypatch.setattr(off, "win32com", None)
    p = tmp_path / "old.xls"
    p.write_bytes(b"\xd0\xcf\x11\xe0" + b"\0" * 100)
    code, out = _main(capsys, "read", "--target", str(p), "--formulas")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "엑셀에서 이 파일을 연 뒤" in out["message"] and "old.xls" in out["message"]


# 열린 문서 (가짜 COM) ------------------------------------------------------------------------------------------

@pytest.fixture
def xl(monkeypatch):
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


def test_com_bulk_formula_hasformula_and_array(xl, capsys):
    code, out = _main(capsys, "read", "--range", "A1:C3", "--formulas")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["formulas"] == [
        {"cell": "C1", "formula": "=A1:A2*2", "value": 20, "array": True, "array_range": "C1:C2"},
        {"cell": "A3", "formula": "=SUM(A1:A2)", "value": 30}]
    assert out["formula_count"] == 2
    reads = [x for x in xl.log if x[0] == "FormulaRead"]
    assert reads == [("FormulaRead", "$A$1:$C$3")]                         # 범위 전체를 한 번에
    has = [x[1] for x in xl.log if x[0] == "HasFormula"]
    assert has == ["$A$1:$C$3", "$B$1"]                                     # 값과 같은 글인 칸만 따로 확인


def test_com_no_formulas_skips_formula_read(xl, capsys):
    code, out = _main(capsys, "read", "--range", "A1:A2", "--formulas")
    assert code == 0 and out["formula_count"] == 0 and "수식이 없습니다" in out["agent_directive"]
    assert not [x for x in xl.log if x[0] == "FormulaRead"]


def test_com_no_arrays_skips_cell_checks(xl, capsys):
    xl.arrays.clear()
    code, out = _main(capsys, "read", "--range", "A3", "--formulas")
    assert code == 0 and out["formulas"] == [{"cell": "A3", "formula": "=SUM(A1:A2)", "value": 30}]
    assert [x[1] for x in xl.log if x[0] == "HasArray"] == ["$A$3"]


def test_com_formula2_preferred(xl, capsys, monkeypatch):
    def formula2(self):
        base = self.Formula
        return tuple(tuple("=SORT(A1:A2)" if f == "=SUM(A1:A2)" else f for f in row) for row in base)
    monkeypatch.setattr(FakeRange, "Formula2", property(formula2), raising=False)
    code, out = _main(capsys, "read", "--range", "A1:C3", "--formulas")
    assert code == 0 and {"cell": "A3", "formula": "=SORT(A1:A2)", "value": 30} in out["formulas"]


def test_com_formula2_error_falls_back_to_formula(xl, capsys, monkeypatch):
    def formula2(self):
        raise Exception("Formula2 없음 (엑셀 2016)")
    monkeypatch.setattr(FakeRange, "Formula2", property(formula2), raising=False)
    code, out = _main(capsys, "read", "--range", "A1:C3", "--formulas")
    assert code == 0 and {"cell": "A3", "formula": "=SUM(A1:A2)", "value": 30} in out["formulas"]


# 오류 로그 -----------------------------------------------------------------------------------------------------

def _log_text(isolated_dirs):
    p = os.path.join(isolated_dirs["data"], "logs", "errors.txt")
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


def test_error_log_has_no_formula_text(tmp_path, capsys, monkeypatch, isolated_dirs):
    p = _mixed_xlsx(tmp_path / "m.xlsx")

    def boom(*a):
        raise ValueError("=SECRET_FORMULA(B2)")
    monkeypatch.setattr(excel, "_closed_formulas", boom)
    code, out = _main(capsys, "read", "--target", p, "--formulas")
    assert code == 1 and out["agent_directive"].startswith("ERROR:"), out
    assert "SECRET_FORMULA" not in json.dumps(out, ensure_ascii=False)
    log = _log_text(isolated_dirs)
    assert "수식 읽기 실패 (ValueError)" in log and "SECRET_FORMULA" not in log


def test_error_log_search_formula_text_hidden(tmp_path, capsys, isolated_dirs):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, out = _main(capsys, "read", "--target", p, "--sheet", "없는시트", "--formulas", "--search", "=SUM(A1:A3)")
    assert code == 1
    log = _log_text(isolated_dirs)
    assert log and "SUM(A1:A3)" not in log and "--formulas" in log


def test_success_writes_no_log(tmp_path, capsys, isolated_dirs):
    p = _mixed_xlsx(tmp_path / "m.xlsx")
    code, _ = _main(capsys, "read", "--target", p, "--formulas")
    assert code == 0 and _log_text(isolated_dirs) == ""
