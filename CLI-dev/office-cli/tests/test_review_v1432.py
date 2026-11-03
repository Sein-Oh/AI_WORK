"""1.4.32 (23차 점검) 회귀 테스트.

H1: 닫힌 xlsx를 범위 없이 읽으면 끝쪽의 빈 행(값이 모두 None)을 잘라, 계산된 값이 없는 수식만 있는 합계 행이 빠지고 VERIFIED였음
    (수식만 있는 시트는 머리글만). 크기 정보(<dimension>)가 없는 파일은 끝쪽 수식 열도 잘렸음.
H2: xlsxwriter는 수식 칸에 <v>0</v>을 쓰고 workbook.xml에 <calcPr fullCalcOnLoad="1"/>을 둠: 0을 값으로 VERIFIED.
L2(안전): 수식 칸 확인이 실패(MemoryError 포함)하면 0개로 보아 VERIFIED였음 -> 확인 못함(UNVERIFIED).
"""
import json
import types
import zipfile

import pytest

import _runtime as rt
import excel
import office
import reader


def _main(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


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
    """머리글, 3행, 그리고 =SUM 합계 행 (리뷰 mk2)."""
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


# H1: 끝쪽 수식 행·열을 자르지 않음 ------------------------------------------------------------------------------------

def test_h1_total_row_kept_and_unverified(tmp_path, capsys):
    p = _total_xlsx(tmp_path / "t.xlsx")
    code, out = _main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["rows"] == 5 and out["range"] == "A1:B5" and out["values"][4] == ["", ""]
    assert out["uncomputed_formulas"] == ["매출!B5"] and out["uncomputed_count"] == 1


def test_h1_all_formula_sheet(tmp_path, capsys):
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
    code, out = _main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["rows"] == 5 and out["uncomputed_count"] == 8


def test_h1_search_save_export_transfer(tmp_path, capsys):
    p = _total_xlsx(tmp_path / "t.xlsx")
    code, out = _main(excel, capsys, "read", "--target", p, "--search", "600")
    assert code == 0 and out["count"] == 0 and out["range"] == "A1:B5" and out["agent_directive"].startswith("UNVERIFIED:"), out
    code, out = _main(excel, capsys, "read", "--target", p, "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["rows"] == 5 and out["agent_directive"].startswith("UNVERIFIED:"), out
    code, out = _main(excel, capsys, "export", "--target", p, "--save", str(tmp_path / "e.csv"))
    assert code == 0 and out["rows"] == 5 and out["agent_directive"].startswith("UNVERIFIED:"), out
    rows, sheet, info = office._excel_range_rows(types.SimpleNamespace(from_target=p, sheet=None, range=None))
    assert len(rows) == 5 and info["cells"] == ["매출!B5"]


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


def test_h1_no_dimension_keeps_formula_column(tmp_path, capsys):
    p = _nodim_xlsx(tmp_path / "n.xlsx")
    code, out = _main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["range"] == "A1:C4" and out["cols"] == 3 and out["rows"] == 4, out
    assert out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 3
    code, out = _main(excel, capsys, "read", "--target", p, "--range", "C:C")
    assert code == 0 and out["range"] == "C1:C4" and out["uncomputed_count"] == 3, out
    code, out = _main(excel, capsys, "read", "--target", p, "--range", "A1:XFD1048576")    # 큰 범위 줄이기
    assert code == 0 and out["range"] == "A1:C4" and out["uncomputed_count"] == 3, out


# H2: fullCalcOnLoad ------------------------------------------------------------------------------------------------

def test_h2_full_calc_on_load_cached_zero(tmp_path, capsys):
    p = _patch(_total_xlsx(tmp_path / "x.xlsx"), "<v></v>", "<v>0</v>")      # xlsxwriter 모양
    assert b'fullCalcOnLoad="1"' in zipfile.ZipFile(p).read("xl/workbook.xml")
    code, out = _main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["values"][4] == ["", 0] and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_formulas"] == ["매출!B5"] and "fullCalcOnLoad" in out["agent_directive"]
    code, out = _main(excel, capsys, "read", "--target", p, "--range", "B5")          # 빈 칸이 없는 범위도
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 1, out
    code, out = _main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_formulas"] == ["매출!B5"], out


def test_h2_true_flag_and_excel_saved(tmp_path, capsys):
    t = _patch(_total_xlsx(tmp_path / "t.xlsx"), "<v></v>", "<v>600</v>", ' fullCalcOnLoad="true"')
    assert rt.xlsx_uncomputed_formulas(t)["count"] == 1 and rt.xlsx_full_calc_on_load(t)
    e = _patch(_total_xlsx(tmp_path / "e.xlsx"), "<v></v>", "<v>600</v>", "")         # 엑셀이 저장하면 표시가 지워짐
    assert rt.xlsx_uncomputed_formulas(e)["count"] == 0 and not rt.xlsx_full_calc_on_load(e)
    code, out = _main(excel, capsys, "read", "--target", e)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["values"][4] == ["", 600], out


# L2: 확인 실패는 확인 못함 -----------------------------------------------------------------------------------------

def _boom(*a, **k):
    raise MemoryError()


def test_l2_scan_failure_is_unknown(tmp_path, capsys, monkeypatch):
    import xml.etree.ElementTree as ET
    p = _total_xlsx(tmp_path / "t.xlsx")
    monkeypatch.setattr(ET, "iterparse", _boom)
    info = rt.xlsx_uncomputed_formulas(p)
    assert info["unknown"] is True and info["error"] == "MemoryError" and rt.uncomputed_found(info)
    assert rt.uncomputed_fields(info) == {"uncomputed_check": "unknown"}
    code, out = _main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_check"] == "unknown" and "수식 계산 여부를 확인하지 못함" in out["agent_directive"]
    code, out = _main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_check"] == "unknown", out


def test_l2_area_stops_after_last_row(tmp_path):
    p = _total_xlsx(tmp_path / "t.xlsx")
    # 범위 끝 행(3) 뒤의 XML이 깨져 있어도 범위 안만 확인하고 끝남
    _patch(p, '<row r="4"', '<row r="4"><broken')
    info = rt.xlsx_uncomputed_formulas(p, "매출", (1, 1, 3, 2))
    assert info["count"] == 0 and "unknown" not in info
    assert rt.xlsx_uncomputed_formulas(p, "매출")["unknown"] is True
