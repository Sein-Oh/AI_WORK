"""BOM 비교 시험 묶음 (tools/make_bom_samples.py, tools/bom-compare/scripts/compare.py): 계층 BOM 샘플을 정답대로 맞대는지.

compare.py는 사용자 조합 스킬의 스크립트(office-cli 도구가 아님)라 Office 없이 CSV만 다룬다. excel.py read --save가 만드는
CSV 모양(A1부터, 빈 행은 빈 줄)을 openpyxl로 만들어 넣는다.
"""
import csv
import importlib.util
import json
import os
import subprocess
import sys

TOOLS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools")
COMPARE = os.path.join(TOOLS, "bom-compare", "scripts", "compare.py")


def _samples(tmp_path):
    spec = importlib.util.spec_from_file_location("make_bom_samples", os.path.join(TOOLS, "make_bom_samples.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    old_argv = sys.argv
    sys.argv = ["make_bom_samples.py", "--dir", str(tmp_path)]
    try:
        assert mod.main() == 0
    finally:
        sys.argv = old_argv
    return mod


def _to_csv(xlsx):
    import openpyxl
    ws = openpyxl.load_workbook(xlsx).active
    rows = [["" if v is None else v for v in r] for r in ws.iter_rows(values_only=True)]
    out = os.path.splitext(xlsx)[0] + "_비교용.csv"
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    return out, f"A1:G{len(rows)}"


def _compare(old, new):
    (o, o_rng), (n, n_rng) = _to_csv(old), _to_csv(new)
    p = subprocess.run([sys.executable, COMPARE, "--old", o, "--new", n, "--new-range", n_rng, "--old-range", o_rng],
                       capture_output=True, text=True, encoding="utf-8", env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    return p.returncode, json.loads(p.stdout.strip().splitlines()[-1])


def test_hierarchical_bom_matches_answer_by_parent_and_part_number(tmp_path, capsys):
    """같은 부품이 여러 중간 조립체에 반복돼도 자리(상위 조립체 + Part Number)마다 맞대고, 함정 8가지는 결과에 없음."""
    mod = _samples(tmp_path)
    rows = mod.old_rows()
    assert sum(1 for r in rows if r[0] == 1) >= 10                                     # 중간 조립체 10개 이상
    kids = mod.ASSEMBLIES_KIDS
    assert all(len(k) >= 5 for k in kids.values())                                    # 조립체당 부품 5개 이상
    uses = {}
    for sa, k in kids.items():
        for pn, _ in k:
            uses.setdefault(pn, set()).add(sa)
    assert sum(1 for s in uses.values() if len(s) >= 3) >= 5                          # 여러 조립체에 반복되는 부품
    code, out = _compare(tmp_path / "BOM_A_Rev1.xlsx", tmp_path / "BOM_A_Rev2.xlsx")
    assert code == 0 and out["preview"] == "추가 4, 삭제 4, 변경 7자리(7칸)", out
    capsys.readouterr()
    assert mod.check(out["result_file"]) == 0, capsys.readouterr().out
    with open(out["memo_file"], encoding="utf-8-sig") as f:
        memos = dict(list(csv.reader(f))[1:])
    assert any("SA-03 아래" in m and "Qty: 4 -> 6" in m for m in memos.values()), memos


def test_dotted_level_gives_same_result(tmp_path, capsys):
    """ERP 들여쓰기 BOM의 점 표기 Level('.1', '..2')도 숫자 Level과 같은 결과 (회사 BOM에 맞출 때 코드를 고치지 않게)."""
    import openpyxl
    mod = _samples(tmp_path)
    for name in ("BOM_A_Rev1.xlsx", "BOM_A_Rev2.xlsx"):
        wb = openpyxl.load_workbook(tmp_path / name)
        ws = wb.active
        for row in ws.iter_rows(min_row=4, max_col=1):
            if isinstance(row[0].value, int) and row[0].value > 0:
                row[0].value = "." * row[0].value + str(row[0].value) if name.endswith("1.xlsx") else "." * row[0].value
        wb.save(tmp_path / name)
    code, out = _compare(tmp_path / "BOM_A_Rev1.xlsx", tmp_path / "BOM_A_Rev2.xlsx")
    assert code == 0 and out["preview"] == "추가 4, 삭제 4, 변경 7자리(7칸)", out
    capsys.readouterr()
    assert mod.check(out["result_file"]) == 0, capsys.readouterr().out


def test_unreadable_level_stops_with_row(tmp_path):
    """Level을 못 읽으면 맞대지 않고 몇 번째 줄인지 알림."""
    import openpyxl
    _samples(tmp_path)
    wb = openpyxl.load_workbook(tmp_path / "BOM_A_Rev2.xlsx")
    wb.active["A6"] = "하위"
    wb.save(tmp_path / "BOM_A_Rev2.xlsx")
    code, out = _compare(tmp_path / "BOM_A_Rev1.xlsx", tmp_path / "BOM_A_Rev2.xlsx")
    assert code == 1 and "6번째 줄" in out["message"] and "점 표기" in out["message"], out


def test_same_part_twice_under_one_assembly_stops(tmp_path):
    """판단 기준 4: 같은 상위 조립체 아래 같은 Part Number가 두 번이면 맞대지 않고 자리와 엑셀 행을 알림."""
    _samples(tmp_path)
    code, out = _compare(tmp_path / "BOM_A_Rev1.xlsx", tmp_path / "BOM_A_Rev2_중복.xlsx")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "SA-03 아래 P-0016" in out["message"] and "엑셀 19, 20행" in out["message"], out
    assert not os.path.exists(tmp_path / "BOM_A_Rev2_중복_비교용_비교결과.csv")
