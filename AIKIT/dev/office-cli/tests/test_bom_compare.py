"""BOM 비교 연습 샘플 (tools/make_bom_samples.py): 정답을 아는 계층 BOM 두 판과 --check.

맞대기 스크립트는 사내 LLM이 가이드(tools/BOM비교_스킬만들기_가이드.md)대로 회사에서 만든다. 여기서는 샘플 모양과 정답 확인만 시험한다.
"""
import csv
import importlib.util
import os
import sys

TOOLS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools")


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


def test_samples_have_hierarchy_and_repeated_parts(tmp_path):
    """중간 조립체 10개 이상, 조립체당 부품 5개 이상, 여러 조립체에 반복되는 부품, 중복 판."""
    mod = _samples(tmp_path)
    assert sum(1 for r in mod.old_rows() if r[0] == 1) >= 10
    kids = mod.ASSEMBLIES_KIDS
    assert all(len(k) >= 5 for k in kids.values())
    uses = {}
    for sa, k in kids.items():
        for pn, _ in k:
            uses.setdefault(pn, set()).add(sa)
    assert sum(1 for s in uses.values() if len(s) >= 3) >= 5
    for name in ("BOM_A_Rev1.xlsx", "BOM_A_Rev2.xlsx", "BOM_A_Rev2_중복.xlsx", "BOM_A_정답.csv"):
        assert (tmp_path / name).exists(), name


def test_check_accepts_answer_and_reports_differences(tmp_path, capsys):
    """--check는 결과 CSV의 앞 네 열(구분, 상위 조립체, Part Number, 바뀐 열)만 보고, 대소문자·공백은 무시."""
    mod = _samples(tmp_path)
    rows = [list(r) for r in csv.reader(open(tmp_path / "BOM_A_정답.csv", encoding="utf-8-sig"))]
    ok = tmp_path / "결과_정답그대로.csv"
    with open(ok, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([rows[0] + ["옛 값", "새 값"]] + [[a, b.lower(), f" {c} ", d, "", ""] for a, b, c, d in rows[1:]])
    assert mod.check(str(ok)) == 0, capsys.readouterr().out
    bad = tmp_path / "결과_하나빠짐.csv"
    with open(bad, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows[:-1])
    assert mod.check(str(bad)) == 1 and "빠진 것" in capsys.readouterr().out
