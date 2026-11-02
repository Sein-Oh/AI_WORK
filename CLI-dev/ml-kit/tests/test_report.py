# -*- coding: utf-8 -*-
"""ML 결과 보고서 (ml_report.py): structure, numbers and protocol checks on generated .docx files."""
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit"))
from make_sample import make_sample, make_class_sample  # noqa: E402
from test_cli_smoke import run_tool, run_fail  # noqa: E402

docx = pytest.importorskip("docx")
pytest.importorskip("matplotlib")

SECTIONS = ["1. 요약", "2. 데이터와 피처", "2.1. 파생변수", "2.2. 상관관계", "3. 분석 방법", "4. 타겟별 결과"]
REMOVED = ["목차", "5. 버전 이력", "6. 운영 가이드", "부록 A. 상관관계", "부록 B. 산출물 및 실행 환경"]  # 2.4.1에서 뺌


def _headings(path):
    d = docx.Document(path)
    return d, [p.text for p in d.paragraphs if p.style.name.startswith("Heading")]


def _all_text(d):
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            parts.extend(c.text for c in row.cells)
    return "\n".join(parts)


@pytest.fixture(scope="module")
def regression_report(tmp_path_factory):
    work = tmp_path_factory.mktemp("rep_reg")
    data = make_sample(str(work / "proc.csv"))
    feats = "반응온도,소비전력,압축압력,촉매비율,수분함량"
    run_tool("ml_train.py", "-f", data, "-x", feats, "-y", "생산효율", "--desc", "효율 예측", "--tag", "base")
    tr = run_tool("ml_train.py", "-f", data, "-x", "반응온도,소비전력,촉매비율,수분함량", "-y", "생산효율", "--tag", "pruned")
    slots = work / "slots.json"
    slots.write_text(json.dumps({"summary": "슬롯으로 작성한 요약 문장",
                                 "targets": {"생산효율(%)": {"features": "슬롯 피처 소견"}}}, ensure_ascii=False),
                     encoding="utf-8")
    out = work / "report.docx"
    payload = run_tool("ml_report.py", "-m", tr["version_dir"], "--slots", str(slots), "-o", str(out))
    return work, tr, payload, out


def test_regression_report_structure(regression_report):
    _, tr, payload, out = regression_report
    assert out.is_file() and payload["report_file"] == str(out)
    d, heads = _headings(str(out))
    for sec in SECTIONS:
        assert sec in heads, sec
    text = "\n".join(p.text for p in d.paragraphs)
    for sec in REMOVED:
        assert sec not in heads and sec not in text, sec
    assert "4.1. 생산효율(%)" in heads
    assert d.paragraphs[0].text == "ML 결과 보고서"  # tools/run_win_tests.py M6 reads the first paragraph
    assert len(d.inline_shapes) == payload["figure_count"] >= 2          # 상관관계 + 타겟마다 1장
    assert payload["table_count"] >= 3 and len(d.tables) >= 5 and payload["verdicts"]["생산효율(%)"] in ("적합", "조건부", "부적합")
    page_breaks = sum(1 for p in d.paragraphs for r in p.runs if 'w:br w:type="page"' in r._r.xml)
    assert page_breaks == 0                                              # 장마다 쪽을 넘기지 않음 (빈 공백)


def test_regression_report_numbers_and_slots(regression_report):
    work, tr, payload, out = regression_report
    d = docx.Document(str(out))
    text = _all_text(d)
    lb = json.loads(open(os.path.join(tr["version_dir"], "leaderboard.json"), encoding="utf-8").read())
    champ = lb["targets"]["생산효율(%)"]["champion"]
    assert f"{champ['val_r2']:.4f}" in text and f"{champ['mae']:.4f}" in text
    assert "슬롯으로 작성한 요약 문장" in text and "슬롯 피처 소견" in text
    assert set(payload["slots_used"]) == {"summary", "targets.생산효율(%).features"}


@pytest.fixture(scope="module")
def binary_report(tmp_path_factory):
    work = tmp_path_factory.mktemp("rep_cls")
    data = make_class_sample(str(work / "q.csv"), n=200, imbalanced=True)
    tr = run_tool("ml_train.py", "-f", data, "-x", "반응온도,촉매비율,수분함량", "-y", "품질등급",
                  "--task", "classification", "--positive-class", "불량")
    payload = run_tool("ml_report.py", "-m", tr["version_dir"])
    return tr, payload


def test_classification_report(binary_report):
    tr, payload = binary_report
    assert payload["task"] == "classification"
    assert os.path.dirname(payload["report_file"]) == os.path.dirname(tr["version_dir"])  # default: data folder
    d, heads = _headings(payload["report_file"])
    text = _all_text(d)
    assert "혼동행렬" in text and "판정 기준" in text and "불량" in text
    assert len(d.inline_shapes) == payload["figure_count"] >= 1  # 혼동행렬 + 피처 기여도 한 장
    assert "클래스별 성능" in text


def test_report_rejects_non_version_folder(regression_report, tmp_path):
    work, _, _, _ = regression_report
    payload = run_fail("ml_report.py", "-m", str(work))
    assert payload["agent_directive"].startswith("STOP_RETRY") and "leaderboard.json" in payload["message"]
    missing = run_fail("ml_report.py", "-m", str(tmp_path / "nope"))
    assert "버전 폴더" in missing["message"]


def test_data_dir_mode_picks_best_version_and_slots(regression_report):
    """-d <data_root> without -v uses the best version of the latest experiment; --slots text lands in the document."""
    work, tr, _, _ = regression_report
    root = os.path.dirname(tr["version_dir"])
    lines = open(os.path.join(root, "history.jsonl"), encoding="utf-8").read().splitlines()
    recs = [json.loads(x) for x in lines]
    best = max(recs, key=lambda r: r["mean_val_score"])["version"]
    slots = json.dumps({"performance": "직접 쓴 성능 소견", "operation": "직접 쓴 운영 소견"}, ensure_ascii=False)
    out = work / "slots.docx"
    payload = run_tool("ml_report.py", "-d", root, "--slots", slots, "-o", str(out))
    assert payload["best_version"] == payload["version"] == best
    text = _all_text(docx.Document(str(out)))
    assert "직접 쓴 성능 소견" in text and "직접 쓴 운영 소견" in text


def test_report_requires_oof_files(regression_report, tmp_path):
    """A version folder without oof_*.csv (retrain with 2.9.1 or later) gets a clear STOP_RETRY."""
    import shutil
    work, tr, _, _ = regression_report
    root = tmp_path / "copy_ml_data"
    shutil.copytree(os.path.dirname(tr["version_dir"]), root)
    vdir = root / os.path.basename(tr["version_dir"])
    for f in vdir.glob("oof_*.csv"):
        f.unlink()
    payload = run_fail("ml_report.py", "-m", str(vdir))
    assert "2.9.1" in payload["message"] and payload["agent_directive"].startswith("STOP_RETRY")


def test_derived_features_are_shown_with_formulas(tmp_path):
    """파생변수는 2.1에 계산식(× ÷ ² 표기)과 타겟별 기여도·순위로 표시하고, 피처 표에서 '파생'으로 구별 (2.4.1)."""
    data = make_sample(str(tmp_path / "d.csv"))
    tr = run_tool("ml_train.py", "-f", data, "-x", "반응온도,압축압력,촉매비율,수분함량", "-y", "생산효율",
                  "--eval", "온도x수분 = 반응온도 * 수분함량; 촉매제곱 = 촉매비율 ** 2; 압력비 = 압축압력 / 촉매비율")
    payload = run_tool("ml_report.py", "-m", tr["version_dir"], "-o", str(tmp_path / "r.docx"))
    text = _all_text(docx.Document(payload["report_file"]))
    assert "온도x수분 = 반응온도(°C) × 수분함량(%)" in text
    assert "촉매제곱 = 촉매비율(wt%)²" in text and "압력비 = 압축압력(bar) ÷ 촉매비율(wt%)" in text
    assert "생산효율(%): " in text and "위)" in text                       # 사용 타겟의 기여도와 순위
    assert "\n파생\n" in text                                               # 입력 피처 표의 유형


def test_pretty_formula():
    import ml_report
    assert ml_report.pretty_formula("a = x ** 2 + y**3 - z ** 0.5") == "x² + y³ - z ^ 0.5"
    assert ml_report.pretty_formula("b = sqrt(x) / y * 2") == "√(x) ÷ y × 2"
