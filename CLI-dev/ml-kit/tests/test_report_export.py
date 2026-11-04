# -*- coding: utf-8 -*-
"""Word 보고서(ml_report.py)와 성능예측표(ml_export.py, 엑셀 계산기).

보고서: 구성, 숫자 형식, Word 2007 스키마 순서, 손상된 템플릿 복구, 보고서와 성능예측표 세트.
성능예측표: 재학습으로 입력 고르기(상호작용 입력은 남김, --keep 존중, 8개까지), 수식만으로 보간하는 잠긴 엑셀,
LibreOffice가 있으면 수식을 다시 계산해 Python 값과 비교.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest
from scipy.interpolate import RegularGridInterpolator

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_sample import make_class_sample, make_sample  # noqa: E402
from mlkit_helpers import (FEATURES, HAS_DOCX, HAS_MPL, KIT, TARGET, ok_res, run_fail, run_raw, run_tool,  # noqa: E402
                           stop_msg, write_df)
import ml_export as mx  # noqa: E402
import ml_predict as mp  # noqa: E402
import ml_report as mr  # noqa: E402
import ml_train as mt  # noqa: E402

try:
    import docx
    import docx.api
    from docx.oxml.ns import qn
except ImportError:          # 보고서 검사(needs_report)는 건너뜀
    docx = qn = None

needs_report = pytest.mark.skipif(not (HAS_DOCX and HAS_MPL), reason="python-docx/matplotlib not installed")


# ---- Word 보고서 구성과 숫자 ----------------------------------------------------------------------------------


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


@needs_report
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


@needs_report
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


@needs_report
def test_classification_report(binary_report):
    tr, payload = binary_report
    assert payload["task"] == "classification"
    assert os.path.dirname(payload["report_file"]) == os.path.dirname(tr["version_dir"])  # default: data folder
    d, heads = _headings(payload["report_file"])
    text = _all_text(d)
    assert "혼동행렬" in text and "판정 기준" in text and "불량" in text
    assert len(d.inline_shapes) == payload["figure_count"] >= 1  # 혼동행렬 + 피처 기여도 한 장
    assert "클래스별 성능" in text


@needs_report
def test_report_rejects_non_version_folder(regression_report, tmp_path):
    work, _, _, _ = regression_report
    payload = run_fail("ml_report.py", "-m", str(work))
    assert payload["agent_directive"].startswith("STOP_RETRY") and "leaderboard.json" in payload["message"]
    missing = run_fail("ml_report.py", "-m", str(tmp_path / "nope"))
    assert "버전 폴더" in missing["message"]


@needs_report
def test_data_dir_mode_picks_best_version_and_slots(regression_report):
    """-d <데이터 폴더>만 주면(-v 없음) 최근 실험의 최고 버전을 쓰고, --slots 글이 문서에 들어간다."""
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


@needs_report
def test_report_requires_oof_files(regression_report, tmp_path):
    """oof_*.csv가 없는 버전 폴더(2.9.1 전에 학습)는 다시 학습하라는 STOP_RETRY."""
    import shutil
    work, tr, _, _ = regression_report
    root = tmp_path / "copy_ml_data"
    shutil.copytree(os.path.dirname(tr["version_dir"]), root)
    vdir = root / os.path.basename(tr["version_dir"])
    for f in vdir.glob("oof_*.csv"):
        f.unlink()
    payload = run_fail("ml_report.py", "-m", str(vdir))
    assert "2.9.1" in payload["message"] and payload["agent_directive"].startswith("STOP_RETRY")


@needs_report
def test_derived_features_are_shown_with_formulas(tmp_path):
    """파생변수는 2.1에 계산식(× ÷ ² 표기)과 타겟별 기여도·순위로 표시하고, 피처 표에서 '파생'으로 구별한다."""
    data = make_sample(str(tmp_path / "d.csv"))
    tr = run_tool("ml_train.py", "-f", data, "-x", "반응온도,압축압력,촉매비율,수분함량", "-y", "생산효율",
                  "--eval", "온도x수분 = 반응온도 * 수분함량; 촉매제곱 = 촉매비율 ** 2; 압력비 = 압축압력 / 촉매비율")
    payload = run_tool("ml_report.py", "-m", tr["version_dir"], "-o", str(tmp_path / "r.docx"))
    text = _all_text(docx.Document(payload["report_file"]))
    assert "온도x수분 = 반응온도(°C) × 수분함량(%)" in text
    assert "촉매제곱 = 촉매비율(wt%)²" in text and "압력비 = 압축압력(bar) ÷ 촉매비율(wt%)" in text
    assert "생산효율(%): " in text and "위)" in text                       # 사용 타겟의 기여도와 순위
    assert "\n파생\n" in text                                               # 입력 피처 표의 유형


@needs_report
def test_pretty_formula():
    import ml_report
    assert ml_report.pretty_formula("a = x ** 2 + y**3 - z ** 0.5") == "x² + y³ - z ^ 0.5"
    assert ml_report.pretty_formula("b = sqrt(x) / y * 2") == "√(x) ÷ y × 2"


def _tag_names(el):
    return [c.tag.rsplit("}", 1)[-1] for c in el]


def _in_order(names, order):
    pos = [order.index(n) for n in names]
    return pos == sorted(pos)


@needs_report
def test_report_docx_follows_schema_for_word_2007(tmp_path):
    """보고서 XML은 WordprocessingML 스키마의 순서와 이름을 따른다 (Word 2007은 w:start/w:end와 순서가 틀린 문서를 거부)."""
    data = tmp_path / "d.csv"
    make_sample(str(data))
    tr = run_tool("ml_train.py", "-f", str(data), "-x", "반응온도,촉매비율,수분함량", "-y", "생산효율")
    rep = run_tool("ml_report.py", "-d", tr["data_root"], "--out", str(tmp_path / "r.docx"))
    body = docx.Document(rep["report_file"]).element.body

    assert not list(body.iter(qn("w:start"))) and not list(body.iter(qn("w:end")))   # Word 2010+ names
    tcprs = list(body.iter(qn("w:tcPr")))
    assert tcprs and list(body.iter(qn("w:tcMar"))) and list(body.iter(qn("w:tcBorders")))
    for tcpr in tcprs:
        assert _in_order(_tag_names(tcpr), mr._TCPR_ORDER), _tag_names(tcpr)
    for tblpr in body.iter(qn("w:tblPr")):
        assert _in_order(_tag_names(tblpr), mr._TBLPR_ORDER), _tag_names(tblpr)
    for ppr in body.iter(qn("w:pPr")):
        assert _in_order(_tag_names(ppr), mr._PPR_ORDER), _tag_names(ppr)
    for box in list(body.iter(qn("w:tcMar"))) + list(body.iter(qn("w:tcBorders"))):
        assert _tag_names(box) == ["top", "left", "bottom", "right"]


def test_report_fmt_keeps_small_values():
    assert mr.fmt(0.123456) == "0.1235" and mr.fmt(1234.5) == "1,234.5000"
    assert mr.fmt(3.2e-6) == "3.200e-06"                                       # 2.9.12: 0.0000
    assert mr.fmt(0.004567) == "0.004567" and mr.fmt(0.0) == "0.0000"           # 2.9.15: 4 significant digits
    t = {"task": "regression", "oof": pd.DataFrame({"actual": [1e-6, 3e-6, 5e-6]}),
         "champ": {"champion_model": "Ridge", "val_r2": 0.95, "mae": 2.5e-7, "overfit_gap": 0.01}}
    text = mr.text_performance(t)
    assert "MAE는 2.500e-07" in text and "0.0000" not in text


def test_report_fmt_no_padding():
    assert mr.fmt(0.0066) == "0.0066"                                          # 2.9.13: 0.006600
    assert mr.fmt(0.004567) == "0.004567" and mr.fmt(4e-5) == "0.00004" and mr.fmt(3.2e-6) == "3.200e-06"  # 2.9.15
    assert mr.fmt(1234.5) == "1,234.5000" and mr.fmt(0.0) == "0.0000"


def test_report_fmt_significant_digits():
    assert mr.fmt(0.00012346) == "0.0001235"                                       # 2.9.14: 0.0001
    assert mr.fmt(0.0066) == "0.0066" and mr.fmt(0.00099) == "0.00099" and mr.fmt(0.001) == "0.0010"
    assert mr.fmt(0.12345678) == "0.1235" and mr.fmt(1234.56789) == "1,234.5679"
    assert mr.fmt(0.00012346, 3) == "0.000123" and mr.fmt(2.5e-7) == "2.500e-07"
    # no jump around 1e-3: neighbours keep the same number of significant digits
    assert mr.fmt(0.0009876) == "0.0009876" and mr.fmt(0.001234) == "0.001234"


def test_report_predict_command_layout_from_history():
    import ml_report as mr
    recs = [{"task": "regression", "data_profile": {"layout": {"header_row": 2, "unit_row": 3, "data_row": None}}}]
    ns = mr._layout_args(argparse.Namespace(header_row=None, unit_row=None, data_row=None), recs, "regression")
    assert mt.layout_flags(ns) == " --header-row 2 --unit-row 3"
    ns = mr._layout_args(argparse.Namespace(header_row=5, unit_row=None, data_row=None), recs, "regression")
    assert mt.layout_flags(ns) == " --header-row 5 --unit-row 3"


def test_report_reuses_single_best_version(tmp_path):
    pytest.importorskip("docx")
    pytest.importorskip("matplotlib")
    import docx
    r = np.random.default_rng(6)
    n = 50
    df = pd.DataFrame({"a": r.uniform(0, 10, n), "b": r.uniform(0, 5, n)})
    df["y1"] = 2 * df.a + r.normal(0, 0.3, n)
    df["y2"] = df.a + 3 * df.b + r.normal(0, 0.3, n)
    df = pd.concat([df, df.iloc[:4]], ignore_index=True)                       # 4 duplicate rows
    data = write_df(tmp_path / "multi.csv", df)
    tr = run_tool("ml_train.py", "-f", data, "-x", "a,b", "-y", "y1,y2")
    rep = run_tool("ml_report.py", "-f", data)
    assert rep["made_best"] is False and rep["best_version"] == tr["version"] == "v1"   # 2.9.13: rebuilt v2
    assert not os.path.isdir(os.path.join(tr["data_root"], "v2"))
    assert rep["best_targets"]["y1"]["source_version"] == "v1"
    text = "\n".join(p.text for p in docx.Document(rep["report_file"]).paragraphs)
    assert "중복 행 4개" in text                                                 # 2.9.13: not in the report


# ---- 손상된 Word 템플릿 (회사 PC 문서 보안) ------------------------------------------------------------


# 회사 PC에서 python-docx의 templates/default.docx가 표준 스타일 없는 대체 문서였다("no style with name 'Caption'").
# 보고서는 풀어 둔 템플릿 폴더(templates/default-docx-template)를 먼저 열고, 손상된 템플릿은 고쳐서 쓴다.


PLACEHOLDER = "보안 정책으로 보호된 문서입니다"


REMOVED_STYLES = ("Caption", "List Bullet", "Heading 2")


def make_damaged_default(path):
    """문서 보안이 바꾼 것 같은 docx: 정상 Word 파일이지만 표준 스타일이 없고 본문에 안내 문구가 있다."""
    d, _ = mr.docx_from_template_dir(mr.docx_template_dir())
    styles_el = d.styles.element
    for st in list(d.styles):
        if st.name in REMOVED_STYLES:
            styles_el.remove(st.element)
    d.add_paragraph(PLACEHOLDER)
    d.save(str(path))
    check = docx.Document(str(path))
    assert "Caption" not in {s.name for s in check.styles}
    return str(path)


@pytest.fixture
def recovered(monkeypatch):
    lst = []
    monkeypatch.setattr(mp, "_RECOVERED", lst)
    return lst


@needs_report
def test_directory_template_used_normally(tmp_path, monkeypatch, recovered):
    """템플릿 폴더가 있으면 default.docx는 열지 않는다 (손상돼 있어도)."""
    assert os.path.isdir(mr.docx_template_dir())
    bad = make_damaged_default(tmp_path / "default.docx")
    monkeypatch.setattr(docx.api, "_default_docx_path", lambda: bad)
    d = mr.new_report_document()
    names = {s.name for s in d.styles}
    assert set(mr.REPORT_STYLES) <= names
    assert len(d.paragraphs) == 0 and len(d.tables) == 0
    assert not any("default.docx" in r for r in recovered)     # folder used, not the damaged default.docx


@needs_report
def test_template_folder_with_missing_part(tmp_path, monkeypatch, recovered):
    """보안 제품이 템플릿 폴더의 docProps/thumbnail.jpeg를 지운 PC: python-docx만으로는 실패하지만(FileNotFoundError)
    보고서는 끊긴 관계를 빼고 계속한다.
    """
    import shutil
    folder = tmp_path / "default-docx-template"
    shutil.copytree(mr.docx_template_dir(), folder)
    thumb = folder / "docProps" / "thumbnail.jpeg"
    if thumb.exists():
        thumb.unlink()
    with pytest.raises(Exception):
        docx.Document(str(folder))                          # the failure seen on the company PC
    monkeypatch.setattr(mr, "docx_template_dir", lambda: str(folder))
    bad = make_damaged_default(tmp_path / "default.docx")
    monkeypatch.setattr(docx.api, "_default_docx_path", lambda: bad)
    d = mr.new_report_document()
    assert set(mr.REPORT_STYLES) <= {s.name for s in d.styles}
    assert any("빠진 파일" in r and "thumbnail.jpeg" in r for r in recovered)
    assert not any("default.docx" in r for r in recovered)
    d.add_paragraph("x", style="Caption")
    out = tmp_path / "saved.docx"
    d.save(str(out))
    assert docx.Document(str(out)).paragraphs[-1].text == "x"


@needs_report
def test_damaged_default_without_template_folder(tmp_path, monkeypatch, recovered):
    bad = make_damaged_default(tmp_path / "default.docx")
    monkeypatch.setattr(docx.api, "_default_docx_path", lambda: bad)
    monkeypatch.setattr(mr, "docx_template_dir", lambda: str(tmp_path / "없는폴더"))
    d = mr.new_report_document()
    assert set(mr.REPORT_STYLES) <= {s.name for s in d.styles}
    assert len(d.paragraphs) == 0
    assert any("default.docx 사용" in r for r in recovered)
    assert any("본문 1개" in r for r in recovered)
    assert any("스타일 없음" in r and "Caption" in r for r in recovered)
    d.add_paragraph("x", style="Caption")  # the call that failed on the company PC
    d.add_heading("h", level=2)


@needs_report
def test_both_templates_broken_is_stop_retry(tmp_path, monkeypatch, recovered):
    not_zip = tmp_path / "default.docx"
    not_zip.write_bytes(b"DRM-ENCRYPTED\x00\x01")  # like pptx/templates/default.pptx on the company PC
    monkeypatch.setattr(docx.api, "_default_docx_path", lambda: str(not_zip))
    monkeypatch.setattr(mr, "docx_template_dir", lambda: str(tmp_path / "없는폴더"))
    with pytest.raises(mp.InputError) as ei:
        mr.new_report_document()
    assert "python -m pip install --force-reinstall python-docx" in str(ei.value)


RUNNER = r"""
import sys
import docx.api
sys.path.insert(0, sys.argv[1])
import ml_report
bad, missing = sys.argv[2], sys.argv[3]
docx.api._default_docx_path = lambda: bad
ml_report.docx_template_dir = lambda: missing
sys.argv = ["ml_report.py"] + sys.argv[4:]
ml_report.main()
"""


@needs_report
def test_report_succeeds_with_damaged_templates(tmp_path):
    """끝까지: 손상된 default.docx + 템플릿 폴더 없음 -> 보고서 성공, 캡션 있음, 안내 문구 없음."""
    data = make_sample(str(tmp_path / "drm.csv"), n=80)
    tr = run_tool("ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET, "--desc", "drm")
    bad = make_damaged_default(tmp_path / "default.docx")
    out = tmp_path / "report.docx"
    proc = subprocess.run(
        [sys.executable, "-c", RUNNER, KIT, bad, str(tmp_path / "없는폴더"), "-m", tr["version_dir"], "-o", str(out)],
        capture_output=True, text=True, encoding="utf-8", timeout=600,
    )
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert len(lines) == 1, proc.stdout + proc.stderr
    payload = json.loads(lines[0])
    assert proc.returncode == 0 and payload["status"] == "ok", proc.stdout + proc.stderr
    assert any("스타일 없음" in r for r in payload["recovered"])
    assert "[WARN] Word" in proc.stderr
    d = docx.Document(str(out))
    captions = [p.text for p in d.paragraphs if p.style.name == "Caption"]
    assert len(captions) >= payload["figure_count"] + 1
    assert d.paragraphs[0].text == "ML 결과 보고서"
    assert PLACEHOLDER not in "\n".join(p.text for p in d.paragraphs)


# ---- 보고서와 성능예측표 세트 --------------------------------------------------------------------------------------


# 회귀 보고서는 같은 <데이터>_ml_data 폴더에 성능예측표_<데이터>_<버전>.xlsx도 만든다 (conftest가 끈 것을 여기서만 켬).
# 성능예측표를 만들지 못해도 보고서는 성공한다.


@pytest.fixture
def with_export(monkeypatch):
    monkeypatch.delenv("ML_KIT_SKIP_EXPORT", raising=False)


def _calc_data(tmp_path):
    r = np.random.default_rng(12)
    n = 250
    df = pd.DataFrame({"온도": r.uniform(0, 1, n), "압력": r.uniform(0, 1, n), "잡음": r.uniform(0, 1, n)})
    df["효율"] = 8 * df.온도 + 3 * df.압력 + r.normal(0, 0.05, n)
    df["등급"] = np.where(df.효율 > df.효율.median(), "상", "하")
    path = tmp_path / "p.csv"
    df.to_csv(path, index=False)
    return path


def test_report_makes_calculator_in_the_same_folder(tmp_path, with_export):
    path = _calc_data(tmp_path)
    tr = run_tool("ml_train.py", "-f", str(path), "-x", "온도,압력,잡음", "-y", "효율")
    rep = run_tool("ml_report.py", "-m", tr["version_dir"], "--title", "시험 효율 예측")
    root = os.path.dirname(tr["version_dir"])
    assert os.path.dirname(rep["report_file"]) == root and os.path.isfile(os.path.join(root, "history.jsonl"))
    assert rep["export_file"] == os.path.join(root, "성능예측표_p_v1.xlsx") and os.path.isfile(rep["export_file"])
    assert rep["export"]["효율"]["inputs"] == ["온도", "압력"] and "성능예측표" in rep["preview"]
    assert "export_file" in rep["agent_directive"]
    a = rep["export"]["효율"]["agreement"]                                       # 2.9.4: agreement with the model
    assert 95 < a["agreement_pct"] <= 100 and a["p95_diff_pct"] >= 0
    import openpyxl
    wb = openpyxl.load_workbook(rep["export_file"])
    assert wb["계산기"]["B2"].value == "시험 효율 예측"
    notes = " ".join(str(c.value) for row in wb["계산기"].iter_rows() for c in row if isinstance(c.value, str))
    assert f"평균 {a['agreement_pct']:.1f}% 일치" in notes
    from docx import Document
    text = "\n".join(p.text for p in Document(rep["report_file"]).paragraphs)
    assert "성능예측표 (엑셀 계산기): 성능예측표_p_v1.xlsx" in text and f"평균 {a['agreement_pct']:.1f}% 일치" in text


def test_classification_report_has_no_calculator(tmp_path, with_export):
    path = _calc_data(tmp_path)
    tr = run_tool("ml_train.py", "-f", str(path), "-x", "온도,압력", "-y", "등급", "--task", "classification")
    rep = run_tool("ml_report.py", "-m", tr["version_dir"])
    assert "export_file" not in rep and "export_note" not in rep


def test_calculator_failure_keeps_the_report(tmp_path, with_export):
    path = _calc_data(tmp_path)
    tr = run_tool("ml_train.py", "-f", str(path), "-x", "온도,압력", "-y", "효율")
    os.remove(os.path.join(tr["version_dir"], "model_dataset.csv"))          # calculator needs the training snapshot
    rep = run_tool("ml_report.py", "-m", tr["version_dir"])
    assert os.path.isfile(rep["report_file"]) and "export_file" not in rep
    assert rep["export_note"].startswith("성능예측표를 만들지 못했습니다") and "export_note" in rep["agent_directive"]


def test_report_unverified_when_calculator_unverified():
    result = {"preview": "p", "verified": True, "agent_directive": "VERIFIED: 보고서 저장."}
    ex = {"result": {"output_file": "x.xlsx", "targets": {}, "verified": False,
                     "unverified": ["엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함"]}}
    mr.attach_export(result, ex)
    assert result["verified"] is False and result["export_verified"] is False      # 2.9.14: VERIFIED
    assert result["agent_directive"].startswith("UNVERIFIED:")
    assert result["unverified"] == ["성능예측표: 엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함"]
    ok = {"preview": "p", "verified": True, "agent_directive": "VERIFIED: r."}
    mr.attach_export(ok, {"result": {"output_file": "x.xlsx", "targets": {}, "verified": True}})
    assert ok["verified"] is True and ok["agent_directive"].startswith("VERIFIED:")


def test_report_passes_overwrite_and_targets(monkeypatch, tmp_path):
    import ml_export
    seen = {}

    def fake(args):
        seen.update(vars(args))
        return {"output_file": "x.xlsx", "targets": {}, "verified": True, "warnings": ["w1"]}
    monkeypatch.delenv("ML_KIT_SKIP_EXPORT", raising=False)
    monkeypatch.setattr(ml_export, "execute_export", fake)
    ex = mr.make_export(str(tmp_path), "regression", None, ["생산효율"], overwrite=True)
    assert seen["overwrite"] is True and seen["targets"] == "생산효율"
    result = {"preview": "p", "agent_directive": "VERIFIED: r."}
    mr.attach_export(result, ex)                                          # E7
    assert result["export_warnings"] == ["w1"] and "export_warnings" in result["agent_directive"]


# ---- 성능예측표: 입력 고르기, 잠긴 엑셀, 수식 다시 계산 ----------------------------------------------------


SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


@pytest.fixture(scope="module")
def export_version(tmp_path_factory):
    """유량 x 농도가 상호작용(유량 혼자는 덜 중요해 보임), 잡음은 무관한 데이터의 버전."""
    d = tmp_path_factory.mktemp("exp")
    r = np.random.default_rng(21)
    n = 500
    df = pd.DataFrame({k: r.uniform(0, 1, n) for k in ["온도", "유량", "농도", "잡음"]})
    df["효율"] = 10 * np.sin(3 * df.온도) + 6 * df.유량 * df.농도 + r.normal(0, 0.2, n)
    df["등급"] = np.where(df.효율 > df.효율.median(), "상", "하")
    path = d / "s.csv"
    df.to_csv(path, index=False)
    res = run_tool("ml_train.py", "-f", str(path), "-x", "온도,유량,농도,잡음", "-y", "효율", "--desc", "시험용 공정 효율 데이터. 두 번째 문장.")
    return d, path, res["version_dir"]


def _grid(wb, sheet, k):
    ws = wb[sheet]
    A = np.array([r for r in ws.iter_rows(min_row=2, max_col=k + 1, values_only=True)], dtype=float)
    axes = [np.unique(A[:, j]) for j in range(k)]
    return RegularGridInterpolator(axes, A[:, k].reshape([len(a) for a in axes]))


def _recalc(path, tmp):
    out = tmp / "calc"
    out.mkdir(exist_ok=True)
    shutil.copy(path, out / "t.xlsx")
    subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out / "o"), str(out / "t.xlsx")],
                   capture_output=True, timeout=600)
    return openpyxl.load_workbook(out / "o" / "t.xlsx", data_only=True)


def test_search_keeps_interacting_input_and_builds_locked_workbook(export_version, tmp_path):
    _, _, vdir = export_version
    r = run_tool("ml_export.py", "-m", vdir)
    t = r["targets"]["효율"]
    assert t["inputs"] == ["온도", "유량", "농도"] and t["dropped"] == ["잡음"]          # 유량 kept despite low importance
    assert [c["count"] for c in t["candidates"]] == [4, 3, 2, 1] and "0.005" in t["reason"]
    assert os.path.isfile(os.path.join(r["export_model_dir"], "model_효율.joblib"))
    assert r["title"] == "시험용 공정 효율 데이터." and r["next_commands"][0]["command"].count("--inputs") == 1
    wb = openpyxl.load_workbook(r["output_file"])
    assert wb.sheetnames == ["계산기", "일괄계산", "사용법", "표_효율", "정확도", "정보"]
    assert all(wb[n].protection.sheet for n in wb.sheetnames) and wb.security.lockStructure
    assert wb["표_효율"].sheet_state == "veryHidden"                                    # grid not shown (2.9.3)
    assert all(wb[n].sheet_state == "visible" for n in wb.sheetnames if not n.startswith("표_"))
    calc = wb["계산기"]
    unlocked = [c.coordinate for row in calc.iter_rows() for c in row if not c.protection.locked]
    assert unlocked == ["C6", "C7", "C8"] and calc["B6"].value == "온도"
    batch = wb["일괄계산"]
    assert not batch["A3"].protection.locked and batch["E3"].protection.locked
    assert wb["표_효율"].max_row - 1 == t["table_rows"]
    a = t["agreement"]
    assert a["points"] == 2000 and a["agreement_pct"] > 90 and a["max_diff_pct"] >= a["p95_diff_pct"]
    assert any("일치도" in str(c.value) for row in wb["정확도"].iter_rows() for c in row)
    assert any(str(c.value).startswith("모델 예측과 평균") for row in wb["정보"].iter_rows() for c in row if c.value)
    assert any("엑셀" in u for u in r.get("unverified", [])) or r["verified"]          # Excel check only on Windows


def test_status_cells_have_visible_colors(export_version):
    # conditional formats: Excel reads bgColor for a solid fill, so fgColor alone shows no color
    _, _, vdir = export_version
    wb = openpyxl.load_workbook(run_tool("ml_export.py", "-m", vdir)["output_file"])
    for sheet in ("계산기", "일괄계산"):
        rules = [(rule.formula[0], rule.dxf) for cf in wb[sheet].conditional_formatting for rule in cf.rules]
        ok = [d for f, d in rules if f.endswith('="정상"')]
        warn = [d for f, d in rules if "범위 밖" in f and "참고용" in f]
        assert ok and warn and len(ok) == len(warn)
        assert all(d.fill.bgColor.rgb.endswith("DCFCE7") and d.font.color.rgb.endswith("166534") for d in ok)
        assert all(d.fill.bgColor.rgb.endswith("FCE4D6") and d.font.color.rgb.endswith("C2410C") for d in warn)


def test_keep_and_inputs_and_errors(export_version):
    _, _, vdir = export_version
    k = run_tool("ml_export.py", "-m", vdir, "--keep", "잡음", "--title", "짧은 제목")
    assert "잡음" in k["targets"]["효율"]["inputs"] and k["title"] == "짧은 제목"
    i = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,유량")
    assert i["targets"]["효율"]["inputs"] == ["온도", "유량"] and len(i["targets"]["효율"]["candidates"]) == 1
    assert "[후보] '온도'" in run_fail("ml_export.py", "-m", vdir, "--keep", "온도x")["message"]     # 2.9.33: was [추천]
    assert "--keep" in run_fail("ml_export.py", "-m", vdir, "--keep", "잡음", "--inputs", "온도")["message"]
    proc, payload = run_raw("ml_export.py", "-h")
    assert proc.returncode == 0 and payload["command"] == "help"


def test_classification_only_version_is_stopped(export_version):
    _, path, _ = export_version
    cls = run_tool("ml_train.py", "-f", str(path), "-x", "온도,유량", "-y", "등급", "--task", "classification")
    assert "숫자 타겟" in run_fail("ml_export.py", "-m", cls["version_dir"])["message"]


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_formulas_recalculate_to_the_python_values(export_version, tmp_path):
    _, _, vdir = export_version
    r = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,유량,농도", "--save", str(tmp_path / "f.xlsx"))
    src = openpyxl.load_workbook(r["output_file"])
    interp = _grid(src, "표_효율", 3)
    # calculator at default inputs (medians) + 3 batch rows, one of them outside the training range
    src["일괄계산"]["A3"], src["일괄계산"]["B3"], src["일괄계산"]["C3"] = 0.21, 0.73, 0.4
    src["일괄계산"]["A4"], src["일괄계산"]["B4"], src["일괄계산"]["C4"] = 0.9, 0.1, 0.95
    src["일괄계산"]["A5"], src["일괄계산"]["B5"], src["일괄계산"]["C5"] = 1.05, 0.5, 0.5          # out of range
    mod = tmp_path / "m.xlsx"
    src.save(mod)
    out = _recalc(mod, tmp_path)
    calc = out["계산기"]
    x = [calc[f"C{r}"].value for r in (6, 7, 8)]
    assert calc["C11"].value == pytest.approx(float(interp([x])[0]), abs=1e-9)                # result row = 5 + inputs + 3
    assert calc["F11"].value == "정상" and calc["F6"].value == "정상"
    b = out["일괄계산"]
    for row, xs in ((3, (0.21, 0.73, 0.4)), (4, (0.9, 0.1, 0.95))):
        assert b[f"D{row}"].value == pytest.approx(float(interp([xs])[0]), abs=1e-9)
        assert b[f"E{row}"].value < b[f"D{row}"].value < b[f"F{row}"].value
    assert b["G5"].value == "참고용 (학습 범위 밖)" and b["D6"].value in (None, "")


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_derived_features_stay_inside_the_model(tmp_path):
    r = np.random.default_rng(4)
    n = 300
    df = pd.DataFrame({"a": r.uniform(1, 2, n), "b": r.uniform(1, 2, n)})
    df["y"] = 3 * df.a * df.b + r.normal(0, 0.05, n)
    p = tmp_path / "e.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "y", "--eval", "ab = a * b")["version_dir"]
    res = run_tool("ml_export.py", "-m", v, "--inputs", "a,b", "--save", str(tmp_path / "e.xlsx"))
    t = res["targets"]["y"]
    assert t["inputs"] == ["a", "b"]                                                     # raw inputs only on the sheet
    wb = openpyxl.load_workbook(res["output_file"])
    assert [wb["계산기"][f"B{r}"].value for r in (6, 7)] == ["a", "b"]
    out = _recalc(res["output_file"], tmp_path)
    x = [out["계산기"]["C6"].value, out["계산기"]["C7"].value]
    assert out["계산기"]["C10"].value == pytest.approx(3 * x[0] * x[1], abs=0.15)


def test_several_inputs_and_outputs(tmp_path):
    """출력마다 자기 입력(학습 순서)을 쓴다. 계산기는 입력의 합집합을 받고, 출력은 자기 입력이 다 차면 계산한다."""
    r = np.random.default_rng(31)
    n = 300
    df = pd.DataFrame({"온도": r.uniform(0, 1, n), "압력": r.uniform(0, 1, n), "유량": r.uniform(0, 1, n), "분진": r.uniform(0, 1, n)})
    df["효율"] = 10 * np.sin(3 * df.온도) + r.normal(0, 0.1, n)
    df["전력"] = 5 * df.압력 * df.유량 + 2 * df.유량 + r.normal(0, 0.05, n)
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "온도,압력,유량,분진", "-y", "효율,전력")["version_dir"]
    res = run_tool("ml_export.py", "-m", v, "--save", str(tmp_path / "m.xlsx"))
    assert list(res["targets"]) == ["효율", "전력"]                                       # -y order
    assert res["targets"]["효율"]["inputs"] == ["온도"] and res["targets"]["전력"]["inputs"] == ["압력", "유량"]
    wb = openpyxl.load_workbook(res["output_file"])
    calc = wb["계산기"]
    assert [calc[f"B{r}"].value for r in (6, 7, 8)] == ["온도", "압력", "유량"]          # union, 분진 dropped everywhere
    assert [calc[f"B{r}"].value for r in (11, 12)] == ["효율", "전력"]
    assert "COUNT($C$6)=1" in calc["C11"].value and "COUNT($C$7,$C$8)=2" in calc["C12"].value
    assert wb.sheetnames[3:5] == ["표_효율", "표_전력"]


def test_workbook_has_no_full_paths(export_version):
    """<데이터>_ml_data 앞의 폴더 경로(사용자 이름, 회사 폴더)는 개인 정보다: 통합 문서에 넣지 않는다."""
    d, _, vdir = export_version
    res = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,유량,농도")
    wb = openpyxl.load_workbook(res["output_file"])
    parent = os.path.dirname(os.path.dirname(vdir))
    texts = [str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert not any(parent in t for t in texts)
    assert any(t == os.path.join("s_ml_data", "v1") for t in texts)


def test_agreement_basis_near_zero():
    """0 근처를 지나는 출력은 범위 대비로 비교한다 (아주 작은 값의 %는 의미가 없음)."""
    import ml_export as ex
    model = np.linspace(-1, 1, 101)
    a = ex.agreement(model + 0.02, model, {"residual_lower": -0.1, "residual_upper": 0.1})
    assert a["basis"] == "출력 범위 대비" and a["agreement_pct"] == pytest.approx(99.0)
    assert a["within_interval_pct"] == 100.0
    b = ex.agreement(np.full(10, 101.0), np.full(10, 100.0) + np.arange(10) * 0.0, None)
    assert b["basis"] == "예측값 대비" and b["agreement_pct"] == pytest.approx(99.0) and b["within_interval_pct"] is None


# ---- 성능예측표: 학습 범위, 일정한 입력, 8개 이하 입력, 파생 입력 --------------------------------------


def _strings(wb, sheet):
    return " ".join(str(c.value) for row in wb[sheet].iter_rows() for c in row if c.value is not None)


def test_calculator_range_ignores_excluded_rows_and_constant_input(tmp_path):
    r = np.random.default_rng(5)
    n = 300
    df = pd.DataFrame({"반응온도": r.uniform(150, 200, n), "압력": r.uniform(1, 3, n), "pH": 7.0})
    df["효율"] = 0.4 * df.반응온도 + 6 * df.압력 + r.normal(0, 0.3, n)
    df.loc[10, "반응온도"] = 2000                                # typo, excluded as an automatic outlier
    p = tmp_path / "c.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "반응온도,압력,pH", "-y", "효율")["version_dir"]
    res = run_tool("ml_export.py", "-m", v, "--save", str(tmp_path / "c.xlsx"))
    t = res["targets"]["효율"]
    assert "pH" not in t["inputs"] and any("pH" in w for w in res["warnings"])
    wb = openpyxl.load_workbook(res["output_file"])
    text = _strings(wb, "계산기") + _strings(wb, "일괄계산") + _strings(wb, "정보")
    assert "2000" not in text and "1999" not in text                       # range / out-of-range checks use 150~200


def test_inputs_and_keep_drop_constant_with_hint(tmp_path):
    r = np.random.default_rng(6)
    df = pd.DataFrame({"a": r.uniform(0, 1, 200), "k": 1.0})
    df["y"] = 3 * df.a + r.normal(0, 0.05, 200)
    p = tmp_path / "k.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "a,k", "-y", "y")["version_dir"]
    bad = run_fail("ml_export.py", "-m", v, "--inputs", "k")
    # 2.9.33: [후보] (another input set is the user's choice), was "[추천] a"
    assert bad["agent_directive"].startswith("STOP_RETRY") and "[후보] a" in bad["message"] and "[추천]" not in bad["message"]
    ok = run_tool("ml_export.py", "-m", v, "--keep", "k", "--save", str(tmp_path / "k.xlsx"))
    assert ok["targets"]["y"]["inputs"] == ["a"]


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_eight_inputs_use_range_sums_and_recalculate(tmp_path):
    r = np.random.default_rng(8)
    n = 400
    cols = [f"x{i}" for i in range(8)]
    df = pd.DataFrame({c: r.uniform(0, 1, n) for c in cols})
    df["y"] = sum((i + 1) * df[c] for i, c in enumerate(cols)) + r.normal(0, 0.05, n)
    p = tmp_path / "e.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", ",".join(cols), "-y", "y")["version_dir"]
    res = run_tool("ml_export.py", "-m", v, "--inputs", ",".join(cols), "--save", str(tmp_path / "e.xlsx"))
    wb = openpyxl.load_workbook(res["output_file"])
    sums = [c.value for row in wb["계산기"].iter_rows() for c in row if isinstance(c.value, str) and "SUM(" in c.value]
    assert sums and all(s.count(",") < 255 for s in sums)
    out = _recalc(res["output_file"], tmp_path)
    x = [out["계산기"][f"C{6 + i}"].value for i in range(8)]
    want = float(_grid(openpyxl.load_workbook(res["output_file"]), "표_y", 8)([x])[0])
    assert out["계산기"][f"C{5 + 8 + 3}"].value == pytest.approx(want, abs=1e-9)


def test_derived_sources_are_whole_names_and_chains_drop_together():
    raw = ["온도", "반응온도", "압력"]
    evals = ["t2 = 반응온도 * 반응온도", "tp = t2 * 압력"]
    assert mx.evals_for(["반응온도", "압력"], evals, raw) == evals            # '온도' is not inside '반응온도'
    assert mx.evals_for(["온도", "압력"], evals, raw) == []                    # t2 dropped -> tp dropped


def test_export_constant_inputs_candidates(name_files):
    proc, res = run_raw("ml_train.py", "-f", name_files["const"], "-x", "온도,압력,라인", "-y", "수율")
    ok_res(res, proc)
    proc, res = run_raw("ml_export.py", "-m", os.path.join(str(name_files["tmp"]), "const_ml_data", "v1"), "--inputs", "라인")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 온도, 압력" in msg, msg


@pytest.fixture(scope="module")
def nine(tmp_path_factory):
    work = tmp_path_factory.mktemp("nine")
    r = np.random.default_rng(9)
    cols = [f"x{i}" for i in range(9)]
    df = pd.DataFrame(r.uniform(0, 1, (120, 9)), columns=cols)
    df["y"] = sum((i + 1) * df[c] for i, c in enumerate(cols)) + r.normal(0, 0.05, 120)
    p = work / "nine.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", ",".join(cols), "-y", "y")["version_dir"]
    return {"work": work, "v": v, "cols": cols}


def test_more_than_eight_inputs_recommends_a_set(nine):
    bad = run_fail("ml_export.py", "-m", nine["v"], "--inputs", ",".join(nine["cols"]))
    assert bad["agent_directive"].startswith("STOP_RETRY") and "[추천]" in bad["message"]
    rec = bad["message"].split("[추천]")[1].strip().split(",")
    assert 1 <= len(rec) <= 8 and set(rec) <= set(nine["cols"])


def test_automatic_search_stays_within_eight(nine):
    keep = ",".join(nine["cols"][2:])                        # 7 kept: only two candidates to try
    res = run_tool("ml_export.py", "-m", nine["v"], "--keep", keep, "--save", str(nine["work"] / "n.xlsx"))
    t = res["targets"]["y"]
    assert len(t["inputs"]) <= 8
    big = [c for c in t["candidates"] if c["count"] == 9][0]
    assert "command" not in big and big["table_rows"] is None and "note" in big
    assert any("8개" in w for w in res["warnings"])


def test_keep_over_eight_is_refused(nine):
    bad = run_fail("ml_export.py", "-m", nine["v"], "--keep", ",".join(nine["cols"]))
    assert bad["agent_directive"].startswith("STOP_RETRY") and "8개" in bad["message"]


@pytest.fixture(scope="module")
def nine2(tmp_path_factory):
    work = tmp_path_factory.mktemp("nine2")
    r = np.random.default_rng(9)
    cols = [f"x{i}" for i in range(9)]
    df = pd.DataFrame(r.uniform(0, 1, (120, 9)), columns=cols)
    df["ya"] = sum((i + 1) * df[c] for i, c in enumerate(cols)) + r.normal(0, 0.05, 120)
    df["yb"] = 3 * df.x0 + 2 * df.x1 + r.normal(0, 0.05, 120)
    p = work / "nine2.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", ",".join(cols), "-y", "ya,yb")["version_dir"]
    return {"work": work, "v": v, "cols": cols}


def test_inputs_over_eight_recommends_every_target(nine2):
    spec = f"ya={','.join(nine2['cols'])};yb=x0,x1"
    bad = run_fail("ml_export.py", "-m", nine2["v"], "--inputs", spec)
    rec = bad["message"].split("[추천] ")[1].strip()
    parts = dict(p.split("=", 1) for p in rec.split(";"))
    assert parts["yb"] == "x0,x1"                                              # 2.9.12: [추천] ya=... only
    assert 1 <= len(parts["ya"].split(",")) <= 8


def test_inputs_spec_safe():
    assert mx.spec_safe({"y": ["a", "b"]})
    assert not mx.spec_safe({"y": ["a,b", "c"]}) and not mx.spec_safe({"y=1": ["a"]}) and not mx.spec_safe({"y": ["a;b"]})


def test_spec_safe_single_target_ignores_target_name():
    assert mx.spec_safe({"y=1,a": ["a", "b"]}, single=True)                    # 2.9.13: False
    assert not mx.spec_safe({"y": ["a,b"]}, single=True) and not mx.spec_safe({"y": ["a=b"]}, single=True)
    assert not mx.spec_safe({"y=1": ["a"]}) and not mx.spec_safe({"y": ["a;b"]})


# ---- 성능예측표: 타겟마다 입력, --overwrite, 작은 값의 정밀도, 시트 이름 ----------------------


@pytest.fixture(scope="module")
def two_export(tmp_path_factory):
    """ya는 a,b, yb는 a,c를 쓰는 최고 버전."""
    work = tmp_path_factory.mktemp("two")
    r = np.random.default_rng(2)
    n = 150
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n), "c": r.uniform(0, 1, n)})
    df["ya"] = 3 * df.a + 2 * df.b + r.normal(0, 0.02, n)
    df["yb"] = 2 * df.a + 4 * df.c + r.normal(0, 0.02, n)
    p = work / "two.csv"
    df.to_csv(p, index=False)
    run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "ya")
    run_tool("ml_train.py", "-f", str(p), "-x", "a,c", "-y", "yb")
    best = run_tool("ml_train.py", "-f", str(p), "--best")
    return {"work": work, "v": best["best_version_dir"]}


def test_keep_and_inputs_apply_to_targets_that_have_them(two_export):
    res = run_tool("ml_export.py", "-m", two_export["v"], "--keep", "b", "--save", str(two_export["work"] / "k.xlsx"))
    assert "b" in res["targets"]["ya"]["inputs"] and res["targets"]["ya"]["keep"] == ["b"]
    assert res["targets"]["yb"]["keep"] == []
    bad = run_fail("ml_export.py", "-m", two_export["v"], "--keep", "zzz")
    assert bad["agent_directive"].startswith("STOP_RETRY")
    # rebuild commands keep the other target's inputs (per-target --inputs), no single -y
    cmds = [c["command"] for c in res["next_commands"]]
    assert cmds and all(" -y " not in c and "ya=" in c and "yb=" in c for c in cmds)
    assert res["truncated"] is False and res["next_commands_count"] == len(cmds)
    alt = [c for c in res["targets"]["ya"]["candidates"] if "command" in c][0]
    spec = alt["command"].split("--inputs ")[1].split(" --title")[0].strip('"')
    again = run_tool("ml_export.py", "-m", two_export["v"], "--inputs", spec, "--save", str(two_export["work"] / "k2.xlsx"))
    assert again["targets"]["ya"]["inputs"] == alt["inputs"]
    assert again["targets"]["yb"]["inputs"] == res["targets"]["yb"]["inputs"]


def test_plain_inputs_list_per_target(two_export):
    res = run_tool("ml_export.py", "-m", two_export["v"], "--inputs", "a,c", "--save", str(two_export["work"] / "p.xlsx"))
    assert res["targets"]["ya"]["inputs"] == ["a"] and res["targets"]["yb"]["inputs"] == ["a", "c"]


def test_export_and_map_overwrite(two_export, tmp_path):
    save = str(two_export["work"] / "ow.xlsx")
    first = run_tool("ml_export.py", "-m", two_export["v"], "--inputs", "a", "--save", save)
    second = run_tool("ml_export.py", "-m", two_export["v"], "--inputs", "a", "--save", save)
    assert second["output_file"] != first["output_file"] and second.get("recovered")
    third = run_tool("ml_export.py", "-m", two_export["v"], "--inputs", "a", "--save", save, "--overwrite")
    assert third["output_file"] == first["output_file"] and not third.get("recovered")
    p = tmp_path / "g.csv"
    pd.DataFrame({"A": [0.0, 10, 20], "C": [1.0, 2, 3]}).to_csv(p, index=False)
    m1 = run_tool("ml_map.py", "-f", str(p), "-i", "A=5", "--save", "out.xlsx")
    m2 = run_tool("ml_map.py", "-f", str(p), "-i", "A=5", "--save", "out.xlsx", "--overwrite")
    assert m1["output_file"] == m2["output_file"]


def test_value_format_follows_magnitude():
    assert mx.value_format(np.array([120.0, 250.0])) == "0.00"
    assert mx.value_format(np.array([0.004, 0.005])) == "0.000000"
    assert mx.value_format(np.array([2e-6, 8e-6])) == "0.000E+00"
    assert mx._sig(0.0000123456) == 1.235e-05


def test_tiny_target_calculator_keeps_precision(tmp_path):
    r = np.random.default_rng(3)
    n = 200
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n)})
    df["y"] = 2e-6 + 4e-6 * df.a + 2e-6 * df.b + r.normal(0, 2e-8, n)
    p = tmp_path / "tiny.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "y")["version_dir"]
    meta = json.load(open(os.path.join(v, "model_y.json"), encoding="utf-8"))
    assert meta["prediction_interval"]["residual_upper"] != 0                     # was rounded to 6 decimals
    res = run_tool("ml_export.py", "-m", v, "--inputs", "a,b", "--save", str(tmp_path / "tiny.xlsx"))
    t = res["targets"]["y"]
    assert t["rmse"] > 0 and t["agreement"]["agreement_pct"] > 99
    wb = openpyxl.load_workbook(res["output_file"])
    grid = np.array([row[2] for row in wb["표_y"].iter_rows(min_row=2, max_col=3, values_only=True)], dtype=float)
    assert grid.min() > 1e-6 and len(np.unique(grid)) > 100                     # not rounded to 0.000002 steps
    calc = wb["계산기"]
    cells = [c for row in calc.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=IF(COUNT(")]
    assert any(c.number_format == "0.000E+00" for c in cells)
    assert any('TEXT(' in str(c.value) and '"0.000E+00"' in str(c.value) for row in calc.iter_rows() for c in row)


class _SumModel:
    def predict(self, X):
        return np.asarray(X, dtype=float).sum(axis=1)


@pytest.mark.parametrize("k", [6, 8])
def test_uncovered_threshold_does_not_flag_training_rows(k):
    r = np.random.default_rng(k)
    cols = [f"x{i}" for i in range(k)]
    data = pd.DataFrame(r.uniform(0, 1, (400, k)), columns=cols)
    data.iloc[:, 0] *= 0.5                                   # no measured data for x0 > 0.5 ...
    data.iloc[:3, :] = 0.0
    data.iloc[:3, 0] = 1.0                                   # ... except one corner, so the range of x0 is 0~1
    cand = {"inputs": cols, "evals": [], "pipe": _SumModel(), "interval": None,
            "schema": {"numeric": cols, "categorical": {}, "datetime": {}, "input_columns": cols}}
    tab = mx.build_grid(cand, data)
    f = RegularGridInterpolator(tab["axes"], tab["dist"].reshape([tab["n"]] * k))
    lo, span = tab["lo"], tab["hi"] - tab["lo"]
    at_rows = f(np.clip(data.to_numpy(), lo, tab["hi"]))
    assert (at_rows > tab["uncovered_dist"]).mean() < 0.05             # 2.9.11: 75-100% of training rows flagged
    q = r.uniform(0, 1, (500, k)) * span + lo
    q[:, 0] = 0.9 + 0.1 * r.uniform(0, 1, 500)                # x0 0.9~1: inside the range, no measured rows near
    assert (f(q) > tab["uncovered_dist"]).mean() > 0.9


def test_sheet_names_unique_without_case():
    used = {"계산기"}
    a = mx._sheet_name("표_", "Temp", used)
    b = mx._sheet_name("표_", "TEMP", used)
    assert a.lower() != b.lower()
