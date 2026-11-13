# -*- coding: utf-8 -*-
"""Word 보고서(ml_report.py)와 성능예측표(ml_export.py, 엑셀 계산기).

보고서: 구성, 숫자 형식, WordprocessingML 스키마 순서, 손상된 템플릿 복구, --file은 최종 버전(마지막 학습), 보고서와 성능예측표 세트.
성능예측표: 버전의 모델을 엑셀 수식으로 그대로 옮김(나무는 잎마다 입력 범위, 선형은 식), 모델 다섯 가지 모두 같은 값,
16자리 경계값, 파생변수 수식 칸, 잠긴 엑셀, LibreOffice가 있으면 다시 계산해 모델 예측과 비교.
"""
import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest

pytestmark = pytest.mark.report    # 보고서·성능예측표 (오래 걸림: 빼려면 -m "not report", pytest.ini)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_sample import make_class_sample, make_sample  # noqa: E402
from mlkit_helpers import (split_command, FEATURES, HAS_DOCX, HAS_MPL, KIT, TARGET, ok_res, recommended_command, run_fail,  # noqa: E402
                           run_raw, run_tool, stop_msg, write_df)
import ml_export as mx  # noqa: E402
import _ml_runtime as rt  # noqa: E402
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
    run_tool("ml_train.py", "--file", data, "--features", feats, "--targets", "생산효율", "--desc", "효율 예측", "--tag", "base")
    tr = run_tool("ml_train.py", "--file", data, "--features", "반응온도,소비전력,촉매비율,수분함량", "--targets", "생산효율", "--tag", "pruned")
    slots = work / "slots.json"
    slots.write_text(json.dumps({"summary": "슬롯으로 작성한 요약 문장",
                                 "targets": {"생산효율(%)": {"features": "슬롯 피처 소견"}}}, ensure_ascii=False),
                     encoding="utf-8")
    out = work / "report.docx"
    payload = run_tool("ml_report.py", "--model", tr["version_dir"], "--slots", str(slots), "--save", str(out))
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
    tr = run_tool("ml_train.py", "--file", data, "--features", "반응온도,촉매비율,수분함량", "--targets", "품질등급",
                  "--task", "classification", "--positive-class", "불량")
    payload = run_tool("ml_report.py", "--model", tr["version_dir"])
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
    payload = run_fail("ml_report.py", "--model", str(work))
    # 2026-11-13: --model이 ml_predict·ml_export와 같은 해석 (데이터 폴더를 주면 그 안의 학습한 데이터 파일을 [추천])
    assert payload["agent_directive"].startswith("STOP_RETRY") and "[추천]" in payload["message"], payload
    script, argv = recommended_command(payload["message"])
    assert script == "ml_report.py" and argv[0] == "--model" and argv[1].endswith(".csv"), argv
    missing = run_fail("ml_report.py", "--model", str(tmp_path / "nope"))
    assert "버전 폴더" in missing["message"]


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
    payload = run_fail("ml_report.py", "--model", str(vdir))
    assert "2.9.1" in payload["message"] and payload["agent_directive"].startswith("STOP_RETRY")


@needs_report
def test_derived_features_are_shown_with_formulas(tmp_path):
    """파생변수는 2.1에 계산식(× ÷ ² 표기)과 타겟별 기여도·순위로 표시하고, 피처 표에서 '파생'으로 구별한다."""
    data = make_sample(str(tmp_path / "d.csv"))
    tr = run_tool("ml_train.py", "--file", data, "--features", "반응온도,압축압력,촉매비율,수분함량", "--targets", "생산효율",
                  "--eval", "온도x수분 = 반응온도 * 수분함량; 촉매제곱 = 촉매비율 ** 2; 압력비 = 압축압력 / 촉매비율")
    payload = run_tool("ml_report.py", "--model", tr["version_dir"], "--save", str(tmp_path / "r.docx"))
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
def test_report_docx_follows_wordprocessingml_schema(tmp_path):
    """보고서 XML은 WordprocessingML 스키마의 순서와 이름을 따른다 (Word는 w:start/w:end와 순서가 틀린 문서를 거부할 수 있음)."""
    data = tmp_path / "d.csv"
    make_sample(str(data))
    tr = run_tool("ml_train.py", "--file", str(data), "--features", "반응온도,촉매비율,수분함량", "--targets", "생산효율")
    rep = run_tool("ml_report.py", "--model", tr["version_dir"], "--save", str(tmp_path / "r.docx"))
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


def test_report_file_uses_final_version(tmp_path):
    """2.9.83: --file은 최종 버전(마지막 학습)으로 보고서 1개. 다시 학습하지 않고 새 버전을 만들지 않는다."""
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
    tr = run_tool("ml_train.py", "--file", data, "--features", "a,b", "--targets", "y1,y2")
    rep = run_tool("ml_report.py", "--model", data)
    assert rep["version"] == tr["version"] == "v1" and rep["targets"] == ["y1", "y2"]
    assert not os.path.isdir(os.path.join(tr["data_root"], "v2")) and "warnings" not in rep
    for gone in ("best_version", "made_best", "best_versions", "best_targets", "best_reports"):
        assert gone not in rep, gone
    text = "\n".join(p.text for p in docx.Document(rep["report_file"]).paragraphs)
    assert "중복 행 4개" in text                                                 # 2.9.13: not in the report
    # 2026-11-13 스킬 검토: 데이터 파일은 ml_predict·ml_export처럼 --model 한 가지. 옛 --file은 [추천] --model,
    # ml-kit 공통 규칙의 --header-row 등은 받고 쓰지 않음 (warnings)
    fail = run_fail("ml_report.py", "--file", data)
    script, argv = recommended_command(fail["message"])
    assert script == "ml_report.py" and argv[:2] == ["--model", data], argv
    again = run_tool("ml_report.py", "--model", data, "--header-row", "1", "--save", str(tmp_path / "r2.docx"))
    assert any("--header-row" in w and "학습 때 정한" in w for w in again["warnings"]), again


def _three_target_data(tmp_path):
    r = np.random.default_rng(16)
    n = 60
    df = pd.DataFrame({"a": r.uniform(0, 10, n), "b": r.uniform(0, 5, n), "c": r.uniform(0, 1, n)})
    df["A"] = 2 * df.a + r.normal(0, 0.2, n)
    df["B"] = 3 * df.b + r.normal(0, 0.2, n)
    df["S"] = df.a + df.b + r.normal(0, 0.2, n)
    return write_df(tmp_path / "hist.csv", df)


@needs_report
def test_report_file_uses_latest_run_and_warns_targets_left_out(tmp_path):
    """집 PC(2.9.82): v1(타겟 A, B, S) 뒤 S를 빼고 v2를 학습했는데 보고서가 최고 버전을 다시 학습하며 S를 다시 넣었다.
    이제 보고서는 v2(A, B)이고, 이력에만 있는 S는 warnings로 알린다."""
    data = _three_target_data(tmp_path)
    run_tool("ml_train.py", "--file", data, "--features", "a,b,c", "--targets", "A,B,S")
    run_tool("ml_train.py", "--file", data, "--features", "a,b,c", "--targets", "A,B")
    rep = run_tool("ml_report.py", "--model", data)
    assert rep["version"] == "v2" and rep["targets"] == ["A", "B"], rep
    gone = [w for w in rep["warnings"] if "S" in w and "v2" in w and "다시 학습" in w]
    assert len(gone) == 1, rep["warnings"]
    assert "warnings" in rep["agent_directive"]
    heads = _headings(rep["report_file"])[1]
    assert "4.1. A" in heads and "4.2. B" in heads and not any(h.endswith(". S") for h in heads)
    assert not os.path.isdir(tmp_path / "hist_ml_data" / "v3")                 # 다시 학습하지 않음
    # 타겟을 골라 만들면(--targets) 이력 안내는 없음
    only = run_tool("ml_report.py", "--model", data, "--targets", "A", "--save", str(tmp_path / "a.docx"))
    assert only["version"] == "v2" and only["targets"] == ["A"] and "warnings" not in only


@needs_report
def test_report_file_warns_about_an_earlier_version_with_a_higher_score(tmp_path):
    """같은 타겟의 예전 버전 점수가 더 높으면 warnings에 그 버전과 그 버전으로 만드는 --model 명령 (그대로 실행됨).
    다른 타겟 묶음(v1: A, B, S)은 비교하지 않는다."""
    data = _three_target_data(tmp_path)
    run_tool("ml_train.py", "--file", data, "--features", "a,b,c", "--targets", "A,B,S")
    v2 = run_tool("ml_train.py", "--file", data, "--features", "a,b,c", "--targets", "A,B")
    v3 = run_tool("ml_train.py", "--file", data, "--features", "c", "--targets", "A,B")          # 나쁜 입력: 점수가 낮음
    assert v2["mean_val_score"] > v3["mean_val_score"]
    rep = run_tool("ml_report.py", "--model", data, "--save", str(tmp_path / "r3.docx"))
    assert rep["version"] == "v3", rep
    better = [w for w in rep["warnings"] if "점수가 더 높았습니다" in w]
    assert len(better) == 1 and "v2" in better[0] and "v1" not in better[0], rep["warnings"]
    argv = split_command(better[0].split("그 버전으로 만들려면: ", 1)[1])
    assert argv[0] == "python" and argv[argv.index("--model") + 1] == v2["version_dir"], argv
    old = run_tool(os.path.basename(argv[1]), *argv[2:])
    assert old["version"] == "v2" and old["targets"] == ["A", "B"] and "warnings" not in old


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
    monkeypatch.setattr(rt, "_RECOVERED", lst)
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
    with pytest.raises(rt.InputError) as ei:
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
    tr = run_tool("ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET, "--desc", "drm")
    bad = make_damaged_default(tmp_path / "default.docx")
    out = tmp_path / "report.docx"
    proc = subprocess.run(
        [sys.executable, "-c", RUNNER, KIT, bad, str(tmp_path / "없는폴더"), "--model", tr["version_dir"], "--save", str(out)],
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
    tr = run_tool("ml_train.py", "--file", str(path), "--features", "온도,압력,잡음", "--targets", "효율")
    rep = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "시험 효율 예측")
    root = os.path.dirname(tr["version_dir"])
    assert os.path.dirname(rep["report_file"]) == root and os.path.isfile(os.path.join(root, "history.jsonl"))
    assert rep["export_file"] == os.path.join(root, "성능예측표_p_v1.xlsx") and os.path.isfile(rep["export_file"])
    assert rep["export"]["효율"]["inputs"] == ["온도", "압력", "잡음"] and "성능예측표" in rep["preview"]   # 2.9.83: 입력 전부
    assert "export_file" in rep["agent_directive"]
    ex = rep["export"]["효율"]                                                   # 2.9.84: the model as formulas
    assert "그대로" in ex["method"] and ex["check"]["max_diff"] < 1e-9
    import openpyxl
    wb = openpyxl.load_workbook(rep["export_file"])
    assert wb["계산기"]["B2"].value == "시험 효율 예측"
    notes = " ".join(str(c.value) for row in wb["계산기"].iter_rows() for c in row if isinstance(c.value, str))
    assert "학습한 모델과 같은 계산" in notes
    from docx import Document
    text = "\n".join(p.text for p in Document(rep["report_file"]).paragraphs)
    assert "성능예측표 (엑셀 계산기): 성능예측표_p_v1.xlsx" in text and "모델과 같은 값 확인" in text


def test_report_again_reuses_the_checked_calculator_of_the_same_model(tmp_path, with_export):
    """결정 1-7 (C영역 D6-3): 문장만 바꿔 보고서를 다시 만들면 같은 내용의 성능예측표가 시각 붙은 새 이름으로 하나씩 더 생겼다.
    같은 모델(파일 해시)·같은 제목·타겟의 성능예측표가 있고 다시 열어 확인되면 재사용한다 (preview '기존 성능예측표 사용').
    파일이 바뀌었거나 모델이 바뀌었으면 다시 만든다."""
    path = _calc_data(tmp_path)
    tr = run_tool("ml_train.py", "--file", str(path), "--features", "온도,압력", "--targets", "효율")
    first = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "효율 계산")
    root = os.path.dirname(tr["version_dir"])
    assert "기존 성능예측표 사용" not in first["preview"]
    again = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "효율 계산")
    assert again["export_file"] == first["export_file"] and "기존 성능예측표 사용" in again["preview"], again["preview"]
    assert again["report_file"] != first["report_file"]                                  # 보고서는 새로 (시각 붙은 이름)
    assert len([n for n in os.listdir(root) if n.endswith(".xlsx")]) == 1
    assert again["export_verified"] == first["export_verified"] and again["export"] == first["export"]
    assert not [m for m in again.get("recovered", []) if "성능예측표" in m], again.get("recovered")
    from docx import Document
    text = "\n".join(p.text for p in Document(again["report_file"]).paragraphs)
    assert "기존 파일 사용" in text
    other = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "다른 제목")      # 제목이 다르면 새로
    assert other["export_file"] != first["export_file"] and "기존 성능예측표 사용" not in other["preview"]
    wb = openpyxl.load_workbook(first["export_file"])                                     # 사용자가 열어 고치고 저장
    wb.save(first["export_file"])
    third = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "효율 계산")
    assert third["export_file"] not in (first["export_file"], other["export_file"]) and "기존 성능예측표 사용" not in third["preview"]
    fourth = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "효율 계산")
    assert fourth["export_file"] == third["export_file"] and "기존 성능예측표 사용" in fourth["preview"]
    vinfo = mx.load_version(tr["version_dir"], None)                                     # 그때는 엑셀 확인을 껐고 지금은 엑셀이 있음
    made = {"ml_kit": rt.VERSION, "title": "효율 계산", "targets": vinfo["targets"],
            "models": mx.model_hashes(vinfo["dir"], vinfo["metas"], vinfo["targets"])}
    assert mx.reusable(mx.default_export_path(vinfo["dir"]), made)[0] == third["export_file"]
    import unittest.mock
    with unittest.mock.patch.object(mx, "excel_state", lambda: "can"):
        assert mx.reusable(mx.default_export_path(vinfo["dir"]), made) is None
    meta = os.path.join(tr["version_dir"], "model_효율.json")                              # 같은 버전 번호로 모델이 바뀜
    with open(meta, "a", encoding="utf-8") as f:
        f.write("\n")
    fifth = run_tool("ml_report.py", "--model", tr["version_dir"], "--title", "효율 계산")
    assert fifth["export_file"] != third["export_file"] and "기존 성능예측표 사용" not in fifth["preview"]


def test_classification_report_has_no_calculator(tmp_path, with_export):
    path = _calc_data(tmp_path)
    tr = run_tool("ml_train.py", "--file", str(path), "--features", "온도,압력", "--targets", "등급", "--task", "classification")
    rep = run_tool("ml_report.py", "--model", tr["version_dir"])
    assert "export_file" not in rep and "export_note" not in rep


def test_calculator_failure_keeps_the_report(tmp_path, with_export):
    path = _calc_data(tmp_path)
    tr = run_tool("ml_train.py", "--file", str(path), "--features", "온도,압력", "--targets", "효율")
    os.remove(os.path.join(tr["version_dir"], "model_dataset.csv"))          # calculator needs the training snapshot
    rep = run_tool("ml_report.py", "--model", tr["version_dir"])
    assert os.path.isfile(rep["report_file"]) and "export_file" not in rep
    assert rep["export_note"].startswith("성능예측표를 만들지 못했습니다") and "export_note" in rep["agent_directive"]


def test_report_unverified_when_calculator_unverified():
    result = {"preview": "p", "verified": True, "agent_directive": "VERIFIED: 보고서 저장."}
    ex = {"result": {"output_file": "x.xlsx", "targets": {}, "verified": False, "preview": "성능예측표 생성: y",
                     "unverified": ["엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함"]}}
    mr.attach_export(result, ex)
    assert result["verified"] is False and result["export_verified"] is False      # 2.9.14: VERIFIED
    assert result["agent_directive"].startswith("UNVERIFIED:")
    assert result["unverified"] == ["성능예측표: 엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함"]
    ok = {"preview": "p", "verified": True, "agent_directive": "VERIFIED: r."}
    mr.attach_export(ok, {"result": {"output_file": "x.xlsx", "targets": {}, "verified": True, "preview": "성능예측표 생성: y"}})
    assert ok["verified"] is True and ok["agent_directive"].startswith("VERIFIED:")
    assert ok["preview"] == "p | 성능예측표 생성: y (x.xlsx)"                     # 2.9.83: 성능예측표 preview를 이어 붙임


def test_report_passes_overwrite_and_targets(monkeypatch, tmp_path):
    import ml_export
    seen = {}

    def fake(args):
        seen.update(vars(args))
        return {"output_file": "x.xlsx", "targets": {}, "verified": True, "warnings": ["w1"], "preview": "성능예측표 생성: y"}
    monkeypatch.delenv("ML_KIT_SKIP_EXPORT", raising=False)
    monkeypatch.setattr(ml_export, "execute_export", fake)
    ex = mr.make_export(str(tmp_path), "regression", None, ["생산효율"], overwrite=True)
    assert seen["overwrite"] is True and seen["targets"] == "생산효율" and "inputs" not in seen and "keep" not in seen
    result = {"preview": "p", "agent_directive": "VERIFIED: r."}
    mr.attach_export(result, ex)                                          # E7
    assert result["export_warnings"] == ["w1"] and "export_warnings" in result["agent_directive"]


# ---- 성능예측표 (2.9.84): 버전의 모델을 엑셀 수식으로 그대로 옮김 (나무는 잎마다 입력 범위, 선형은 식) ----------------


SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


@pytest.fixture(scope="module")
def export_version(tmp_path_factory):
    """유량 x 농도가 상호작용, 잡음은 무관한 데이터의 버전 (입력 4개)."""
    d = tmp_path_factory.mktemp("exp")
    r = np.random.default_rng(21)
    n = 500
    df = pd.DataFrame({k: r.uniform(0, 1, n) for k in ["온도", "유량", "농도", "잡음"]})
    df["효율"] = 10 * np.sin(3 * df.온도) + 6 * df.유량 * df.농도 + r.normal(0, 0.2, n)
    df["등급"] = np.where(df.효율 > df.효율.median(), "상", "하")
    path = d / "s.csv"
    df.to_csv(path, index=False)
    res = run_tool("ml_train.py", "--file", str(path), "--features", "온도,유량,농도,잡음", "--targets", "효율", "--desc", "시험용 공정 효율 데이터. 두 번째 문장.")
    return d, path, res["version_dir"]


def _recalc(path, tmp):
    out = tmp / "calc"
    out.mkdir(exist_ok=True)
    shutil.copy(path, out / "t.xlsx")
    subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out / "o"), str(out / "t.xlsx")],
                   capture_output=True, timeout=600)
    return openpyxl.load_workbook(out / "o" / "t.xlsx", data_only=True)


def _model_predict(vdir, target, rows):
    meta = json.load(open(os.path.join(vdir, f"model_{target}.json"), encoding="utf-8"))
    cand = {"inputs": meta["input_columns_required"], "evals": meta.get("derived_evals") or [], "schema": meta["feature_schema"],
            "pipe": mx.load_pipe(vdir, meta)}
    return mx.predict_frame(cand, pd.DataFrame(rows, columns=cand["inputs"]))


def test_export_moves_the_version_model_as_it_is(export_version):
    """2.9.84: 다시 학습하지 않고 버전의 모델을 그대로 옮긴다 (입력 전부, 잎 시트, 모델과 같은 값 확인)."""
    _, _, vdir = export_version
    proc, r = run_raw("ml_export.py", "--model", vdir)
    ok_res(r, proc)
    t = r["targets"]["효율"]
    assert t["inputs"] == ["온도", "유량", "농도", "잡음"], t
    champ = json.load(open(os.path.join(vdir, "leaderboard.json"), encoding="utf-8"))["targets"]["효율"]["champion"]
    assert t["model"] == champ["champion_model"] and t["val_r2"] == round(champ["val_r2"], 4), (t, champ)
    assert "그대로" in t["method"] and t["check"]["points"] >= 2000 + 400 and t["check"]["max_diff"] < 1e-9, t
    assert not os.path.isdir(os.path.join(vdir, "export"))                       # 버전 폴더는 바뀌지 않음
    assert r["title"] == "시험용 공정 효율 데이터." and "next_commands" not in r
    assert "preview" in r["agent_directive"] and r["preview"].startswith("성능예측표 생성: 효율: 입력 4개")
    assert "모델과 같은 계산" in r["preview"]
    wb = openpyxl.load_workbook(r["output_file"])
    assert wb.sheetnames == ["계산기", "일괄계산", "사용법", "모델_효율", "정확도", "정보"]
    assert all(wb[n].protection.sheet for n in wb.sheetnames) and wb.security.lockStructure
    assert wb["모델_효율"].sheet_state == "veryHidden"
    assert all(wb[n].sheet_state == "visible" for n in wb.sheetnames if not n.startswith("모델_"))
    calc = wb["계산기"]
    unlocked = [c.coordinate for row in calc.iter_rows() for c in row if not c.protection.locked]
    assert unlocked == ["C6", "C7", "C8", "C9"] and calc["B6"].value == "온도" and calc["B9"].value == "잡음"
    batch = wb["일괄계산"]
    assert not batch["A3"].protection.locked and not batch["D3"].protection.locked and batch["E3"].protection.locked
    if t["leaves"]:
        assert "SUMPRODUCT(" in calc["C12"].value and wb["모델_효율"]["A1"].value.endswith("하한(초과)")
    assert any("같은지 확인" in str(c.value) for row in wb["정확도"].iter_rows() for c in row)
    assert any("엑셀" in u for u in r.get("unverified", [])) or r["verified"]          # Excel check only on Windows


@pytest.mark.skipif(not (sys.platform == "win32" and importlib.util.find_spec("win32com")), reason="Windows + Excel only")
def test_excel_recalculates_like_the_model(export_version, monkeypatch):
    """Windows + 엑셀: 일괄계산에 무작위 조건과 경계값 위·아래 조건을 한 번에 넣고 한 번 다시 계산한 값 = 모델 (그 밖의 시험은 이 확인을 끔)."""
    _, _, vdir = export_version
    monkeypatch.delenv("ML_KIT_SKIP_EXCEL_CHECK", raising=False)
    r = run_tool("ml_export.py", "--model", vdir)
    assert r["verified"] is True and not r.get("unverified"), r


def test_calculator_works_after_the_other_sheets_are_deleted(export_version, tmp_path):
    """최소판은 사용자가 만든다: 일괄계산·사용법·정확도·정보 시트를 지워도 계산기는 숨김 모델 시트만 읽어 그대로 계산 (안내문에 '사용법' 없음)."""
    _, _, vdir = export_version
    r = run_tool("ml_export.py", "--model", vdir)
    wb = openpyxl.load_workbook(r["output_file"])
    note = " ".join(str(c.value) for row in wb["계산기"].iter_rows() for c in row if isinstance(c.value, str))
    assert "사용법" not in note and "같은 계산" in note
    for name in ("일괄계산", "사용법", "정확도", "정보"):
        sheet = f"'{name}'!"
        assert not any(sheet in str(c.value) for row in wb["계산기"].iter_rows() for c in row), name
    if SOFFICE:
        for name in ("일괄계산", "사용법", "정확도", "정보"):
            wb.remove(wb[name])
        wb.save(tmp_path / "min.xlsx")
        out = _recalc(tmp_path / "min.xlsx", tmp_path)
        x = [out["계산기"][f"C{i}"].value for i in (6, 7, 8, 9)]
        assert out["계산기"]["C12"].value == pytest.approx(_model_predict(vdir, "효율", [x])[0], rel=1e-12, abs=1e-12)


def test_status_cells_have_visible_colors(export_version):
    # conditional formats: Excel reads bgColor for a solid fill, so fgColor alone shows no color
    _, _, vdir = export_version
    wb = openpyxl.load_workbook(run_tool("ml_export.py", "--model", vdir)["output_file"])
    for sheet in ("계산기", "일괄계산"):
        rules = [(rule.formula[0], rule.dxf) for cf in wb[sheet].conditional_formatting for rule in cf.rules]
        ok = [d for f, d in rules if f.endswith('="정상"')]
        warn = [d for f, d in rules if "범위 밖" in f and "참고용" in f]
        assert ok and warn and len(ok) == len(warn)
        assert all(d.fill.bgColor.rgb.endswith("DCFCE7") and d.font.color.rgb.endswith("166534") for d in ok)
        assert all(d.fill.bgColor.rgb.endswith("FCE4D6") and d.font.color.rgb.endswith("C2410C") for d in warn)


def test_title_and_removed_options(export_version):
    """--title은 계산기 제목. --inputs(입력 고르기, 다시 학습)는 없앰 (2.9.84: 모델 그대로라 입력 수 상한도 없음)."""
    _, _, vdir = export_version
    i = run_tool("ml_export.py", "--model", vdir, "--title", "짧은 제목")
    assert i["title"] == "짧은 제목"
    for opt in (("--inputs", "온도,유량"), ("--keep", "잡음")):
        proc, payload = run_raw("ml_export.py", "--model", vdir, *opt)
        assert proc.returncode == 2 and payload["status"] == "error", payload
    proc, payload = run_raw("ml_export.py", "-h")
    assert proc.returncode == 0 and payload["command"] == "help"


def test_classification_only_version_is_stopped(export_version):
    _, path, _ = export_version
    cls = run_tool("ml_train.py", "--file", str(path), "--features", "온도,유량", "--targets", "등급", "--task", "classification")
    assert "숫자 타겟" in run_fail("ml_export.py", "--model", cls["version_dir"])["message"]


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_formulas_recalculate_to_the_model(export_version, tmp_path):
    """LibreOffice로 다시 계산한 계산기·일괄계산 값 = 모델 예측 (부동소수점 끝자리까지)."""
    _, _, vdir = export_version
    r = run_tool("ml_export.py", "--model", vdir, "--save", str(tmp_path / "f.xlsx"))
    src = openpyxl.load_workbook(r["output_file"])
    rows = [(0.21, 0.73, 0.4, 0.5), (0.9, 0.1, 0.95, 0.05), (0.33333333333333, 0.5, 0.25, 0.75)]
    for i, xs in enumerate(rows + [(1.05, 0.5, 0.5, 0.5)]):
        for j, v in enumerate(xs):
            src["일괄계산"].cell(row=3 + i, column=1 + j, value=v)
    mod = tmp_path / "m.xlsx"
    src.save(mod)
    out = _recalc(mod, tmp_path)
    calc = out["계산기"]
    x = [calc[f"C{r}"].value for r in (6, 7, 8, 9)]
    want = _model_predict(vdir, "효율", [x] + rows)
    assert calc["C12"].value == pytest.approx(want[0], rel=1e-12, abs=1e-12)          # result row = 5 + inputs + 3
    assert calc["F12"].value == "정상" and calc["F6"].value == "정상"
    b = out["일괄계산"]
    for i in range(3):
        assert b[f"E{3 + i}"].value == pytest.approx(want[1 + i], rel=1e-12, abs=1e-12)
        assert b[f"F{3 + i}"].value < b[f"E{3 + i}"].value < b[f"G{3 + i}"].value
    assert b["H6"].value == "참고용 (학습 범위 밖)" and b["E7"].value in (None, "")


def _pipe(est, cols):
    schema = {"input_columns": cols, "categorical": {}, "numeric": cols, "datetime": {}}
    return mt.build_pipeline(est, schema, "regressor")


@pytest.mark.parametrize("kind", ["ExtraTrees", "RandomForest", "HistGradientBoosting", "Ridge", "ElasticNet"])
def test_every_regression_model_moves_exactly(kind):
    """ml-kit의 회귀 모델 다섯 가지 모두 옮긴 수식 = 모델 예측. 경계값과 바로 위·아래 수, 15자리로 줄인 값 포함."""
    r = np.random.default_rng(7)
    cols = ["a", "b", "c"]
    X = pd.DataFrame(np.round(r.uniform(0, 100, (300, 3)), 2), columns=cols)
    y = np.sin(X.a / 10) * 5 + X.b * X.c / 100 + r.normal(0, 0.1, 300)
    est = mt.get_regression_models(seed=0, n_rows=300)[kind]
    pipe = _pipe(est, cols).fit(X, y)
    form = mx.model_formula(pipe)
    pts = [r.uniform(-10, 110, (500, 3)), X.to_numpy()]
    if form["kind"] == "trees":
        b = np.unique(form["hi"][form["hi"] < mx.OPEN])[:200]
        for f in range(3):
            for edge in (b, np.nextafter(b, np.inf), np.nextafter(b, -np.inf), np.array([float("%.15g" % v) for v in b])):
                P = np.tile(X.median().to_numpy(), (len(edge), 1))
                P[:, f] = edge
                pts.append(P)
    P = pd.DataFrame(np.vstack(pts), columns=cols)
    want = pipe.predict(P)
    got = mx.mirror(form, P.to_numpy(dtype=float))
    np.testing.assert_allclose(got, want, rtol=1e-12, atol=1e-12)


@pytest.mark.parametrize("kind", ["ExtraTrees", "RandomForest", "HistGradientBoosting", "Ridge", "ElasticNet"])
def test_every_regression_model_with_a_text_input_moves_exactly(kind):
    """결정 2-2: 글자 입력(원-핫)이 있는 모델 다섯 가지 모두 옮긴 수식(0/1 칸) = 모델 예측. 종류마다, 학습에 없던 종류(모두 0),
    빈 칸('__missing__'를 학습한 모델은 그 칸), 앞뒤 공백, 숫자 모양의 종류 포함."""
    r = np.random.default_rng(9)
    n = 300
    X = pd.DataFrame({"a": np.round(r.uniform(0, 100, n), 2), "라인": r.choice(["A", "B", "C", "7"], n).astype(object)})
    X.loc[:9, "라인"] = np.nan                                              # 빈 칸도 학습: '__missing__' 종류
    y = np.sin(X.a / 10) * 5 + X.라인.map({"A": 3.0, "B": 0.0, "C": -2.0, "7": 1.0}).fillna(0.5) + r.normal(0, 0.1, n)
    schema = {"input_columns": ["a", "라인"], "numeric": ["a"], "datetime": {},
              "categorical": {"라인": ["7", "A", "B", "C"]}}
    est = mt.get_regression_models(seed=0, n_rows=n)[kind]
    pipe = mt.build_pipeline(est, schema, "regressor").fit(rt.prepare_features(X, schema), y)
    form = mx.model_formula(pipe)
    assert {o["category"] for o in form["onehot"]} == {"7", "A", "B", "C", mx.MISSING}
    cand = {"inputs": ["a", "라인"], "evals": [], "schema": schema, "pipe": pipe, "form": form}
    P = mx.check_points(cand, X)
    assert P.attrs["edge"] == 4 + 3
    extra = pd.DataFrame({"a": [10.0, 50.0, 90.0, 20.0, 30.0], "라인": ["Z", None, "  B ", 7, "-"]})
    P = pd.concat([P, extra], ignore_index=True)
    np.testing.assert_allclose(mx.mirror(form, mx.model_features(cand, P)), mx.predict_frame(cand, P), rtol=1e-12, atol=1e-12)


def test_bounds_are_written_with_every_digit(tmp_path):
    """openpyxl은 숫자를 16자리로 써서 갈림 경계가 움직인다: 잎 시트는 모든 자리(repr)로 써서 다시 읽으면 같은 double."""
    blocks = [np.array([[194.44514502053806, -1e300], [0.1 + 0.2, 1e300]]), np.full((2, 1), np.nan), np.array([[1 / 3]])]
    xml = mx._rows_xml(blocks).decode("ascii")
    assert "194.44514502053806" in xml and "0.30000000000000004" in xml and 'r="C2"' not in xml and 'r="D3"' not in xml
    assert '<c r="D2" t="n"><v>0.3333333333333333</v></c>' in xml


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_derived_features_become_cell_formulas(tmp_path):
    """파생변수(--eval)는 엑셀 수식 칸으로 계산된다 (함수·거듭제곱·나누기 0이면 모델처럼 중앙값)."""
    r = np.random.default_rng(4)
    n = 300
    df = pd.DataFrame({"a": r.uniform(1, 2, n), "b": r.uniform(1, 2, n)})
    df["y"] = 3 * df.a * df.b + np.log(df.a) + r.normal(0, 0.05, n)
    p = tmp_path / "e.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "--file", str(p), "--features", "a,b", "--targets", "y",
                 "--eval", "ab = a * b; q = log(a) - b ** 2 / (a - 1.5)")["version_dir"]
    res = run_tool("ml_export.py", "--model", v, "--save", str(tmp_path / "e.xlsx"))
    assert res["targets"]["y"]["inputs"] == ["a", "b"]                                   # raw inputs only on the sheet
    src = openpyxl.load_workbook(res["output_file"])
    assert [src["계산기"][f"B{r}"].value for r in (6, 7)] == ["a", "b"]
    rows = [(1.2, 1.7), (1.5, 1.3), (1.9, 1.1)]                                          # a = 1.5: division by 0
    for i, xs in enumerate(rows):
        src["일괄계산"].cell(row=3 + i, column=1, value=xs[0])
        src["일괄계산"].cell(row=3 + i, column=2, value=xs[1])
    src.save(tmp_path / "m.xlsx")
    out = _recalc(tmp_path / "m.xlsx", tmp_path)
    want = _model_predict(v, "y", rows)
    for i in range(3):
        assert out["일괄계산"][f"C{3 + i}"].value == pytest.approx(want[i], rel=1e-9, abs=1e-9), i


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_derived_power_and_chained_errors_match_the_model(tmp_path):
    """결정 1-8 (C영역 D5): 파이썬은 음수 밑의 분수 거듭제곱이 빈 값(-> 학습 중앙값), 0 ** 0은 1인데 엑셀(LibreOffice도)은
    (-1)^(1/3) = -1, 엑셀은 0^0 = #NUM!. 앞 파생변수가 빈 값이면 그것을 쓰는 뒤 파생변수도 빈 값(-> 그 중앙값)인데 엑셀은 앞 칸의
    중앙값으로 뒤 칸을 계산했다. 이제 IF로 감싸고, 뒤 파생변수는 채우기 전 값(오류 그대로)을 읽는다."""
    r = np.random.default_rng(5)
    n = 300
    df = pd.DataFrame({"a": r.uniform(-1, 2, n), "b": r.uniform(0.2, 0.8, n)})
    df.loc[:4, ["a", "b"]] = 0.0                                                          # 0 ** 0 행
    pp = np.where(df.a >= 0, np.abs(df.a) ** df.b, 0.0)
    df["y"] = 4 * pp + 0.5 * df.a + r.normal(0, 0.05, n)
    p = tmp_path / "pw.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "--file", str(p), "--features", "b", "--targets", "y",
                 "--eval", "p = a ** b; q = p * 3 + a")["version_dir"]
    res = run_tool("ml_export.py", "--model", v, "--save", str(tmp_path / "pw.xlsx"))
    src = openpyxl.load_workbook(res["output_file"])
    texts = _strings(src, "일괄계산")
    assert "INT(" in texts and "NA()" in texts, texts[:400]                               # 음수 밑 분수 지수 -> 빈 값
    inputs = res["targets"]["y"]["inputs"]
    assert sorted(inputs) == ["a", "b"]
    conds = [{"a": -1.0, "b": 1 / 3}, {"a": 0.0, "b": 0.0}, {"a": 0.0, "b": -1.0}, {"a": 1.5, "b": 0.5}, {"a": -0.5, "b": 0.7}]
    rows = [[c[k] for k in inputs] for c in conds]
    for i, xs in enumerate(rows):
        for j, x in enumerate(xs):
            src["일괄계산"].cell(row=3 + i, column=1 + j, value=x)
    src.save(tmp_path / "m.xlsx")
    out = _recalc(tmp_path / "m.xlsx", tmp_path)
    want = _model_predict(v, "y", rows)
    pred_col = 3
    for i in range(len(rows)):
        got = out["일괄계산"].cell(row=3 + i, column=pred_col).value
        assert got == pytest.approx(want[i], rel=1e-9, abs=1e-9), (rows[i], got, want[i])
    assert "파생변수" in res["preview"] and "엑셀" in res["preview"], res["preview"]       # 파이썬 확인 밖이라는 정직한 문구


def test_excel_check_adds_zero_and_negative_conditions_for_derived_features():
    """결정 1-8: 파생변수 수식은 엑셀에서만 확인되므로 엑셀 확인 조건에 0과 음수 입력을 더한다 (파생변수가 없으면 더하지 않음)."""
    data = pd.DataFrame({"a": [1.0, 2.0, 3.0], "b": [4.0, 5.0, 6.0]})
    p = {"cand": {"inputs": ["a", "b"], "evals": ["p = a ** b"]}, "data": data}
    pts = mx.derived_points(p)
    assert len(pts) == 6 and (pts.iloc[0] == 0).all() and (pts.iloc[1] == -1.5).all()
    assert pts.iloc[3].tolist() == [-3.0, 5.0] and pts.iloc[4].tolist() == [2.0, 0.0]
    assert mx.derived_points({"cand": {"inputs": ["a"], "evals": []}, "data": data}).empty
    assert "엑셀이 있는 PC에서만" in mx.derived_note(2, False) and "엑셀에서 다시 계산해 확인" in mx.derived_note(2, True)


def test_several_outputs_share_the_inputs(tmp_path):
    r = np.random.default_rng(31)
    n = 300
    df = pd.DataFrame({"온도": r.uniform(0, 1, n), "압력": r.uniform(0, 1, n), "유량": r.uniform(0, 1, n)})
    df["효율"] = 10 * np.sin(3 * df.온도) + r.normal(0, 0.1, n)
    df["전력"] = 5 * df.압력 * df.유량 + 2 * df.유량 + r.normal(0, 0.05, n)
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "--file", str(p), "--features", "온도,압력,유량", "--targets", "효율,전력")["version_dir"]
    res = run_tool("ml_export.py", "--model", v, "--save", str(tmp_path / "m.xlsx"))
    assert list(res["targets"]) == ["효율", "전력"]                                       # --targets order
    wb = openpyxl.load_workbook(res["output_file"])
    calc = wb["계산기"]
    assert [calc[f"B{r}"].value for r in (6, 7, 8)] == ["온도", "압력", "유량"]
    assert [calc[f"B{r}"].value for r in (11, 12)] == ["효율", "전력"]
    assert "COUNT($C$6,$C$7,$C$8)=3" in calc["C11"].value
    assert wb.sheetnames[3:5] == ["모델_효율", "모델_전력"]
    one = run_tool("ml_export.py", "--model", v, "--targets", "전력", "--save", str(tmp_path / "one.xlsx"))
    assert list(one["targets"]) == ["전력"]


def test_workbook_has_no_full_paths(export_version):
    """<데이터>_ml_data 앞의 폴더 경로(사용자 이름, 회사 폴더)는 개인 정보다: 통합 문서에 넣지 않는다."""
    d, _, vdir = export_version
    res = run_tool("ml_export.py", "--model", vdir)
    wb = openpyxl.load_workbook(res["output_file"])
    parent = os.path.dirname(os.path.dirname(vdir))
    texts = [str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert not any(parent in t for t in texts)
    assert any(t == os.path.join("s_ml_data", "v1") for t in texts)


def _strings(wb, sheet):
    return " ".join(str(c.value) for row in wb[sheet].iter_rows() for c in row if c.value is not None)


def test_calculator_range_ignores_excluded_rows_and_keeps_a_constant_input(tmp_path):
    """학습에서 뺀 행(이상치)은 학습 범위에 넣지 않는다. 값이 하나뿐인 입력도 모델 입력이라 계산기에 남고 warnings로 알린다."""
    r = np.random.default_rng(5)
    n = 300
    df = pd.DataFrame({"반응온도": r.uniform(150, 200, n), "압력": r.uniform(1, 3, n), "pH": 7.0})
    df["효율"] = 0.4 * df.반응온도 + 6 * df.압력 + r.normal(0, 0.3, n)
    df.loc[10, "반응온도"] = 2000                                # typo, excluded as an automatic outlier
    p = tmp_path / "c.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "--file", str(p), "--features", "반응온도,압력,pH", "--targets", "효율")["version_dir"]
    res = run_tool("ml_export.py", "--model", v, "--save", str(tmp_path / "c.xlsx"))
    t = res["targets"]["효율"]
    assert t["inputs"] == ["반응온도", "압력", "pH"] and any("pH" in w and "하나뿐" in w for w in res["warnings"])
    wb = openpyxl.load_workbook(res["output_file"])
    text = _strings(wb, "계산기") + _strings(wb, "일괄계산") + _strings(wb, "정보")
    assert "2000" not in text and "1999" not in text                       # range / out-of-range checks use 150~200
    assert wb["계산기"]["C8"].value == 7.0


def test_export_and_map_overwrite(export_version, tmp_path):
    d, _, vdir = export_version
    save = str(d / "ow.xlsx")
    first = run_tool("ml_export.py", "--model", vdir, "--save", save)
    second = run_tool("ml_export.py", "--model", vdir, "--save", save)
    assert second["output_file"] != first["output_file"] and second.get("recovered")
    third = run_tool("ml_export.py", "--model", vdir, "--save", save, "--overwrite")
    assert third["output_file"] == first["output_file"] and not third.get("recovered")
    p = tmp_path / "g.csv"
    pd.DataFrame({"A": [0.0, 10, 20], "C": [1.0, 2, 3]}).to_csv(p, index=False)
    m1 = run_tool("ml_map.py", "--file", str(p), "--input", "A=5", "--save", "out.xlsx")
    m2 = run_tool("ml_map.py", "--file", str(p), "--input", "A=5", "--save", "out.xlsx", "--overwrite")
    assert m1["output_file"] == m2["output_file"]


def test_too_many_leaves_is_stopped(export_version, monkeypatch):
    _, _, vdir = export_version
    monkeypatch.setattr(mx, "MAX_LEAVES", 10)
    form = mx.model_formula(mx.load_pipe(vdir, json.load(open(os.path.join(vdir, "model_효율.json"), encoding="utf-8"))))
    if form["kind"] != "trees":
        pytest.skip("champion is a linear model")
    with pytest.raises(rt.InputError, match="잎"):
        mx.execute_export(argparse.Namespace(model=vdir, targets=None, title=None, save=None, overwrite=False))


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
    v = run_tool("ml_train.py", "--file", str(p), "--features", "a,b", "--targets", "y")["version_dir"]
    meta = json.load(open(os.path.join(v, "model_y.json"), encoding="utf-8"))
    assert meta["prediction_interval"]["residual_upper"] != 0                     # was rounded to 6 decimals
    res = run_tool("ml_export.py", "--model", v, "--save", str(tmp_path / "tiny.xlsx"))
    t = res["targets"]["y"]
    assert t["rmse"] > 0 and t["check"]["max_rel_pct"] < 1e-6
    calc = openpyxl.load_workbook(res["output_file"])["계산기"]
    cells = [c for row in calc.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=IF(COUNT(")]
    assert any(c.number_format == "0.000E+00" for c in cells)
    assert any('TEXT(' in str(c.value) and '"0.000E+00"' in str(c.value) for row in calc.iter_rows() for c in row)


@pytest.mark.parametrize("k", [2, 6])
def test_uncovered_threshold_follows_the_data(k):
    """'실측 없는 영역' 기준은 학습 행끼리의 간격에 맞춘다: 학습 행 근처는 거의 표시하지 않고(입력이 많아 데이터가 성겨도),
    학습 범위 안이지만 실측이 없는 구석은 표시한다."""
    from scipy.spatial import cKDTree
    r = np.random.default_rng(k)
    cols = [f"x{i}" for i in range(k)]
    data = pd.DataFrame(r.uniform(0, 1, (400, k)), columns=cols)
    data.iloc[:, 0] *= 0.5                                   # no measured data for x0 > 0.5 ...
    data.iloc[:3, :] = 0.0
    data.iloc[:3, 0] = 1.0                                   # ... except one corner, so the range of x0 is 0~1
    T = data.to_numpy()
    th = mx.uncovered_distance(T, data, cols)
    near = T[3:] + r.normal(0, 0.01, T[3:].shape)
    assert (cKDTree(T).query(near)[0] > th).mean() < 0.1
    q = 0.3 + 0.7 * r.uniform(0, 1, (500, k))               # away from the corner rows (x0 = 1, others 0)
    q[:, 0] = 0.9 + 0.1 * r.uniform(0, 1, 500)
    assert (cKDTree(T).query(q)[0] > th).mean() > 0.9


def test_sheet_names_unique_without_case():
    used = {"계산기"}
    a = mx._sheet_name("모델_", "Temp", used)
    b = mx._sheet_name("모델_", "TEMP", used)
    assert a.lower() != b.lower()


def test_unfinished_passes_through_the_reports_optional_workbook(monkeypatch, with_export):
    """보고서는 성능예측표 실패를 note로 넘기지만(보고서는 성공), 시간 제한 멈춤은 삼키지 않고 CONTINUE로 끝낸다."""
    import ml_export
    import _ml_runtime as rt
    import ml_report

    def stop(args):
        raise rt.Unfinished("후보 2개 학습 저장됨")
    monkeypatch.setattr(ml_export, "execute_export", stop)
    with pytest.raises(rt.Unfinished):
        ml_report.make_export("v", "regression", None)


def test_time_limit_setting(monkeypatch):
    import _ml_runtime as rt
    for raw, want in (("", rt.DEFAULT_TIME_LIMIT), ("abc", rt.DEFAULT_TIME_LIMIT), ("0", float("inf")), ("250", 250.0)):
        monkeypatch.setenv("ML_KIT_TIME_LIMIT", raw)
        assert rt.time_limit() == want
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    rt.check_time(10 ** 6, "x")                                   # 제한 없음: 멈추지 않음
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "31")
    monkeypatch.setattr(rt, "RUN_START", rt.time.monotonic() - 5)
    with pytest.raises(rt.Unfinished):
        rt.check_time(0, "x")


def test_excel_edge_point_is_beyond_excel_comparison_precision():
    """회사 PC 2026-11-10 F7: 엑셀은 15자리쯤까지 같은 두 수를 같다고 비교 -> 경계값 바로 다음 double은 엑셀에서 경계값 쪽 잎.
    엑셀 확인의 '경계 바로 위' 조건은 엑셀이 구분하는 가장 작은 걸음 위 (모델도 그 값은 오른쪽 잎)."""
    import ml_export
    for v in (76.5671, 180.49999237060547, -3.25, 0.0, 1e-7, 12345678.9):
        up = ml_export.above_for_excel(v)
        assert up > v and float("%.15g" % up) != float("%.15g" % v), (v, up)
        assert up - v <= max(abs(v), 1.0) * 1e-12, (v, up)


def test_excel_check_lost_excel_connection_is_reported_not_raised(monkeypatch):
    """회사 PC 2026-11-10 beta F1: 엑셀 연결이 끊겨 xl.Quit 속성 찾기가 AttributeError -> finally 밖으로 올라가 성능예측표 전체 실패.
    닫기 실패는 삼키고 확인 실패만 알림."""
    import types
    import importlib.machinery
    import ml_export

    class LostExcel:
        Visible = DisplayAlerts = False

        @property
        def Workbooks(self):
            raise RuntimeError("RPC 서버를 사용할 수 없습니다")

        def __getattr__(self, name):
            raise AttributeError(f"Excel.Application.{name}")
    client = types.ModuleType("win32com.client")
    client.DispatchEx = lambda prog: LostExcel()
    pkg = types.ModuleType("win32com")
    pkg.__spec__ = importlib.machinery.ModuleSpec("win32com", None)
    pkg.client = client
    monkeypatch.setitem(sys.modules, "win32com", pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.delenv("ML_KIT_SKIP_EXCEL_CHECK", raising=False)
    got = ml_export.excel_check("없는파일.xlsx", [])
    assert got.startswith("엑셀로 계산 확인 실패: RuntimeError"), got


def test_excel_check_passes_within_one_percent():
    """사용자 결정 2026-11-10: 성능예측표 엑셀 확인은 99% 일치(차이 1% 이내)면 통과, 다르면 warnings로 알림, 1%를 넘으면 실패.
    결정 1-5 (C영역 C3): 1%의 기준은 확인한 조건들의 예측 범위(최대 - 최소). 예측값 크기 기준이면 1000 근처 타겟은
    범위의 절반만큼 틀려도 통과하고, 0 근처 타겟은 아주 작은 차이로 실패했다."""
    import ml_export
    w = ["일괄계산 3행", "일괄계산 4행", "계산기"]
    assert ml_export.compare_excel("y", [1.0, 2.0, 1.0], [1.0, 2.0, 1.0], w) == (None, None)
    problem, note = ml_export.compare_excel("y", [76.5689, 2.0, 1.0], [76.5671, 2.0, 1.0], w)     # 회사 PC F7 (범위 75.6의 0.0024%)
    assert problem is None and "일괄계산 3행" in note and "허용 1%" in note and "예측 범위" in note, note
    problem, note = ml_export.compare_excel("y", [1.02, 2.0, 1.0], [1.0, 2.0, 1.0], w)      # 범위 1의 2%
    assert note is None and "2.00% 다름" in problem and "예측 범위" in problem, problem
    problem, _ = ml_export.compare_excel("y", ["#VALUE!", 2.0, 1.0], [1.0, 2.0, 1.0], w)
    assert "숫자가 아님" in problem
    # 1000 근처에서 990~1010만 움직이는 타겟: 값의 1%(10)가 아니라 범위 20의 1%(0.2)까지만
    problem, _ = ml_export.compare_excel("y", [1005.0, 990.0, 1010.0], [1000.0, 990.0, 1010.0], w)
    assert problem and "25.00% 다름" in problem, problem
    # 0 근처 타겟: 범위 2의 0.05%는 통과 (값 기준이면 0.001 / 0.002 = 50%로 실패)
    problem, note = ml_export.compare_excel("y", [0.003, -1.0, 1.0], [0.002, -1.0, 1.0], w)
    assert problem is None and note, (problem, note)
    # 모든 조건의 예측이 같음 (범위 0): 예측값 크기의 1%로
    problem, _ = ml_export.compare_excel("y", [5.0, 5.0, 5.0], [5.0, 5.0, 5.0], w)
    assert problem is None
    problem, _ = ml_export.compare_excel("y", [5.2, 5.0, 5.0], [5.0, 5.0, 5.0], w)
    assert problem and "4.00% 다름" in problem, problem


# ---- 2.9.88 검토 C: 보고서·성능예측표 결함 (재현 -> 수정) ------------------------------------------------------------


@pytest.fixture(scope="module")
def three_targets(tmp_path_factory):
    """타겟 A, B, S 회귀 버전 (머리글 행을 직접 줌: 보고서의 표 구조 문구)."""
    d = tmp_path_factory.mktemp("rev_c")
    data = _three_target_data(d)
    tr = run_tool("ml_train.py", "--file", data, "--features", "a,b,c", "--targets", "A,B,S", "--header-row", "1")
    return d, data, tr


def _doc_text(path):
    return _all_text(docx.Document(path))


@needs_report
def test_unreadable_slots_is_stop_retry_and_unused_slots_are_warned(three_targets):
    """A1: --slots를 읽지 못하면 자동 소견으로 만들고 VERIFIED였음 -> STOP_RETRY (보고서 만들지 않음).
    모르는 항목, 보고서에 없는 타겟의 문장은 warnings (slots_used에 없는 문장은 넣지 않았다고 알림)."""
    d, data, tr = three_targets
    reports = lambda: sorted(p.name for p in d.glob("hist_ml_data/*.docx"))         # noqa: E731
    before = reports()
    msg = run_fail("ml_report.py", "--model", tr["version_dir"], "--slots", str(d / "없음.json"))["message"]
    assert "--slots 파일이 없습니다" in msg and '"summary"' in msg, msg
    bad = d / "bad.json"
    bad.write_text('{"summary": "요약",}', encoding="utf-8")
    msg = run_fail("ml_report.py", "--model", tr["version_dir"], "--slots", str(bad))["message"]
    assert "JSON" in msg and "끝 쉼표" in msg, msg
    assert reports() == before
    odd = d / "odd.json"
    odd.write_text(json.dumps({"summary": "내가 정한 요약", "nokey": "x",
                               "targets": {"A": {"features": "A 피처 소견"}, "Z": {"features": "z"}}}, ensure_ascii=False),
                   encoding="utf-8")
    rep = run_tool("ml_report.py", "--model", tr["version_dir"], "--slots", str(odd), "--save", str(d / "odd.docx"))
    assert set(rep["slots_used"]) == {"summary", "targets.A.features"}, rep["slots_used"]
    assert any("'nokey'" in w for w in rep["warnings"]) and any("targets.Z" in w for w in rep["warnings"]), rep["warnings"]
    assert "slots_used" in rep["agent_directive"] and "warnings" in rep["agent_directive"]
    assert "내가 정한 요약" in _doc_text(rep["report_file"])


@needs_report
def test_report_of_some_targets_averages_only_those(three_targets):
    """A2: --targets A 보고서의 '평균 검증 R²' 카드가 버전 전체(A, B, S) 평균이었고, 학습 사용 행에 빠진 타겟도 적었다.
    D4: 표 구조는 내부 키 이름(header 1행)이 아니라 '머리글 1행', 상관이 작은 입력은 '상관이 큰 입력'에 적지 않음."""
    d, data, tr = three_targets
    rep = run_tool("ml_report.py", "--model", tr["version_dir"], "--targets", "A", "--save", str(d / "a_only.docx"))
    lb = json.load(open(os.path.join(tr["version_dir"], "leaderboard.json"), encoding="utf-8"))
    want = lb["targets"]["A"]["champion"]["val_score"]
    assert abs(want - tr["mean_val_score"]) > 1e-3
    doc = docx.Document(rep["report_file"])
    card = " ".join(c.text for row in doc.tables[0].rows for c in row.cells)
    assert f"{want:.4f}" in card and f"{tr['mean_val_score']:.4f}" not in card, card
    text = "\n".join(p.text for p in doc.paragraphs)
    used = text.split("학습 사용 ", 1)[1].split(".")[0]
    assert used.startswith("A: ") and "B:" not in used and "S:" not in used, used
    assert "표 구조: 머리글 1행" in text and "header" not in text
    full = _all_text(doc)
    import re
    assert re.search(r"A: a\([+-]\d\.\d\d\)(;|\n|$)", full), full[full.find("상관이 큰"):][:200]


def test_pretty_formula_keeps_slash_and_star_in_column_names():
    """A3: 열 이름 안의 '/', '*'(L/min, m/s, x*y)를 연산자로 바꿔 '유량(L ÷ min)'으로 썼다."""
    assert mr.pretty_formula("유량비 = 유량(L/min) / 압축압력(bar)", ["유량(L/min)", "압축압력(bar)"]) == "유량(L/min) ÷ 압축압력(bar)"
    assert mr.pretty_formula("곱 = 유속 (m/s) * x*y", ["유속 (m/s)", "x*y"]) == "유속 (m/s) × x*y"
    t = {"name": "y", "champ": {"feature_importances": []},
         "meta": {"input_columns_required": ["유량(L/min)", "압력"], "derived_evals": ["비 = 유량(L/min) / 압력", "비2 = 비 ** 2"]}}
    rows, names = mr.derived_feature_rows([t])
    assert [r[1] for r in rows] == ["비 = 유량(L/min) ÷ 압력", "비2 = 비²"] and names == ["비", "비2"]


@needs_report
def test_report_save_must_be_a_word_file_and_relative_paths_go_to_the_data_folder(three_targets):
    """A5: --save out.pdf에 Word 파일을 저장하고 VERIFIED였다 -> STOP_RETRY + [추천] .docx 전체 명령. 확장자가 없으면 .docx.
    B3: 상대 경로는 성능예측표처럼 한 기준 (없는 폴더는 만듦). 결정 1-6: 그 기준은 원본 데이터 파일 옆, [추천] 명령도 같은 곳."""
    d, data, tr = three_targets
    proc, res = run_raw("ml_report.py", "--model", tr["version_dir"], "--targets", "A", "--save", str(d / "out.pdf"))
    msg = stop_msg(res, proc)
    script, argv = recommended_command(msg)
    assert script == "ml_report.py" and argv[argv.index("--save") + 1] == str(d / "out.docx"), argv
    assert not (d / "out.pdf").exists()
    assert run_tool(script, *argv)["report_file"] == str(d / "out.docx")
    assert run_tool("ml_report.py", "--model", tr["version_dir"], "--targets", "A",
                    "--save", str(d / "noext"))["report_file"] == str(d / "noext.docx")
    rel = run_tool("ml_report.py", "--model", tr["version_dir"], "--targets", "A", "--save", os.path.join("새폴더", "rel.docx"))
    # 결정 1-6: 상대 경로는 원본 데이터 파일 옆 (성능예측표·데이터셋과 같음, 2.9.88은 <데이터>_ml_data)
    assert rel["report_file"] == os.path.join(os.path.dirname(str(data)), "새폴더", "rel.docx"), rel["report_file"]
    proc, res = run_raw("ml_report.py", "--model-dir", tr["version_dir"], "--targets", "A", "--save", "추천.docx")
    script, argv = recommended_command(res["message"])
    assert proc.returncode == 2 and argv[argv.index("--save") + 1] == os.path.join(os.path.dirname(str(data)), "추천.docx"), argv


def _fake_tree(**champ):
    base = {"champion_model": "Ridge", "val_r2": 0.95, "mae": 0.1, "overfit_gap": 0.0, "residual_skew": 0.0,
            "top_features": [], "feature_importances": [], "weak_features": []}
    base.update(champ)
    return {"task": "regression", "name": "강도", "oof": pd.DataFrame({"actual": [1.0, 2.0, 3.0]}), "champ": base,
            "meta": {"feature_schema": {"numeric": ["온도", "압력"], "categorical": {"작업조": ["주간", "야간"]},
                                        "datetime": {"일자": ["month", "day", "dayofweek"]}}}}


def test_findings_do_not_contradict_each_other():
    """A6: 기여도 0%인 범주가 '상위 기여 피처에 포함'과 '기여도 1% 이하'에 함께 나왔고, 간극 0.10~0.15는 카드로는 정상인데
    소견은 과적합이었다. D4: 날짜 분해 단위를 영어 내부 이름(dayofweek)으로, 판정 이유를 소수 3자리로 썼다."""
    ranks = [{"feature": "온도", "pct": 95.2}, {"feature": "압력", "pct": 4.8}, {"feature": "작업조", "pct": 0.0}]
    t = _fake_tree(top_features=ranks, feature_importances=ranks, weak_features=["작업조"], overfit_gap=0.12)
    feats = mr.text_features(t)
    assert "상위 기여 피처에 포함" not in feats and "기여도 1% 이하" in feats, feats
    assert "월, 일, 요일" in feats and "dayofweek" not in feats, feats
    assert "과적합 위험이 낮습니다" in mr.text_performance(t)
    assert "과적합 경향" in mr.text_performance(_fake_tree(overfit_gap=0.16))
    useful = [{"feature": "작업조", "pct": 40.0}, {"feature": "온도", "pct": 60.0}]
    assert "상위 기여 피처에 포함" in mr.text_features(_fake_tree(top_features=useful, feature_importances=useful))
    assert mr.verdict_of(0.99961, 0.01, "R²")[1] == "검증 R² 0.9996 >= 0.9"


@needs_report
def test_report_rows_count_without_the_unit_row(tmp_path):
    """A4 (보고서 쪽): 단위 행을 알려 주지 않고 학습한 데이터의 보고서가 '학습 데이터 61행'(단위 행 포함)이라고 썼다.
    카드는 학습에 쓴 행 수, 본문은 '읽은 행'과 '빈 칸 또는 숫자가 아닌 칸'으로 쓴다 (원인은 학습 원장의 프로파일)."""
    r = np.random.default_rng(4)
    rows = [["온도", "압력", "수율"], ["°C", "bar", "%"]]
    for _ in range(60):
        t, p = round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2)
        rows.append([t, p, round(0.3 * t + 5 * p + float(r.normal(0, 0.3)), 2)])
    from mlkit_helpers import write_rows
    f = write_rows(str(tmp_path / "unit.csv"), rows, [])
    tr = run_tool("ml_train.py", "--file", f, "--features", "온도,압력", "--targets", "수율")
    rep = run_tool("ml_report.py", "--model", tr["version_dir"])
    doc = docx.Document(rep["report_file"])
    card = " ".join(c.text for row in doc.tables[0].rows for c in row.cells)
    assert "60행" in card and "61행" not in card, card
    text = "\n".join(p.text for p in doc.paragraphs)
    # 1-9 (C영역 A4 뿌리, 2026-11-10): 학습 프로파일이 단위 행을 행 수·결측 수에서 뺀다 (예전 61행)
    assert "읽은 행 60행" in text and "61행" not in text and "학습 데이터 기준값으로 대치" not in text, text


@needs_report
def test_report_is_saved_before_the_calculator(three_targets, monkeypatch, tmp_path):
    """B1: 보고서 명령이 성능예측표를 먼저 저장하고 Word 보고서를 만들어, 그 사이 시간 제한에 닿으면 보고서 없이 ok +
    '다시 실행하지 마세요'였다. 이제 보고서를 먼저 저장하고, 성능예측표를 만드는 동안 멈추면 결과에 보고서 경로와
    성능예측표를 따로 만드는 전체 명령이 들어간다. 끝나면 성능예측표 줄을 채워 같은 파일에 다시 저장."""
    import ml_export
    d, data, tr = three_targets
    monkeypatch.delenv("ML_KIT_SKIP_EXPORT", raising=False)
    out = tmp_path / "b1.docx"
    seen = {}

    def fake(args):
        seen["on_disk"] = out.is_file()
        seen["payload"] = rt._watchdog_payload()[0]
        return {"output_file": str(tmp_path / "x.xlsx"), "title": "t", "verified": True, "preview": "성능예측표 생성: y",
                "targets": {t: {"inputs": ["a", "b", "c"], "model": "Ridge", "val_r2": 0.9, "method": "Ridge 선형식 그대로",
                                "check": {"points": 2060, "max_diff": 0.0, "max_rel_pct": 0.0}} for t in ("A", "B", "S")}}
    monkeypatch.setattr(ml_export, "execute_export", fake)
    rt.RUN_COMMITTED.clear()
    try:
        res = mr.execute_report(argparse.Namespace(model=tr["version_dir"], file=None, targets=None, save=str(out),
                                                   overwrite=False, slots=None, title="제목"))
    finally:
        rt.RUN_COMMITTED.clear()
        rt.WATCHDOG_HOLD.clear()
    assert seen["on_disk"] is True, seen
    p = seen["payload"]
    assert p["status"] == "ok" and p["report_file"] == str(out) and "ml_export.py" in p["agent_directive"], p
    argv = split_command(p["agent_directive"].split("원하면 ", 1)[1].split(" 로 따로", 1)[0])
    assert os.path.basename(argv[1]) == "ml_export.py" and argv[argv.index("--model") + 1] == os.path.abspath(tr["version_dir"])
    assert argv[argv.index("--title") + 1] == "제목"
    assert res["report_file"] == str(out) and res["export_file"] == str(tmp_path / "x.xlsx")
    text = "\n".join(q.text for q in docx.Document(str(out)).paragraphs)
    assert "성능예측표 (엑셀 계산기): x.xlsx" in text and "모델과 같은 값 확인" in text and "만드는 도중" not in text
    assert text.count("입력 3개") == 1                                       # B7: 같은 입력은 한 번


@needs_report
def test_other_task_target_gets_the_report_command_of_its_version(tmp_path):
    """B5: 분류 타겟(v1) 뒤 회귀 타겟(v2)을 학습하면 --file 보고서가 '그 타겟을 --targets에 넣어 다시 학습'이라고 했다
    (분류와 회귀는 함께 학습할 수 없음). 이제 그 타겟 버전의 보고서를 만드는 전체 명령을 준다."""
    path = _calc_data(tmp_path)
    v1 = run_tool("ml_train.py", "--file", str(path), "--features", "온도,압력", "--targets", "등급", "--task", "classification")
    run_tool("ml_train.py", "--file", str(path), "--features", "온도,압력", "--targets", "효율")
    rep = run_tool("ml_report.py", "--model", str(path))
    w = [x for x in rep["warnings"] if "등급" in x]
    assert len(w) == 1 and "다시 학습" not in w[0], rep["warnings"]
    argv = split_command(w[0].split("만들려면: ", 1)[1])
    assert os.path.basename(argv[1]) == "ml_report.py" and argv[argv.index("--model") + 1] == v1["version_dir"]
    old = run_tool("ml_report.py", *argv[2:])
    assert old["version"] == "v1" and old["task"] == "classification"


def test_calculator_formula_over_the_excel_limit_is_stopped(export_version, monkeypatch, tmp_path):
    """B2: 입력이 많으면 수식이 엑셀 한도(8,192자, 함수 인자 255개)를 넘는데 검사 없이 저장했다 (입력 140개: 9,033자).
    한도를 낮춰 재현: 저장 전에 STOP_RETRY(입력 수, 가장 긴 수식, 한도) + 파일 없음."""
    _, _, vdir = export_version
    monkeypatch.setattr(mx, "FORMULA_MAX", 300)
    save = tmp_path / "limit.xlsx"
    with pytest.raises(rt.InputError, match="한도"):
        mx.execute_export(argparse.Namespace(model=vdir, targets=None, title=None, save=str(save), overwrite=False))
    assert not save.exists()


def test_export_save_folder_extension_and_missing_version(export_version, tmp_path):
    """B3: 없는 폴더로 --save는 [Errno 2]만, 'my.xls'는 'my.xls.xlsx'로 저장했다 -> 폴더를 만들고, 다른 확장자는 [추천] .xlsx.
    B8: 없는 버전 폴더는 '버전 폴더가 없습니다'뿐이었다 -> 있는 버전의 [추천] 전체 명령 (ml_report와 같음)."""
    d, _, vdir = export_version
    deep = tmp_path / "없는폴더" / "아래" / "calc.xlsx"
    assert run_tool("ml_export.py", "--model", vdir, "--save", str(deep))["output_file"] == str(deep) and deep.is_file()
    proc, res = run_raw("ml_export.py", "--model", vdir, "--save", str(tmp_path / "my.xls"))
    script, argv = recommended_command(stop_msg(res, proc))
    assert script == "ml_export.py" and argv[argv.index("--save") + 1] == str(tmp_path / "my.xlsx"), argv
    assert not list(tmp_path.glob("my.xls*"))
    # 결정 1-6: 상대 --save는 원본 데이터 파일 옆 (보고서·데이터셋과 같음), 다른 확장자의 [추천]도 같은 곳
    rel = run_tool("ml_export.py", "--model", vdir, "--save", os.path.join("표", "rel"))
    assert rel["output_file"] == str(d / "표" / "rel.xlsx"), rel["output_file"]
    proc, res = run_raw("ml_export.py", "--model", vdir, "--save", "rel.xls")
    script, argv = recommended_command(stop_msg(res, proc))
    assert argv[argv.index("--save") + 1] == str(d / "rel.xlsx"), argv
    root = tmp_path / "s_ml_data"                     # 학습한 버전이 v1 하나인 데이터 폴더 (다른 테스트가 버전을 더 만들어도)
    shutil.copytree(vdir, root / "v1")
    proc, res = run_raw("ml_export.py", "--model", str(root / "v9"))
    script, argv = recommended_command(stop_msg(res, proc))
    assert script == "ml_export.py" and argv[argv.index("--model") + 1] == str(root / "v1"), argv


def test_failed_export_does_not_report_a_new_name(export_version, monkeypatch):
    """B6: 같은 이름 성능예측표가 있을 때 저장 전에 실패해도 recovered에 '이미 있음 -> ..._<시각>.xlsx로 저장'이 남았다
    (그 파일은 없음). 새 이름 알림은 저장한 뒤에만."""
    _, _, vdir = export_version
    first = run_tool("ml_export.py", "--model", vdir)
    assert os.path.isfile(first["output_file"])

    def boom(*a, **k):
        raise RuntimeError("저장 전 실패")
    monkeypatch.setattr(mx, "write_workbook", boom)
    rt._RECOVERED.clear()
    try:
        with pytest.raises(RuntimeError):
            mx.execute_export(argparse.Namespace(model=vdir, targets=None, title=None, save=None, overwrite=False))
        assert not [m for m in rt._RECOVERED if "이미 있음" in m], rt._RECOVERED
    finally:
        rt._RECOVERED.clear()


def test_brief_lists_few_input_names_once():
    """B7: 입력 140개 모델의 preview가 2,041자(이름 140개)였고, 타겟마다 같은 목록을 반복했다.
    D2: 괄호 안만 다른 입력이 '온도, 온도'로 나왔다."""
    names = [f"측정값_{i:03d}(mm)" for i in range(140)]
    brief = mx.names_brief(names)
    assert brief.endswith("외 132개") and brief.count("측정값_") == 8 and "(mm)" not in brief
    assert mx.names_brief(["온도(입구)", "온도(출구)", "압력(bar)"]) == "온도(입구), 온도(출구), 압력"
    v = {"inputs": names, "model": "Ridge", "val_r2": 0.9, "check": {"points": 2000}}
    text = mx.export_brief({"A": v, "B": dict(v)})
    assert text.count("측정값_000") == 1 and len(text) < 400, text


def _train_csv(tmp_path, name, df, features, target):
    p = tmp_path / name
    df.to_csv(p, index=False, encoding="utf-8-sig")
    return p, run_tool("ml_train.py", "--file", str(p), "--features", features, "--targets", target)


def _cat_data(n=150, seed=8):
    r = np.random.default_rng(seed)
    df = pd.DataFrame({"온도": r.uniform(150, 200, n), "압력": r.uniform(1, 5, n), "라인": r.choice(["A", "B", "C"], n)})
    df["강도"] = 0.5 * df.온도 + 3 * df.압력 + df.라인.map({"A": 5.0, "B": 0.0, "C": -3.0}) + r.normal(0, 0.1, n)
    return df


@needs_report
def test_text_input_report_makes_the_calculator_and_date_input_says_why_not(tmp_path, with_export):
    """C1 / 결정 2-2: 글자(범주) 입력 모델도 성능예측표를 만든다 (드롭다운 + 숨김 0/1 칸). 날짜 입력은 여전히 만들지 않고 까닭을
    보고서와 export_note에 쓴다 (학습 권고의 '함께 만들어집니다' 문장은 뺌)."""
    df = _cat_data()
    _, tr = _train_csv(tmp_path, "cat.csv", df, "온도,압력,라인", "강도")
    rep = run_tool("ml_report.py", "--model", tr["version_dir"])
    assert os.path.isfile(rep["export_file"]) and "export_note" not in rep, rep
    assert rep["export"]["강도"]["inputs"] == ["온도", "압력", "라인"]
    dd = df.copy()
    dd["일자"] = pd.date_range("2024-01-01", periods=len(dd), freq="D").strftime("%Y-%m-%d")
    _, tr2 = _train_csv(tmp_path, "date.csv", dd, "온도,압력,일자", "강도")
    rep2 = run_tool("ml_report.py", "--model", tr2["version_dir"])
    assert "export_file" not in rep2 and rep2["export_note"].startswith("성능예측표(엑셀 계산기)는 만들지 않았습니다"), rep2
    assert "일자" in rep2["export_note"] and "날짜" in rep2["export_note"] and "Error" not in rep2["export_note"]
    text = _doc_text(rep2["report_file"])
    assert "함께 만들어집니다" not in text and "성능예측표 (엑셀 계산기): 만들지 않음" in text


def test_text_input_calculator_has_a_dropdown_and_matches_the_model(tmp_path):
    """결정 2-2: 글자(범주) 입력은 계산기에 드롭다운(학습 종류 목록) + 숨김 0/1 칸으로 옮긴다. 파이썬 사본 확인은 종류마다,
    학습에 없던 종류, 빈 칸 조건을 포함하고, 학습에 없던 종류·빈 칸은 모델처럼 계산하되 판단은 '참고용'."""
    df = _cat_data()
    _, tr = _train_csv(tmp_path, "cat2.csv", df, "온도,압력,라인", "강도")
    v = tr["version_dir"]
    res = run_tool("ml_export.py", "--model", v)
    t = res["targets"]["강도"]
    assert t["inputs"] == ["온도", "압력", "라인"] and t["check"]["max_diff"] < 1e-9, t
    assert all("엑셀 계산 확인을 껐음" in u for u in res["unverified"]), res["unverified"]     # 구조·목록·드롭다운 확인은 통과
    meta = json.load(open(os.path.join(v, "model_강도.json"), encoding="utf-8"))
    cand = {"inputs": t["inputs"], "evals": [], "schema": meta["feature_schema"], "pipe": mx.load_pipe(v, meta)}
    cand["form"] = mx.model_formula(cand["pipe"])
    P = mx.check_points(cand, pd.read_csv(os.path.join(v, "model_dataset.csv")))
    labels = set(P["라인"].map(lambda x: "" if not isinstance(x, str) else x))
    assert {"A", "B", "C", ""} <= labels and any(x not in ("A", "B", "C", "") for x in labels), labels
    wb = openpyxl.load_workbook(res["output_file"])
    assert "목록" in wb.sheetnames and wb["목록"].sheet_state == "veryHidden"
    calc = wb["계산기"]
    row = next(i for i in range(6, 9) if calc.cell(row=i, column=2).value == "라인")
    cell = calc.cell(row=row, column=3)
    assert cell.value in ("A", "B", "C") and not cell.protection.locked and cell.number_format == "@"
    lists = [dv for dv in calc.data_validations.dataValidation if dv.type == "list"]
    assert lists and cell.coordinate in str(lists[0].sqref) and "목록" in lists[0].formula1, lists
    assert "종류 3개" in str(calc.cell(row=row, column=5).value)
    batch_lists = [dv for dv in wb["일괄계산"].data_validations.dataValidation if dv.type == "list"]
    assert batch_lists, "일괄계산의 글자 열에도 드롭다운"
    assert "학습에 없던 종류" in _strings(wb, "계산기") and "학습에 없던 종류" in _strings(wb, "일괄계산")
    if SOFFICE is None:
        return
    inputs = t["inputs"]
    conds = [{"온도": 170.0, "압력": 2.0, "라인": "A"}, {"온도": 180.0, "압력": 3.0, "라인": "B"},
             {"온도": 190.0, "압력": 4.0, "라인": "C"}, {"온도": 175.0, "압력": 2.5, "라인": "Z"},
             {"온도": 185.0, "압력": 1.5, "라인": None}, {"온도": 160.0, "압력": 4.5, "라인": " C "}]
    for i, c in enumerate(conds):
        for j, k in enumerate(inputs):
            wb["일괄계산"].cell(row=3 + i, column=1 + j, value=c[k])
    wb.save(tmp_path / "m.xlsx")
    out = _recalc(tmp_path / "m.xlsx", tmp_path)
    want = _model_predict(v, "강도", [[c[k] for k in inputs] for c in conds])
    for i in range(len(conds)):
        got = out["일괄계산"].cell(row=3 + i, column=4).value
        assert got == pytest.approx(want[i], rel=1e-9, abs=1e-9), (conds[i], got, want[i])
    judge = [out["일괄계산"].cell(row=3 + i, column=7).value for i in range(len(conds))]
    assert judge[0] == "정상" and judge[3].startswith("참고용") and judge[4].startswith("참고용") and judge[5] == "정상", judge


def test_text_input_with_many_categories_respects_the_formula_limit(tmp_path, monkeypatch):
    """결정 2-2: 종류가 많으면 수식이 길어진다 -> 수식 길이 한도 검사(C영역 B2)를 그대로 지킨다 (저장 전 STOP_RETRY)."""
    df = _cat_data(n=120, seed=3)
    _, tr = _train_csv(tmp_path, "cat3.csv", df, "온도,압력,라인", "강도")
    monkeypatch.setattr(mx, "FORMULA_MAX", 120)
    save = tmp_path / "lim.xlsx"
    with pytest.raises(rt.InputError, match="한도"):
        mx.execute_export(argparse.Namespace(model=tr["version_dir"], targets=None, title=None, save=str(save), overwrite=False))
    assert not save.exists()


def test_low_score_calculator_says_reference_only(tmp_path):
    """C2: 검증 R²가 0.7 미만(보고서 '부적합')인 모델도 경고 없이 '모델과 같은 계산'으로만 만들었다 -> warnings,
    계산기 제목 아래 주의 문장, 판단 칸 '참고용 (모델 정확도 낮음)'."""
    r = np.random.default_rng(2)
    n = 200
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n)})
    df["y"] = r.normal(0, 1, n)
    _, tr = _train_csv(tmp_path, "noise.csv", df, "a,b", "y")
    res = run_tool("ml_export.py", "--model", tr["version_dir"])
    assert res["targets"]["y"]["val_r2"] < 0.7
    assert any("참고용" in w and "검증 R²" in w for w in res["warnings"]) and "참고용" in res["preview"], res
    wb = openpyxl.load_workbook(res["output_file"])
    assert "정확도가 낮아" in wb["계산기"]["B4"].value
    assert "참고용 (모델 정확도 낮음)" in _strings(wb, "계산기") and "참고용 (모델 정확도 낮음)" in _strings(wb, "일괄계산")


def test_tiny_and_same_name_inputs_in_the_calculator(tmp_path):
    """D1: 1e-7 크기 입력의 계산기 기본값이 0, 학습 범위가 '0.000000 ~ 0.000000'이었다.
    D2: 괄호 안만 다른 입력이 계산기에 '온도', '온도'로 나왔다. D6: 사용법에 빈 입력 칸 안내."""
    r = np.random.default_rng(3)
    n = 150
    df = pd.DataFrame({"두께(m)": r.uniform(1.2e-7, 3.4e-7, n), "온도(입구)": r.uniform(10, 20, n), "온도(출구)": r.uniform(30, 40, n)})
    df["y"] = 1e7 * df["두께(m)"] + df["온도(입구)"] - 0.5 * df["온도(출구)"] + r.normal(0, 0.05, n)
    _, tr = _train_csv(tmp_path, "tiny_in.csv", df, "두께(m),온도(입구),온도(출구)", "y")
    res = run_tool("ml_export.py", "--model", tr["version_dir"])
    calc = openpyxl.load_workbook(res["output_file"])["계산기"]
    rows = {calc.cell(row=i, column=2).value: i for i in range(6, 9)}
    assert set(rows) == {"두께", "온도(입구)", "온도(출구)"}, rows
    i = rows["두께"]
    assert 1.2e-7 <= calc.cell(row=i, column=3).value <= 3.4e-7 and "E-07" in calc.cell(row=i, column=5).value
    assert "온도(입구)" in res["preview"] and "온도(출구)" in res["preview"]
    assert "비어 있으면" in _strings(openpyxl.load_workbook(res["output_file"]), "사용법")
