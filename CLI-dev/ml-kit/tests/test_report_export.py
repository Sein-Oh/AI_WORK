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
from mlkit_helpers import (split_command, FEATURES, HAS_DOCX, HAS_MPL, KIT, TARGET, ok_res, run_fail, run_raw, run_tool,  # noqa: E402
                           write_df)
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
    assert payload["agent_directive"].startswith("STOP_RETRY") and "leaderboard.json" in payload["message"]
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
    rep = run_tool("ml_report.py", "--file", data)
    assert rep["version"] == tr["version"] == "v1" and rep["targets"] == ["y1", "y2"]
    assert not os.path.isdir(os.path.join(tr["data_root"], "v2")) and "warnings" not in rep
    for gone in ("best_version", "made_best", "best_versions", "best_targets", "best_reports"):
        assert gone not in rep, gone
    text = "\n".join(p.text for p in docx.Document(rep["report_file"]).paragraphs)
    assert "중복 행 4개" in text                                                 # 2.9.13: not in the report


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
    rep = run_tool("ml_report.py", "--file", data)
    assert rep["version"] == "v2" and rep["targets"] == ["A", "B"], rep
    gone = [w for w in rep["warnings"] if "S" in w and "v2" in w and "다시 학습" in w]
    assert len(gone) == 1, rep["warnings"]
    assert "warnings" in rep["agent_directive"]
    heads = _headings(rep["report_file"])[1]
    assert "4.1. A" in heads and "4.2. B" in heads and not any(h.endswith(". S") for h in heads)
    assert not os.path.isdir(tmp_path / "hist_ml_data" / "v3")                 # 다시 학습하지 않음
    # 타겟을 골라 만들면(--targets) 이력 안내는 없음
    only = run_tool("ml_report.py", "--file", data, "--targets", "A", "--save", str(tmp_path / "a.docx"))
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
    rep = run_tool("ml_report.py", "--file", data, "--save", str(tmp_path / "r3.docx"))
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
