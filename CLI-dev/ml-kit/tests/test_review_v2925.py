# -*- coding: utf-8 -*-
"""
2.9.25 (14차 점검) regressions.

1. The 2.9.24 unparsed-number warning read every loaded row: the unit row right below the header (auto-detected,
   target empty, never trained) gave one false warning per numeric column ('°C', 'bar' ...) on the SKILL sample,
   and the 5-warning cap hid real ones. Only rows a model learns from are checked now (target present, not marked
   in '제외', not the unit row), one warning per column for all targets, further columns counted.
2. Only columns the model reads as numbers (feature schema numeric, --eval sources): no warning for a category code
   column ('101', 'A1'), a 40% decimal-comma column that became a category, or a --dtypes datetime column.
3. --eval source columns are checked ('175,6' in T1 of dT=T1-T2).
4. Totals wording of daily reports ('일계', '총 합계', 'Total:', '소계(A)', '평균값' ...) was matched exactly and
   trained without a warning. Ordinary words ending with '계' ('설계', '1단계', '온도계') stay quiet.
5. A selected feature of decimal-comma texts ('0,149') trained silently as a high-cardinality category: refused
   (STOP_RETRY) from half of the filled cells, a warning below that.
"""
import os
import random
import sys
from datetime import datetime

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tools"))
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _number_warnings(res):
    return [w for w in res.get("warnings", []) if "숫자로 읽지 못해" in w or "소수점 쉼표" in w]


def test_skill_sample_unit_row_no_false_warning(tmp_path):
    import make_beta_samples
    make_beta_samples.main(str(tmp_path))
    data = tmp_path / "공정_다중출력.xlsx"      # header row 3, unit row 4 (auto-detected: no layout options)
    _, res = run_raw("ml_train.py", "-f", str(data), "-x", "반응온도,압축압력,촉매비율,수분함량,유량,분진",
                     "-y", "생산효율,불순물")
    assert res["status"] == "ok", res
    assert _number_warnings(res) == [], res.get("warnings")


def _train_rows(tmp_path, name, rows, *args):
    f = _write(str(tmp_path / name), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, *args)
    return res


def test_category_code_column_not_warned(tmp_path):
    r = random.Random(3)
    rows = [["금형", "온도", "수율"]]
    for i in range(60):
        t = 150 + i % 30
        m = r.choice(["101", "102", "103", "104", "105", "106", "107", "A1", "A2", "B1"])
        rows.append([m, t, 0.5 * t + (5 if m.startswith("A") else 0) + r.uniform(-1, 1)])
    res = _train_rows(tmp_path, "u1.xlsx", rows, "-x", "금형,온도", "-y", "수율")
    assert res["status"] == "ok" and _number_warnings(res) == [], res


def test_excluded_and_target_missing_rows_not_warned(tmp_path):
    r = random.Random(3)
    rows = [["온도", "압력", "수율", "제외"]]
    for i in range(40):
        t, p, y, ex = 150 + i, round(r.uniform(2, 4), 2), 0.5 * (150 + i) + r.uniform(-1, 1), None
        if i in (5, 15):
            t, ex = "설비정지", "정지"
        if i == 25:
            t, y = "점검", None
        rows.append([t, p, y, ex])
    res = _train_rows(tmp_path, "u2.xlsx", rows, "-x", "온도,압력", "-y", "수율")
    assert res["status"] == "ok" and _number_warnings(res) == [], res


def test_forced_datetime_column_not_warned(tmp_path):
    r = random.Random(3)
    rows = [["일자", "온도", "수율"]]
    for i in range(40):
        d = datetime(2026, 1, 1 + i % 28)
        rows.append([(d.strftime("%Y-%m-%d") if i in (4, 8) else int(d.strftime("%Y%m%d"))), 150 + i,
                     0.5 * (150 + i) + r.uniform(-1, 1)])
    res = _train_rows(tmp_path, "u5.xlsx", rows, "-x", "일자,온도", "-y", "수율", "--dtypes", "일자:datetime")
    assert res["status"] == "ok" and _number_warnings(res) == [], res


def test_eval_source_column_warned(tmp_path):
    r = random.Random(3)
    rows = [["T1", "T2", "수율"]]
    for i in range(40):
        a, b = 150 + i, 100 + r.uniform(0, 10)
        rows.append([("175,6" if i in (3, 9, 20) else a), b, 0.5 * (a - b) + r.uniform(-1, 1)])
    res = _train_rows(tmp_path, "u3.xlsx", rows, "-x", "T2", "--eval", "dT=T1-T2", "-y", "수율")
    assert res["status"] == "ok", res
    assert any("'T1'의 3칸" in w and "175,6" in w for w in res["warnings"]), res["warnings"]


def test_one_warning_per_column_and_further_columns_counted(tmp_path):
    r = random.Random(5)
    cols = [f"x{k}" for k in range(7)]
    rows = [cols + ["y1", "y2"]]
    for i in range(40):
        vals = [round(r.uniform(0, 10), 2) for _ in cols]
        y = sum(vals)
        for k in range(7):
            if i == 3 + k:
                vals[k] = f"{k},5"
        rows.append(vals + [y, 2 * y])
    res = _train_rows(tmp_path, "many.xlsx", rows, "-x", ",".join(cols), "-y", "y1,y2")
    assert res["status"] == "ok", res
    named = [w for w in res["warnings"] if w.startswith("숫자 열")]
    assert len(named) == 5 and len(set(named)) == 5, res["warnings"]
    assert all("의 1칸" in w for w in named), named
    assert any("2개 더" in w and "'x5'" in w and "'x6'" in w for w in res["warnings"]), res["warnings"]


@pytest.mark.parametrize("word", ["일계", "월계", "누계", "총 합계", "전체합계", "합계:", "Total:", "소계(A)", "주간계",
                                  "일합계", "Grand Total"])
def test_daily_report_totals_wording_warned(tmp_path, word):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 8):
        a, b = [150 + d, 300 + d], [160 + d, 310 + d]
        rows += [[f"10/{d}", "주간", *a], [None, "야간", *b], [word, None, a[0] + b[0], a[1] + b[1]]]
    got = mp.load_data_file(_write(str(tmp_path / "w.xlsx"), rows, [])).attrs["summary_rows"]
    assert len(got) == 7, (word, got)


@pytest.mark.parametrize("word", ["평균값", "AVG.", "AVG"])
def test_average_wording_warned(tmp_path, word):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 8):
        rows += [[f"10/{d}", "주간", 150 + d, 300 + d], [None, "야간", 160 + d, 310 + d]]
    rows.append([word, None, 160, 310])
    got = mp.load_data_file(_write(str(tmp_path / "a.xlsx"), rows, [])).attrs["summary_rows"]
    assert [r[2] for r in got] == [word], got


def test_words_ending_with_gye_stay_categories(tmp_path):
    r = random.Random(2)
    rows = [["부서", "공정", "설비", "온도", "수율"]]
    for i in range(40):
        rows.append([r.choice(["설계", "계측", "생산", "품질"]), r.choice(["1단계", "2단계", "3단계"]),
                     r.choice(["온도계", "압력계", "유량계", "하계", "동계"]), 150 + i, 70 + i / 2])
    got = mp.load_data_file(_write(str(tmp_path / "c.xlsx"), rows, [])).attrs["summary_rows"]
    assert got == [], got


def test_word_kind_units():
    k = mp._word_kind
    assert [k(w) for w in ("일계", "년계", "Sub Total", "sum", "평균(A조)", "Mean:", "최대")] == \
        ["total", "total", "total", "total", "avg", "avg", "agg"]
    assert [k(w) for w in ("설계", "계측", "summary", "통계", "2단계", "A조")] == [""] * 6


def _textdc_rows(share):
    r = random.Random(4)
    rows = [["온도", "압력", "두께", "수율"]]
    for i in range(40):
        t, th = 150 + i, round(r.uniform(0.1, 0.3), 3)
        cell = f"{th:.3f}".replace(".", ",") if (i % 10) < share * 10 else th
        rows.append([t, round(r.uniform(2, 4), 2), cell, 0.5 * t + 100 * th + r.uniform(-1, 1)])
    return rows


def test_decimal_comma_text_feature_refused(tmp_path):
    res = _train_rows(tmp_path, "c.xlsx", _textdc_rows(1.0), "-x", "온도,압력,두께", "-y", "수율")
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY"), res
    assert "'두께'" in res["message"] and "0," in res["message"] and "[추천]" not in res["message"], res["message"]


def test_decimal_comma_text_minority_warned(tmp_path):
    res = _train_rows(tmp_path, "c4.xlsx", _textdc_rows(0.4), "-x", "온도,압력,두께", "-y", "수율")
    assert res["status"] == "ok", res
    assert any("피처 '두께'" in w and "소수점 쉼표" in w for w in res["warnings"]), res["warnings"]


def test_decimal_comma_text_forced_category_allowed(tmp_path):
    res = _train_rows(tmp_path, "c5.xlsx", _textdc_rows(1.0), "-x", "온도,압력,두께", "-y", "수율",
                      "--dtypes", "두께:str")
    assert res["status"] == "ok", res
