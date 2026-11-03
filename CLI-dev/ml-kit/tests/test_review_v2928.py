# -*- coding: utf-8 -*-
"""
2.9.28 (17차 점검) regressions.

1. 2.9.27 confirmed a loose totals row ('총 합계', '월합계', '일합계') only when at least half of its number cells were
   block sums. A totals row that averages the process columns (온도/압력/속도) and sums only 생산량 was not warned.
   A cell equal to the block mean is now neutral: confirmed with 1+ sum cell and (sums + means) * 2 >= filled.
2. -i '압력(MPa)=0.3' for a model input '압력(bar)' got "[추천] '압력(bar)'": the retry kept the MPa value under
   bar and returned VERIFIED. A name with another unit gets no [추천]: STOP_RETRY asking to convert the value to the
   model's unit (also ml_map -i / --find and ml_predict --find -i).
3. 2.9.27 refused -i '온도=180,온도(°C)=182' while parsing, without the model: a model with both '온도' and
   '온도(°C)' features could not be used. The check now runs where names are matched to the model inputs (ml_map:
   data columns): refused only when the unit-less name is not itself a feature / column.
4. JSON -i: a repeated key ({"온도":180,"온도":190}) and an alias pair ('온도' and '온도(°C)' for the feature
   '온도(°C)') silently kept one value. Both stop (STOP_RETRY) like 'k=v'.
"""
import json
import os
import random
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tools"))
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _summary(tmp_path, name, rows, merges=()):
    return mp.load_data_file(_write(str(tmp_path / name), rows, list(merges))).attrs["summary_rows"]


def _avg(rows, c):
    return round(sum(r[c] for r in rows) / len(rows), 2)


def _sum(rows, c):
    return round(sum(r[c] for r in rows), 2)


def _process(seed=7, n=30):
    r = random.Random(seed)
    return [[f"10/{i + 1}", round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2), round(r.uniform(40, 50), 1),
             r.randint(280, 320)] for i in range(n)]


H = ["일자", "온도", "압력", "속도", "생산량"]


# ---------------------------------------------------------------------------------------------------------------
# 1. averaged process columns + summed quantity
# ---------------------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("word", ["총 합계", "합계", "전체합계", "Grand Total", "월합계"])
def test_totals_row_with_averaged_columns_warned(tmp_path, word):
    d = _process()
    rows = [H] + d + [[word, _avg(d, 1), _avg(d, 2), _avg(d, 3), _sum(d, 4)]]
    got = _summary(tmp_path, "a.xlsx", rows)
    assert [r[0] for r in got] == [32] and got[0][2] == word, got


def test_totals_row_ratio_column_warned(tmp_path):
    r = random.Random(3)
    d = []
    for i in range(30):
        p, b = r.randint(280, 320), r.randint(0, 9)
        d.append([f"10/{i + 1}", round(r.uniform(175, 185), 1), p, b, round(b / p * 100, 2)])
    P, B = sum(x[2] for x in d), sum(x[3] for x in d)
    rows = [["일자", "온도", "생산량", "불량수", "불량률"]] + d + [["총 합계", _avg(d, 1), P, B, round(B / P * 100, 2)]]
    assert [x[0] for x in _summary(tmp_path, "r.xlsx", rows)] == [32]


@pytest.mark.parametrize("word,last", [("일합계", "월합계"), ("일계", "월합계"), ("주간계", "월간계"), ("합계", "총 합계")])
def test_daily_report_averaged_columns_all_warned(tmp_path, word, last):
    # merged-date daily report, 3 shifts: 온도/압력/속도 averaged, 생산량 summed, per day and in the last row
    r = random.Random(5)
    rows, merges, every = [["일자", "교대", "온도", "압력", "속도", "생산량"]], [], []
    for day in range(1, 11):
        r0, blk = len(rows) + 1, []
        for j, s in enumerate(["A조", "B조", "C조"]):
            x = [f"10/{day}" if j == 0 else None, s, round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2),
                 round(r.uniform(40, 50), 1), r.randint(280, 320)]
            rows.append(x)
            blk.append(x)
            every.append(x)
        merges.append(f"A{r0}:A{r0 + 2}")
        rows.append([word, None, _avg(blk, 2), _avg(blk, 3), _avg(blk, 4), _sum(blk, 5)])
    rows.append([last, None, _avg(every, 2), _avg(every, 3), _avg(every, 4), _sum(every, 5)])
    got = _summary(tmp_path, "d.xlsx", rows, merges)
    assert len(got) == 11 and got[-1][0] == 42 and got[-1][2] == last, got


def test_category_rows_still_quiet(tmp_path):
    # categories hold numbers like the data rows (2.9.27 c2 probes): a value near a column mean alone confirms nothing
    r = random.Random(2)
    labels = ["A계"] * 3 + ["B조"] * 17
    r.shuffle(labels)
    rows = [["구분", "온도편차", "생산량"]] + [[lab, round(r.uniform(-2, 2), 2), r.randint(280, 320)] for lab in labels]
    assert _summary(tmp_path, "c.xlsx", rows) == []
    r = random.Random(6)
    labels = ["1계"] * 14 + ["2계"] * 6
    r.shuffle(labels)
    rows = [["계통", "생산량", "원료투입", "전력"]]
    for lab in labels:
        s = 3 if lab == "2계" else 1
        rows.append([lab, r.randint(280, 320) * s, r.randint(300, 340) * s, r.randint(50, 60) * s])
    assert _summary(tmp_path, "g.xlsx", rows) == []
    # a category row whose temperature equals the mean and quantity is no sum: quiet
    d = _process(seed=9, n=20)
    rows = [H] + d[:10] + [["A계", _avg(d[:10], 1), _avg(d[:10], 2), _avg(d[:10], 3), 301]] + d[10:]
    assert _summary(tmp_path, "m.xlsx", rows) == []


def test_mean_neutral_needs_a_sum(tmp_path):
    a = mp._sums_match
    import numpy as np
    sums = np.array([[300.0, 30.0, 3000.0]])
    count = np.array([[10.0, 10.0, 10.0]])
    # 30.0 (mean of col 0), 3.0 (mean of col 1), 3000 (sum) -> confirmed; all means, no sum -> not
    assert a(np.array([[30.0, 3.0, 3000.0]]), sums, count, sums)[0]
    assert not a(np.array([[30.0, 3.0, 300.0]]), sums, count, sums)[0]
    # one sum among two unrelated cells -> not (1 * 2 < 3)
    assert not a(np.array([[17.0, 9.0, 3000.0]]), sums, count, sums)[0]


# ---------------------------------------------------------------------------------------------------------------
# 2-4. -i names: other units, alias pairs, repeated JSON keys
# ---------------------------------------------------------------------------------------------------------------

def _train(tmp_path, name, cols):
    r = random.Random(1)
    rows = [cols + ["수율"]]
    for _ in range(60):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2)
        v = {"온도": t, "온도(°C)": round(t + r.uniform(-0.5, 0.5), 1), "압력(bar)": p}
        rows.append([v[c] for c in cols] + [round(0.5 * t + 10 * p + r.uniform(-1, 1), 2)])
    f = _write(str(tmp_path / name), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", ",".join(cols), "-y", "수율")
    assert res["status"] == "ok", res
    return os.path.join(str(tmp_path), os.path.splitext(name)[0] + "_ml_data", "v1"), f


@pytest.fixture(scope="module")
def models(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2928")
    both, both_f = _train(tmp, "al.xlsx", ["온도", "온도(°C)", "압력(bar)"])
    unit, unit_f = _train(tmp, "al2.xlsx", ["온도(°C)", "압력(bar)"])
    return {"both": both, "unit": unit, "unit_f": unit_f, "tmp": tmp}


def _pred(model, inline):
    return run_raw("ml_predict.py", "-m", model, "-i", inline)


def test_other_unit_no_rename_suggestion(models):
    proc, res = _pred(models["unit"], "온도(°C)=180,압력(MPa)=0.3")
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "[추천]" not in res["message"] and "단위가 다릅니다" in res["message"], res["message"]
    assert "bar 단위로 환산" in res["message"], res["message"]
    # a unit-less name is still matched; a spacing typo of the same unit still gets [추천]
    assert _pred(models["unit"], "온도=180,압력=3")[1]["status"] == "ok"
    proc, res = _pred(models["unit"], "온도(°C)=180,압력(bar )=3")
    assert "[추천] '압력(bar)'" in res["message"], res["message"]


def test_other_unit_ml_map_and_find(models):
    proc, res = run_raw("ml_map.py", "-f", models["unit_f"], "-i", "온도(°C)=180,압력(MPa)=0.3", "-y", "수율")
    assert proc.returncode == 1 and "[추천]" not in res["message"] and "단위가 다릅니다" in res["message"], res
    proc, res = run_raw("ml_predict.py", "-m", models["unit"], "--find", "수율=100", "--vary", "온도(°C)",
                        "-i", "압력(MPa)=0.3")
    assert proc.returncode == 1 and "[추천]" not in res["message"] and "단위가 다릅니다" in res["message"], res


def test_alias_pair_with_both_features_accepted(models):
    assert mp.parse_inline_input("온도=180,온도(°C)=182").shape == (1, 2)
    for inline in ("온도=180,온도(°C)=182,압력(bar)=3", '{"온도":180,"온도(°C)":182,"압력(bar)":3}'):
        proc, res = _pred(models["both"], inline)
        assert res["status"] == "ok", res


@pytest.mark.parametrize("inline", ["온도=180,온도(°C)=182,압력(bar)=3", '{"온도":180,"온도(°C)":182,"압력(bar)":3}',
                                    '[{"온도":180,"온도(°C)":182,"압력(bar)":3}]'])
def test_alias_pair_one_feature_refused(models, inline):
    proc, res = _pred(models["unit"], inline)
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "같은 열 '온도(°C)'" in res["message"] and "[후보] 온도=180, 온도(°C)=182" in res["message"], res


def test_alias_pair_ml_map_and_find(models):
    proc, res = run_raw("ml_map.py", "-f", models["unit_f"], "-i", '{"온도":180,"온도(°C)":182}', "-y", "수율")
    assert proc.returncode == 1 and "[후보] 온도=180, 온도(°C)=182" in res["message"], res
    proc, res = run_raw("ml_predict.py", "-m", models["unit"], "--find", "수율=100", "--vary", "압력(bar)",
                        "-i", "온도=180,온도(°C)=182")
    assert proc.returncode == 1 and "[후보] 온도=180, 온도(°C)=182" in res["message"], res


def test_json_repeated_key_refused(models):
    with pytest.raises(mp.InputError) as e:
        mp.parse_inline_input('{"온도(°C)":180,"온도(°C)":190}')
    assert "두 번" in str(e.value) and "[후보] 온도(°C)=180, 온도(°C)=190" in str(e.value)
    p = os.path.join(str(models["tmp"]), "in.json")
    with open(p, "w", encoding="utf-8") as f:
        f.write('[{"온도(°C)":180,"압력(bar)":3,"압력(bar)":4}]')
    proc, res = _pred(models["unit"], p)
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "[후보] 압력(bar)=3, 압력(bar)=4" in res["message"], res
    # a plain JSON row still predicts
    assert _pred(models["unit"], json.dumps({"온도(°C)": 180, "압력(bar)": 3}, ensure_ascii=False))[1]["status"] == "ok"


def test_other_unit_rules():
    assert mp.other_unit("압력(MPa)", "압력(bar)") and mp.other_unit("온도(°F)", "온도(°C)")
    assert mp.other_unit("무게(mg)", "무게(Mg)")                    # case is a different unit
    for a, b in (("온도(C)", "온도(°C)"), ("온도(℃)", "온도(°C)"), ("압력 [bar]", "압력( bar )"), ("온도", "온도(°C)")):
        assert not mp.other_unit(a, b), (a, b)
