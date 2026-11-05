# -*- coding: utf-8 -*-
"""예측(ml_predict.py): -i 값 읽기, 예측 결과와 결과 파일, --find 역산, 실측 비교(map_check).

열 이름 맞추기(단위, 별칭, 오타 [후보])는 test_names.py.
"""
import json
import os
import shutil
import sys
import time

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import split_command, KIT, run_fail, run_raw, run_tool, stop_msg, write_df, write_rows  # noqa: E402
import ml_predict as mp  # noqa: E402


# ---- -i 값 읽기: 글자 그대로, 천 단위, 같은 이름 두 번, 키 오타 ------------------------------------


def test_inline_pairs_stay_text():
    df = mp.parse_inline_input("코드=007,a=1.5")
    assert df.loc[0, "코드"] == "007" and df.loc[0, "a"] == "1.5"


def test_inline_category_code_with_leading_zero(coded):
    res = run_tool("ml_predict.py", "-m", coded["v"], "-i", "a=5,b=2,코드=007")
    assert res["unseen_category_warning"] is False and res["missing_input_warning"] is False
    js = coded["work"] / "one.json"
    js.write_text(json.dumps({"a": 5, "b": 2, "코드": "007"}), encoding="utf-8")
    same = run_tool("ml_predict.py", "-m", coded["v"], "-i", str(js))
    assert res["predictions_summary"]["y"]["predictions"] == same["predictions_summary"]["y"]["predictions"]


def test_inline_thousands_separator():
    df = mp.parse_inline_input("온도=3,압력=500,000,라인=A")
    assert df.iloc[0].to_dict() == {"온도": "3", "압력": "500,000", "라인": "A"}
    assert float(mp.coerce_numeric(df["압력"]).iloc[0]) == 500000.0
    with pytest.raises(mp.InputError):
        mp.parse_inline_input("500,온도=3")


@pytest.fixture(scope="module")
def small(tmp_path_factory):
    work = tmp_path_factory.mktemp("comma")
    r = np.random.default_rng(4)
    n = 60
    df = pd.DataFrame({"온도": r.uniform(170, 190, n), "압력": r.uniform(1000, 3500, n), "라인": r.choice(["A", "B"], n)})
    df["y"] = df.온도 * 0.5 + df.압력 * 0.01 + (df.라인 == "A") * 3 + r.normal(0, 0.3, n)
    data = write_df(work / "small.csv", df)
    return {"data": data, "v": run_tool("ml_train.py", "-f", data, "-x", "온도,압력,라인", "-y", "y")["version_dir"]}


def test_inline_comma_number_stops_with_command(small):
    bad = run_fail("ml_predict.py", "-m", small["v"], "-i", "온도=180,190,압력=3000,라인=A")
    assert bad["agent_directive"].startswith("STOP_RETRY")                      # 2.9.13: 180190, VERIFIED
    rec = bad["message"].split("[추천] ")[1].strip()
    assert rec.startswith("python") and " -i 온도=180190,압력=3000,라인=A" in rec and "천 단위" in bad["message"]
    assert run_tool("ml_predict.py", *split_command(rec)[2:])["sample_count"] == 1
    bad = run_fail("ml_predict.py", "-m", small["v"], "-i", "온도=1,5,압력=3000,라인=A")
    assert "[추천]" not in bad["message"]
    # text values keep the joined text (not a number): unseen category, no stop
    ok = run_tool("ml_predict.py", "-m", small["v"], "-i", "온도=180,압력=3000,라인=A,B")
    assert ok["unseen_category_warning"]


def test_kv_repeated_name_refused():
    with pytest.raises(mp.InputError) as e:
        mp.parse_inline_input("온도=195,압력=3.8;온도=165,압력=2.1")
    assert "두 번" in str(e.value) and "JSON 배열" in str(e.value)
    with pytest.raises(mp.InputError):
        mp.parse_inline_input("온도=195, 온도=165")


def test_kv_alias_and_full_name_refused():
    # 2.9.28: refused where the names are matched to the model / data columns (check_same_column), not while parsing
    # (a model can have both '온도' and '온도(°C)'; test_names.py의 별칭 쌍 검사)
    with pytest.raises(mp.InputError) as e:
        mp.check_same_column(["온도", "온도(°C)"], ["온도(°C)", "온도(°C)"], "-i", {"온도": "180", "온도(°C)": "190"})
    assert "[후보] 온도=180, 온도(°C)=190" in str(e.value)
    # two different units are two columns; a plain row still parses
    assert list(mp.parse_inline_input("온도(°C)=180,온도(K)=453").columns) == ["온도(°C)", "온도(K)"]
    assert mp.parse_inline_input("라인=A,B,온도=180").iloc[0].to_dict() == {"라인": "A,B", "온도": "180"}


def test_kv_repeated_name_cli_stop_retry(tmp_path):
    rows = [["온도", "압력", "수율"]] + [[170 + i, 2 + i / 10, 50 + i] for i in range(12)]
    f = write_rows(str(tmp_path / "k.xlsx"), rows, [])
    proc, res = run_raw("ml_map.py", "-f", f, "-i", "온도=175,온도=176", "-y", "수율")
    assert proc.returncode == 1 and res["status"] == "error", res
    assert res["agent_directive"].startswith("STOP_RETRY:"), res


@pytest.fixture(scope="module")
def tiny(tmp_path_factory):
    work = tmp_path_factory.mktemp("tiny")
    r = np.random.default_rng(1)
    n = 150
    df = pd.DataFrame({"도핑농도": r.uniform(1e-6, 9e-6, n), "온도": r.uniform(0, 10, n)})
    df["저항"] = 1e6 * df["도핑농도"] + 0.5 * df["온도"] + r.normal(0, 0.05, n)
    p = work / "tiny.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "도핑농도,온도", "-y", "저항")["version_dir"]
    return {"work": work, "v": v}


def test_inline_typo_points_at_the_user_key(tiny):
    bad = run_fail("ml_predict.py", "-m", tiny["v"], "-i", "도핑농돜=5e-6,온도=5")
    # 2.9.33: [후보] (was "-i의 '도핑농돜' -> [추천] '도핑농도'"; 2.9.12: "'도핑농도' -> [추천] ['도핑농두']")
    assert "-i의 [후보] 도핑농돜 중 '도핑농도'" in bad["message"] and "[추천]" not in bad["message"], bad["message"]
    # 2.9.32: '도핑농두' (one common syllable for another) is [후보], the user picks
    bad = run_fail("ml_predict.py", "-m", tiny["v"], "-i", "도핑농두=5e-6,온도=5")
    assert "[추천]" not in bad["message"] and "-i의 [후보] 도핑농두 중 '도핑농도'" in bad["message"], bad["message"]


# ---- 예측 결과: 작은 값, 외삽 경고, 빈 입력, 결과 파일 --------------------------------------------------


def test_small_inputs_keep_training_range(tiny):
    meta = json.load(open(os.path.join(tiny["v"], "model_저항.json"), encoding="utf-8"))
    st = meta["feature_stats"]["도핑농도"]
    assert 1e-6 <= st["min"] < st["median"] < st["max"] <= 9e-6           # 2.9.12: all 0.0
    res = run_tool("ml_predict.py", "-m", tiny["v"], "-i", "도핑농도=5e-6,온도=5")
    assert res["drift_warning"] is False                                      # 2.9.12: +500% warning on every predict
    out = run_tool("ml_predict.py", "-m", tiny["v"], "-i", "도핑농도=2e-5,온도=5")
    assert out["drift_warning"] and "0.00" not in out["extrapolation_warnings"][0]
    f = run_tool("ml_predict.py", "-m", tiny["v"], "--find", "저항=6", "--vary", "도핑농도")
    lo, hi = f["training_range"]["도핑농도"]
    assert 0 < lo < hi and f["count"] >= 1                                    # 2.9.12: [0, 0] and count 0
    assert 1e-6 < f["solutions"][0]["도핑농도"] < 9e-6
    assert f["filled_defaults"]["온도"] > 0


def test_round_pred_keeps_small_values():
    import ml_predict as mp
    assert mp.round_pred(5.123456e-6) == pytest.approx(5.123e-6)
    assert mp.round_pred(123.456789) == 123.4568
    assert mp.round_pred(0.0) == 0.0
    arr = mp.round_pred(np.array([1.23456e-7, 2.5, np.nan]))
    assert arr[0] == pytest.approx(1.235e-7) and arr[1] == 2.5 and np.isnan(arr[2])


def test_extrapolation_warnings_capped(tmp_path):
    r = np.random.default_rng(5)
    cols = [f"x{i}" for i in range(25)]
    df = pd.DataFrame(r.uniform(0, 1, (40, 25)), columns=cols)
    df["y"] = df.sum(axis=1)
    v = run_tool("ml_train.py", "-f", write_df(tmp_path / "w.csv", df), "-x", ",".join(cols), "-y", "y")["version_dir"]
    res = run_tool("ml_predict.py", "-m", v, "-i", ",".join(f"{c}=5" for c in cols))
    assert "drift_warnings" not in res                                         # 2.9.13: two copies of every warning
    assert len(res["extrapolation_warnings"]) == 20 and res["extrapolation_warning_count"] == 25
    assert res["extrapolation_warnings_truncated"] is True and "25" in res["agent_directive"]


def test_prediction_with_named_empty_column_imputes(coded, tmp_path):
    df = pd.read_csv(coded["file"], dtype=str).head(10)
    df["b"] = ""
    p = tmp_path / "blank_b.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_predict.py", "-m", coded["v"], "-f", str(p), "--save", str(tmp_path / "out.csv"))
    assert res["missing_input_warning"] is True and res["missing_input_counts"]["b"] == 10


def test_inline_over_five_records_saved_to_file(coded):
    recs = [{"a": float(i), "b": 1.0, "코드": "012"} for i in range(8)]
    js = coded["work"] / "many.json"
    js.write_text(json.dumps(recs), encoding="utf-8")
    res = run_tool("ml_predict.py", "-m", coded["v"], "-i", str(js))
    assert res["truncated"] is True and res["count"] == 8 and res["output_file"]
    assert os.path.dirname(res["output_file"]) == str(coded["work"]) and res["verified"] is True
    assert len(pd.read_csv(res["output_file"], encoding="utf-8-sig")) == 8
    one = run_tool("ml_predict.py", "-m", coded["v"], "-i", "a=1,b=1,코드=012")
    assert one["truncated"] is False and one["output_file"] is None


def test_inline_json_string_over_five_saved_in_data_root(coded):
    recs = json.dumps([{"a": float(i), "b": 1.0, "코드": "012"} for i in range(6)])
    res = run_tool("ml_predict.py", "-m", coded["v"], "-i", recs)
    assert res["truncated"] is True
    assert os.path.dirname(res["output_file"]) == os.path.dirname(coded["v"].rstrip("\\/"))


def test_truncated_inline_has_full_result_file(tiny, tmp_path):
    rows = [{"도핑농도": 2e-6 + i * 1e-6, "온도": 5.0} for i in range(7)]
    p = tmp_path / "in.json"
    p.write_text(json.dumps(rows), encoding="utf-8")
    res = run_tool("ml_predict.py", "-m", tiny["v"], "-i", str(p))
    assert res["truncated"] is True and res["full_result_file"] == res["output_file"]


def test_out_extension_checked_first(coded, tmp_path):
    bad = run_fail("ml_predict.py", "-m", coded["v"], "-f", coded["file"], "--save", str(tmp_path / "r.txt"))
    assert bad["agent_directive"].startswith("STOP_RETRY") and "[추천]" in bad["message"]
    assert bad["message"].rstrip().endswith("r.csv") and not (tmp_path / "r.txt").exists()


def test_model_from_before_291_asks_for_retraining(tmp_path):
    """베타라 호환 경로가 없다: 2.9.1 전 형식의 모델은 반쪽 지원 대신 다시 학습하라는 STOP_RETRY."""
    import json
    import shutil
    r = np.random.default_rng(2)
    df = pd.DataFrame({"a": r.uniform(0, 1, 120), "b": r.uniform(0, 1, 120)})
    df["y"] = 2 * df.a + df.b + r.normal(0, 0.05, 120)
    p = tmp_path / "o.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "y")["version_dir"]
    old = tmp_path / "v_old"
    shutil.copytree(v, old)
    meta_p = old / "model_y.json"
    meta = json.loads(meta_p.read_text(encoding="utf-8"))
    del meta["input_columns_required"]
    meta_p.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    bad = run_fail("ml_predict.py", "-m", str(old), "-i", "a=0.5,b=0.5")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "다시 학습" in bad["message"]
    res = run_tool("ml_train.py", "-f", str(p), "--history")
    assert "mean_val_r2" not in res["records"][0] and "best_mean_val_r2" not in res


# ---- --find --vary 모델 역산: 모든 답, 90% 예측 구간의 운전 구간, 학습 중앙값 기본값, 실측 대조 ----


# y = 100 + 8 sin(pi (t - 170) / 25) + 6 (p - 3)^2: at p = 3, y = 105 at t = 175.37 and 189.63 (rises, then falls)
TRUE_105 = (175.37, 189.63)


@pytest.fixture(scope="module")
def hill(tmp_path_factory):
    d = tmp_path_factory.mktemp("find")
    r = np.random.default_rng(5)
    n = 400
    t, p = r.uniform(170, 195, n), r.uniform(2.5, 3.5, n)
    y = 100 + 8 * np.sin((t - 170) / 25 * np.pi) + 6 * (p - 3) ** 2 + r.normal(0, 0.4, n)
    path = d / "hill.csv"
    pd.DataFrame({"반응온도": t, "압축압력": p, "라인": np.where(np.arange(n) % 2, "A", "B"), "생산효율": y,
                  "등급": np.where(y > 104, "상", "하")}).to_csv(path, index=False)
    res = run_tool("ml_train.py", "-f", str(path), "-x", "반응온도,압축압력", "-y", "생산효율")
    return d, path, res["version_dir"]


def test_two_answers_windows_and_measured(hill):
    _, _, vdir = hill
    r = run_tool("ml_predict.py", "-m", vdir, "--find", "생산효율=105", "--vary", "반응온도", "-i", "압축압력=3")
    assert r["count"] == 2 and r["verified"] and r["agent_directive"].startswith("VERIFIED:")
    for ans, truth in zip(r["solutions"], TRUE_105):
        assert abs(ans["반응온도"] - truth) < 1.0
    wins = r["operating_windows"]
    assert wins and all(any(w["from"] <= a["반응온도"] <= w["to"] for w in wins) for a in r["solutions"])
    m = r["measured"]
    assert m["count"] == 2 and all(abs(v - t) < 1.0 for v, t in zip(m["solutions"], TRUE_105))
    assert len(r["nearest_rows"]) == 5 and r["filled_defaults"] == {}
    assert "운전 구간" in r["preview"] and "실측 기준" in r["preview"]


def test_unreachable_and_default_inputs(hill):
    _, _, vdir = hill
    r = run_tool("ml_predict.py", "-m", vdir, "--find", "생산효율=112", "--vary", "반응온도", "-i", "압축압력=3")
    assert r["count"] == 0 and r["solutions"] == [] and r["reachable"]["max"] < 112
    assert "찾은 값 없음" in r["preview"] and "count 0" in r["agent_directive"]
    d = run_tool("ml_predict.py", "-m", vdir, "--find", "생산효율=105", "--vary", "반응온도")   # 압축압력 not given
    assert set(d["filled_defaults"]) == {"압축압력(bar)"} or set(d["filled_defaults"]) == {"압축압력"}
    assert "중앙값" in d["warnings"][0] and d["count"] >= 1


def test_out_of_range_fixed_input_warns(hill):
    _, _, vdir = hill
    r = run_tool("ml_predict.py", "-m", vdir, "--find", "생산효율=105", "--vary", "반응온도", "-i", "압축압력=9")
    assert any("학습 범위" in w for w in r["warnings"])


def test_inputs_errors(hill):
    d, path, vdir = hill
    cat = run_tool("ml_train.py", "-f", str(path), "-x", "반응온도,압축압력,라인", "-y", "생산효율")
    miss = run_fail("ml_predict.py", "-m", cat["version_dir"], "--find", "생산효율=105", "--vary", "반응온도", "-i", "압축압력=3")
    assert "[후보] A, B" in miss["message"]                                        # category never guessed
    ok = run_tool("ml_predict.py", "-m", cat["version_dir"], "--find", "생산효율=105", "--vary", "반응온도",
                  "-i", "압축압력=3,라인=A")
    assert ok["status"] == "ok"
    assert "숫자 입력이 아니라" in run_fail("ml_predict.py", "-m", cat["version_dir"], "--find", "생산효율=105",
                                           "--vary", "라인", "-i", "압축압력=3")["message"]
    cls = run_tool("ml_train.py", "-f", str(path), "-x", "반응온도,압축압력", "-y", "등급", "--task", "classification")
    assert "분류" in run_fail("ml_predict.py", "-m", cls["version_dir"], "--find", "등급=1", "--vary", "반응온도")["message"]
    proc, payload = run_raw("ml_predict.py", "-m", vdir, "--find", "생산효율=105")                 # --vary missing
    assert proc.returncode == 2 and payload["status"] == "error"
    proc, payload = run_raw("ml_predict.py", "-m", vdir, "--find", "생산효율=105", "--vary", "반응온도", "-f", str(path))
    assert proc.returncode == 2


def test_linear_model_exact_answer(tmp_path):
    r = np.random.default_rng(7)
    n = 200
    t, p = r.uniform(170, 195, n), r.uniform(2.5, 3.5, n)
    path = tmp_path / "lin.csv"
    pd.DataFrame({"반응온도": t, "압축압력": p, "생산효율": 0.5 * t + 8 * p + r.normal(0, 0.05, n)}).to_csv(path, index=False)
    v = run_tool("ml_train.py", "-f", str(path), "-x", "반응온도,압축압력", "-y", "생산효율")["version_dir"]
    res = run_tool("ml_predict.py", "-m", v, "--find", "생산효율=114", "--vary", "반응온도", "-i", "압축압력=3")
    assert res["count"] == 1 and res["solutions"][0]["반응온도"] == pytest.approx((114 - 24) / 0.5, abs=0.5)


def test_many_inputs_skip_measured_lookup(tmp_path):
    r = np.random.default_rng(8)
    n = 200
    df = pd.DataFrame({c: r.uniform(0, 1, n) for c in ("a", "b", "c", "d")})
    df["y"] = df.a * 10 + df.b + df.c + df.d
    path = tmp_path / "m.csv"
    df.to_csv(path, index=False)
    v = run_tool("ml_train.py", "-f", str(path), "-x", "a,b,c,d", "-y", "y")["version_dir"]
    res = run_tool("ml_predict.py", "-m", v, "--find", "y=6", "--vary", "a", "-i", "b=0.5,c=0.5,d=0.5")
    assert "생략" in res["measured"]["note"] and len(res["nearest_rows"]) == 5
    assert res["count"] >= 1 and abs(res["solutions"][0]["a"] - 0.45) < 0.1


def test_find_with_text_inline_values(coded):
    res = run_tool("ml_predict.py", "-m", coded["v"], "--find", "y=30", "--vary", "a", "-i", "b=2,코드=007")
    assert res["fixed"] == {"b": 2.0, "코드": "007"} and res["count"] >= 1


def test_find_comma_number_stops_with_command(small):
    bad = run_fail("ml_predict.py", "-m", small["v"], "--find", "y=1,000", "--vary", "온도", "-i", "압력=3000,라인=A")
    assert '--find y=1000' in bad["message"].split("[추천] ")[1]
    bad = run_fail("ml_predict.py", "-m", small["v"], "--find", "y=120", "--vary", "온도", "-i", "압력=3,000,라인=A")
    rec = bad["message"].split("[추천] ")[1].strip()
    assert ' -i 압력=3000,라인=A' in rec
    assert run_tool("ml_predict.py", *split_command(rec)[2:])["command"] == "find"


def test_find_multi_row_input_refused(p1_model):
    proc, res = run_raw("ml_predict.py", "-m", p1_model["v"], "--find", "생산량=530", "--vary", "온도",
                        "-i", '[{"압력":2.2},{"압력":3.8}]')
    assert "한 줄" in stop_msg(res, proc)
    proc, res = run_raw("ml_predict.py", "-m", p1_model["v"], "--find", '[{"생산량":530},{"생산량":520}]',
                        "--vary", "온도")
    assert "한 줄" in stop_msg(res, proc)
    proc, res = run_raw("ml_predict.py", "-m", p1_model["v"], "--find", '[{"생산량":530}]', "--vary", "온도",
                        "-i", '[{"압력":2.2}]')
    assert res["status"] == "ok" and res["fixed"] == {"압력(bar)": 2.2}, res


def test_nearest_rows_are_numbers_and_find_rejects_text_input(tmp_path):
    r = np.random.default_rng(3)
    df = pd.DataFrame({"온도": r.uniform(150, 200, 200), "압력": r.uniform(1, 3, 200)})
    df["효율"] = 0.4 * df.온도 + 5 * df.압력 + r.normal(0, 0.2, 200)
    p = tmp_path / "n.csv"
    df.to_csv(p, index=False)
    m = run_tool("ml_map.py", "-f", str(p), "--find", "효율=80", "--vary", "온도", "-i", "압력=2")
    assert m["nearest_rows"] and all(isinstance(nr["온도"], float) for nr in m["nearest_rows"])
    v = run_tool("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "효율")["version_dir"]
    f = run_tool("ml_predict.py", "-m", v, "--find", "효율=80", "--vary", "온도", "-i", "압력=2")
    assert all(isinstance(nr["온도"], float) for nr in f["nearest_rows"])
    bad = run_fail("ml_predict.py", "-m", v, "--find", "효율=80", "--vary", "온도", "-i", "압력=x")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "압력" in bad["message"] and "숫자" in bad["message"]


# ---- 실측 비교 map_check (회귀 예측마다) --------------------------------------------------------------


# 버전 폴더의 model_dataset.csv로 계산한다: 숫자 입력 3개 이하면 표 보간, 더 많으면 가까운 5행의 IDW.


N = 200


def _mc_data(path, marks=False):
    r = np.random.default_rng(3)
    t, p, c, w = r.uniform(170, 195, N), r.uniform(2.5, 3.5, N), r.uniform(2, 3, N), r.uniform(0.5, 1.5, N)
    y = 0.5 * t + 8 * p - 3 * c - 2 * w + r.normal(0, 0.5, N)
    df = pd.DataFrame({"라인": np.where(np.arange(N) % 2, "A", "B"), "반응온도": t, "압축압력": p,
                       "촉매비율": c, "수분함량": w, "생산효율": y, "등급": np.where(y > np.median(y), "상", "하")})
    if marks:
        df["제외"] = [""] * N
        df.loc[10, "제외"] = "센서오류"
    df.to_excel(path, index=False)
    return df


@pytest.fixture(scope="module")
def mc_model(tmp_path_factory):
    d = tmp_path_factory.mktemp("mc")
    data = d / "d.xlsx"
    df = _mc_data(data, marks=True)
    res = run_tool("ml_train.py", "-f", str(data), "-x", "반응온도,압축압력", "-y", "생산효율")
    return d, df, res["version_dir"]


def test_snapshot_has_rows_and_exclusions(mc_model):
    _, _, vdir = mc_model
    snap = pd.read_csv(os.path.join(vdir, "model_dataset.csv"), encoding="utf-8-sig")
    assert list(snap["원본행"][:3]) == [2, 3, 4]                       # Excel rows (header is row 1)
    assert snap.loc[snap["원본행"] == 12, "학습제외_생산효율"].item() == 1   # '센서오류' on data row 10 = Excel row 12


def test_inline_prediction_has_map_check(mc_model):
    _, df, vdir = mc_model
    res = run_tool("ml_predict.py", "-m", vdir, "-i", "반응온도=180,압축압력=3")
    ps = res["predictions_summary"]["생산효율"]
    mc = ps["map_check"][0]
    assert mc["method"].startswith("둘러싼 실측 3점") and len(mc["basis"]) == 3
    assert all(b["row"] >= 2 and b["row"] != 12 for b in mc["basis"])  # Excel rows; the excluded row is never used
    assert mc["diff"] == pytest.approx(ps["predictions"][0] - mc["value"], abs=1e-3)
    assert not res["map_check_warning"] and ps["map_check_keys"] == ["반응온도", "압축압력"]
    assert "실측 비교값" in res["agent_directive"]


def test_file_prediction_columns_and_extrapolation(mc_model, tmp_path):
    _, _, vdir = mc_model
    new = tmp_path / "new.csv"
    pd.DataFrame({"반응온도": [180, 260, None], "압축압력": [3, 3, 3]}).to_csv(new, index=False)
    res = run_tool("ml_predict.py", "-m", vdir, "-f", str(new))
    out = pd.read_csv(res["output_file"], encoding="utf-8-sig")
    assert {"map_생산효율", "map_diff_생산효율", "map_방식_생산효율"} <= set(out.columns)
    assert out["map_방식_생산효율"][1].endswith("[외삽]") and out["map_방식_생산효율"][2].startswith("계산 불가")
    assert out["map_diff_생산효율"][0] == pytest.approx(out["pred_생산효율"][0] - out["map_생산효율"][0], abs=1e-3)
    assert not res["map_check_warning"]                                  # extrapolated comparisons never warn


def test_warning_when_measured_data_disagrees(mc_model, tmp_path):
    _, _, vdir = mc_model
    v2 = tmp_path / "v_tampered"
    shutil.copytree(vdir, v2)
    snap_path = v2 / "model_dataset.csv"
    snap = pd.read_csv(snap_path, encoding="utf-8-sig")
    snap["생산효율"] += 10                                                # measured data now 10 above the model
    snap.to_csv(snap_path, index=False, encoding="utf-8-sig")
    res = run_tool("ml_predict.py", "-m", str(v2), "-i", "반응온도=180,압축압력=3")
    assert res["map_check_warning"] and "90% 예측 범위 밖" in res["map_check_warnings"][0]
    assert res["predictions_summary"]["생산효율"]["map_check"][0]["outside_interval"]
    assert "실측 비교 경고" in res["preview"]


def test_old_snapshot_and_missing_snapshot(mc_model, tmp_path):
    _, _, vdir = mc_model
    old = tmp_path / "v_old"
    shutil.copytree(vdir, old)
    snap = pd.read_csv(old / "model_dataset.csv", encoding="utf-8-sig").drop(columns=["원본행", "학습제외_생산효율"])
    snap.to_csv(old / "model_dataset.csv", index=False, encoding="utf-8-sig")
    res = run_tool("ml_predict.py", "-m", str(old), "-i", "반응온도=180,압축압력=3")     # snapshot from before 2.7.0
    assert "map_check" not in res["predictions_summary"]["생산효율"] and "다시 학습" in res["map_check_notes"][0]

    os.remove(old / "model_dataset.csv")
    res = run_tool("ml_predict.py", "-m", str(old), "-i", "반응온도=180,압축압력=3")
    assert "map_check" not in res["predictions_summary"]["생산효율"] and "model_dataset.csv" in res["map_check_notes"][0]
    assert res["predictions_summary"]["생산효율"]["predictions"]           # prediction itself unaffected


def test_many_inputs_use_nearest_rows_and_category(tmp_path):
    data = tmp_path / "d.xlsx"
    _mc_data(data)
    res = run_tool("ml_train.py", "-f", str(data), "-x", "반응온도,압축압력,촉매비율,수분함량,라인", "-y", "생산효율")
    out = run_tool("ml_predict.py", "-m", res["version_dir"], "-i", "반응온도=180,압축압력=3,촉매비율=2.5,수분함량=1,라인=A")
    mc = out["predictions_summary"]["생산효율"]["map_check"][0]
    assert mc["method"].startswith("가까운 실측") and 1 <= len(mc["basis"]) <= 5
    assert sum(b["weight"] for b in mc["basis"]) == pytest.approx(1.0, abs=1e-4)
    assert all(b["row"] % 2 == 1 for b in mc["basis"])                  # 라인 A = odd data rows = odd Excel rows (row = i + 2)


def test_classification_has_no_map_check_and_standalone(mc_model, tmp_path):
    d, _, vdir = mc_model
    cls = run_tool("ml_train.py", "-f", str(d / "d.xlsx"), "-x", "반응온도,압축압력", "-y", "등급", "--task", "classification")
    res = run_tool("ml_predict.py", "-m", cls["version_dir"], "-i", "반응온도=180,압축압력=3")
    assert "map_check" not in res["predictions_summary"]["등급"] and not res["map_check_warning"]

    deploy = tmp_path / "deploy"
    deploy.mkdir()
    shutil.copy(os.path.join(KIT, "ml_predict.py"), deploy / "ml_predict.py")
    shutil.copytree(vdir, deploy / "v1")
    import subprocess
    import json
    proc = subprocess.run([sys.executable, str(deploy / "ml_predict.py"), "-m", str(deploy / "v1"), "-i", "반응온도=180,압축압력=3"],
                          capture_output=True, text=True, encoding="utf-8", cwd=str(deploy))
    payload = json.loads(proc.stdout.strip().splitlines()[-1])
    assert payload["status"] == "ok" and payload["predictions_summary"]["생산효율"]["map_check"][0]["value"] is not None


def _snapshot(version_dir, n=3000):
    r = np.random.default_rng(1)
    df = pd.DataFrame(r.uniform(0, 1, (n, 4)), columns=list("abcd"))
    df["y"] = df.sum(axis=1)
    df["원본행"] = np.arange(2, n + 2)
    df.to_csv(os.path.join(version_dir, "model_dataset.csv"), index=False, encoding="utf-8-sig")
    return df


def test_idw_comparison_converts_snapshot_once_and_keeps_values(tmp_path, monkeypatch):
    """IDW 비교는 입력 행마다 스냅숏 열을 다시 변환하지 않는다 (2만 행에서 15분이 넘던 문제), 값은 손 계산과 같다."""
    df = _snapshot(str(tmp_path))
    meta = {"feature_schema": {}, "input_columns_required": list("abcd")}
    monkeypatch.setattr(mp, "model_raw_inputs", lambda m: list("abcd"))
    rows = df[list("abcd")].iloc[:2000].reset_index(drop=True) + 0.001

    calls = {"n": 0}
    real = mp.coerce_numeric

    def counting(s):
        calls["n"] += 1
        return real(s)

    monkeypatch.setattr(mp, "coerce_numeric", counting)
    t0 = time.time()
    out = mp.compare_with_measured(meta, "y", str(tmp_path), rows, {})
    took = time.time() - t0

    assert calls["n"] < 30                                            # not once per input row
    assert took < 20
    assert np.isfinite(out["values"]).all() and len(out["basis"][0]) == 5

    # same answer as computing the first row by hand
    x = rows.iloc[0].to_numpy()
    K = df[list("abcd")].to_numpy()
    span = K.max(axis=0) - K.min(axis=0)
    d = np.sqrt((((K - x) / span) ** 2).sum(axis=1))
    near = np.argsort(d, kind="stable")[:5]
    w = 1 / d[near]
    assert abs(out["values"][0] - float((w / w.sum()) @ df["y"].to_numpy()[near])) < 1e-9


# ---- 조건별 예측(--sweep)과 조건 찾기(--maximize/--minimize) (2.9.58) ------------------------------------------


@pytest.fixture(scope="module")
def whatif(tmp_path_factory):
    work = tmp_path_factory.mktemp("whatif")
    r = np.random.default_rng(11)
    n = 300
    t = r.uniform(170, 200, n)
    t[:2] = 170, 200                     # the training range is exactly 170~200
    pr = r.uniform(2, 4, n)
    c = r.uniform(0.5, 1.5, n)
    eff = 0.5 * t + 5 * pr + 2 * c + r.normal(0, 0.3, n)
    imp = 0.2 * (t - 185) ** 2 + 10 * pr + 30 + r.normal(0, 0.5, n)
    data = write_df(work / "w.csv", pd.DataFrame({"반응온도": t, "압축압력": pr, "촉매비율": c, "생산효율": eff, "불순물": imp}))
    v = run_tool("ml_train.py", "-f", data, "-x", "반응온도,압축압력,촉매비율", "-y", "생산효율,불순물")["version_dir"]
    return {"v": v, "work": work}


def test_sweep_table_chart_trend(whatif):
    res = run_tool("ml_predict.py", "-m", whatif["v"], "--sweep", "반응온도=170:200:5", "-i", "압축압력=3,촉매비율=1")
    assert res["command"] == "sweep" and res["count"] == 7 and res["truncated"] is False and len(res["points"]) == 7
    assert [p["반응온도"] for p in res["points"]] == [170, 175, 180, 185, 190, 195, 200]
    assert res["trend"]["생산효율"]["direction"] == "increasing", res["trend"]
    assert res["trend"]["불순물"]["direction"] == "valley" and abs(res["trend"]["불순물"]["min"]["반응온도"] - 185) <= 5
    assert "생산효율_90%" in res["points"][0] and res["out_of_range_points"] == 0 and res["filled_defaults"] == {}
    assert res["output_file"].endswith(".xlsx") and os.path.dirname(res["output_file"]) == os.path.dirname(whatif["v"])
    tab = pd.read_excel(res["output_file"])
    assert len(tab) == 7 and {"pred_생산효율", "pred_lo_생산효율", "pred_hi_불순물", "학습범위"} <= set(tab.columns)
    assert os.path.getsize(res["chart_file"]) > 0 and res["chart_file"].endswith(".png")
    assert res["agent_directive"].startswith("VERIFIED") and "trend" in res["agent_directive"]


def test_sweep_marks_extrapolation_and_fills_medians(whatif):
    res = run_tool("ml_predict.py", "-m", whatif["v"], "--sweep", "반응온도=160:210:10", "-y", "생산효율")
    assert res["targets"] == ["생산효율"] and res["out_of_range_points"] == 2
    assert [p["in_range"] for p in res["points"]] == [False, True, True, True, True, False]
    assert any("외삽" in w for w in res["warnings"]) and "중앙값" in res["warnings"][0]
    assert set(res["filled_defaults"]) == {"압축압력", "촉매비율"}


def test_sweep_caps_points(whatif):
    bad = run_fail("ml_predict.py", "-m", whatif["v"], "--sweep", "반응온도=170:200:0.01", "-i", "압축압력=3,촉매비율=1")
    rec = bad["message"].split("[추천] ")[1]
    assert rec.startswith("python ") and "--sweep" in rec
    res = run_tool("ml_predict.py", *split_command(rec)[2:])
    assert 25 < res["count"] <= 500 and res["truncated"] is True and res["shown"] == len(res["points"]) == 25
    assert res["full_result_file"] == res["output_file"] and len(pd.read_excel(res["output_file"])) == res["count"]
    assert res["points"][0]["반응온도"] == 170 and res["points"][-1]["반응온도"] == 200


def test_maximize_with_constraint(whatif):
    res = run_tool("ml_predict.py", "-m", whatif["v"], "--maximize", "생산효율", "--where", "불순물 <= 80",
                   "--vary", "반응온도,압축압력", "-i", "촉매비율=1")
    assert res["command"] == "maximize" and res["count"] >= 1 and res["where"] == "불순물 <= 80"
    best = res["solutions"][0]
    assert best["predictions"]["불순물"] <= 80 and best["inputs"]["압축압력"] > 3.5, best
    assert 185 < best["inputs"]["반응온도"] < 197, best
    assert all(170 <= s["inputs"]["반응온도"] <= 200 and 2 <= s["inputs"]["압축압력"] <= 4 for s in res["solutions"])
    effs = [s["predictions"]["생산효율"] for s in res["solutions"]]
    assert effs == sorted(effs, reverse=True) and len(res["nearest_rows"]) == 5
    assert "interval_90" in best and res["agent_directive"].startswith("VERIFIED")


def test_search_rechecks_answers(whatif, monkeypatch):
    """2.9.59: 'VERIFIED: 찾은 조합을 모델로 다시 예측해 조건을 확인했습니다'라고 했지만 다시 예측하지 않았다.
    이제 적힌 inputs로 다시 예측해 predictions와 --where를 확인하고, 다르면 UNVERIFIED."""
    import argparse
    args = argparse.Namespace(model=whatif["v"], maximize="생산효율", minimize=None, vary="반응온도,압축압력",
                              where="불순물 <= 80", input="촉매비율=1", targets=None)
    res = mp.execute_search(args)
    assert res["verified"] is True and res["agent_directive"].startswith("VERIFIED: 답마다 적힌 입력값"), res
    real = mp._predict_rows
    # the re-check predicts the (at most 5) answers in one call: a model that answers otherwise there is caught
    monkeypatch.setattr(mp, "_predict_rows", lambda it, rows: real(it, rows) + (50.0 if len(rows) <= 5 else 0.0))
    res = mp.execute_search(args)
    assert res["verified"] is False and res["agent_directive"].startswith("UNVERIFIED"), res
    assert res["unverified"] and "답 1" in res["unverified"][0], res["unverified"]


def test_minimize_impossible_constraint_is_zero(whatif):
    res = run_tool("ml_predict.py", "-m", whatif["v"], "--minimize", "불순물", "--where", "생산효율 > 1000",
                   "--vary", "반응온도", "-i", "압축압력=3,촉매비율=1")
    assert res["count"] == 0 and res["solutions"] == [] and "best_without_where" in res
    assert res["agent_directive"].startswith("VERIFIED: 0건")
    assert abs(res["best_without_where"]["inputs"]["반응온도"] - 185) <= 3


def test_whatif_argument_rules(whatif):
    proc, res = run_raw("ml_predict.py", "-m", whatif["v"], "--where", "불순물 <= 80", "-i", "반응온도=180")
    assert proc.returncode == 2, res
    proc, res = run_raw("ml_predict.py", "-m", whatif["v"], "--sweep", "반응온도", "--find", "생산효율=100", "--vary", "반응온도")
    assert proc.returncode == 2, res
    bad = run_fail("ml_predict.py", "-m", whatif["v"], "--maximize", "생산효율", "--vary", "반응온도", "--where", "불순물 == 80")
    assert "--where '불순물 = 80'" in bad["message"].split("[추천] ")[1]


# ---- 범주 입력을 학습 범주와 맞추기 (2.9.60) --------------------------------------------------------------------


@pytest.fixture(scope="module")
def line_model(tmp_path_factory):
    work = tmp_path_factory.mktemp("line")
    r = np.random.default_rng(5)
    n = 240
    t = r.uniform(200, 280, n)
    line = np.array(["A", "B", "C"] * (n // 3))
    y = 0.2 * t + np.select([line == "A", line == "B", line == "C"], [0, 10, 20]) + r.normal(0, 0.3, n)
    data = write_df(work / "cat.csv", pd.DataFrame({"온도": t, "라인": line, "수율": y}))
    return run_tool("ml_train.py", "-f", data, "-x", "온도,라인", "-y", "수율")["version_dir"]


def test_whatif_unknown_category_stops_with_candidates(line_model):
    """2.9.59: -i "라인=라인C"를 학습 범주와 대조하지 않아 기준값으로 예측하고 VERIFIED였다."""
    for opt in (["--sweep", "온도"], ["--maximize", "수율", "--vary", "온도"], ["--find", "수율=60", "--vary", "온도"]):
        bad = run_fail("ml_predict.py", "-m", line_model, *opt, "-i", "라인=라인C")
        assert bad["agent_directive"].startswith("STOP_RETRY"), bad
        assert "[후보] A, B, C" in bad["message"] and "라인C" in bad["message"], bad["message"]


def test_whatif_category_case_width_space_recovered(line_model):
    ref = run_tool("ml_predict.py", "-m", line_model, "--sweep", "온도=200:280:40", "-i", "라인=C")
    for given in ("c", " C ", "Ｃ"):              # 소문자, 앞뒤 공백, 전각
        res = run_tool("ml_predict.py", "-m", line_model, "--sweep", "온도=200:280:40", "-i", f"라인={given}")
        assert res["fixed"] == {"라인": "C"}, res
        assert [p["수율"] for p in res["points"]] == [p["수율"] for p in ref["points"]]
        if given != " C ":                            # 앞뒤 공백은 읽을 때 이미 지워져 바뀐 것이 없다
            assert any("'C'" in m for m in res.get("recovered", [])), res
    res = run_tool("ml_predict.py", "-m", line_model, "--maximize", "수율", "--vary", "온도", "-i", "라인=b")
    assert res["fixed"] == {"라인": "B"} and res.get("recovered"), res
    res = run_tool("ml_predict.py", "-m", line_model, "--find", "수율=60", "--vary", "온도", "-i", "라인=a")
    assert res.get("recovered") and res["fixed"] == {"라인": "A"}, res


def test_predict_category_case_matched_like_whatif(line_model):
    """일반 예측 -i도 같은 공통 함수: 대소문자만 다르면 맞추고 recovered, 그 밖은 unseen_category_warnings."""
    ref = run_tool("ml_predict.py", "-m", line_model, "-i", "온도=250,라인=C")
    res = run_tool("ml_predict.py", "-m", line_model, "-i", "온도=250,라인=c")
    assert res["unseen_category_warning"] is False and res.get("recovered"), res
    assert res["predictions_summary"]["수율"]["predictions"] == ref["predictions_summary"]["수율"]["predictions"]
    res = run_tool("ml_predict.py", "-m", line_model, "-i", "온도=250,라인=라인C")
    assert res["unseen_category_warning"] is True and "recovered" not in res, res


def test_match_categories_ambiguous_not_fixed():
    fixes, unseen = mp.match_categories(["a", "b ", "x"], ["A", "a", "B"])
    assert fixes == {"b ": "B"} and unseen == ["x"]      # 'a'는 학습 범주 그대로 ('A'로 바꾸지 않음)
    fixes, unseen = mp.match_categories(["ａ"], ["A", "a"])
    assert fixes == {} and unseen == ["ａ"]          # 'A'와 'a' 둘 다 맞아 고를 수 없음
