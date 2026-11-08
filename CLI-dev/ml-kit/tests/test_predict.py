# -*- coding: utf-8 -*-
"""예측(ml_predict.py): --input 값 읽기, 예측 결과와 결과 파일, --find 역산, 실측 비교(map_check).

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
import _ml_runtime as rt  # noqa: E402


# ---- --input 값 읽기: 글자 그대로, 천 단위, 같은 이름 두 번, 키 오타 ------------------------------------


def test_inline_pairs_stay_text():
    df = rt.parse_inline_input("코드=007,a=1.5")
    assert df.loc[0, "코드"] == "007" and df.loc[0, "a"] == "1.5"


def test_inline_category_code_with_leading_zero(coded):
    res = run_tool("ml_predict.py", "--model", coded["v"], "--input", "a=5,b=2,코드=007")
    assert res["unseen_category_warning"] is False and res["missing_input_warning"] is False
    js = coded["work"] / "one.json"
    js.write_text(json.dumps({"a": 5, "b": 2, "코드": "007"}), encoding="utf-8")
    same = run_tool("ml_predict.py", "--model", coded["v"], "--input", str(js))
    assert res["predictions_summary"]["y"]["predictions"] == same["predictions_summary"]["y"]["predictions"]


def test_inline_thousands_separator():
    df = rt.parse_inline_input("온도=3,압력=500,000,라인=A")
    assert df.iloc[0].to_dict() == {"온도": "3", "압력": "500,000", "라인": "A"}
    assert float(rt.coerce_numeric(df["압력"]).iloc[0]) == 500000.0
    with pytest.raises(rt.InputError):
        rt.parse_inline_input("500,온도=3")


@pytest.fixture(scope="module")
def small(tmp_path_factory):
    work = tmp_path_factory.mktemp("comma")
    r = np.random.default_rng(4)
    n = 60
    df = pd.DataFrame({"온도": r.uniform(170, 190, n), "압력": r.uniform(1000, 3500, n), "라인": r.choice(["A", "B"], n)})
    df["y"] = df.온도 * 0.5 + df.압력 * 0.01 + (df.라인 == "A") * 3 + r.normal(0, 0.3, n)
    data = write_df(work / "small.csv", df)
    return {"data": data, "v": run_tool("ml_train.py", "--file", data, "--features", "온도,압력,라인", "--targets", "y")["version_dir"]}


def test_inline_comma_number_stops_with_command(small):
    bad = run_fail("ml_predict.py", "--model", small["v"], "--input", "온도=180,190,압력=3000,라인=A")
    assert bad["agent_directive"].startswith("STOP_RETRY")                      # 2.9.13: 180190, VERIFIED
    rec = bad["message"].split("[추천] ")[1].strip()
    assert rec.startswith("python") and " --input 온도=180190,압력=3000,라인=A" in rec and "천 단위" in bad["message"]
    assert run_tool("ml_predict.py", *split_command(rec)[2:])["sample_count"] == 1
    bad = run_fail("ml_predict.py", "--model", small["v"], "--input", "온도=1,5,압력=3000,라인=A")
    assert "[추천]" not in bad["message"]
    # text values keep the joined text (not a number): unseen category, no stop
    ok = run_tool("ml_predict.py", "--model", small["v"], "--input", "온도=180,압력=3000,라인=A,B")
    assert ok["unseen_category_warning"]


def test_kv_repeated_name_refused():
    with pytest.raises(rt.InputError) as e:
        rt.parse_inline_input("온도=195,압력=3.8;온도=165,압력=2.1")
    assert "두 번" in str(e.value) and "JSON 배열" in str(e.value)
    with pytest.raises(rt.InputError):
        rt.parse_inline_input("온도=195, 온도=165")


def test_kv_alias_and_full_name_refused():
    # 2.9.28: refused where the names are matched to the model / data columns (check_same_column), not while parsing
    # (a model can have both '온도' and '온도(°C)'; test_names.py의 별칭 쌍 검사)
    with pytest.raises(rt.InputError) as e:
        rt.check_same_column(["온도", "온도(°C)"], ["온도(°C)", "온도(°C)"], "--input", {"온도": "180", "온도(°C)": "190"})
    assert "[후보] 1) 온도=180 2) 온도(°C)=190" in str(e.value)
    # two different units are two columns; a plain row still parses
    assert list(rt.parse_inline_input("온도(°C)=180,온도(K)=453").columns) == ["온도(°C)", "온도(K)"]
    assert rt.parse_inline_input("라인=A,B,온도=180").iloc[0].to_dict() == {"라인": "A,B", "온도": "180"}


def test_kv_repeated_name_cli_stop_retry(tmp_path):
    rows = [["온도", "압력", "수율"]] + [[170 + i, 2 + i / 10, 50 + i] for i in range(12)]
    f = write_rows(str(tmp_path / "k.xlsx"), rows, [])
    proc, res = run_raw("ml_map.py", "--file", f, "--input", "온도=175,온도=176", "--targets", "수율")
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
    v = run_tool("ml_train.py", "--file", str(p), "--features", "도핑농도,온도", "--targets", "저항")["version_dir"]
    return {"work": work, "v": v}


def test_inline_typo_points_at_the_user_key(tiny):
    bad = run_fail("ml_predict.py", "--model", tiny["v"], "--input", "도핑농돜=5e-6,온도=5")
    # 2.9.33: [후보] (was "--input의 '도핑농돜' -> [추천] '도핑농도'"; 2.9.12: "'도핑농도' -> [추천] ['도핑농두']")
    assert "--input의 이름 중 '도핑농도'(으)로 고칠 이름을 사용자에게 물으십시오: [후보] 1) 도핑농돜" in bad["message"] and "[추천]" not in bad["message"], bad["message"]
    # 2.9.32: '도핑농두' (one common syllable for another) is [후보], the user picks
    bad = run_fail("ml_predict.py", "--model", tiny["v"], "--input", "도핑농두=5e-6,온도=5")
    assert "[추천]" not in bad["message"] and "--input의 이름 중 '도핑농도'(으)로 고칠 이름을 사용자에게 물으십시오: [후보] 1) 도핑농두" in bad["message"], bad["message"]


# ---- 예측 결과: 작은 값, 외삽 경고, 빈 입력, 결과 파일 --------------------------------------------------


def test_small_inputs_keep_training_range(tiny):
    meta = json.load(open(os.path.join(tiny["v"], "model_저항.json"), encoding="utf-8"))
    st = meta["feature_stats"]["도핑농도"]
    assert 1e-6 <= st["min"] < st["median"] < st["max"] <= 9e-6           # 2.9.12: all 0.0
    res = run_tool("ml_predict.py", "--model", tiny["v"], "--input", "도핑농도=5e-6,온도=5")
    assert res["drift_warning"] is False                                      # 2.9.12: +500% warning on every predict
    out = run_tool("ml_predict.py", "--model", tiny["v"], "--input", "도핑농도=2e-5,온도=5")
    assert out["drift_warning"] and "0.00" not in out["extrapolation_warnings"][0]
    f = run_tool("ml_predict.py", "--model", tiny["v"], "--find", "저항=6", "--vary", "도핑농도")
    lo, hi = f["training_range"]["도핑농도"]
    assert 0 < lo < hi and f["count"] >= 1                                    # 2.9.12: [0, 0] and count 0
    assert 1e-6 < f["solutions"][0]["도핑농도"] < 9e-6
    assert f["filled_defaults"]["온도"] > 0


def test_round_pred_keeps_small_values():
    import _ml_runtime as rt
    assert rt.round_pred(5.123456e-6) == pytest.approx(5.123e-6)
    assert rt.round_pred(123.456789) == 123.4568
    assert rt.round_pred(0.0) == 0.0
    arr = rt.round_pred(np.array([1.23456e-7, 2.5, np.nan]))
    assert arr[0] == pytest.approx(1.235e-7) and arr[1] == 2.5 and np.isnan(arr[2])


def test_extrapolation_warnings_capped(tmp_path):
    r = np.random.default_rng(5)
    cols = [f"x{i}" for i in range(25)]
    df = pd.DataFrame(r.uniform(0, 1, (40, 25)), columns=cols)
    df["y"] = df.sum(axis=1)
    v = run_tool("ml_train.py", "--file", write_df(tmp_path / "w.csv", df), "--features", ",".join(cols), "--targets", "y")["version_dir"]
    res = run_tool("ml_predict.py", "--model", v, "--input", ",".join(f"{c}=5" for c in cols))
    assert "drift_warnings" not in res                                         # 2.9.13: two copies of every warning
    assert len(res["extrapolation_warnings"]) == 20 and res["extrapolation_warning_count"] == 25
    assert res["extrapolation_warnings_truncated"] is True and "25" in res["agent_directive"]


def test_prediction_with_named_empty_column_imputes(coded, tmp_path):
    df = pd.read_csv(coded["file"], dtype=str).head(10)
    df["b"] = ""
    p = tmp_path / "blank_b.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_predict.py", "--model", coded["v"], "--file", str(p), "--save", str(tmp_path / "out.csv"))
    assert res["missing_input_warning"] is True and res["missing_input_counts"]["b"] == 10


def test_inline_over_five_records_saved_to_file(coded):
    recs = [{"a": float(i), "b": 1.0, "코드": "012"} for i in range(8)]
    js = coded["work"] / "many.json"
    js.write_text(json.dumps(recs), encoding="utf-8")
    res = run_tool("ml_predict.py", "--model", coded["v"], "--input", str(js))
    assert res["truncated"] is True and res["count"] == 8 and res["output_file"]
    assert os.path.dirname(res["output_file"]) == str(coded["work"]) and res["verified"] is True
    assert len(pd.read_csv(res["output_file"], encoding="utf-8-sig")) == 8
    one = run_tool("ml_predict.py", "--model", coded["v"], "--input", "a=1,b=1,코드=012")
    assert one["truncated"] is False and one["output_file"] is None


def test_inline_json_string_over_five_saved_in_data_root(coded):
    recs = json.dumps([{"a": float(i), "b": 1.0, "코드": "012"} for i in range(6)])
    res = run_tool("ml_predict.py", "--model", coded["v"], "--input", recs)
    assert res["truncated"] is True
    assert os.path.dirname(res["output_file"]) == os.path.dirname(coded["v"].rstrip("\\/"))


def test_truncated_inline_has_full_result_file(tiny, tmp_path):
    rows = [{"도핑농도": 2e-6 + i * 1e-6, "온도": 5.0} for i in range(7)]
    p = tmp_path / "in.json"
    p.write_text(json.dumps(rows), encoding="utf-8")
    res = run_tool("ml_predict.py", "--model", tiny["v"], "--input", str(p))
    assert res["truncated"] is True and res["full_result_file"] == res["output_file"]


def test_out_extension_checked_first(coded, tmp_path):
    bad = run_fail("ml_predict.py", "--model", coded["v"], "--file", coded["file"], "--save", str(tmp_path / "r.txt"))
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
    v = run_tool("ml_train.py", "--file", str(p), "--features", "a,b", "--targets", "y")["version_dir"]
    old = tmp_path / "v_old"
    shutil.copytree(v, old)
    meta_p = old / "model_y.json"
    meta = json.loads(meta_p.read_text(encoding="utf-8"))
    del meta["input_columns_required"]
    meta_p.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    bad = run_fail("ml_predict.py", "--model", str(old), "--input", "a=0.5,b=0.5")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "다시 학습" in bad["message"]
    res = run_tool("ml_train.py", "--file", str(p), "--history")
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
    res = run_tool("ml_train.py", "--file", str(path), "--features", "반응온도,압축압력", "--targets", "생산효율")
    return d, path, res["version_dir"]


def test_two_answers_windows_and_measured(hill):
    _, _, vdir = hill
    r = run_tool("ml_predict.py", "--model", vdir, "--find", "생산효율=105", "--vary", "반응온도", "--input", "압축압력=3")
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
    r = run_tool("ml_predict.py", "--model", vdir, "--find", "생산효율=112", "--vary", "반응온도", "--input", "압축압력=3")
    assert r["count"] == 0 and r["solutions"] == [] and r["reachable"]["max"] < 112
    assert "찾은 값 없음" in r["preview"] and "count 0" in r["agent_directive"]
    d = run_tool("ml_predict.py", "--model", vdir, "--find", "생산효율=105", "--vary", "반응온도")   # 압축압력 not given
    assert set(d["filled_defaults"]) == {"압축압력(bar)"} or set(d["filled_defaults"]) == {"압축압력"}
    assert "중앙값" in d["warnings"][0] and d["count"] >= 1


def test_out_of_range_fixed_input_warns(hill):
    _, _, vdir = hill
    r = run_tool("ml_predict.py", "--model", vdir, "--find", "생산효율=105", "--vary", "반응온도", "--input", "압축압력=9")
    assert any("학습 범위" in w for w in r["warnings"])


def test_inputs_errors(hill):
    d, path, vdir = hill
    cat = run_tool("ml_train.py", "--file", str(path), "--features", "반응온도,압축압력,라인", "--targets", "생산효율")
    miss = run_fail("ml_predict.py", "--model", cat["version_dir"], "--find", "생산효율=105", "--vary", "반응온도", "--input", "압축압력=3")
    assert "[후보] 1) A 2) B" in miss["message"]                                        # category never guessed
    ok = run_tool("ml_predict.py", "--model", cat["version_dir"], "--find", "생산효율=105", "--vary", "반응온도",
                  "--input", "압축압력=3,라인=A")
    assert ok["status"] == "ok"
    assert "숫자 입력이 아니라" in run_fail("ml_predict.py", "--model", cat["version_dir"], "--find", "생산효율=105",
                                           "--vary", "라인", "--input", "압축압력=3")["message"]
    cls = run_tool("ml_train.py", "--file", str(path), "--features", "반응온도,압축압력", "--targets", "등급", "--task", "classification")
    assert "분류" in run_fail("ml_predict.py", "--model", cls["version_dir"], "--find", "등급=1", "--vary", "반응온도")["message"]
    proc, payload = run_raw("ml_predict.py", "--model", vdir, "--find", "생산효율=105")                 # --vary missing
    assert proc.returncode == 2 and payload["status"] == "error"
    proc, payload = run_raw("ml_predict.py", "--model", vdir, "--find", "생산효율=105", "--vary", "반응온도", "--file", str(path))
    assert proc.returncode == 2


def test_linear_model_exact_answer(tmp_path):
    r = np.random.default_rng(7)
    n = 200
    t, p = r.uniform(170, 195, n), r.uniform(2.5, 3.5, n)
    path = tmp_path / "lin.csv"
    pd.DataFrame({"반응온도": t, "압축압력": p, "생산효율": 0.5 * t + 8 * p + r.normal(0, 0.05, n)}).to_csv(path, index=False)
    v = run_tool("ml_train.py", "--file", str(path), "--features", "반응온도,압축압력", "--targets", "생산효율")["version_dir"]
    res = run_tool("ml_predict.py", "--model", v, "--find", "생산효율=114", "--vary", "반응온도", "--input", "압축압력=3")
    assert res["count"] == 1 and res["solutions"][0]["반응온도"] == pytest.approx((114 - 24) / 0.5, abs=0.5)


def test_many_inputs_skip_measured_lookup(tmp_path):
    r = np.random.default_rng(8)
    n = 200
    df = pd.DataFrame({c: r.uniform(0, 1, n) for c in ("a", "b", "c", "d")})
    df["y"] = df.a * 10 + df.b + df.c + df.d
    path = tmp_path / "m.csv"
    df.to_csv(path, index=False)
    v = run_tool("ml_train.py", "--file", str(path), "--features", "a,b,c,d", "--targets", "y")["version_dir"]
    res = run_tool("ml_predict.py", "--model", v, "--find", "y=6", "--vary", "a", "--input", "b=0.5,c=0.5,d=0.5")
    assert "생략" in res["measured"]["note"] and len(res["nearest_rows"]) == 5
    assert res["count"] >= 1 and abs(res["solutions"][0]["a"] - 0.45) < 0.1


def test_find_with_text_inline_values(coded):
    res = run_tool("ml_predict.py", "--model", coded["v"], "--find", "y=30", "--vary", "a", "--input", "b=2,코드=007")
    assert res["fixed"] == {"b": 2.0, "코드": "007"} and res["count"] >= 1


def test_find_comma_number_stops_with_command(small):
    bad = run_fail("ml_predict.py", "--model", small["v"], "--find", "y=1,000", "--vary", "온도", "--input", "압력=3000,라인=A")
    assert '--find y=1000' in bad["message"].split("[추천] ")[1]
    bad = run_fail("ml_predict.py", "--model", small["v"], "--find", "y=120", "--vary", "온도", "--input", "압력=3,000,라인=A")
    rec = bad["message"].split("[추천] ")[1].strip()
    assert ' --input 압력=3000,라인=A' in rec
    assert run_tool("ml_predict.py", *split_command(rec)[2:])["command"] == "find"


def test_find_multi_row_input_refused(p1_model):
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--find", "생산량=530", "--vary", "온도",
                        "--input", '[{"압력":2.2},{"압력":3.8}]')
    assert "한 줄" in stop_msg(res, proc)
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--find", '[{"생산량":530},{"생산량":520}]',
                        "--vary", "온도")
    assert "한 줄" in stop_msg(res, proc)
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--find", '[{"생산량":530}]', "--vary", "온도",
                        "--input", '[{"압력":2.2}]')
    assert res["status"] == "ok" and res["fixed"] == {"압력(bar)": 2.2}, res


def test_nearest_rows_are_numbers_and_find_rejects_text_input(tmp_path):
    r = np.random.default_rng(3)
    df = pd.DataFrame({"온도": r.uniform(150, 200, 200), "압력": r.uniform(1, 3, 200)})
    df["효율"] = 0.4 * df.온도 + 5 * df.압력 + r.normal(0, 0.2, 200)
    p = tmp_path / "n.csv"
    df.to_csv(p, index=False)
    m = run_tool("ml_map.py", "--file", str(p), "--find", "효율=80", "--vary", "온도", "--input", "압력=2")
    assert m["nearest_rows"] and all(isinstance(nr["온도"], float) for nr in m["nearest_rows"])
    v = run_tool("ml_train.py", "--file", str(p), "--features", "온도,압력", "--targets", "효율")["version_dir"]
    f = run_tool("ml_predict.py", "--model", v, "--find", "효율=80", "--vary", "온도", "--input", "압력=2")
    assert all(isinstance(nr["온도"], float) for nr in f["nearest_rows"])
    bad = run_fail("ml_predict.py", "--model", v, "--find", "효율=80", "--vary", "온도", "--input", "압력=x")
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
    res = run_tool("ml_train.py", "--file", str(data), "--features", "반응온도,압축압력", "--targets", "생산효율")
    return d, df, res["version_dir"]


def test_snapshot_has_rows_and_exclusions(mc_model):
    _, _, vdir = mc_model
    snap = pd.read_csv(os.path.join(vdir, "model_dataset.csv"), encoding="utf-8-sig")
    assert list(snap["원본행"][:3]) == [2, 3, 4]                       # Excel rows (header is row 1)
    assert snap.loc[snap["원본행"] == 12, "학습제외_생산효율"].item() == 1   # '센서오류' on data row 10 = Excel row 12


def test_inline_prediction_has_map_check(mc_model):
    _, df, vdir = mc_model
    res = run_tool("ml_predict.py", "--model", vdir, "--input", "반응온도=180,압축압력=3")
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
    res = run_tool("ml_predict.py", "--model", vdir, "--file", str(new))
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
    res = run_tool("ml_predict.py", "--model", str(v2), "--input", "반응온도=180,압축압력=3")
    assert res["map_check_warning"] and "90% 예측 범위 밖" in res["map_check_warnings"][0]
    assert res["predictions_summary"]["생산효율"]["map_check"][0]["outside_interval"]
    assert "실측 비교 경고" in res["preview"]


def test_old_snapshot_and_missing_snapshot(mc_model, tmp_path):
    _, _, vdir = mc_model
    old = tmp_path / "v_old"
    shutil.copytree(vdir, old)
    snap = pd.read_csv(old / "model_dataset.csv", encoding="utf-8-sig").drop(columns=["원본행", "학습제외_생산효율"])
    snap.to_csv(old / "model_dataset.csv", index=False, encoding="utf-8-sig")
    res = run_tool("ml_predict.py", "--model", str(old), "--input", "반응온도=180,압축압력=3")     # snapshot from before 2.7.0
    assert "map_check" not in res["predictions_summary"]["생산효율"] and "다시 학습" in res["map_check_notes"][0]

    os.remove(old / "model_dataset.csv")
    res = run_tool("ml_predict.py", "--model", str(old), "--input", "반응온도=180,압축압력=3")
    assert "map_check" not in res["predictions_summary"]["생산효율"] and "model_dataset.csv" in res["map_check_notes"][0]
    assert res["predictions_summary"]["생산효율"]["predictions"]           # prediction itself unaffected


def test_many_inputs_use_nearest_rows_and_category(tmp_path):
    data = tmp_path / "d.xlsx"
    _mc_data(data)
    res = run_tool("ml_train.py", "--file", str(data), "--features", "반응온도,압축압력,촉매비율,수분함량,라인", "--targets", "생산효율")
    out = run_tool("ml_predict.py", "--model", res["version_dir"], "--input", "반응온도=180,압축압력=3,촉매비율=2.5,수분함량=1,라인=A")
    mc = out["predictions_summary"]["생산효율"]["map_check"][0]
    assert mc["method"].startswith("가까운 실측") and 1 <= len(mc["basis"]) <= 5
    assert sum(b["weight"] for b in mc["basis"]) == pytest.approx(1.0, abs=1e-4)
    assert all(b["row"] % 2 == 1 for b in mc["basis"])                  # 라인 A = odd data rows = odd Excel rows (row = i + 2)


def test_classification_has_no_map_check_and_standalone(mc_model, tmp_path):
    d, _, vdir = mc_model
    cls = run_tool("ml_train.py", "--file", str(d / "d.xlsx"), "--features", "반응온도,압축압력", "--targets", "등급", "--task", "classification")
    res = run_tool("ml_predict.py", "--model", cls["version_dir"], "--input", "반응온도=180,압축압력=3")
    assert "map_check" not in res["predictions_summary"]["등급"] and not res["map_check_warning"]

    deploy = tmp_path / "deploy"
    deploy.mkdir()
    for f in ("ml_predict.py", "_ml_runtime.py"):
        shutil.copy(os.path.join(KIT, f), deploy / f)
    shutil.copytree(vdir, deploy / "v1")
    import subprocess
    import json
    proc = subprocess.run([sys.executable, str(deploy / "ml_predict.py"), "--model", str(deploy / "v1"), "--input", "반응온도=180,압축압력=3"],
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
    monkeypatch.setattr(mp, "model_raw_inputs", lambda m: list("abcd"))   # compare_with_measured (ml_predict)
    rows = df[list("abcd")].iloc[:2000].reset_index(drop=True) + 0.001

    calls = {"n": 0}
    real = rt.coerce_numeric

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
    v = run_tool("ml_train.py", "--file", data, "--features", "반응온도,압축압력,촉매비율", "--targets", "생산효율,불순물")["version_dir"]
    return {"v": v, "work": work}


def test_sweep_table_chart_trend(whatif):
    res = run_tool("ml_predict.py", "--model", whatif["v"], "--sweep", "반응온도=170:200:5", "--input", "압축압력=3,촉매비율=1")
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
    res = run_tool("ml_predict.py", "--model", whatif["v"], "--sweep", "반응온도=160:210:10", "--targets", "생산효율")
    assert res["targets"] == ["생산효율"] and res["out_of_range_points"] == 2
    assert [p["in_range"] for p in res["points"]] == [False, True, True, True, True, False]
    assert any("외삽" in w for w in res["warnings"]) and "중앙값" in res["warnings"][0]
    assert set(res["filled_defaults"]) == {"압축압력", "촉매비율"}


def test_sweep_caps_points(whatif):
    bad = run_fail("ml_predict.py", "--model", whatif["v"], "--sweep", "반응온도=170:200:0.01", "--input", "압축압력=3,촉매비율=1")
    rec = bad["message"].split("[추천] ")[1]
    assert rec.startswith("python ") and "--sweep" in rec
    res = run_tool("ml_predict.py", *split_command(rec)[2:])
    assert 25 < res["count"] <= 500 and res["truncated"] is True and res["shown"] == len(res["points"]) == 25
    assert res["full_result_file"] == res["output_file"] and len(pd.read_excel(res["output_file"])) == res["count"]
    assert res["points"][0]["반응온도"] == 170 and res["points"][-1]["반응온도"] == 200


def test_maximize_with_constraint(whatif):
    res = run_tool("ml_predict.py", "--model", whatif["v"], "--maximize", "생산효율", "--where", "불순물 <= 80",
                   "--vary", "반응온도,압축압력", "--input", "촉매비율=1")
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
    res = run_tool("ml_predict.py", "--model", whatif["v"], "--minimize", "불순물", "--where", "생산효율 > 1000",
                   "--vary", "반응온도", "--input", "압축압력=3,촉매비율=1")
    assert res["count"] == 0 and res["solutions"] == [] and "best_without_where" in res
    assert res["agent_directive"].startswith("VERIFIED: 0건")
    assert abs(res["best_without_where"]["inputs"]["반응온도"] - 185) <= 3


def test_whatif_argument_rules(whatif):
    proc, res = run_raw("ml_predict.py", "--model", whatif["v"], "--where", "불순물 <= 80", "--input", "반응온도=180")
    assert proc.returncode == 2, res
    proc, res = run_raw("ml_predict.py", "--model", whatif["v"], "--sweep", "반응온도", "--find", "생산효율=100", "--vary", "반응온도")
    assert proc.returncode == 2, res
    bad = run_fail("ml_predict.py", "--model", whatif["v"], "--maximize", "생산효율", "--vary", "반응온도", "--where", "불순물 == 80")
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
    return run_tool("ml_train.py", "--file", data, "--features", "온도,라인", "--targets", "수율")["version_dir"]


def test_whatif_unknown_category_stops_with_candidates(line_model):
    """2.9.59: --input "라인=라인C"를 학습 범주와 대조하지 않아 기준값으로 예측하고 VERIFIED였다."""
    for opt in (["--sweep", "온도"], ["--maximize", "수율", "--vary", "온도"], ["--find", "수율=60", "--vary", "온도"]):
        bad = run_fail("ml_predict.py", "--model", line_model, *opt, "--input", "라인=라인C")
        assert bad["agent_directive"].startswith("STOP_RETRY"), bad
        assert "[후보] 1) A 2) B 3) C" in bad["message"] and "라인C" in bad["message"], bad["message"]


def test_whatif_category_case_width_space_recovered(line_model):
    ref = run_tool("ml_predict.py", "--model", line_model, "--sweep", "온도=200:280:40", "--input", "라인=C")
    for given in ("c", " C ", "Ｃ"):              # 소문자, 앞뒤 공백, 전각
        res = run_tool("ml_predict.py", "--model", line_model, "--sweep", "온도=200:280:40", "--input", f"라인={given}")
        assert res["fixed"] == {"라인": "C"}, res
        assert [p["수율"] for p in res["points"]] == [p["수율"] for p in ref["points"]]
        if given != " C ":                            # 앞뒤 공백은 읽을 때 이미 지워져 바뀐 것이 없다
            assert any("'C'" in m for m in res.get("recovered", [])), res
    res = run_tool("ml_predict.py", "--model", line_model, "--maximize", "수율", "--vary", "온도", "--input", "라인=b")
    assert res["fixed"] == {"라인": "B"} and res.get("recovered"), res
    res = run_tool("ml_predict.py", "--model", line_model, "--find", "수율=60", "--vary", "온도", "--input", "라인=a")
    assert res.get("recovered") and res["fixed"] == {"라인": "A"}, res


def test_predict_category_case_matched_like_whatif(line_model):
    """일반 예측 --input도 같은 공통 함수: 대소문자만 다르면 맞추고 recovered, 그 밖은 unseen_category_warnings."""
    ref = run_tool("ml_predict.py", "--model", line_model, "--input", "온도=250,라인=C")
    res = run_tool("ml_predict.py", "--model", line_model, "--input", "온도=250,라인=c")
    assert res["unseen_category_warning"] is False and res.get("recovered"), res
    assert res["predictions_summary"]["수율"]["predictions"] == ref["predictions_summary"]["수율"]["predictions"]
    res = run_tool("ml_predict.py", "--model", line_model, "--input", "온도=250,라인=라인C")
    assert res["unseen_category_warning"] is True and "recovered" not in res, res


def test_match_categories_ambiguous_not_fixed():
    fixes, unseen = rt.match_categories(["a", "b ", "x"], ["A", "a", "B"])
    assert fixes == {"b ": "B"} and unseen == ["x"]      # 'a'는 학습 범주 그대로 ('A'로 바꾸지 않음)
    fixes, unseen = rt.match_categories(["ａ"], ["A", "a"])
    assert fixes == {} and unseen == ["ａ"]          # 'A'와 'a' 둘 다 맞아 고를 수 없음


def test_model_given_as_data_file_uses_best_version(p1_model, tmp_path):
    """2.9.71: --model <데이터 파일>이면 도구가 best_version.json(학습이 원장마다 씀)의 최고 버전으로 예측 (LLM이 버전 폴더를 고르지 않음)."""
    root = os.path.join(os.path.dirname(p1_model["f"]), "p1_ml_data")
    pointer = json.load(open(os.path.join(root, "best_version.json"), encoding="utf-8"))
    assert pointer["tasks"]["regression"]["version"] == "v1", pointer
    res = run_tool("ml_predict.py", "--model", p1_model["f"], "--input", "온도=180,압력=3,속도=45")
    by_dir = run_tool("ml_predict.py", "--model", p1_model["v"], "--input", "온도=180,압력=3,속도=45")
    assert res["predictions_summary"] == by_dir["predictions_summary"] and "v1" in res["model_chosen"], res
    assert "model_chosen" not in by_dir
    # 기록이 없으면 고르지 않고 버전 폴더를 [후보] 전체 명령으로
    copy = tmp_path / "p1_ml_data"
    shutil.copytree(root, copy)
    os.remove(copy / "best_version.json")
    shutil.copy(p1_model["f"], tmp_path / "p1.csv")
    err = run_fail("ml_predict.py", "--model", str(tmp_path / "p1.csv"), "--input", "온도=180,압력=3,속도=45")
    assert "[후보] 1) v1" in err["message"] and err["error_type"] == "needs_user_decision", err
    argv = split_command(err["confirm_commands"][0]["command"])
    assert argv[0] == "python" and argv[argv.index("--model") + 1] == str(copy / "v1"), err


# ---- 2.9.72: 다음 실험 제안(--suggest), 입력별 영향(explain), 실제값 비교와 재학습 알림(actual_check) -------------


@pytest.fixture(scope="module")
def lin(tmp_path_factory):
    """수율 = 3*온도 - 2*압력 (+ 속도 영향 없음) 회귀 모델과, 새 데이터 두 개 (같은 관계 / 바뀐 관계)."""
    tmp = str(tmp_path_factory.mktemp("lin"))
    r = np.random.default_rng(1)
    n, m = 80, 30
    a, b, c = r.uniform(0, 10, n), r.uniform(0, 10, n), r.uniform(0, 10, n)
    y = 3 * a - 2 * b + r.normal(0, 0.3, n)
    f = write_df(os.path.join(tmp, "lin.csv"), pd.DataFrame({"온도": a.round(3), "압력": b.round(3), "속도": c.round(3),
                                                             "수율": y.round(3)}))
    run_tool("ml_train.py", "--file", f, "--features", "온도,압력,속도", "--targets", "수율")
    a2, b2, c2 = r.uniform(1, 9, m), r.uniform(1, 9, m), r.uniform(1, 9, m)
    same = write_df(os.path.join(tmp, "new_same.csv"), pd.DataFrame(
        {"온도": a2, "압력": b2, "속도": c2, "수율": 3 * a2 - 2 * b2 + r.normal(0, 0.3, m)}))
    shift = write_df(os.path.join(tmp, "new_shift.csv"), pd.DataFrame(
        {"온도": a2, "압력": b2, "속도": c2, "수율": 3 * a2 - 2 * b2 + 5 + c2 + r.normal(0, 0.3, m)}))
    return {"f": f, "same": same, "shift": shift, "tmp": tmp, "ys": y}


def _in_range(res):
    for p in res["points"]:
        for v, x in p["inputs"].items():
            lo, hi = res["training_range"][v]
            assert lo <= x <= hi, (v, x, lo, hi)


def test_suggest_explore_in_range_and_saved(lin):
    import openpyxl
    res = run_tool("ml_predict.py", "--model", lin["f"], "--suggest", "4")
    assert res["mode"] == "explore" and res["count"] == 4 and res["verified"] is True
    assert res["vary"] == ["온도", "압력", "속도"]
    _in_range(res)
    pts = [tuple(p["inputs"].values()) for p in res["points"]]
    assert len(set(pts)) == 4
    assert all(p["nearest_distance"] > 0 and p["interval_width"]["수율"] > 0 for p in res["points"])
    wb = openpyxl.load_workbook(res["output_file"], read_only=True)
    assert wb.sheetnames == ["실험계획"]
    rows = list(wb["실험계획"].iter_rows(values_only=True))
    wb.close()
    assert len(rows) == 5 and rows[0][:4] == ("번호", "온도", "압력", "속도")
    assert "다음 실험 제안 4개" in res["preview"] and res["agent_directive"].startswith("VERIFIED:")


def test_suggest_spreads_away_from_measured_rows(lin):
    """제안 점은 실측 행끼리의 보통 거리보다 멀리 떨어진 곳 (빈 곳을 채움)."""
    res = run_tool("ml_predict.py", "--model", lin["f"], "--suggest", "3", "--vary", "온도", "--input", "압력=3")
    assert res["vary"] == ["온도"] and res["fixed"] == {"압력": 3.0} and "속도" in res["filled_defaults"]
    assert any("중앙값" in w for w in res["warnings"])
    xs = sorted(p["inputs"]["온도"] for p in res["points"])
    assert min(np.diff(xs)) > 0.5
    _in_range(res)


def test_suggest_maximize_prefers_good_values(lin):
    res = run_tool("ml_predict.py", "--model", lin["f"], "--suggest", "3", "--maximize", "수율")
    assert res["mode"] == "maximize" and res["goal"] == "수율" and res["verified"] is True
    _in_range(res)
    assert min(p["predictions"]["수율"] for p in res["points"]) > np.quantile(lin["ys"], 0.9)
    assert "수율 최대" in res["preview"]


def test_suggest_count_limit_recommends_full_command(lin):
    proc, res = run_raw("ml_predict.py", "--model", lin["f"], "--suggest", "50")
    msg = stop_msg(res, proc)
    rec = msg.split("[추천] ")[1].strip()
    assert rec.startswith("python") and split_command(rec)[-2:] == ["--suggest", "30"]


def test_explain_signs_on_linear_model(lin):
    res = run_tool("ml_predict.py", "--model", lin["f"], "--input", "온도=8,압력=6,속도=5")
    ex = res["predictions_summary"]["수율"]["explain"][0]
    eff = {e["input"]: e["effect"] for e in ex["effects"]}
    assert eff["온도"] > 0 and eff["압력"] < 0
    assert abs(eff["속도"]) < 0.1 * abs(eff["온도"]) and abs(eff["속도"]) < 0.1 * abs(eff["압력"])
    assert [e["input"] for e in ex["effects"]][:1] == ["온도"]                 # 가장 큰 영향이 먼저
    med = {e["input"]: e["baseline"] for e in ex["effects"]}
    assert abs(eff["온도"] - 3 * (8 - med["온도"])) < 1.0 and abs(eff["압력"] + 2 * (6 - med["압력"])) < 1.0
    assert "높인 입력: 온도" in res["preview"] and "낮춘 입력: 압력" in res["preview"]
    assert "입력별 영향" in res["agent_directive"] and "입력별 영향" in res["preview"]     # 2.9.78: 알릴 글은 preview에
    # 2.9.77 집 PC Gemma: '반응온도(↓)'처럼 화살표로 줄여 써 채점 실패 -> 글 그대로 옮기라고 지시
    assert "글 그대로" in res["agent_directive"] and "기호로 바꾸지 않음" in res["agent_directive"]


def test_explain_classification_probability(lin):
    tmp = lin["tmp"]
    df = pd.read_csv(lin["f"], encoding="utf-8-sig")
    df["등급"] = np.where(df["수율"] > df["수율"].median(), "상", "하")
    f = write_df(os.path.join(tmp, "lin_cls.csv"), df.drop(columns="수율"))
    run_tool("ml_train.py", "--file", f, "--features", "온도,압력,속도", "--targets", "등급", "--task", "classification")
    res = run_tool("ml_predict.py", "--model", f, "--input", "온도=9,압력=1,속도=5")
    ps = res["predictions_summary"]["등급"]
    assert ps["predictions"] == ["상"]
    ex = ps["explain"][0]
    assert ex["class"] == "상"
    eff = {e["input"]: e["effect"] for e in ex["effects"]}
    assert eff["온도"] > 0 and eff["압력"] > 0                                  # 둘 다 '상' 확률을 높임
    assert "'상' 확률" in res["preview"]


def test_actual_check_similar_data_no_alert(lin):
    res = run_tool("ml_predict.py", "--model", lin["f"], "--file", lin["same"])
    chk = res["actual_check"]["수율"]
    assert chk["rows"] == 30 and chk["judged"] is True and chk["worse"] is False
    assert chk["rmse"] < 1.5 * chk["validation"]["rmse"]
    assert res["retrain_recommended"] is False and "next_commands" not in res
    assert "실제값 비교" in res["preview"] and "(비슷함)" in res["preview"]


def test_actual_check_shifted_data_alerts_and_retrains(lin):
    res = run_tool("ml_predict.py", "--model", lin["f"], "--file", lin["shift"])
    chk = res["actual_check"]["수율"]
    assert chk["worse"] is True and res["retrain_recommended"] is True
    assert any("다시 학습" in w for w in res["warnings"])
    assert "next_commands[0]" in res["agent_directive"] and "재학습 권장" in res["preview"]
    nc = res["next_commands"][0]
    assert nc["needs_input"] is False
    args = split_command(nc["command"])
    assert os.path.basename(args[1]) == "ml_train.py"
    assert args[args.index("--file") + 1] == os.path.abspath(lin["shift"]) and args[args.index("--features") + 1] == "온도,압력,속도"
    again = run_tool("ml_train.py", *args[2:])
    assert again["targets"] == ["수율"] and again["verified"] is True
    assert os.path.dirname(again["data_root"]) == os.path.dirname(os.path.abspath(lin["shift"]))


def test_actual_check_few_rows_not_judged(lin, tmp_path):
    df = pd.read_csv(lin["shift"], encoding="utf-8-sig").head(5)
    f = write_df(os.path.join(str(tmp_path), "few.csv"), df)
    res = run_tool("ml_predict.py", "--model", lin["f"], "--file", f)
    chk = res["actual_check"]["수율"]
    assert chk["judged"] is False and chk["worse"] is False and "note" in chk and "next_commands" not in res


def test_one_case_preview_has_everything_to_tell_and_directive_is_short(p1_model):
    """2.9.78 집 PC r5: 지시문이 길어(약 350자) Gemma가 뒤쪽 '실측 비교값 함께'를 놓침 -> 건별 예측값·90% 범위·실측 비교값은 preview에,
    지시문은 'preview 그대로' 짧게 (높인/낮춘 입력 글 그대로는 유지)."""
    res = run_tool("ml_predict.py", "--model", p1_model["v"], "--input", "온도(°C)=180,압력(bar)=3,속도=45")
    pv, d = res["preview"], res["agent_directive"]
    assert "생산량 " in pv and "90% 범위 " in pv and "실측 비교값 " in pv, pv
    assert len(d) < 260 and d.startswith("VERIFIED: 타겟 1개를 예측했습니다. preview의 예측값, 90% 범위, 실측 비교값"), d
    assert "글 그대로" in d and "화살표" in d, d
    two = run_tool("ml_predict.py", "--model", p1_model["v"], "--input", '[{"온도(°C)":180,"압력(bar)":3,"속도":45},{"온도(°C)":185,"압력(bar)":3,"속도":45}]')
    assert "생산량 1번 " in two["preview"] and "생산량 2번 " in two["preview"], two["preview"]
