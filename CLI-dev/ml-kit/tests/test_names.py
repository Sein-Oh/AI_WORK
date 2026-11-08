# -*- coding: utf-8 -*-
"""열 이름 맞추기 (모든 도구): 단위, 별칭, 오타, 다른 역할의 이름, ml_map --points 열.

규칙: 비슷한 이름은 [추천]이 아니라 [후보]다 (에이전트가 [추천]을 스스로 재시도해 '속도' 값이 '가속도'에 들어갈 수 있다).
같은 열을 다르게 적은 것(대소문자, 띄어쓰기, 괄호 모양, '°C'/'℃'/'C')만 [INFO]와 함께 자동으로 맞춘다.
단위가 다른 이름은 값 환산이 필요하므로 [추천] 없이 멈춘다. 이미 다른 이름이나 옵션(--targets, --vary, --find)에 쓰인 열은 후보가 아니다.
"""
import json
import os
import random
import sys

import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import data_file, data_path, ok_res, run_raw, run_tool, split_command, stop_msg, write_rows  # noqa: E402
import ml_export  # noqa: E402
import ml_map  # noqa: E402
import ml_predict as mp  # noqa: E402
import _ml_runtime as rt  # noqa: E402
import ml_report  # noqa: E402


@pytest.fixture(scope="module")
def two_model(tmp_path_factory):
    """온도, 압력 -> 수율(1차), 수율(최종): 단위를 빼면 이름이 같은 두 타겟의 모델."""
    tmp = tmp_path_factory.mktemp("two")
    r = random.Random(5)
    rows = [["온도", "압력", "수율(1차)", "수율(최종)"]]
    for _ in range(80):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 3), 2)
        rows.append([t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2), round(0.2 * t + p + r.uniform(-1, 1), 2)])
    f = write_rows(os.path.join(str(tmp), "two.csv"), rows, [])
    _, res = run_raw("ml_train.py", "--file", f, "--features", "온도,압력", "--targets", "수율(1차),수율(최종)")
    assert res["status"] == "ok", res
    return os.path.join(str(tmp), "two_ml_data", "v1")


# ---- 규칙: resolve_column_names, 단위, 오타 ------------------------------------------------


def test_ambiguous_alias_not_resolved():
    with pytest.raises(rt.InputError):
        rt.resolve_column_names(["온도"], ["온도(°C)", "온도(K)"], "피처")


def _resolve_error(names, cols, **kw):
    with pytest.raises(rt.InputError) as e:
        rt.resolve_column_names(names, cols, "피처(--features)", **kw)
    return str(e.value)


def test_no_sure_typo_function():
    assert not hasattr(mp, "sure_typo") and not hasattr(mp, "one_or_ask")


@pytest.mark.parametrize("given,cols", [
    ("속도", ["가속도", "압력(bar)"]), ("수율", ["함수율(%)", "압력(bar)"]), ("속두", ["속도", "압력"]),
    ("압룍", ["압력(bar)", "온도"]), ("온도x", ["온도(°C)", "압력"]), ("tmep", ["temp", "press"]),
    ("압력", ["반응압력(bar)", "온도"]), ("생산량(kg)", ["생산량", "온도"]), ("온돜", ["온도", "압력"]),
])
def test_fuzzy_name_is_a_candidate(given, cols):
    msg = _resolve_error([given], cols)
    assert "[추천]" not in msg and f"[후보] 1) '{cols[0]}'" in msg and "사용자에게 물으십시오" in msg, msg


@pytest.mark.parametrize("given,col", [
    ("온도(C)", "온도(°C)"), ("온도(℃)", "온도(°C)"), ("TEMP(°C)", "temp(℃)"), ("압력 [bar]", "압력(bar)"),
    ("수율( 1차)", "수율(1차)"), ("반응 온도(C)", "반응온도(°C)"), ("생산효율", "생산효율(%)"),
])
def test_same_column_written_otherwise_matches(given, col, capsys):
    got, applied = rt.resolve_column_names([given], [col, "압력(MPa)"], "피처(--features)")
    assert got == [col] and applied == {given: col}
    assert "[INFO] 컬럼명 자동 매칭" in capsys.readouterr().err


def test_equivalence_is_not_unit_case_or_ambiguous():
    msg = _resolve_error(["질량(mg)"], ["질량(Mg)"])                  # unit case counts: another unit
    assert "[추천]" not in msg and "[후보]" in msg, msg
    msg = _resolve_error(["온도(C)"], ["온도(°C)", "온도(℃)"])         # two equivalent columns: the user picks
    assert "[추천]" not in msg and "[후보] 1) '온도(°C)' 2) '온도(℃)'" in msg, msg


@pytest.mark.report
def test_excluded_column_never_listed():
    msg = _resolve_error(["수율"], ["온도", "수율(1차)", "수율(최종)"], exclude=["수율(최종)"])
    head = msg.split("(비슷한")[0]
    assert "[후보] 1) '수율(1차)'" in head and "수율(최종)" not in head and "[추천]" not in msg, msg
    assert "이 명령에서 이미 다른 이름이나 옵션에 쓰인 열" in msg, msg
    with pytest.raises(rt.InputError) as e:
        ml_export._check_known(["압력", "압룍"], {"y": ["압력", "온도"]}, "꼭 포함할 입력(--keep)")
    msg = str(e.value)
    assert "[추천]" not in msg and "[후보] '압력'" not in msg and "이미 다른 이름이나 옵션에 쓰인 열" in msg, msg


OTHER_WORDS = [("습도", "속도"), ("온도", "농도"), ("T1", "T2"), ("온도1", "온도2"), ("반응온도", "반응농도"),
               ("Temp", "Time"), ("CO2", "O2"), ("pH", "H"), ("Ax", "A")]


# 2.9.33부터 sure_typo가 없다. 비슷한 이름은 모두 [후보]다 (아래 오타는 전에 [추천]이었음). 같은 열을 다르게 적은 것
# ('반응 온도', '반응온도(C)' -> '반응온도(°C)')은 맞춘다. 속돜, 압력(bar), 온도x, tmep는 SURE 목록이 같은 입력으로 확인한다.


TYPOS = [("도핑농돜", "도핑농도"), ("생산량(kg)", "생산량")]


@pytest.mark.parametrize("a,b", OTHER_WORDS)
def test_other_word_is_a_candidate(a, b):
    with pytest.raises(rt.InputError) as e:
        rt.resolve_column_names([a], [b, "생산량"], "피처(--features)")
    assert "[추천]" not in str(e.value) and f"[후보] 1) '{b}'" in str(e.value), str(e.value)


@pytest.mark.parametrize("a,b", TYPOS)
def test_typo_is_a_candidate(a, b):
    # 2.9.33: a typo is a [후보] too (was [추천])
    with pytest.raises(rt.InputError) as e:
        rt.resolve_column_names([a], [b, "zzz"], "피처(--features)")
    assert "[추천]" not in str(e.value) and f"[후보] 1) '{b}'" in str(e.value), str(e.value)


def test_typo_resolve_messages():
    for a, b in (("반응 온도", "반응온도"), ("반응온도(C)", "반응온도(°C)")):        # 2.9.33: matched (was [추천])
        assert rt.resolve_column_names([a], [b, "생산량"], "피처(--features)")[0] == [b]


ASK = [("속두", "속도"), ("습두", "습도"), ("온두", "온도"), ("도핑농두", "도핑농도"), ("강도", "경도"), ("점도", "정도"),
       ("생산랑", "생산량"), ("TempA", "TempB"), ("온도A", "온도B"), ("Heat", "Head"), ("T1", "X1"), ("O2", "T2"),
       ("Ti", "Tim"), ("Ti", "Time"), ("온도", "온도차"), ("설정온도", "온도"), ("반응온도", "온도(°C)"),
       ("수율(1차)", "수율(최종)"), ("생산량(kg)", "생산량(t)"), ("압력(MPa)", "반응압력(bar)")]


# 2.9.33: sure_typo is gone; these were [추천] and are [후보] now ('반응온도(C)' / '반응온도(°C)' matches instead)
SURE = [("압룍", "압력"), ("속돜", "속도"), ("온도x", "온도"), ("tmep", "temp"), ("압력", "반응압력"),
        ("압력(bar)", "반응압력(bar)"), ("Prssure", "Pressure"), ("수율( 1차)", "수율")]


@pytest.mark.parametrize("a,b", ASK)
def test_close_word_is_a_candidate(a, b):
    msg = _resolve_error([a], [b, "zzz"])
    assert "[추천]" not in msg and "[후보]" in msg, msg


@pytest.mark.parametrize("a,b", SURE)
def test_former_sure_typo_is_a_candidate(a, b):
    msg = _resolve_error([a], [b, "zzz"])
    assert "[추천]" not in msg and f"[후보] 1) '{b}'" in msg, msg


def test_known_columns():
    assert rt.known_columns(["온도", "압력( bar)", "없음"], ["온도(°C)", "압력(bar)"]) == ["온도(°C)", "압력(bar)"]


def test_unitless_name_against_unit_columns():
    msg = _resolve_error(["온두"], ["온도(°C)", "압력(bar)"])
    assert "[추천]" not in msg and "[후보] 1) '온도(°C)'" in msg, msg           # 2.9.31: 유사 컬럼 없음
    assert "[후보] 1) '온도(°C)'" in _resolve_error(["온돜"], ["온도(°C)", "압력(bar)"])       # 2.9.33: [추천]
    hint = mp._typo_hints(["온도(°C)"], ["온두"], ["온도(°C)"], True)
    assert len(hint) == 1 and "[후보] 1) 온두" in hint[0], hint


def test_same_name_candidates_whatever_the_suffix():
    # difflib alone misses a long unit suffix ('수율' / '수율(최종공정)' ratio 0.4)
    got = rt.suggest_columns("수율", ["수율(1차)", "수율(최종공정)", "온도"])
    assert set(got[:2]) == {"수율(1차)", "수율(최종공정)"}, got
    with pytest.raises(rt.InputError) as e:
        rt.resolve_column_names(["수율"], ["수율(1차)", "수율(최종공정)", "온도"], "목표(--find)", values=True)
    assert "[후보]" in str(e.value) and "수율(최종공정)" in str(e.value)


def test_other_unit_rules():
    assert rt.other_unit("압력(MPa)", "압력(bar)") and rt.other_unit("온도(°F)", "온도(°C)")
    assert rt.other_unit("무게(mg)", "무게(Mg)")                    # case is a different unit
    for a, b in (("온도(C)", "온도(°C)"), ("온도(℃)", "온도(°C)"), ("압력 [bar]", "압력( bar )"), ("온도", "온도(°C)")):
        assert not rt.other_unit(a, b), (a, b)


def test_other_unit_rule():
    assert rt.other_unit("압력(MPa)", "반응압력(bar)") and rt.other_unit("반응압력(bar)", "압력(MPa)")
    assert not rt.other_unit("압력(bar)", "반응압력(bar)") and not rt.other_unit("압력(MPa)", "반응압력")
    assert not rt.other_unit("온도(C)", "반응온도(℃)")


# ---- 학습 --features, --targets, --positive-class, --eval 이름 ------------------------------------


def test_train_accel_and_moisture_are_candidates(name_files):
    """'속도' -> '가속도', '수율' -> '함수율(%)'은 [후보]다 (2.9.32의 [추천]은 재시도에서 다른 양을 학습했음)."""
    proc, res = run_raw("ml_train.py", "--file", name_files["acc"], "--features", "속도,압력", "--targets", "진동")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '가속도'" in msg, msg          # 2.9.71: 번호 + next_commands
    proc, res = run_raw("ml_train.py", "--file", name_files["acc"], "--features", "가속도,압력", "--targets", "수율")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '함수율(%)'" in msg, msg


def test_degree_spelling_matches_in_train(name_files):
    proc, res = run_raw("ml_train.py", "--file", name_files["d"], "--features", "온도(C),압력(bar)", "--targets", "수율(1차)")
    ok_res(res, proc)


@pytest.fixture(scope="module")
def role_data(tmp_path_factory):
    """두 타겟 파일(two), 습도가 있는 공정 파일(p)과 그 모델(pm), 이진 분류 파일(c)."""
    tmp = tmp_path_factory.mktemp("roles")
    r = random.Random(3)
    two = [["온도", "압력", "수율(1차)", "수율(최종)"]]
    for _ in range(60):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 3), 2)
        two.append([t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2), round(0.2 * t + p + r.uniform(-1, 1), 2)])
    proc = [["온도(°C)", "압력(bar)", "속도", "습도", "생산량"]]
    for _ in range(60):
        t, p, s, h = (round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1),
                      round(r.uniform(30, 70), 1))
        proc.append([t, p, s, h, round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    cls = [["온도", "압력", "판정"]]
    for i in range(40):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2)
        cls.append([t, p, "불량" if t + 10 * p > 215 or i % 7 == 0 else "양품"])
    out = {"tmp": tmp, "two": write_rows(os.path.join(str(tmp), "two.csv"), two, []),
           "p": write_rows(os.path.join(str(tmp), "p.csv"), proc, []),
           "c": write_rows(os.path.join(str(tmp), "c.csv"), cls, [])}
    _, res = run_raw("ml_train.py", "--file", out["p"], "--features", "온도(°C),압력(bar),속도,습도", "--targets", "생산량")
    assert res["status"] == "ok", res
    out["pm"] = os.path.join(str(tmp), "p_ml_data", "v1")
    return out


def test_train_x_typo_never_suggests_the_target(role_data):
    proc, res = run_raw("ml_train.py", "--file", role_data["p"], "--features", "온도,압력,생산랑", "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "'생산량'" in msg, msg
    proc, res = run_raw("ml_train.py", "--file", role_data["two"], "--features", "온도,압력,수율", "--targets", "수율(최종)")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '수율(1차)' (비슷한 '수율(최종)'" in msg, msg        # 2.9.33: no --targets column
    # the reverse: a --targets typo never suggests a feature
    proc, res = run_raw("ml_train.py", "--file", role_data["p"], "--features", "온도,압력", "--targets", "압룍")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg, msg


def test_eval_never_suggests_a_target():
    df = pd.DataFrame({"온도": [1.0, 2.0], "압력": [1.0, 2.0], "생산량": [3.0, 4.0]})
    with pytest.raises(rt.InputError) as e:
        rt.apply_eval_expressions(df, ["d = 생산랑 * 2"], target_cols=["생산량"])
    assert "[추천]" not in str(e.value) and "타겟" in str(e.value), str(e.value)
    with pytest.raises(rt.InputError) as e:
        rt.apply_eval_expressions(df, ["d = 온돜 * 2"], target_cols=["생산량"])
    assert "[후보] 1) 온도" in str(e.value) and "[추천]" not in str(e.value), str(e.value)      # 2.9.33: [추천]


def test_train_eval_typo_of_the_target(role_data):
    proc, res = run_raw("ml_train.py", "--file", role_data["p"], "--features", "온도,압력", "--targets", "생산량", "--eval", "d = 생산랑 * 2")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg, msg


def test_positive_class_candidates(role_data):
    proc, res = run_raw("ml_train.py", "--file", role_data["c"], "--features", "온도,압력", "--targets", "판정", "--task", "classification",
                        "--positive-class", "불랑")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) 불량 2) 양품" in msg and "['" not in msg, msg


def _humid_rows(n=60, seed=1):
    r = random.Random(seed)
    rows = [["온도(°C)", "압력(bar)", "속도", "습도", "생산량"]]
    for _ in range(n):
        t, p, s, h = (round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1),
                      round(r.uniform(30, 70), 1))
        rows.append([t, p, s, h, round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    return rows


@pytest.fixture(scope="module")
def humid_data(tmp_path_factory, two_model):
    """온도(°C), 압력(bar), 속도, 습도 -> 생산량 파일(p)과 모델(pm), 두 타겟 모델(two)."""
    tmp = tmp_path_factory.mktemp("humid")
    out = {"tmp": tmp, "p": write_rows(os.path.join(str(tmp), "p.csv"), _humid_rows(), []), "two": two_model}
    _, res = run_raw("ml_train.py", "--file", out["p"], "--features", "온도(°C),압력(bar),속도,습도", "--targets", "생산량")
    assert res["status"] == "ok", res
    out["pm"] = os.path.join(str(tmp), "p_ml_data", "v1")
    return out


def test_train_and_predict_other_word(humid_data):
    proc, res = run_raw("ml_train.py", "--file", humid_data["p"], "--features", "온도,압력,습두", "--targets", "생산량")
    msg = stop_msg(res, proc)                                    # 2.9.32: one common syllable for another, [후보]
    assert "[추천]" not in msg and "[후보] 1) '습도'" in msg, msg
    proc, res = run_raw("ml_train.py", "--file", humid_data["p"], "--features", "온도,압력,습돜", "--targets", "생산량")
    assert "[후보] 1) '습도'" in stop_msg(res, proc)                               # 2.9.33: [추천] -> [후보]
    proc, res = run_raw("ml_predict.py", "--model", humid_data["two"], "--input", "온도=180,압력=2.5,농도=3")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "농도" in msg, msg


def _formula_error(expr, cols):
    with pytest.raises(rt.InputError) as e:
        rt.apply_eval_expressions(pd.DataFrame({c: [1.0, 2.0] for c in cols}), [expr])
    return str(e.value)


def test_eval_unknown_name_wording():
    msg = _formula_error("d = 수율 * 2", ["수율(1차)", "수율(최종)", "온도"])
    assert "[후보] 1) 수율(1차) 2) 수율(최종)" in msg and "['" not in msg.split("[사용 가능한")[0], msg
    assert "-> [후보] 1) 온도 " in _formula_error("d = 온두 * 2", ["온도", "압력"])     # 2.9.32: [추천] -> [후보]
    assert "-> [후보] 1) 온도 " in _formula_error("d = 온돜 * 2", ["온도", "압력"])     # 2.9.33: [추천] -> [후보]
    msg = _formula_error("d = T3 - T1", ["T1", "T2"])
    assert "[추천]" not in msg and "[후보] 1) T2" in msg, msg


# ---- --eval 수식 속 숫자와 숫자 이름 열 ('2', '2(mm)', '0.5(mm)') --------------


def test_number_named_column_in_formula_stops(name_files):
    with pytest.raises(rt.InputError) as e:
        rt.compile_formula("온도*2", ["온도", "2", "수율"])
    assert "숫자 이름의 열 '2'" in str(e.value)
    # 2.9.32: '2.5' became <column 2>.5; 2.9.33/2.9.34 stopped; 2.9.35: a number inside a larger number is no column
    assert rt.compile_formula("온도*3 + 2.5", ["온도", "2"])[2] == "온도*3 + 2.5"
    # no collision: a formula without that number, or another number that holds its digit
    assert rt.compile_formula("온도*3 + 12", ["온도", "2"])[1] == {"__COL_TOKEN_0__": "온도"}
    proc, res = run_raw("ml_train.py", "--file", name_files["num"], "--features", "온도", "--targets", "수율", "--eval", "x2=온도*2")
    msg = stop_msg(res, proc)
    assert "숫자 이름의 열" in msg, msg


@pytest.mark.parametrize("formula,col", [("온도*2", "2(mm)"), ("온도*0.5", "0.5(mm)")])   # '온도+2.5': 2.9.35 compiles
def test_number_alias_in_formula_stops(formula, col):
    with pytest.raises(rt.InputError) as e:
        rt.compile_formula(formula, ["온도", "0.5(mm)", "2(mm)", "수율"])
    assert f"'{col}'" in str(e.value)


def test_number_alias_never_a_formula_name():
    tok, cols, canon = rt.compile_formula("온도*3", ["온도", "2(mm)", "수율"])
    assert list(cols.values()) == ["온도"] and canon == "온도*3"


def test_train_number_alias_formula_stops(percent_files):
    proc, res = run_raw("ml_train.py", "--file", percent_files["sieve"], "--features", "온도,0.5(mm)", "--targets", "수율", "--eval", "a=온도*2")
    assert "'2(mm)'" in stop_msg(res, proc)


SIEVE = ["온도", "0.5(mm)", "2(mm)", "수율"]


@pytest.mark.parametrize("formula,cols", [("0.5(mm)*2(mm)", ["0.5(mm)", "2(mm)"]), ("온도+2(mm)", ["온도", "2(mm)"]),
                                          ("2(mm)*2.5", ["2(mm)"])])
def test_full_number_name_compiles(formula, cols):
    tok, t2c, canon = rt.compile_formula(formula, SIEVE)
    assert sorted(t2c.values()) == sorted(cols) and canon == formula, (tok, t2c)


@pytest.mark.parametrize("formula,columns", [("온도*2.5", SIEVE), ("온도*0.5", ["온도", "5(mm)"]),
                                             ("온도*2.5", ["온도", "2"])])
def test_number_inside_number_is_no_column(formula, columns):
    tok, t2c, canon = rt.compile_formula(formula, columns)
    assert list(t2c.values()) == ["온도"] and canon == formula, (tok, t2c)


@pytest.mark.parametrize("formula,columns,col", [("온도*2", SIEVE, "2(mm)"), ("온도*2", ["온도", "2"], "2"),
                                                 ("온도*0.5", SIEVE, "0.5(mm)")])
def test_bare_number_of_column_still_stops(formula, columns, col):
    with pytest.raises(rt.InputError) as e:
        rt.compile_formula(formula, columns)
    assert f"'{col}'" in str(e.value)


def test_train_full_number_names(near_name_files):
    proc, res = run_raw("ml_train.py", "--file", near_name_files["sieve"], "--features", "온도", "--targets", "수율",
                        "--eval", "합=0.5(mm)*2(mm);차=온도+2(mm)")
    ok_res(res, proc)


# ---- 예측 --input 키, --targets, --find 목표: 타겟 이름과 모델이 쓰지 않는 키 ----------------------------


@pytest.fixture(scope="module")
def target_models(tmp_path_factory, two_model, p1_model):
    """두 타겟 모델(two)과 온도(°C), 압력(bar), 속도 -> 생산량 모델(p1)."""
    return {"tmp": tmp_path_factory.mktemp("targets"), "two": two_model, "p1": p1_model["v"], "p1_f": p1_model["f"]}


def test_find_ambiguous_goal_refused(target_models):
    proc, res = run_raw("ml_predict.py", "--model", target_models["two"], "--find", "수율=65", "--vary", "온도", "--input", "압력=2.5")
    msg = stop_msg(res, proc)
    assert "[후보]" in msg and "수율(1차)" in msg and "수율(최종)" in msg, msg
    proc, res = run_raw("ml_predict.py", "--model", target_models["two"], "--find", "수율(1차)=65", "--vary", "온도",
                        "--input", "압력=2.5")
    assert ok_res(res, proc)["target"] == "수율(1차)"


def test_predict_targets_filter(target_models):
    proc, res = run_raw("ml_predict.py", "--model", target_models["two"], "--targets", "수율", "--input", "온도=180,압력=2.5")
    msg = stop_msg(res, proc)
    assert "[후보]" in msg and "수율(1차)" in msg and "수율(최종)" in msg, msg
    # an exact name takes that target only (it also took '수율(최종)': same name without unit)
    proc, res = run_raw("ml_predict.py", "--model", target_models["two"], "--targets", "수율(최종)", "--input", "온도=180,압력=2.5")
    assert ok_res(res, proc)["targets_predicted"] == ["수율(최종)"]


@pytest.mark.report
def test_report_target_filter(target_models):
    with pytest.raises(rt.InputError) as e:
        ml_report.load_version(target_models["two"], ["수율"])
    assert "[후보]" in str(e.value) and "수율(1차)" in str(e.value) and "수율(최종)" in str(e.value)
    ctx = ml_report.load_version(target_models["two"], ["수율(1차)"])
    assert [t["name"] for t in ctx["targets"]] == ["수율(1차)"]


def test_inline_other_unit_extra_key_refused(target_models):
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--input", "압력(bar)=3,압력(MPa)=0.5,온도=180,속도=45")
    msg = stop_msg(res, proc)
    assert "압력(MPa)" in msg and "단위가 다릅니다" in msg and "[추천]" not in msg, msg


def test_inline_unknown_key_refused(target_models):
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--input", "압력=3,온도=180,속도=45,습도=50")
    msg = stop_msg(res, proc)
    assert "습도" in msg and "[추천]" not in msg and "[모델 입력 목록]" in msg and "속도" in msg, msg
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--input", '{"압력": 3, "온도": 180, "속도": 45, "id": 7}')
    msg = stop_msg(res, proc)
    assert "'id'" in msg, msg


def test_inline_typo_key_is_a_candidate(target_models):
    # a typo of a model input that is not given: [후보] (2.9.33, was [추천] '속도'; a close name may be another word)
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--input", "압력=3,온도=180,속돜=45")
    msg = stop_msg(res, proc)
    assert "이름 중 '속도'(으)로 고칠 이름을 사용자에게 물으십시오: [후보] 1) 속돜" in msg and "[추천]" not in msg, msg
    # 2.9.32: '속두' (one common syllable for another) is [후보]
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--input", "압력=3,온도=180,속두=45")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) 속두" in msg, msg


def test_inline_exact_keys_ok_and_file_extra_columns_kept(target_models):
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--input", "압력=3,온도=180,속도=45")
    ok_res(res, proc)
    f = write_rows(os.path.join(str(target_models["tmp"]), "in_extra.csv"),
               [["번호", "온도(°C)", "압력(bar)", "속도", "습도"], [1, 180, 3, 45, 50], [2, 182, 3.1, 46, 51]], [])
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--file", f, "--save", str(target_models["tmp"] / "in_extra_out.csv"))
    ok_res(res, proc)


def test_find_and_map_inline_unknown_key_refused(target_models):
    proc, res = run_raw("ml_predict.py", "--model", target_models["p1"], "--find", "생산량=470", "--vary", "온도",
                        "--input", "압력=3,속도=45,습도=50")
    msg = stop_msg(res, proc)
    assert "습도" in msg and "[추천]" not in msg, msg     # '속도' is given: no typo target
    proc, res = run_raw("ml_map.py", "--file", target_models["p1_f"], "--input", "압력=3,속도=45,습도=50")
    msg = stop_msg(res, proc)
    assert "습도" in msg and "[추천]" not in msg, msg


def test_one_name_for_two_inputs(humid_data):
    f = write_rows(os.path.join(str(humid_data["tmp"]), "sol.csv"), [["온도(°C)", "압력(bar)", "솔도"], [180, 3, 45]], [])
    proc, res = run_raw("ml_predict.py", "--model", humid_data["pm"], "--file", f, "--save", str(humid_data["tmp"] / "sol_out.csv"))
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and msg.count("[후보]") == 1 and "[후보] 1) 속도 2) 습도" in msg, msg
    hints = mp._typo_hints(["속도", "습도"], ["온도(°C)", "솔도"], ["온도(°C)", "속도", "습도"], False)
    assert len(hints) == 1 and "[후보] 1) 속도 2) 습도" in hints[0], hints
    assert mp._typo_hints(["속도"], ["속돜"], ["속도"], True) == \
        ["'속도'이(가) 없습니다. --input의 이름 중 '속도'(으)로 고칠 이름을 사용자에게 물으십시오: [후보] 1) 속돜"]      # 2.9.33: [추천]
    assert "[후보] 1) 속두" in mp._typo_hints(["속도"], ["속두"], ["속도"], True)[0]     # 2.9.32: [추천] -> [후보]
    assert "[추천]" not in mp._typo_hints(["속도"], ["습도"], ["속도"], True)[0]


@pytest.mark.report
def test_spaced_unit_is_exact(humid_data):
    ctx = ml_report.load_version(humid_data["two"], ["수율( 1차)"])
    assert [t["name"] for t in ctx["targets"]] == ["수율(1차)"]
    proc, res = run_raw("ml_predict.py", "--model", humid_data["two"], "--targets", "수율( 1차)", "--input", "온도=180,압력=2.5")
    assert ok_res(res, proc)["targets_predicted"] == ["수율(1차)"]
    assert rt.resolve_column_names(["압력 ( bar )"], ["압력(bar)", "온도"], "피처(--features)")[0] == ["압력(bar)"]
    proc, res = run_raw("ml_predict.py", "--model", humid_data["pm"], "--input", "온도( °C)=180,압력 [bar]=3,속도=45,습도=50")
    assert ok_res(res, proc)["status"] == "ok"                                 # model inputs '온도(°C)', '압력(bar)'
    # 2.9.33: another spelling of the same unit matches too (was [추천])
    assert rt.resolve_column_names(["온도(C)"], ["온도(°C)"], "피처(--features)")[0] == ["온도(°C)"]


def test_common_syllable_typo_cli_is_a_candidate(role_data):
    proc, res = run_raw("ml_predict.py", "--model", role_data["pm"], "--input", "온도=180,압력=3,속도=45,습두=50")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) 습두" in msg, msg
    proc, res = run_raw("ml_predict.py", "--model", role_data["pm"], "--input", "온도=180,압력=3,속도=45,습돜=50")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) 습돜" in msg, msg                 # 2.9.33: was [추천] '습도'


def test_find_inputs_never_suggest_each_other(role_data):
    proc, res = run_raw("ml_predict.py", "--model", role_data["pm"], "--find", "생산량=480", "--vary", "속도", "--input", "속돜=45")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg, msg                                   # M2: the --vary input
    proc, res = run_raw("ml_predict.py", "--model", role_data["pm"], "--find", "생산량=480", "--vary", "속돜", "--input", "속도=45")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg, msg                                   # an input fixed by --input


# ---- 다른 단위의 이름, 별칭 쌍, JSON의 같은 키 ----------------------------------------------------------


def _train_alias(tmp_path, name, cols):
    r = random.Random(1)
    rows = [cols + ["수율"]]
    for _ in range(60):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2)
        v = {"온도": t, "온도(°C)": round(t + r.uniform(-0.5, 0.5), 1), "압력(bar)": p}
        rows.append([v[c] for c in cols] + [round(0.5 * t + 10 * p + r.uniform(-1, 1), 2)])
    f = write_rows(str(tmp_path / name), rows, [])
    _, res = run_raw("ml_train.py", "--file", f, "--features", ",".join(cols), "--targets", "수율")
    assert res["status"] == "ok", res
    return os.path.join(str(tmp_path), os.path.splitext(name)[0] + "_ml_data", "v1"), f


@pytest.fixture(scope="module")
def alias_models(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("alias")
    both, both_f = _train_alias(tmp, "al.xlsx", ["온도", "온도(°C)", "압력(bar)"])
    unit, unit_f = _train_alias(tmp, "al2.xlsx", ["온도(°C)", "압력(bar)"])
    return {"both": both, "unit": unit, "unit_f": unit_f, "tmp": tmp}


def _pred(model, inline):
    return run_raw("ml_predict.py", "--model", model, "--input", inline)


def test_other_unit_no_rename_suggestion(alias_models):
    proc, res = _pred(alias_models["unit"], "온도(°C)=180,압력(MPa)=0.3")
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "[추천]" not in res["message"] and "단위가 다릅니다" in res["message"], res["message"]
    assert "bar 단위로 환산" in res["message"], res["message"]
    # a unit-less name is still matched; the same unit spaced otherwise is the same name (2.9.31, was [추천])
    assert _pred(alias_models["unit"], "온도=180,압력=3")[1]["status"] == "ok"
    proc, res = _pred(alias_models["unit"], "온도(°C)=180,압력(bar )=3")
    assert res["status"] == "ok", res


def test_other_unit_ml_map_and_find(alias_models):
    proc, res = run_raw("ml_map.py", "--file", alias_models["unit_f"], "--input", "온도(°C)=180,압력(MPa)=0.3", "--targets", "수율")
    assert proc.returncode == 1 and "[추천]" not in res["message"] and "단위가 다릅니다" in res["message"], res
    proc, res = run_raw("ml_predict.py", "--model", alias_models["unit"], "--find", "수율=100", "--vary", "온도(°C)",
                        "--input", "압력(MPa)=0.3")
    assert proc.returncode == 1 and "[추천]" not in res["message"] and "단위가 다릅니다" in res["message"], res


def test_alias_pair_with_both_features_accepted(alias_models):
    assert rt.parse_inline_input("온도=180,온도(°C)=182").shape == (1, 2)
    for inline in ("온도=180,온도(°C)=182,압력(bar)=3", '{"온도":180,"온도(°C)":182,"압력(bar)":3}'):
        proc, res = _pred(alias_models["both"], inline)
        assert res["status"] == "ok", res


@pytest.mark.parametrize("inline", ["온도=180,온도(°C)=182,압력(bar)=3", '{"온도":180,"온도(°C)":182,"압력(bar)":3}',
                                    '[{"온도":180,"온도(°C)":182,"압력(bar)":3}]'])
def test_alias_pair_one_feature_refused(alias_models, inline):
    proc, res = _pred(alias_models["unit"], inline)
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "같은 열 '온도(°C)'" in res["message"], res
    if inline.startswith("온도="):        # 2.9.73: 번호와 같은 번호의 전체 명령 (다른 이름 하나를 뺀 --input)
        assert "[후보] 1) 온도=180 2) 온도(°C)=182" in res["message"], res
        got = [split_command(c["command"])[-1] for c in res["confirm_commands"]]
        assert got == ["온도=180,압력(bar)=3", "온도(°C)=182,압력(bar)=3"], got
    else:                                 # JSON 형식은 목록을 나눌 수 없어 값 목록만
        assert "[후보] 1) 온도=180 2) 온도(°C)=182" in res["message"] and "confirm_commands" not in res, res


@pytest.mark.parametrize("other", ["온도(C)", "온도(℃)", "온도( °C)"])
def test_alias_pair_with_degree_spelling_refused(alias_models, other):
    # 23차 점검 L1 (2.9.44): '온도(C)' is '온도(°C)' written otherwise, so '온도' + '온도(C)' are the same column twice
    # (2.9.43: "'온도'은(는) 모델 입력이 아님")
    proc, res = _pred(alias_models["unit"], f"온도=180,{other}=182,압력(bar)=3")
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "같은 열 '온도(°C)'" in res["message"] and f"[후보] 1) 온도=180 2) {other}=182" in res["message"], res
    assert "모델 입력이 아니" not in res["message"], res
    mp.check_inline_aliases(rt.parse_inline_input(f"{other}=182,압력(bar)=3"), ["온도(°C)", "압력(bar)"])   # one name: ok


def test_alias_pair_ml_map_and_find(alias_models):
    proc, res = run_raw("ml_map.py", "--file", alias_models["unit_f"], "--input", '{"온도":180,"온도(°C)":182}', "--targets", "수율")
    assert proc.returncode == 1 and "[후보] 1) 온도=180 2) 온도(°C)=182" in res["message"], res
    proc, res = run_raw("ml_predict.py", "--model", alias_models["unit"], "--find", "수율=100", "--vary", "압력(bar)",
                        "--input", "온도=180,온도(°C)=182")
    assert proc.returncode == 1 and "[후보] 1) 온도=180 2) 온도(°C)=182" in res["message"], res


def test_json_repeated_key_refused(alias_models):
    with pytest.raises(rt.InputError) as e:
        rt.parse_inline_input('{"온도(°C)":180,"온도(°C)":190}')
    assert "두 번" in str(e.value) and "[후보] 1) 온도(°C)=180 2) 온도(°C)=190" in str(e.value)
    p = os.path.join(str(alias_models["tmp"]), "in.json")
    with open(p, "w", encoding="utf-8") as f:
        f.write('[{"온도(°C)":180,"압력(bar)":3,"압력(bar)":4}]')
    proc, res = _pred(alias_models["unit"], p)
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "[후보] 1) 압력(bar)=3 2) 압력(bar)=4" in res["message"], res
    # a plain JSON row still predicts
    assert _pred(alias_models["unit"], json.dumps({"온도(°C)": 180, "압력(bar)": 3}, ensure_ascii=False))[1]["status"] == "ok"


def _pressure_rows(cols, target, n=60, seed=1):
    """공정 파일 행: 온도(°C), 주어진 이름의 압력 열, 속도, 타겟."""
    r = random.Random(seed)
    rows = [cols + [target]]
    for _ in range(n):
        t, p, s = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1)
        v = {"온도(°C)": t, "속도": s}
        for c in cols:
            if "압력" in c:
                v[c] = p if c not in v else v[c]
        if "압력차(bar)" in cols:
            v["압력차(bar)"] = round(p * 0.3 + r.uniform(0, 0.2), 2)
        rows.append([v[c] for c in cols] + [round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    return rows


def _train_pressure(tmp, name, cols, target="생산량"):
    f = write_rows(os.path.join(str(tmp), name), _pressure_rows(cols, target), [])
    _, res = run_raw("ml_train.py", "--file", f, "--features", ",".join(cols), "--targets", rt.strip_unit(target))
    assert res["status"] == "ok", res
    return os.path.join(str(tmp), os.path.splitext(name)[0] + "_ml_data", "v1"), f


@pytest.fixture(scope="module")
def unit_models(tmp_path_factory):
    """압력 열 이름과 단위가 다른 모델들: p4 반응압력(bar), p3 압력(bar)+압력차(bar), p2 타겟 생산량(t)."""
    tmp = tmp_path_factory.mktemp("units")
    out = {"tmp": tmp}
    out["p4"], out["p4_f"] = _train_pressure(tmp, "p4.csv", ["온도(°C)", "반응압력(bar)", "속도"])
    out["p3"], out["p3_f"] = _train_pressure(tmp, "p3.csv", ["온도(°C)", "압력(bar)", "압력차(bar)", "속도"])
    out["p2"], _ = _train_pressure(tmp, "p2.csv", ["온도(°C)", "압력(bar)", "속도"], target="생산량(t)")
    return out


def test_other_unit_with_other_base_name_predict(unit_models):
    proc, res = run_raw("ml_predict.py", "--model", unit_models["p4"], "--input", "압력(MPa)=0.3,온도=180,속도=45")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg and "bar 단위로 환산" in msg, msg
    # file input column: same rule
    f = write_rows(os.path.join(str(unit_models["tmp"]), "in_mpa.csv"), [["압력(MPa)", "온도(°C)", "속도"], [0.3, 180, 45]], [])
    proc, res = run_raw("ml_predict.py", "--model", unit_models["p4"], "--file", f)
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg


def test_other_unit_with_other_base_name_ml_map(unit_models):
    proc, res = run_raw("ml_map.py", "--file", unit_models["p4_f"], "--input", "압력(MPa)=0.3", "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg
    proc, res = run_raw("ml_map.py", "--file", unit_models["p4_f"], "--find", "생산량=530", "--vary", "온도",
                        "--input", "압력(MPa)=0.3")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg
    # --points: a points column with the base name of a data column and another unit (other base names are
    # passed through as extra columns of the points file)
    pts = write_rows(os.path.join(str(unit_models["tmp"]), "pts.csv"), [["압력(MPa)", "온도"], [0.3, 180]], [])
    proc, res = run_raw("ml_map.py", "--file", unit_models["p3_f"], "--points", pts, "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "'압력(bar)'(단위 다름: 값 환산 필요)" in msg, msg


def test_other_unit_with_other_base_name_find(unit_models):
    proc, res = run_raw("ml_predict.py", "--model", unit_models["p4"], "--find", "생산량=530", "--vary", "온도",
                        "--input", "압력(MPa)=0.3")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg


def test_same_unit_close_name_is_a_candidate(unit_models):
    # 2.9.33: a close name of the same unit is a [후보] (was [추천]): no unit message, the user picks
    proc, res = run_raw("ml_predict.py", "--model", unit_models["p4"], "--input", "압력(bar)=3,온도=180,속도=45")
    msg = stop_msg(res, proc)
    assert "[후보]" in msg and "반응압력(bar)" in msg and "[추천]" not in msg and "단위가 다릅니다" not in msg, msg
    proc, res = run_raw("ml_map.py", "--file", unit_models["p4_f"], "--input", "반응압룍(bar)=3", "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[후보] 1) '반응압력(bar)'" in msg and "[추천]" not in msg, msg
    # 2.9.31: the same unit spaced otherwise is the same name (resolve_column_names), no [추천] (was the test above)
    proc, res = run_raw("ml_map.py", "--file", unit_models["p4_f"], "--input", "반응압력( bar)=3", "--targets", "생산량")
    assert proc.returncode == 0 and res["keys"] == ["반응압력(bar)"], res


def test_candidates_marked(unit_models):
    proc, res = run_raw("ml_map.py", "--file", unit_models["p3_f"], "--input", "압력(MPa)=0.3", "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[후보] 1) '압력(bar)'(단위 다름: 값 환산 필요) 2) '압력차(bar)'(단위 다름: 값 환산 필요)" in msg, msg
    proc, res = run_raw("ml_predict.py", "--model", unit_models["p3"], "--find", "생산량=530", "--vary", "온도",
                        "--input", "압력(MPa)=0.3")
    msg = stop_msg(res, proc)
    assert "'압력(bar)'(단위 다름: 값 환산 필요)" in msg and "'압력차(bar)'(단위 다름: 값 환산 필요)" in msg, msg
    # names without values (--targets) are not marked; a same-unit candidate is not marked
    with pytest.raises(rt.InputError) as e:
        rt.resolve_column_names(["압력(MPa)"], ["압력(bar)", "압력차(bar)"], "구할 열(--targets)")
    assert "단위 다름" not in str(e.value)
    with pytest.raises(rt.InputError) as e:
        rt.resolve_column_names(["압력(bar)"], ["압력차(bar)", "압력차2(bar)"], "기준(--input)", values=True)
    assert "[후보]" in str(e.value) and "단위 다름" not in str(e.value)


def test_find_goal_other_unit_refused(unit_models):
    proc, res = run_raw("ml_predict.py", "--model", unit_models["p2"], "--find", "생산량(kg)=500", "--vary", "온도",
                        "--input", "압력=3")
    msg = stop_msg(res, proc)
    assert "목표(--find)" in msg and "단위가 다릅니다" in msg and "t 단위로 환산" in msg, msg
    for goal in ("생산량(t)=500", "생산량=500", "생산량( t )=500"):
        proc, res = run_raw("ml_predict.py", "--model", unit_models["p2"], "--find", goal, "--vary", "온도", "--input", "압력=3")
        assert res["status"] == "ok", (goal, res)


# ---- ml_map --input, --find, --vary, --targets 이름 ------------------------------------------------


def test_map_inline_typo_is_a_candidate(name_files):
    proc, res = run_raw("ml_map.py", "--file", name_files["d"], "--input", "온도=180,압룍=3", "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '압력(bar)'" in msg, msg


def test_inline_keys_never_suggest_the_value_column(humid_data):
    # --input next to --targets / --find / --vary: the column those take is no key ([추천] '생산량' would fail on the retry)
    proc, res = run_raw("ml_map.py", "--file", humid_data["p"], "--input", "온도=180,생산랑=470", "--targets", "생산량")
    assert "[추천] '생산량'" not in stop_msg(res, proc)
    proc, res = run_raw("ml_map.py", "--file", humid_data["p"], "--find", "생산량=470", "--vary", "온도", "--input", "생산랑=3")
    assert "[추천] '생산량'" not in stop_msg(res, proc)
    proc, res = run_raw("ml_map.py", "--file", humid_data["p"], "--input", "온도=180,압룍(bar)=3", "--targets", "생산량")
    assert "[후보] 1) '압력(bar)'" in stop_msg(res, proc)                       # 2.9.33: [추천] -> [후보]


def test_inline_unitless_name_next_to_the_target_asks(role_data):
    proc, res = run_raw("ml_map.py", "--file", role_data["two"], "--input", "온도=180,수율=40", "--targets", "수율(최종)")
    msg = stop_msg(res, proc)
    # 2.9.33: the excluded '수율(최종)' (the --targets column) is never a [후보] (was "[후보] '수율(1차)', '수율(최종)'")
    assert "[추천]" not in msg and "[후보] 1) '수율(1차)' (비슷한 '수율(최종)'" in msg, msg
    msg = _resolve_error(["수율"], ["온도", "수율(1차)", "수율(최종)"], exclude=["수율(최종)"])
    assert "[추천]" not in msg and "[후보] 1) '수율(1차)' (비슷한 '수율(최종)'" in msg, msg


def test_map_find_vary_never_suggest_each_other(role_data):
    proc, res = run_raw("ml_map.py", "--file", role_data["p"], "--find", "생산량=480", "--vary", "생산랑")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "생산량" in msg, msg               # M3: the --find goal
    proc, res = run_raw("ml_map.py", "--file", role_data["p"], "--find", "속돜=45", "--vary", "속도")
    assert "[추천]" not in stop_msg(res, proc)                            # the goal is no --vary column
    proc, res = run_raw("ml_map.py", "--file", role_data["p"], "--input", "속도=45", "--targets", "속돜")
    assert "[추천]" not in stop_msg(res, proc)                            # --targets is no --input key


# ---- ml_map --points 열: 키 열, 실측값 열, 메모 열 ------------------------------------------


# 조회 파일의 열이 쓰지 않은 데이터 열과 비슷하면(키 열일 수 있음) [후보]로 멈춘다. --targets 이름을 품은 실측값·메모 열
# ('수율_실측', '목표수율')은 비교할 때 --targets 이름을 빼고 보며, 그래도 비슷한 열이 없으면 통과해 결과 시트에 남는다.


@pytest.fixture(scope="module")
def points_map_data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("pointsmap")
    rows = [["온도", "반응압력(bar)", "생산량"]]
    for t in (170, 180, 190):
        for p in (2.0, 2.5, 3.0, 3.5, 4.0):
            rows.append([t, p, round(2 * t + 30 * p, 1)])
    return tmp, write_rows(os.path.join(str(tmp), "map.csv"), rows, [])


def _points_in(tmp, name, head, rows):
    return write_rows(os.path.join(str(tmp), name), [head] + rows, [])


def test_points_near_name_refused(points_map_data):
    tmp, data = points_map_data
    p = _points_in(tmp, "sp.csv", ["온도", "반응압룍(bar)"], [[180, 2.0], [180, 4.0]])
    proc, res = run_raw("ml_map.py", "--file", data, "--points", p, "--targets", "생산량", "--save", str(tmp / "o1.xlsx"))
    msg = stop_msg(res, proc)
    assert "[후보] 1) '반응압력(bar)'" in msg and "반응압룍(bar)" in msg and "[추천]" not in msg, msg     # 2.9.33: [추천]. 이름이 조회 파일 머리글에 있어 명령으로 고칠 수 없음 (번호 없음)
    # 2.9.31: '반응 압력(bar)' (the same name and unit spaced otherwise) is the key itself
    p = _points_in(tmp, "sp1.csv", ["온도", "반응 압력(bar)"], [[180, 2.0], [180, 4.0]])
    proc, res = run_raw("ml_map.py", "--file", data, "--points", p, "--targets", "생산량", "--save", str(tmp / "o1b.xlsx"))
    assert ok_res(res, proc)["keys"] == ["온도", "반응압력(bar)"], res
    p = _points_in(tmp, "mpa.csv", ["온도", "압력(MPa)"], [[180, 0.3], [180, 0.4]])
    proc, res = run_raw("ml_map.py", "--file", data, "--points", p, "--targets", "생산량", "--save", str(tmp / "o2.xlsx"))
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg and "bar 단위로 환산" in msg, msg


def test_points_memo_columns_pass(points_map_data):
    tmp, data = points_map_data
    p = _points_in(tmp, "memo.csv", ["온도", "반응압력(bar)", "번호", "비고"], [[180, 2.2, 1, "시험"], [185, 3.1, 2, "x"]])
    proc, res = run_raw("ml_map.py", "--file", data, "--points", p, "--targets", "생산량", "--save", str(tmp / "o3.xlsx"))
    res = ok_res(res, proc)
    assert res["keys"] == ["온도", "반응압력(bar)"] and res["count"] == 2, res


def _map_points(humid_data, pts, *extra):
    out = str(humid_data["tmp"] / (os.path.basename(pts) + str(len(extra)) + "_o.xlsx"))
    return run_raw("ml_map.py", "--file", humid_data["p"], "--points", pts, "--save", out, "--overwrite", *extra)


@pytest.mark.parametrize("col", ["생산량_실측", "생산량(실측)", "생산량"])
def test_points_target_measured_column_passes_with_y(humid_data, col):
    pts = data_file(humid_data, f"m_{len(col)}_{col[-2:]}.csv", ["온도(°C)", "압력(bar)", col], [[180, 3, 470], [185, 3.5, 490]])
    res = ok_res(*reversed(_map_points(humid_data, pts, "--targets", "생산량")))
    assert res["keys"] == ["온도(°C)", "압력(bar)"] and res["columns"] == ["생산량"] and res["count"] == 2, res
    got = openpyxl.load_workbook(res["output_file"])["결과"]
    assert col in [c.value for c in got[1]]                      # the measured column is kept in the result sheet


@pytest.mark.parametrize("col", ["생산량_실측", "생산량(실측)"])
def test_points_without_y_asks(humid_data, col):
    pts = data_file(humid_data, f"n_{col[-3:-1]}.csv", ["온도(°C)", "압력(bar)", col], [[180, 3, 470], [185, 3.5, 490]])
    proc, res = _map_points(humid_data, pts)
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보]" in msg and "--targets" in msg and "기준(--input)" not in msg, msg


def test_points_other_word_is_no_recommendation(humid_data):
    # '습도2' next to the data column '습도': another number, [후보]; a sure typo keeps [추천] and its retry works
    # (2.9.32: '속두' is one common syllable for another, [후보]; the rare-syllable typo '속돜' keeps [추천])
    pts = data_file(humid_data, "h2.csv", ["온도(°C)", "압력(bar)", "습도2"], [[180, 3, 1], [185, 3.5, 2]])
    proc, res = _map_points(humid_data, pts, "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '습도'" in msg, msg
    pts = data_file(humid_data, "sp.csv", ["온도(°C)", "압력 (bar)", "속두"], [[180, 3, 45], [185, 3.5, 46]])
    proc, res = _map_points(humid_data, pts, "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '속도'" in msg, msg
    pts = data_file(humid_data, "sp1.csv", ["온도(°C)", "압력 (bar)", "속돜"], [[180, 3, 45], [185, 3.5, 46]])
    proc, res = _map_points(humid_data, pts, "--targets", "생산량")
    msg = stop_msg(res, proc)                                       # 2.9.33: every close name is [후보] (was [추천])
    assert "[추천]" not in msg and "[후보] 1) '속도'" in msg, msg
    pts = data_file(humid_data, "sp2.csv", ["온도(°C)", "압력 (bar)", "속도"], [[180, 3, 45], [185, 3.5, 46]])
    ok_res(*reversed(_map_points(humid_data, pts, "--targets", "생산량")))


def test_points_unitless_name_of_the_target_is_no_key(role_data):
    pts = data_file(role_data, "pts.csv", ["온도", "압력", "수율"], [[180, 2.5, 40.1], [185, 2.6, 41]])
    out = str(role_data["tmp"] / "pts_o.xlsx")
    proc, res = run_raw("ml_map.py", "--file", role_data["two"], "--points", pts, "--targets", "수율(최종)", "--save", out, "--overwrite")
    res = ok_res(res, proc)
    assert res["keys"] == ["온도", "압력"] and res["columns"] == ["수율(최종)"], res     # 2.9.31: 수율(1차) a key
    assert "수율" in [c.value for c in openpyxl.load_workbook(res["output_file"])["결과"][1]]


def test_points_without_y_candidate_has_the_y_guidance(role_data):
    pts = data_file(role_data, "h2.csv", ["온도(°C)", "압력(bar)", "습도2"], [[180, 3, 1], [185, 3.5, 2]])
    out = str(role_data["tmp"] / "h2_o.xlsx")
    proc, res = run_raw("ml_map.py", "--file", role_data["p"], "--points", pts, "--save", out, "--overwrite")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '습도'" in msg and "--targets에 구할 열" in msg, msg


@pytest.mark.parametrize("col", ["수율 실측", "수율_실측", "목표수율", "수율목표"])
@pytest.mark.parametrize("target", ["수율(최종)", "생산량"])
def test_points_measured_column_passes(name_files, col, target):
    pts = data_file(name_files, f"p_{col}_{target}.csv".replace(" ", "_").replace("(", "").replace(")", ""),
                  ["온도", "압력", col], [[170, 3, 60], [180, 2.5, 61]])
    proc, res = run_raw("ml_map.py", "--file", name_files["d"], "--points", pts, "--targets", target)
    res = ok_res(res, proc)
    assert res["keys"] == ["온도(°C)", "압력(bar)"] and res["columns"] == [target], res


def test_points_near_key_name_still_stops(name_files):
    pts = data_file(name_files, "p_near.csv", ["온도", "압력", "솎도"], [[170, 3, 20]])
    proc, res = run_raw("ml_map.py", "--file", name_files["d"], "--points", pts, "--targets", "생산량")
    msg = stop_msg(res, proc)
    assert "[추천]" not in msg and "[후보] 1) '속도'" in msg, msg


def test_points_without_y_guidance_only_for_names(name_files):
    bad = data_file(name_files, "p_bad.csv", ["온도", "압력"], [["1:30", 3]])
    proc, res = run_raw("ml_map.py", "--file", name_files["d"], "--points", bad)
    msg = stop_msg(res, proc)
    assert "숫자가 아닙니다" in msg and "--targets가 없어" not in msg, msg
    dup = data_file(name_files, "p_dup.csv", ["온도", "온도(°C)", "압력"], [[170, 170, 3]])
    proc, res = run_raw("ml_map.py", "--file", name_files["d"], "--points", dup)
    msg = stop_msg(res, proc)
    assert "같은 열" in msg and "--targets가 없어" not in msg, msg
    name = data_file(name_files, "p_name.csv", ["온도", "압력", "수율"], [[170, 3, 60]])
    proc, res = run_raw("ml_map.py", "--file", name_files["d"], "--points", name)
    msg = stop_msg(res, proc)
    assert "[후보]" in msg and "--targets가 없어" in msg, msg


def test_points_key_column_with_target_name_stops(percent_files):
    pts = data_file(percent_files, "p_cool.csv", ["냉각수 온도(°C)", "유량"], [[20, 3], [28, 3]])
    proc, res = run_raw("ml_map.py", "--file", percent_files["cool"], "--points", pts, "--targets", "온도", "--save", data_path(percent_files, "o1"))
    msg = stop_msg(res, proc)
    assert "[후보] 1) '냉각수온도'" in msg and "[추천]" not in msg, msg


@pytest.mark.parametrize("src,col,target,cand", [
    ("moist", "함수율%", "수율", "함수율"),
    ("moist_u", "함수율%", "수율(최종)", "함수율(%)"),
])
def test_points_moisture_column_with_target_name_stops(percent_files, src, col, target, cand):
    pts = data_file(percent_files, f"p_{src}.csv", ["온도" if src == "moist" else "반응온도(°C)", col, "메모"],
                [[160, 6, "a"], [180, 14, "b"]])
    proc, res = run_raw("ml_map.py", "--file", percent_files[src], "--points", pts, "--targets", target, "--save", data_path(percent_files, "o2"))
    msg = stop_msg(res, proc)
    assert f"[후보] 1) '{cand}'" in msg and "[추천]" not in msg, msg


@pytest.mark.parametrize("col,target", [("수율_실측", "수율(최종)"), ("목표수율", "수율(최종)"),
                                        ("생산량_실측", "생산량"), ("수율_실측", "생산량")])
def test_points_measured_column_passes_next_to_target_sibling(percent_files, col, target):
    pts = data_file(percent_files, f"p_m_{col}_{target}.csv".replace("(", "").replace(")", ""), ["온도", "압력", col],
                [[170, 3, 60], [180, 2.5, 61]])
    out = data_path(percent_files, f"o_m_{col}.xlsx")
    proc, res = run_raw("ml_map.py", "--file", percent_files["d"], "--points", pts, "--targets", target, "--save", out, "--overwrite")
    res = ok_res(res, proc)
    assert res["keys"] == ["온도(°C)", "압력(bar)"] and res["columns"] == [target], res
    assert col in pd.read_excel(res["output_file"], sheet_name="결과").columns


def test_check_near_points_both_directions():
    pts = pd.DataFrame({"냉각수 온도(°C)": [20.0], "유량": [3.0]})
    with pytest.raises(rt.ColumnNameError):
        ml_map._check_near_points(pts, ["유량"], ["냉각수온도"], {}, used=["유량", "온도"])
    ok = pd.DataFrame({"온도": [170.0], "수율_실측": [60.0]})
    ml_map._check_near_points(ok, ["온도"], ["압력(bar)", "수율(1차)"], {}, used=["온도(°C)", "수율(최종)"])


@pytest.mark.parametrize("col,target", [("수율_실측", "수율"), ("목표수율", "수율"), ("수율_실측", "수율,생산량"),
                                        ("목표수율", "수율,생산량"), ("생산량_실측", "수율,생산량")])
def test_points_measured_column_passes_next_to_moisture(near_name_files, col, target):
    pts = data_file(near_name_files, f"p_{col}.csv", ["온도", col], [[160, 30], [180, 20]])
    proc, res = run_raw("ml_map.py", "--file", near_name_files["grid"], "--points", pts, "--targets", target,
                        "--save", data_path(near_name_files, f"o_{col}_{len(target)}"))
    res = ok_res(res, proc)
    assert res["agent_directive"].startswith("VERIFIED:"), res


@pytest.mark.parametrize("col", ["함수율%", "함수 율(%)", "함수율_수율", "함수율 수율", "수율실측함수율"])
def test_points_moisture_like_column_stops(near_name_files, col):
    pts = data_file(near_name_files, f"p_s{len(col)}.csv", ["온도", col], [[160, 8], [180, 12]])
    proc, res = run_raw("ml_map.py", "--file", near_name_files["grid"], "--points", pts, "--targets", "수율",
                        "--save", data_path(near_name_files, "o_s"))
    msg = stop_msg(res, proc)
    assert "[후보] 1) '함수율'" in msg and "[추천]" not in msg, msg


def test_points_cooling_key_still_stops(near_name_files):
    pts = data_file(near_name_files, "p_cool.csv", ["시간", "냉각수 온도(°C)"], [[160, 25]])
    proc, res = run_raw("ml_map.py", "--file", near_name_files["cool"], "--points", pts, "--targets", "온도", "--save", data_path(near_name_files, "o_c"))
    assert "[후보] 1) '냉각수온도'" in stop_msg(res, proc)


# 위치·단 키 열('온도_1' / 데이터 '온도(1)', --targets '온도(3)')은 구두점을 빼고 비교해 가깝다고 본다 (2.9.35는 버려서 VERIFIED).


@pytest.fixture(scope="module")
def position_tmp(tmp_path_factory):
    return str(tmp_path_factory.mktemp("positions"))


def _position_data(position_tmp, name, cols):
    r = random.Random(3)
    rows = [cols]
    for _ in range(30):
        a, b, c = (round(r.uniform(100, 200), 1) for _ in range(3))
        rows.append([round(r.uniform(10, 30), 1), a, b, c])
    return write_rows(os.path.join(position_tmp, name), rows, [])


@pytest.mark.parametrize("pcol,dcols,target", [
    ("온도_1", ["시간", "온도(1)", "온도(2)", "온도(3)"], "온도(3)"),
    ("온도#1", ["시간", "온도(1)", "온도(2)", "온도(3)"], "온도(3)"),
    ("유량-1", ["시간", "유량(1)", "유량(2)", "유량(3)"], "유량(3)"),
    ("온도_상", ["시간", "온도(상)", "온도(중)", "온도(하)"], "온도(하)"),
    ("압력1", ["시간", "압력(1단)", "압력(3단)", "압력(2단)"], "압력(2단)"),
    ("수율_1", ["시간", "수율(1차)", "수율(2차)", "수율(최종)"], "수율(최종)"),
])
def test_points_position_key_column_stops(position_tmp, pcol, dcols, target):
    d = _position_data(position_tmp, f"d_{len(pcol)}_{pcol[-1]}.csv", dcols)
    p = write_rows(os.path.join(position_tmp, f"p_{len(pcol)}_{pcol[-1]}.csv"), [["시간", pcol], [20, 150], [20, 170]], [])
    proc, res = run_raw("ml_map.py", "--file", d, "--points", p, "--targets", target,
                        "--save", os.path.join(position_tmp, "o.xlsx"), "--overwrite")
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "[후보]" in res["message"] and "[추천]" not in res["message"], res


@pytest.mark.parametrize("pcol", ["수율_실측", "목표수율"])
def test_points_measured_column_passes_beside_positions(position_tmp, pcol):
    d = _position_data(position_tmp, "dm.csv", ["온도", "함수율", "수율", "생산량"])
    p = write_rows(os.path.join(position_tmp, f"pm_{len(pcol)}.csv"), [["온도", pcol], [15, 30], [25, 40]], [])
    proc, res = run_raw("ml_map.py", "--file", d, "--points", p, "--targets", "수율",
                        "--save", os.path.join(position_tmp, "om.xlsx"), "--overwrite")
    assert proc.returncode == 0 and res["agent_directive"].startswith("VERIFIED:"), res


# ---- ml_export 이름 ----------------------------------------------------------------------------------------


@pytest.mark.report
def test_export_keep_typo_never_suggests_a_given_input():
    with pytest.raises(rt.InputError) as e:
        ml_export._check_known(["압력", "압룍"], {"t": ["압력(bar)", "온도"]}, "입력(--keep)")
    assert "[추천]" not in str(e.value), str(e.value)
    with pytest.raises(rt.InputError) as e:
        ml_export._check_known(["압룍"], {"t": ["압력(bar)", "온도"]}, "입력(--keep)")
    assert "[후보] 1) '압력(bar)'" in str(e.value), str(e.value)                  # 2.9.33: [추천]


@pytest.mark.report
def test_export_labels_two_targets(humid_data):
    assert ml_export.target_labels(["수율(1차)", "수율(최종)", "온도(°C)"]) == \
        {"수율(1차)": "수율(1차)", "수율(최종)": "수율(최종)", "온도(°C)": "온도"}
    out = str(humid_data["tmp"] / "exp.xlsx")
    proc, res = run_raw("ml_export.py", "--model", humid_data["two"], "--inputs", "온도,압력", "--save", out)
    res = ok_res(res, proc)
    wb = openpyxl.load_workbook(res["output_file"])
    heads = [c.value for c in wb["일괄계산"][2] if c.value]
    assert "예측 수율(1차)" in heads and "예측 수율(최종)" in heads and "예측 수율" not in heads, heads
    note = "\n".join(str(c.value) for row in wb["계산기"].iter_rows() for c in row if isinstance(c.value, str))
    assert "- 수율(1차):" in note and "- 수율(최종):" in note and "- 수율:" not in note, note
    assert {"표_수율(1차)", "표_수율(최종)"} <= set(wb.sheetnames), wb.sheetnames


def test_column_candidate_has_numbered_full_command(tmp_path):
    """2.9.71: 열 이름 하나가 틀리면 [후보]에 번호를 붙이고 같은 번호의 next_commands가 고친 전체 명령 (LLM이 이름을 끼워 넣지 않음)."""
    from make_sample import make_sample
    data = make_sample(str(tmp_path / "d.csv"))
    proc, res = run_raw("ml_train.py", "--file", data, "--features", "반응온도,소비전력,압축압렬", "--targets", "생산효율")
    msg = stop_msg(res, proc)
    assert "[후보] 1) '압축압력(bar)'" in msg and len(res["confirm_commands"]) == 1, res
    argv = split_command(res["confirm_commands"][0]["command"])
    assert argv[argv.index("--features") + 1] == "반응온도,소비전력,압축압력(bar)", argv
    assert run_tool(os.path.basename(argv[1]), *argv[2:])["version"] == "v1"


# ---- 2.9.73: 남은 [후보]도 번호와 같은 번호의 전체 명령 ------------------------------------


def test_two_missing_names_numbered_one_at_a_time(name_files):
    """없는 이름이 둘이면 첫 이름의 [후보]에 번호와 전체 명령, 둘째 이름은 다음 실행에서 물음 (명령을 섞어 만들지 않음)."""
    proc, res = run_raw("ml_train.py", "--file", name_files["acc"], "--features", "속도,압룍", "--targets", "진동")
    msg = stop_msg(res, proc)
    assert "[후보] 1) '가속도'" in msg and "다음 실행에서 묻습니다" in msg, msg
    fixed = split_command(res["confirm_commands"][0]["command"])
    assert fixed[fixed.index("--features") + 1] == "가속도,압룍", fixed


def test_command_replacing_refuses_a_value_given_twice(monkeypatch):
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["ml_map.py", "--file", "a.csv", "--input", "라인=A,등급=A"])
    assert rt.command_replacing("A", "B") is None                   # 어느 A인지 추측하지 않음
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["ml_map.py", "--file", "a.csv", "--input", "라인=A,등급=C"])
    assert split_command(rt.command_replacing("A", "B"))[-1] == "라인=B,등급=C"


def test_command_dropping_keeps_one_name(monkeypatch):
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["ml_predict.py", "--model", "v1", "--input", "온도=1,온도(°C)=2,압력=3"])
    assert split_command(rt.command_dropping(["온도(°C)"]))[-1] == "온도=1,압력=3"
    assert rt.command_dropping(["없는이름"]) is None
