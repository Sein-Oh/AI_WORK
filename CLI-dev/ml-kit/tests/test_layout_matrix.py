# -*- coding: utf-8 -*-
"""머리글·단위 행 레이아웃 매트릭스: 현장에서 본 모양의 작은 .xlsx/.csv로 학습과 예측을 끝까지 실행한다.

레이아웃마다:
  (a) 맞는 옵션으로 학습하면 VERIFIED, 열 이름과 학습 행이 맞다.
  (b) 옵션 없이 학습하면 VERIFIED이거나, STOP_RETRY의 [추천] 하나를 한 번 따르면 (a)와 같다.
  (c) 같은 레이아웃의 새 파일을 단위로 학습한 모델과 단위 없이 학습한 모델로, 기본 명령과 next_commands의 예측 명령으로
      예측하면 데이터 행만 모두 예측한다 (단위·머리글·메모 행은 데이터로 예측하지 않음). 또는 [추천] 하나가 그렇게 만든다.

묶음 셋: 기본 레이아웃(머리글, 단위, 제목, 병합 머리글, 메모, 날짜, 범주 먼저), 첫 행 측정불가·좁은 제목 CSV·반복 머리글·합계 행,
비고가 든 단위 행·배율 숫자 단위 행과 타겟 없는 예측 파일. 매트릭스마다 4개씩 병렬로 실행한다 (별도 프로세스).
"""
import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import split_command, FAMILIES, H, U, family_rows, run_raw, write_rows  # noqa: E402
import _ml_runtime as rt  # noqa: E402

N_TRAIN, N_PRED = 20, 5
X, Y = "반응온도,압력,라인", "수율"


# ---- 기본 레이아웃 (모든 매트릭스가 쓰는 실행 도구 포함) ----------------------------------------------------


def _set(data, row, col, value):
    data[row][col] = value
    return data


# id: (family, top rows, data edit, merges, correct options, unit row used for the names or None, file extension)
LAYOUTS = {
    "L1_hdr": ("A", [H], None, [], [], None, ".csv"),
    "L2_hdr_unit": ("A", [H, U], None, [], ["--unit-row", "2"], U, ".xlsx"),
    "L3_title_hdr_unit": ("A", [["반응 공정 데이터 2026"], [None], H, U], None, [],
                          ["--header-row", "3", "--unit-row", "4"], U, ".xlsx"),
    "L4_merged": ("A", [["공정 조건", None, None, "결과"], H, U], None, ["A1:C1"],
                  ["--header-row", "2", "--unit-row", "3"], U, ".xlsx"),
    "L5_unit_NG_first": ("A", [H, U], lambda d: _set(_set(d, 0, 0, "측정불가"), 1, 1, "NG"), [],
                         ["--unit-row", "2"], U, ".xlsx"),
    "L6_cat_first": ("C", [FAMILIES["C"]["cols"]], None, [], [], None, ".xlsx"),
    "L6b_cat_first_unit": ("C", [FAMILIES["C"]["cols"], FAMILIES["C"]["units"]], None, [],
                           ["--unit-row", "2"], FAMILIES["C"]["units"], ".xlsx"),
    "L7_blank_before": ("A", [H, U, [None], [None]], None, [], ["--unit-row", "2"], U, ".xlsx"),
    "L8_notes": ("A", [H, ["비고: 2026년 1월 측정", None, None, None]], None, [], ["--data-row", "3"], None, ".xlsx"),
    "L9_partial_unit": ("A", [H, ["°C", None, None, None]], None, [], ["--unit-row", "2"],
                        ["°C", None, None, None], ".xlsx"),
    "L10_degC": ("A", [H, ["℃", "bar", "-", "%"]], None, [], ["--unit-row", "2"], ["℃", "bar", "-", "%"], ".xlsx"),
    "L11_unit_blank_first": ("A", [H, U], lambda d: _set(d, 0, 0, None), [], ["--unit-row", "2"], U, ".xlsx"),
    "L12_pct_text": ("A", [H, U], lambda d: _set(d, 0, 3, "12%"), [], ["--unit-row", "2"], U, ".xlsx"),
    "L13_title_hdr": ("A", [["반응 공정 데이터"], H], None, [], ["--header-row", "2"], None, ".xlsx"),
    "L14_date_unit": ("D", [FAMILIES["D"]["cols"], FAMILIES["D"]["units"]], None, [], ["--unit-row", "2"],
                      FAMILIES["D"]["units"], ".xlsx"),
    "L15_date_hdr": ("D", [FAMILIES["D"]["cols"]], None, [], [], None, ".xlsx"),
}


# models for (c): per family, one trained with units and one without
MODELS = {"A": ("L2_hdr_unit", "L1_hdr"), "C": ("L6b_cat_first_unit", "L6_cat_first"),
          "D": ("L14_date_unit", "L15_date_hdr")}


def _make(folder, lid, n, seed):
    """(파일, 데이터 행의 1부터 센 원본 행 번호, 그 가운데 타겟이 있는 행)"""
    fam, top, edit, merges, _, _, ext = LAYOUTS[lid]
    data = family_rows(fam, n, seed)
    if edit:
        data = edit(data)
    path = write_rows(os.path.join(str(folder), lid + ext), [list(r) for r in top] + data, merges)
    rows = list(range(len(top) + 1, len(top) + n + 1))
    y = FAMILIES[fam]["cols"].index(FAMILIES[fam]["y"])
    valid = [rw for rw, d in zip(rows, data) if FAMILIES[fam]["task"] or rt.coerce_numeric(pd.Series([d[y]])).notna().all()]
    return path, rows, valid


def _names(lid, with_units):
    fam, units = LAYOUTS[lid][0], LAYOUTS[lid][5]
    cols = FAMILIES[fam]["cols"]
    if with_units and units:
        return [rt._with_unit(c, u) for c, u in zip(cols, units)]
    return list(cols)


def _recommendation(message):
    assert message.count("[추천]") == 1, message
    assert "[후보]" not in message, message
    return message.split("[추천] ")[1].split("\n")[0].strip()


def _follow(script, args, message):
    """[추천] 하나를 한 번 따른 명령: 전체 명령이면 그것, 옵션 값이면 명령에 넣거나 바꾼다."""
    rec = _recommendation(message)
    if rec.startswith("python"):
        parts = split_command(rec)
        return os.path.basename(parts[1]), parts[2:]
    toks = split_command(rec)
    assert toks and toks[0].startswith("--"), f"[추천] is not a command or an option: {rec}"
    new = list(args)
    for i in range(0, len(toks), 2):
        if toks[i] in new:
            new[new.index(toks[i]) + 1] = toks[i + 1]
        else:
            new += toks[i:i + 2]
    return script, new


def _run_with_one_retry(script, args):
    """(첫 결과, 최종 결과): STOP_RETRY의 [추천]을 한 번 따른다."""
    _, first = run_raw(script, *args)
    if first["status"] == "ok":
        return first, first
    assert first["agent_directive"].startswith("STOP_RETRY"), first
    if "[추천]" not in first["message"]:
        return first, first          # nothing to follow (2.9.34: L12's mixed '%' target asks the user)
    script2, args2 = _follow(script, args, first["message"])
    _, final = run_raw(script2, *args2)
    return first, final


def _job(fn, *a):
    try:
        return fn(*a)
    except BaseException as e:   # reported by the test of this case
        return e


def _train(path, lid, opts):
    fam = FAMILIES[LAYOUTS[lid][0]]
    return _run_with_one_retry("ml_train.py", ["--file", path, "--features", fam["x"], "--targets", fam["y"], *fam["task"], *opts])


def _predict_command(train_res):
    # 파일 예측 틀 (2.9.75부터 한 건 예측 --input 틀도 있음)
    cmds = [c["command"] for c in train_res.get("next_commands", []) if "ml_predict.py" in c["command"] and " --file " in c["command"]]
    assert len(cmds) == 1, train_res.get("next_commands")
    return cmds[0]


PRED_CASES = [(lid, kind, how) for lid in LAYOUTS for kind in ("units", "plain") for how in ("template", "next")
              if not (kind == "plain" and how == "next")]    # the unit-less models were trained without options


def _predict(folder, lid, kind, how, model_res):
    path = _make(folder, lid, N_PRED, 2)[0]                  # its own copy: outputs are named by the second
    if isinstance(model_res, BaseException):
        raise model_res
    trained = model_res[1]
    if how == "template":
        args = ["--model", trained["version_dir"], "--file", path]
    else:
        args = split_command(_predict_command(trained).replace("<새데이터.xlsx>", path))[2:]
    return _run_with_one_retry("ml_predict.py", args)


def run_matrix(mktemp):
    """매트릭스 전체를 4개씩 실행한다(별도 프로세스, 약 1분): 맞는 옵션으로 학습, 같은 파일을 옵션 없이 학습
    (다음 버전), 첫 학습 모델로 예측. mktemp(name)은 새 폴더를 만든다. 결과는 (첫 결과, 최종 결과)
    또는 그 실행이 낸 예외.
    """
    from concurrent.futures import ThreadPoolExecutor
    files, out = {}, {}
    for lid in LAYOUTS:
        folder = mktemp(lid.split("_")[0] + "t")
        files[lid] = _make(folder, lid, N_TRAIN, 1)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {("a", lid): pool.submit(_job, _train, files[lid][0], lid, LAYOUTS[lid][4]) for lid in LAYOUTS}
        out.update({k: f.result() for k, f in futs.items()})
        futs = {("b", lid): pool.submit(_job, _train, files[lid][0], lid, []) for lid in LAYOUTS}
        for lid, kind, how in PRED_CASES:
            model_lid = MODELS[LAYOUTS[lid][0]][0 if kind == "units" else 1]
            folder = mktemp(lid.split("_")[0] + "p")
            futs[("c", lid, kind, how)] = pool.submit(_job, _predict, folder, lid, kind, how, out[("a", model_lid)])
        out.update({k: f.result() for k, f in futs.items()})
    out["files"] = files
    return out


@pytest.fixture(scope="module")
def matrix_basic(tmp_path_factory):
    return run_matrix(tmp_path_factory.mktemp)


def _result(matrix_basic, key):
    res = matrix_basic[key]
    if isinstance(res, BaseException):
        raise res
    return res


def _trained_layout(res):
    """(피처 + 타겟 이름, 숫자 피처, 학습한 1부터 센 원본 행 번호)"""
    assert res["status"] == "ok" and res["agent_directive"].startswith("VERIFIED"), res
    ft = res["feature_types"]
    names = sorted(list(ft["numeric"]) + list(ft["categorical"]) + list(ft["datetime"]) + list(res["targets"]))
    ds = rt.load_data_file(os.path.join(res["version_dir"], "model_dataset.csv"))
    tgt = ds[res["targets"][0]]
    used = ds.loc[tgt.notna() & (tgt.astype(str).str.strip() != ""), "원본행"]
    return names, sorted(ft["numeric"]), sorted(int(v) for v in used)


def _expect(lid, with_units):
    fam = FAMILIES[LAYOUTS[lid][0]]
    names = _names(lid, with_units)
    numeric = sorted(n for n, c in zip(names, fam["cols"]) if c in ("반응온도", "압력"))
    return sorted(names), numeric


def _mixed_percent_stops(final, lid):
    """L12의 타겟은 '12%'와 일반 숫자가 섞여 배율을 정할 수 없다: STOP_RETRY (2.9.34부터, 전에는 12로 읽음)."""
    if lid != "L12_pct_text":
        return False
    assert final["agent_directive"].startswith("STOP_RETRY:") and "%를 붙인 숫자 칸('12%')" in final["message"], final
    return True


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_with_correct_options(matrix_basic, lid):
    first, final = _result(matrix_basic, ("a", lid))
    if _mixed_percent_stops(final, lid):
        return
    assert first is final, first.get("message")             # no [추천] detour with the right options
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) == _expect(lid, True)
    assert used == matrix_basic["files"][lid][2]


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_without_options(matrix_basic, lid):
    first, final = _result(matrix_basic, ("b", lid))
    if _mixed_percent_stops(final, lid):
        return
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) in (_expect(lid, True), _expect(lid, False)), first.get("message")
    assert used == matrix_basic["files"][lid][2]


@pytest.mark.parametrize("lid,kind,how", PRED_CASES)
def test_predict_same_layout(matrix_basic, lid, kind, how):
    first, final = _result(matrix_basic, ("c", lid, kind, how))
    assert final["status"] == "ok" and final["agent_directive"].startswith("VERIFIED"), (first, final)
    rows = len(range(len(LAYOUTS[lid][1]) + 1, len(LAYOUTS[lid][1]) + N_PRED + 1))
    assert final["sample_count"] == rows, (first.get("message"), final.get("preview"))
    assert len(rt.load_data_file(final["output_file"])) == rows


# ---- 첫 행 측정불가, 좁은 제목 CSV, 반복 머리글, 합계 행 ----------------------------------------------


# 반복 머리글 행은 경고와 함께 뺀다. 합계 행은 남기고('제외' 열에 표시할지는 사용자 판단) 경고에 그 행을 적는다.


def _meas_first(d):
    d[0][0], d[0][1] = "측정불가", "측정불가"
    return d


def _blank_first(d):
    d[0][0], d[0][1], d[0][2] = None, None, None
    return d


def _total(d):
    nums = [r for r in d if all(isinstance(r[k], float) for k in (0, 1, 3))]
    return d + [[round(sum(r[0] for r in nums), 1), round(sum(r[1] for r in nums), 2), "합계",
                 round(sum(r[3] for r in nums), 2)]]


# id: (top rows, data edit, correct options, units of the names or None, extension, repeat the top rows after data row k)
LAYOUTS_B = {
    "M1_meas_first": ([H], _meas_first, [], None, ".xlsx", None),
    "M1c_meas_first_unit": ([H, U], _meas_first, ["--unit-row", "2"], U, ".xlsx", None),
    "N9_blank_first_unit": ([H, U], _blank_first, ["--unit-row", "2"], U, ".xlsx", None),
    "T5_ragged_csv": ([["반응 공정 데이터"], H], None, ["--header-row", "2"], None, ".csv", None),
    "T5b_ragged_csv_unit": ([["시험 기록", "작성자"], H, U], None, ["--header-row", "2", "--unit-row", "3"], U,
                            ".csv", None),
    "T3_title2_hdr": ([["시험 기록", "작성 2026-11-02"], H], None, ["--header-row", "2"], None, ".xlsx", None),
    "R1_repeat_hdr_unit": ([H, U], None, ["--unit-row", "2"], U, ".xlsx", 2),
    "S1_total_row": ([H, U], _total, ["--unit-row", "2"], U, ".xlsx", None),
}


MODELS_B = ("M1c_meas_first_unit", "M1_meas_first")      # trained with units, without units (round a)


def _make_b(folder, lid, n, seed):
    """(파일, 데이터 행의 1부터 센 원본 행 번호, 그 가운데 숫자 타겟이 있는 행)"""
    top, edit, _, _, ext, repeat = LAYOUTS_B[lid]
    data = family_rows("A", n, seed)
    if edit:
        data = edit(data)
    rows, src = [list(r) for r in top], []
    for k, d in enumerate(data):
        if repeat is not None and k == repeat:
            rows += [list(r) for r in top]
        rows.append(d)
        src.append(len(rows))
    path = write_rows(os.path.join(str(folder), lid + ext), rows, [])
    valid = [rw for rw, d in zip(src, data) if rt.coerce_numeric(pd.Series([d[3]])).notna().all()]
    return path, src, valid


def _names_b(lid, with_units):
    units = LAYOUTS_B[lid][3]
    cols = FAMILIES["A"]["cols"]
    return [rt._with_unit(c, u) for c, u in zip(cols, units)] if with_units and units else list(cols)


def _train_b(path, opts):
    return _run_with_one_retry("ml_train.py", ["--file", path, "--features", X, "--targets", Y, *opts])


PRED_CASES_B = [(lid, kind, how) for lid in LAYOUTS_B for kind in ("units", "plain") for how in ("template", "next")
              if not (kind == "plain" and how == "next")]


def _predict_b(folder, lid, how, model_res):
    path = _make_b(folder, lid, N_PRED, 2)[0]
    if isinstance(model_res, BaseException):
        raise model_res
    trained = model_res[1]
    if how == "template":
        args = ["--model", trained["version_dir"], "--file", path]
    else:
        args = split_command(_predict_command(trained).replace("<새데이터.xlsx>", path))[2:]
    return _run_with_one_retry("ml_predict.py", args)


@pytest.fixture(scope="module")
def matrix_marks(tmp_path_factory):
    from concurrent.futures import ThreadPoolExecutor
    mktemp = tmp_path_factory.mktemp
    files, out = {}, {}
    for lid in LAYOUTS_B:
        files[lid] = _make_b(mktemp(lid.split("_")[0] + "t"), lid, N_TRAIN, 1)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {("a", lid): pool.submit(_job, _train_b, files[lid][0], LAYOUTS_B[lid][2]) for lid in LAYOUTS_B}
        out.update({k: f.result() for k, f in futs.items()})
        futs = {("b", lid): pool.submit(_job, _train_b, files[lid][0], []) for lid in LAYOUTS_B}
        for lid, kind, how in PRED_CASES_B:
            model = out[("a", MODELS_B[0 if kind == "units" else 1])]
            futs[("c", lid, kind, how)] = pool.submit(_job, _predict_b, mktemp(lid.split("_")[0] + "p"), lid, how, model)
        out.update({k: f.result() for k, f in futs.items()})
    out["files"] = files
    return out


def _expect_b(lid, with_units):
    names = _names_b(lid, with_units)
    return sorted(names), sorted(n for n, c in zip(names, FAMILIES["A"]["cols"]) if c in ("반응온도", "압력"))


@pytest.mark.parametrize("lid", list(LAYOUTS_B))
def test_marks_train_with_correct_options(matrix_marks, lid):
    first, final = _result(matrix_marks, ("a", lid))
    assert first is final, first.get("message")
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) == _expect_b(lid, True)
    data_rows = matrix_marks["files"][lid][2]
    if lid == "S1_total_row":       # the totals row is kept (its values are the user's decision), with a warning
        assert used in (data_rows, data_rows + [data_rows[-1] + 1])     # it may be an automatic outlier
        assert any("합계" in w and "'제외' 열" in w for w in final.get("warnings", [])), final.get("warnings")
    else:
        assert used == data_rows
    if lid == "R1_repeat_hdr_unit":
        assert any("5, 6행" in w and "머리글" in w for w in final.get("warnings", [])), final.get("warnings")


@pytest.mark.parametrize("lid", [lid for lid in LAYOUTS_B if lid != "S1_total_row"])
def test_marks_train_without_options(matrix_marks, lid):
    first, final = _result(matrix_marks, ("b", lid))
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) in (_expect_b(lid, True), _expect_b(lid, False)), first.get("message")
    assert used == matrix_marks["files"][lid][2]


@pytest.mark.parametrize("lid,kind,how", PRED_CASES_B)
def test_marks_predict_same_layout(matrix_marks, lid, kind, how):
    first, final = _result(matrix_marks, ("c", lid, kind, how))
    assert final["status"] == "ok" and final["agent_directive"].startswith("VERIFIED"), (first, final)
    rows = N_PRED + (1 if lid == "S1_total_row" else 0)
    assert final["sample_count"] == rows, (first.get("message"), final.get("preview"))
    assert len(rt.load_data_file(final["output_file"])) == rows
    if lid == "R1_repeat_hdr_unit":
        assert any("5, 6행" in w for w in final.get("warnings", [])), final.get("warnings")
    if lid == "S1_total_row":
        assert any("합계" in w for w in final.get("warnings", [])), final.get("warnings")


# ---- 비고가 든 단위 행, 배율 숫자 단위 행, 타겟 없는 예측 파일 ------------------------------------------


SCALE_U = ["°C", 1000, None, "%"]


# id: (top rows, correct options, units of the names)
LAYOUTS_C = {
    "N11_blank_hdr_note": ([H + [None], U + ["비고"]], ["--unit-row", "2"], U),
    "L3s_scale_unit": ([H, SCALE_U], ["--unit-row", "2"], SCALE_U),
}


# models for (c): trained with units (header + unit row), without units (header only)
MODEL_LAYOUTS = {"units": ([H, U], ["--unit-row", "2"]), "plain": ([H], [])}


def _make_c(folder, lid, n, seed, top=None):
    """(파일, 데이터 행의 1부터 센 원본 행 번호)"""
    top = LAYOUTS_C[lid][0] if top is None else top
    data = family_rows("A", n, seed)
    path = write_rows(os.path.join(str(folder), lid + ".xlsx"), [list(r) for r in top] + data, [])
    return path, list(range(len(top) + 1, len(top) + n + 1))


def _names_c(lid, with_units):
    units = LAYOUTS_C[lid][2]
    cols = FAMILIES["A"]["cols"]
    return [rt._with_unit(c, u) for c, u in zip(cols, units)] if with_units else list(cols)


def _train_c(path, opts):
    return _run_with_one_retry("ml_train.py", ["--file", path, "--features", X, "--targets", Y, *opts])


PRED_CASES_C = [(lid, kind, how) for lid in LAYOUTS_C for kind in ("units", "plain") for how in ("template", "next")
              if not (kind == "plain" and how == "next")]


def _predict_args(trained, path, how):
    if how == "template":
        return ["--model", trained["version_dir"], "--file", path]
    return split_command(_predict_command(trained).replace("<새데이터.xlsx>", path))[2:]


def _predict_c(folder, lid, how, model_res):
    path = _make_c(folder, lid, N_PRED, 2)[0]
    if isinstance(model_res, BaseException):
        raise model_res
    return _run_with_one_retry("ml_predict.py", _predict_args(model_res[1], path, how))


# prediction files without a target column, first row of missing marks (B1, B1b, B1c)
NOTARGET = {
    "B1_meas_cat": ["측정불가", "측정불가", "A"],
    "B1b_meas_nocat": ["측정불가", "측정불가", None],
    "B1c_ng_cat": ["NG", "NG", "A"],
}


def _make_notarget(folder, bid):
    data = [r[:3] for r in family_rows("A", 6, 2)]
    data[0] = list(NOTARGET[bid])
    return write_rows(os.path.join(str(folder), bid + ".xlsx"), [H[:3]] + data, [])


def _predict_notarget(folder, bid, model_res):
    if isinstance(model_res, BaseException):
        raise model_res
    path = _make_notarget(folder, bid)
    return run_raw("ml_predict.py", "--model", model_res[1]["version_dir"], "--file", path)[1]


@pytest.fixture(scope="module")
def matrix_note_scale(tmp_path_factory):
    from concurrent.futures import ThreadPoolExecutor
    mktemp = tmp_path_factory.mktemp
    files, out = {}, {}
    for lid in LAYOUTS_C:
        files[lid] = _make_c(mktemp(lid.split("_")[0] + "t"), lid, N_TRAIN, 1)
    for kind, (top, _) in MODEL_LAYOUTS.items():
        files[kind] = _make_c(mktemp("m" + kind), "M_" + kind, N_TRAIN, 1, top)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {("a", lid): pool.submit(_job, _train_c, files[lid][0], LAYOUTS_C[lid][1]) for lid in LAYOUTS_C}
        futs.update({("m", kind): pool.submit(_job, _train_c, files[kind][0], opts)
                     for kind, (_, opts) in MODEL_LAYOUTS.items()})
        out.update({k: f.result() for k, f in futs.items()})
        futs = {("b", lid): pool.submit(_job, _train_c, files[lid][0], []) for lid in LAYOUTS_C}
        for lid, kind, how in PRED_CASES_C:
            futs[("c", lid, kind, how)] = pool.submit(_job, _predict_c, mktemp(lid.split("_")[0] + "p"), lid, how,
                                                      out[("m", kind)])
        for bid in NOTARGET:
            for kind in MODEL_LAYOUTS:
                futs[("n", bid, kind)] = pool.submit(_job, _predict_notarget, mktemp(bid.split("_")[0]), bid,
                                                     out[("m", kind)])
        out.update({k: f.result() for k, f in futs.items()})
    out["files"] = files
    return out


def _expect_c(lid, with_units):
    names = _names_c(lid, with_units)
    return sorted(names), sorted(n for n, c in zip(names, FAMILIES["A"]["cols"]) if c in ("반응온도", "압력"))


@pytest.mark.parametrize("lid", list(LAYOUTS_C))
def test_note_scale_train_with_correct_options(matrix_note_scale, lid):
    first, final = _result(matrix_note_scale, ("a", lid))
    assert first is final, first.get("message")     # 2.9.18: both refused as "not a unit row"
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) == _expect_c(lid, True)
    assert used == matrix_note_scale["files"][lid][1]


@pytest.mark.parametrize("lid", list(LAYOUTS_C))
def test_note_scale_train_without_options(matrix_note_scale, lid):
    first, final = _result(matrix_note_scale, ("b", lid))
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) in (_expect_c(lid, True), _expect_c(lid, False)), first.get("message")
    assert used == matrix_note_scale["files"][lid][1]


@pytest.mark.parametrize("lid,kind,how", PRED_CASES_C)
def test_note_scale_predict_same_layout(matrix_note_scale, lid, kind, how):
    first, final = _result(matrix_note_scale, ("c", lid, kind, how))
    assert final["status"] == "ok" and final["agent_directive"].startswith("VERIFIED"), (first, final)
    assert final["sample_count"] == N_PRED, (first.get("message"), final.get("preview"))   # 2.9.18: 6 (unit row)
    assert len(rt.load_data_file(final["output_file"])) == N_PRED


@pytest.mark.parametrize("bid", list(NOTARGET))
@pytest.mark.parametrize("kind", list(MODEL_LAYOUTS))
def test_predict_first_row_missing_marks_without_target(matrix_note_scale, bid, kind):
    res = _result(matrix_note_scale, ("n", bid, kind))
    # 2.9.18: '[추천] --data-row 3' (a real data row dropped) for B1 / B1c, B1b
    assert res["status"] == "ok" and res["agent_directive"].startswith("VERIFIED"), res
    assert res["sample_count"] == 6


def test_train_unit_row_option_on_missing_marks_row(tmp_path):
    data = family_rows("A", 20, 1)
    data[0] = ["측정불가", "측정불가", "A", None]
    path = write_rows(str(tmp_path / "b2.xlsx"), [H] + data, [])
    first, final = _train_c(path, ["--unit-row", "2"])
    assert first["status"] == "error" and first["agent_directive"].startswith("STOP_RETRY"), first
    assert "--unit-row" not in first["message"].split("[추천]")[1], first["message"]    # 2.9.18: VERIFIED
    names, numeric, used = _trained_layout(final)
    assert names == sorted(H) and used == list(range(3, 22))
