# -*- coding: utf-8 -*-
"""학습 데이터셋 만들기 (ml_dataset.py, 2.9.85).

pick(맞는 행만 새 파일), exclude(모든 행 + 맞는 행의 '제외' 열에 이유), --where / --rows, CSV 기본·xlsx, 기본 이름 <원본>_뽑기N.csv,
원본은 그대로, 다시 읽어 확인, 0행, --dry-run, 다음 학습 명령(원본의 마지막 학습과 같은 타겟·입력).
조건 식 자체(코드 열, 쉼표, [후보])와 exclude 뒤 학습은 test_train.py에 있다.
"""
import os
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import ok_res, run_fail, run_raw, run_tool, split_command  # noqa: E402


def _data(path, n=80, seed=1):
    r = np.random.default_rng(seed)
    df = pd.DataFrame({"반응온도(°C)": r.uniform(160, 200, n).round(2), "압력": r.uniform(1, 3, n).round(3),
                       "라인": ["A", "B"] * (n // 2)})
    df["수율(%)"] = (0.3 * df["반응온도(°C)"] + 5 * df["압력"] + r.normal(0, 0.3, n)).round(3)
    if str(path).endswith(".xlsx"):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    return df


def test_pick_writes_only_matching_rows_and_keeps_the_source(tmp_path):
    f = tmp_path / "공정.csv"
    df = _data(f)
    before = f.read_bytes()
    proc, res = run_raw("ml_dataset.py", "pick", "--file", str(f), "--where", "반응온도 >= 180; 라인 = B")
    ok_res(res, proc)
    want = df[(df["반응온도(°C)"] >= 180) & (df["라인"] == "B")]
    assert res["count"] == len(want) > 0 and res["rows_saved"] == len(want) and res["verified"] is True, res
    assert res["output_file"] == str(tmp_path / "공정_뽑기1.csv") and f.read_bytes() == before
    assert res["chosen_by"] == "조건 반응온도(°C) >= 180; 라인 = B" and res["agent_directive"].startswith("VERIFIED:")
    out = pd.read_csv(res["output_file"], encoding="utf-8-sig")
    assert list(out.columns) == list(df.columns) and np.allclose(out["수율(%)"], want["수율(%)"])
    assert res["rows"] == [i + 2 for i in want.index][:20] and "원본은 그대로" in res["preview"]
    again = run_tool("ml_dataset.py", "pick", "--file", str(f), "--where", "라인 = A")
    assert again["output_file"] == str(tmp_path / "공정_뽑기2.csv")                  # the next free number


def test_exclude_marks_rows_by_number_and_xlsx_output(tmp_path):
    f = tmp_path / "d.xlsx"
    df = _data(f)
    res = run_tool("ml_dataset.py", "exclude", "--file", str(f), "--rows", "3,10", "--reason", "센서 고장", "--save", "표시.xlsx")
    assert res["output_file"] == str(tmp_path / "표시.xlsx") and res["marked"] == 2 and res["rows_saved"] == len(df)
    out = pd.read_excel(res["output_file"])
    assert list(out.columns) == list(df.columns) + ["제외"]
    marked = out.index[out["제외"].fillna("") != ""].tolist()
    assert marked == [1, 8] and set(out.loc[marked, "제외"]) == {"센서 고장"}                # rows 3 and 10 of the sheet
    bad = run_fail("ml_dataset.py", "exclude", "--file", str(f), "--rows", "1,500")
    assert "데이터 행이 아닙니다" in bad["message"] and "[1, 500]" in bad["message"], bad
    bad = run_fail("ml_dataset.py", "exclude", "--file", str(f), "--rows", "3-5")
    assert "범위는 쓰지 않음" in bad["message"], bad


def test_exclude_keeps_existing_marks(tmp_path):
    f = tmp_path / "m.csv"
    df = _data(f)
    df["제외"] = ""
    df.loc[0, "제외"] = "포함"
    df.loc[1, "제외"] = "오타"
    df.to_csv(f, index=False, encoding="utf-8-sig")
    res = run_tool("ml_dataset.py", "exclude", "--file", str(f), "--rows", "2,3,4")
    assert res["count"] == 3 and res["marked"] == 1 and any("이미 표시" in w for w in res["warnings"]), res
    out = pd.read_csv(res["output_file"], encoding="utf-8-sig").fillna("")
    assert out["제외"].tolist()[:4] == ["포함", "오타", "행 2, 3, 4", ""]
    bad = run_fail("ml_dataset.py", "exclude", "--file", str(f), "--rows", "5", "--reason", "포함")
    assert "포함" in bad["message"], bad


def test_zero_rows_dry_run_and_options(tmp_path):
    f = tmp_path / "z.csv"
    _data(f)
    none = run_tool("ml_dataset.py", "pick", "--file", str(f), "--where", "반응온도 > 999")
    assert none["count"] == 0 and "output_file" not in none and "찾은 행 없음" in none["agent_directive"]
    assert not list(tmp_path.glob("z_*.csv"))
    dry = run_tool("ml_dataset.py", "exclude", "--file", str(f), "--where", "라인 = A", "--dry-run")
    assert dry["dry_run"] is True and not os.path.exists(dry["output_file"]) and dry["count"] == 40
    argv = split_command(dry["confirm_commands"][0]["command"])
    done = run_tool("ml_dataset.py", *argv[2:])
    assert done["output_file"] == dry["output_file"] and done["marked"] == 40
    for extra in ((), ("--where", "라인 = A", "--rows", "2")):
        proc, out = run_raw("ml_dataset.py", "pick", "--file", str(f), *extra)
        assert proc.returncode == 1 and "하나만" in out["message"], out
    proc, out = run_raw("ml_dataset.py", "copy", "--file", str(f), "--rows", "2")
    assert proc.returncode == 2 and out["status"] == "error"
    bad = run_fail("ml_dataset.py", "pick", "--file", str(f), "--rows", "2", "--save", str(f))
    assert "원본" in bad["message"]
    warn = run_tool("ml_dataset.py", "pick", "--file", str(f), "--rows", "2", "--reason", "x")
    assert any("--reason" in w for w in warn["warnings"])


def test_next_command_trains_the_new_file_like_the_last_training(tmp_path):
    f = tmp_path / "t.csv"
    _data(f)
    plain = run_tool("ml_dataset.py", "pick", "--file", str(f), "--where", "라인 = A")
    assert plain["next_commands"][0]["needs_input"] is True and "<타겟>" in plain["next_commands"][0]["command"]
    run_tool("ml_train.py", "--file", str(f), "--features", "반응온도,압력", "--targets", "수율", "--eval", "tp = 반응온도 * 압력")
    res = run_tool("ml_dataset.py", "pick", "--file", str(f), "--where", "라인 = A")
    nxt = res["next_commands"][0]
    assert nxt["needs_input"] is False and "v1 학습과 같은" in nxt["purpose"], nxt
    argv = split_command(nxt["command"])
    assert argv[argv.index("--file") + 1] == res["output_file"]
    assert argv[argv.index("--features") + 1] == "반응온도(°C),압력" and argv[argv.index("--targets") + 1] == "수율(%)"
    assert argv[argv.index("--eval") + 1] == "tp = 반응온도(°C) * 압력"
    tr = run_tool("ml_train.py", *argv[2:])
    assert tr["version"] == "v1" and os.path.basename(tr["data_root"]) == "t_뽑기2_ml_data"   # its own folder, v1


@pytest.mark.parametrize("ext", [".csv", ".xlsx"])
def test_saved_file_reads_back_like_the_source(tmp_path, ext):
    """다시 읽기 확인: 날짜·글자·빈 칸·퍼센트가 섞여도 저장한 값이 원본과 같다 (CSV, xlsx 모두)."""
    f = tmp_path / f"s{ext}"
    df = pd.DataFrame({"일시": pd.date_range("2026-01-01", periods=6, freq="D"), "값": [1.5, None, 3.25, 4.0, 5.0, 6.0],
                       "메모": ["a", "", "b c", "d", "e", "f"], "y": [1, 2, 3, 4, 5, 6]})
    if ext == ".csv":
        df.to_csv(f, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(f, index=False)
    res = run_tool("ml_dataset.py", "pick", "--file", str(f), "--where", "y >= 2")
    assert res["verified"] is True and res["rows_saved"] == 5, res
