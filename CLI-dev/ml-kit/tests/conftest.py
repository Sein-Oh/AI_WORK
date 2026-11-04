# -*- coding: utf-8 -*-
"""ml-kit 테스트 설정과 여러 테스트 파일이 같이 쓰는 fixture.

- 오류 로그(R6)가 실제 도구 폴더에 쌓이지 않게 세션 동안 ML_KIT_DATA_DIR을 임시 폴더로 둔다.
  테스트가 띄우는 하위 프로세스도 이 값을 물려받는다.
- 학습 한 번이 드는 모델과 데이터 파일 묶음은 세션에 한 번만 만든다 (같은 입력을 쓰는 테스트 파일이 둘 이상).
"""
import os
import random
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import run_raw, run_tool, write_rows  # noqa: E402


@pytest.fixture(scope="session", autouse=True)
def _mlkit_data_dir(tmp_path_factory):
    old = os.environ.get("ML_KIT_DATA_DIR")
    os.environ["ML_KIT_DATA_DIR"] = str(tmp_path_factory.mktemp("mlkit_data"))
    # ml_report.py가 성능예측표(재학습, 몇 분)까지 만들지 않게 한다: test_report_export.py의 성능예측표 검사만 켠다
    os.environ["ML_KIT_SKIP_EXPORT"] = "1"
    # 명령 시간 제한(2.9.40)은 끔: 세션이 길어도 학습이 멈추지 않게. 시간 제한 테스트는 값을 직접 줌
    os.environ["ML_KIT_TIME_LIMIT"] = "0"
    yield os.environ["ML_KIT_DATA_DIR"]
    if old is None:
        os.environ.pop("ML_KIT_DATA_DIR", None)
    else:
        os.environ["ML_KIT_DATA_DIR"] = old


@pytest.fixture(scope="session")
def p1_model(tmp_path_factory):
    """온도(°C), 압력(bar), 속도 -> 생산량 회귀 모델 (test_names, test_predict). {"v": 버전 폴더, "f": 데이터 파일}"""
    tmp = tmp_path_factory.mktemp("p1")
    r = random.Random(1)
    rows = [["온도(°C)", "압력(bar)", "속도", "생산량"]]
    for _ in range(60):
        t, p, s = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1)
        rows.append([t, p, s, round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    f = write_rows(os.path.join(str(tmp), "p1.csv"), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도(°C),압력(bar),속도", "-y", "생산량")
    assert res["status"] == "ok", res
    return {"v": os.path.join(str(tmp), "p1_ml_data", "v1"), "f": f}


@pytest.fixture(scope="session")
def coded(tmp_path_factory):
    """숫자 a, b와 앞자리 0이 있는 범주 코드('007'이 7이 되면 안 됨) -> y (test_predict, test_train)."""
    work = tmp_path_factory.mktemp("coded")
    r = np.random.default_rng(11)
    n = 240
    codes = np.array(["007", "012", "101"])[r.integers(0, 3, n)]
    df = pd.DataFrame({"a": r.uniform(0, 10, n), "b": r.uniform(0, 5, n), "코드": codes})
    df["y"] = 3 * df.a + 2 * df.b + np.select([df.코드 == "007", df.코드 == "012"], [20.0, 0.0], -20.0) + r.normal(0, 0.2, n)
    p = work / "coded.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_train.py", "-f", str(p), "-x", "a,b,코드", "-y", "y", "--dtypes", "코드:str")
    return {"work": work, "file": str(p), "v": res["version_dir"], "train": res}


@pytest.fixture(scope="session")
def name_files(tmp_path_factory):
    """비슷한 이름·% 숫자·숫자 이름 열·일정한 입력 데이터 파일 (학습 없음; test_names, test_loading_numbers,
    test_report_export)."""
    tmp = tmp_path_factory.mktemp("names")
    r = random.Random(7)
    d = [["온도(°C)", "압력(bar)", "속도", "습도(%)", "수율(1차)", "수율(최종)", "생산량"]]
    pct = [["온도(°C)", "압력(bar)", "습도(%)", "수율(%)"]]
    plain = [["온도(°C)", "압력(bar)", "습도(%)", "수율(%)"]]
    acc = [["가속도", "압력(bar)", "함수율(%)", "진동"]]
    num = [["온도", "2", "수율"]]
    const = [["온도", "압력", "라인", "수율"]]
    for _ in range(60):
        t, p, s, h = (round(r.uniform(150, 200), 2), round(r.uniform(1, 5), 3), round(r.uniform(10, 30), 2),
                      round(r.uniform(30, 70), 1))
        y1 = round(0.3 * t + 5 * p + r.uniform(-1, 1), 2)
        y2 = round(y1 + s / 5 + r.uniform(-0.5, 0.5), 2)
        d.append([t, p, s, h, y1, y2, round(y2 * 10, 1)])
        yp = round(0.2 * t + 0.3 * h + r.uniform(-1, 1), 2)
        pct.append([t, p, f"{h}%", f"{yp}%"])
        plain.append([t, p, h, yp])
        a, w = round(r.uniform(1, 4), 2), round(r.uniform(5, 12), 2)
        acc.append([a, p, w, round(5 * a + w + r.uniform(-1, 1), 2)])
        num.append([t, round(r.uniform(0, 9), 2), y1])
        const.append([t, p, 5, y1])
    out = {"tmp": tmp}
    for name, rows in (("d", d), ("pct", pct), ("plain", plain), ("acc", acc), ("num", num), ("const", const)):
        out[name] = write_rows(os.path.join(str(tmp), f"{name}.csv"), rows, [])
    return out


@pytest.fixture(scope="session")
def percent_files(tmp_path_factory):
    """-y 이름을 품은 열, % 숫자, %와 일반 숫자가 섞인 열, 숫자 이름 열 데이터 파일 (학습 없음; test_names,
    test_loading_numbers)."""
    tmp = tmp_path_factory.mktemp("percent")
    r = random.Random(11)
    cool = [["냉각수온도", "유량", "온도"]]                          # -y 온도 옆의 냉각수온도
    moist = [["온도", "함수율", "수율"]]
    moist_u = [["반응온도(°C)", "함수율(%)", "수율(최종)"]]
    d = [["온도(°C)", "압력(bar)", "수율(1차)", "수율(최종)", "생산량"]]
    grid_pct = [["온도", "함수율", "수율"]]                         # 모든 칸이 '%'
    grid = [["온도", "함수율", "수율"]]
    mix = [["온도", "함수율", "수율"]]                              # '11.9%' 옆에 '0.077'
    mix_y = [["온도", "함수율", "수율"]]
    sieve = [["온도", "0.5(mm)", "2(mm)", "수율"]]
    for i in range(40):
        t, w = round(r.uniform(150, 200), 1), round(r.uniform(5, 15), 1)
        c, f = round(r.uniform(15, 30), 2), round(r.uniform(1, 5), 3)
        cool.append([c, f, round(60 + 0.8 * c + 2 * f + r.uniform(-1, 1), 2)])
        y = round(0.3 * t - w + r.uniform(-1, 1), 2)
        moist.append([t, w, y])
        moist_u.append([t, w, y])
        p = round(r.uniform(1, 5), 3)
        y1 = round(0.3 * t + 5 * p + r.uniform(-1, 1), 2)
        d.append([t, p, y1, round(y1 + 2, 2), round(y1 * 10, 1)])
        mix.append([t, f"{w}%" if i % 2 else round(w / 100, 4), y])
        mix_y.append([t, w, f"{y}%" if i == 3 else y])
        a, b = round(r.uniform(10, 30), 2), round(r.uniform(1, 30), 2)
        sieve.append([t, a, b, round(t + a + b + r.uniform(-1, 1), 2)])
    for t in (150, 170, 190):
        for w in (5, 10, 15):
            grid_pct.append([t, f"{w}%", 0.2 * t - w])
            grid.append([t, w, 0.2 * t - w])
    out = {"tmp": tmp}
    for name, rows in (("cool", cool), ("moist", moist), ("moist_u", moist_u), ("d", d), ("grid_pct", grid_pct),
                       ("grid", grid), ("mix", mix), ("mix_y", mix_y), ("sieve", sieve)):
        out[name] = write_rows(os.path.join(str(tmp), f"{name}.csv"), rows, [])
    return out


@pytest.fixture(scope="session")
def near_name_files(tmp_path_factory):
    """함수율 옆 실측값 열, 냉각수온도, 숫자 이름 열, '제외' 행의 % 칸 데이터 파일 (학습 없음; test_names,
    test_loading_numbers)."""
    tmp = tmp_path_factory.mktemp("near")
    r = random.Random(7)
    grid = [["온도", "함수율", "수율", "생산량"]]
    cool = [["시간", "냉각수온도", "온도"]]
    sieve = [["온도", "0.5(mm)", "2(mm)", "수율"]]
    excl = [["온도", "함수율", "수율", "제외"]]
    for t in (150, 170, 190):
        for w in (5, 10, 15):
            grid.append([t, w, 0.2 * t - w, 2 * t + w])
            cool.append([t, w + 15, 0.2 * t - w])
    for i in range(40):
        t, w = round(r.uniform(150, 200), 1), round(r.uniform(5, 15), 1)
        a, b = round(r.uniform(10, 30), 2), round(r.uniform(1, 30), 2)
        sieve.append([t, a, b, round(t + a + b + r.uniform(-1, 1), 2)])
        y = round(0.3 * t - w + r.uniform(-1, 1), 2)
        excl.append([t, f"{w}%" if i == 4 else w, y, "측정 오류" if i == 4 else ""])
    out = {"tmp": tmp}
    for name, rows in (("grid", grid), ("cool", cool), ("sieve", sieve), ("excl", excl)):
        out[name] = write_rows(os.path.join(str(tmp), f"{name}.csv"), rows, [])
    return out
