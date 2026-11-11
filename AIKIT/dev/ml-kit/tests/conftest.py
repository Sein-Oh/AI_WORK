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


# 매개변수 id 길이 상한: pytest는 테스트 id를 환경변수 PYTEST_CURRENT_TEST에 넣는데, Windows는 환경변수 한 개가
# 32767자를 넘으면 ValueError로 그 테스트가 ERROR가 된다 (회사 PC 실기 2026-11-05). 긴 글자 값은 앞부분과 길이만 id로.
PARAM_ID_MAX = 60


# 평소(날마다) 실행에서 빼는 무거운 시험 (사용자 결정 2026-11-12: 전체가 리눅스 26분·회사 PC 1시간 이상).
# - report 표시(보고서·성능예측표), 읽기 배치 조합 파일 3개, slow_tests.txt의 함수(한 번에 2초 이상 걸린 회귀 시험).
# - 전부 돌리려면 환경 변수 AIKIT_FULL_TEST=1 (PR 전·릴리스 전, run_all_tests.py --full). -m을 직접 주면 그 선택을 따름.
FULL_ENV = "AIKIT_FULL_TEST"
HEAVY_FILES = ("test_layout_matrix.py", "test_loading_layout.py", "test_loading_numbers.py")
SLOW_LIST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "slow_tests.txt")


def slow_names():
    with open(SLOW_LIST, encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip() and not line.startswith("#")}


def is_heavy(item, names):
    base = os.path.basename(str(item.fspath))
    return (base in HEAVY_FILES or item.get_closest_marker("report") is not None
            or f"{base}::{getattr(item, 'originalname', item.name)}" in names)


def pytest_collection_modifyitems(config, items):
    if os.environ.get(FULL_ENV) == "1" or config.getoption("markexpr", ""):
        return
    names = slow_names()
    keep, skipped = [], []
    for item in items:
        (skipped if is_heavy(item, names) else keep).append(item)
    if skipped:
        config.hook.pytest_deselected(items=skipped)
        items[:] = keep


def pytest_make_parametrize_id(config, val, argname):
    if isinstance(val, (str, bytes)):
        text = val.decode("latin-1") if isinstance(val, bytes) else val
        esc = text.encode("unicode_escape").decode("ascii")
        if len(esc) > PARAM_ID_MAX:
            return f"{esc[:PARAM_ID_MAX - 12]}..len{len(text)}"
    return None


@pytest.fixture(scope="session", autouse=True)
def _mlkit_data_dir(tmp_path_factory):
    old = os.environ.get("ML_KIT_DATA_DIR")
    os.environ["ML_KIT_DATA_DIR"] = str(tmp_path_factory.mktemp("mlkit_data"))
    # ml_report.py가 성능예측표(재학습, 몇 분)까지 만들지 않게 한다: test_report_export.py의 성능예측표 검사만 켠다
    os.environ["ML_KIT_SKIP_EXPORT"] = "1"
    # 성능예측표의 엑셀 다시 계산 확인(Windows + 엑셀)은 test_excel_recalculates_like_the_model 하나만 (회사 PC 2026-11-09:
    # 성능예측표를 만드는 시험마다 엑셀을 열어 pytest가 62분)
    os.environ["ML_KIT_SKIP_EXCEL_CHECK"] = "1"
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
    _, res = run_raw("ml_train.py", "--file", f, "--features", "온도(°C),압력(bar),속도", "--targets", "생산량")
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
    res = run_tool("ml_train.py", "--file", str(p), "--features", "a,b,코드", "--targets", "y", "--dtypes", "코드:str")
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
    """--targets 이름을 품은 열, % 숫자, %와 일반 숫자가 섞인 열, 숫자 이름 열 데이터 파일 (학습 없음; test_names,
    test_loading_numbers)."""
    tmp = tmp_path_factory.mktemp("percent")
    r = random.Random(11)
    cool = [["냉각수온도", "유량", "온도"]]                          # --targets 온도 옆의 냉각수온도
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
