# -*- coding: utf-8 -*-
"""bom-kit 시험 도우미: bom.py를 서브프로세스로 실행해 stdout 1줄 JSON을 읽는다.
입력 csv는 pandas-kit으로 만든다 (pandas-kit 시험 도우미 pk_helpers의 env·ok를 함께 씀)."""
import json
import os
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(os.path.dirname(HERE))                            # AIKIT/dev
AIKIT = os.path.dirname(DEV)
BOM = os.path.join(AIKIT, "bom-kit", "bom.py")
SAMPLES = os.path.join(DEV, "pandas-kit", "tools")                      # make_mbom_samples.py (pandas-kit 시험도 씀)
sys.path.insert(0, os.path.join(DEV, "pandas-kit", "tests"))
sys.path.insert(0, SAMPLES)
import pk_helpers  # noqa: E402
from pk_helpers import command_args, ok, recommended, run  # noqa: E402,F401


@pytest.fixture
def env(env_pk, monkeypatch):
    """pandas-kit의 env(데이터 폴더, bash 따옴표)에 bom-kit도 같게."""
    monkeypatch.setenv("BOM_KIT_DATA_DIR", str(env_pk["dir"] / "bomdata"))
    monkeypatch.setenv("BOM_KIT_SHELL", "bash")
    return env_pk


env_pk = pk_helpers.env


def bom(*args):
    """(exit code, JSON). stdout은 정확히 1줄이어야 한다."""
    p = subprocess.run([sys.executable, BOM, *map(str, args)], capture_output=True, text=True, encoding="utf-8",
                       env=os.environ.copy())
    lines = [l for l in p.stdout.splitlines() if l.strip()]
    assert len(lines) == 1, (p.stdout, p.stderr)
    return p.returncode, json.loads(lines[0])


def make_samples(out):
    p = subprocess.run([sys.executable, os.path.join(SAMPLES, "make_mbom_samples.py"), "--out", str(out)],
                       capture_output=True, text=True, encoding="utf-8")
    assert p.returncode == 0, p.stderr
    return out
