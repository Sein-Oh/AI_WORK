# -*- coding: utf-8 -*-
"""bom-kit 시험 설정: 도우미(bk_helpers.py)의 env 고정 장치."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bk_helpers import env, env_pk  # noqa: E402,F401
