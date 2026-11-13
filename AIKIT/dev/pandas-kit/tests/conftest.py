# -*- coding: utf-8 -*-
"""pandas-kit 시험 설정: 도우미(pk_helpers.py)의 env 고정 장치."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pk_helpers import env  # noqa: E402,F401
