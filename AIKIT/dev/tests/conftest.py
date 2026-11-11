# -*- coding: utf-8 -*-
"""AIKIT/dev/tests 공용 pytest 설정."""


# 매개변수 id 길이 상한: pytest는 테스트 id를 환경변수 PYTEST_CURRENT_TEST에 넣는데, Windows는 환경변수 한 개가
# 32767자를 넘으면 ValueError로 그 테스트가 ERROR가 된다 (회사 PC 실기 2026-11-05). 긴 글자 값은 앞부분과 길이만 id로.
PARAM_ID_MAX = 60


def pytest_make_parametrize_id(config, val, argname):
    if isinstance(val, (str, bytes)):
        text = val.decode("latin-1") if isinstance(val, bytes) else val
        esc = text.encode("unicode_escape").decode("ascii")
        if len(esc) > PARAM_ID_MAX:
            return f"{esc[:PARAM_ID_MAX - 12]}..len{len(text)}"
    return None
