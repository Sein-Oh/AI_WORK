import os
import re
import sys

import pytest

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(TESTS_DIR))), "CLI", "office-cli")  # 실행 파일
# 저장소 루트 pytest.ini는 --import-mode=importlib를 쓰므로 테스트 폴더(fake_*, office_helpers)와 도구 폴더를 직접 경로에 등록
for p in (ROOT, TESTS_DIR):
    if p not in sys.path:
        sys.path.insert(0, p)

import _office as off  # noqa: E402
from fake_excel import FakeExcelWorkbook, FakeXlApp  # noqa: E402
from fake_ppt import FakePptWindowApp, FakePresentation  # noqa: E402
from fake_word import FakeWordDocument  # noqa: E402
from office_helpers import TABLE, set_active, stub_com, unsafe_next_commands  # noqa: E402


@pytest.fixture(autouse=True)
def isolated_dirs(tmp_path, monkeypatch):
    """캐시/데이터 폴더를 테스트별 임시 폴더로 격리합니다."""
    cache = tmp_path / "cache"
    data = tmp_path / "data"
    cache.mkdir()
    data.mkdir()
    monkeypatch.setenv("OFFICE_CLI_CACHE_DIR", str(cache))
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(data))
    monkeypatch.setenv("OFFICE_CLI_TIME_LIMIT", "0")      # R10: 테스트는 시간 제한 없음 (CONTINUE 시험만 따로 켬)
    return {"cache": cache, "data": data}


# next_commands = 묻지 않고 실행해도 되는 것(읽기·미리보기)만 (1.4.75, 11차 검토). 문서를 바꾸는 이어 하기는 confirm_commands.
# 모든 테스트의 stdout JSON을 보고, 바꾸는 명령이 next_commands에 들면 그 테스트를 실패시킨다.
@pytest.fixture(autouse=True)
def _next_commands_are_safe(monkeypatch):
    import _runtime as rt
    seen = []
    orig = rt.emit_json

    def emit(payload):
        seen.extend(unsafe_next_commands(payload))
        return orig(payload)
    monkeypatch.setattr(rt, "emit_json", emit)
    yield
    assert not seen, f"next_commands에 문서를 바꾸는 명령 (confirm_commands로): {seen}"


# ---------------------------------------------------------------------------
# 활성 문서가 가짜 Office 문서인 상태 (win32com도 가짜, 열린 문서 목록은 비어 있음)
# ---------------------------------------------------------------------------

def _excel(monkeypatch, name, sheets):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, name, sheets=sheets)
    stub_com(monkeypatch)
    set_active(monkeypatch, app, wb)
    return app, wb


@pytest.fixture
def xl(monkeypatch):
    """활성 통합문서 '보고.xlsx' (Sheet1, 요약). (app, wb) 반환."""
    return _excel(monkeypatch, "보고.xlsx", ("Sheet1", "요약"))


@pytest.fixture
def xl_ab(monkeypatch):
    """활성 통합문서 '보고.xlsx' (Sheet1, a_b): 밑줄이 든 시트 이름. (app, wb) 반환."""
    return _excel(monkeypatch, "보고.xlsx", ("Sheet1", "a_b"))


@pytest.fixture
def xl_s2(monkeypatch):
    """활성 통합문서 'b.xlsx' (Sheet1, S2). (app, wb) 반환."""
    return _excel(monkeypatch, "b.xlsx", ("Sheet1", "S2"))


@pytest.fixture
def xl_table(monkeypatch):
    """활성 통합문서 '보고.xlsx'의 Sheet1에 TABLE(머리글 + 4행)을 채운 상태. 시트(ws) 반환."""
    _, wb = _excel(monkeypatch, "보고.xlsx", ("Sheet1",))
    ws = wb.Worksheets("Sheet1")
    for i, row in enumerate(TABLE, start=1):
        for j, v in enumerate(row, start=1):
            ws.values[(i, j)] = v
    return ws


@pytest.fixture
def wd(monkeypatch):
    """활성 워드 문서 (추진 배경 / 2026년 계획입니다. / 일정). holder['doc']로 바꿔 끼울 수 있음, 스냅샷은 만들지 않음."""
    holder = {"doc": FakeWordDocument("추진 배경\r2026년 계획입니다.\r일정")}
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": holder["doc"].Name, "path": holder["doc"].FullName,
                                                          "app": object(), "doc": holder["doc"]})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    return holder


@pytest.fixture
def pp(monkeypatch):
    """활성 프레젠테이션 (슬라이드 3장). holder['prs']로 바꿔 끼울 수 있고 holder['app'].ActiveWindow.View.goto에
    화면 이동이 남음. 스냅샷은 만들지 않음."""
    holder = {"prs": FakePresentation(3)}
    holder["app"] = FakePptWindowApp(holder["prs"])
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)

    def active(kind):
        holder["app"].ActiveWindow.Presentation = holder["prs"]
        return {"name": holder["prs"].Name, "path": holder["prs"].FullName, "app": holder["app"], "doc": holder["prs"]}
    monkeypatch.setattr(off, "get_active", active)
    return holder
