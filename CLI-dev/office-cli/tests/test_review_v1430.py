"""1.4.30 (21차 점검) 회귀 테스트.

L1: outlook list --folder '<계정>\\없는폴더'의 오류 로그 '결과:' 줄에 메일 계정 이름(메일 주소)이 남았음
    (not_found의 extra ' (경로: ...)'가 --folder 원문 그대로였음). 1.4.29 가드 테스트는 소문자로 넣고 대소문자를 구분해 확인해
    실패할 수 없었음.
L2: transfer --method relay --to ppt가 --slide를 확인하기 전에 클립보드에 복사하고 임시 워드를 띄웠음
    (2장짜리에 --slide 9: 클립보드를 덮어쓴 뒤 오류).
L3: transfer --method relay --to-target ""가 KeyError('word')를 STOP_RETRY로 냈음. 빈 대상은 활성 문서 (export·redraw와 같음).
"""
import contextlib
import json

import pytest

import _office as off
import _runtime as rt
import office
import outlook
import test_office_cmd
import test_review_v1429 as v1429

ROOT_MAIL = v1429.ROOT_MAIL
mail = v1429.mail                       # fixture (메일 계정 ROOT_MAIL 아래 폴더 몇 개)
chart_env = test_office_cmd.chart_env   # fixture (엑셀 차트 하나)
run = test_office_cmd.run
_real_relay = office.relay_chart


def _result_lines(d):
    return [ln for ln in v1429._log(d).splitlines() if ln.startswith("결과:")]


# L1: 경로(extra)의 메일 계정 이름도 로그에 남기지 않음 ---------------------------------------------------------------

@pytest.mark.parametrize("root", [ROOT_MAIL, ROOT_MAIL.lower()])
def test_folder_path_extra_masks_account_in_log(mail, capsys, monkeypatch, isolated_dirs, root):
    code, out = v1429._main(outlook, capsys, monkeypatch, "list", "--folder", f"{root}\\없음")
    assert code == 1 and f"(경로: {root}\\없음)" in out["message"], out      # 사용자에게는 준 경로 그대로
    res = "\n".join(_result_lines(isolated_dirs))
    assert res and ROOT_MAIL.lower() not in res.lower(), res
    assert "(경로: <메일 계정>\\없음)" in res, res


def test_folder_path_extra_keeps_alias_path_in_log(mail, capsys, monkeypatch, isolated_dirs):
    # 기본 폴더 별칭으로 시작한 경로는 계정 이름이 아니므로 그대로 남김
    code, out = v1429._main(outlook, capsys, monkeypatch, "list", "--folder", "inbox\\없음")
    assert code == 1, out
    res = "\n".join(_result_lines(isolated_dirs))
    assert "(경로: inbox\\없음)" in res and ROOT_MAIL.lower() not in res.lower(), res


def test_not_found_log_extra():
    e = rt.not_found("메일 폴더", "없음", [], extra=" (경로: 계정@x.com\\없음)", log_extra=" (경로: <메일 계정>\\없음)")
    assert "계정@x.com" in str(e) and "계정@x.com" not in e.log_message and "<메일 계정>" in e.log_message
    e = rt.not_found("메일 폴더", "없음", ["ZQX가"], extra=" (경로: 계정@x.com\\없음)", log_extra=" (경로: <계정>)")
    assert "계정@x.com" not in e.log_message and "ZQX" not in e.log_message


# L2: relay는 --slide를 클립보드·임시 워드보다 먼저 확인 -------------------------------------------------------------

class _Shapes:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Paste(self):
        pic = type("Pic", (), {"Width": 400.0, "Height": 300.0, "Left": 0.0, "Top": 0.0, "LockAspectRatio": 0})()
        self.items.append(pic)
        return type("SR", (), {"Count": 1, "Item": lambda s, i: pic})()


class _Slides:
    def __init__(self, n):
        self.slides = [type("Slide", (), {"Shapes": _Shapes()})() for _ in range(n)]

    @property
    def Count(self):
        return len(self.slides)

    def __call__(self, i):
        return self.slides[i - 1]

    def Add(self, i, layout):
        s = type("Slide", (), {"Shapes": _Shapes()})()
        self.slides.insert(i - 1, s)
        return s


def _prs(n=2):
    return type("P", (), {"Slides": _Slides(n), "PageSetup": type("PS", (), {"SlideWidth": 960, "SlideHeight": 540})()})()


class _T:
    def __init__(self, kind, doc):
        self.kind, self.doc = kind, doc

    def set_focus(self, **k):
        pass

    def resolved(self):
        return {"app": off.APP_LABEL[self.kind]}


class _WordDoc:
    def __init__(self):
        self.InlineShapes = type("IS", (), {"Count": 1, "Item": lambda s, i: type("I", (), {
            "Range": type("R", (), {"CopyAsPicture": lambda s: None})()})()})()

    def Range(self, a, b):
        doc = self
        return type("R", (), {"InsertAfter": lambda s, t: None, "PasteSpecial": lambda s, **k: None,
                              "InlineShapes": doc.InlineShapes})()

    def Close(self, **k):
        pass


@pytest.fixture
def relay_env(chart_env, monkeypatch):
    """relay를 진짜로 부르는 상태. 클립보드 복사·임시 워드·문서 열기를 기록."""
    log = {"copied": [], "dispatch": [], "opened": []}
    chart = type("C", (), {"Name": "c", "CopyPicture": lambda self, *a: log["copied"].append(a)})()
    import excel
    import word
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(word, "_insert_position", lambda doc, anchor, pos, occurrence=None: (0, "end"))
    monkeypatch.setattr(office, "relay_chart", _real_relay)
    monkeypatch.setattr(office.time, "sleep", lambda s: None)
    prs, wdoc = _prs(2), _WordDoc()
    log["prs"] = prs
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": "x", "path": "", "app": None,
                                                         "doc": prs if kind == "ppt" else wdoc})
    tapp = type("A", (), {"Documents": type("Ds", (), {"Add": lambda s, **k: _WordDoc()})(), "Quit": lambda s, **k: None})()
    monkeypatch.setattr(off.win32com.client, "DispatchEx", lambda n: log["dispatch"].append(n) or tapp, raising=False)

    @contextlib.contextmanager
    def ot(kind, target, **k):
        if kind != "excel":                               # 보내는 엑셀 문서는 빼고 받는 문서만
            log["opened"].append((kind, target))
        yield _T(kind, prs if kind == "ppt" else wdoc)
    monkeypatch.setattr(off, "open_target", ot)
    return log


@pytest.mark.parametrize("target", [None, "file"])
def test_relay_ppt_bad_slide_before_clipboard(relay_env, capsys, tmp_path, target):
    argv = ["transfer", "--from", "excel", "--chart", "c", "--to", "ppt", "--slide", "9", "--method", "relay"]
    if target == "file":
        p = tmp_path / "deck.pptx"
        p.write_bytes(b"x")
        argv += ["--to-target", str(p)]                   # 열려 있지 않은 파일: 열어서 슬라이드를 먼저 확인
    code, out = run(capsys, *argv)
    assert code == 1 and "슬라이드 9번이 없습니다" in out["message"], out
    assert relay_env["copied"] == [] and relay_env["dispatch"] == []      # 클립보드·임시 워드 그대로


def test_relay_ppt_ok_after_slide_check(relay_env, capsys):
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "c", "--to", "ppt", "--slide", "2", "--method", "relay")
    assert code == 0 and out["results"]["ppt"]["method"] == "relay", out
    assert len(relay_env["copied"]) == 1 and relay_env["dispatch"] == ["Word.Application"]
    assert relay_env["prs"].Slides(2).Shapes.Count == 1 and relay_env["opened"] == [("ppt", "active")]


# L3: --to-target ""는 활성 문서 -----------------------------------------------------------------------------------

@pytest.mark.parametrize("kind", ["word", "ppt"])
def test_relay_empty_target_is_active(relay_env, capsys, kind):
    slide = ["--slide", "1"] if kind == "ppt" else []
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "c", "--to", kind, "--to-target", "", "--method", "relay",
                    *slide)
    assert code == 0 and out["results"][kind]["method"] == "relay", out["message"]
    assert "KeyError" not in json.dumps(out, ensure_ascii=False)
    assert relay_env["opened"] == [(kind, "")] and len(relay_env["copied"]) == 1


def test_relay_unit_empty_target_returns_result(relay_env):
    done = office.relay_chart(type("C", (), {"CopyPicture": lambda s, *a: None})(), "", None, None, None, "t")
    assert list(done) == ["word"]
