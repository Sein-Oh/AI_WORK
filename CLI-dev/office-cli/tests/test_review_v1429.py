"""1.4.29 (20차 점검) 회귀 테스트.

M1: transfer --chart가 받는 문서 오타([추천])에도 다음 방식(redraw, relay)으로 넘어가 클립보드를 덮어쓰고,
    [추천]이 '; relay: ...' 앞, 오류 가운데에 묻혔음. 대상을 찾지 못한 오류는 바로 원래 오류 그대로 멈춥니다.
M2: outlook 폴더를 찾지 못하면 전체 경로끼리 비교해 같은 상위 경로(메일 계정) 때문에 관계없는 폴더가 모두 [후보]였고,
    오류 로그에 메일 계정 이름(메일 주소)이 남았음.
L1: 차트 다시 그리기(redraw)가 제목·계열 이름의 $...$를 수식으로 읽어 깨지거나 ValueError. transfer의 오류 로그에
    내부 오류 글이 그대로 남았음.
L2: reader search/index --file <파일 하나>가 파서 내부 오류를 STOP_RETRY 입력 오류로 바꿨음.
L3: run_win_tests.py가 pywin32 확인을 인자 해석보다 먼저 해 --agent -h가 exit 1이었음.
"""
import importlib.util
import json
import os
import sys

import pytest

import _office as off
import _runtime as rt
import office
import outlook
import ppt
import reader
import test_office_cmd

# 문서에서 가져온 글을 나타내는 표지. 오류 로그에 하나도 남으면 안 됨
S = "ZQX"
ROOT_MAIL = "ZQMhong@corp.com"           # 메일 계정 이름 (보통 메일 주소)
TOOLS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools")

chart_env = test_office_cmd.chart_env       # fixture (엑셀 차트 하나, 삽입·그림 방식은 기록용 대역)
run = test_office_cmd.run
_real_insert_word = office.insert_picture_word
_real_relay = office.relay_chart


def _log(d):
    p = os.path.join(d["data"], "logs", "errors.txt")
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


def _main(mod, capsys, monkeypatch, *argv):
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


# M1: 대상 문제는 다음 방식으로 넘어가지 않음 ---------------------------------------------------------------------

@pytest.fixture
def relay_spy(monkeypatch):
    calls = []
    monkeypatch.setattr(office, "relay_chart", lambda *a, **k: calls.append(a) or (_ for _ in ()).throw(RuntimeError("relay")))
    return calls


def test_chart_target_recommend_stops_unchanged(chart_env, capsys, monkeypatch, isolated_dirs, relay_spy):
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "ZQX_보고서.docx", "path": "", "app": None, "doc": None}])
    monkeypatch.setattr(office, "insert_picture_word", _real_insert_word)
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "word",
                    "--to-target", "zqx 보고서.docx")
    assert code == 1 and out["message"].endswith("[추천] ZQX_보고서.docx"), out
    assert "relay" not in out["message"] and "--probe" not in out["message"] and "시도한 방식" not in out["message"]
    assert relay_spy == [] and chart_env["calls"] == ["export", "redraw"]       # 그림 문제(export)만 다음 방식으로
    log = _log(isolated_dirs)
    assert log and S not in log and "[추천 생략]" in log, log


def test_chart_target_bad_extension_stops(chart_env, capsys, monkeypatch, tmp_path, relay_spy):
    monkeypatch.setattr(office, "insert_picture_word", _real_insert_word)
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "word",
                    "--to-target", str(tmp_path / "새 문서.txt"))
    assert code == 1 and "확장자" in out["message"] and "relay" not in out["message"], out
    assert relay_spy == []


def test_relay_resolves_target_before_clipboard(chart_env, capsys, monkeypatch):
    copied = []
    chart = type("C", (), {"Name": "사인 2주기", "CopyPicture": lambda self, *a: copied.append(a)})()
    import excel
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(office, "relay_chart", _real_relay)
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "ZQX_보고서.docx", "path": "", "app": None, "doc": None}])
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "word",
                    "--to-target", "zqx 보고서.docx", "--method", "relay")
    assert code == 1 and out["message"].endswith("[추천] ZQX_보고서.docx"), out
    assert copied == []                                   # 사용자의 클립보드를 덮어쓰지 않음


def test_resolve_target_errors_are_no_fallback(monkeypatch, tmp_path):
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "ZQX가.docx", "path": "", "app": None, "doc": None},
                                                       {"name": "ZQX나.docx", "path": "", "app": None, "doc": None}])
    for target in ("zqx 가.docx", "가나다.docx", "없는.docx", str(tmp_path / "a.txt")):
        with pytest.raises(rt.InputError) as ei:
            off.resolve_target("word", target, readonly=False, create=True, option="--to-target")
        assert ei.value.no_fallback, (target, str(ei.value))


def test_slide_errors_are_no_fallback():
    prs = type("P", (), {"Slides": type("S", (), {"Count": 2})()})()
    for value in ("9", "x"):
        with pytest.raises(rt.InputError) as ei:
            ppt._resolve_slide(prs, value, default_new=False)
        assert ei.value.no_fallback


# M2: outlook 폴더 [후보]는 끝 이름으로 비교, 메일 계정 이름은 로그에 남기지 않음 -----------------------------------

class _F:
    def __init__(self, name, subs=()):
        self.Name = name
        self.Folders = list(subs)
        self.Parent = None
        for f in self.Folders:
            f.Parent = self


class _NS:
    def __init__(self):
        self.root = _F(ROOT_MAIL, [_F("ZQX받은편지함"), _F("ZQX프로젝트A", [_F("ZQX회의록")]), _F("ZQX보관")])
        self.Folders = [self.root]

    def GetDefaultFolder(self, n):
        return self.root.Folders[0]

    def Logon(self, *a):
        pass


@pytest.fixture
def mail(monkeypatch):
    app = type("A", (), {"GetNamespace": lambda self, k: _NS()})()
    monkeypatch.setattr(outlook, "get_outlook", lambda: app)


def test_folder_candidates_match_leaf_names(mail, capsys, monkeypatch, isolated_dirs):
    code, out = _main(outlook, capsys, monkeypatch, "list", "--folder", "프로젝트")
    assert code == 1 and out["candidates"] == [f"{ROOT_MAIL}\\ZQX프로젝트A"], out
    code, out = _main(outlook, capsys, monkeypatch, "list", "--folder", "회의")    # 비슷한 끝 이름 없음: [후보] 아님
    assert code == 1 and "candidates" not in out and "사용 가능한 메일 폴더" in out["message"], out
    assert f"{ROOT_MAIL}\\ZQX보관" in out["message"]
    code, out = _main(outlook, capsys, monkeypatch, "list", "--folder", "zqx프로젝트a\\회의")
    assert code == 1 and out["candidates"] == [f"{ROOT_MAIL}\\ZQX프로젝트A\\ZQX회의록"], out


def test_folder_same_name_recommends_full_path(mail, capsys, monkeypatch):
    code, out = _main(outlook, capsys, monkeypatch, "list", "--folder", "zqx_보관")
    assert code == 1 and out["message"].endswith(f"[추천] {ROOT_MAIL}\\ZQX보관"), out


@pytest.mark.parametrize("folder", ["프로젝트", "회의", "zqx_보관", f"{ROOT_MAIL}\\없음"])
def test_guard_folder_log_has_no_mailbox_root(mail, capsys, monkeypatch, isolated_dirs, folder):
    code, out = _main(outlook, capsys, monkeypatch, "list", "--folder", folder)
    assert code == 1 and ROOT_MAIL in out["message"], out
    log = _log(isolated_dirs)
    assert log and ROOT_MAIL not in log.replace(f"--folder {folder}", "") and S not in log, log
    assert "<메일 계정>" in log


def test_not_found_display_and_log_name():
    e = rt.not_found("메일 폴더", "프로젝트", ["ZQX프로젝트A", "기타"], display=lambda n: "ROOT\\" + n, log_name="<계정>\\프로젝트")
    assert e.candidates == ["ROOT\\ZQX프로젝트A"] and "'ROOT\\프로젝트'" in str(e)
    assert "ROOT" not in e.log_message and S not in e.log_message and "<계정>" in e.log_message


# L1: 차트 다시 그리기의 $, 오류 로그의 내부 오류 글 ---------------------------------------------------------------

def test_render_chart_does_not_parse_mathtext(tmp_path, monkeypatch):
    import matplotlib.text as mtext
    seen = []
    real_draw = mtext.Text.draw

    def draw(self, renderer):
        if "$" in self.get_text():
            seen.append(self.get_parse_math())
        return real_draw(self, renderer)
    monkeypatch.setattr(mtext.Text, "draw", draw)
    bad = "단가 $\\frac 와 $"
    for kind in (51, 4, 5):
        data = {"type": kind, "title": bad, "width": 300, "height": 180,
                "series": [{"name": bad, "values": [1, 2], "x": [bad, "$5 ~ $10"]},
                           {"name": "$x^2$", "values": [2, 3], "x": [bad, "b"]}]}
        path = office.render_chart(data, str(tmp_path / f"{kind}.png"))
        assert os.path.getsize(path) > 0
    assert seen and not any(seen), seen                   # 제목, 계열 이름(범례), 눈금 글 모두 글자 그대로
    import matplotlib
    assert matplotlib.rcParams["text.parse_math"] is True   # 전역 설정은 그대로


def test_redraw_with_dollar_title_succeeds(chart_env, capsys, monkeypatch, relay_spy):
    monkeypatch.setattr(office, "image_by_redraw", lambda obj, stem: office.render_chart(
        {"type": 51, "title": "단가 $\\frac 와 $", "width": 300, "height": 180,
         "series": [{"name": "s", "values": [1, 2], "x": ["a", "b"]}]},
        rt.resolve_output_path(None, stem + "_redraw", ".png")))
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "ppt")
    assert code == 0 and out["results"]["ppt"]["method"] == "redraw" and relay_spy == [], out


def test_guard_chart_internal_error_text_not_logged(chart_env, capsys, monkeypatch, isolated_dirs):
    def boom(kind):
        def f(*a, **k):
            raise kind("ZQX 매출 차트 제목")
        return f
    monkeypatch.setattr(office, "image_by_export", boom(RuntimeError))
    monkeypatch.setattr(office, "image_by_redraw", boom(ValueError))
    monkeypatch.setattr(office, "relay_chart", boom(KeyError))
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "word")
    assert code == 1 and S in out["message"], out
    log = _log(isolated_dirs)
    assert log and S not in log, log
    assert "export: RuntimeError" in log and "redraw: ValueError" in log and "relay: KeyError" in log, log


# L2: reader --file 파일 하나의 내부 오류는 내부 오류 그대로 -------------------------------------------------------

@pytest.fixture
def one_file(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("hello", encoding="utf-8")
    return p


@pytest.mark.parametrize("argv", [["search", "-q", "hello", "--file"], ["index", "--file"]])
def test_single_file_internal_error_is_internal(one_file, capsys, monkeypatch, isolated_dirs, argv):
    def bug(path, check_index=True):
        raise KeyError("heading")
    monkeypatch.setattr(reader, "load_document_status", bug)
    code, out = _main(reader, capsys, monkeypatch, *(argv + [str(one_file)]))
    assert code == 1 and out["error_type"] == "internal" and out["agent_directive"].startswith("ERROR:"), out
    assert out["message"] == "KeyError: 'heading'"
    assert "traceback:" in _log(isolated_dirs)


@pytest.mark.parametrize("argv", [["search", "-q", "hello", "--file"], ["index", "--file"]])
def test_single_file_input_error_keeps_log_message(one_file, capsys, monkeypatch, isolated_dirs, argv):
    def bad(path, check_index=True):
        raise rt.InputError("문서를 읽을 수 없습니다: ZQX 본문", log_message="문서를 읽을 수 없습니다 (본문 생략)")
    monkeypatch.setattr(reader, "load_document_status", bad)
    code, out = _main(reader, capsys, monkeypatch, *(argv + [str(one_file)]))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and S in out["message"], out
    log = _log(isolated_dirs)
    assert S not in log and "본문 생략" in log, log


def test_folder_search_failure_reason_names_error_type(one_file, capsys, monkeypatch):
    (one_file.parent / "b.txt").write_text("hello", encoding="utf-8")
    real = reader.load_document_status

    def bug(path, check_index=True):
        if path.endswith("a.txt"):
            raise KeyError("heading")
        return real(path, check_index=check_index)
    monkeypatch.setattr(reader, "load_document_status", bug)
    code, out = _main(reader, capsys, monkeypatch, "search", "-q", "hello", "--file", str(one_file.parent))
    assert code == 0, out
    assert [s["reason"] for s in out["skipped"]] == ["KeyError: 'heading'"], out


# L3: run_win_tests.py는 인자를 먼저 해석 ------------------------------------------------------------------------

def _runner_without_pywin32(monkeypatch, argv):
    monkeypatch.setitem(sys.modules, "pythoncom", None)          # import pythoncom -> ImportError
    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    monkeypatch.setattr(sys, "argv", ["run_win_tests.py"] + argv)
    if TOOLS not in sys.path:
        monkeypatch.syspath_prepend(TOOLS)
    spec = importlib.util.spec_from_file_location("run_win_tests_l3", os.path.join(TOOLS, "run_win_tests.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.mark.parametrize("argv,code,status", [(["--agent", "-h"], 0, "ok"), (["--agent", "--nope"], 2, "error"),
                                              (["--agent", "--only", "Z"], 2, "error"), (["--agent"], 1, "error")])
def test_runner_parses_args_before_pywin32_check(monkeypatch, capsys, argv, code, status):
    mod = _runner_without_pywin32(monkeypatch, argv)
    with pytest.raises(SystemExit) as ei:
        mod.main()
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert ei.value.code == code and len(out) == 1, out
    payload = json.loads(out[0])
    assert payload["status"] == status
    if code == 1:
        assert payload["error_type"] == "dependency" and "python -m pip install pywin32" in payload["agent_directive"]
