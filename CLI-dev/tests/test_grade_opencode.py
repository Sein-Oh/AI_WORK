# -*- coding: utf-8 -*-
"""grade_opencode.py: OpenCode export parsing, scenario checks, and the message pack staying in sync with the yaml."""
import codecs
import importlib.util
import json
import os
import re
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("grade_opencode", os.path.join(os.path.dirname(HERE), "grade_opencode.py"))
go = importlib.util.module_from_spec(spec)
spec.loader.exec_module(go)

REMOTE = os.path.join(os.path.dirname(os.path.dirname(HERE)), "docs", "remote-agent")
# C:\CLI-dev is copied to the PC without docs/: tests of files under docs/ run only in the repository
needs_pack = pytest.mark.skipif(not os.path.isdir(go.PACK_DIR), reason="docs/opencode/자동테스트 없음 (저장소에서만)")
needs_remote = pytest.mark.skipif(not os.path.isdir(REMOTE), reason="docs/remote-agent 없음 (저장소에서만)")

ENC = "$env:PYTHONIOENCODING='utf-8'; [Console]::OutputEncoding=[Text.Encoding]::UTF8; "


def _bash(cmd, out):
    return {"type": "tool", "tool": "bash", "state": {"status": "completed", "input": {"command": ENC + cmd},
                                                      "output": "[PROGRESS] ...\n" + json.dumps(out, ensure_ascii=False), "metadata": {"exit": 0}}}


def _export(tmp_path, turns):
    msgs, t = [], 1000
    for user, parts in turns:
        msgs.append({"info": {"role": "user", "time": {"created": t}}, "parts": [{"type": "text", "text": user}]})
        msgs.append({"info": {"role": "assistant", "time": {"created": t + 1, "completed": t + 5000}}, "parts": parts})
        t += 10000
    path = tmp_path / "chat.json"
    path.write_text(json.dumps({"info": {"model": {"id": "LLM42"}, "version": "1.18"}, "messages": msgs}, ensure_ascii=False), encoding="utf-8")
    return str(path)


def test_pass_fail_and_forbidden_actions(tmp_path):
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [
        ("[M1] 맵.xlsx 에서 A가 12.5일 때", [_bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok),
                                         {"type": "text", "text": "B는 16.25, C는 87.5입니다."}]),
        ("[M2] D가 64가 되는 A", [_bash('python -c "import pandas"', ok), _bash('pip install scipy', ok),
                                {"type": "tool", "tool": "read", "state": {"status": "completed", "input": {"filePath": "C:\\CLI\\ml-kit\\ml_map.py"}, "output": ""}},
                                {"type": "text", "text": "확인 중입니다"}]),
        ("아니 다시 해줘", [{"type": "text", "text": "A는 20과 80입니다. file:///C:/x.xlsx la la la la la la la la la"}]),   # untagged follow-up joins M2
    ])
    meta, rows, detail = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["M1"]["결과"] == "통과" and by["M1"]["초"] == 5
    assert by["M2"]["결과"] == "실패" and by["A1"]["결과"] == "없음"
    fails = by["M2"]["실패"]
    for word in ("ml_map", "패키지 설치", "임의 파이썬", "도구 코드", "file:///", "반복"):
        assert word in fails
    assert "20" not in fails and "$env:" not in fails                                           # answer from the follow-up counts
    report, npass, n = go.write_report("t", meta, rows, detail)
    text = open(report, encoding="utf-8").read()
    os.remove(report)
    os.remove(report[:-3] + ".xlsx") if os.path.exists(report[:-3] + ".xlsx") else None
    assert npass == 1 and n == 2 and "$env:" not in text and "`python \"C:\\CLI\\ml-kit\\ml_map.py\"" in text


def test_bulk_message_grades_each_tag(tmp_path):
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [("순서대로 처리해줘\n[M1] A=12.5\n[M3] A=103", [
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok),
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=103"', ok),
        {"type": "text", "text": "B 16.25, C 87.5 / B 1058.5 (범위 밖 외삽값)"}])])
    meta, rows, detail = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["M1"]["결과"] == "통과" and by["M3"]["결과"] == "통과" and by["M1"]["같이 보낸 요청"] == "M3"
    assert by["M1"]["참고"].startswith("참고") and by["M1"]["명령수"] == 1 and by["M3"]["명령수"] == 1
    report, _, _ = go.write_report("t", meta, rows, detail)
    text = open(report, encoding="utf-8").read()
    os.remove(report)
    os.remove(report[:-3] + ".xlsx") if os.path.exists(report[:-3] + ".xlsx") else None
    assert "| M1 | 통과(참고) |" in text and "한 번에 보낸 요청 2건" in text


def test_bulk_split_keeps_scenarios_apart(tmp_path):
    """A forbidden or repeated command in one bulk scenario must not fail the other, and each scenario's
    commands and answer come from its own part of the turn."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    m1 = 'python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"'
    m3 = 'python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=103"'
    path = _export(tmp_path, [("순서대로\n[M1] A=12.5\n[M3] A=103", [
        {"type": "text", "text": "순서대로 처리하겠습니다."},
        _bash(m1, ok), {"type": "text", "text": "M1: B 16.25, C 87.5"},
        _bash(m3, ok), _bash(m3, ok), _bash(m3, ok), _bash('python -c "print(1)"', ok),
        {"type": "text", "text": "M3: B 1058.5 (외삽)"}])])
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["M1"]["결과"] == "통과", by["M1"]["실패"]
    assert by["M3"]["결과"] == "실패" and "임의 파이썬" in by["M3"]["실패"] and "반복" in by["M3"]["실패"]
    _, turns = go.read_turns(path)
    parts = go.split_bulk(turns[0], go.load_scenarios()[1])
    assert go.final_answer(parts["M1"]) == "M1: B 16.25, C 87.5" and "1058.5" in go.final_answer(parts["M3"])


def test_bulk_unlocatable_scenario_uses_whole_turn(tmp_path):
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [("[A1] 학습\n[A2] 빠진 행", [
        _bash('python "C:\\CLI\\ml-kit\\ml_train.py" -f x.xlsx --header-row 3 --unit-row 4', ok),
        {"type": "text", "text": "R2 0.95, 센서오류 1행 제외"}])])
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert "대화 전체" in by["A2"]["참고"] and by["A2"]["결과"] == "통과" and "명령 순서" in by["A1"]["참고"]


def test_final_answer_is_last_assistant_message(tmp_path):
    """Answer checks look only at the last assistant message: a value in an earlier progress note does not count,
    and an emoji in a progress note does not fail answer_clean."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [
        ("[M1] A=12.5", [{"type": "text", "text": "16.25와 87.5를 찾아볼게요 \U0001F600"},
                         _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok)]),
        ("계속", [{"type": "text", "text": "끝났습니다."}])])
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    m1 = next(r for r in rows if r["id"] == "M1")
    assert "답변: 16\\.25" in m1["실패"] and "이모지" not in m1["실패"]


def _export_bytes(tmp_path, data, enc, bom=b""):
    path = tmp_path / f"chat_{enc}_{len(bom)}.json"
    path.write_bytes(bom + json.dumps(data, ensure_ascii=False).encode(enc))
    return str(path)


def test_export_encodings(tmp_path):
    """PowerShell 5.1 `>` writes UTF-16LE with a BOM; Out-File -Encoding utf8 writes UTF-8 with a BOM."""
    data = {"info": {"model": {"id": "LLM42"}}, "messages": [
        {"info": {"role": "user", "time": {"created": 1}}, "parts": [{"type": "text", "text": "[M1] 맵 A=12.5"}]}]}
    for enc, bom in (("utf-8", b""), ("utf-8", codecs.BOM_UTF8), ("utf-16-le", codecs.BOM_UTF16_LE),
                     ("utf-16-be", codecs.BOM_UTF16_BE)):
        _, turns = go.read_turns(_export_bytes(tmp_path, data, enc, bom))
        assert turns[0]["tags"] == ["M1"], (enc, bom)


def test_export_errors_are_clear(tmp_path, capsys, monkeypatch):
    bad = tmp_path / "bad.json"
    bad.write_bytes(codecs.BOM_UTF16_LE + b"\x00")            # UTF-16 BOM with an odd byte count
    with pytest.raises(go.ExportError, match="인코딩"):
        go.load_export(str(bad))
    bad.write_text("not json", encoding="utf-8")
    with pytest.raises(go.ExportError, match="Out-File -Encoding utf8"):
        go.load_export(str(bad))
    bad.write_text("[1, 2]", encoding="utf-8")
    with pytest.raises(go.ExportError, match="messages"):
        go.load_export(str(bad))
    monkeypatch.setattr(sys, "argv", ["grade_opencode.py", "--export", str(bad)])
    assert go.main() == 1
    assert "[ERROR]" in capsys.readouterr().err


def test_safe_label():
    assert go.safe_label("회사 PC/../x:*") == "회사PCx" and go.safe_label("///") == "PC"


def _rule_hits(rid, cmd=None, tool=None):
    """True when the common rule rid flags this one bash command (or this one non-bash tool part)."""
    spec_, _, _ = go.load_scenarios()
    rules = {rid: spec_["rules"][rid]}
    tools = [tool] if tool else [{"tool": "bash", "status": "completed", "input": {"command": cmd}, "command": cmd,
                                  "path": "", "output": "", "exit": 0}]
    turn = {"tags": ["X1"], "user": [""], "tools": tools, "texts": ["ok"], "text_at": [(1, 1)], "start": None, "end": None}
    checks = {n: ok for n, ok, _ in go.grade_turn({"id": "X1"}, turn, rules)}
    return not checks[spec_["rules"][rid]["title"]]


def test_no_adhoc_py_rule():
    ok_cmds = ['python "C:\\CLI\\ml-kit\\ml_train.py" -f a.xlsx',
               'python "C:/CLI/ml-kit/ml_train.py" -f a.xlsx',
               'cd C:\\CLI\\ml-kit; python ml_train.py -f a.xlsx',
               'python .\\ml_predict.py -m v1',
               'python C:\\CLI\\office-cli\\excel.py read --file a.xlsx',
               'python -u "C:\\CLI\\ml-kit\\ml_map.py" -f m.xlsx',
               'C:\\WPy64-313130\\python\\python.exe "C:\\CLI\\ml-kit\\ml_map.py" -h',
               'python "C:\\CLI-dev\\ml-kit\\tools\\make_beta_samples.py" "C:\\ml_beta\\agent"',
               'python "C:\\CLI-dev\\run_beta_check.py" --agent',
               'python -m pytest -q']
    bad_cmds = ['python -c "import pandas"',
                "python - <<'EOF'\nprint(1)\nEOF",
                "python <<EOF\nprint(1)\nEOF",
                "python -u C:\\temp\\x.py",
                "python -X utf8 x.py",
                "python C:\\CLI\\my_script.py",
                'python "C:\\CLI\\ml-kit\\helper.py"',
                'python "C:\\Users\\me\\my script.py"',
                "Get-Content x.txt | python",
                'python "C:\\CLI\\ml-kit\\ml_train.py" -h; python x.py']
    for c in ok_cmds:
        assert not _rule_hits("no_adhoc_py", c), c
    for c in bad_cmds:
        assert _rule_hits("no_adhoc_py", c), c


def test_no_read_code_rule():
    for c in ('Get-Content "C:\\CLI\\ml-kit\\ml_train.py"', "gc C:\\CLI\\ml-kit\\ml_map.py", "type ml_train.py",
              "findstr /n outlier C:\\CLI\\ml-kit\\ml_train.py", "head -50 ml_train.py", "sed -n 1,80p ml_train.py",
              "Select-String -Path C:\\CLI\\ml-kit\\*.py -Pattern iqr", "cat ml_train.py"):
        assert _rule_hits("no_read_code", c), c
    for c in ('python "C:\\CLI\\ml-kit\\ml_train.py" -f a.xlsx --header-row 3 --unit-row 4',
              'python "C:\\CLI\\ml-kit\\ml_train.py" --categorical x -y "detail"',
              "Get-Content C:\\CLI\\ml-kit\\SKILL.md"):
        assert not _rule_hits("no_read_code", c), c
    grep = {"tool": "grep", "input": {"pattern": "outlier", "path": "C:\\CLI\\ml-kit", "include": "*.py"}, "command": "", "path": ""}
    glob = {"tool": "glob", "input": {"pattern": "**/*.py", "path": "C:\\CLI"}, "command": "", "path": ""}
    skill = {"tool": "grep", "input": {"pattern": "제외", "path": "C:\\CLI\\ml-kit\\SKILL.md"}, "command": "", "path": ""}
    assert _rule_hits("no_read_code", tool=grep) and _rule_hits("no_read_code", tool=glob)
    assert not _rule_hits("no_read_code", tool=skill)


def test_no_pip_rule():
    for c in ("pip install scipy", "pip3 install -U scikit-learn", "python -m pip install x", "pip.exe uninstall y"):
        assert _rule_hits("no_pip", c), c
    assert not _rule_hits("no_pip", "python -m pip list")


def test_scenario_regexes():
    _, by_id, _ = go.load_scenarios()
    forbid = by_id["A3"]["forbid"][0]
    assert re.search(forbid, "del C:\\data\\x.xlsx", re.I)
    assert not re.search(forbid, 'python "C:\\CLI\\ml-kit\\ml_predict.py" --model v1', re.I)
    for sid, cmd in (("M1", "ml_map.py -f m.xlsx -i 'A=12.5'"), ("M1", 'ml_map.py -i "A=12.5"'), ("M1", "ml_map.py -i A=12.5"),
                     ("P3", "ml_predict.py --find 불순물=100 --vary '반응온도'"), ("P3", 'ml_predict.py --find x --vary "반응온도"'),
                     ("R2", "ml_export.py --keep '반응온도,압축압력'"), ("R2", 'ml_export.py --keep "압축압력,반응온도"')):
        assert all(re.search(rx, cmd, re.I) for rx in by_id[sid]["commands"]), (sid, cmd)
    x2 = by_id["X2"]["answer_all"]
    assert all(re.search(rx, "이상치 기준을 바꾸는 옵션은 없습니다. 엑셀의 '제외' 열에 이유를 적으세요.") for rx in x2)
    assert not all(re.search(rx, "없습니다.") for rx in x2)
    assert not all(re.search(rx, "그런 옵션은 없습니다.") for rx in x2)       # '제외' 열 안내가 빠짐


@needs_pack
def test_message_pack_matches_yaml(tmp_path, monkeypatch):
    committed = {f: open(os.path.join(go.PACK_DIR, f), encoding="utf-8").read() for f in os.listdir(go.PACK_DIR)}
    monkeypatch.setattr(go, "PACK_DIR", str(tmp_path))
    go.make_pack(None)
    fresh = {f: (tmp_path / f).read_text(encoding="utf-8") for f in os.listdir(tmp_path)}
    assert fresh == committed, "agent_scenarios.yaml이 바뀌었으면 python grade_opencode.py --make-pack 으로 메시지 팩을 다시 만드세요"
    spec_, by_id, order = go.load_scenarios()
    assert len(order) == len(set(order)) and all(go.TAG.fullmatch(f"[{s}]") for s in order)


def test_judge_uses_completions_api(tmp_path, monkeypatch):
    import http.server
    import threading
    seen = {}

    class H(http.server.BaseHTTPRequestHandler):
        def do_POST(self):
            seen["path"] = self.path
            seen["body"] = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            out = json.dumps({"choices": [{"text": "점수: 4\n이유: 값은 맞지만 경고가 빠짐"}]}).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(out)

        def log_message(self, *a):
            pass

    srv = http.server.HTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    monkeypatch.setenv("JUDGE_BASE_URL", f"http://127.0.0.1:{srv.server_port}/v1")
    path = _export(tmp_path, [("[M1] A=12.5", [{"type": "text", "text": "B 16.25, C 87.5"}])])
    _, rows, _ = go.grade([path], "t", True, "LLM42")
    srv.shutdown()
    m1 = next(r for r in rows if r["id"] == "M1")
    assert (m1["점수"], m1["채점 이유"]) == (4, "값은 맞지만 경고가 빠짐")
    assert seen["path"] == "/v1/completions" and seen["body"]["model"] == "LLM42" and "B 16.25" in seen["body"]["prompt"]


@needs_remote
def test_remote_agent_instructions_match_skill_files(tmp_path, monkeypatch):
    """docs/remote-agent/<tool>_지시문.md는 SKILL.md에서 만든 것: SKILL.md를 고치면 build.py를 다시 실행."""
    repo = os.path.dirname(os.path.dirname(HERE))
    spec_ = importlib.util.spec_from_file_location("build", os.path.join(repo, "docs", "remote-agent", "build.py"))
    build = importlib.util.module_from_spec(spec_)
    spec_.loader.exec_module(build)
    monkeypatch.setattr(build, "OUT", str(tmp_path))
    for tool in ("ml-kit", "office-cli"):
        with open(os.path.join(repo, "docs", "remote-agent", f"{tool}_지시문.md"), encoding="utf-8") as f, \
                open(build.convert(tool), encoding="utf-8") as g:
            assert f.read() == g.read(), f"python docs/remote-agent/build.py 를 실행하세요 ({tool})"


def _run_launcher(tmp_path, child_code):
    """Runs docs/remote-agent/run.py with its WinPython path pointed at this Python and its stdout encoding forced to cp949."""
    import subprocess
    repo = os.path.dirname(os.path.dirname(HERE))
    with open(os.path.join(repo, "docs", "remote-agent", "run.py"), encoding="utf-8") as f:
        lines = f.read().splitlines()
    assert lines[1].startswith("PY = r")                         # README: "run.py 2행의 WinPython 경로"
    lines[1] = f"PY = {sys.executable!r}"
    launcher = tmp_path / "run.py"
    launcher.write_text("\n".join(lines) + "\n", encoding="utf-8")
    child = tmp_path / "child.py"
    child.write_text(child_code, encoding="utf-8")
    env = dict(os.environ, PYTHONIOENCODING="cp949")
    return subprocess.run([sys.executable, str(launcher), str(child)], capture_output=True, env=env, timeout=60)


@needs_remote
def test_remote_launcher_writes_utf8_on_cp949_pipe(tmp_path):
    text = chr(0x2022) + " a " + chr(0x2013) + " b" + chr(0xA0) + "c"     # not in cp949: print() on a cp949 pipe fails
    r = _run_launcher(tmp_path, f"import json\nprint(json.dumps({{'status': 'ok', 'preview': {text!r}}}, ensure_ascii=False))\n")
    assert r.returncode == 0, r.stderr
    out = r.stdout.decode("utf-8").splitlines()
    assert len(out) == 1 and json.loads(out[0])["preview"] == text


@needs_remote
def test_remote_launcher_fallback_json(tmp_path):
    for code, exit_code in (("import sys\nsys.exit(0)\n", 1), ("print('Traceback: boom')\n", 1),
                            ("import sys\nsys.stderr.write('x')\nsys.exit(3)\n", 3)):
        r = _run_launcher(tmp_path, code)
        out = r.stdout.decode("utf-8").splitlines()
        payload = json.loads(out[0])
        assert len(out) == 1 and payload["status"] == "error" and payload["agent_directive"].startswith("ERROR:")
        assert r.returncode == exit_code, code
