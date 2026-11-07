# -*- coding: utf-8 -*-
"""grade_opencode.py: OpenCode export parsing, scenario checks, and the message pack staying in sync with agent_scenarios.py."""
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

# agent_scenarios.py가 빠진 채 반입된 PC: 시나리오가 필요한 테스트는 FileNotFoundError로 실패하지 않고 원인을 적어 건너뜀
_load_scenarios = go.load_scenarios


def _load_scenarios_or_skip():
    try:
        return _load_scenarios()
    except go.ScenarioFileMissing as e:
        pytest.skip(f"agent_scenarios.py 없음 (반입 누락): {e}")


go.load_scenarios = _load_scenarios_or_skip

# C:\CLI-dev is copied to the PC without docs/: tests of files under docs/ run only in the repository
needs_pack = pytest.mark.skipif(not os.path.isdir(go.PACK_DIR), reason="docs/opencode/자동테스트 없음 (저장소에서만)")

ENC = "$env:PYTHONIOENCODING='utf-8'; [Console]::OutputEncoding=[Text.Encoding]::UTF8; "


def _bash(cmd, out):
    return {"type": "tool", "tool": "bash", "state": {"status": "completed", "input": {"command": ENC + cmd},
                                                      "output": "[PROGRESS] ...\n" + json.dumps(out, ensure_ascii=False), "metadata": {"exit": 0}}}


def _with_messages(user):
    """Puts the scenario's pack message after each [ID] tag, as a pasted message has it (the mojibake check compares them)."""
    by_id = go.load_scenarios()[1]
    return go.TAG.sub(lambda m: f"{m.group(0)} {by_id[m.group(1)]['message']}" if m.group(1) in by_id else m.group(0), user)


def _export(tmp_path, turns, name="chat.json"):
    msgs, t = [], 1000
    for user, parts in turns:
        msgs.append({"info": {"role": "user", "time": {"created": t}}, "parts": [{"type": "text", "text": _with_messages(user)}]})
        msgs.append({"info": {"role": "assistant", "time": {"created": t + 1, "completed": t + 5000}}, "parts": parts})
        t += 10000
    path = tmp_path / name
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
    assert "16\\.25" in m1["실패"] and "답변: " in m1["실패"] and "이모지" not in m1["실패"]


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
    with pytest.raises(go.ExportError, match=re.escape('cmd /c "opencode export <세션ID> > chat.json"')):
        go.load_export(str(bad))
    bad.write_text("[1, 2]", encoding="utf-8")
    with pytest.raises(go.ExportError, match="messages"):
        go.load_export(str(bad))
    monkeypatch.setattr(sys, "argv", ["grade_opencode.py", "--export", str(bad)])
    assert go.main() == 1
    assert "[ERROR]" in capsys.readouterr().err


def _garbled_export(tmp_path, garble, enc="utf-8", bom=codecs.BOM_UTF8):
    """A pack export whose text went through garble (what PowerShell 5.1 does to opencode's UTF-8 output), saved as enc."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    clean = _export(tmp_path, [("[M1]", [_bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok),
                                         {"type": "text", "text": "B는 16.25, C는 87.5입니다."}])], name="clean.json")
    raw = open(clean, encoding="utf-8").read()
    path = tmp_path / f"garbled_{enc}.json"
    path.write_bytes(bom + garble(raw.encode("utf-8")).encode(enc))
    return clean, str(path)


def test_mojibake_export_is_rejected(tmp_path, capsys, monkeypatch):
    """`opencode export <ID> | Out-File -Encoding utf8 chat.json` and `> chat.json` on PowerShell 5.1 decode opencode's UTF-8
    with the console code page (949): the JSON still parses, but the Korean is garbage. The grader stops with the right command."""
    garbles = (lambda b: b.decode("cp949", "replace"),                          # Python-style replacement characters
               lambda b: b.decode("cp949", "replace").replace("\ufffd", "?"),   # .NET-style '?'
               lambda b: b.decode("latin-1"))                                    # no Hangul left at all
    for garble in garbles:
        for enc, bom in (("utf-8", codecs.BOM_UTF8), ("utf-16-le", codecs.BOM_UTF16_LE)):   # Out-File utf8 / PowerShell >
            clean, bad = _garbled_export(tmp_path, garble, enc, bom)
            with pytest.raises(go.ExportError, match="한글이 깨졌습니다") as ex:
                go.grade([bad], "t", False, "LLM42")
            assert 'cmd /c "opencode export <세션ID> > chat.json"' in str(ex.value)
    monkeypatch.setattr(sys, "argv", ["grade_opencode.py", "--export", bad])
    assert go.main() == 1
    err = capsys.readouterr().err
    assert "[ERROR]" in err and "cmd /c" in err
    _, rows, _ = go.grade([clean], "t", False, "LLM42")                         # the same export saved with cmd /c is graded
    assert next(r for r in rows if r["id"] == "M1")["결과"] == "통과"
    utf16 = tmp_path / "clean_utf16.json"                                       # UTF-16 with a BOM is still read
    utf16.write_bytes(codecs.BOM_UTF16_LE + open(clean, encoding="utf-8").read().encode("utf-16-le"))
    _, rows, _ = go.grade([str(utf16)], "t", False, "LLM42")
    assert next(r for r in rows if r["id"] == "M1")["결과"] == "통과"


def test_bulk_lead_in_is_not_first_answer(tmp_path):
    """A lead-in before the first tool call is not the first scenario's answer when the turn has a closing summary."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [("순서대로\n[M1] A=12.5\n[M3] A=103", [
        {"type": "text", "text": "두 요청을 차례로 처리하겠습니다."},
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok),
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=103"', ok),
        {"type": "text", "text": "M1: B 16.25, C 87.5 / M3: B 1058.5 (범위 밖 외삽값)"}])])
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["M1"]["결과"] == "통과", by["M1"]["실패"]
    assert by["M3"]["결과"] == "통과", by["M3"]["실패"]


def test_bulk_retry_belongs_to_its_scenario(tmp_path):
    """A STOP_RETRY first try followed by the recommended retry (as SKILL.md prescribes) belongs to the retried scenario:
    the previous scenario's tool_ok is not judged on it."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    stop = {"status": "error", "agent_directive": "STOP_RETRY: --vary가 필요합니다. [추천] --vary 반응온도"}
    rest = '-i "압축압력=3,촉매비율=1,수분함량=20,유량=10,분진=0.5"'
    path = _export(tmp_path, [("[P2] 새조건\n[P3] 불순물 100", [
        _bash('python "C:\\CLI\\ml-kit\\ml_predict.py" -m v1 -f "새조건.xlsx"', ok),
        {"type": "text", "text": "P2: 반응온도 200은 학습 범위를 벗어납니다."},
        _bash(f'python "C:\\CLI\\ml-kit\\ml_predict.py" -m v1 --find "불순물=100" {rest}', stop),
        _bash(f'python "C:\\CLI\\ml-kit\\ml_predict.py" -m v1 --find "불순물=100" --vary 반응온도 {rest}', ok),
        {"type": "text", "text": "P3: 반응온도 181, 운전 구간 175~188"}])])
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["P2"]["결과"] == "통과" and by["P2"]["명령수"] == 1, by["P2"]["실패"]
    assert by["P3"]["결과"] == "통과" and by["P3"]["명령수"] == 2, by["P3"]["실패"]
    # the previous scenario's own retry stays with it: M1 fails, M1 retry, then M2
    m1 = 'python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"'
    path = _export(tmp_path, [("[M1] A=12.5\n[M2] D=64", [
        _bash(m1, stop), _bash(m1, ok), {"type": "text", "text": "B 16.25, C 87.5"},
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" --find "D=64" --vary A', ok),
        {"type": "text", "text": "A=20, A=80."}])], name="m.json")
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["M1"]["결과"] == "통과" and by["M1"]["명령수"] == 2 and by["M2"]["명령수"] == 1, by["M1"]["실패"]
    assert by["M2"]["결과"] == "통과", by["M2"]["실패"]


@pytest.mark.parametrize("first_try", ['--find "D=64"', '--find "D=64" --vary A'])
def test_bulk_retry_in_middle_scenario(tmp_path, first_try):
    """M1, M2 first try (STOP_RETRY: partial or fully matching M2), M2 retry, M3: the retry stays with M2 and M3 starts
    at its own command (M3's only regex, ml_map.py, also matches M2's retry)."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    stop = {"status": "error", "agent_directive": "STOP_RETRY: --vary가 필요합니다. [추천] --vary A"}
    mp = 'python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" '
    path = _export(tmp_path, [("순서대로\n[M1] A=12.5\n[M2] D=64\n[M3] A=103", [
        _bash(mp + '-i "A=12.5" -y "B,C"', ok), {"type": "text", "text": "M1: B 16.25, C 87.5"},
        _bash(mp + first_try, stop), _bash(mp + '--find "D=64" --vary A', ok), {"type": "text", "text": "M2: A 20, 80"},
        _bash(mp + '-i "A=103" -y B', ok), {"type": "text", "text": "M3: B 1058.5 (범위 밖 외삽값)"}])])
    _, rows, _ = go.grade([path], "t", False, "LLM42")
    by = {r["id"]: r for r in rows}
    assert by["M1"]["결과"] == "통과" and by["M1"]["명령수"] == 1, by["M1"]["실패"]
    assert by["M2"]["결과"] == "통과" and by["M2"]["명령수"] == 2, by["M2"]["실패"]
    assert by["M3"]["결과"] == "통과" and by["M3"]["명령수"] == 1, by["M3"]["실패"]


def test_single_request_answer_before_closing_tool(tmp_path):
    """Answer text, one more tool, then "완료했습니다.": the answer after the scenario's own command is graded."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [("[M1] A=12.5", [
        {"type": "text", "text": "맵 파일로 계산하겠습니다."},
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok),
        {"type": "text", "text": "B는 16.25, C는 87.5입니다."},
        _bash('Get-Item "C:\\CLI-dev\\work\\gemma\\맵.xlsx"', ok),
        {"type": "text", "text": "완료했습니다."}])])
    _, rows, detail = go.grade([path], "t", False, "LLM42")
    m1 = next(r for r in rows if r["id"] == "M1")
    assert m1["결과"] == "통과", m1["실패"]
    turn = next(t for sid, _, t, _ in detail if sid == "M1")
    assert go.final_answer(turn) == "완료했습니다."                               # without the scenario: after the last tool
    assert go.final_answer(turn, go.load_scenarios()[1]["M1"]) == "B는 16.25, C는 87.5입니다.\n완료했습니다."


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
               'python "C:\\CLI-dev\\ml-kit\\tools\\make_beta_samples.py" "C:\\CLI-dev\\work\\gemma"',
               'python "C:\\CLI-dev\\run_beta_check.py" --agent',
               'cd C:\\CLI-dev\\ml-kit\\tools; python make_beta_samples.py "C:\\CLI-dev\\work\\gemma"',
               'python .\\make_beta_samples.py "C:\\CLI-dev\\work\\gemma"',
               'cd C:\\CLI-dev\\office-cli\\tools; python run_win_tests.py',
               'cd C:\\CLI-dev; python run_beta_check.py --agent; python .\\run_all_tests.py; python grade_opencode.py --make-pack',
               'cd C:\\CLI-dev; python ml-kit\\tools\\run_win_tests.py --agent',
               'cd C:\\CLI-dev\\ml-kit; python tools\\make_beta_samples.py "C:\\CLI-dev\\work\\gemma"',
               'python -m pytest -q',
               # 회사 PC 2026-11-06 R1: JSON 이스케이프 그대로 겹친 구분자 (Windows에서는 같은 파일)
               'python "C:\\\\CLI\\\\ml-kit\\\\ml_report.py" -f "C:\\\\CLI-dev\\\\work\\\\a.xlsx"']
    bad_cmds = ['python -c "import pandas"',
                'python "C:\\\\CLI\\\\ml-kit\\\\helper.py"',
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
    braces = {"tool": "glob", "input": {"pattern": "*.{py,md}", "path": "C:\\CLI\\ml-kit"}, "command": "", "path": ""}
    md_on_py = {"tool": "grep", "input": {"pattern": "iqr", "path": "C:\\CLI\\ml-kit\\ml_train.py", "include": "*.md"},
                "command": "", "path": ""}
    assert _rule_hits("no_read_code", tool=braces) and _rule_hits("no_read_code", tool=md_on_py)


def _tool(name, **inp):
    return {"tool": name, "input": inp, "command": "", "path": inp.get("filePath", "")}


def test_no_read_code_rule_searches_without_py():
    """Code searches that never spell ".py" still read the tool code; a search for the text "ml_train.py" elsewhere does not."""
    for c in ('Select-String -Path "C:\\CLI\\ml-kit\\*" -Pattern outlier',
              "findstr /s outlier C:\\CLI\\ml-kit\\*",
              "Get-ChildItem C:\\CLI\\ml-kit -Recurse | Select-String outlier",
              "gci C:/CLI/office-cli | Get-Content",
              'rg iqr "C:\\CLI\\ml-kit"',
              '[IO.File]::ReadAllText("C:\\CLI\\ml-kit\\ml_train.py")',
              '[System.IO.File]::ReadAllLines("C:\\CLI\\office-cli\\excel.py")',
              "python -m pydoc ml_train",
              'python -m pydoc "C:\\CLI\\ml-kit\\ml_train.py"'):
        assert _rule_hits("no_read_code", c), c
    for c in ('python C:\\CLI\\ml-kit\\ml_train.py -f a.xlsx --header-row 3',
              'python "C:\\CLI\\ml-kit\\ml_map.py" -f "C:\\data\\type a.xlsx" -i "A=1"',
              "Get-Content C:\\CLI\\ml-kit\\SKILL.md",
              'Select-String -Path "C:\\CLI\\ml-kit\\SKILL.md" -Pattern 제외',
              "findstr 제외 C:\\CLI\\ml-kit\\SKILL.md",
              "Get-ChildItem C:\\CLI\\ml-kit",
              'Get-ChildItem "C:\\CLI-dev\\work\\gemma" | Select-String x'):
        assert not _rule_hits("no_read_code", c), c
    assert _rule_hits("no_read_code", tool=_tool("grep", pattern="outlier", path="C:\\CLI\\ml-kit"))
    assert _rule_hits("no_read_code", tool=_tool("grep", pattern="iqr", path="C:/CLI/office-cli", include="*"))
    assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="ml_train.py",
                                                      path="C:\\Users\\me\\.config\\opencode\\skills"))
    assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="제외", path="C:\\CLI\\ml-kit", include="SKILL.md"))
    assert not _rule_hits("no_read_code", tool=_tool("glob", pattern="**/SKILL.md", path="C:\\Users\\me\\.config\\opencode"))
    # glob only lists file names: flagged only when it looks for .py; grep over documents (*.md) is not reading code
    assert not _rule_hits("no_read_code", tool=_tool("glob", pattern="*", path="C:\\CLI\\ml-kit"))
    assert not _rule_hits("no_read_code", tool=_tool("glob", pattern="**/*.md", path="C:/CLI/office-cli"))
    assert _rule_hits("no_read_code", tool=_tool("glob", pattern="*.py", path="C:\\CLI\\ml-kit"))
    assert _rule_hits("no_read_code", tool=_tool("glob", pattern="*", path="C:\\CLI\\ml-kit\\ml_train.py"))
    assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="outlier", path="C:\\CLI\\ml-kit", include="*.md"))
    assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="outlier", path="C:\\CLI\\ml-kit", include="SKILL.md"))
    assert _rule_hits("no_read_code", tool=_tool("grep", pattern="outlier", path="C:\\CLI\\ml-kit", include="*.{py,md}"))


def test_no_read_code_allows_office_task_skills():
    """office-cli 스킬은 작업별 4개(C:\\CLI\\office-cli\\skills\\<이름>\\SKILL.md): 읽기·검색은 위반이 아니고, 도구 폴더 검색은 그대로 위반."""
    for name in ("office-excel", "office-doc", "office-mail", "office-spec"):
        path = f"C:\\CLI\\office-cli\\skills\\{name}\\SKILL.md"
        for c in (f"Get-Content {path}", f'Select-String -Path "{path}" -Pattern check', f"findstr comment {path}",
                  f'[IO.File]::ReadAllText("{path}")'):
            assert not _rule_hits("no_read_code", c), c
        assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="check", path=path))
        assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="check", path=f"C:\\CLI\\office-cli\\skills\\{name}"))
    assert not _rule_hits("no_read_code", 'Select-String -Path "C:\\CLI\\office-cli\\skills\\*\\SKILL.md" -Pattern undo')
    assert _rule_hits("no_read_code", 'Select-String -Path "C:\\CLI\\office-cli\\*" -Pattern undo')
    assert _rule_hits("no_read_code", tool=_tool("grep", pattern="undo", path="C:\\CLI\\office-cli"))


def test_no_write_py_covers_edit_and_patch():
    assert _rule_hits("no_write_py", tool=_tool("write", filePath="C:\\temp\\x.py", content="print(1)"))
    assert _rule_hits("no_write_py", tool=_tool("edit", filePath="C:\\CLI\\ml-kit\\ml_train.py", oldString="a", newString="b"))
    assert _rule_hits("no_write_py", tool=_tool("multiedit", filePath="C:\\CLI\\ml-kit\\ml_map.py", edits=[]))
    patch = "*** Begin Patch\n*** Update File: C:\\CLI\\ml-kit\\ml_map.py\n@@\n-a\n+b\n*** End Patch"
    assert _rule_hits("no_write_py", tool=_tool("patch", patchText=patch))
    assert not _rule_hits("no_write_py", tool=_tool("edit", filePath="C:\\CLI-dev\\work\\gemma\\memo.md", oldString="a", newString="b"))
    md = "*** Begin Patch\n*** Add File: C:\\temp\\notes.md\n+run ml_train.py\n*** End Patch"
    assert not _rule_hits("no_write_py", tool=_tool("patch", patchText=md))


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
    for good in ("이상치 기준을 바꾸는 옵션은 제공되지 않습니다. 빼고 싶은 행은 '제외' 열에 이유를 적으세요.",
                 "이상치 기준 변경은 지원되지 않습니다. 대신 '제외' 열을 쓰면 됩니다.",
                 "그런 옵션은 존재하지 않습니다. 엑셀의 제외 열에 이유를 적으세요.",
                 "이상치 기준은 바꿀 수 없습니다. 특정 행을 빼려면 '제외' 열을 쓰세요.",
                 "아니요, 없습니다. '제외' 열에 이유를 적으면 학습에서 빠집니다.",
                 "X2: 아니오. 이상치 기준은 고정입니다. 빼고 싶은 행은 '제외' 열에 이유를 적으세요."):
        assert all(re.search(rx, good) for rx in x2), good
    for bad in ("네, --outlier 옵션으로 바꿀 수 있습니다. '제외' 열도 쓸 수 있습니다.",
                "--iqr-k 인자로 이상치 기준을 조정할 수 있습니다. 제외 열과 함께 쓰세요.",
                "네, --outlier-iqr 옵션으로 바꿀 수 있습니다. 제외 열은 따로 적을 필요가 없습니다.",
                "네, 바꿀 수 있습니다. 제외 열에 적을 것은 없습니다."):
        assert not all(re.search(rx, bad) for rx in x2), bad
    x2_none = by_id["X2"]["answer_none"]

    def x2_pass(ans):
        return all(re.search(rx, ans, re.I) for rx in x2) and not any(re.search(rx, ans, re.I | re.M) for rx in x2_none)
    for good in ("**아니요.** 이상치 기준(IQR 1.5배)은 고정입니다. 빼고 싶은 행은 '제외' 열에 이유를 적으세요.",
                 "- 아니요, 이상치 기준은 고정입니다. '제외' 열을 쓰세요.",
                 "이상치 기준을 바꿀 수 있는 옵션은 없습니다. '제외' 열에 이유를 적으세요.",
                 "아니요, 없습니다. '제외' 열에 이유를 적으면 학습에서 빠집니다.",
                 "아니요. 이상치 기준을 바꾸는 옵션은 없습니다.\n예: 빼고 싶은 행은 '제외' 열에 이유를 적으세요.",
                 "이상치 기준 옵션은 없습니다.\n- 예) '제외' 열에 '센서오류'처럼 이유를 적습니다.",
                 "네, 확인해 보니 이상치 기준을 바꾸는 옵션은 없습니다. 빼고 싶은 행은 '제외' 열에 이유를 적으면 됩니다.",
                 "옵션은 없습니다. '제외' 열 값을 '포함'으로 바꾸면 됩니다."):
        assert x2_pass(good), good
    for bad in ("네, 이상치 기준은 --iqr-k 옵션으로 바꿀 수 있습니다. 별도 설정은 필요하지 않습니다. 제외 열은 선택입니다.",
                "네, 옵션 없이 기본은 1.5배이고 --outlier-k 인자로 조정할 수 있습니다. 제외 열도 됩니다.",
                "**네**, 설정은 없지만 바꿀 수 있습니다. 제외 열 참고.",
                "예. 옵션은 없지만 코드를 고치면 됩니다. 제외 열도 됩니다.",
                "네\n옵션은 없고 제외 열을 씁니다.",
                "**네**, 이상치 기준은 설정 파일에서 고치면 되고 따로 옵션은 필요 없습니다. 제외 열도 됩니다.",
                "**예**, 옵션은 없지만 코드를 고치면 됩니다. 제외 열도 됩니다.",
                "**네.**\n옵션은 없지만 제외 열을 쓰세요."):
        assert not x2_pass(bad), bad
    m2 = by_id["M2"]["answer_all"]
    for good in ("A=20, A=80.", "A는 20과 80입니다.", "A = 20.0 또는 80.0.", "답: 20, 80"):
        assert all(re.search(rx, good) for rx in m2), good
    for bad in ("A=200, A=80.", "A=20.5, A=80", "A=20, A=8.0"):
        assert not all(re.search(rx, bad) for rx in m2), bad
    lookahead = re.compile(r"\(\?!\[\\d\.\]\)")                                # (?![\d.]) rejects a sentence-ending period
    assert not any(lookahead.search(rx) for s in by_id.values() for rx in s.get("answer_all", []) + s.get("answer_any", []))


@needs_pack
def test_message_pack_matches_scenarios(tmp_path, monkeypatch):
    committed = {f: open(os.path.join(go.PACK_DIR, f), encoding="utf-8").read() for f in os.listdir(go.PACK_DIR)}
    monkeypatch.setattr(go, "PACK_DIR", str(tmp_path))
    go.make_pack(None)
    fresh = {f: (tmp_path / f).read_text(encoding="utf-8") for f in os.listdir(tmp_path)}
    assert fresh == committed, "agent_scenarios.py가 바뀌었으면 python grade_opencode.py --make-pack 으로 메시지 팩을 다시 만드세요"
    spec_, by_id, order = go.load_scenarios()
    assert len(order) == len(set(order)) and all(go.TAG.fullmatch(f"[{s}]") for s in order)


def test_missing_scenario_file_names_the_cause(tmp_path, monkeypatch):
    """C:\\CLI-dev에 agent_scenarios.py가 빠진 채 반입되면: 원인(반입 누락)을 적은 오류, 테스트는 실패 대신 건너뜀."""
    monkeypatch.setattr(go, "SCENARIOS_FILE", str(tmp_path / "agent_scenarios.py"))
    with pytest.raises(go.ScenarioFileMissing, match="반입"):
        _load_scenarios()
    with pytest.raises(pytest.skip.Exception, match="반입 누락"):
        go.load_scenarios()


def test_scenarios_module_is_data_only():
    """agent_scenarios.py는 데이터 상수만 (함수·import 없음): 채점기가 경로로 읽어 실행한다."""
    import ast
    tree = ast.parse(open(go.SCENARIOS_FILE, encoding="utf-8").read())
    assert not [n for n in ast.walk(tree) if isinstance(n, (ast.Import, ast.ImportFrom, ast.FunctionDef, ast.Call))]
    spec_, _, _ = go.load_scenarios()
    spec_["groups"].clear()                                       # 읽은 쪽이 고쳐도 다음 읽기는 그대로
    assert go.load_scenarios()[0]["groups"]


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


def test_report_table_escapes_pipes(tmp_path, monkeypatch):
    """Failed check names hold regexes with '|': the summary table row must keep its column count."""
    monkeypatch.setattr(go, "HERE", str(tmp_path))
    rows = [{"id": "X2", "결과": "실패", "명령수": 0, "초": None, "실패": "답변: a|b|c"}]
    out = go.write_report("t", ["x.json"], rows, [])
    path = out[0] if isinstance(out, (list, tuple)) else out
    text = open(path, encoding="utf-8-sig").read()
    line = [ln for ln in text.splitlines() if ln.startswith("| X2 ")][0]
    assert line.count("|") - line.count("\\|") == 6, line


def _answer_pass(sc, ans):
    """answer_all, answer_any and answer_none of one scenario on one answer (as grade_turn checks them)."""
    return (all(re.search(rx, ans, re.I) for rx in sc.get("answer_all", []))
            and (not sc.get("answer_any") or any(re.search(rx, ans, re.I) for rx in sc["answer_any"]))
            and not any(re.search(rx, ans, re.I | re.M) for rx in sc.get("answer_none", [])))


def _x2_corpus():
    with open(os.path.join(os.path.dirname(__file__), "x2_answers.json"), encoding="utf-8") as f:
        return json.load(f)


def test_x2_answer_corpus():
    """X2 정답·오답 말뭉치(good/bad, 실제 답변 형태 114개): answer_all과 answer_none으로 모두 바르게 가려야 함."""
    sc = go.load_scenarios()[1]["X2"]
    corpus = _x2_corpus()
    assert len(corpus["good"]) + len(corpus["bad"]) >= 110
    wrong = [a for a in corpus["good"] if not _answer_pass(sc, a)] + [a for a in corpus["bad"] if _answer_pass(sc, a)]
    assert not wrong, wrong


def test_x2_holdout_is_reference_only():
    """holdout 20개는 정규식을 맞출 때 보지 않은 답변이다 (8차: 14/20). 자유 문장은 정규식으로 다 가릴 수 없으므로
    X2 답변 판정은 참고(answer_reference)다. 여기서는 정확도가 떨어지지 않는지만 본다 (holdout에 맞춰 고치지 말 것)."""
    sc = go.load_scenarios()[1]["X2"]
    hold = _x2_corpus()["holdout"]
    right = sum(_answer_pass(sc, a) for a in hold["good"]) + sum(not _answer_pass(sc, a) for a in hold["bad"])
    assert len(hold["good"]) == len(hold["bad"]) == 10
    assert right >= 14, right
    assert sc.get("answer_reference") is True


def test_answer_reference_verdict(tmp_path, monkeypatch):
    """answer_reference(X2): 답변 정규식 결과는 '통과(참고)'/'실패(참고)', 명령·금지 규칙 위반은 그대로 '실패'.
    보고서에 최종 판정은 --judge나 사람이 정한다는 줄이 있고, --judge 점수가 있으면 결과 칸에 함께 보인다."""
    monkeypatch.setattr(go, "HERE", str(tmp_path))
    good = "이상치 기준을 바꾸는 옵션은 없습니다. '제외' 열에 이유를 적으세요."
    bad = "네, --outlier-k 옵션으로 바꿀 수 있습니다. 제외 열도 됩니다."
    for answer, tools, want in ((good, [], "통과(참고)"), (good, [_bash('python -c "print(1)"', {})], "실패"),
                                (bad, [], "실패(참고)")):
        path = _export(tmp_path, [("[X2] 옵션 있어?", tools + [{"type": "text", "text": answer}])])
        meta, rows, detail = go.grade([path], "t", False, "LLM42")
        x2 = next(r for r in rows if r["id"] == "X2")
        assert x2["결과"] == want, (answer, x2)
        if want == "실패":
            assert "임의 파이썬" in x2["실패"] and "참고 답변" not in x2["실패"]
        if want == "실패(참고)":
            assert x2["실패"].startswith("참고 답변")
    report, npass, n = go.write_report("t", meta, rows, detail)
    text = open(report, encoding="utf-8").read()
    assert "--judge" in text and "답변 판정이 참고인 요청: X2 실패" in text
    assert "- [OK(참고)] 답변: 제외" in text and "- [FAIL(참고)] 답변에 없어야 함: --" in text
    assert "- [OK] 임의 파이썬" in text
    rows = [dict(r, 점수=2) if r["id"] == "X2" else r for r in rows]
    report = go.write_report("t2", meta, rows, detail)[0]
    text = open(report, encoding="utf-8").read()
    assert [ln for ln in text.splitlines() if ln.startswith("| X2 ")][0].startswith("| X2 | 실패(참고) 2점 |")
    rows = [dict(r, 결과="통과(참고)", 실패="") if r["id"] == "X2" else r for r in rows]
    text = open(go.write_report("t3", meta, rows, detail)[0], encoding="utf-8").read()
    assert "| X2 | 통과(참고) 2점 |" in text and "X2 통과(참고) (채점 LLM 2점)" in text


def test_answer_variants_other_scenarios():
    """P1/P2/P3/M1/M3/R1 답변 규칙: 도구 preview 그대로 전달한 답과 흔한 바꿔 쓰기는 통과, 값이 틀리거나 경고가 빠지면 실패."""
    by_id = go.load_scenarios()[1]
    cases = {
        "P1": (["생산효율 76.33 (90% 범위 72.1~80.5), 실측 비교값 76.1",
                "생산효율 76.33 (90% 범위 72.1~80.5), 측정 데이터 비교: 76.0",
                "| 출력 | 예측 | 90% 범위 | 실측 비교값 |\n|---|---|---|---|\n| 생산효율 | 76.33 | 72~80 | 76.1 |",
                "생산효율 76.3 / 소비전력 42.9 / 불순물 81.6 (실측값과 차이 0.2 이내)",
                "[추론 완료] 소스: 인라인 | 타겟 예측 요약: [생산효율(76.33)] | 인라인 반환 | 실측 비교 경고 1건"],
               ["생산효율 76.3, 소비전력 42.9, 불순물 81.6입니다."]),
        "P2": (["[추론 완료] 소스: 새조건.xlsx | 타겟 예측 요약: [생산효율(75.96)] | 저장: 새조건_pred_x.csv | 외삽 경고 1건",
                "저장: 새조건_pred_x.csv. 외삽 경고 1건: 반응온도 상한 초과(학습 최대 195, 입력 200).",
                "결과는 새조건_pred_x.csv 에 있습니다. 1개 행은 반응온도가 학습 데이터 최댓값을 넘었습니다(외삽).",
                "새조건_pred_x.csv로 저장했습니다. 반응온도 200은 학습 범위 밖이라 참고만 하세요.",
                "새조건_pred_x.csv로 저장했습니다. 반응온도 200은 학습 구간을 벗어납니다."],
               ["새조건_pred_x.csv로 저장했습니다. 평균 생산효율 75.96.",
                "새조건_pred_x.csv로 저장했습니다. 생산효율 75.96 (90% 범위 72.1~80.5)."]),
        "P3": (["모델 기준 불순물=100이 되는 반응온도: 190 | 운전 구간(90% 오차 안): 188.84~191.45",
                "반응온도 189.998. 90% 오차 범위 안의 운전 범위는 188.84~191.449입니다. 보장된 값은 아닙니다.",
                "약 190°C입니다. 운전범위 188.8~191.4°C (모델 예측).",
                "반응온도 약 190도, 구간 188.84~191.45",
                "반응온도는 190.0°C, 운전 구간은 188.84~191.45°C입니다.",
                "반응온도 190°C, 허용 운전 폭 188.8~191.4°C",
                "반응온도 190, operating window 188.8~191.4"],
               ["반응온도는 190°C입니다."]),
        "M1": (["격자 보간: B=16.25, C=87.5",
                "A=12.5일 때 B는 16.25, C는 87.5입니다.",
                "| 열 | 값 |\n|---|---|\n| B | 16.25 |\n| C | 87.50 |",
                "B = 16.250, C = 87.500",
                "**B**: `16.25`, **C**: `87.5`"],
               ["보간 결과 B 116.25, C 187.5", "B는 16.25, C는 87.55입니다.", "B는 10, C는 90입니다."]),
        "M3": (["격자 보간: B=1058.5 [외삽]",
                "B는 1,058.5입니다 (외삽).",
                "B = 1058.50 (범위 초과, 참고용)",
                "B는 1058.5입니다. A가 데이터 범위(0~100)를 넘어서 끝 구간을 직선 연장한 추정값입니다.",
                "B 1058.5 (extrapolated: A가 최대 100보다 3 큼)",
                "B는 1058.5입니다. A가 데이터 범위를 초과합니다."],
               ["B는 1058.5입니다.", "B는 11058.5입니다 (외삽).", "B는 1000입니다 (범위 밖)."]),
        "R1": (["[보고서 완료] v1 / 타겟 3개 / 종합 판정 양호 / 저장: 공정_보고서.docx | 성능예측표: 공정_성능예측표.xlsx, 일치도 92%",
                "보고서와 성능 예측표를 만들었습니다. 일치도 92%.",
                "성능예측표(공정_성능예측표.xlsx)를 만들었고 생산효율 일치도는 95%입니다.",
                "보고서: a.docx, 성능예측표: b.xlsx. 출력별 일치도 생산효율 92%, 불순물 88%.",
                "보고서와 성능 예측표 생성 완료 (일치도 93%, 판정 양호)"],
               ["보고서를 만들었습니다: 공정_보고서.docx", "성능예측표를 만들었습니다."]),
    }
    for sid, (goods, bads) in cases.items():
        for a in goods:
            assert _answer_pass(by_id[sid], a), (sid, a)
        for a in bads:
            assert not _answer_pass(by_id[sid], a), (sid, a)


# ---------------------------------------------------------------------------
# opencode run --format json 기록 (OpenCode 2.x: 셸 도구 이름 shell, 파일 경로 키 path)
# ---------------------------------------------------------------------------
def _run_events(sid, events):
    lines = ["opencode 안내 글 (JSON 아님)"]
    for i, (kind, part) in enumerate(events):
        part = dict(part, sessionID="ses_x", messageID=f"msg_{i // 2}")
        lines.append(json.dumps({"type": kind, "timestamp": 1000 * (i + 1), "sessionID": "ses_x", "part": part}, ensure_ascii=False))
    return "\n".join(lines) + "\n"


def _shell(cmd, out):
    return ("tool_use", {"type": "tool", "tool": "shell", "state": {"status": "completed", "input": {"command": ENC + cmd},
                                                                   "output": json.dumps(out, ensure_ascii=False)}})


def test_read_run_maps_shell_to_bash_and_path_key(tmp_path):
    p = tmp_path / "S0.jsonl"
    p.write_text(_run_events("S0", [
        ("step_start", {"type": "step-start"}),
        _shell('python "C:\\CLI-dev\\ml-kit\\tools\\make_beta_samples.py" "C:\\CLI-dev\\work\\gemma"', {"status": "ok"}),
        ("tool_use", {"type": "tool", "tool": "write", "state": {"status": "completed",
                                                                 "input": {"path": "C:\\temp\\x.py", "content": "print(1)"}}}),
        ("text", {"type": "text", "text": "만들었습니다."}),
        ("error", {}),
    ]), encoding="utf-8")
    turn, session, errors = go.read_run(str(p), "S0", "[S0] 요청")
    assert session == "ses_x" and len(errors) == 1
    assert [t["tool"] for t in turn["tools"]] == ["bash", "write"] and turn["tools"][1]["path"] == "C:\\temp\\x.py"
    assert go.final_answer(turn) == "만들었습니다." and turn["end"] - turn["start"] == 4000
    sc = go.load_scenarios()[1]["S0"]
    checks = {n: ok for n, ok, _ in go.grade_turn(sc, turn, go.load_scenarios()[0]["rules"])}
    assert checks["명령: make_beta_samples\\.py"] and not checks[".py 파일 만들기·고치기 없음"]


def test_grade_runs_reads_folder_and_marks_missing(tmp_path):
    (tmp_path / "S0.jsonl").write_text(_run_events("S0", [
        _shell('python "C:\\CLI-dev\\ml-kit\\tools\\make_beta_samples.py" "C:\\CLI-dev\\work\\gemma"', {"status": "ok"}),
        ("text", {"type": "text", "text": "샘플을 만들었습니다."})]), encoding="utf-8")
    (tmp_path / go.RUN_MESSAGES).write_text(json.dumps({"S0": "[S0] 보낸 문장"}, ensure_ascii=False), encoding="utf-8")
    meta, rows, detail = go.grade_runs(str(tmp_path))
    by = {r["id"]: r for r in rows}
    assert by["S0"]["결과"] == "통과" and by["A1"]["결과"] == "없음" and "ses_x" in meta[0]
    assert detail[0][2]["user"] == ["[S0] 보낸 문장"]
    report, npass, n = go.write_report("t", meta, rows, detail, out_dir=str(tmp_path))
    assert os.path.dirname(report) == str(tmp_path) and (npass, n) == (1, 1)


def test_reruns_asked_by_continue_are_not_repeats():
    """ml-kit가 시간 제한으로 CONTINUE(continue: true)를 내면 같은 명령을 다시 실행하는 것이 맞음 (R10): 반복 위반 아님.
    CONTINUE 없이 같은 명령을 3번 이상 실행하면 그대로 위반."""
    cmd = 'python "C:\\CLI\\ml-kit\\ml_export.py" -m "C:\\x\\v1"'
    cont = {"status": "ok", "continue": True, "agent_directive": "CONTINUE: ..."}
    done = {"status": "ok", "agent_directive": "VERIFIED: ..."}
    rules = go.load_scenarios()[0]["rules"]

    def turn(outs):
        tools = [go.tool_entry(_bash(cmd, o)) for o in outs]
        return {"tags": ["R2"], "user": ["[R2]"], "tools": tools, "texts": ["끝"], "text_at": [(0, len(tools))]}
    ok = {n: v for n, v, _ in go.grade_turn({}, turn([cont, cont, cont, done]), rules)}
    assert ok["같은 명령 3번 이상 반복 없음"]
    bad = {n: v for n, v, _ in go.grade_turn({}, turn([done, done, done]), rules)}
    assert not bad["같은 명령 3번 이상 반복 없음"]


# ---------------------------------------------------------------------------
# office-cli 묶음 O (회사 PC 오류 로그: 경량 LLM이 옵션을 지어낸 경우) - opencode run 기록으로 합격/불합격
# ---------------------------------------------------------------------------
G = "C:\\CLI-dev\\work\\gemma"
EX, WD, OL, RD, OF = (f'python "C:\\CLI\\office-cli\\{n}.py"' for n in ("excel", "word", "outlook", "reader", "office"))
OK = {"status": "ok", "agent_directive": "VERIFIED: ok"}
ARG_ERR = {"status": "error", "error_type": "invalid_argument", "agent_directive": "STOP_RETRY: unrecognized arguments"}
CHECK = f'{EX} check --target "{G}\\규격대조.xlsx" --range B3:C7 --item-col B --value-col C --spec "{G}\\규격서.txt"'
MEMO = f'{EX} comment --target "{G}\\규격대조.xlsx" --action add --file "C:\\CLI\\office-cli\\data\\_work\\check_memo_검사.csv"'
O1_ANSWER = ("5행 대조: 일치 2, 불일치 2, 근거없음 1. 불일치: 소비 전류 250 (규격서.txt 3.2, 200 mA 이하), "
             "동작 온도 70 (3.3, -20~60 ℃). 메모 5개를 미리보기했습니다. 이대로 달까요?")

# 시나리오: (합격 기록들, [(불합격 기록, 실패 이유에 들어갈 글)]). 기록 = [(명령, 결과 JSON), ..., 답변]
OFFICE_CASES = {
    "O1": ([[(CHECK, OK), (MEMO + " --dry-run", OK), O1_ANSWER],
            [(CHECK.replace('"', ""), OK), (MEMO + " --dry-run", OK), O1_ANSWER]],
           [([(CHECK, OK), (MEMO + " --dry-run", OK), (MEMO, OK), O1_ANSWER], "금지 명령"),        # 확인 없이 실제로 담
            ([(CHECK, OK), (MEMO, OK), O1_ANSWER], "--dry-run"),                                   # 미리보기 없이 바로 담
            ([(CHECK, OK), "불일치 2건입니다. 소비 전류, 동작 온도. 규격서 3.2, 3.3"], "--dry-run"),   # 메모 미리보기 없음
            ([(f'{RD} search --query "정격 전압" --file "{G}\\규격서.txt"', OK), (MEMO + " --dry-run", OK), O1_ANSWER], "check"),
            # 회귀: ;로 이은 다른 명령의 --dry-run을 보고 실제 달기를 놓침
            ([(CHECK, OK), (f"{MEMO}; {MEMO} --dry-run", OK), O1_ANSWER], "금지 명령"),
            # 회귀: 메모 글(--text) 안의 "--dry-run" 글자를 옵션으로 봄
            ([(CHECK, OK), (f'{EX} comment --target "{G}\\규격대조.xlsx" --action add --range C5 --text "불일치 (--dry-run 아님)"', OK),
              O1_ANSWER], "금지 명령")]),
    "O2": ([[(f'{EX} write --target "{G}\\단위표.xlsx" --cells "C6=ft; C7=℃"', OK), "C6에 ft, C7에 ℃를 넣었습니다."],
            [(f"{EX} write -t '{G}\\단위표.xlsx' --cells 'C6 = ft;C7 = ℃'", OK), "C6, C7에 썼습니다."]],
           [([(f'{EX} write --target "{G}\\단위표.xlsx" --cell C6 --text ft', ARG_ERR),
              (f'{EX} write --target "{G}\\단위표.xlsx" --cells "C6=ft; C7=℃"', OK), "C6, C7 완료"], "없는 명령"),
            ([(f'{EX} write --target "{G}\\단위표.xlsx" --range C6 --value ft', OK),
              (f'{EX} write --target "{G}\\단위표.xlsx" --range C7 --value ℃', OK), "C6, C7 완료"], "--cells"),
            ([(f'{EX} write --target "{G}\\단위표.xlsx" --cells "C6=ft"', OK),
              (f'{EX} write --target "{G}\\단위표.xlsx" --cells "C7=℃"', OK), "C6, C7 완료"], "금지 명령"),
            ([(f'{EX} write --target "{G}\\단위표.xlsx" --cells "C6=ft; C7=°C"', OK), "C6, C7 완료"], "--cells")]),
    "O3": ([[(f"{OL} search --query 시험성적서", OK), (f'{OL} attachments --id 1 --save "{G}\\메일첨부"', OK),
             (f'{RD} read --file "{G}\\메일첨부\\시험성적서.docx"', OK),
             f"{G}\\메일첨부 에 시험성적서.docx를 저장했습니다. 판정: 합격"]],
           [([(f'{OL} attachments --id 1 --save "{G}\\메일첨부"', OK), (f'{WD} read --target "{G}\\메일첨부\\시험성적서.docx"', OK),
              "메일첨부에 저장했고 판정은 합격입니다."], "reader"),
            ([(f'{OL} attachments --id 1 --save "{G}\\메일첨부"', OK), (f'{RD} read --file "{G}\\메일첨부\\시험성적서.docx"', OK),
              (f'{OL} reply --body "확인했습니다" --send', OK), "메일첨부, 합격. 답장도 보냈습니다."], "금지 명령")]),
    "O4": ([[(f'{WD} write --target "{G}\\주간보고.docx" --text "다음 주 계획: 시험 장비 교정"', OK), "문서 끝에 덧붙였습니다."],
            [(f'{WD} write --target "{G}\\주간보고.docx" --file "{G}\\추가.md"', OK), "추가했습니다."]],
           [([(f'{OF} write --target "{G}\\주간보고.docx" --text "다음 주 계획: 시험 장비 교정"', ARG_ERR),
              (f'{WD} write --target "{G}\\주간보고.docx" --text "다음 주 계획: 시험 장비 교정"', OK), "덧붙였습니다."], "없는 명령"),
            ([(f'{OF} transfer --from word --to word --text "다음 주 계획: 시험 장비 교정"', ARG_ERR), "추가하지 못했습니다."], "없는 명령"),
            ([(f'{WD} write --target "{G}\\주간보고.docx" --text "다음 주 계획: 시험 장비 교정" --pos replace', OK), "넣었습니다."],
             "금지 명령"),
            # 회귀: 앵커 뒤에 넣기는 문서 중간 (요청은 문서 끝)
            ([(f'{WD} write --target "{G}\\주간보고.docx" --anchor "주간 보고" --pos after --text "다음 주 계획: 시험 장비 교정"', OK),
              "덧붙였습니다."], "금지 명령"),
            # 문서 자체를 --file로 줌 (--target이 아님)
            ([(f'{WD} write --file "{G}\\주간보고.docx" --text "다음 주 계획: 시험 장비 교정"', OK), "덧붙였습니다."], "word")]),
    "O5": ([[(f'{WD} fill --target "{G}\\성적서양식.docx" --file "{G}\\값.csv" --dry-run', OK),
             (f'{WD} fill --target "{G}\\성적서양식.docx" --file "{G}\\값.csv"', OK), "업체명 한빛전자, 시험일 2026-11-04를 채웠습니다."]],
           [([(f'{WD} fill --target "{G}\\성적서양식.docx" --file "{G}\\값.csv" --dry-run', OK), "미리보기입니다. 채울까요?"], "fill"),
            ([(f'{WD} replace --target "{G}\\성적서양식.docx" --find "{{{{업체명}}}}" --replace 한빛전자', OK), "채웠습니다."], "fill"),
            ([(f'{WD} fill --target "{G}\\성적서양식.docx" --text "업체명=한빛전자"', ARG_ERR),
              (f'{WD} fill --target "{G}\\성적서양식.docx" --file "{G}\\값.csv"', OK), "한빛전자로 채웠습니다."], "금지 명령"),
            # 회귀: 다른 문서를 채우고 값 파일 이름에 양식 이름이 들어 있음
            ([(f'{WD} fill --target "{G}\\성적서양식.docx" --file "{G}\\값.csv" --dry-run', OK),
              (f'{WD} fill --target "{G}\\다른.docx" --file "{G}\\성적서양식.docx_values.csv"', OK), "한빛전자로 채웠습니다."], "fill")]),
}


def _office_run(tmp_path, sid, record):
    """record ([(명령, 결과), ..., 답변])를 opencode run 기록 <sid>.jsonl 하나로 쓰고 grade_runs의 그 요청 행을 돌려줍니다."""
    for f in tmp_path.glob("*.jsonl"):
        f.unlink()
    events = [_shell(c, out) for c, out in record[:-1]] + [("text", {"type": "text", "text": record[-1]})]
    (tmp_path / f"{sid}.jsonl").write_text(_run_events(sid, events), encoding="utf-8")
    rows = go.grade_runs(str(tmp_path))[1]
    return next(r for r in rows if r["id"] == sid)


def test_office_scenarios_pass_and_fail(tmp_path):
    """묶음 O: 맞는 명령(check -> comment --dry-run, write --cells 한 번, attachments --save -> reader read, word.py write,
    word.py fill --file)은 통과. 지어낸 옵션, 칸마다 반복, 확인 없이 메모 달기, 다른 도구로 대신하기는 실패."""
    by_id = go.load_scenarios()[1]
    assert {s for s in by_id if s.startswith("O")} == set(OFFICE_CASES)
    for sid, (goods, bads) in OFFICE_CASES.items():
        for rec in goods:
            row = _office_run(tmp_path, sid, rec)
            assert row["결과"] == "통과", (sid, rec, row["실패"])
        for rec, why in bads:
            row = _office_run(tmp_path, sid, rec)
            assert row["결과"] == "실패" and why in row["실패"], (sid, rec, row["실패"])


def test_command_segments_split_outside_quotes():
    assert go.command_segments('a --cells "C6=ft; C7=x"; b --dry-run && c | d') == [
        'a --cells "C6=ft; C7=x"', "b --dry-run", "c", "d"]
    assert go.command_segments("a --cells 'C6 = ft;C7 = x'") == ["a --cells 'C6 = ft;C7 = x'"]
    assert go.cmd_hit(r"comment\b(?![^\n]*--dry-run)", "comment --add; other --dry-run")


def test_office_scenarios_judge_by_command_not_skill_name():
    """채점은 스킬 이름(office-cli, office-excel 등)에 기대지 않는다: 명령 정규식과 금지 규칙에 스킬 이름이 없음."""
    by_id = go.load_scenarios()[1]
    for sid in OFFICE_CASES:
        sc = by_id[sid]
        regexes = sc.get("commands", []) + [sc.get("tool_ok", "")] + sc.get("forbid", [])
        assert not any(re.search(r"office-(cli|excel|doc|mail|spec)", rx) for rx in regexes), sid


def test_no_fake_office_rule():
    """회사 PC 오류 로그에서 나온 지어낸 office-cli 명령·옵션은 어느 요청에서든 공통 규칙 위반."""
    for c in (f'{EX} write --range C6 --text ft', f'{EX} write --cell C6 --value ft', f'{OF} write --text "x"',
              f'{OF} transfer --source excel --to word', f'{OF} transfer --from word --to excel --text "x"',
              f'{OF} undo --target "{G}\\a.xlsx"', f'{EX} list', 'cd C:\\CLI\\office-cli; python excel.py write --cell C6',
              # 1.4.79: transfer --to/--from, ppt slide --to는 옛 이름 (지금은 --to-app/--from-app, --dest)
              f'{OF} transfer --from excel --range A1:D10 --to word', f'{OF} transfer --from-app excel --range A1 --to word',
              'python "C:\\CLI\\office-cli\\ppt.py" slide --action move --slide 1 --to 3'):
        assert _rule_hits("no_fake_office", c), c
    for c in (f'{EX} write --cells "C6=ft; C7=℃"', f'{EX} comment --action add --range B5 --text "의견"',
              f'{WD} write --text "## 제목"', f'{OF} transfer --from-app excel --range A1:D10 --to-app word',
              f'{OF} transfer --from-app word --from-target "{G}\\a.docx" --to-app excel --to-target active',
              'python "C:\\CLI\\office-cli\\ppt.py" slide --action move --slide 1 --dest 3',
              'python "C:\\CLI\\office-cli\\outlook.py" search --query x --from 홍길동',
              f'{OF} undo --list', f'{EX} write --cells C6=ft; {WD} write --text "x"', f'{EX} read --range A1:B2',
              'python "C:\\CLI\\ml-kit\\ml_train.py" -f a.xlsx'):
        assert not _rule_hits("no_fake_office", c), c


def test_make_office_samples_is_allowed_script():
    for c in (f'python "C:\\CLI-dev\\office-cli\\tools\\make_office_samples.py" "{G}"',
              f'cd C:\\CLI-dev\\office-cli\\tools; python make_office_samples.py "{G}"'):
        assert not _rule_hits("no_adhoc_py", c), c


def test_manual_scenario_is_marked_and_left_out_of_bulk(tmp_path, monkeypatch):
    """manual 시나리오(O3, Outlook 메일 준비): 기록이 없으면 '수동 전용'으로 없음, 메시지 팩은 준비 안내를 붙이고 한 번에 보내기에서 뺌."""
    by_id = go.load_scenarios()[1]
    assert [s for s in by_id if by_id[s].get("manual")] == ["O3"]
    rows = {r["id"]: r for r in go.grade_runs(str(tmp_path))[1]}
    assert rows["O3"]["결과"] == "없음" and "수동 전용" in rows["O3"]["실패"]
    assert "수동 전용" not in rows["O1"]["실패"]
    monkeypatch.setattr(go, "PACK_DIR", str(tmp_path / "pack"))
    go.make_pack(None)
    text = (tmp_path / "pack" / "O_office-cli_문서_작업.md").read_text(encoding="utf-8")
    single, bulk = text.split("## 한 번에 보내기")
    assert "수동 준비: Outlook" in single and f"{G}\\시험성적서.docx" in single and "[O3]" in single
    assert "[O3]" not in bulk and "[O1]" in bulk and "[O5]" in bulk
    assert "S -> A -> M -> P -> R -> O -> X" in (tmp_path / "pack" / "README.md").read_text(encoding="utf-8")


def test_office_samples_match_scenarios(tmp_path):
    """S1 샘플(make_office_samples.py)이 시나리오가 기대하는 파일과 값을 만든다 (O1 대조 항목, O2의 빈 칸, O5의 자리)."""
    openpyxl = pytest.importorskip("openpyxl")
    docx = pytest.importorskip("docx")
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), "office-cli", "tools"))
    try:
        import make_office_samples
    finally:
        sys.path.pop(0)
    made = make_office_samples.main(str(tmp_path))
    assert set(made) == {"규격서.txt", "규격대조.xlsx", "단위표.xlsx", "주간보고.docx", "성적서양식.docx", "시험성적서.docx"}
    ws = openpyxl.load_workbook(tmp_path / "규격대조.xlsx").active
    assert [ws[f"B{r}"].value for r in range(3, 8)] == ["정격 전압", "소비 전류", "동작 온도", "절연 저항", "포장 색상"]
    assert ws["D2"].value == "단위" and ws["D3"].value == "V"       # check가 머리글 '단위' 열을 행마다 읽음
    ws = openpyxl.load_workbook(tmp_path / "단위표.xlsx").active
    assert ws["C6"].value is None and ws["C7"].value is None and ws["C5"].value == "단위"
    form = "\n".join(p.text for p in docx.Document(str(tmp_path / "성적서양식.docx")).paragraphs)
    assert "{{업체명}}" in form and "{{시험일}}" in form
    assert "판정: 합격" in "\n".join(p.text for p in docx.Document(str(tmp_path / "시험성적서.docx")).paragraphs)
    by_id = go.load_scenarios()[1]
    for sid, files in (("O1", ["규격대조.xlsx", "규격서.txt"]), ("O2", ["단위표.xlsx"]), ("O4", ["주간보고.docx"]),
                       ("O5", ["성적서양식.docx"])):
        assert all(f in by_id[sid]["message"] for f in files), sid
