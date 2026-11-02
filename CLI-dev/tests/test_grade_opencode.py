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


def test_single_request_answer_before_closing_tool(tmp_path):
    """Answer text, one more tool, then "완료했습니다.": the answer after the scenario's own command is graded."""
    ok = {"status": "ok", "agent_directive": "VERIFIED: ok"}
    path = _export(tmp_path, [("[M1] A=12.5", [
        {"type": "text", "text": "맵 파일로 계산하겠습니다."},
        _bash('python "C:\\CLI\\ml-kit\\ml_map.py" -f "맵.xlsx" -i "A=12.5"', ok),
        {"type": "text", "text": "B는 16.25, C는 87.5입니다."},
        _bash('Get-Item "C:\\ml_beta\\agent\\맵.xlsx"', ok),
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
              'Get-ChildItem "C:\\ml_beta\\agent" | Select-String x'):
        assert not _rule_hits("no_read_code", c), c
    assert _rule_hits("no_read_code", tool=_tool("grep", pattern="outlier", path="C:\\CLI\\ml-kit"))
    assert _rule_hits("no_read_code", tool=_tool("grep", pattern="iqr", path="C:/CLI/office-cli", include="*"))
    assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="ml_train.py",
                                                      path="C:\\Users\\me\\.config\\opencode\\skills"))
    assert not _rule_hits("no_read_code", tool=_tool("grep", pattern="제외", path="C:\\CLI\\ml-kit", include="SKILL.md"))
    assert not _rule_hits("no_read_code", tool=_tool("glob", pattern="**/SKILL.md", path="C:\\Users\\me\\.config\\opencode"))


def test_no_write_py_covers_edit_and_patch():
    assert _rule_hits("no_write_py", tool=_tool("write", filePath="C:\\temp\\x.py", content="print(1)"))
    assert _rule_hits("no_write_py", tool=_tool("edit", filePath="C:\\CLI\\ml-kit\\ml_train.py", oldString="a", newString="b"))
    assert _rule_hits("no_write_py", tool=_tool("multiedit", filePath="C:\\CLI\\ml-kit\\ml_map.py", edits=[]))
    patch = "*** Begin Patch\n*** Update File: C:\\CLI\\ml-kit\\ml_map.py\n@@\n-a\n+b\n*** End Patch"
    assert _rule_hits("no_write_py", tool=_tool("patch", patchText=patch))
    assert not _rule_hits("no_write_py", tool=_tool("edit", filePath="C:\\ml_beta\\agent\\memo.md", oldString="a", newString="b"))
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
                 "아니요, 없습니다. '제외' 열에 이유를 적으면 학습에서 빠집니다."):
        assert all(re.search(rx, good) for rx in x2), good
    for bad in ("네, --outlier 옵션으로 바꿀 수 있습니다. '제외' 열도 쓸 수 있습니다.",
                "--iqr-k 인자로 이상치 기준을 조정할 수 있습니다. 제외 열과 함께 쓰세요."):
        assert not all(re.search(rx, bad) for rx in x2), bad
    m2 = by_id["M2"]["answer_all"]
    for good in ("A=20, A=80.", "A는 20과 80입니다.", "A = 20.0 또는 80.0.", "답: 20, 80"):
        assert all(re.search(rx, good) for rx in m2), good
    for bad in ("A=200, A=80.", "A=20.5, A=80", "A=20, A=8.0"):
        assert not all(re.search(rx, bad) for rx in m2), bad
    lookahead = re.compile(r"\(\?!\[\\d\.\]\)")                                # (?![\d.]) rejects a sentence-ending period
    assert not any(lookahead.search(rx) for s in by_id.values() for rx in s.get("answer_all", []) + s.get("answer_any", []))


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


def _run_launcher(tmp_path, child_code, encoding="cp949"):
    """Runs docs/remote-agent/run.py with its WinPython path pointed at this Python and its stdout encoding forced to encoding."""
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
    env = dict(os.environ, PYTHONIOENCODING=encoding)
    return subprocess.run([sys.executable, str(launcher), str(child)], capture_output=True, env=env, timeout=60)


@needs_remote
def test_remote_launcher_writes_in_pipe_encoding(tmp_path):
    """The host may decode the launcher's stdout with the system code page (cp949): the bytes follow sys.stdout.encoding,
    Korean stays readable and a character cp949 lacks becomes '?' instead of an error."""
    odd = chr(0x2022) + " a " + chr(0x2013) + " b" + chr(0xA0) + "c"     # not in cp949: print() on a cp949 pipe fails
    code = f"import json\nprint(json.dumps({{'status': 'ok', 'message': '한글 결과', 'preview': {odd!r}}}, ensure_ascii=False))\n"
    r = _run_launcher(tmp_path, code)
    assert r.returncode == 0, r.stderr
    out = r.stdout.decode("cp949").splitlines()
    payload = json.loads(out[0])
    assert len(out) == 1 and payload["message"] == "한글 결과" and payload["preview"] == "? a ? b?c"
    r = _run_launcher(tmp_path, code, encoding="utf-8")
    out = r.stdout.decode("utf-8").splitlines()
    assert r.returncode == 0 and len(out) == 1 and json.loads(out[0])["preview"] == odd


@needs_remote
def test_remote_launcher_fallback_json(tmp_path):
    for code, exit_code in (("import sys\nsys.exit(0)\n", 1), ("print('Traceback: boom')\n", 1),
                            ("import sys\nsys.stderr.write('x')\nsys.exit(3)\n", 3)):
        r = _run_launcher(tmp_path, code)
        out = r.stdout.decode("cp949").splitlines()
        payload = json.loads(out[0])
        assert len(out) == 1 and payload["status"] == "error" and payload["agent_directive"].startswith("ERROR:")
        assert r.returncode == exit_code, code
