# -*- coding: utf-8 -*-
"""OpenCode 대화 채점기 (사람이 실행하는 개발 도구). 실행 기록 묶음(write_bundle)은 run_all_tests.py gemma 단계가 씀

채점:
    python grade_opencode.py --export <대화.json> [<대화2.json> ...] [--label 회사PC] [--judge]
    python grade_opencode.py --runs <기록 폴더> [--label 회사PC]   # run_all_tests.py --gemma가 남긴 opencode run 기록(<ID>.jsonl)

대화 내보내기: PowerShell에서 `cmd /c "opencode export <세션ID> > chat.json"`
  (cmd의 `>`는 opencode가 쓴 UTF-8 바이트를 그대로 저장한다. PowerShell 5.1의 `| Out-File`이나 `>`는 출력을
   콘솔 코드 페이지(cp949)로 읽어 한글이 깨진다. 채점기는 한글이 깨진 대화를 찾으면 [ERROR]로 멈춘다.)

채점 기준은 agent_scenarios.py 한 곳에 있다. 사용자 메시지 앞의 "[ID]" 표시로 시나리오를 찾고,
그 요청부터 다음 표시가 나오기 전까지의 명령·도구 결과·답변을 본다.
답변 검사(answer_all, answer_any, answer_clean)는 최종 답변만 본다: 그 시나리오의 tool_ok/commands에 맞는 마지막 명령
뒤의 글 전부 (맞는 명령이 없으면 마지막 도우미 메시지의 마지막 명령 뒤 글). 답 뒤에 명령을 하나 더 실행하고
"완료했습니다."로 끝내도 답은 채점된다. 맞는 명령 전의 진행 안내는 채점하지 않는다.
한 번에 보낸 대화(메시지 1개에 [ID] 여러 개)는 도구 실행 순서로 시나리오마다 나눠 채점하고 결과에 "참고"를 붙인다
(시나리오마다 처음 맞는 명령부터 다음 시나리오의 명령 전까지. 첫 시도가 실패해 다시 실행한 경우처럼 앞에 일부만 맞는
명령이 있으면 그 명령부터. 명령으로 찾을 수 없는 시나리오는 대화 전체로 채점).

판정 (도구로 확인, LLM 판단 없음):
  - 필요한 명령을 실행했는가 (commands), 도구 결과 JSON이 ok인가 (tool_ok)
  - 답변에 꼭 들어갈 값이 있는가 (answer_all, answer_any)
  - 공통 금지: 패키지 설치, py/python3, 임의 파이썬 코드, 도구 코드 읽기, .py 쓰기, office-cli에 없는 명령·옵션(지어낸 것),
    같은 명령 반복, 답변의 이모지·file:///·같은 말 반복
  - manual이 있는 시나리오(O3, Outlook 메일 준비 필요)는 run_all_tests.py --gemma가 보내지 않는다. 기록이 없으면 "없음"(수동 전용)
  - answer_reference: true인 시나리오(X2처럼 자유 문장 답)는 답변 정규식 판정을 참고로만 보인다 ("통과(참고)"/"실패(참고)").
    최종 판정은 --judge 점수나 사람이 답변을 읽고 정한다. 명령·금지 규칙 위반은 그대로 "실패".
--judge: 사내 OpenAI 호환 API(completions)로 시나리오별 1~5점과 이유를 붙인다 (참고용, 통과 여부에는 넣지 않음).
  환경변수 JUDGE_BASE_URL (예: http://서버/v1), JUDGE_API_KEY(없으면 생략), --judge-model (기본 LLM42).

결과: 이 파일 옆에 grade_<라벨>_<시각>.md (요약 + 실패 상세 + 명령 기록)와 .xlsx (시나리오별 표).
"""
import argparse
import codecs
import json
import os
import re
import sys
import urllib.request
from collections import Counter
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
SCENARIOS_FILE = os.path.join(HERE, "agent_scenarios.py")      # 데이터만 담은 모듈 (회사 반입 시스템이 .yaml을 막음)
TAG = re.compile(r"\[([A-Z]\d+)\]")
EMOJI = re.compile("[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\u2460-\u24FF\u2190-\u21FF]"
                   r"|\\(?:right|left|up|down|Right|Left|Up|Down)arrow\b")      # LaTeX 화살표 ($\downarrow$, 집 PC 06)
REPEAT = re.compile(r"(\S{1,10})(?:\s*\1){7,}")
# python [옵션] <스크립트.py>: the script a python call runs (checked against no_adhoc_py.script_allow)
PY_SCRIPT = re.compile(r"""(?<![\w.-])python(?:\.exe)?"?(?:\s+-(?:[XW]\s*\S+|[bBdEiIOPqsSuv]+))*\s+"""
                       r"""("[^"]*\.py"|'[^']*\.py'|[^\s"']*\.py)(?=["'\s;|&)]|$)""", re.I)
WRITE_TOOLS = ("write", "edit", "multiedit", "patch", "apply_patch")
TOOL_ALIASES = {"shell": "bash"}           # OpenCode 2.x는 셸 도구 이름이 shell (채점은 bash 하나로 봄)
PATCH_FILE = re.compile(r"^\*\*\* (?:(?:Add|Update|Delete) File|Move to):\s*(.+?)\s*$", re.M)   # file lines of a patch tool's text
ENC_PREFIX = re.compile(r"^(\s*(\$env:\w+\s*=\s*[^;]+|\[Console\]::\w+\s*=\s*[^;]+);)+\s*")   # PowerShell encoding line from the global rules


class ScenarioFileMissing(FileNotFoundError):
    """agent_scenarios.py가 없음 (반입 때 빠짐)."""


def load_scenarios():
    """agent_scenarios.py의 SCENARIOS -> (spec, {id: 시나리오}, [id 순서]). 파일은 경로로 읽는다 (sys.path와 상관없음)."""
    import copy
    import importlib.util
    if not os.path.isfile(SCENARIOS_FILE):
        raise ScenarioFileMissing(f"시나리오 파일이 없습니다: {SCENARIOS_FILE} (반입 때 빠졌는지 확인하세요. "
                                  f"저장소의 AIKIT\\dev\\agent_scenarios.py를 이 폴더에 두면 됩니다)")
    mod_spec = importlib.util.spec_from_file_location("agent_scenarios", SCENARIOS_FILE)
    mod = importlib.util.module_from_spec(mod_spec)
    mod_spec.loader.exec_module(mod)
    spec = copy.deepcopy(mod.SCENARIOS)                  # 부르는 쪽이 고쳐도 다음 읽기에 섞이지 않게
    for g in spec["groups"]:                             # 묶음의 paused(잠시 시험에서 뺌)는 그 요청마다 적어 둠
        for s in g["scenarios"]:
            if g.get("paused"):
                s["paused"] = g["paused"]
    order = [s for g in spec["groups"] for s in g["scenarios"]]
    return spec, {s["id"]: s for s in order}, [s["id"] for s in order]


# ------------------------------------------------------------------------------------------------
# 대화 읽기
class ExportError(Exception):
    """The export file could not be read as JSON (wrong encoding, not an OpenCode export, ...) or its Korean text is garbled."""


EXPORT_CMD = 'cmd /c "opencode export <세션ID> > chat.json"'
EXPORT_HINT = f"PowerShell에서 `{EXPORT_CMD}`로 다시 내보내세요 (cmd의 >는 UTF-8 바이트를 그대로 저장합니다)."
HANGUL = re.compile("[\uac00-\ud7a3]")
HANJA = re.compile("[\u4e00-\u9fff\uf900-\ufaff]")


def load_export(path):
    """Reads an OpenCode export: UTF-8 with or without a BOM (cmd /c ... > file), or UTF-16 with a BOM (PowerShell `>`).
    The encoding of the file says nothing about garbled text inside it: see check_mojibake."""
    text = _read_text(path)
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ExportError(f"대화 파일이 JSON이 아닙니다 ({e.msg}, {e.lineno}행): {path}. {EXPORT_HINT}")
    if not isinstance(data, dict) or not isinstance(data.get("messages"), list):
        raise ExportError(f"OpenCode 대화 내보내기 형식이 아닙니다 (messages 목록 없음): {path}. {EXPORT_HINT}")
    return data


def _read_text(path):
    """UTF-8 (BOM 있거나 없거나) 또는 BOM 있는 UTF-16 파일을 글로 읽습니다."""
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError as e:
        raise ExportError(f"대화 파일을 열 수 없습니다: {path} ({type(e).__name__}: {e})")
    if raw.startswith(codecs.BOM_UTF8):
        encodings = ["utf-8-sig"]
    elif raw.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        encodings = ["utf-16"]
    else:
        encodings = ["utf-8", "utf-16"]
    text, tried = None, []
    for enc in encodings:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            tried.append(enc)
    if text is None:
        raise ExportError(f"대화 파일의 인코딩을 알 수 없습니다 ({', '.join(tried)} 실패): {path}. {EXPORT_HINT}")
    return text.lstrip("\ufeff")


def tool_entry(p):
    """OpenCode tool part -> 채점용 도구 기록 (내보내기와 run 기록 공통). 파일 경로는 filePath(1.x) 또는 path(2.x)."""
    st = p.get("state", {}) or {}
    inp = st.get("input", {}) or {}
    name = p.get("tool")
    return {"tool": TOOL_ALIASES.get(name, name), "status": st.get("status"), "input": inp,
            "command": inp.get("command", ""), "path": inp.get("filePath") or inp.get("path") or "",
            "output": st.get("output") or st.get("error") or "",
            "exit": (st.get("metadata") or {}).get("exit")}


def read_run(path, sid, message):
    """`opencode run --format json` 기록 하나(요청 1개, 한 줄에 사건 하나)를 채점용 차례 하나로 읽습니다.
    요청 문장은 기록에 없으므로 보낸 message를 씁니다. (차례, 세션ID, 오류 사건 목록)을 돌려줍니다.
    JSON이 아닌 줄(opencode 안내 글)은 건너뜁니다."""
    turn = {"tags": [sid], "user": [message], "tools": [], "texts": [], "text_at": [], "start": None, "end": None,
            "file": os.path.basename(path), "bulk": False, "shared": []}
    msgs, session, errors = {}, None, []
    for line in _read_text(path).splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            ev = json.loads(line)
        except json.JSONDecodeError:
            continue
        if not isinstance(ev, dict):
            continue
        ts = ev.get("timestamp")
        if isinstance(ts, (int, float)):
            turn["start"] = ts if turn["start"] is None else turn["start"]
            turn["end"] = ts
        session = session or ev.get("sessionID")
        part = ev.get("part") or {}
        mi = msgs.setdefault(part.get("messageID"), len(msgs))
        kind = ev.get("type")
        if kind == "text" and str(part.get("text", "")).strip():
            turn["texts"].append(part["text"])
            turn["text_at"].append((mi, len(turn["tools"])))
        elif kind == "tool_use" or part.get("type") == "tool":
            turn["tools"].append(tool_entry(part))
        elif kind == "error":
            errors.append(json.dumps(ev.get("error") or ev, ensure_ascii=False)[:300])
    return turn, session, errors


def read_turns(path):
    """Splits one export into turns: a user message with a [ID] tag plus everything after it up to the next tag.
    Each text part keeps (assistant message number, tools seen before it) so the final answer and bulk splits can be found."""
    data = load_export(path)
    turns, cur = [], None
    for mi, m in enumerate(data.get("messages", [])):
        info, parts = m.get("info", {}), m.get("parts", [])
        t = info.get("time", {})
        if info.get("role") == "user":
            text = "\n".join(p.get("text", "") for p in parts if p.get("type") == "text")
            tags = list(dict.fromkeys(TAG.findall(text)))
            if tags or cur is None:
                cur = {"tags": tags, "user": [text], "tools": [], "texts": [], "text_at": [],
                       "start": t.get("created"), "end": t.get("created"), "file": os.path.basename(path)}
                turns.append(cur)
            else:
                cur["user"].append(text)               # follow-up without a tag (e.g. answering Gemma's question)
            continue
        if cur is None:
            continue
        cur["end"] = t.get("completed") or t.get("created") or cur["end"]
        for p in parts:
            if p.get("type") == "text" and p.get("text", "").strip():
                cur["texts"].append(p["text"])
                cur["text_at"].append((mi, len(cur["tools"])))
            elif p.get("type") == "tool":
                cur["tools"].append(tool_entry(p))
    return data.get("info", {}), turns


def check_mojibake(path, turns, by_id):
    """Raises ExportError when the export's Korean text was garbled on the way to the file (opencode's UTF-8 output read as
    cp949 by PowerShell 5.1 `| Out-File` or `>`). Looks at the tagged user messages, whose text is known from agent_scenarios.py:
    no Hangul at all, few of the expected Hangul syllables, or Hanja / replacement characters the messages do not have."""
    tagged = [(t, [by_id[x]["message"] for x in t["tags"] if x in by_id]) for t in turns]
    tagged = [(t, msgs) for t, msgs in tagged if msgs]
    if not tagged:
        return
    got = "\n".join("\n".join(t["user"]) for t, _ in tagged)
    want = "\n".join(m for _, msgs in tagged for m in msgs)
    expected = set(HANGUL.findall(want))
    reason = None
    if expected and not HANGUL.search(got):
        reason = "요청 메시지에 한글이 없음"
    elif expected and len(expected & set(HANGUL.findall(got))) < len(expected) / 2:
        reason = "요청 메시지의 한글이 시나리오 문장과 거의 다름"
    elif "\ufffd" in got or (HANJA.search(got) and not HANJA.search(want)):
        reason = "요청 메시지에 깨진 문자(한자, U+FFFD)가 있음"
    if reason:
        raise ExportError(f"대화 파일의 한글이 깨졌습니다 ({reason}): {path}. PowerShell 5.1의 `| Out-File`이나 `>`는 "
                          f"opencode 출력을 cp949로 읽어 한글을 깨뜨립니다. {EXPORT_HINT}")


def final_answer(turn, sc=None):
    """The final answer: every text after the last tool call that matches the scenario's tool_ok/commands regexes
    (a closing note such as "완료했습니다." after one more tool does not hide the answer before it).
    Without such a call (or without sc): the text parts of the last assistant message that has text, after its last tool call.
    Progress notes before the scenario's own command are not graded."""
    if not turn["texts"]:
        return ""
    rxs = ([sc["tool_ok"]] if sc and sc.get("tool_ok") else []) + list((sc or {}).get("commands") or [])
    hits = [i for i, c in bash_commands(turn) if any(re.search(rx, c, re.I) for rx in rxs)]
    if hits:
        after = [t for t, (_, p) in zip(turn["texts"], turn["text_at"]) if p > hits[-1]]
        if after:
            return "\n".join(after)
    last = turn["text_at"][-1]
    return "\n".join(t for t, at in zip(turn["texts"], turn["text_at"]) if at == last)


_SEGMENT = re.compile(r"""(?:"[^"\n]*"?|'[^'\n]*'?|[^;&|\n"'])+""")


def command_segments(cmd):
    """One shell line -> the commands joined by ; && || | or newlines (separators inside quotes are kept:
    --cells "C6=ft; C7=ft" stays one command). Scenario commands/forbid regexes are matched per command, so a
    lookahead such as (?![^\n]*--dry-run) does not see another command's --dry-run."""
    return [m.group(0).strip() for m in _SEGMENT.finditer(str(cmd)) if m.group(0).strip()] or [str(cmd)]


def cmd_hit(rx, cmd):
    return any(re.search(rx, seg, re.I) for seg in command_segments(cmd))


def bash_commands(turn):
    return [(i, ENC_PREFIX.sub("", t["command"])) for i, t in enumerate(turn["tools"]) if t["tool"] == "bash"]


def locate(sc, cmds, start, prev=None):
    """(first, hit) for this scenario among the tool calls at or after start, or None when the scenario has no command
    regexes or nothing matches. hit is the first call matching all its command regexes, otherwise any of them.
    Calls that fully match the previous scenario's regexes (prev) are that scenario's own (re)runs: a hit that is not
    such a call is preferred, and one that is counts only when nothing else matches.
    first moves back from hit to an earlier partial attempt (e.g. a STOP_RETRY first try before the retry the tool
    recommended), which belongs to this scenario too, but not past a call that fully matches prev.
    The caller searches the next scenario after hit (not after first), so it never starts inside this scenario's retry."""
    rxs = sc.get("commands") or []
    if not rxs:
        return None

    def full(regexes, c):
        return bool(regexes) and all(re.search(rx, c, re.I) for rx in regexes)

    for skip_prev in (True, False):
        for need in (all, any):
            hit = next((i for i, c in cmds if i >= start and need(re.search(rx, c, re.I) for rx in rxs)
                        and not (skip_prev and full(prev, c))), None)
            if hit is None:
                continue
            floor = start
            for j, d in cmds:
                if start <= j < hit and full(prev, d):
                    floor = j + 1
            first = next((j for j, d in cmds if floor <= j < hit and any(re.search(rx, d, re.I) for rx in rxs)), hit)
            return first, hit
    return None


def split_bulk(turn, by_id):
    """Splits a turn that carries several [ID] tags (one bulk message) into one sub-turn per scenario by the order of tool calls:
    a scenario owns the tools from its first matching command up to the next located scenario's first command.
    Its answer is the text after its first command up to the next scenario's; with no such text, the closing summary of
    the whole turn (so a lead-in before the first command is the first scenario's answer only when it is the turn's last text).
    Scenarios that cannot be located (no command regex or no match) get the whole turn, marked split=False."""
    tags = [x for x in turn["tags"] if x in by_id]
    cmds = bash_commands(turn)
    starts, pos, prev = {}, 0, None
    for sid in tags:
        found = locate(by_id[sid], cmds, pos, prev)
        if found is not None:
            starts[sid] = found[0]
            pos, prev = found[1] + 1, by_id[sid].get("commands")
    located = [sid for sid in tags if sid in starts]
    out = {}
    for sid in tags:
        sub = dict(turn, tags=[sid], bulk=True, shared=[x for x in turn["tags"] if x != sid])
        if sid not in starts:
            out[sid] = dict(sub, split=False)
            continue
        k = located.index(sid)
        lo = starts[sid]
        hi = starts[located[k + 1]] if k + 1 < len(located) else len(turn["tools"])
        idx = [j for j, (_, p) in enumerate(turn["text_at"]) if lo < p <= hi]
        texts, text_at = [turn["texts"][j] for j in idx], [turn["text_at"][j] for j in idx]
        if not texts and turn["texts"]:                 # no own text: the closing summary of the whole turn (a lead-in
                                                        # before the first command counts only when it is the turn's last text)
            texts, text_at = [final_answer(turn)], [(turn["text_at"][-1][0], len(turn["tools"]))]
        text_at = [(m, p - lo) for m, p in text_at]     # tool positions relative to this sub-turn's tools
        out[sid] = dict(sub, split=True, tools=turn["tools"][lo:hi], texts=texts, text_at=text_at, start=None, end=None)
    return out


def last_json(output):
    for line in reversed(str(output).splitlines()):
        line = line.strip()
        if line.startswith("{") and line.endswith("}"):
            try:
                return json.loads(line)
            except json.JSONDecodeError:
                continue
    return None


# ------------------------------------------------------------------------------------------------
# 채점
def continued_runs(turn):
    """Commands run again because the previous run of the same command answered CONTINUE (continue: true, R10):
    the tool asked for that rerun, so it is not a repeat."""
    out, last = [], {}
    for t in turn["tools"]:
        if t["tool"] != "bash":
            continue
        cmd = ENC_PREFIX.sub("", t["command"])
        key = re.sub(r"\s+", " ", cmd).strip()
        if last.get(key):
            out.append(cmd)
        payload = last_json(t["output"])
        last[key] = bool(payload and payload.get("continue"))
    return out


def grade_turn(sc, turn, rules):
    """Returns a list of (check name, ok, detail) for one scenario on one turn."""
    cmds = [c for _, c in bash_commands(turn)]
    answer = final_answer(turn, sc)
    out = []
    for rx in sc.get("commands", []):
        hit = any(cmd_hit(rx, c) for c in cmds)
        out.append((f"명령: {rx}", hit, "" if hit else "맞는 명령 없음"))
    if sc.get("tool_ok") and not sc.get("tool_plain"):       # tool_plain: 사용자 스크립트 (1줄 JSON을 내지 않음, Y1)
        runs = [t for t in turn["tools"] if t["tool"] == "bash" and re.search(sc["tool_ok"], t["command"], re.I)]
        payload = last_json(runs[-1]["output"]) if runs else None
        ok = bool(payload) and payload.get("status") == "ok"
        detail = "실행 없음" if not runs else "JSON 없음" if not payload else str(payload.get("agent_directive", ""))[:200]
        out.append(("도구 결과 ok", ok, "" if ok else detail))
    for rx in sc.get("answer_all", []):
        hit = bool(re.search(rx, answer, re.I))
        out.append((f"답변: {rx}", hit, "" if hit else "답변에 없음"))
    if sc.get("answer_any"):
        hit = any(re.search(rx, answer, re.I) for rx in sc["answer_any"])
        out.append(("답변: " + " 또는 ".join(sc["answer_any"]), hit, "" if hit else "답변에 없음"))
    for rx in sc.get("answer_none", []):
        bad = re.search(rx, answer, re.I | re.M)
        out.append((f"답변에 없어야 함: {rx}", not bad, bad.group(0)[:100] if bad else ""))
    for rx in sc.get("forbid", []):
        bad = [c for c in cmds if cmd_hit(rx, c)]
        out.append((f"금지 명령: {rx}", not bad, "; ".join(bad)[:300]))

    allow = set(sc.get("allow", []))
    for rid, rule in rules.items():
        if rid in allow:
            continue
        bad = []
        if rule.get("command"):                        # one regex or a list of them (any match is a violation)
            rx = rule["command"] if isinstance(rule["command"], str) else "|".join(f"(?:{r})" for r in rule["command"])
            bad += [c for c in cmds if re.search(rx, c, re.I | re.M)]
        if rule.get("script_allow"):
            # 겹친 경로 구분자(C:\\AIKIT\\ml-kit\\ml_report.py: JSON 이스케이프를 그대로 옮긴 것)도 Windows에서는 같은 파일
            bad += [c for c in cmds if any(not re.fullmatch(rule["script_allow"],
                                                            re.sub(r"[\\/]{2,}", lambda g: g.group(0)[0], m.group(1).strip("\"'")), re.I)
                                           for m in PY_SCRIPT.finditer(c)) and c not in bad]
        if rule.get("read"):
            bad += [t["path"] for t in turn["tools"] if t["tool"] == "read" and re.search(rule["read"], t["path"], re.I)]
        if rule.get("write"):                          # OpenCode write/edit/multiedit tools (filePath) and patch (paths in the patch text)
            for t in turn["tools"]:
                if t["tool"] in WRITE_TOOLS:
                    paths = [t["path"]] if t["path"] else PATCH_FILE.findall(" ".join(str(v) for v in t["input"].values()))
                    bad += [p for p in paths if re.search(rule["write"], p.strip(), re.I)]
        if rule.get("search"):                         # OpenCode grep: where it looks (path, include; not the search text)
            for t in turn["tools"]:                    # glob only lists file names: its path/pattern must match rule["glob"]
                if t["tool"] in ("grep", "glob"):
                    inp = t["input"]
                    if t["tool"] == "glob":
                        where, rx = [str(inp.get(k) or "") for k in ("path", "pattern")], rule.get("glob", rule["search"])
                    else:
                        where, rx = [str(inp.get(k) or "") for k in ("path", "include")], rule["search"]
                        if (re.search(r"\.md$", str(inp.get("include") or "").strip(), re.I)   # *.md, SKILL.md: documents only
                                and not re.search(r"\.py\b", str(inp.get("path") or ""), re.I)):  # unless path is a .py file
                            continue
                    if any(re.search(r"SKILL\.md$", w, re.I) for w in where):
                        continue
                    if any(re.search(rx, w, re.I) for w in where):
                        bad.append(f"{t['tool']} " + " ".join(f"{k}={v}" for k, v in inp.items()))
        if rule.get("max_repeat"):                     # a run right after a CONTINUE result is the tool's own request (R10)
            asked = continued_runs(turn)
            counts = Counter(re.sub(r"\s+", " ", c).strip() for c in cmds)
            for c in asked:
                counts[re.sub(r"\s+", " ", c).strip()] -= 1
            limit = int(sc.get("max_repeat") or rule["max_repeat"])       # 시나리오가 정한 상한 (Y1: 고쳐 다시 실행 2번까지 = 3회)
            bad += [f"{n}회: {c}" for c, n in counts.items() if n > limit]
        if rid == "answer_clean":
            bad += [f"이모지 {m}" for m in set(EMOJI.findall(answer))]
            bad += ["file:/// 링크"] if "file:///" in answer else []
            bad += [f"반복 '{m.group(0)[:40]}'" for m in REPEAT.finditer(answer) if re.search(r"[^\W\d_]", m.group(1))][:3]   # words, not table rules
        out.append((rule["title"], not bad, "; ".join(map(str, bad))[:300]))
    if not answer:
        out.append(("최종 답변 있음", False, "답변 문장이 없음"))
    return out


def judge(sc, turn, rubric, model):
    base = os.environ.get("JUDGE_BASE_URL", "").rstrip("/")
    if not base:
        return None, "JUDGE_BASE_URL 환경변수가 없어 건너뜀"
    tools = "\n".join(f"$ {t['command'] or t['tool'] + ' ' + t['path']}\n{str(t['output'])[-1500:]}" for t in turn["tools"])[-6000:]
    prompt = ("너는 업무 도우미 AI의 답변을 채점하는 평가자다.\n"
              f"[채점 기준]\n{rubric}\n\n[사용자 요청]\n{chr(10).join(turn['user'])}\n\n"
              f"[도우미가 실행한 명령과 결과]\n{tools or '(없음)'}\n\n[도우미의 최종 답변]\n{final_answer(turn, sc)[-3000:]}\n\n"
              "기준에 맞으면 5, 전혀 맞지 않으면 1이다. 아래 형식으로 두 줄만 쓴다.\n점수: <1~5>\n이유: <한 문장>\n")
    body = json.dumps({"model": model, "prompt": prompt, "max_tokens": 200, "temperature": 0}).encode("utf-8")
    req = urllib.request.Request(base + "/completions", data=body, headers={"Content-Type": "application/json"})
    if os.environ.get("JUDGE_API_KEY"):
        req.add_header("Authorization", "Bearer " + os.environ["JUDGE_API_KEY"])
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            text = json.loads(r.read().decode("utf-8"))["choices"][0]["text"]
    except Exception as e:
        return None, f"채점 API 오류: {type(e).__name__}: {e}"[:200]
    m = re.search(r"점수\s*[:：]\s*([1-5])", text)
    why = re.search(r"이유\s*[:：]\s*(.+)", text)
    return (int(m.group(1)) if m else None), (why.group(1).strip() if why else text.strip())[:300]


def grade(paths, label, use_judge, model):
    spec, by_id, order = load_scenarios()
    found, meta = {}, []
    for p in paths:
        info, turns = read_turns(p)
        check_mojibake(p, turns, by_id)
        meta.append(f"{os.path.basename(p)} (모델 {info.get('model', {}).get('id', '?')}, OpenCode {info.get('version', '?')})")
        for t in turns:
            tags = [x for x in t["tags"] if x in by_id]
            if len(tags) > 1:                            # bulk message: split by the order of tool calls
                for tag, sub in split_bulk(t, by_id).items():
                    found.setdefault(tag, []).append(sub)
            elif tags:
                found.setdefault(tags[0], []).append(dict(t, bulk=False, shared=[x for x in t["tags"] if x != tags[0]]))
    return meta, *score(found, spec, by_id, order, use_judge, model)


RUN_MESSAGES = "messages.json"     # run_all_tests.py --gemma가 기록 폴더에 남기는 {요청 ID: 보낸 문장}
RUN_MODELS = "models.json"         # 같은 기록 폴더의 {요청 ID: 답한 모델} (--gemma-model, 사용량 제한 때 예비 모델)


def grade_runs(folder, use_judge=False, model="LLM42"):
    """opencode run 기록 폴더(<ID>.jsonl + messages.json)를 채점합니다. 요청마다 run 한 번이라 나눌 필요가 없습니다.
    기록이 없는 요청은 "없음"."""
    spec, by_id, order = load_scenarios()
    try:
        with open(os.path.join(folder, RUN_MESSAGES), encoding="utf-8") as f:
            sent = json.load(f)
    except (OSError, ValueError):
        sent = {}
    found, sessions = {}, set()
    for sid in order:
        path = os.path.join(folder, f"{sid}.jsonl")
        if not os.path.isfile(path):
            continue
        turn, session, _ = read_run(path, sid, sent.get(sid) or f"[{sid}] " + by_id[sid]["message"])
        sessions.add(session or "?")
        found[sid] = [turn]
    meta = [f"opencode run 기록 {folder} ({len(found)}건, 세션 {', '.join(sorted(sessions)) or '없음'})"]
    try:                                    # run_all_tests.py --gemma-model: 요청별로 실제로 답한 모델
        with open(os.path.join(folder, RUN_MODELS), encoding="utf-8") as f:
            models = json.load(f)
        if models:
            meta.append("요청별 모델: " + ", ".join(f"{k}={v}" for k, v in models.items()))
    except (OSError, ValueError):
        pass
    rows, detail = score(found, spec, by_id, order, use_judge, model)
    return meta, rows, detail



BUNDLE_OUTPUT = 3000      # 기록 묶음: 도구 결과 하나의 글자 상한 (마지막 JSON 1줄, 없으면 출력 끝)
BUNDLE_LIMIT = 3_000_000  # 기록 묶음 전체 글자 상한
FENCE = "`" * 3


def _mask(text):
    """결과 파일과 같은 가림 (경로 속 사용자 폴더, Office 작성자). run_all_tests.py를 못 읽으면 그대로."""
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("run_all_tests_mask", os.path.join(HERE, "run_all_tests.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.mask_private(text)
    except Exception:
        return text


def _block(text):
    return ["", FENCE + "text", text.replace(FENCE, "'" * 3), FENCE]


def write_bundle(folder, out_path):
    """opencode run 기록 폴더(요청마다 <ID>.jsonl, <ID>.stderr.txt 등 여러 파일)를 반출용 md 하나로 묶습니다.
    요청마다 보낸 문장, 답한 모델, 모델의 글과 도구 호출을 일어난 순서대로(명령, 결과 JSON 전체 또는 출력 끝), 오류 사건, stderr 끝.
    결과 파일(result_*.md)의 채점 표는 줄인 것이라 되묻기·[후보] 처리·지침 위반처럼 대화 흐름을 봐야 하는 검토에 씀."""
    spec, by_id, order = load_scenarios()

    def _load(name):
        try:
            with open(os.path.join(folder, name), encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}
    sent, models = _load(RUN_MESSAGES), _load(RUN_MODELS)
    ids = [sid for sid in order if os.path.isfile(os.path.join(folder, f"{sid}.jsonl"))]
    L = ["# OpenCode 실행 기록 묶음", "", f"- 기록 폴더: {folder}", f"- 요청 {len(ids)}건: {', '.join(ids) or '없음'}",
         "- 요청마다 모델의 글(>)과 도구 호출(명령, 결과)을 일어난 순서대로 적었습니다. 결과는 마지막 JSON 1줄 전체(없으면 출력 끝)입니다.", ""]
    for sid in ids:
        msg = sent.get(sid) or f"[{sid}] " + by_id[sid]["message"]
        turn, session, errors = read_run(os.path.join(folder, f"{sid}.jsonl"), sid, msg)
        L += [f"## {sid}", "", f"- 모델: {models.get(sid, '(opencode 설정)')}, 세션: {session or '없음'}", f"- 요청: {msg}", ""]
        texts = list(zip(turn["texts"], turn["text_at"]))

        def _texts_before(k):
            out = []
            for t, (_, before) in texts:
                if before == k:
                    out += ["> " + t.strip().replace("\n", "\n> "), ""]
            return out
        for k, t in enumerate(turn["tools"]):
            L += _texts_before(k)
            what = ENC_PREFIX.sub("", t["command"]) or f"{t['tool']} {t['path']}".strip()
            L.append(f"{k + 1}. 도구 {t['tool']} ({t.get('status') or '?'}"
                     + (f", exit {t['exit']}" if t.get("exit") is not None else "") + f"): `{what[:600]}`")
            payload = last_json(t["output"]) if t["tool"] == "bash" else None
            res = json.dumps(payload, ensure_ascii=False) if payload else str(t["output"] or "")[-BUNDLE_OUTPUT:]
            if res.strip():
                L += _block(res[:BUNDLE_OUTPUT])
            L.append("")
        L += _texts_before(len(turn["tools"]))
        if errors:
            L += ["오류 사건:", ""] + [f"- {e}" for e in errors] + [""]
        err_path = os.path.join(folder, f"{sid}.stderr.txt")
        err = _read_text(err_path).strip() if os.path.isfile(err_path) else ""
        if err:
            L += ["stderr 끝:"] + _block(err[-1500:]) + [""]
        limited = sorted(n for n in os.listdir(folder) if n.startswith(f"{sid}.limited") and n.endswith(".jsonl"))
        if limited:
            L += [f"사용량 제한으로 예비 모델에 다시 보냄: {', '.join(limited)} (이름만)", ""]
    text = "\n".join(L) + "\n"
    if len(text) > BUNDLE_LIMIT:
        text = text[:BUNDLE_LIMIT] + f"\n\n(묶음이 {BUNDLE_LIMIT}자를 넘어 잘림)\n"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(_mask(text))
    return out_path, len(ids)

def score(found, spec, by_id, order, use_judge, model):
    """{요청 ID: [차례, ...]} -> (rows, detail). 요청마다 마지막으로 보낸 차례로 채점합니다."""
    rows, detail = [], []
    rules = spec["rules"]
    for sid in order:
        sc = by_id[sid]
        if sid not in found:
            why = (f"시험에서 뺀 요청: {sc['paused']}" if sc.get("paused") else
                   "수동 전용 요청 (자동 실행에서는 보내지 않음)" if sc.get("manual") else
                   "준비 요청 (run_all_tests.py가 직접 실행, 모델에 보내지 않음)" if sc.get("setup") else "대화에 이 요청이 없음")
            rows.append({"id": sid, "결과": "없음", "실패": why, "명령수": 0, "초": None, "보낸횟수": 0})
            continue
        turn = found[sid][-1]                            # graded on the last time the request was sent
        checks = grade_turn(sc, turn, rules)
        ref = bool(sc.get("answer_reference"))            # answer regexes are a reference verdict only (free-text answers)
        hard = [f"{n} ({d})" if d else n for n, ok, d in checks if not ok and not (ref and is_answer_check(n))]
        soft = [f"참고 {n} ({d})" if d else f"참고 {n}" for n, ok, d in checks if not ok and ref and is_answer_check(n)]
        note = ""
        if turn.get("bulk"):
            note = "참고: 한 번에 보낸 요청, " + ("명령 순서로 나눠 채점" if turn.get("split") else "명령으로 나눌 수 없어 대화 전체로 채점")
        result = "실패" if hard else (("실패(참고)" if soft else "통과(참고)") if ref else "통과")
        row = {"id": sid, "결과": result, "실패": " | ".join(hard + soft), "명령수": len(turn["tools"]),
               "초": round((turn["end"] - turn["start"]) / 1000) if turn["start"] and turn["end"] else None,
               "보낸횟수": len(found[sid]), "같이 보낸 요청": ",".join(turn.get("shared", [])), "참고": note}
        if use_judge:
            row["점수"], row["채점 이유"] = judge(sc, turn, sc.get("judge") or spec["default_judge"], model)
        rows.append(row)
        detail.append((sid, sc, turn, checks))
    return rows, detail


def is_answer_check(name):
    """answer_all/answer_any/answer_none checks: only these become reference verdicts in answer_reference scenarios."""
    return name.startswith(("답변: ", "답변에 없어야 함: "))


def safe_label(label):
    """Keeps only characters that are safe in a Windows file name (same rule as run_all_tests.py)."""
    return "".join(c for c in str(label) if c.isalnum() or c in "-_") or "PC"


def write_report(label, meta, rows, detail, out_dir=None):
    label = safe_label(label)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(out_dir or HERE, f"grade_{label}_{stamp}")
    graded = [r for r in rows if r["결과"] != "없음"]
    npass = sum(r["결과"] in ("통과", "통과(참고)") for r in graded)
    nbulk = sum(bool(r.get("참고")) for r in graded)
    refs = [r for r in graded if r["결과"].endswith("(참고)")]
    scores = [r["점수"] for r in graded if isinstance(r.get("점수"), int)]
    lines = [f"# OpenCode 종합 테스트 채점 ({label})", "", f"- 채점 시각: {datetime.now():%Y-%m-%d %H:%M}"]
    lines += [f"- 대화 파일: {m}" for m in meta]
    lines += [f"- 결과: 채점 {len(graded)}건 중 통과 {npass}건 ({npass / max(len(graded), 1) * 100:.0f}%), 채점하지 않음 {len(rows) - len(graded)}건 (준비·수동 전용 요청 포함, 이유는 표의 실패 항목)"]
    if nbulk:
        lines.append(f"- 한 번에 보낸 요청 {nbulk}건은 결과에 '(참고)'를 붙였습니다. 요청이 한 대화 턴에 섞여 있어 "
                     "명령 순서로 나눠 채점했으므로 하나씩 보낸 결과만큼 정확하지 않습니다.")
    if refs:
        lines.append("- 답변 판정이 참고인 요청: " + ", ".join(
            f"{r['id']} {r['결과']}" + (f" (채점 LLM {r['점수']}점)" if isinstance(r.get("점수"), int) else "") for r in refs)
            + ". 자유 문장 답이라 정규식 판정은 참고이며 통과 수에는 정규식 판정대로 넣었습니다. "
            "최종 판정은 --judge 점수나 사람이 아래 '최종 답변'을 읽고 정하세요 (명령·금지 규칙 위반은 그대로 실패).")
    if scores:
        lines.append(f"- 채점 LLM 평균 점수: {sum(scores) / len(scores):.2f} / 5 ({len(scores)}건, 참고용)")
    has_judge = any("점수" in r for r in rows)
    lines += ["", "| 요청 | 결과 | 명령 수 | 걸린 초 | 실패 항목 |" + (" 점수 |" if has_judge else ""),
              "|---|---|---|---|---|" + ("---|" if has_judge else "")]
    for r in rows:
        cell = r["결과"] + ("(참고)" if r.get("참고") and not r["결과"].endswith("(참고)") else "")
        if r["결과"].endswith("(참고)") and isinstance(r.get("점수"), int):
            cell += f" {r['점수']}점"
        lines.append(f"| {r['id']} | {cell} | {r['명령수']} | {r['초'] if r['초'] is not None else ''} | {(r['실패'] or '').replace('|', chr(92) + '|')} |"
                     + (f" {r.get('점수') if r.get('점수') is not None else ''} |" if has_judge else ""))
    lines += ["", "## 요청별 기록", ""]
    for sid, sc, turn, checks in detail:
        lines += [f"### {sid}", "", "요청: " + " / ".join(u.strip() for u in turn["user"])[:400], ""]
        if turn.get("bulk"):
            lines += [next(r for r in rows if r["id"] == sid)["참고"], ""]
        ref = bool(sc.get("answer_reference"))
        for n, ok, d in checks:
            tag = ("OK" if ok else "FAIL") + ("(참고)" if ref and is_answer_check(n) else "")
            lines.append(f"- [{tag}] {n}" + (f": {d}" if d and not ok else ""))
        lines += ["", "명령:"]
        for t in turn["tools"]:
            what = ENC_PREFIX.sub("", t["command"]) or f"{t['tool']} {t['path']}".strip()
            payload = last_json(t["output"]) if t["tool"] == "bash" else None
            res = f" -> {payload.get('status')} / {str(payload.get('agent_directive', ''))[:60]}" if payload else ""
            lines.append(f"- `{what[:300]}`{res}")
        answer = final_answer(turn, sc).strip() or "(답변 없음)"
        lines += ["", "최종 답변:", "", "> " + answer[:1500].replace("\n", "\n> "), ""]
        row = next(r for r in rows if r["id"] == sid)
        if row.get("채점 이유"):
            lines += [f"채점 LLM: {row.get('점수')}점, {row['채점 이유']}", ""]
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    try:
        import pandas as pd
        pd.DataFrame(rows).to_excel(base + ".xlsx", index=False)
    except Exception as e:
        print(f"[WARN] xlsx 저장 실패: {e}", file=sys.stderr)
    return base + ".md", npass, len(graded)


# ------------------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="OpenCode 대화 채점")
    ap.add_argument("--export", nargs="+", help="OpenCode에서 내보낸 대화 JSON (여러 개 가능)")
    ap.add_argument("--runs", help="run_all_tests.py --gemma가 남긴 opencode run 기록 폴더 (<ID>.jsonl)")
    ap.add_argument("--label", default="PC")
    ap.add_argument("--judge", action="store_true", help="사내 OpenAI 호환 API로 참고 점수 (JUDGE_BASE_URL)")
    ap.add_argument("--judge-model", default="LLM42")
    args = ap.parse_args()
    if not args.export and not args.runs:
        ap.error("--export 또는 --runs 가 필요합니다")
    try:
        if args.runs:
            meta, rows, detail = grade_runs(args.runs, args.judge, args.judge_model)
        else:
            meta, rows, detail = grade(args.export, args.label, args.judge, args.judge_model)
    except ExportError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    report, npass, n = write_report(args.label, meta, rows, detail)
    print(f"[STATUS] 통과 {npass} / {n} -> {report}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
