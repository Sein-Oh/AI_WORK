# -*- coding: utf-8 -*-
"""mbom-diff 조합 스킬(tools/skills/mbom-diff)을 SKILL.md 순서 그대로 실행해, make_mbom_samples의 정답(MOD_CHANGES)과 맞춰 본다.
스킬 스크립트가 낸 next_commands는 글자 그대로 실행한다 (LLM이 하듯)."""
import json
import os
import subprocess
import sys

import pytest

from pk_helpers import command_args, env, ok  # noqa: F401  (env는 fixture)

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.join(os.path.dirname(HERE), "tools")
SKILL = os.path.join(TOOLS, "skills", "mbom-diff")
sys.path.insert(0, TOOLS)
import make_mbom_samples as samples  # noqa: E402


def script(name, *args):
    p = subprocess.run([sys.executable, os.path.join(SKILL, "scripts", name), *map(str, args)],
                       capture_output=True, text=True, encoding="utf-8", env=os.environ.copy())
    lines = [l for l in p.stdout.splitlines() if l.strip()]
    assert len(lines) == 1, (p.stdout, p.stderr)
    return p.returncode, json.loads(lines[0])


@pytest.fixture
def files(env):
    out = env["dir"] / "mbom"
    p = subprocess.run([sys.executable, os.path.join(TOOLS, "make_mbom_samples.py"), "--out", str(out)],
                       capture_output=True, text=True, encoding="utf-8")
    assert p.returncode == 0, p.stderr
    return out


def compare(env, old, new):
    """SKILL.md 1~8단계. (diff 결과 JSON, diff 표의 행들, 요약 JSON)."""
    ok(env, "new")
    ok(env, "select", "--file", old, "--name", "old")
    ok(env, "select", "--file", new, "--name", "new")
    o = ok(env, "export", "--from", "old", "--save", "old.csv")["output_file"]
    n = ok(env, "export", "--from", "new", "--save", "new.csv")["output_file"]
    code, path = script("mbom_path.py", "--old", o, "--new", n)
    assert code == 0 and path["agent_directive"].startswith("VERIFIED:"), path
    assert [c["command"].split()[2] for c in path["next_commands"]] == ["select", "join", "select", "join", "diff", "export"]
    outs = [ok(env, *command_args(c["command"])) for c in path["next_commands"]]
    diff, last = outs[4], outs[5]
    rows = ok(env, "show", "--from", "chg")["rows"]
    code, summary = script("mbom_summary.py", "--file", last["output_file"])
    assert code == 0 and summary["agent_directive"].startswith("VERIFIED:"), summary
    return diff, rows, summary


def norm(v):
    if v is None:
        return None
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return str(v)


def test_v2_mod_changes_are_found_exactly_and_traps_are_not_changes(env, files):
    diff, rows, summary = compare(env, files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2_mod.xlsx")
    got = sorted((r["구분"], r["경로"], r["열"], norm(r["옛 값"]), norm(r["새 값"])) for r in rows)
    want = sorted((k, p, c, norm(a), norm(b)) for k, p, c, a, b in samples.MOD_CHANGES)
    assert got == want            # 행 밀림·순서만 바뀜·'18' -> 18·제목 블록은 나오지 않음
    for r in rows:                # 원본 엑셀 행을 잃지 않음 (join이 엑셀 행을 그대로 둠)
        assert (r["oldp 행"] is not None) == (r["구분"] != "추가") and (r["newp 행"] is not None) == (r["구분"] != "삭제"), r
    purposes = [c["purpose"] for c in diff["confirm_commands"]]   # 두 원본에 칠하는 office-cli 명령, 행과 칸을 나눠 셈
    assert len(purposes) == 2 and "삭제된 3행, 바뀐 칸 8곳" in purposes[0] and "추가된 3행, 바뀐 칸 8곳" in purposes[1], purposes
    assert summary["counts"] == {"이동": 1, "대체 후보": 1, "추가": 1, "삭제": 1, "바뀐 칸": 8}
    text = summary["preview"]
    assert text.startswith("변경 이동 1건, 대체 후보 1건, 추가 1건, 삭제 1건, 바뀐 칸 8곳 (7개 품목).")
    for line in ("[이동] 110040: DR-200/100100 아래 -> DR-200 아래",
                 "[대체 후보] 600160 -> 600161: DR-200/100600 아래 같은 자리",
                 "[추가] 600180: DR-200/100600 아래", "[삭제] 900020: DR-200 아래",
                 "[바뀜] 300213 (DR-200/100300/300210#1 아래): Qty 3 -> 4",
                 "[바뀜] RM-1001 (DR-200/100300/300210#2/300212 아래): Qty 0.045 -> 0.05",
                 "[바뀜] 000123 (DR-200/100500 아래): Qty 2 -> 3"):
        assert line in text, (line, text)
    assert "대체 후보" in summary["agent_directive"]     # 짐작이라 사용자에게 확인


def test_v1_to_v2_and_same_file(env, files):
    _d, _r, summary = compare(env, files / "MBOM_DR200.xlsx", files / "MBOM_DR200_v2.xlsx")
    assert summary["counts"] == {"이동": 0, "대체 후보": 0, "추가": 0, "삭제": 1, "바뀐 칸": 2}
    assert "[삭제] 110050: DR-200/100100 아래" in summary["preview"]
    _d, rows, summary = compare(env, files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2.xlsx")
    assert rows == [] and summary["count"] == 0 and summary["preview"] == "두 판이 같습니다 (찾은 변경 없음)."


def test_path_script_stops_with_column_list_when_level_column_is_missing(env, tmp_path):
    a = tmp_path / "old.csv"
    a.write_text("엑셀 행,Lv,Part Number\n7,0,A\n", encoding="utf-8-sig")
    code, out = script("mbom_path.py", "--old", a, "--new", a)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "Lv" in out["message"]
    code, out = script("mbom_path.py", "--old", a, "--new", a, "--level", "Lv")
    assert code == 1 and "이름(표 이름)이 같습니다" in out["message"]
    code, out = script("mbom_path.py", "--bogus")
    assert code == 2 and out["status"] == "error"


def test_paths_number_repeated_parts_and_warn_level_jumps(env, tmp_path):
    rows = ["엑셀 행,Level,Part Number", "2,0,TOP", "3,1,S", "4,2,P", "5,1,S", "6,2,P", "7,3,X", "8,1,Q", "9,3,Z", "10,x,BAD"]
    for name in ("old", "new"):
        (tmp_path / f"{name}.csv").write_text("\n".join(rows) + "\n", encoding="utf-8-sig")
    code, out = script("mbom_path.py", "--old", tmp_path / "old.csv", "--new", tmp_path / "new.csv")
    assert code == 0 and "S#1, S#2" in out["preview"]
    assert any("건너뛴 행 1개 (엑셀 행 9)" in w for w in out["warnings"])
    assert any("품번이 빈 행 1개" in w and "엑셀 행 10" in w for w in out["warnings"])
    got = (tmp_path / "old_경로.csv").read_text(encoding="utf-8-sig").splitlines()
    assert got == ["엑셀 행,상위 품번,경로", "2,,TOP", "3,TOP,TOP/S#1", "4,S,TOP/S#1/P", "5,TOP,TOP/S#2", "6,S,TOP/S#2/P",
                   "7,P,TOP/S#2/P/X", "8,TOP,TOP/Q", "9,Q,TOP/Q/Z"]
