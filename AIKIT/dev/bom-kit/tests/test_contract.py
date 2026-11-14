# -*- coding: utf-8 -*-
"""bom-kit ACI 계약: stdout 1줄 JSON, 도움말·인자 오류도 JSON, 같은 뜻 옵션은 [추천] 전체 명령, 실패는 errors.txt,
pandas-kit을 import하지 않음 (도구 폴더는 단독 배포, 실행만 함), 스킬은 bom.py 명령 하나."""
import ast
import os
import re

from bk_helpers import AIKIT, bom, recommended

KIT = os.path.join(AIKIT, "bom-kit")


def test_help_and_argument_errors_are_one_json_line(env):
    code, out = bom("-h")
    assert code == 0 and out["status"] == "ok"
    code, out = bom("where-used", "--file", "x.csv")                       # --find 빠짐
    assert code == 2 and out["status"] == "error" and out["agent_directive"].startswith("STOP_RETRY:")


def test_same_meaning_option_gives_the_full_fixed_command(env, tmp_path):
    code, out = bom("where-used", "--file", tmp_path / "bom.csv", "--search", "000125")
    assert code == 2 and recommended(out["message"])[:2] == ["where-used", "--file"]
    assert recommended(out["message"])[-2:] == ["--find", "000125"]


def test_failures_go_to_the_error_log(env, tmp_path):
    code, out = bom("where-used", "--file", tmp_path / "없음.xlsx", "--find", "000125")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "파일이 없습니다" in out["message"]
    log = env["dir"] / "bomdata" / "logs" / "errors.txt"
    assert log.is_file() and "bom-kit" in log.read_text(encoding="utf-8")


def test_tool_folder_has_only_its_own_python_files():
    assert sorted(f for f in os.listdir(KIT) if f.endswith(".py")) == ["_bom_runtime.py", "bom.py"]
    for name in ("bom.py", "_bom_runtime.py"):
        tree = ast.parse(open(os.path.join(KIT, name), encoding="utf-8").read())
        mods = {a.name.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.Import) for a in n.names}
        mods |= {n.module.split(".")[0] for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module}
        assert not mods & {"pandas_kit", "_pd_runtime", "_runtime", "_ml_runtime", "pandas", "openpyxl"}, (name, mods)


def test_skills_call_bom_kit_and_have_no_scripts_folder():
    for skill, cmds in (("mbom-diff", ("diff",)), ("where-used", ("where-used",)), ("bom-history", ("history",))):
        folder = os.path.join(AIKIT, ".opencode", "skills", skill)
        assert not os.path.exists(os.path.join(folder, "scripts")), skill        # BOM 계산은 bom-kit 한 곳
        with open(os.path.join(folder, "SKILL.md"), "rb") as f:
            raw = f.read().replace(b"\r\n", b"\n")
        assert len(raw) <= 8192
        text = raw.decode("utf-8")
        used = re.findall(r'python "C:\\AIKIT\\bom-kit\\bom\.py" ([a-z-]+)', text)
        assert used == list(cmds), (skill, used)                         # 명령 하나 (경량 모델이 여러 단계를 잇지 않게)
        assert "pandas_kit.py" not in text, skill


def test_pandas_kit_state_of_the_user_is_not_touched(env, tmp_path):
    """bom-kit이 실행하는 pandas-kit은 데이터 폴더를 따로 써서 사용자의 지금 pandas-kit 작업(지금 표)을 바꾸지 않음."""
    from bk_helpers import make_samples, ok
    files = make_samples(tmp_path / "mbom")
    mine = ok(env, "select", "--file", str(files / "MBOM_DR200.xlsx"), "--name", "mine")["work"]
    code, _ = bom("where-used", "--file", files / "MBOM_DR200_v2.xlsx", "--find", "000125")
    assert code == 0
    assert ok(env, "show", "--from", "mine")["work"] == mine and ok(env, "show", "--from", "mine")["table"] == "mine"


def test_skill_versions_match_the_tool():
    import _bom_runtime as rt
    for skill in ("mbom-diff", "where-used", "bom-history"):
        with open(os.path.join(AIKIT, ".opencode", "skills", skill, "SKILL.md"), encoding="utf-8") as f:
            assert f'version: "{rt.VERSION}"' in f.read(), skill


def test_misspelt_command_name_gives_candidates(env):
    code, out = bom("dif", "--old", "a.xlsx", "--new", "b.xlsx")
    assert code == 2 and out["message"].startswith("CLI 인자 오류: 'dif' 명령은 없습니다. [후보] 1) diff")
    assert "invalid choice" not in out["message"]                    # argparse 영어 문장은 뺌 (0.1.2)
    import shlex                                     # 시험은 BOM_KIT_SHELL=bash: Windows 경로는 '...'로 감싸짐
    assert shlex.split(out["confirm_commands"][0]["command"])[2:4] == ["diff", "--old"]


def _work_folders(folder):
    return sorted(n for n in os.listdir(folder) if os.path.isdir(os.path.join(folder, n)))


def test_results_stay_in_one_result_folder_and_work_folders_are_removed(env, tmp_path):
    """결과는 원본 옆 <원본 이름>_bom 폴더에 실행마다 시각을 붙인 새 이름으로 (덮어쓰지 않음, R4), pandas-kit 작업 폴더(캐시)는
    실행마다 지움 (0.1.2, 사용자 결정 2026-11-14). 알려 준 output_file과 칠할 칸 파일은 다른 실행 뒤에도 남음
    (Fable 검토 B1: 지우면 다음 턴의 confirm_commands가 실패)."""
    import re as _re
    from bk_helpers import command_args, make_samples
    files = make_samples(tmp_path / "mbom")
    f, folder = files / "MBOM_DR200.xlsx", files / "MBOM_DR200_bom"
    stamp = r"_\d{8}_\d{6}(_\d+)?"
    code, out = bom("where-used", "--file", f, "--find", "000125", "--sheet", "Routing")      # 수량 열이 없어 멈춤
    assert code == 1 and "[후보]" in out["message"] and _work_folders(files) == []
    code, first = bom("where-used", "--file", f, "--find", "000125")
    assert code == 0 and _work_folders(files) == ["MBOM_DR200_bom"] and "work" not in first and first["result_folder"] == str(folder)
    assert os.path.dirname(first["output_file"]) == str(folder)
    assert _re.fullmatch("쓰인곳_000125" + stamp + r"\.csv", os.path.basename(first["output_file"]))
    code, again = bom("where-used", "--file", f, "--find", "000125")                          # 같은 질문도 새 파일 (같은 초면 _2)
    assert code == 0 and again["output_file"] != first["output_file"] and os.path.isfile(first["output_file"])
    code, other = bom("where-used", "--file", f, "--find", "나사")
    assert code == 0 and all(os.path.isfile(o["output_file"]) for o in (first, again, other))
    assert _work_folders(files) == ["MBOM_DR200_bom"]
    old, new, v1 = files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2_mod.xlsx", files / "MBOM_DR200.xlsx"
    code, d1 = bom("diff", "--old", old, "--new", new)
    code, d2 = bom("diff", "--old", old, "--new", v1)                                         # 같은 옛 판, 다른 새 판
    code, d3 = bom("diff", "--old", old, "--new", new)                                        # 같은 비교 다시
    assert code == 0 and d3["counts"] == d1["counts"] and len({d["output_file"] for d in (d1, d2, d3)}) == 3
    assert all(os.path.isfile(d["output_file"]) and os.path.isfile(d["changes_file"]) for d in (d1, d2, d3))
    assert sorted(_work_folders(files)) == ["MBOM_DR200_bom", "MBOM_DR200_v2_bom"]
    names = sorted(os.listdir(files / "MBOM_DR200_v2_bom"))
    assert len(names) == 1 + 3 * 5 and ".created_by_bom-kit" in names
    assert all(_re.fullmatch(r"diff_(MBOM_DR200_v2_mod|MBOM_DR200)" + stamp + r"_(요약\.txt|변경내용\.csv|칠할칸_(옛판|새판)\.csv|기록\.json)", n)
               for n in names if n != ".created_by_bom-kit"), names
    stem = d1["output_file"][:-len("_요약.txt")]
    paint = [command_args(c["command"]) for c in d1["confirm_commands"]]                       # 칠하는 명령은 결과 폴더의 파일
    assert [p[p.index("--file") + 1] for p in paint] == [f"{stem}_칠할칸_{s}.csv" for s in ("옛판", "새판")]
    assert all(os.path.isfile(p[p.index("--file") + 1]) for p in paint)
    runs = env["dir"] / "bomdata" / "pandas-kit" / "runs"
    assert not runs.exists() or not os.listdir(runs)


def test_two_bom_runs_at_once_both_finish(env, tmp_path):
    """bom-kit 두 실행을 동시에 (Fable 검토 2026-11-14 A2): pandas-kit 지금 작업 포인터를 같이 써서 둘 다 멈췄음.
    이제 실행마다 pandas-kit 데이터 폴더를 따로 써서 둘 다 끝나고, 원본 옆에는 결과 폴더만 남음."""
    import json as _json
    import subprocess
    import sys
    from bk_helpers import BOM, make_samples
    files = make_samples(tmp_path / "mbom")
    jobs = [["diff", "--old", files / "MBOM_DR200_v2.xlsx", "--new", files / "MBOM_DR200_v2_mod.xlsx"],
            ["where-used", "--file", files / "MBOM_DR200_v2.xlsx", "--find", "000125"]]
    procs = [subprocess.Popen([sys.executable, BOM, *map(str, j)], stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                              env=os.environ.copy()) for j in jobs]
    outs = [_json.loads(p.communicate(timeout=300)[0].decode("utf-8").strip().splitlines()[-1]) for p in procs]
    assert [o["status"] for o in outs] == ["ok", "ok"], outs
    assert _work_folders(files) == ["MBOM_DR200_v2_bom"]


def test_a_work_folder_with_an_open_file_is_left_whole_and_removed_next_time(env, tmp_path, monkeypatch):
    """안의 파일이 열려 있는 폴더 (집 PC 2026-11-14 B9): 반쯤 지우지 않고 그대로 두며, 기억했다가 다음 실행 때 지움.
    Windows는 열린 파일이 든 폴더의 이름을 바꾸지 못함 -> 여기서는 os.rename이 실패하는 것으로 흉내."""
    import json
    import bom as bom_mod
    folder = tmp_path / "MBOM_2611141834"
    (folder / "tables").mkdir(parents=True)
    for name in (bom_mod.MARKER, "bom.csv", "bom_쓰인곳.csv", os.path.join("tables", "s1.pkl")):
        (folder / name).write_text("x", encoding="utf-8")
    real = os.rename

    def locked(src, dst):
        raise PermissionError(32, "다른 프로세스가 사용 중", str(src))
    monkeypatch.setattr(bom_mod.os, "rename", locked)
    assert bom_mod.PandasKit.drop(str(folder), "시험") is False
    assert sorted(p.name for p in folder.rglob("*")) == sorted([bom_mod.MARKER, "bom.csv", "bom_쓰인곳.csv", "tables", "s1.pkl"])
    bom_mod._remember_pending(str(folder))
    bom_mod._retry_pending()                                       # 아직 열림: 그대로 기억
    assert folder.is_dir() and json.loads(open(bom_mod._works_path(), encoding="utf-8").read())[bom_mod.PENDING] == [str(folder)]
    monkeypatch.setattr(bom_mod.os, "rename", real)                # 파일을 닫음
    bom_mod._retry_pending()
    assert not folder.exists() and not list(tmp_path.glob("*.지우는중*"))
    assert bom_mod.PENDING not in json.loads(open(bom_mod._works_path(), encoding="utf-8").read())


def test_level_that_is_not_a_finite_number_is_a_skipped_row(env, tmp_path):
    """Level 칸의 'inf', 'nan' 글자가 내부 오류(OverflowError)를 냈음 (0.1.1): 보지 않은 행으로 알림."""
    import bom as bom_mod
    assert bom_mod.level_of("inf") is None and bom_mod.level_of("nan") is None and bom_mod.level_of("2.0") == 2
    assert bom_mod.qty_of("inf") is None and bom_mod.qty_of("1,000") == 1000
    p = tmp_path / "e.csv"
    p.write_text("Level,Part Number,품명,Qty\n0,TOP,제품,1\n1,A,조립,2\ninf,B,볼트,3\n2,C,볼트,4\n", encoding="utf-8")
    code, out = bom("where-used", "--file", p, "--find", "볼트")
    assert code == 0 and out["count"] == 1 and any("보지 않았습니다" in w for w in out["warnings"])

def _csvs(tmp_path, pairs):
    paths = []
    for name, text in pairs:
        p = tmp_path / name
        p.write_text(text, encoding="utf-8")
        paths.append(p)
    return paths


def test_irregular_boms_are_compared_or_said_plainly(env, tmp_path):
    """비표준 BOM에서 틀린 답을 VERIFIED로 내던 것 (Fable 검토 2026-11-14, 0.2.0): 점 표기 Level, 한 판만 다른 열 이름,
    같은 상위 아래 둘이던 품번이 하나가 됨, 품번에 '/', 한 판만 Level을 못 읽음."""
    dot_o, dot_n = _csvs(tmp_path, [("dot_o.csv", "Level,Part Number,Qty\n0,TOP,1\n.1,A,2\n..2,B,4\n"),
                                    ("dot_n.csv", "Level,Part Number,Qty\n0,TOP,1\n.1,A,2\n..2,B,9\n")])
    code, d = bom("diff", "--old", dot_o, "--new", dot_n)
    assert code == 0 and d["verified"] and d["lines"] == ["[바뀜] B (TOP/A 아래): Qty 4 -> 9 (옛 엑셀 행 4, 새 엑셀 행 4)"]
    col_o, col_n = _csvs(tmp_path, [("col_o.csv", "Level,Part Number,Qty\n0,TOP,1\n1,A,2\n2,B,4\n"),
                                    ("col_n.csv", "Level,Part Number,수량\n0,TOP,1\n1,A,2\n2,B,9\n")])
    code, d = bom("diff", "--old", col_o, "--new", col_n)
    assert code == 0 and d["lines"] == ["[바뀜] B (TOP/A 아래): Qty(새 판 수량) 4 -> 9 (옛 엑셀 행 4, 새 엑셀 행 4)"]
    assert any("옛 판 Qty = 새 판 수량" in w for w in d["warnings"])
    dup_o, dup_n = _csvs(tmp_path, [("dup_o.csv", "Level,Part Number,Qty\n0,TOP,1\n1,G,1\n2,X,1\n1,G,1\n2,Y,1\n"),
                                    ("dup_n.csv", "Level,Part Number,Qty\n0,TOP,1\n1,G,1\n2,X,1\n")])
    code, d = bom("diff", "--old", dup_o, "--new", dup_n)
    assert code == 0 and d["counts"]["추가"] == 0
    assert d["lines"] == ["[삭제] G: TOP 아래 (옛 엑셀 행 5)", "[삭제] Y: TOP/G#2 아래 (옛 엑셀 행 6)"]
    sl_o, sl_n = _csvs(tmp_path, [("sl_o.csv", "Level,Part Number,Qty\n0,TOP,1\n1,A/100,2\n1,A,1\n2,100,3\n"),
                                  ("sl_n.csv", "Level,Part Number,Qty\n0,TOP,1\n1,A/100,2\n1,A,1\n2,100,5\n")])
    code, d = bom("diff", "--old", sl_o, "--new", sl_n)                          # 품번의 '/'가 경로와 겹치지 않음
    assert code == 0 and d["lines"] == ["[바뀜] 100 (TOP/A 아래): Qty 3 -> 5 (옛 엑셀 행 5, 새 엑셀 행 5)"]
    bad_o, bad_n = _csvs(tmp_path, [("bad_o.csv", "Level,Part Number,Qty\n0,TOP,1\n1,A,2\n2,B,4\n"),
                                    ("bad_n.csv", "Level,Part Number,Qty\n0,TOP,1\nL1,A,2\nL2,B,4\n")])
    code, d = bom("diff", "--old", bad_o, "--new", bad_n)                        # 못 읽은 행은 첫 줄에, UNVERIFIED
    assert code == 0 and not d["verified"] and d["agent_directive"].startswith("UNVERIFIED:")
    assert d["preview"].startswith("주의: 새 판 3행 중 2행은 Level·품번을 읽지 못해 비교하지 않았습니다")
    code, w = bom("where-used", "--file", bad_n, "--find", "B")
    assert code == 0 and not w["verified"] and w["preview"].startswith("주의: BOM 3행 중 2행은")


def test_history_collects_the_diffs_of_connected_versions(env, tmp_path):
    """변경 이력 (0.2.0): diff마다 결과 폴더에 기록을 남기고, history는 이어진 판들의 기록을 시간순으로 모아 csv 하나로."""
    import csv as _csv
    import time as _t
    from bk_helpers import make_samples
    files = make_samples(tmp_path / "mbom")
    v1, v2, mod = files / "MBOM_DR200.xlsx", files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2_mod.xlsx"
    code, h = bom("history", "--file", mod)
    assert code == 0 and h["count"] == 0 and "비교 기록이 없습니다" in h["preview"] and "0건" in h["agent_directive"]
    bom("diff", "--old", v1, "--new", v2)
    bom("diff", "--old", v2, "--new", mod)
    bom("diff", "--old", v2, "--new", mod)                                      # 같은 비교 다시: 마지막 것만
    for start in (v1, v2, mod):                                                  # 어느 판에서 물어도 같은 이력
        code, h = bom("history", "--file", start)
        assert code == 0 and h["verified"] and h["count"] == 2, h
        assert [line.split(": ", 1)[0].split(" ", 3)[3] for line in h["lines"]] == [
            "MBOM_DR200.xlsx -> MBOM_DR200_v2.xlsx", "MBOM_DR200_v2.xlsx -> MBOM_DR200_v2_mod.xlsx"]
    assert h["lines"][1].endswith("변경 대체 후보 1건, 추가 2건, 삭제 2건, 바뀐 칸 8곳 (7개 품목).")
    with open(h["output_file"], encoding="utf-8-sig") as f:
        rows = list(_csv.DictReader(f))
    assert {r["새 판"] for r in rows} == {"MBOM_DR200_v2.xlsx", "MBOM_DR200_v2_mod.xlsx"} and "/" in rows[0]["경로"]
    _t.sleep(1.1)
    v2.write_bytes(v2.read_bytes() + b"")
    os.utime(v2, (_t.time(), _t.time()))                                         # 비교 뒤 원본이 바뀜
    code, h = bom("history", "--file", mod)
    assert code == 0 and any("MBOM_DR200_v2.xlsx이(가) 바뀌었습니다" in w for w in h["warnings"])
