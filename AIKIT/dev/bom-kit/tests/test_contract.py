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
    for skill, cmds in (("mbom-diff", ("diff",)), ("where-used", ("where-used",))):
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
    assert ok(env, "show")["work"] == mine and ok(env, "show")["table"] == "mine"
