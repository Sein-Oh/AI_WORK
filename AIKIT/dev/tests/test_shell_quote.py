# -*- coding: utf-8 -*-
"""도구가 내는 명령의 인자 따옴표(quote_arg, office-cli 1.4.76 / ml-kit 2.9.68)를 셸 시뮬레이터(shell_sim.py)로 확인한다.

까다로운 값마다 quote_arg 결과로 `python <스크립트> <값>`을 만들고, 그 셸(Windows PowerShell 5.1, PowerShell 7.3+,
cmd, bash)이 실행할 때 파이썬이 받는 argv가 원래 값과 같은지 본다. 두 도구의 quote_arg는 복사본이라 결과도 같아야 한다.
"""
import importlib.util
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import shell_sim  # noqa: E402
from shell_sim import ShellError, run_shell  # noqa: E402

sys.path.pop(0)


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


OFFICE = _load("office_runtime_for_quote", os.path.join(REPO, "office-cli", "_runtime.py"))
_grade = _load("grade_opencode_for_quote", os.path.join(REPO, "dev", "grade_opencode.py"))
try:
    ML = _load("ml_runtime_for_quote", os.path.join(REPO, "ml-kit", "_ml_runtime.py"))
except ImportError:                    # ml-kit 라이브러리가 없는 PC: office-cli만 확인
    ML = None
TOOLS = [OFFICE] + ([ML] if ML else [])

SCRIPT = r"C:\AIKIT\office-cli\excel.py"
VALUES = [
    "plain", "a b", "  두 칸  ", "\t탭\t", "", "-", "--", "-x", "--cells", "-5", "-1.5", "-a.b", "-Dfoo.bar=1",
    "C:\\a b\\", "C:\\a b\\\\", "\\\\server\\share\\x y\\", "C:\\끝\\", "C:\\AIKIT\\office-cli\\excel.py",
    'say "hi"', '"', '\\"', '\\\\"', 'a\\"b c', 'x"', '"앞뒤"', "it's", "'", "''", "\u2018둥근\u2019 따옴표", "\u201c큰\u201d",
    "$x", "${x}", "$env:PATH", "a$b c", "`", "a`nb", "100`$",
    "{{업체명}}", "{a}", "@a", "a@b.com", "#x", "a#b", "x;y", "C6=ft; C7=\u2103", "(1)", "f(x)", "[0]", "a[1]",
    "%PATH%", "100%", "%x% %USERNAME%", "a&b", "a|b", "<>", "a>b.txt", "^", "a^b", "!x", "S!Z200000",
    "한글", "전각\uff21\uff11　공백", "\u3000", "\xa0nbsp", "007", "0x10", "1e3", "1.50", ".5", "+5", "2-8", "5,9", "a,b", ",a", "a,",
    "a,,b", "=SUM(A1:A3)", "=$A$1*2", '=IF(A1="a;b",1,2)', '부서 = "영업"', "비고 = '불량 포함 시료'", "반응온도(°C)=240,소비전력(kW)=65",
    '[{"a": 1}, {"b": "x y"}]', "a\nb", "*.xlsx", "~", "a~b", "--%", "&&", "a||b",
]


def _cmd(tool, value, shell):
    return f"python {tool.quote_arg(SCRIPT, shell)} {tool.quote_arg(value, shell)}"


def test_value_table_is_big_enough():
    assert len(VALUES) >= 40 and len(set(VALUES)) == len(VALUES)


@pytest.mark.parametrize("tool", TOOLS, ids=lambda m: m.__name__.split("_")[0])
@pytest.mark.parametrize("shell", ["powershell", "pwsh"])
@pytest.mark.parametrize("value", VALUES)
def test_powershell_passes_value_exactly(tool, shell, value):
    """PowerShell 두 종류는 모든 값이 정확히 그대로 (셸 문법 단계 + 네이티브 전달 단계)."""
    cmd = _cmd(tool, value, shell)
    assert run_shell(cmd, shell) == ["python", SCRIPT, value], cmd


@pytest.mark.parametrize("tool", TOOLS, ids=lambda m: m.__name__.split("_")[0])
@pytest.mark.parametrize("shell", ["cmd", "bash"])
@pytest.mark.parametrize("value", VALUES)
def test_cmd_and_bash_best_effort(tool, shell, value):
    """cmd·bash는 최선 노력: 줄바꿈이 든 값(cmd에서 명령이 나뉨)만 빼고 그대로."""
    cmd = _cmd(tool, value, shell)
    if shell == "cmd" and "\n" in value:
        with pytest.raises(ShellError):
            run_shell(cmd, shell)
        return
    assert run_shell(cmd, shell) == ["python", SCRIPT, value], cmd


@pytest.mark.parametrize("shell", shell_sim.SHELLS)
def test_both_tools_quote_the_same(shell):
    """quote_arg는 도구마다 복사본 (도구 간 import 금지): 두 복사본이 같은 결과를 내야 한다."""
    if ML is None:
        pytest.skip("ml-kit 라이브러리 없음")
    assert [OFFICE.quote_arg(v, shell) for v in VALUES] == [ML.quote_arg(v, shell) for v in VALUES]


def test_shapes():
    """단순 토큰은 따옴표 없이, 나머지는 셸마다 정해진 모양."""
    q = OFFICE.quote_arg
    assert q("C:\\AIKIT\\office-cli\\excel.py", "powershell") == "C:\\AIKIT\\office-cli\\excel.py"
    assert q("반응온도(°C)", "pwsh") == "'반응온도(°C)'" and q("--cells", "powershell") == "--cells" and q("12", "pwsh") == "12"
    assert q("it's", "pwsh") == "'it''s'" and q("$x", "pwsh") == "'$x'"
    assert q('a "b"', "pwsh") == "'a \"b\"'"                         # 7.3+는 스스로 이스케이프: 미리 바꾸지 않음
    assert q('a "b"', "powershell") == "'a \\\"b\\\"'"               # 5.1은 \" 로 미리
    assert q("C:\\a b\\", "powershell") == "'C:\\a b\\\\'" and q("C:\\a b\\", "pwsh") == "'C:\\a b\\'"
    assert q("", "powershell") == "'\"\"'" and q("", "pwsh") == "''"   # 5.1은 빈 인자를 버림
    assert q("100%", "cmd") == '"100"^%""' and q("a b", "bash") == "'a b'" and q("C:\\x", "bash") == "'C:\\x'"


def test_old_double_quote_style_breaks_in_windows_powershell():
    """옛 quote_arg("...", 안의 "는 \\")가 실제 셸에서 깨지던 모양을 시뮬레이터가 재현한다 (시뮬레이터 자체 확인)."""
    assert run_shell('python x.py "C:\\a b\\"', "powershell")[2] != "C:\\a b\\"       # 끝 백슬래시
    with pytest.raises(ShellError):
        run_shell('python x.py "부서 = \\"영업\\""', "powershell")                # PowerShell에서 \"는 이스케이프가 아님
    with pytest.raises(ShellError):
        run_shell('python x.py "=$A$1*2"', "powershell")                            # 큰따옴표 안의 $ 확장
    with pytest.raises(ShellError):
        run_shell("python x.py --cells C6=a;C7=b", "pwsh")                         # 따옴표 밖의 ;
    assert run_shell('python x.py ""', "powershell") == ["python", "x.py"]          # 5.1은 빈 값을 버림
    assert run_shell('python x.py "%X%"', "cmd")[2] != "%X%"                        # cmd는 따옴표 안의 %X%도 확장


def test_shell_detection_follows_opencode(monkeypatch):
    """환경변수 강제 > pwsh가 PATH에 있으면 pwsh > Windows PowerShell 5.1 (OpenCode 셸 도구와 같은 순서)."""
    tools = [(OFFICE, "OFFICE_CLI_SHELL")] + ([(ML, "ML_KIT_SHELL")] if ML else [])
    for tool, var in tools:
        monkeypatch.delenv(var, raising=False)
        tool._FOUND_SHELL.clear()
        monkeypatch.setattr("shutil.which", lambda name: "C:\\pwsh\\pwsh.exe" if name == "pwsh" else None)
        assert tool.command_shell() == "pwsh" and tool.quote_arg('a "b"') == "'a \"b\"'"
        tool._FOUND_SHELL.clear()
        monkeypatch.setattr("shutil.which", lambda name: None)
        assert tool.command_shell() == "powershell" and tool.quote_arg('a "b"') == "'a \\\"b\\\"'"
        for forced in shell_sim.SHELLS:
            monkeypatch.setenv(var, forced.upper())
            assert tool.command_shell() == forced
        monkeypatch.setenv(var, "zsh")                                              # 모르는 값은 찾은 셸
        assert tool.command_shell() == "powershell"
        tool._FOUND_SHELL.clear()


def test_grader_splits_single_quoted_commands():
    """채점기 command_segments는 작은따옴표 안의 ; | & 와 ''를 한 명령으로 둔다."""
    cmds = ["python C:\\AIKIT\\office-cli\\excel.py write --cells 'C6=ft; C7=\u2103' --dry-run",
            "python C:\\AIKIT\\ml-kit\\ml_dataset.py exclude --file 'C:\\a b\\d.xlsx' --where '비고 = ''a;b''|x > 1'",
            "python C:\\AIKIT\\office-cli\\word.py fill --file 'C:\\값 $x (1)\\it''s.csv' --where '부서 = \\\"영업 & 기획\\\"'"]
    for cmd in cmds:
        assert _grade.command_segments(cmd) == [cmd]
    two = "$env:PYTHONIOENCODING='utf-8'; python a.py read --target 'x; y' ; python b.py"
    assert _grade.command_segments(_grade.ENC_PREFIX.sub("", two)) == ["python a.py read --target 'x; y'", "python b.py"]
