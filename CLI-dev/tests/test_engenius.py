"""En-Genius 스킬 zip(docs/engenius/<tool>.zip)과 실행 입구(run.py) 테스트.

- 저장소의 zip이 CLI-dev/engenius/build.py로 새로 만든 것과 같은지 (SKILL.md·run.py·버전을 고치면 build.py를 다시 실행)
- zip에는 SKILL.md와 run.py만: run.py가 C:\\CLI\\<tool>의 도구를 실행한다 (OpenCode와 같은 폴더)
- En-Genius 가져오기 규칙: SKILL.md 위치, 이름(영문·숫자·한글·-·_ 1~64자), 경로(4칸, 합계 120자, 한 칸 64자),
  본문 30,000자, SKILL.md 512KB, 파일 합계는 실행 프레임 상한(약 11MB) 안
- run.py: 명령 한 줄 나누기, python·경로 앞부분 무시, @input 자리 내용, 오류는 1줄 JSON(exit 2), 도구 결과는 그대로
"""
import importlib.util
import json
import os
import re
import subprocess
import sys
import zipfile

import pytest

DEV = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.dirname(DEV)
BUILD = os.path.join(DEV, "engenius", "build.py")
TOOLS = ("ml-kit", "office-cli")


def _load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


build = _load(BUILD, "engenius_build")


@pytest.mark.parametrize("tool", TOOLS)
def test_committed_zip_is_current(tool):
    with open(os.path.join(ROOT, "docs", "engenius", f"{tool}.zip"), "rb") as f:
        committed = f.read()
    assert committed == build.zip_bytes(tool), "docs/engenius zip이 최신이 아닙니다: python CLI-dev/engenius/build.py"


@pytest.mark.parametrize("tool", TOOLS)
def test_zip_follows_import_rules(tool):
    z = zipfile.ZipFile(os.path.join(ROOT, "docs", "engenius", f"{tool}.zip"))
    names = [i.filename for i in z.infolist()]
    assert sorted(names) == [f"{tool}/SKILL.md", f"{tool}/run.py"]                # 도구 코드는 C:\\CLI에
    for n in names:
        rel = n.split("/", 1)[1]
        parts = rel.split("/")
        assert len(parts) <= 4 and len(rel) <= 120 and all(0 < len(p) <= 64 for p in parts), rel
    assert sum(i.file_size for i in z.infolist()) < 11 * 1024 * 1024
    md = z.read(f"{tool}/SKILL.md")
    assert len(md) < 512 * 1024
    text = md.decode("utf-8")
    lines = text.splitlines()
    assert lines[0] == "---"
    end = lines.index("---", 1)
    meta = dict(re.match(r"^([A-Za-z0-9_-]+):[ \t]*(.*)$", ln).groups() for ln in lines[1:end])
    assert re.fullmatch(r"[0-9A-Za-z가-힣_-]{1,64}", meta["name"]) and meta["name"] == tool
    desc = json.loads(meta["description"])
    assert 0 < len(desc) <= 200                                   # 도구 설명에는 200자까지만 실림
    body = "\n".join(lines[end + 1:]).strip()
    assert body and len(body) <= 30000
    assert "run_skill_script" in body and "PowerShell" not in body and "파일 쓰기 도구" not in body
    for ch in text:                                               # CP949로 표현 가능한 글자만 (문자 규칙)
        ch.encode("cp949")


# ---------------------------------------------------------------------------------------------------------------
# run.py
# ---------------------------------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def unpacked(tmp_path_factory):
    """트레이처럼 zip을 임시 폴더에 푼다."""
    out = {}
    for tool in TOOLS:
        d = tmp_path_factory.mktemp(f"engenius-skill-{tool}")
        with zipfile.ZipFile(os.path.join(ROOT, "docs", "engenius", f"{tool}.zip")) as z:
            z.extractall(d)
        out[tool] = os.path.join(str(d), tool)
    return out


def run(unpacked, tool, args, tmp_path, py_flags=(), **extra):
    """트레이처럼 실행: 풀린 스킬 폴더가 cwd, C:\\CLI 대신 저장소 CLI 폴더(ENGENIUS_CLI_DIR)."""
    env = {k: v for k, v in os.environ.items()
           if not k.startswith(("OFFICE_CLI_", "ML_KIT_", "ENGENIUS_"))}
    env.update(ENGENIUS_OUT=str(tmp_path / "out"), USERPROFILE=str(tmp_path / "home"), PYTHONUTF8="1",
               ENGENIUS_CLI_DIR=os.path.join(ROOT, "CLI"), OFFICE_CLI_DATA_DIR=str(tmp_path / "odata"),
               OFFICE_CLI_CACHE_DIR=str(tmp_path / "ocache"), ML_KIT_DATA_DIR=str(tmp_path / "mdata"))
    env.update(extra)
    proc = subprocess.run([sys.executable, *py_flags, "run.py", *args], cwd=unpacked[tool], capture_output=True,
                          timeout=300, env=env)
    lines = [ln for ln in proc.stdout.decode("utf-8").splitlines() if ln.strip()]
    assert len(lines) == 1, (proc.stdout, proc.stderr.decode("utf-8", "replace")[-2000:])
    return proc.returncode, json.loads(lines[0])


runpy_mod = _load(os.path.join(DEV, "engenius", "office-cli", "run.py"), "engenius_run_office")


@pytest.mark.parametrize("tool", TOOLS)
def test_each_tool_has_its_own_run_py(tool):
    mod = _load(os.path.join(DEV, "engenius", tool, "run.py"), f"engenius_run_{tool}")
    assert mod.TOOL == tool and mod.OTHER_TOOL == next(t for t in TOOLS if t != tool)
    assert set(mod.SCRIPTS) == {f for f in os.listdir(os.path.join(ROOT, "CLI", tool))
                                if f.endswith(".py") and not f.startswith("_")}
    z = zipfile.ZipFile(os.path.join(ROOT, "docs", "engenius", f"{tool}.zip"))
    packed = z.read(f"{tool}/run.py").decode("utf-8")
    assert packed == build.run_py(tool)
    m = re.search(r'^SKILL_VERSION = "([^"]+)"', packed, re.M)
    with open(os.path.join(ROOT, "CLI", tool, mod.VERSION_FILE), encoding="utf-8") as f:
        assert m and m.group(1) == re.search(r'^VERSION\s*=\s*"([^"]+)"', f.read(), re.M).group(1)
    with open(os.path.join(ROOT, "CLI", tool, "SKILL.md"), encoding="utf-8") as f:
        digest = mod.skill_hash(f.read())
    assert re.search(r'^SKILL_HASH = "([0-9a-f]{16})"', packed, re.M).group(1) == digest


def test_run_py_files_differ_only_in_tool_settings():
    texts = [open(os.path.join(DEV, "engenius", t, "run.py"), encoding="utf-8").read() for t in TOOLS]
    strip = [re.sub(r'(?m)^(TOOL|SCRIPTS|VERSION_FILE|NEEDS|OTHER_TOOL|OTHER_SCRIPTS|EXAMPLE) = .*$', "",
                    t.split('"""', 2)[2]) for t in texts]                        # 설명(docstring)과 도구 설정 줄만 다름
    assert strip[0] == strip[1], "두 run.py의 공통 부분이 다릅니다 (한쪽만 고쳤는지 확인)"


@pytest.mark.parametrize("line,expected", [
    ('excel.py read --range A1:D20', ["excel.py", "read", "--range", "A1:D20"]),
    ('python "C:\\CLI\\ml-kit\\ml_train.py" -f "C:\\데이터\\공정 1.xlsx" -y 수율',
     ["python", "C:\\CLI\\ml-kit\\ml_train.py", "-f", "C:\\데이터\\공정 1.xlsx", "-y", "수율"]),
    ('reader.py index --folder "C:\\문서 폴더\\"', ["reader.py", "index", "--folder", "C:\\문서 폴더\\"]),
    ('excel.py write --formula "=IF(A1=\\"x\\",1,0)"', ["excel.py", "write", "--formula", '=IF(A1="x",1,0)']),
    ("excel.py write --formula '=$A$1*2'", ["excel.py", "write", "--formula", "=$A$1*2"]),
    ('word.py replace --find "" --replace x', ["word.py", "replace", "--find", "", "--replace", "x"]),
])
def test_split_command(line, expected):
    assert runpy_mod.split_command(line) == expected


def test_split_command_round_trips_quote_arg():
    sys.path.insert(0, os.path.join(ROOT, "CLI", "office-cli"))
    import _runtime as rt
    values = ["C:\\a b\\c.xlsx", "수율,강도", 'x"y', "=SUM(A1:A3)", "100%", "a;b", "plain"]
    assert runpy_mod.split_command(" ".join(rt.quote_arg(v) for v in values)) == values


def test_help_is_json(unpacked, tmp_path):
    code, res = run(unpacked, "office-cli", [], tmp_path)
    assert code == 0 and res["status"] == "ok" and "excel.py" in res["scripts"] and "run.py" not in res["scripts"]


def test_full_command_line_runs_tool(unpacked, tmp_path):
    csv = tmp_path / "표 1.csv"
    csv.write_text("x,y\n1,2\n3,4\n", encoding="utf-8")
    code, res = run(unpacked, "office-cli", [f'python "C:\\CLI\\office-cli\\excel.py" read --target "{csv}"'], tmp_path)
    assert code == 0 and res["status"] == "ok" and res["agent_directive"].startswith("VERIFIED"), res


def test_split_args_also_run(unpacked, tmp_path):
    code, res = run(unpacked, "ml-kit", ["ml_train.py", "-h"], tmp_path)
    assert code == 0 and res["status"] == "ok"


def test_input_placeholder_becomes_file(unpacked, tmp_path):
    code, res = run(unpacked, "office-cli", ["excel.py read --target @input1.csv", "a,b\n1,2\n5,6\n"], tmp_path)
    assert code == 0 and res["status"] == "ok" and res["rows"] == 3, res


def test_tool_error_keeps_tool_json(unpacked, tmp_path):
    code, res = run(unpacked, "office-cli", ["excel.py --bogus"], tmp_path)
    assert code == 2 and res["status"] == "error" and res["error_type"] != "setup"
    assert (tmp_path / "odata" / "logs" / "errors.txt").is_file()          # 도구의 오류 로그 (C:\\CLI 쪽 설정 그대로)


def test_python_without_libraries_falls_back_to_winpython(unpacked, tmp_path):
    """트레이 내장 파이썬처럼 도구 라이브러리가 없으면(-S: site-packages 없음) WinPython(ENGENIUS_PYTHON)으로 실행."""
    code, res = run(unpacked, "ml-kit", ["ml_train.py -h"], tmp_path, py_flags=("-S",), ENGENIUS_PYTHON=sys.executable)
    assert code == 0 and res["status"] == "ok", res
    code, res = run(unpacked, "ml-kit", ["ml_train.py -h"], tmp_path, py_flags=("-S",),
                    ENGENIUS_PYTHON=str(tmp_path / "없음" / "python.exe"))
    assert code == 2 and res["error_type"] == "setup" and "WinPython" in res["message"]


def test_missing_cli_folder_is_setup_error(unpacked, tmp_path):
    code, res = run(unpacked, "office-cli", ["excel.py read"], tmp_path, ENGENIUS_CLI_DIR=str(tmp_path / "noCLI"))
    assert code == 2 and res["error_type"] == "setup" and "C:\\CLI" in res["message"]


def test_changed_skill_text_asks_to_reimport_zip(unpacked, tmp_path):
    other = tmp_path / "skill" / "office-cli"
    other.mkdir(parents=True)
    text = open(os.path.join(unpacked["office-cli"], "run.py"), encoding="utf-8").read()
    (other / "run.py").write_text(re.sub(r'^SKILL_HASH = "[^"]+"', 'SKILL_HASH = "0000000000000000"', text, flags=re.M),
                                  encoding="utf-8")
    code, res = run({"office-cli": str(other)}, "office-cli", ["excel.py read"], tmp_path)
    assert code == 2 and res["error_type"] == "setup" and "다시 가져온" in res["message"]


def test_tool_only_update_needs_no_new_zip(unpacked, tmp_path):
    """도구만 고친 업데이트: C:\\CLI의 버전 줄(SKILL.md metadata, VERSION)만 바뀌면 zip은 그대로 실행된다."""
    import shutil
    cli = tmp_path / "CLI"
    shutil.copytree(os.path.join(ROOT, "CLI", "office-cli"), cli / "office-cli",
                    ignore=shutil.ignore_patterns("data", "cache", "config.json", "__pycache__"))
    for name, pat, new in (("SKILL.md", r'version: "[^"]+"', 'version: "9.9.9"'),
                           ("_runtime.py", r'^VERSION = "[^"]+"', 'VERSION = "9.9.9"')):
        f = cli / "office-cli" / name
        f.write_text(re.sub(pat, new, f.read_text(encoding="utf-8"), count=1, flags=re.M), encoding="utf-8")
    code, res = run(unpacked, "office-cli", [], tmp_path, ENGENIUS_CLI_DIR=str(cli))
    assert code == 0 and res["tool_version"] == "9.9.9" and res["skill_current"] is True
    csv = tmp_path / "t.csv"
    csv.write_text("a,b\n1,2\n", encoding="utf-8")
    code, res = run(unpacked, "office-cli", [f'excel.py read --target "{csv}"'], tmp_path, ENGENIUS_CLI_DIR=str(cli))
    assert code == 0 and res["status"] == "ok", res


@pytest.mark.parametrize("tool,args,needle", [
    ("ml-kit", ["excel.py read --target a.xlsx"], 'name="office-cli"'),
    ("office-cli", ["python ml_predict.py -m x"], 'name="ml-kit"'),
    ("office-cli", ["notepad.exe a.txt"], "쓸 수 있는 스크립트"),
    ("office-cli", ['excel.py read --target "a.xlsx'], "큰따옴표"),
    ("office-cli", ["word.py write --file @input2.md", "only one"], "@input2"),
    ("office-cli", ["word.py write --file @input1.md", "a", "b"], "쓰이지 않은 칸"),
    ("office-cli", ["python"], "실행할 스크립트"),
])
def test_wrapper_errors_are_one_json_line(unpacked, tmp_path, tool, args, needle):
    code, res = run(unpacked, tool, args, tmp_path)
    assert code == 2 and res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:")
    assert needle in res["message"], res


def test_agent_xml_is_current_and_well_formed():
    import xml.etree.ElementTree as ET
    with open(build.AGENT_XML, encoding="utf-8") as f:
        text = f.read()
    assert text == build.agent_xml(), "docs/engenius 에이전트 XML이 최신이 아닙니다: python CLI-dev/engenius/build.py"
    root = ET.fromstring(text.encode("utf-8"))
    assert root.tag == "EnGeniusAgent" and root.get("format") == "1"
    assert root.findtext("Name") == build.AGENT_NAME and len(root.findtext("Prompt")) <= 8000   # 역할 프롬프트 상한
    assert root.find("Servers") is not None and len(root.find("Servers")) == 0               # 내장 도구 서버 없음
    assert [t.get("name") for t in root.find("Tools")] == ["run_skill_script"]               # 빈 Tools는 400
    assert root.find("Skill") is None and root.find("Provider") is None
    text.encode("cp949")
