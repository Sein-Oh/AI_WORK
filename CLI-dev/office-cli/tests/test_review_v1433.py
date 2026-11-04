"""1.4.33: En-Genius 스킬로 실행하면(트레이가 ENGENIUS_OUT을 넣고, 스킬 파일을 실행마다 임시 폴더에 풀었다가 지움)
config.json, data, cache를 도구 폴더가 아니라 고정 폴더(<USERPROFILE>/Documents/En-Genius/CLI-data/office-cli)에 둔다.
OFFICE_CLI_CONFIG로 config.json 위치를 정할 수 있다.
1.4.35: 고정 폴더는 도구 파일이 트레이의 임시 폴더(engenius-skill)에서 실행될 때만 쓴다. 배포 스킬은 C:\\CLI의 도구를 실행하므로
ENGENIUS_OUT이 있어도 도구 폴더(OpenCode와 같은 data, cache)를 쓴다."""
import shutil
import json
import os
import subprocess
import sys

TOOL_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))),
                        "CLI", "office-cli")
PROBE = ("import sys, json\n"
         "sys.path.insert(0, sys.argv[1])\n"
         "import _runtime as rt\n"
         "print(json.dumps({'config': rt.CONFIG_FILE, 'data': rt.get_data_dir(), 'cache': rt.get_cache_dir()}))\n")


def _env(tmp_path, **extra):
    env = {k: v for k, v in os.environ.items() if not k.startswith("OFFICE_CLI_") and k != "ENGENIUS_OUT"}
    env["USERPROFILE"] = str(tmp_path / "home")
    env.update(extra)
    return env


def _unpacked(tmp_path):
    """트레이처럼 도구를 %TEMP%\\engenius-skill\\<id>\\에 푼 자리."""
    d = tmp_path / "engenius-skill" / "r1"
    if not d.exists():
        shutil.copytree(TOOL_DIR, d, ignore=shutil.ignore_patterns("data", "cache", "config.json", "__pycache__"))
    return str(d)


def _probe(tmp_path, tool_dir=TOOL_DIR, **extra):
    out = subprocess.run([sys.executable, "-c", PROBE, tool_dir], capture_output=True, timeout=60,
                         env=_env(tmp_path, **extra))
    return json.loads(out.stdout.decode("utf-8").strip().splitlines()[-1])


def test_engenius_temp_run_uses_fixed_folder(tmp_path):
    got = _probe(tmp_path, _unpacked(tmp_path), ENGENIUS_OUT=str(tmp_path / "out"))
    home = os.path.join(str(tmp_path / "home"), "Documents", "En-Genius", "CLI-data", "office-cli")
    assert got == {"config": os.path.join(home, "config.json"), "data": os.path.join(home, "data"),
                   "cache": os.path.join(home, "cache")}


def test_engenius_run_from_cli_folder_shares_tool_folder(tmp_path):
    got = _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out"))
    assert got["config"] == os.path.join(TOOL_DIR, "config.json")
    assert got["data"] == os.path.join(TOOL_DIR, "data") and got["cache"] == os.path.join(TOOL_DIR, "cache")


def test_normal_run_keeps_tool_folder(tmp_path):
    got = _probe(tmp_path)
    assert got["config"] == os.path.join(TOOL_DIR, "config.json")
    assert got["data"] == os.path.join(TOOL_DIR, "data") and got["cache"] == os.path.join(TOOL_DIR, "cache")


def test_env_settings_win_over_engenius(tmp_path):
    got = _probe(tmp_path, _unpacked(tmp_path), ENGENIUS_OUT=str(tmp_path / "out"), OFFICE_CLI_CONFIG=str(tmp_path / "c.json"),
                 OFFICE_CLI_DATA_DIR=str(tmp_path / "d"), OFFICE_CLI_CACHE_DIR=str(tmp_path / "k"))
    assert got == {"config": str(tmp_path / "c.json"), "data": str(tmp_path / "d"), "cache": str(tmp_path / "k")}


def test_engenius_error_log_lands_in_fixed_folder(tmp_path):
    proc = subprocess.run([sys.executable, os.path.join(_unpacked(tmp_path), "excel.py"), "--bogus"], capture_output=True,
                          timeout=60, env=_env(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")))
    assert proc.returncode == 2
    log = tmp_path / "home" / "Documents" / "En-Genius" / "CLI-data" / "office-cli" / "data" / "logs" / "errors.txt"
    assert log.is_file() and "--bogus" in log.read_text(encoding="utf-8")
