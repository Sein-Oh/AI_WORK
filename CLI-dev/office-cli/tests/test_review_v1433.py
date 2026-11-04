"""1.4.33: En-Genius 스킬로 실행하면(트레이가 ENGENIUS_OUT을 넣고, 스킬 파일을 실행마다 임시 폴더에 풀었다가 지움)
config.json, data, cache를 도구 폴더가 아니라 고정 폴더(<USERPROFILE>/Documents/En-Genius/CLI-data/office-cli)에 둔다.
OFFICE_CLI_CONFIG로 config.json 위치를 정할 수 있다."""
import json
import os
import subprocess
import sys

TOOL_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))),
                        "CLI", "office-cli")
PROBE = ("import sys, json\n"
         f"sys.path.insert(0, {TOOL_DIR!r})\n"
         "import _runtime as rt\n"
         "print(json.dumps({'config': rt.CONFIG_FILE, 'data': rt.get_data_dir(), 'cache': rt.get_cache_dir()}))\n")


def _env(tmp_path, **extra):
    env = {k: v for k, v in os.environ.items() if not k.startswith("OFFICE_CLI_") and k != "ENGENIUS_OUT"}
    env["USERPROFILE"] = str(tmp_path / "home")
    env.update(extra)
    return env


def _probe(tmp_path, **extra):
    out = subprocess.run([sys.executable, "-c", PROBE], capture_output=True, timeout=60, env=_env(tmp_path, **extra))
    return json.loads(out.stdout.decode("utf-8").strip().splitlines()[-1])


def test_engenius_run_uses_fixed_folder(tmp_path):
    got = _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out"))
    home = os.path.join(str(tmp_path / "home"), "Documents", "En-Genius", "CLI-data", "office-cli")
    assert got == {"config": os.path.join(home, "config.json"), "data": os.path.join(home, "data"),
                   "cache": os.path.join(home, "cache")}


def test_normal_run_keeps_tool_folder(tmp_path):
    got = _probe(tmp_path)
    assert got["config"] == os.path.join(TOOL_DIR, "config.json")
    assert got["data"] == os.path.join(TOOL_DIR, "data") and got["cache"] == os.path.join(TOOL_DIR, "cache")


def test_env_settings_win_over_engenius(tmp_path):
    got = _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out"), OFFICE_CLI_CONFIG=str(tmp_path / "c.json"),
                 OFFICE_CLI_DATA_DIR=str(tmp_path / "d"), OFFICE_CLI_CACHE_DIR=str(tmp_path / "k"))
    assert got == {"config": str(tmp_path / "c.json"), "data": str(tmp_path / "d"), "cache": str(tmp_path / "k")}


def test_engenius_error_log_lands_in_fixed_folder(tmp_path):
    proc = subprocess.run([sys.executable, os.path.join(TOOL_DIR, "excel.py"), "--bogus"], capture_output=True,
                          timeout=60, env=_env(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")))
    assert proc.returncode == 2
    log = tmp_path / "home" / "Documents" / "En-Genius" / "CLI-data" / "office-cli" / "data" / "logs" / "errors.txt"
    assert log.is_file() and "--bogus" in log.read_text(encoding="utf-8")
