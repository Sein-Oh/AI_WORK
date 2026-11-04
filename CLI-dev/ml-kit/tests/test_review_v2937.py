"""2.9.37: run as an En-Genius skill (the tray sets ENGENIUS_OUT and deletes the unpacked skill folder after each run),
the tool data folder (error log) is <USERPROFILE>/Documents/En-Genius/CLI-data/ml-kit/data, not <tool folder>/data.
ML_KIT_DATA_DIR still wins."""
import json
import os
import subprocess
import sys

TOOL_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))),
                        "CLI", "ml-kit")
PROBE = ("import sys\n"
         f"sys.path.insert(0, {TOOL_DIR!r})\n"
         "import ml_predict\n"
         "print(ml_predict.get_data_dir())\n")


def _env(tmp_path, **extra):
    env = {k: v for k, v in os.environ.items() if k not in ("ML_KIT_DATA_DIR", "ENGENIUS_OUT")}
    env["USERPROFILE"] = str(tmp_path / "home")
    env.update(extra)
    return env


def _probe(tmp_path, **extra):
    out = subprocess.run([sys.executable, "-c", PROBE], capture_output=True, timeout=120, env=_env(tmp_path, **extra))
    return out.stdout.decode("utf-8").strip().splitlines()[-1]


def test_engenius_run_uses_fixed_folder(tmp_path):
    assert _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")) == os.path.join(
        str(tmp_path / "home"), "Documents", "En-Genius", "CLI-data", "ml-kit", "data")


def test_normal_run_keeps_tool_folder(tmp_path):
    assert _probe(tmp_path) == os.path.join(TOOL_DIR, "data")


def test_env_wins_over_engenius(tmp_path):
    assert _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out"), ML_KIT_DATA_DIR=str(tmp_path / "d")) == str(tmp_path / "d")


def test_engenius_error_log_lands_in_fixed_folder(tmp_path):
    proc = subprocess.run([sys.executable, os.path.join(TOOL_DIR, "ml_predict.py"), "--bogus"], capture_output=True,
                          timeout=120, env=_env(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")))
    assert proc.returncode == 2 and json.loads(proc.stdout.decode("utf-8").strip().splitlines()[-1])["status"] == "error"
    assert (tmp_path / "home" / "Documents" / "En-Genius" / "CLI-data" / "ml-kit" / "data" / "logs" / "errors.txt").is_file()
