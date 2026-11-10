# -*- coding: utf-8 -*-
"""CLI 계약: stdout 1줄 JSON, exit 코드(0/1/2/130), 도움말, 오류 로그, 데이터 폴더, 학습 -> 예측 -> 보고서 기본 흐름.

에이전트(경량 LLM)가 기대는 출력 형식을 고정한다. 성공은 exit 0 + "status": "ok", 실패는 exit != 0 +
"status": "error" + agent_directive(STOP_RETRY/ERROR).
"""
import ast
import contextlib
import io
import json
import os
import shutil
import subprocess
import sys

import numpy as np
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_sample import make_sample  # noqa: E402
from mlkit_helpers import shell_sim, split_command, FEATURES, HAS_DOCX, INLINE, KIT, TARGET, recommended_command, run_fail, run_raw, run_tool  # noqa: E402
import _ml_runtime as rt  # noqa: E402


# ---- 학습 -> 예측 -> 보고서 기본 흐름 (make_sample 데이터, 모듈에서 한 번 학습) --------


@pytest.fixture(scope="module")
def trained(tmp_path_factory):
    work = tmp_path_factory.mktemp("mlkit")
    data = make_sample(str(work / "sample.csv"))
    payload = run_tool(
        "ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET,
        "--desc", "smoke test", "--tag", "baseline",
        "--eval", "온도전력비 = 반응온도(°C) / (소비전력(kW) + 1e-6)",
    )
    return work, data, payload


def test_train_creates_version(trained):
    work, _, payload = trained
    data_root = work / "sample_ml_data"
    assert (data_root / "history.jsonl").is_file()
    assert (data_root / "history.json").is_file()
    assert (data_root / "v1").is_dir()
    oof = (data_root / "v1" / "oof_생산효율(%).csv")
    assert oof.is_file() and oof.read_text(encoding="utf-8-sig").splitlines()[0] == "data_row,actual,predicted,residual"
    rec = json.loads((data_root / "history.jsonl").read_text(encoding="utf-8").splitlines()[0])
    assert rec["data_profile"]["rows_loaded"] == 120 and "feature_missing" in rec["data_profile"]
    assert not any(p.name.startswith(".tmp_") for p in data_root.iterdir())
    assert payload["agent_directive"].startswith("VERIFIED")
    assert payload["goal_met"] == (payload["mean_val_score"] >= 0.9 and all(
        b["champion"]["val_score"] >= 0.9 for b in json.loads(
            (data_root / "v1" / "leaderboard.json").read_text(encoding="utf-8"))["targets"].values()))
    meta = json.loads((data_root / "v1" / "model_생산효율(%).json").read_text(encoding="utf-8"))
    assert "sklearn_version" in meta


def test_train_unitless_names_and_eval_alias(trained):
    """단위를 뺀 이름(README 예시 B)도 전체 열 이름으로 맞추고, 수식은 전체 이름으로 저장해 예측 때 그대로 다시 계산한다."""
    work, data, _ = trained
    payload = run_tool(
        "ml_train.py", "--file", data, "--features", "반응온도,소비전력,압축압력,촉매비율,수분함량", "--targets", "생산효율",
        "--eval", "에너지집중도 = 소비전력 / (반응온도 + 1e-6); 로그비 = log1p(에너지집중도)",
        "--tag", "alias",
    )
    assert payload["targets"] == [TARGET]
    assert payload["column_aliases"]["생산효율"] == TARGET
    assert payload["description"] == "smoke test"  # inherited from v1
    meta = json.loads((work / "sample_ml_data" / payload["version"] / "model_생산효율(%).json").read_text(encoding="utf-8"))
    # stored formulas use full names so inference replays them exactly
    assert meta["derived_evals"][0] == "에너지집중도 = 소비전력(kW) / (반응온도(°C) + 1e-6)"
    assert "로그비" in meta["features_required"]


def test_train_column_typo_suggests(trained):
    _, data, _ = trained
    # 2.9.33: '반응온도(C)' is the same column (°C / C) and matches; a typo is a [후보] (never an automatic [추천])
    payload = run_fail("ml_train.py", "--file", data, "--features", "반응온도x,소비전력(kW)", "--targets", TARGET)
    assert payload["agent_directive"].startswith("STOP_RETRY")
    assert "[후보]" in payload["message"] and "[추천]" not in payload["message"] and "반응온도(°C)" in payload["message"]


def test_train_eval_leakage_blocked(trained):
    _, data, _ = trained
    payload = run_fail("ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET, "--eval", "누수 = 생산효율 * 2")
    assert "타겟 누수" in payload["message"]


def test_train_eval_rejects_code(trained):
    _, data, _ = trained
    payload = run_fail("ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET,
                       "--eval", "x = __import__('os').getcwd()")
    assert payload["agent_directive"].startswith("STOP_RETRY")


def test_train_strings_with_commas_and_placeholders(tmp_path):
    """천 단위 쉼표와 '-' 자리표시 칸을 숫자로 읽는다 (pandas 3 'str' dtype). 상수 열의 NaN 상관계수가 JSON을 깨지 않는다."""
    import pandas as pd
    p = tmp_path / "comma.csv"
    make_sample(str(p))
    df = pd.read_csv(p, encoding="utf-8-sig")
    df["소비전력(kW)"] = df["소비전력(kW)"].map(lambda v: f"{v * 100:,.1f}")
    df["수분함량(%)"] = df["수분함량(%)"].astype(object)
    df.loc[3, "수분함량(%)"] = "-"
    df["상수열"] = 1.0  # constant column -> NaN correlation must not break JSON
    df.to_csv(p, index=False, encoding="cp949")
    payload = run_tool("ml_train.py", "--file", str(p), "--features", FEATURES + ",상수열", "--targets", TARGET)
    hist = (tmp_path / "comma_ml_data" / "history.jsonl").read_text(encoding="utf-8")
    assert "NaN" not in hist
    assert payload["mean_val_score"] > 0.5 and "mean_val_r2" not in payload


def test_train_history_mode_is_json(trained):
    _, data, _ = trained
    payload = run_tool("ml_train.py", "--file", data, "--history")
    assert payload["record_count"] >= 1 and payload["records"][0]["version"] == "v1"


def test_history_without_records_is_ok_zero(tmp_path):
    # R2 (2.9.46): 0 records is a result (2.9.45: "status": "no_history", no verified key)
    proc, res = run_raw("ml_train.py", "--file", str(tmp_path / "새데이터.csv"), "--history")
    assert proc.returncode == 0 and res["status"] == "ok" and res["verified"] is True, res
    assert res["record_count"] == 0 and res["count"] == 0 and res["records"] == [] and res["no_history"] is True, res
    assert res["agent_directive"].startswith("VERIFIED: 0건") and "next_commands" not in res, res


def test_blank_rows_only_input_says_zero_rows(tmp_path):
    p = tmp_path / "blank.csv"
    p.write_text("온도,압력,수율\n,,\n,,\n", encoding="utf-8-sig")
    for script, args in (("ml_train.py", ["--features", "온도,압력", "--targets", "수율"]), ("ml_map.py", ["--input", "온도=1"])):
        res = run_fail(script, "--file", str(p), *args)
        assert "데이터가 0행" in res["message"] and res["agent_directive"].startswith("STOP_RETRY"), res


def test_bad_cli_args_emit_json():
    proc, payload = run_raw("ml_train.py", "--no-such-option")
    assert proc.returncode == 2 and payload["status"] == "error"


def test_predict_inline(trained):
    work, _, _ = trained
    payload = run_tool("ml_predict.py", "--model", str(work / "sample_ml_data" / "v1"), "--input", INLINE)
    assert payload["agent_directive"].startswith("VERIFIED")
    assert payload["drift_warning"] is False
    assert len(payload["predictions_summary"][TARGET]["predictions"]) == 1


def test_predict_inline_json_file_and_unitless_keys(trained):
    work, _, _ = trained
    f = work / "input.json"
    # PowerShell 5 'Out-File' style: UTF-16 with BOM, keys without units
    f.write_text(json.dumps({"반응온도": 240, "소비전력": 65, "압축압력": 3, "촉매비율": 0.8, "수분함량": 20},
                            ensure_ascii=False), encoding="utf-16")
    payload = run_tool("ml_predict.py", "--model", str(work / "sample_ml_data" / "v1"), "--input", str(f))
    assert payload["column_aliases"]["반응온도"] == "반응온도(°C)"


def test_predict_extrapolation_flag(trained):
    work, _, _ = trained
    payload = run_tool("ml_predict.py", "--model", str(work / "sample_ml_data" / "v1"),
                       "--input", INLINE.replace("반응온도(°C)=240", "반응온도(°C)=900"))
    assert payload["drift_warning"] is True
    assert payload["extrapolation_warnings"]


def test_predict_batch(trained):
    work, data, _ = trained
    out = work / "pred.csv"
    payload = run_tool("ml_predict.py", "--model", str(work / "sample_ml_data" / "v1"), "--file", data, "--save", str(out))
    assert out.is_file() and payload["total_rows"] == 120


def test_predict_missing_inputs_offer_median_fill(trained):
    """2.9.55: --input without some numeric inputs -> needs_user_decision + a command that fills training medians (never run)."""
    work, _, _ = trained
    vdir = str(work / "sample_ml_data" / "v1")
    payload = run_fail("ml_predict.py", "--model", vdir, "--input", "반응온도=240,소비전력=65")    # 2.9.54: '누락' only
    assert payload["error_type"] == "needs_user_decision" and payload["agent_directive"].startswith("STOP_RETRY:")
    assert "next_commands[0]" in payload["agent_directive"] and "[추천]" not in payload["message"]
    assert set(payload["missing_inputs"]) == {"압축압력(bar)", "촉매비율(wt%)", "수분함량(%)"}, payload
    meta = json.load(open(os.path.join(vdir, "model_생산효율(%).json"), encoding="utf-8"))
    med = meta["feature_stats"]["수분함량(%)"]["median"]
    assert abs(payload["fill_values"]["수분함량(%)"] - med) < 1e-3 and "수분함량(%)=" in payload["message"]
    nxt = payload["next_commands"][0]
    assert nxt["needs_user_confirm"] is True and nxt["command"].startswith("python ")
    ok = run_tool("ml_predict.py", *split_command(nxt["command"])[2:])               # the offered command works as given
    assert ok["predictions_summary"][TARGET]["predictions"] and not ok.get("missing_input_warning")
    # a misspelled key is not filled: the usual [후보] for the name
    bad = run_fail("ml_predict.py", "--model", vdir, "--input", "반응온도=240,소비전력=65,압축압력=3,촉매비율=0.8,수분함럄=20")
    assert bad["error_type"] == "needs_user_decision" and "[후보] 1) " in bad["message"] and bad["confirm_commands"], bad


def test_train_target_in_features_offers_both_commands(trained):
    """2.9.55: a target also in --features -> [후보] of two full commands (out of --targets / out of --features). The user picks: they are
    confirm_commands (needs_user_decision), never next_commands (run without asking)."""
    _, data, _ = trained
    payload = run_fail("ml_train.py", "--file", data, "--features", "반응온도(°C),소비전력(kW),수분함량(%)",
                       "--targets", "생산효율(%),수분함량(%)", "--desc", "누수 시험")
    assert "타겟 누수:" in payload["message"] and "[후보] 1) " in payload["message"] and " 2) " in payload["message"]
    assert payload["error_type"] == "needs_user_decision" and "next_commands" not in payload, payload
    one, two = (split_command(c["command"]) for c in payload["confirm_commands"])
    assert one[one.index("--targets") + 1] == "생산효율(%)" and "수분함량(%)" in one[one.index("--features") + 1]
    assert two[two.index("--features") + 1] == "반응온도(°C),소비전력(kW)" and two[two.index("--targets") + 1] == "생산효율(%),수분함량(%)"
    assert "--desc" in one and "--desc" in two                                    # the other options are kept


@pytest.mark.parametrize("shell", ["powershell", "pwsh", "cmd", "bash"])
def test_offered_commands_survive_each_shell(trained, monkeypatch, shell):
    """2.9.68: next_commands의 따옴표는 실행 셸에 맞춤 (ML_KIT_SHELL, 기본은 OpenCode처럼 pwsh > Windows PowerShell 5.1).
    그 셸이 실행하면(shell_sim) --input·--desc 값이 글자 그대로 도구에 들어간다."""
    work, data, _ = trained
    monkeypatch.setenv("ML_KIT_SHELL", shell)
    vdir = str(work / "sample_ml_data" / "v1")
    payload = run_fail("ml_predict.py", "--model", vdir, "--input", "반응온도=240,소비전력=65")
    nxt = payload["next_commands"][0]["command"]
    argv = shell_sim.run_shell(nxt, shell)
    assert argv[argv.index("--model") + 1] == vdir and argv[argv.index("--input") + 1].startswith("반응온도=240,소비전력=65,"), nxt
    assert run_tool("ml_predict.py", *argv[2:])["predictions_summary"][TARGET]["predictions"]
    desc = "누수 $x `t '시험' \"따옴표\" 100% {{a}} @b #c; C:\\a b\\"
    payload = run_fail("ml_train.py", "--file", data, "--features", "반응온도(°C),소비전력(kW),수분함량(%)",
                       "--targets", "생산효율(%),수분함량(%)", "--desc", desc)
    assert "타겟 누수:" in payload["message"], payload
    for c in payload["confirm_commands"]:
        argv = shell_sim.run_shell(c["command"], shell)
        assert argv[argv.index("--desc") + 1] == desc, c["command"]


DEPLOY_FILES = ("ml_predict.py", "_ml_runtime.py")      # 예측만 하는 PC에 복사할 파일 (README "단독 배포")


def test_predict_without_runtime_is_one_json_line(tmp_path):
    """ml_predict.py만 복사하고 _ml_runtime.py를 빠뜨리면 traceback 대신 JSON 1줄로 무엇을 복사할지 알린다."""
    import shutil
    shutil.copy(os.path.join(KIT, "ml_predict.py"), tmp_path / "ml_predict.py")
    proc = subprocess.run([sys.executable, "-I", str(tmp_path / "ml_predict.py"), "--model", "v1", "--input", "a=1"],
                          capture_output=True, text=True, encoding="utf-8", timeout=60, cwd=str(tmp_path))
    assert proc.returncode == 1 and len(proc.stdout.strip().splitlines()) == 1, proc.stdout + proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["status"] == "error" and "_ml_runtime.py" in payload["message"]
    assert payload["agent_directive"].startswith("STOP_RETRY:")
    # Windows 파이프의 기본 CP949를 흉내 냄 (2.9.80 집 PC: 이 JSON이 CP949로 나가 UTF-8로 읽지 못함). -I는 PYTHONIOENCODING을 무시함
    env = {k: v for k, v in os.environ.items() if k not in ("PYTHONPATH", "PYTHONUTF8")}
    env["PYTHONIOENCODING"] = "cp949"
    proc = subprocess.run([sys.executable, str(tmp_path / "ml_predict.py"), "--model", "v1", "--input", "a=1"], env=env,
                          capture_output=True, timeout=60, cwd=str(tmp_path))
    assert json.loads(proc.stdout.decode("utf-8"))["error_type"] == "missing_file", proc.stdout


def test_predict_standalone_deployment(trained, tmp_path):
    """ml_predict.py + _ml_runtime.py와 버전 폴더만으로 예측이 된다 (단독 배포)."""
    import shutil
    work, _, _ = trained
    for f in DEPLOY_FILES:
        shutil.copy(os.path.join(KIT, f), tmp_path / f)
    shutil.copytree(work / "sample_ml_data" / "v1", tmp_path / "v1")
    proc = subprocess.run(
        [sys.executable, "-I", str(tmp_path / "ml_predict.py"), "--model", str(tmp_path / "v1"), "--input", INLINE],
        capture_output=True, text=True, encoding="utf-8", timeout=300, cwd=str(tmp_path),
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    payload = json.loads(proc.stdout.strip().splitlines()[-1])
    assert payload["status"] == "ok" and payload["predictions_summary"][TARGET]["predictions"]


needs_docx = pytest.mark.skipif(not HAS_DOCX, reason="python-docx not installed in this environment")


@pytest.mark.report
@pytest.mark.skipif(HAS_DOCX, reason="only meaningful when python-docx is missing")
def test_report_without_docx_returns_json(trained):
    work, _, _ = trained
    payload = run_fail("ml_report.py", "--model", str(work / "sample_ml_data" / "v1"))
    assert payload["agent_directive"].startswith("STOP_RETRY") and "python-docx" in payload["message"]


@needs_docx
@pytest.mark.report
def test_report(trained):
    work, _, _ = trained
    out = work / "report.docx"
    payload = run_tool("ml_report.py", "--model", str(work / "sample_ml_data" / "v1"), "--targets", TARGET, "--save", str(out))
    assert out.is_file() and payload["version"] == "v1" and "best_version" not in payload


@needs_docx
@pytest.mark.report
def test_report_unknown_version_fails(trained):
    work, _, _ = trained
    payload = run_fail("ml_report.py", "--model", str(work / "sample_ml_data" / "v99"))
    assert "v99" in payload["message"]
    assert "[후보] v1" in payload["message"] or "[추천] python" in payload["message"], payload     # 2.9.48


@needs_docx
@pytest.mark.report
def test_report_unknown_version_recommends_the_only_one(tmp_path):
    root = tmp_path / "y_ml_data"
    (root / "v1").mkdir(parents=True)
    (root / "v1" / "leaderboard.json").write_text("{}", encoding="utf-8")
    payload = run_fail("ml_report.py", "--model", str(root / "v5"), "--title", "제목")
    # 2.9.48: the one trained version is recommended as a full command (--model), the other options kept
    rec = payload["message"].split("[추천] ")[1]
    assert rec.startswith("python ") and rec.endswith(str(root / "v1")) and "--title" in rec, rec
    assert rec.count(" --model ") == 1, rec


@needs_docx
@pytest.mark.report
def test_report_unknown_version_lists_candidates(tmp_path):
    root = tmp_path / "x_ml_data"
    for v in ("v1", "v3"):
        (root / v).mkdir(parents=True)
        (root / v / "leaderboard.json").write_text("{}", encoding="utf-8")
    (root / "v2").mkdir()                                            # no leaderboard.json: not a trained version
    payload = run_fail("ml_report.py", "--model", str(root / "v9"))
    assert "[후보] v1, v3" in payload["message"] and "[추천]" not in payload["message"], payload


# ---- 도움말과 STOP_RETRY 안내 문구 ----------------------------------------------------------------------


@pytest.mark.parametrize("script", ["ml_train.py", "ml_predict.py", "ml_report.py"])
def test_help_is_stderr_text_and_one_json_line(script, data_dir):
    """-h 텍스트는 stderr, stdout은 SKILL.md를 가리키는 JSON 1줄 (R7). 오류로 기록하지 않는다."""
    proc, payload = run_raw(script, "-h")          # run_raw asserts exactly 1 stdout line
    assert proc.returncode == 0 and "usage:" in proc.stderr
    assert payload["status"] == "ok" and payload["command"] == "help" and payload["verified"] is True
    assert payload["agent_directive"].startswith("VERIFIED:") and "SKILL.md" in payload["agent_directive"]
    assert os.path.isfile(payload["skill_file"]) and os.path.dirname(payload["skill_file"]).startswith(os.path.join(KIT, "skills"))
    assert _log(data_dir) == ""


def test_stop_retry_directive_mentions_full_command():
    d = rt.error_payload(rt.InputError("x [추천] python a.py"))["agent_directive"]
    assert "전체 명령" in d and "그대로 1번만 실행" in d and "confirm_commands" in d


# ---- 오류 로그 <데이터 폴더>/logs/errors.txt (AGENTS.md 오류 로그 저장, CLI-dev/README.md R6) ----


@pytest.fixture
def data_dir(tmp_path, monkeypatch):
    d = tmp_path / "data"
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(d))
    monkeypatch.setattr(rt, "_RECOVERED", [])
    monkeypatch.setattr(rt, "_RECOVERY_ATTEMPTS", [])
    return d


def _log(d):
    p = d / "logs" / "errors.txt"
    return p.read_text(encoding="utf-8") if p.exists() else ""


def _run_inproc(capsys, body):
    """mp.run_cli를 같은 프로세스에서 실행하고 (exit 코드 또는 None, stdout JSON 1줄)을 돌려준다."""
    code = None
    try:
        rt.run_cli(body)
    except SystemExit as e:
        code = e.code
    lines = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    return code, json.loads(lines[0])


def test_default_data_dir_is_tool_folder(monkeypatch):
    monkeypatch.delenv("ML_KIT_DATA_DIR", raising=False)
    assert rt.get_data_dir() == os.path.join(KIT, "data")
    assert rt.error_log_path() == os.path.join(KIT, "data", "logs", "errors.txt")


def test_input_error_exit1_logged_without_traceback(data_dir, tmp_path):
    proc, payload = run_raw("ml_predict.py", "--model", str(tmp_path / "없는폴더"), "--input", "a=1")
    assert proc.returncode == 1 and payload["agent_directive"].startswith("STOP_RETRY")
    log = _log(data_dir)
    head = log.splitlines()[0]
    assert head.startswith("===== ") and f"| ml-kit {rt.VERSION} | Python " in head
    assert "명령: ml_predict.py --model " in log and "없는폴더" in log
    assert "결과: input | " in log
    assert "traceback:" not in log and "복구 시도:" not in log
    assert log.endswith("\n\n")
    assert "errors.txt" not in json.dumps(payload, ensure_ascii=False)  # log path is never in stdout JSON


def test_argument_error_exit2_logged(data_dir):
    proc, payload = run_raw("ml_predict.py", "--input", "a=1")  # --model missing
    assert proc.returncode == 2 and payload["status"] == "error"
    log = _log(data_dir)
    assert "명령: ml_predict.py --input a=1" in log and "결과: input | CLI 인자 오류" in log


def test_internal_error_logged_with_traceback(data_dir, capsys):
    def body():
        raise RuntimeError("내부 버그")
    code, payload = _run_inproc(capsys, body)
    assert code == 1 and payload["error_type"] == "internal" and payload["agent_directive"].startswith("ERROR:")
    log = _log(data_dir)
    assert "결과: internal | RuntimeError: 내부 버그" in log
    assert "traceback:\nTraceback (most recent call last):" in log and "RuntimeError: 내부 버그" in log


def test_success_and_interrupt_not_logged(data_dir, capsys):
    code, payload = _run_inproc(capsys, lambda: {"status": "ok", "agent_directive": "VERIFIED: ok"})
    assert code is None and payload["status"] == "ok"

    def interrupted():
        raise KeyboardInterrupt()
    code, payload = _run_inproc(capsys, interrupted)
    assert code == 130 and payload["error_type"] == "interrupted"
    assert not (data_dir / "logs" / "errors.txt").exists()


def test_recovery_attempt_line(data_dir, capsys):
    def body():
        def locked():
            raise PermissionError("locked")
        rt.retry_fs(locked, attempts=2, delay=0)
    code, payload = _run_inproc(capsys, body)
    assert code == 1
    assert "복구 시도: 파일 작업 2회 재시도 실패" in _log(data_dir)


def test_rotation_at_size_limit(data_dir, monkeypatch):
    logs = data_dir / "logs"
    logs.mkdir(parents=True)
    (logs / "errors.txt").write_text("OLD ENTRY\n" * 10, encoding="utf-8")
    monkeypatch.setattr(rt, "ERROR_LOG_MAX", 50)
    rt.write_error_log({"error_type": "input", "message": "새 항목"}, argv=["ml_train.py", "--file", "a.csv"])
    assert "OLD ENTRY" in (logs / "errors.old.txt").read_text(encoding="utf-8")
    new = (logs / "errors.txt").read_text(encoding="utf-8")
    assert "OLD ENTRY" not in new and "결과: input | 새 항목" in new and "명령: ml_train.py --file a.csv" in new


def test_rotation_failure_keeps_entry(data_dir, monkeypatch):
    logs = data_dir / "logs"
    logs.mkdir(parents=True)
    (logs / "errors.txt").write_text("OLD ENTRY\n" * 10, encoding="utf-8")
    monkeypatch.setattr(rt, "ERROR_LOG_MAX", 50)

    def deny(*_a, **_k):
        raise PermissionError("rotation locked")
    monkeypatch.setattr(rt.os, "replace", deny)
    rt.write_error_log({"error_type": "input", "message": "회전 실패에도 기록"}, argv=["ml_report.py"])
    assert "회전 실패에도 기록" in (logs / "errors.txt").read_text(encoding="utf-8")


def test_logging_failure_keeps_json_contract(tmp_path, monkeypatch, capsys):
    blocker = tmp_path / "not_a_dir"
    blocker.write_text("x", encoding="utf-8")  # data dir is a file -> makedirs fails inside write_error_log
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(blocker))

    def body():
        raise rt.InputError("입력 오류")
    code, payload = _run_inproc(capsys, body)
    assert code == 1 and payload["status"] == "error" and payload["message"] == "입력 오류"


# ---- En-Genius 스킬로 실행할 때의 데이터 폴더 ----------------------------------------------------------


# En-Genius 트레이는 ENGENIUS_OUT을 두고, 압축을 푼 스킬 폴더(engenius-skill)를 실행 뒤 지운다. 그 임시 폴더에서 실행할 때만
# 데이터 폴더(오류 로그)가 <USERPROFILE>/Documents/En-Genius/CLI-data/ml-kit/data다. C:\CLI에서 실행하면 ENGENIUS_OUT이
# 있어도 도구 폴더의 data를 쓴다. ML_KIT_DATA_DIR이 늘 우선한다.


PROBE = ("import sys\n"
         "sys.path.insert(0, sys.argv[1])\n"
         "import _ml_runtime\n"
         "print(_ml_runtime.get_data_dir())\n")


def _env(tmp_path, **extra):
    env = {k: v for k, v in os.environ.items() if k not in ("ML_KIT_DATA_DIR", "ENGENIUS_OUT")}
    env["USERPROFILE"] = str(tmp_path / "home")
    env.update(extra)
    return env


def _unpacked(tmp_path):
    d = tmp_path / "engenius-skill" / "r1"
    if not d.exists():
        shutil.copytree(KIT, d, ignore=shutil.ignore_patterns("data", "__pycache__"))
    return str(d)


def _probe(tmp_path, tool_dir=KIT, **extra):
    out = subprocess.run([sys.executable, "-c", PROBE, tool_dir], capture_output=True, timeout=120,
                         env=_env(tmp_path, **extra))
    return out.stdout.decode("utf-8").strip().splitlines()[-1]


def test_engenius_run_from_cli_folder_shares_tool_folder(tmp_path):
    assert _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")) == os.path.join(KIT, "data")


def test_engenius_temp_run_uses_fixed_folder(tmp_path):
    assert _probe(tmp_path, _unpacked(tmp_path), ENGENIUS_OUT=str(tmp_path / "out")) == os.path.join(
        str(tmp_path / "home"), "Documents", "En-Genius", "CLI-data", "ml-kit", "data")


def test_normal_run_keeps_tool_folder(tmp_path):
    assert _probe(tmp_path) == os.path.join(KIT, "data")


def test_env_wins_over_engenius(tmp_path):
    assert _probe(tmp_path, _unpacked(tmp_path), ENGENIUS_OUT=str(tmp_path / "out"), ML_KIT_DATA_DIR=str(tmp_path / "d")) == str(tmp_path / "d")


def test_engenius_error_log_lands_in_fixed_folder(tmp_path):
    proc = subprocess.run([sys.executable, os.path.join(_unpacked(tmp_path), "ml_predict.py"), "--bogus"], capture_output=True,
                          timeout=120, env=_env(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")))
    assert proc.returncode == 2 and json.loads(proc.stdout.decode("utf-8").strip().splitlines()[-1])["status"] == "error"
    assert (tmp_path / "home" / "Documents" / "En-Genius" / "CLI-data" / "ml-kit" / "data" / "logs" / "errors.txt").is_file()


# ---- 공통 런타임 (_ml_runtime.py) ----------------------------------------------------------------


def test_json_safe_nan_and_numpy():
    assert rt.dumps_json({"a": np.nan, "b": np.int64(3), "c": [np.float32(1.5)]}) == '{"a": null, "b": 3, "c": [1.5]}'


def test_read_text_file_encodings(tmp_path):
    for enc in ("utf-8", "utf-8-sig", "utf-16", "cp949"):
        p = tmp_path / f"t_{enc}.txt"
        p.write_text("반응온도(°C)", encoding=enc)
        assert rt.read_text_file(str(p)) == "반응온도(°C)"


ENTRY_POINTS = ("ml_train.py", "ml_predict.py", "ml_report.py", "ml_export.py", "ml_map.py", "ml_dataset.py")


def test_entry_points_import_only_the_runtime():
    """진입점은 _ml_runtime만 공통으로 import한다. 아무도 ml_predict를 import하지 않고(예측 단독 배포), _ml_runtime은 어떤
    진입점도 import하지 않는다. 예외는 보고서와 함께 성능예측표를 만드는 ml_report -> ml_export 하나
    (2.9.84: 성능예측표는 다시 학습하지 않아 ml_train을 쓰지 않음). AGENTS.md ml-kit 규칙."""
    allowed = {"ml_report.py": {"ml_export"}}
    bad = []
    for name in ENTRY_POINTS + ("_ml_runtime.py",):
        tree = ast.parse(open(os.path.join(KIT, name), encoding="utf-8").read())
        for node in ast.walk(tree):
            mods = [a.name for a in node.names] if isinstance(node, ast.Import) else \
                [node.module or ""] if isinstance(node, ast.ImportFrom) else []
            for m in mods:
                if m.split(".")[0] + ".py" in ENTRY_POINTS and m not in allowed.get(name, ()):
                    bad.append(f"{name}: import {m}")
    assert not bad, bad


def test_no_module_level_name_assigned_twice():
    """ml_predict.py와 _ml_runtime.py의 모듈 수준 이름은 두 파일을 합쳐 한 번만 정의한다 (2.9.20: _THOUSANDS_RE를 두 번
    정의해 소수점 쉼표 CSV 검사가 뒤의 정규식을 썼음. 예측부가 공통 런타임의 이름을 다시 정의해도 같은 일이 남).
    """
    seen, dup = set(), []
    for name in ("_ml_runtime.py", "ml_predict.py"):
        tree = ast.parse(open(os.path.join(KIT, name), encoding="utf-8").read())
        for node in tree.body:
            targets = node.targets if isinstance(node, ast.Assign) else \
                [node.target] if isinstance(node, (ast.AnnAssign, ast.AugAssign)) else []
            names = [n.id for t in targets for n in ast.walk(t) if isinstance(n, ast.Name)]
            if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
                names.append(node.name)
            for n in names:
                (dup.append(f"{name}: {n}") if n in seen else seen.add(n))
    assert not dup, dup


# ---- 라이브러리 없음 (R5, 2.9.45): traceback 대신 missing_library JSON 1줄 ------------------------------------------


def _without(tmp_path, *modules):
    """modules를 import하면 ModuleNotFoundError가 나는 폴더 (PYTHONPATH 맨 앞)."""
    fake = tmp_path / "fake_site"
    fake.mkdir(exist_ok=True)
    for m in modules:
        (fake / f"{m}.py").write_text(f"raise ModuleNotFoundError(\"No module named '{m}'\", name='{m}')\n",
                                      encoding="utf-8")
    return str(fake)


def _run_without(tmp_path, fake, script, *args):
    env = {**os.environ, "PYTHONPATH": fake, "ML_KIT_DATA_DIR": str(tmp_path / "data")}
    proc = subprocess.run([sys.executable, os.path.join(KIT, script), *args], capture_output=True, text=True,
                          encoding="utf-8", env=env, timeout=120)
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert len(lines) == 1, (proc.stdout, proc.stderr)
    return proc, json.loads(lines[0])


@pytest.mark.parametrize("script,module,pip", [
    ("ml_predict.py", "numpy", "numpy"), ("ml_train.py", "sklearn", "scikit-learn"), ("ml_map.py", "pandas", "pandas"),
    pytest.param("ml_report.py", "joblib", "joblib", marks=pytest.mark.report),
    pytest.param("ml_export.py", "sklearn", "scikit-learn", marks=pytest.mark.report)])
def test_missing_library_is_one_json_line(tmp_path, script, module, pip):
    args = {"ml_predict.py": ["--model", str(tmp_path), "--input", "a=1"], "ml_train.py": ["--file", "x.csv", "--features", "a", "--targets", "b"],
            "ml_map.py": ["--file", "x.csv", "--input", "a=1"], "ml_report.py": ["--model", str(tmp_path)],
            "ml_export.py": ["--model", str(tmp_path)]}[script]
    proc, res = _run_without(tmp_path, _without(tmp_path, module), script, *args)
    assert proc.returncode == 1 and res["status"] == "error", res              # 2.9.44: traceback, no JSON
    assert res["error_type"] == "missing_library" and res["missing_package"] == pip, res
    assert res["install_command"] == f"python -m pip install {pip}", res
    assert res["agent_directive"].startswith("STOP_RETRY:") and "직접 설치하지 마세요" in res["agent_directive"], res
    assert "Traceback" not in proc.stderr
    log = (tmp_path / "data" / "logs" / "errors.txt").read_text(encoding="utf-8")
    assert "missing_library" in log


def test_import_error_inside_run_is_missing_library():
    res = rt.error_payload(ModuleNotFoundError("No module named 'scipy.spatial'", name="scipy.spatial"))
    assert res["error_type"] == "missing_library" and res["install_command"] == "python -m pip install scipy", res
    res = rt.error_payload(ImportError("No module named docx"))
    assert res["missing_package"] == "python-docx" and res["agent_directive"].startswith("STOP_RETRY:"), res


# ---- 로그 태그 (AGENTS.md 문자 규칙, 2.9.50) --------------------------------------------------------------------

LOG_TAGS = ("[INFO]", "[PROGRESS]", "[WARN]", "[ERROR]", "[STATUS]")


def test_stderr_uses_only_shared_log_tags(trained, tmp_path):
    import re
    work, data, _ = trained
    proc, _ = run_raw("ml_train.py", "--file", data, "--features", "반응온도(°C),소비전력(kW)", "--targets", TARGET, "--tag", "tags")
    hist, _ = run_raw("ml_train.py", "--file", data, "--history")
    for err in (proc.stderr, hist.stderr):
        assert "\r" not in err.replace("\r\n", "\n")
        tagged = [ln for ln in err.splitlines() if re.match(r"\s*\[[A-Z]", ln)]
        bad = [ln for ln in tagged if not ln.lstrip().startswith(LOG_TAGS)]
        assert not bad, bad                                           # 2.9.49: '[PROGRESS  5%] [=---]', '[ML EXPERIMENT ...]'
    prog = [ln for ln in proc.stderr.splitlines() if ln.startswith("[PROGRESS]")]
    assert prog and all(re.match(r"\[PROGRESS\] \d{1,3}% \S", ln) for ln in prog), prog[:3]
    assert "[STATUS] 실험 이력" in hist.stderr


# ---- 인자 이름 통일 (AGENTS.md ACI 설계 권장 5, 2.9.51): 저장은 --save, 타겟은 --targets/--targets. 옛 이름은 받지 않음 (베타) ----


@pytest.mark.parametrize("script,args", [
    ("ml_predict.py", ["--model", "x", "--input", "a=1", "--out", "p.csv"]), ("ml_predict.py", ["--model", "x", "--input", "a=1", "-o", "p.csv"]),
    ("ml_report.py", ["--model", "x", "--out", "r.docx"]), ("ml_report.py", ["--model", "x", "-o", "r.docx"]),
    ("ml_report.py", ["--model", "x", "-t", "y"]),         # '--target' is argparse's abbreviation of --targets (same meaning)
    # 2.9.52: removed options (user decision)
    ("ml_predict.py", ["--model", "x", "--input", "a=1", "--meta", "m.json"]), ("ml_predict.py", ["--model", "x", "--input", "a=1", "--encoding", "cp949"]),
    ("ml_train.py", ["--file", "x.csv", "--features", "a", "--targets", "b", "--encoding", "cp949"]),
    ("ml_report.py", ["-d", "x_ml_data"]), ("ml_report.py", ["--data-dir", "x_ml_data"]),
    ("ml_report.py", ["--model", "x", "--version", "v1"]), ("ml_report.py", ["--model", "x", "--version", "v1"])])
def test_old_option_names_are_argument_errors(data_dir, script, args):
    proc, res = run_raw(script, *args)
    assert proc.returncode == 2 and res["status"] == "error", res              # 2.9.50: accepted


@needs_docx
@pytest.mark.report
def test_report_targets_and_save(trained):
    work, _, _ = trained
    out = work / "r_targets.docx"
    payload = run_tool("ml_report.py", "--model", str(work / "sample_ml_data" / "v1"), "--targets", TARGET, "--save", str(out))
    assert out.is_file() and payload["report_file"] == str(out), payload


# ---- 인자 오류 힌트 (2.9.69): 같은 뜻의 지어낸 옵션 -> [추천] 고친 전체 명령, 줄임말은 같은 뜻 표만 -----------------


@pytest.mark.parametrize("script,args,want", [
    ("ml_predict.py", ["--model-dir", "v1", "--input", "a=1"], ["--model"]),                 # required 오류여도
    ("ml_predict.py", ["--model", "v1", "--inputs", "a=1", "--json"], ["--input"]),
    ("ml_predict.py", ["--model", "v1", "--input", "a=1", "--output", "p.csv", "--encoding", "cp949"], ["--save"]),
    ("ml_train.py", ["--data", "d.csv", "--columns", "a,b", "--target-col", "y"], ["--file", "--features", "--targets"]),
    ("ml_map.py", ["--file", "d.csv", "--query", "a=1"], ["--input"]),
])
def test_same_meaning_option_gives_fixed_command(data_dir, script, args, want):
    proc, res = run_raw(script, *args)
    assert proc.returncode == 2 and res["agent_directive"].startswith("STOP_RETRY:") and "[추천]" in res["agent_directive"]
    name, fixed = recommended_command(res["message"])
    assert name == script and all(o in fixed for o in want), fixed
    real = {"--model", "--input", "--file", "--features", "--targets", "--save"}       # 그 명령의 진짜 옵션은 남음
    invented = {a for a in args if a.startswith("--") and a not in real} | {"cp949"}
    assert not invented & set(fixed), fixed


def test_inline_values_in_wrong_option_recommend_i(p1_model, tmp_path):
    """2.9.74 집 PC opencode: 한 건 예측에서 ml_train의 --features를 쓰거나('--features "온도=180,..."') 값을 --file에 넣거나, 한 건을 파일로 만들려다
    없는 파일을 --file로 줬을 때 '알 수 없는 옵션'/'파일 없음'만 받아 멈췄음 -> --input로 고친 전체 명령([추천]) 또는 --input 명령(next_commands)."""
    vals = "온도=180,압력=3,속도=45"
    # Gemma 4: '--eval "온도=180; 압력=3; ..."' (ml_train 옵션, ; 구분) -> 예전에는 철자만 비슷한 [후보] --vary/--model
    for args in (["--features", vals], ["--features", vals], ["--file", vals], ["--eval", "온도=180; 압력=3; 속도=45"]):
        proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], *args)
        assert proc.returncode in (1, 2) and "[추천]" in res["message"] and "[후보]" not in res["message"], res
        name, fixed = recommended_command(res["message"])
        assert name == "ml_predict.py" and fixed[fixed.index("--input") + 1] == vals and args[0] not in fixed, fixed
        assert run_tool("ml_predict.py", *fixed)["status"] == "ok"
    missing = str(tmp_path / "새조건.csv")
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--file", missing)
    assert proc.returncode == 1 and "--input" in res["message"] and "[추천]" not in res["message"], res
    nxt = res["next_commands"][0]
    cmd = split_command(nxt["command"])
    assert nxt["needs_input"] is True and "--file" not in cmd and cmd[cmd.index("--input") + 1].startswith("온도(°C)=<값>"), nxt
    assert nxt["placeholder"] == "<값>"                                    # 2.9.75: 다른 needs_input 명령과 같은 키
    assert not rt.looks_inline_input(str(tmp_path)) and not rt.looks_inline_input("a.csv")
    assert not rt.looks_inline_input("C:\\d\\a=1.csv") and rt.looks_inline_input("반응온도(°C)=182, 압력=3")


def test_model_path_underscore_and_made_up_model_folder_recommend_the_data_file(p1_model):
    """2.9.77 집 PC Gemma: '--model_path "<데이터 폴더>\\models"'(밑줄, 없는 폴더)에 철자만 비슷한 [후보]를 받아 되물었고,
    없는 폴더·데이터 폴더를 --model에 주면 '찾을 수 없습니다'만 받았음 -> 그 폴더의 학습한 데이터 파일로 고친 [추천] 하나."""
    folder = os.path.dirname(p1_model["f"])
    vals = "온도=180,압력=3,속도=45"
    for args in (["--model_path", os.path.join(folder, "models")], ["--model", os.path.join(folder, "models")], ["--model", folder]):
        proc, res = run_raw("ml_predict.py", "--input", vals, *args)
        assert proc.returncode in (1, 2) and "[추천]" in res["message"] and "[후보]" not in res["message"], res
        name, fixed = recommended_command(res["message"])
        k = fixed.index("--model") if "--model" in fixed else fixed.index("--model")
        assert name == "ml_predict.py" and fixed[k + 1] == p1_model["f"] and "--model_path" not in fixed, fixed
        assert run_tool("ml_predict.py", *fixed)["status"] == "ok"
    proc, res = run_raw("ml_predict.py", "--input", vals, "--model", p1_model["v"])          # 버전 폴더는 그대로
    assert proc.returncode == 0, res


def test_target_value_with_placeholder_input_recommends_find(p1_model):
    """2.9.79 집 PC Gemma P3: '--input "불순물=100,반응온도=<값>,..."'(타겟 값 + 자리표시자 하나)로 실행해 '<값>을 물으라'만 받고
    ml_map을 학습 데이터에 돌려 실패 -> 그 뜻 하나인 --find/--vary 명령을 [추천]. 타겟 없이 자리표시자만이면 원래대로 값을 물음."""
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "생산량=500,온도(°C)=<값>,압력(bar)=3,속도=45")
    assert proc.returncode == 1 and "[추천]" in res["message"], res
    name, fixed = recommended_command(res["message"])
    assert name == "ml_predict.py" and fixed[fixed.index("--find") + 1] == "생산량=500", fixed
    assert fixed[fixed.index("--vary") + 1] == "온도(°C)" and fixed[fixed.index("--input") + 1] == "압력(bar)=3,속도=45", fixed
    assert run_tool("ml_predict.py", *fixed)["status"] == "ok"
    # 찾을 입력을 아예 뺌 (P3 세 번째): 한 행을 붙이다 빠뜨렸을 수도 있어 [후보] 1) 찾기 2) 타겟 값 빼고 중앙값으로 예측
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "생산량=500,압력(bar)=3,속도=45")
    assert "[추천]" not in res["message"] and "[후보] 1) " in res["message"] and res["error_type"] == "needs_user_decision", res
    find, fill = (split_command(c["command"]) for c in res["confirm_commands"])
    assert find[find.index("--vary") + 1] == "온도(°C)" and find[find.index("--input") + 1] == "압력(bar)=3,속도=45", res
    assert "--find" not in fill and "생산량" not in fill[fill.index("--input") + 1] and "온도(°C)=" in fill[fill.index("--input") + 1]
    assert run_tool("ml_predict.py", *find[2:])["status"] == "ok" and run_tool("ml_predict.py", *fill[2:])["status"] == "ok"
    # 모델이 모르는 이름(오타)이 있으면 빈 입력을 셀 수 없음: 찾기 명령 없이 원래 이름 안내
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "생산량=500,온돜=<값>,압력(bar)=3,속도=45")
    assert proc.returncode == 1 and "--find" not in json.dumps(res, ensure_ascii=False), res
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "생산량=500,속도=45")                 # 빈 입력 둘: 고르지 않음
    assert proc.returncode == 1 and "[추천]" not in res["message"], res
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "온도(°C)=<값>,압력(bar)=3,속도=45")
    assert proc.returncode == 1 and "[추천]" not in res["message"] and "물은 뒤" in res["message"], res


def test_leading_verb_and_x_y_cols_recommend_fixed_command(data_dir):
    """2.9.77 집 PC solar-mini4: 'ml_train.py train --x-cols ... --y-cols ...' -> 'unrecognized arguments: train'만 받아 헤맴."""
    proc, res = run_raw("ml_train.py", "train", "--file", "d.csv", "--x-cols", "a,b", "--y-cols", "y")
    assert proc.returncode == 2 and "[추천]" in res["message"], res
    name, fixed = recommended_command(res["message"])
    assert name == "ml_train.py" and "train" not in fixed and fixed[fixed.index("--features") + 1] == "a,b"
    assert fixed[fixed.index("--targets") + 1] == "y" and "--x-cols" not in fixed, fixed


def test_abbreviations_only_through_the_same_meaning_table(data_dir):
    """allow_abbrev=False: '--data-ro 2'는 --data-row로 몰래 읽지 않고 [후보], '--overwrit'는 어떤 후보도 주지 않음.
    예전에 줄임말로 되던 같은 뜻('--target' = --targets)도 읽지 않고 고친 전체 명령을 [추천] (2.9.71: 한 가지 사용법)."""
    proc, res = run_raw("ml_train.py", "--file", "없음.csv", "--features", "a", "--targets", "b", "--data-ro", "2")
    assert proc.returncode == 2 and "[후보] 1) --data-row" in res["message"] and "[추천]" not in res["message"], res
    fixed = split_command(res["confirm_commands"][0]["command"])        # 2.9.73: 고른 번호의 전체 명령
    assert fixed[fixed.index("--data-row") + 1] == "2" and "--data-ro" not in fixed, fixed
    proc, res = run_raw("ml_train.py", "--file", "없음.csv", "--features", "a", "--targets", "b", "--overwrit")
    assert proc.returncode == 2 and "[후보]" not in res["message"] and "--overwrite" not in res["message"], res
    proc, res = run_raw("ml_train.py", "--file", "없음.csv", "--features", "a", "--target", "b")
    assert proc.returncode == 2 and "[INFO] 옵션 --target" not in proc.stderr, res
    _, fixed = recommended_command(res["message"])
    assert fixed[fixed.index("--targets") + 1] == "b" and "--target" not in fixed, fixed
    proc, res = run_raw("ml_train.py", "--file", "없음.csv", "--features", "a", "--targets", "b", "--description", "설명")      # 별칭 없음
    _, fixed = recommended_command(res["message"])
    assert proc.returncode == 2 and fixed[fixed.index("--desc") + 1] == "설명" and "--description" not in fixed, fixed
    p = rt.JsonArgumentParser()
    p.add_argument("--data-row", type=int)
    p.add_argument("--file")
    with pytest.raises(SystemExit):
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            p.parse_args(["--data-r", "3"])


def test_continue_is_counted_by_the_tool(monkeypatch, tmp_path):
    """2.9.71: CONTINUE 횟수는 도구가 셈. 새로 저장한 것 없이 바로 앞과 같은 progress면 STOP_RETRY(unfinished_limit),
    새로 저장했으면 같은 글이어도 계속, CONTINUE_MAX를 넘으면 멈춤."""
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path))
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["ml_train.py", "--file", "d.csv", "--features", "a", "--targets", "b"])
    rt.RUN_SAVED.clear()
    assert rt.unfinished_payload(rt.Unfinished("학습 1/5"))["continue_runs"] == 1
    rt.RUN_SAVED.append("x")
    assert rt.unfinished_payload(rt.Unfinished("학습 1/5"))["continue_runs"] == 2      # 저장했으면 진행한 것
    rt.RUN_SAVED.clear()
    stuck = rt.unfinished_payload(rt.Unfinished("학습 1/5"))
    assert stuck["status"] == "error" and stuck["error_type"] == "unfinished_limit", stuck
    monkeypatch.setattr(rt, "CONTINUE_MAX", 2)
    for i in range(2):
        assert rt.unfinished_payload(rt.Unfinished(f"단계 {i}"))["status"] == "ok"
    assert rt.unfinished_payload(rt.Unfinished("단계 9"))["error_type"] == "unfinished_limit"
    rt.continue_clear()
    assert rt._continue_state() == {}


def test_continue_next_commands_have_the_common_shape(monkeypatch, tmp_path):
    """C3: CONTINUE의 next_commands도 다른 결과처럼 {command, needs_input} (스킬 규칙 'next_commands의 command'), preview는 진행."""
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path))
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["ml_train.py", "--file", "d.csv", "--targets", "b"])
    rt.RUN_SAVED.clear()
    out = rt.unfinished_payload(rt.Unfinished("학습 2/5"))
    nxt = out["next_commands"][0]
    assert nxt["needs_input"] is False and nxt["command"].startswith("python ") and "ml_train.py" in nxt["command"], out
    assert out["preview"] == "진행 중: 학습 2/5" and "next_commands[0]의 command" in out["agent_directive"], out
    rt.continue_clear()


def test_time_limit_clock_starts_with_the_process():
    """B7: 시간 제한 감시(1줄 JSON을 직접 내는 시각)는 run_cli가 아니라 프로세스 시작부터 잰다 (무거운 import 시간 포함,
    에이전트의 120초와 같은 기준). 단계의 시계 확인(check_time)은 run_cli부터 (여유 30초가 시작 시간을 덮음)."""
    code = ("import sys, os, time; time.sleep(1.5); sys.path.insert(0, sys.argv[1]); sys.argv = ['ml_x.py']; "
            "os.environ['ML_KIT_TIME_LIMIT'] = '100'; import _ml_runtime as rt; "
            "rt.run_cli(lambda: {'status': 'ok', 'agent_directive': 'VERIFIED: x', "
            "'delay': round(rt.watchdog_delay(), 2), 'run': round(time.monotonic() - rt.RUN_START, 2)})")
    proc = subprocess.run([sys.executable, "-c", code, KIT], capture_output=True, text=True, encoding="utf-8", timeout=60)
    out = json.loads(proc.stdout.strip().splitlines()[-1])
    if sys.platform.startswith(("linux", "win")):
        assert out["delay"] <= 100 + rt.WATCHDOG_GRACE - 1.4, (out, proc.stderr)
    assert out["run"] < 1.0, out


def test_missing_prediction_file_typo_gives_numbered_files(p1_model, tmp_path):
    """2.9.75: 없는 --file가 같은 폴더 파일의 오타면 '--input로 예측' 안내가 아니라 그 파일들을 [후보] 1) 2)와 같은 번호의 전체 명령으로."""
    shutil.copy(p1_model["f"], tmp_path / "새조건_1차.csv")
    (tmp_path / "관계없음.txt").write_text("x", encoding="utf-8")
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--file", str(tmp_path / "새조건_1챠.csv"))
    assert proc.returncode == 1 and "[후보] 1) " in res["message"] and "관계없음" not in res["message"], res
    assert "next_commands" not in res, res
    cmd = split_command(res["confirm_commands"][0]["command"])
    assert cmd[cmd.index("--file") + 1] == str(tmp_path / "새조건_1차.csv"), cmd
    assert run_tool("ml_predict.py", *cmd[2:])["status"] == "ok"


def test_placeholder_or_text_in_inline_numbers_stops(p1_model):
    """2.9.76 검토: 학습 결과의 '--input "온도=<값>,..."' 틀을 그대로 실행하면 중앙값으로 채운 예측이 VERIFIED로 나갔음 -> 멈추고 값을 물으라고 함."""
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "온도(°C)=<값>,압력(bar)=3,속도=45")
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:") and "사용자에게 물은 뒤" in res["message"], res
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--input", "온도(°C)=170:190:10,압력(bar)=3,속도=45")
    assert proc.returncode == 1 and "--sweep" in res["message"], res
    assert run_tool("ml_predict.py", "--model", p1_model["v"], "--input", "온도(°C)=180,압력(bar)=3,속도=45")["status"] == "ok"
    assert rt.re.fullmatch(r"\s*<[^<>]*>\s*", "<값>")                                 # 글자 입력의 <값>도 막는 규칙


def test_constant_eval_is_refused(p1_model):
    """2.9.76 검토: '--eval "온도=182; 압력=3"'(한 건 예측 값을 학습 --eval에 넣은 실수)이 상수 파생변수로 학습되어 최고 버전이 됐음."""
    proc, res = run_raw("ml_train.py", "--file", p1_model["f"], "--targets", "생산량", "--eval", "온도=182; 압력=3")
    assert proc.returncode == 1 and "상수식" in res["message"] and "ml_predict.py --input" in res["message"], res
    proc, res = run_raw("ml_train.py", "--file", p1_model["f"], "--targets", "생산량", "--eval", "온도 = 압력(bar) * 2")
    assert proc.returncode == 1 and "데이터 열 '온도(°C)'과(와) 같습니다" in res["message"], res


def test_find_goal_in_wrong_option_is_not_turned_into_inputs(p1_model):
    """2.9.76 검토: '--goal "생산량=60" --vary 온도'가 [추천] --input 생산량=60으로 고쳐져 목표값이 입력으로 들어갔음."""
    proc, res = run_raw("ml_predict.py", "--model", p1_model["v"], "--goal", "생산량=600", "--vary", "온도(°C)")
    assert proc.returncode == 2 and " --input " not in res["message"], res


# ---- 2.9.79: 구조 정리 (스레드, [후보] 전체 명령, 파일 재시도) ------------------------------------------------------

def test_small_tables_fit_with_one_thread():
    """작은 표는 한 스레드가 더 빠름 (60행: 모든 코어 6.2초, 한 스레드 0.01초). 결과는 스레드 수와 무관."""
    assert rt.model_jobs(100) == 1 and rt.model_jobs(rt.MANY_THREADS_ROWS) == -1
    import ml_train
    for m in ml_train.get_regression_models(n_rows=100).values():
        assert getattr(m, "n_jobs", None) in (None, 1), m


def test_numbered_candidates_carry_their_full_commands(monkeypatch):
    """[후보]는 늘 번호. ask()로 만든 [후보]가 든 message의 InputError는 같은 번호의 confirm_commands를 가진다
    (글만 돌려주는 도우미도 전체 명령을 준다: 사용법 단순성 2번)."""
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["ml_train.py", "--file", "d.csv", "--features", "온두,압력", "--targets", "y"])
    monkeypatch.setattr(rt, "_ASKED", [])
    text = rt.ask(["온도", "온도2"], [rt.command_replacing("온두", c) for c in ("온도", "온도2")], "{}")
    assert text.startswith("[후보] 1) 온도 2) 온도2")
    e = rt.InputError(f"'온두' 없음: {text}")
    assert e.error_type == "needs_user_decision" and len(e.extra["confirm_commands"]) == 2
    assert "--features 온도2,압력" in e.extra["confirm_commands"][1]["command"]
    assert "confirm_commands" not in rt.InputError("다른 오류").extra            # 그 글이 없는 오류에는 붙지 않음
    assert rt.numbered(["a", "b"], [None, None], "{}") == ("[후보] 1) a 2) b", [])   # 명령이 없어도 번호


def test_retry_fs_waits_for_windows_lock_errors(monkeypatch):
    """winerror 32/33(공유·잠금 위반)도 재시도하고(office-cli와 같은 규칙), 없는 파일은 바로 올린다."""
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)
    calls = []

    def locked():
        calls.append(1)
        if len(calls) < 3:
            e = OSError("locked")
            e.winerror = 32
            raise e
        return "ok"
    assert rt.retry_fs(locked) == "ok" and len(calls) == 3
    with pytest.raises(FileNotFoundError):
        rt.retry_fs(lambda: (_ for _ in ()).throw(FileNotFoundError("x")))


@pytest.mark.parametrize("script,argv,want", [
    ("ml_train.py", ["-f", "없음.csv", "-x", "a", "-y", "b"], {"--file": "없음.csv", "--features": "a", "--targets": "b"}),
    ("ml_predict.py", ["-m", "v1", "-i", "a=1"], {"--model": "v1", "--input": "a=1"}),
    ("ml_map.py", ["-f", "없음.csv", "-i", "a=1", "-y", "b"], {"--file": "없음.csv", "--input": "a=1", "--targets": "b"}),
])
def test_short_options_are_refused_with_the_long_command(script, argv, want):
    """2.9.82: 한 글자 옵션은 없음 (사용법 단순성 1번, office-cli 1.4.88과 같음). 쓰면 실행하지 않고(exit 2) 긴 이름으로 고친 전체 명령을
    [추천]. 필수 옵션(--file)을 한 글자로 쓴 경우도 같음."""
    proc, res = run_raw(script, *argv)
    assert proc.returncode == 2 and "[추천]" in res["message"], res
    name, fixed = recommended_command(res["message"])
    assert name == script and not any(a in fixed for a in ("-f", "-m", "-i", "-x", "-y")), fixed
    for opt, val in want.items():                       # 경로 옵션(--file, --model)은 도구가 절대 경로로 넣음
        got = fixed[fixed.index(opt) + 1]
        assert (os.path.basename(got) if opt in ("--file", "--model") else got) == val, fixed
