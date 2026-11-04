# -*- coding: utf-8 -*-
"""사용자 확인 없이 지우거나 덮어쓰지 않는다 (AGENTS.md JIT 자가복구 3번, CLI-dev/README.md R3/R4).

학습 버전 폴더의 트랜잭션(되돌리기, 원장), 예측·보고서 결과 파일의 새 이름 저장, 잠김 복구,
실기 러너(tools/run_win_tests.py)의 폴더 정리와 AI 판독용 로그.
"""
import importlib.util
import json
import os
import shutil
import sys
import types

import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_sample import make_sample  # noqa: E402
from mlkit_helpers import DEV, FEATURES, TARGET, run_raw, run_tool  # noqa: E402
import ml_predict as mp  # noqa: E402


# ---- 학습 버전 폴더: 불완전한 폴더 보존, --overwrite, 트랜잭션 ----------------------------------


def _ledger(root):
    p = root / "history.jsonl"
    return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] if p.exists() else []


@pytest.fixture(scope="module")
def work(tmp_path_factory):
    """데이터 폴더에 이미 불완전한 v1(예: 복사하다 만 폴더)이 있는 데이터. 첫 학습은 v2가 된다."""
    w = tmp_path_factory.mktemp("safety")
    data = make_sample(str(w / "safe.csv"))
    root = w / "safe_ml_data"
    (root / "v1").mkdir(parents=True)
    (root / "v1" / "내 메모.txt").write_text("사용자 파일", encoding="utf-8")
    first = run_tool("ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET, "--desc", "safety")
    return {"dir": w, "data": data, "root": root, "first": first}


def test_incomplete_version_folder_is_kept(work):
    first, root = work["first"], work["root"]
    assert (root / "v1" / "내 메모.txt").exists()                    # not deleted
    assert first["version"] == "v2" and first["incomplete_versions"] == ["v1"]
    assert first["verified"] is True and first["agent_directive"].startswith("VERIFIED:")
    assert [r["version"] for r in _ledger(root)] == ["v2"]


def test_existing_version_refused_then_overwrite(work):
    root, data = work["root"], work["data"]
    before = (root / "v2" / "leaderboard.json").read_bytes()
    proc, payload = run_raw("ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET, "-v", "v2")
    assert proc.returncode == 1 and payload["agent_directive"].startswith("STOP_RETRY:")
    assert "[추천] --version v3" in payload["message"] and "--overwrite" in payload["message"]
    assert (root / "v2" / "leaderboard.json").read_bytes() == before
    assert len(_ledger(root)) == 1

    again = run_tool("ml_train.py", "-f", data, "-x", "반응온도,촉매비율", "-y", TARGET, "-v", "v2", "--overwrite")
    assert again["version"] == "v2" and again["verified"] is True
    meta = json.loads((root / "v2" / f"model_{TARGET}.json").read_text(encoding="utf-8"))
    assert meta["feature_count"] == 2                                 # replaced only with --overwrite
    assert not any(p.name.startswith((".bak_", ".tmp_")) for p in root.iterdir())
    # 2.9.46: one record per version (2.9.45 kept the replaced v2 record too: two 'v2' lines, and the new v2 was
    # compared with the old one for delta_score / is_best)
    ledger = _ledger(root)
    assert [r["version"] for r in ledger] == ["v2"], ledger
    assert ledger[0]["mean_val_score"] == again["mean_val_score"] and ledger[0]["delta_score"] == "base"
    assert ledger[0]["is_best"] is True
    hist = run_tool("ml_train.py", "-f", data, "--history")
    assert hist["record_count"] == 1 and hist["best_version"] == "v2"
    synced = json.loads((root / "history.json").read_text(encoding="utf-8"))
    assert [r["version"] for r in synced] == ["v2"]


def _train_in_process(monkeypatch, capsys, data, *extra):
    import ml_train
    monkeypatch.setattr(sys, "argv", ["ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET, *extra])
    with pytest.raises(SystemExit) as ex:
        ml_train.main()
    out = [x for x in capsys.readouterr().out.splitlines() if x.strip()]
    return ex.value.code, json.loads(out[-1])


def test_failed_commit_leaves_no_ledger_record(tmp_path, monkeypatch, capsys):
    import ml_train
    data = make_sample(str(tmp_path / "c.csv"))

    def broken_commit(self):
        raise PermissionError("rename blocked")
    monkeypatch.setattr(ml_train.TransactionGuard, "commit", broken_commit)
    code, payload = _train_in_process(monkeypatch, capsys, data)
    root = tmp_path / "c_ml_data"
    assert code == 1 and payload["status"] == "error"
    assert _ledger(root) == [] and not (root / "v1").exists()
    assert not any(p.name.startswith(".tmp_") for p in root.iterdir())


def test_failed_ledger_write_rolls_back_and_restores_old_version(work, tmp_path, monkeypatch, capsys):
    import ml_train
    w = tmp_path / "copy"
    shutil.copytree(work["dir"], w)
    data, root = str(w / "safe.csv"), w / "safe_ml_data"
    before = (root / "v2" / "leaderboard.json").read_bytes()

    def broken_ledger(*a, **k):
        raise OSError("disk full")
    monkeypatch.setattr(ml_train, "append_history_ledger", broken_ledger)
    code, payload = _train_in_process(monkeypatch, capsys, data, "-v", "v2", "--overwrite")
    assert code == 1 and payload["status"] == "error"
    assert (root / "v2" / "leaderboard.json").read_bytes() == before   # old version restored
    assert not any(p.name.startswith((".bak_", ".tmp_")) for p in root.iterdir())


def test_interrupt_after_ledger_line_keeps_version(work, tmp_path, monkeypatch, capsys):
    """history.jsonl에 버전이 적힌 뒤 history.json 동기화 중에 중단돼도 그 버전을 되돌리지 않는다."""
    import ml_train
    w = tmp_path / "copy"
    shutil.copytree(work["dir"], w)
    data, root = str(w / "safe.csv"), w / "safe_ml_data"
    real = ml_train.atomic_write_text

    def interrupt_json_sync(path, *a, **k):
        if os.path.basename(path) == "history.json":
            raise KeyboardInterrupt
        return real(path, *a, **k)
    monkeypatch.setattr(ml_train, "atomic_write_text", interrupt_json_sync)
    n_before = len(_ledger(root))
    code, payload = _train_in_process(monkeypatch, capsys, data)
    assert code == 130
    recorded = _ledger(root)
    assert len(recorded) == n_before + 1
    new_version = recorded[-1]["version"]
    assert (root / new_version / "leaderboard.json").exists()          # ledger and folders agree


def test_leftover_backup_blocks_overwrite(work, tmp_path, monkeypatch, capsys):
    """남은 .bak_vX는 유일한 정상 사본일 수 있다: --overwrite는 지우지 않고 멈춘다."""
    w = tmp_path / "copy"
    shutil.copytree(work["dir"], w)
    data, root = str(w / "safe.csv"), w / "safe_ml_data"
    shutil.copytree(root / "v2", root / ".bak_v2")
    marker = root / ".bak_v2" / "leaderboard.json"
    before = marker.read_bytes()
    code, payload = _train_in_process(monkeypatch, capsys, data, "-v", "v2", "--overwrite")
    assert code == 1 and payload["agent_directive"].startswith("STOP_RETRY")
    assert ".bak_v2" in payload["message"] and marker.read_bytes() == before


def test_parked_backup_restored_on_next_run(tmp_path):
    import ml_train
    root = tmp_path / "d_ml_data"
    (root / ".bak_v3").mkdir(parents=True)
    (root / ".bak_v3" / "leaderboard.json").write_text("[]", encoding="utf-8")
    (root / ".bak_v3" / "model_dataset.csv").write_text("a\n1\n", encoding="utf-8")
    g = ml_train.TransactionGuard(str(root))
    assert g.setup() == "v4"
    assert (root / "v3" / "leaderboard.json").exists() and not (root / ".bak_v3").exists()
    g.rollback()


def test_verify_version_dir_detects_broken_model(work, tmp_path):
    import ml_train
    vd = tmp_path / "v2"
    shutil.copytree(work["root"] / "v2", vd)
    ledger = str(work["root"] / "history.jsonl")
    version = _ledger(work["root"])[-1]["version"]
    assert ml_train.verify_version_dir(str(vd), [TARGET], ledger, version) == []
    (vd / f"model_{TARGET}.joblib").write_bytes(b"broken")
    problems = ml_train.verify_version_dir(str(vd), [TARGET], ledger, version)
    assert any("joblib" in p for p in problems)


# ---- 결과 파일: 덮어쓰지 않고 새 이름, 잠김 복구 ------------------------------------------------------------


def test_predict_out_is_not_overwritten(work):
    vd = work["first"]["version_dir"]
    out = work["dir"] / "pred.csv"
    out.write_text("사용자 결과\n", encoding="cp949")
    p = run_tool("ml_predict.py", "-m", vd, "-f", work["data"], "--out", str(out))
    assert os.path.normcase(p["output_file"]) != os.path.normcase(str(out))
    assert os.path.basename(p["output_file"]).startswith("pred_") and os.path.exists(p["output_file"])
    assert out.read_text(encoding="cp949") == "사용자 결과\n"
    assert any("pred.csv 이미 있음" in x for x in p["recovered"])

    p2 = run_tool("ml_predict.py", "-m", vd, "-f", work["data"], "--out", str(out), "--overwrite")
    assert os.path.normcase(p2["output_file"]) == os.path.normcase(str(out)) and "recovered" not in p2
    assert out.read_text(encoding="cp949") != "사용자 결과\n"


def test_report_default_name_is_not_overwritten(work):
    vd = work["first"]["version_dir"]
    r1 = run_tool("ml_report.py", "-m", vd)
    first_bytes = open(r1["report_file"], "rb").read()
    r2 = run_tool("ml_report.py", "-m", vd)
    assert os.path.normcase(r2["report_file"]) != os.path.normcase(r1["report_file"])
    assert open(r1["report_file"], "rb").read() == first_bytes and any("이미 있음" in x for x in r2["recovered"])
    r3 = run_tool("ml_report.py", "-m", vd, "--overwrite")
    assert os.path.normcase(r3["report_file"]) == os.path.normcase(r1["report_file"])


def test_timestamp_suffix_is_not_doubled(tmp_path):
    p = tmp_path / "a_pred_20260101_101010.csv"
    alt = mp._timestamped(str(p))
    stem = os.path.splitext(os.path.basename(alt))[0]
    assert stem.startswith("a_pred_") and stem.count("_") == 3       # a_pred_<date>_<time>


def test_lock_fallback_reports_recovered(tmp_path, monkeypatch):
    monkeypatch.setattr(mp, "_RECOVERED", [])
    calls = []

    def writer(p):
        calls.append(p)
        if len(calls) == 1:
            raise PermissionError("locked")
        open(p, "w").close()
    saved = mp.save_with_lock_fallback(str(tmp_path / "r.docx"), writer, what="Word 보고서")
    assert saved != str(tmp_path / "r.docx") and os.path.exists(saved)
    assert len(mp._RECOVERED) == 1 and "r.docx 잠김" in mp._RECOVERED[0]


# ---- 실기 러너 tools/run_win_tests.py (Windows 전용 모듈은 가짜로 채움) --------


def _load_runner():
    for m in ("pythoncom", "pywintypes", "win32com", "win32com.client"):
        sys.modules.setdefault(m, types.ModuleType(m))
    sys.modules["win32com"].client = sys.modules["win32com.client"]
    spec = importlib.util.spec_from_file_location("run_win_tests", os.path.join(DEV, "tools", "run_win_tests.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_runner_does_not_delete_user_folder(tmp_path):
    rwt = _load_runner()
    user_dir = tmp_path / "내 문서"
    user_dir.mkdir()
    (user_dir / "중요.xlsx").write_text("x", encoding="utf-8")
    with pytest.raises(RuntimeError, match="지우지 않았습니다"):
        rwt.prepare_test_dir(str(user_dir))
    assert (user_dir / "중요.xlsx").exists()

    new_dir = tmp_path / "ml 테스트"
    rwt.prepare_test_dir(str(new_dir))                                # new path: created + marker
    (new_dir / "old_result.txt").write_text("x", encoding="utf-8")
    rwt.prepare_test_dir(str(new_dir))                                # runner's own folder: emptied
    assert sorted(os.listdir(new_dir)) == [rwt.DIR_MARKER]


def test_testlog_contains_criteria_calls_and_traceback(tmp_path):
    rwt = _load_runner()
    ctx = rwt.Ctx(type("Args", (), {"dir": str(tmp_path)})())

    def m13():
        d, c, o, e = rwt.run_tool("ml_train.py", "--없는옵션")
        rwt.record(ctx, "M13", "잘못된 옵션", False, f"exit={c}", d, c, e)

    def m6():
        raise RuntimeError("Word COM 실패 (테스트)")

    rwt.run_test(ctx, "M13", "오류 계약", m13)
    rwt.run_test(ctx, "M6", "보고서를 Word로", m6)
    path = rwt.write_testlog(ctx, str(tmp_path / "logs"), "테스트PC", "20260101_000000", fatal=None)
    text = open(path, encoding="utf-8").read()
    assert os.path.basename(path) == "testlog_테스트PC_20260101_000000.txt"
    assert "결과: ERROR=1, FAIL=1" in text
    assert rwt.EXPECT["M13"] in text and rwt.EXPECT["M6"] in text          # pass criteria
    assert "$ python ml_train.py --없는옵션" in text and "exit=2" in text  # recorded call
    assert '"status": "error"' in text                                     # full stdout JSON
    assert "RuntimeError: Word COM 실패" in text and "Traceback" in text   # exception trace
    assert "[환경]" in text and "scikit-learn" in text and "[전체 호출 기록]" in text


def test_only_rejects_unknown_ids(monkeypatch, capsys):
    import pytest
    rwt = _load_runner()
    assert rwt.TEST_IDS == tuple(f"M{i}" for i in range(1, 15))
    assert rwt.parse_only(None) == (None, [])
    assert rwt.parse_only("m1, M4") == ({"M1", "M4"}, [])
    assert rwt.parse_only("M1,M15,X") == (None, ["X", "M15"])
    assert rwt.parse_only(",") == (None, ["(비어 있음)"])
    monkeypatch.setattr(sys, "argv", ["run_win_tests.py", "--only", "M1,M99"])
    with pytest.raises(SystemExit) as ex:
        rwt.main()
    assert ex.value.code == 2 and "M99" in capsys.readouterr().err


def test_decode_console_reads_locale_text():
    import locale
    rwt = _load_runner()
    assert rwt.decode_console(b"Active code page: 65001") == "Active code page: 65001"
    text = "활성 코드 페이지: 949"
    try:                                     # the console code page of this PC (Korean Windows: cp949)
        raw = text.encode("oem" if os.name == "nt" else locale.getpreferredencoding(False))
    except (LookupError, UnicodeEncodeError):
        return
    assert rwt.decode_console(raw) == text
