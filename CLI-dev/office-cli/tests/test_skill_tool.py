"""skill.py: 사용자가 만든 OpenCode 스킬 등록·목록·점검·지우기 (머리말은 도구가 붙임, 형식·크기·문자 확인, 다시 읽어 확인)."""
import os

import pytest

import skill
from office_helpers import recommended_argv, run_main, split_command

DESC = "주간 실적 엑셀을 부서별로 집계해 보고서 표로 만든다. 주간 보고 정리 요청에 사용."
BODY = "# 주간 보고 정리\n\n## 순서\n1. `python \"C:\\CLI\\data\\scripts\\weekly.py\" \"<엑셀 파일>\"` 을 1번 실행한다.\n2. 결과 파일 경로를 알린다.\n"


@pytest.fixture
def cfg(tmp_path, monkeypatch):
    monkeypatch.setenv("OFFICE_CLI_OPENCODE_DIR", str(tmp_path / "cfg"))
    body = tmp_path / "weekly_skill.md"
    body.write_text(BODY, encoding="utf-8")
    return tmp_path / "cfg" / "skills", str(body)


def _save(capsys, body, *extra, name="weekly-report", desc=DESC):
    return run_main(skill, capsys, "save", "--name", name, "--description", desc, "--file", body, *extra)


def test_save_adds_front_matter_and_verifies(cfg, capsys):
    root, body = cfg
    code, out = _save(capsys, body)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["updated"] is False, out
    text = (root / "weekly-report" / "SKILL.md").read_text(encoding="utf-8")
    assert text.startswith(f"---\nname: weekly-report\ndescription: {DESC}\n---\n\n# 주간 보고 정리"), text[:120]
    assert out["skill_file"] == str(root / "weekly-report" / "SKILL.md") and "다시 시작" in out["preview"]
    assert out["bytes"] == len(text.encode("utf-8")) <= out["limit_bytes"] == 8192


def test_save_existing_asks_and_overwrite_command_runs(cfg, capsys):
    root, body = cfg
    _save(capsys, body)
    code, out = _save(capsys, body)
    assert code == 1 and out["error_type"] == "needs_user_decision" and "[후보] 1)" in out["message"], out
    cmd = out["confirm_commands"][0]
    assert cmd["destructive"] is True
    argv = split_command(cmd["command"])[2:]
    assert "--overwrite" in argv
    code, out = run_main(skill, capsys, *argv)
    assert code == 0 and out["updated"] is True and "바꿨습니다" in out["preview"], out


def test_name_is_fixed_by_recommend_or_asked(cfg, capsys):
    root, body = cfg
    code, out = _save(capsys, body, name="Weekly_Report")
    assert code == 1 and "[추천]" in out["message"] and not root.exists(), out
    argv = recommended_argv(out["message"], "skill")
    assert argv[argv.index("--name") + 1] == "weekly-report"
    code, out = run_main(skill, capsys, *argv)
    assert code == 0 and (root / "weekly-report" / "SKILL.md").is_file(), out
    code, out = _save(capsys, body, name="주간보고")                 # 한글 이름: 고칠 수 없어 사용자에게 물음
    assert code == 1 and "[추천]" not in out["message"] and "영문 이름" in out["message"], out
    code, out = _save(capsys, body, name="office-excel")             # 기본 스킬 이름
    assert code == 1 and "기본 스킬" in out["message"], out


def test_problems_stop_before_writing(cfg, capsys, tmp_path):
    root, body = cfg
    code, out = _save(capsys, body, desc="짧음")
    assert code == 1 and "description이 너무 짧습니다" in out["message"] and not root.exists(), out
    big = tmp_path / "big.md"
    big.write_text("# 큰 스킬\n\n1. `x`\n" + "가" * 4000, encoding="utf-8")
    code, out = _save(capsys, str(big))
    assert code == 1 and "너무 큽니다" in out["message"] and "8192" in out["message"], out
    emoji = tmp_path / "e.md"
    emoji.write_text("# 제목\n\n1. `x` 완료 \u2705 \U0001F600 다음\n", encoding="utf-8")
    code, out = _save(capsys, str(emoji))
    assert code == 1 and "U+2705" in out["message"] and "U+1F600" in out["message"], out
    assert "같은 명령을 1번 다시 실행" in out["agent_directive"]


def test_body_front_matter_is_replaced_and_warned(cfg, capsys, tmp_path):
    root, _ = cfg
    b = tmp_path / "with_front.md"
    b.write_text("---\nname: wrong\ndescription: x\n---\n\n" + BODY + "\npip install pandas\n", encoding="utf-8")
    code, out = _save(capsys, str(b))
    text = (root / "weekly-report" / "SKILL.md").read_text(encoding="utf-8")
    assert code == 0 and text.count("---\n") == 2 and "name: wrong" not in text, text[:200]
    assert any("머리말" in w for w in out["warnings"]) and any("폐쇄망" in w for w in out["warnings"]), out["warnings"]


def test_list_check_delete(cfg, capsys, tmp_path, monkeypatch):
    root, body = cfg
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path / "data"))
    code, out = run_main(skill, capsys, "list")
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"], out
    _save(capsys, body)
    base = root / "office-excel"
    base.mkdir()
    (base / "SKILL.md").write_text("---\nname: office-excel\ndescription: 엑셀 읽기와 쓰기에 쓰는 기본 스킬이다.\n---\n\n# x\n1. `a`\n", encoding="utf-8")
    code, out = run_main(skill, capsys, "list")
    kinds = {x["name"]: x["kind"] for x in out["skills"]}
    assert out["count"] == 2 and out["user_count"] == 1 and kinds == {"office-excel": "기본", "weekly-report": "사용자"}, out
    code, out = run_main(skill, capsys, "check", "--name", "weekly-report")
    assert code == 0 and out["count"] == 0 and "문제 없음" in out["preview"], out
    (root / "weekly-report" / "SKILL.md").write_text("머리말 없는 글\n", encoding="utf-8")
    code, out = run_main(skill, capsys, "check", "--name", "weekly-report")
    assert code == 0 and out["count"] >= 1 and any("머리말" in p for p in out["problems"]), out
    code, out = run_main(skill, capsys, "check", "--name", "weekly")
    assert code == 1 and "weekly-report" in out["message"], out
    code, out = run_main(skill, capsys, "delete", "--name", "office-excel")
    assert code == 1 and out["error_type"] == "not_supported" and base.is_dir(), out
    code, out = run_main(skill, capsys, "delete", "--name", "weekly-report", "--dry-run")
    assert code == 0 and out["dry_run"] is True and (root / "weekly-report").is_dir(), out
    argv = split_command(out["confirm_commands"][0]["command"])[2:]
    code, out = run_main(skill, capsys, *argv)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and not (root / "weekly-report").exists(), out
    assert os.path.isfile(os.path.join(out["backup"], "SKILL.md"))          # 지운 내용은 보관
