"""skill.py: 사용자가 만든 OpenCode 스킬 등록·목록·점검·지우기 (머리말은 도구가 붙임, 형식·크기·문자 확인, 다시 읽어 확인)."""
import os

import pytest

import skill
from office_helpers import recommended_argv, run_main, split_command

DESC = "주간 실적 엑셀을 부서별로 집계해 보고서 표로 만든다. 주간 보고 정리 요청에 사용."
BODY = "# 주간 보고 정리\n\n## 순서\n1. `python \"C:\\AIKIT\\data\\scripts\\weekly.py\" \"<엑셀 파일>\"` 을 1번 실행한다.\n2. 결과 파일 경로를 알린다.\n"


@pytest.fixture
def cfg(tmp_path, monkeypatch):
    """OpenCode 설정 폴더(OPENCODE_CONFIG_DIR)를 tmp_path로 (실제 C:\\AIKIT\\skills를 건드리지 않음). 기본·사용자 스킬 모두 <설정 폴더>\\skills."""
    monkeypatch.setenv("OPENCODE_CONFIG_DIR", str(tmp_path / "cfg"))
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
    code, out = _save(capsys, body, name="excel-kit")             # 기본 스킬 이름
    assert code == 1 and "기본 스킬" in out["message"], out
    code, out = _save(capsys, body, name="office-excel")          # 1.6.7에서 바꾼 기본 스킬의 옛 이름: setup이 지우므로 못 만듦
    assert code == 1 and "기본 스킬" in out["message"] and not (root / "office-excel").exists(), out


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
    base = root / "excel-kit"                                         # 기본 스킬도 같은 skills 폴더 (이름으로 기본/사용자를 나눔, 1.6.6)
    base.mkdir(parents=True)
    (base / "SKILL.md").write_text("---\nname: excel-kit\ndescription: 엑셀 읽기와 쓰기에 쓰는 기본 스킬이다.\n---\n\n# x\n1. `a`\n", encoding="utf-8")
    code, out = run_main(skill, capsys, "list")
    kinds = {x["name"]: x["kind"] for x in out["skills"]}
    assert out["count"] == 2 and out["user_count"] == 1 and kinds == {"excel-kit": "기본", "weekly-report": "사용자"}, out
    code, out = run_main(skill, capsys, "check", "--name", "weekly-report")
    assert code == 0 and out["count"] == 0 and "문제 없음" in out["preview"], out
    (root / "weekly-report" / "SKILL.md").write_text("머리말 없는 글\n", encoding="utf-8")
    code, out = run_main(skill, capsys, "check", "--name", "weekly-report")
    assert code == 0 and out["count"] >= 1 and any("머리말" in p for p in out["problems"]), out
    code, out = run_main(skill, capsys, "check", "--name", "weekly")
    assert code == 1 and "weekly-report" in out["message"], out
    code, out = run_main(skill, capsys, "delete", "--name", "excel-kit")
    assert code == 1 and out["error_type"] == "not_supported" and base.is_dir(), out
    code, out = run_main(skill, capsys, "delete", "--name", "weekly-report", "--dry-run")
    assert code == 0 and out["dry_run"] is True and (root / "weekly-report").is_dir(), out
    argv = split_command(out["confirm_commands"][0]["command"])[2:]
    code, out = run_main(skill, capsys, *argv)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and not (root / "weekly-report").exists(), out
    assert os.path.isfile(os.path.join(out["backup"], "SKILL.md"))          # 지운 내용은 보관


def test_project_skill_goes_to_project_opencode_folder(cfg, capsys, tmp_path, monkeypatch):
    """1.6.5: --project <폴더> -> <폴더>\\.opencode\\skills (OpenCode 2.2.5가 그 폴더에서 연 세션에만 더함, 집 PC 2026-11-12 확인).
    없는 폴더는 만들지 않고 STOP_RETRY (오타 경로에 .opencode가 생기지 않게)."""
    root, body = cfg
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path / "data"))
    proj = tmp_path / "시험A"
    proj.mkdir()
    code, out = _save(capsys, body, "--project", str(proj))
    path = proj / ".opencode" / "skills" / "weekly-report" / "SKILL.md"
    assert code == 0 and out["skill_file"] == str(path) and out["kind"] == "프로젝트" and path.is_file(), out
    assert str(proj) in out["preview"] and "에서만" in out["preview"] and not root.exists()
    code, out = run_main(skill, capsys, "list")                            # --project 없이는 보이지 않음
    assert out["count"] == 0, out
    code, out = run_main(skill, capsys, "list", "--project", str(proj))
    assert out["count"] == 1 and out["skills"][0]["kind"] == "프로젝트", out
    code, out = run_main(skill, capsys, "check", "--name", "weekly-report", "--project", str(proj))
    assert code == 0 and out["skill_file"] == str(path), out
    code, out = run_main(skill, capsys, "delete", "--name", "weekly-report", "--project", str(proj))
    assert code == 0 and not path.parent.exists(), out
    code, out = _save(capsys, body, "--project", str(tmp_path / "없는 폴더"))
    assert code == 1 and "--project 폴더가 없습니다" in out["message"] and not (tmp_path / "없는 폴더").exists(), out


def test_default_folder_is_the_aikit_skills_folder(monkeypatch):
    """1.6.6: OPENCODE_CONFIG_DIR가 없으면 도구 폴더의 부모(C:\\AIKIT)가 설정 폴더, 스킬은 그 skills (기본 스킬과 사용자 스킬이 함께).
    저장소에서는 그 폴더에 기본 스킬 12개가 있고 모두 '기본'으로 보임."""
    monkeypatch.delenv("OPENCODE_CONFIG_DIR", raising=False)
    kit = os.path.dirname(os.path.dirname(os.path.abspath(skill.__file__)))
    assert skill.config_dir() == kit and skill.skills_dir() == os.path.join(kit, "skills")
    base = {x["name"] for x in skill._installed() if x["kind"] == "기본"}
    assert base == set(skill.BASE_SKILLS) - {"office-cli-test"}, base
