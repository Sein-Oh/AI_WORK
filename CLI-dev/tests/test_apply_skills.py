# -*- coding: utf-8 -*-
"""apply_skills.py: 스킬(ml-kit, office-cli 작업별 4개)과 전역 AGENTS.md를 OpenCode 설정 폴더로 복사, 바뀐 것만, 덮어쓰기 전 보관,
옛 통합 스킬 폴더(skills/office-cli)는 보관 후 지움, --check, run_all_tests.py diag의 스킬 확인."""
import importlib.util
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(HERE)


def _load(name, file):
    spec = importlib.util.spec_from_file_location(name, os.path.join(DEV, file))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


aps = _load("apply_skills_t", "apply_skills.py")


def _iso(monkeypatch, tmp_path):
    """원본은 저장소 그대로, 보관 폴더(work)만 tmp_path로."""
    dev = tmp_path / "CLI-dev"
    dev.mkdir()
    real_here = aps.HERE
    monkeypatch.setattr(aps, "HERE", str(dev))
    for src_name, rel in (("opencode", os.path.join("opencode", "AGENTS.md")),
                          ("office-cli", os.path.join("office-cli", "tools", "office-cli-test", "SKILL.md"))):
        dst = dev / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(open(os.path.join(real_here, rel), "rb").read())
    return dev, tmp_path / "opencode_cfg"


def test_sources_exist_in_repo_layout():
    for name, src, rel in aps.targets(include_test=True):
        assert os.path.isfile(src), (name, src)
    assert aps.default_config_dir().endswith(os.path.join(".config", "opencode"))
    on_disk = sorted(os.listdir(os.path.join(aps.CLI_DIR, "office-cli", "skills")))
    assert sorted(aps.OFFICE_SKILLS) == on_disk                          # 새 스킬 폴더를 만들면 apply_skills에도 올림
    assert {os.path.dirname(rel) for _n, _s, rel in aps.targets()} >= {os.path.join("skills", n) for n in on_disk}


def test_apply_copies_new_then_reports_same(monkeypatch, tmp_path, capsys):
    dev, cfg = _iso(monkeypatch, tmp_path)
    assert aps.main(["--config-dir", str(cfg)]) == 0
    assert (cfg / "skills" / "ml-kit" / "SKILL.md").read_bytes() == open(os.path.join(aps.CLI_DIR, "ml-kit", "SKILL.md"), "rb").read()
    assert (cfg / "AGENTS.md").is_file() and not (cfg / "skills" / "office-cli-test").exists()   # 보조 스킬은 --test일 때만
    assert not (dev / "work").exists()                                                   # 덮어쓴 것이 없으면 보관도 없음
    assert aps.main(["--config-dir", str(cfg), "--check"]) == 0
    assert "다름" not in capsys.readouterr().out


def test_changed_file_is_backed_up_and_check_fails_before_apply(monkeypatch, tmp_path, capsys):
    dev, cfg = _iso(monkeypatch, tmp_path)
    aps.main(["--config-dir", str(cfg)])
    old = cfg / "skills" / "office-excel" / "SKILL.md"
    old.write_text('---\nname: office-excel\nmetadata:\n  version: "1.0.0"\n---\n옛 내용', encoding="utf-8")
    (cfg / "skills" / "ml-kit" / "ml_train.py").write_text("x", encoding="utf-8")
    capsys.readouterr()
    assert aps.main(["--config-dir", str(cfg), "--check"]) == 1
    out = capsys.readouterr().out
    assert "office-excel 스킬: 다름 (1.0.0 -> " in out
    assert "ml_train.py" not in out                     # 스킬 폴더의 .py는 알리지 않음 (사용자 결정)
    assert aps.main(["--config-dir", str(cfg), "--test"]) == 0
    kept = list((dev / "work").glob("skill_backup_*/skills/office-excel/SKILL.md"))
    assert len(kept) == 1 and "옛 내용" in kept[0].read_text(encoding="utf-8")
    assert old.read_bytes() == open(os.path.join(aps.CLI_DIR, "office-cli", "skills", "office-excel", "SKILL.md"), "rb").read()
    assert (cfg / "skills" / "office-cli-test" / "SKILL.md").is_file()


def test_old_office_cli_skill_is_backed_up_and_removed(monkeypatch, tmp_path, capsys):
    """1.4.61까지의 통합 스킬 폴더가 남으면 OpenCode가 옛 스킬도 고름: --check와 diag는 실패, 적용하면 보관 후 지움."""
    dev, cfg = _iso(monkeypatch, tmp_path)
    aps.main(["--config-dir", str(cfg)])
    old = cfg / "skills" / "office-cli" / "SKILL.md"
    old.parent.mkdir(parents=True)
    old.write_text('---\nname: office-cli\nmetadata:\n  version: "1.4.61"\n---\n옛 통합 스킬', encoding="utf-8")
    capsys.readouterr()
    assert aps.main(["--config-dir", str(cfg), "--check"]) == 1
    assert "옛 office-cli 스킬: 지울 것" in capsys.readouterr().out and old.is_file()       # --check는 지우지 않음
    assert aps.main(["--config-dir", str(cfg)]) == 0
    assert not old.parent.exists()
    kept = list((dev / "work").glob("skill_backup_*/skills/office-cli/SKILL.md"))
    assert len(kept) == 1 and "옛 통합 스킬" in kept[0].read_text(encoding="utf-8")
    assert "보관 후 지움" in capsys.readouterr().out
    assert aps.main(["--config-dir", str(cfg), "--check"]) == 0
    for name in aps.OFFICE_SKILLS:
        assert (cfg / "skills" / name / "SKILL.md").is_file()


def test_old_office_cli_skill_link_removes_only_link(monkeypatch, tmp_path, capsys):
    """옛 skills\\office-cli가 링크(정션)이면 링크만 지우고 가리키는 폴더(사용자 것일 수 있음)는 그대로."""
    dev, cfg = _iso(monkeypatch, tmp_path)
    aps.main(["--config-dir", str(cfg)])
    real = tmp_path / "내 스킬 원본"
    real.mkdir()
    (real / "SKILL.md").write_text("---\nname: office-cli\n---\n원본", encoding="utf-8")
    link = cfg / "skills" / "office-cli"
    try:
        os.symlink(real, link, target_is_directory=True)
    except (OSError, NotImplementedError):
        import pytest
        pytest.skip("이 환경에서는 링크를 만들 수 없음")
    capsys.readouterr()
    assert aps.main(["--config-dir", str(cfg)]) == 0
    out = capsys.readouterr().out
    assert not os.path.lexists(link) and (real / "SKILL.md").is_file(), out
    assert "링크만 지움" in out and str(real) in out
    assert not list((dev / "work").glob("skill_backup_*/skills/office-cli"))


def test_missing_source_is_error(monkeypatch, tmp_path, capsys):
    dev, cfg = _iso(monkeypatch, tmp_path)
    os.remove(dev / "opencode" / "AGENTS.md")
    assert aps.main(["--config-dir", str(cfg)]) == 1
    assert "원본이 없습니다" in capsys.readouterr().out and not (cfg / "AGENTS.md").exists()


def test_diag_fails_when_installed_skill_differs(tmp_path):
    rat = _load("run_all_tests_s", "run_all_tests.py")
    cfg = tmp_path / "cfg"
    lines = []
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(cfg))                       # 설정 폴더 없음: 확인 안 함
    assert st.failed == 0 and st.passed == 0 and "확인 안 함" in lines[-1]
    aps.apply(str(cfg))
    rat.diag_skills(st, lines, config_dir=str(cfg))
    assert st.failed == 0 and st.passed == 1
    (cfg / "AGENTS.md").write_text("옛 지침", encoding="utf-8")
    rat.diag_skills(st, lines, config_dir=str(cfg))
    assert st.failed == 1 and "apply_skills.py" in st.failures[0] and "전역 지침 다름" in st.failures[0]
    cfg2 = tmp_path / "cfg2"                                                 # 새 폴더라 보관(work)이 생기지 않음
    aps.apply(str(cfg2))
    (cfg2 / "skills" / "office-cli").mkdir()                                 # 옛 통합 스킬 폴더가 남음
    (cfg2 / "skills" / "office-cli" / "SKILL.md").write_text("옛 스킬", encoding="utf-8")
    rat.diag_skills(st, lines, config_dir=str(cfg2))
    assert st.failed == 2 and "옛 office-cli 스킬 지울 것" in st.failures[1]
