# -*- coding: utf-8 -*-
"""apply_skills.py: SKILL.md 2개와 전역 AGENTS.md를 OpenCode 설정 폴더로 복사, 바뀐 것만, 덮어쓰기 전 보관, --check,
run_all_tests.py diag의 스킬 확인."""
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
    old = cfg / "skills" / "office-cli" / "SKILL.md"
    old.write_text('---\nname: office-cli\nmetadata:\n  version: "1.0.0"\n---\n옛 내용', encoding="utf-8")
    (cfg / "skills" / "ml-kit" / "ml_train.py").write_text("x", encoding="utf-8")
    capsys.readouterr()
    assert aps.main(["--config-dir", str(cfg), "--check"]) == 1
    out = capsys.readouterr().out
    assert "office-cli 스킬: 다름 (1.0.0 -> " in out
    assert "ml_train.py" not in out                     # 스킬 폴더의 .py는 알리지 않음 (사용자 결정)
    assert aps.main(["--config-dir", str(cfg), "--test"]) == 0
    kept = list((dev / "work").glob("skill_backup_*/skills/office-cli/SKILL.md"))
    assert len(kept) == 1 and "옛 내용" in kept[0].read_text(encoding="utf-8")
    assert old.read_bytes() == open(os.path.join(aps.CLI_DIR, "office-cli", "SKILL.md"), "rb").read()
    assert (cfg / "skills" / "office-cli-test" / "SKILL.md").is_file()


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
