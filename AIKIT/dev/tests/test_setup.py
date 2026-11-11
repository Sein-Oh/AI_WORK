# -*- coding: utf-8 -*-
"""AIKIT/setup.py (배포 설치): OpenCode 설정 폴더를 C:\\AIKIT으로 (사용자 환경 변수 OPENCODE_CONFIG_DIR, 가짜 winreg),
이전 구성(C:\\AIKIT\\opencode, 도구 폴더 안의 skills) 정리,
옛 전역 설정 폴더(%USERPROFILE%\\.config\\opencode) 정리(우리 스킬·지침은 보관 후 지움, 사용자 스킬·opencode.jsonc는 C:\\AIKIT으로),
시스템 PATH 등록(가짜 winreg), 동작 확인, opencode.json + opencode.jsonc 합친 설정 점검, --check, setup.bat,
run_all_tests.py diag의 환경 변수·PATH·옛 복사본 확인."""
import importlib.util
import json
import os
import shutil

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(HERE)
CLI = os.path.dirname(DEV)                  # AIKIT (dev의 부모)
REPO = os.path.dirname(CLI)


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


su = _load("cli_setup_t", os.path.join(CLI, "setup.py"))


class _Key:
    def __init__(self, root):
        self.root = root

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


class FakeReg:
    """winreg 대신: HKLM 시스템 환경 변수 Path 하나와 HKCU 사용자 환경 변수(user dict)."""
    HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER, KEY_READ, KEY_SET_VALUE, REG_SZ, REG_EXPAND_SZ = "HKLM", "HKCU", 1, 2, 1, 2

    def __init__(self, value=None, kind=2, writable=True, user=None):
        self.value, self.kind, self.writable, self.writes = value, kind, writable, []
        self.user = dict(user or {})

    def OpenKey(self, root, sub, reserved=0, access=1):
        assert (root, sub) in (("HKLM", su.ENV_KEY), ("HKCU", su.USER_ENV_KEY))
        if root == "HKLM" and access & self.KEY_SET_VALUE and not self.writable:
            raise PermissionError(5, "액세스가 거부되었습니다")
        return _Key(root)

    def QueryValueEx(self, key, name):
        if key.root == "HKCU":
            if name not in self.user:
                raise FileNotFoundError(2, "없음")
            return self.user[name], self.REG_SZ
        assert name == "Path"
        if self.value is None:
            raise FileNotFoundError(2, "없음")
        return self.value, self.kind

    def SetValueEx(self, key, name, reserved, kind, value):
        if key.root == "HKCU":
            self.user[name] = value
            return
        self.writes.append((name, kind, value))
        self.value, self.kind = value, kind


def _iso(monkeypatch, tmp_path, reg=None):
    """가짜 C:\\AIKIT(OpenCode가 읽는 파일만 저장소에서 복사)와 가짜 옛 전역 설정 폴더. 동작 확인·opencode 확인은 끔."""
    cli = tmp_path / "AIKIT"
    for _n, src in su.sources():
        dst = cli / os.path.relpath(src, CLI)
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(open(src, "rb").read())
    old = tmp_path / "old_cfg"
    monkeypatch.setattr(su, "HERE", str(cli))
    monkeypatch.setattr(su, "old_config_dir", lambda: str(old))
    monkeypatch.setattr(su, "step_verify", lambda rep: None)
    monkeypatch.setattr(su, "step_opencode_paths", lambda rep: None)
    monkeypatch.setattr(su, "_winreg", lambda: reg)
    monkeypatch.setattr(su, "broadcast_env_change", lambda: True)
    monkeypatch.delenv(su.CONFIG_ENV, raising=False)                    # step_env가 이 프로세스에 넣는 값을 시험 뒤 되돌림
    return cli, old


def _backups(cli, pattern):
    return list((cli / "data").glob("setup_backup_*/" + pattern))


def _skill(folder, name, text="본문"):
    path = folder / "skills" / name / "SKILL.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"---\nname: {name}\n---\n{text}", encoding="utf-8")
    return path


def test_sources_exist_in_repo_layout():
    """1.6.6: 설정 파일 없이 폴더 모양만으로: AIKIT/AGENTS.md(전역 지침), AIKIT/skills/<이름>/SKILL.md(기본 스킬 12개).
    opencode.json은 셸·권한 기본 설정만 (스킬 경로 없음)."""
    for name, src in su.sources():
        assert os.path.isfile(src), (name, src)
    on_disk = set(os.listdir(os.path.join(CLI, "skills")))
    assert set(su.ML_SKILLS + su.OFFICE_SKILLS) <= on_disk                # 새 스킬 폴더를 만들면 setup.py에도 올림
    assert not os.path.exists(os.path.join(CLI, "opencode"))              # 이전 구성의 설정 폴더는 없음
    for tool in ("ml-kit", "office-cli"):
        assert not os.path.exists(os.path.join(CLI, tool, "skills")), tool
    cfg = json.load(open(os.path.join(CLI, "opencode.json"), encoding="utf-8"))
    assert set(cfg) - {"$schema"} <= set(su.OUR_KEYS) and "skills" not in cfg   # 사내 LLM 설정(provider, model)은 opencode.jsonc에만
    assert cfg["shell"] == "powershell" and cfg["autoupdate"] is False
    assert cfg["permission"] == {"external_directory": "allow", "webfetch": "deny", "websearch": "deny"}
    gitignore = os.path.join(REPO, ".gitignore")
    if os.path.isfile(gitignore):        # 저장소에서만 (반입한 PC에는 .gitignore가 없음)
        ignore = open(gitignore, encoding="utf-8").read().splitlines()
        assert {"AIKIT/opencode.jsonc", "AIKIT/skills/*", "AIKIT/data/", "AIKIT/service.json"} <= set(ignore)
        assert {f"!AIKIT/skills/{n}/" for n in su.ML_SKILLS + su.OFFICE_SKILLS} <= set(ignore)    # 기본 스킬만 저장소에
        assert "AIKIT/opencode.json" not in ignore
    assert su.config_dir() == su.HERE and su.old_config_dir().endswith(os.path.join(".config", "opencode"))
    for dist in ("pandas", "scikit-learn", "openpyxl", "python-docx", "python-pptx", "pymupdf", "pywin32"):
        assert dist in dict(su.LIBS)


def test_apply_sets_user_env_then_check_passes(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows", user={su.CONFIG_ENV: r"D:\옛 설정"})
    cli, old = _iso(monkeypatch, tmp_path, reg)
    assert su.main(["--no-path"]) == 0
    assert reg.user[su.CONFIG_ENV] == str(cli) and os.environ[su.CONFIG_ENV] == str(cli)
    kept = _backups(cli, "user_env.txt")
    assert len(kept) == 1 and r"D:\옛 설정" in kept[0].read_text(encoding="utf-8")      # 쓰기 전 원래 값
    assert reg.writes == []                                               # --no-path: 시스템 PATH는 그대로
    assert (cli / "data" / "scripts").is_dir() and (cli / "work").is_dir()
    out = capsys.readouterr().out
    assert "[STATUS] 끝." in out and "옛 전역 설정 폴더에 정리할 것 없음" in out
    assert su.main(["--no-path", "--check"]) == 0
    out = capsys.readouterr().out
    assert "다름" not in out.replace("다름 0", "") and "[FAIL]" not in out and "사용자 환경 변수 OPENCODE_CONFIG_DIR = " in out


def test_check_reports_and_apply_tidies_old_global_config(monkeypatch, tmp_path, capsys):
    """2026-11-12 전 setup이 복사한 옛 전역 설정: 우리 스킬(지금 이름, 옛 통합 폴더)과 AGENTS.md는 보관 후 지우고,
    사용자가 만든 스킬은 C:\\AIKIT\\skills, opencode.jsonc는 C:\\AIKIT으로 옮김. --check는 바꾸지 않음."""
    reg = FakeReg(r"C:\Windows")
    cli, old = _iso(monkeypatch, tmp_path, reg)
    ours = _skill(old, "excel-kit", "옛 복사본")
    combined = _skill(old, "office-cli", "옛 통합 스킬")
    mine = _skill(old, "weekly-report", "내 스킬")
    (old / "AGENTS.md").write_text("옛 전역 지침", encoding="utf-8")
    (old / "opencode.jsonc").write_text('{"model": "corp/gemma"}', encoding="utf-8")
    (old / "node_modules").mkdir()                                        # OpenCode가 만든 것은 그대로
    assert su.main(["--no-path", "--check"]) == 1
    out = capsys.readouterr().out
    assert "OPENCODE_CONFIG_DIR: 없음" in out and su.CONFIG_ENV not in reg.user
    assert "옛 전역 설정 우리 스킬" in out and "옛 전역 설정 사용자 스킬" in out and "옛 전역 설정 opencode.jsonc" in out
    assert ours.is_file() and mine.is_file() and (old / "AGENTS.md").is_file()
    assert su.main(["--no-path"]) == 0
    out = capsys.readouterr().out
    assert not ours.parent.exists() and not combined.parent.exists() and not (old / "AGENTS.md").exists()
    assert (old / "node_modules").is_dir()
    assert (cli / "skills" / "weekly-report" / "SKILL.md").read_text(encoding="utf-8").endswith("내 스킬")
    assert not mine.exists() and "로 옮김" in out
    assert (cli / "opencode.jsonc").read_text(encoding="utf-8") == '{"model": "corp/gemma"}'
    assert not (old / "opencode.jsonc").exists()
    for rel in ("skills/excel-kit/SKILL.md", "skills/office-cli/SKILL.md", "skills/weekly-report/SKILL.md", "AGENTS.md", "opencode.jsonc"):
        assert len(_backups(cli, "old_config/" + rel)) == 1, rel           # 지우거나 옮긴 것은 모두 보관
    assert su.main(["--no-path", "--check"]) == 0


def test_old_jsonc_and_user_skill_do_not_replace_existing_ones(monkeypatch, tmp_path, capsys):
    """C:\\AIKIT에 이미 opencode.jsonc와 같은 이름의 사용자 스킬이 있으면 그것을 쓰고, 옛 것은 보관 후 지움."""
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    (cli / "opencode.jsonc").write_text('{"model": "새 것"}', encoding="utf-8")
    _skill(cli, "weekly-report", "새 스킬")
    _skill(old, "weekly-report", "옛 스킬")
    (old / "opencode.jsonc").write_text('{"model": "옛 것"}', encoding="utf-8")
    assert su.main(["--no-path"]) == 0
    out = capsys.readouterr().out
    assert '"새 것"' in (cli / "opencode.jsonc").read_text(encoding="utf-8")
    assert (cli / "skills" / "weekly-report" / "SKILL.md").read_text(encoding="utf-8").endswith("새 스킬")
    assert not (old / "opencode.jsonc").exists() and "이미 있는" in out
    assert "옛 것" in _backups(cli, "old_config/opencode.jsonc")[0].read_text(encoding="utf-8")


def test_old_skill_link_removes_only_link(monkeypatch, tmp_path, capsys):
    """옛 skills\\office-cli가 링크(정션)이면 링크만 지우고 가리키는 폴더(사용자 것일 수 있음)는 그대로."""
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    real = tmp_path / "내 스킬 원본"
    real.mkdir()
    (real / "SKILL.md").write_text("---\nname: office-cli\n---\n원본", encoding="utf-8")
    (old / "skills").mkdir(parents=True)
    link = old / "skills" / "office-cli"
    try:
        os.symlink(real, link, target_is_directory=True)
    except (OSError, NotImplementedError):
        pytest.skip("이 환경에서는 링크를 만들 수 없음")
    assert su.main(["--no-path"]) == 0
    assert not os.path.lexists(link) and (real / "SKILL.md").is_file()
    assert not _backups(cli, "old_config/skills/office-cli")


def test_missing_source_is_error(monkeypatch, tmp_path, capsys):
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    os.remove(cli / "AGENTS.md")
    shutil.rmtree(cli / "skills" / "excel-kit")
    assert su.main(["--no-path"]) == 1
    out = capsys.readouterr().out
    assert "[FAIL] 전역 지침: 파일이 없습니다" in out and "[FAIL] excel-kit 스킬: 파일이 없습니다" in out and "[STATUS] 끝." not in out


def test_previous_layout_is_moved_to_the_aikit_root(monkeypatch, tmp_path, capsys):
    """1.6.6: 바로 전 구성(2026-11-12 낮: 설정 폴더 C:\\AIKIT\\opencode, 스킬은 도구 폴더 안)에서 옮김. opencode.jsonc와 사용자 스킬은
    맨 위로, AGENTS.md·opencode.json·service.json·도구 폴더 안의 skills는 보관 후 지우고, 빈 opencode 폴더도 지움. --check는 바꾸지 않음."""
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    oc = cli / "opencode"
    oc.mkdir()
    for name, text in (("AGENTS.md", "이전 지침"), ("opencode.json", '{"skills": {"paths": ["../ml-kit/skills"]}}'),
                       ("opencode.jsonc", '{"model": "corp/gemma"}'), ("service.json", "{}")):
        (oc / name).write_text(text, encoding="utf-8")
    _skill(oc, "weekly-report", "내 스킬")
    tool_skill = cli / "ml-kit" / "skills" / "ml-train" / "SKILL.md"
    tool_skill.parent.mkdir(parents=True)
    tool_skill.write_text("이전 위치의 기본 스킬", encoding="utf-8")
    assert su.main(["--no-path", "--check"]) == 1
    out = capsys.readouterr().out
    assert "이전 구성 우리 스킬" in out and "이전 구성 사용자 스킬" in out and "이전 구성 opencode.jsonc" in out and oc.is_dir()
    assert su.main(["--no-path"]) == 0
    assert (cli / "opencode.jsonc").read_text(encoding="utf-8") == '{"model": "corp/gemma"}'
    assert (cli / "skills" / "weekly-report" / "SKILL.md").read_text(encoding="utf-8").endswith("내 스킬")
    assert not oc.exists() and not (cli / "ml-kit" / "skills").exists()
    for rel in ("opencode/AGENTS.md", "opencode/opencode.json", "opencode/opencode.jsonc", "ml-kit/skills/ml-train/SKILL.md"):
        assert len(_backups(cli, "old_layout/" + rel)) == 1, rel
    assert su.main(["--no-path", "--check"]) == 0


def test_user_skill_in_old_tool_skills_folder_is_moved_not_deleted(monkeypatch, tmp_path, capsys):
    """1.6.6 검토: 이전 구성의 도구 폴더 안 skills에 사용자가 둔 스킬은 폴더째 지우지 않고 C:\\AIKIT\\skills로 옮김.
    모르는 파일이 남으면 그 폴더는 그대로 둠. C:\\AIKIT\\skills에 이미 있는 사용자 스킬은 바꾸지 않음."""
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    tool = cli / "office-cli"
    _skill(tool, "excel-kit", "이전 위치의 기본 스킬")
    _skill(tool, "my-check", "도구 폴더에 둔 내 스킬")
    _skill(cli / "ml-kit", "my-ml", "ml-kit 쪽 내 스킬")
    (cli / "ml-kit" / "skills" / "notes.txt").write_text("메모", encoding="utf-8")
    _skill(cli, "weekly-report", "있던 내 스킬")
    assert su.main(["--no-path"]) == 0
    out = capsys.readouterr().out
    assert (cli / "skills" / "my-check" / "SKILL.md").read_text(encoding="utf-8").endswith("도구 폴더에 둔 내 스킬")
    assert (cli / "skills" / "my-ml" / "SKILL.md").read_text(encoding="utf-8").endswith("ml-kit 쪽 내 스킬")
    assert (cli / "skills" / "weekly-report" / "SKILL.md").read_text(encoding="utf-8").endswith("있던 내 스킬")
    assert not (tool / "skills").exists()                                 # 다 비어서 지움
    assert (cli / "ml-kit" / "skills" / "notes.txt").is_file()            # 모르는 파일은 그대로
    assert not (cli / "ml-kit" / "skills" / "my-ml").exists() and "이전 구성 사용자 스킬" in out
    for rel in ("office-cli/skills/excel-kit/SKILL.md", "office-cli/skills/my-check/SKILL.md", "ml-kit/skills/my-ml/SKILL.md"):
        assert len(_backups(cli, "old_layout/" + rel)) == 1, rel
    assert su.main(["--no-path", "--check"]) == 0


def test_renamed_base_skill_folders_are_removed_user_skills_kept(monkeypatch, tmp_path, capsys):
    """1.6.7: 기본 스킬 이름을 바꿈 (office-excel -> excel-kit, office-doc -> word-kit·ppt-kit 등). 폴더를 덮어써 배포하면 옛 이름 폴더가
    C:\\AIKIT\\skills에 남아 OpenCode가 옛 스킬도 읽음 -> 보관 후 지움. 사용자 스킬은 그대로, 새 기본 스킬도 그대로."""
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    for name in su.RENAMED_SKILLS:
        _skill(cli, name, "옛 이름의 기본 스킬")
    mine = _skill(cli, "weekly-report", "내 스킬")
    assert su.main(["--no-path", "--check"]) == 1
    out = capsys.readouterr().out
    assert "이전 구성 우리 스킬(옛 이름)" in out and (cli / "skills" / "office-doc").is_dir()
    assert su.main(["--no-path"]) == 0
    for name in su.RENAMED_SKILLS:
        assert not (cli / "skills" / name).exists(), name
        assert len(_backups(cli, f"old_layout/skills/{name}/SKILL.md")) == 1, name
    assert mine.read_text(encoding="utf-8").endswith("내 스킬")
    for name in su.OFFICE_SKILLS + su.ML_SKILLS:
        assert (cli / "skills" / name / "SKILL.md").is_file(), name
    assert su.main(["--no-path", "--check"]) == 0


def test_test_skill_goes_to_user_skills(monkeypatch, tmp_path, capsys):
    cli, old = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    test_skill = os.path.join(DEV, "office-cli", "tools", "office-cli-test", "SKILL.md")
    assert su.main(["--no-path", "--test-skill", test_skill]) == 0
    assert (cli / "skills" / "office-cli-test" / "SKILL.md").read_bytes() == open(test_skill, "rb").read()
    assert "office-cli-test 스킬(개발용)" in capsys.readouterr().out


def test_not_windows_env_is_warning_only(monkeypatch, tmp_path, capsys):
    cli, old = _iso(monkeypatch, tmp_path, None)
    assert su.main(["--no-path"]) == 0
    assert "[WARN] Windows가 아니라 사용자 환경 변수 OPENCODE_CONFIG_DIR" in capsys.readouterr().out


def test_opencode_paths_must_be_the_aikit_folder(monkeypatch, tmp_path, capsys):
    """opencode debug paths의 config 줄로 OpenCode가 OPENCODE_CONFIG_DIR를 따르는지 확인 (OpenCode 2.2.5 출력 모양)."""
    monkeypatch.setattr(su, "HERE", str(tmp_path))
    rep = su.Report()
    su.step_opencode_paths(rep, run=lambda cmd, t: (0, "", ""))
    assert rep.ok == 0 and rep.fails == []                                # opencode.exe 없음: 지나감
    (tmp_path / "opencode.exe").write_bytes(b"")
    good = f"home       C:\\Users\\x\ndata       C:\\Users\\x\\.local\\share\\opencode\nconfig     {tmp_path}\n"
    su.step_opencode_paths(rep, run=lambda cmd, t: (0, good, ""))
    assert rep.ok == 1 and rep.fails == []
    su.step_opencode_paths(rep, run=lambda cmd, t: (0, good.replace(f"config     {tmp_path}", r"config     C:\Users\x\.config\opencode"), ""))
    assert len(rep.fails) == 1 and "로그아웃" in rep.fails[0]
    rep = su.Report()
    su.step_opencode_paths(rep, run=lambda cmd, t: (1, "unknown command", ""))
    assert rep.fails == [] and "config 줄이 없어" in capsys.readouterr().out


# ---- 시스템 PATH (가짜 winreg) ----
def test_merge_path_adds_only_missing_and_keeps_the_rest():
    a, b = r"C:\AIKIT\python", r"C:\AIKIT\python\Scripts"
    base = r"%SystemRoot%\system32;C:\Tools;;D:\x\ "
    new, add = su.merge_path(base, [a, b])
    assert add == [a, b] and new == base + ";" + a + ";" + b               # 원래 값은 글자 그대로, 뒤에만 붙임
    new, add = su.merge_path(r"c:\aikit\PYTHON\;C:\Tools;" + "\"C:\\AIKIT\\python\\Scripts\";", [a, b])
    assert add == [] and new == r"c:\aikit\PYTHON\;C:\Tools;" + "\"C:\\AIKIT\\python\\Scripts\";"   # 대소문자·끝 \·따옴표 무시
    new, add = su.merge_path(r"C:\Tools;", [a, b])
    assert new == r"C:\Tools;" + a + ";" + b                              # 끝 ;이 있으면 ;를 더 넣지 않음
    assert su.merge_path("", [a, b]) == (a + ";" + b, [a, b])


def test_path_added_with_expand_sz_backup_and_broadcast(monkeypatch, tmp_path, capsys):
    base = r"%SystemRoot%\system32;%SystemRoot%;C:\Program Files\Git\cmd"
    reg = FakeReg(base, FakeReg.REG_EXPAND_SZ)
    cli, _old = _iso(monkeypatch, tmp_path, reg)
    sent = []
    monkeypatch.setattr(su, "broadcast_env_change", lambda: sent.append(1) or True)
    assert su.main([]) == 0
    want = [str(cli / "python"), str(cli / "python" / "Scripts")]
    assert reg.writes == [("Path", FakeReg.REG_EXPAND_SZ, base + ";" + ";".join(want))]
    assert reg.value.startswith(base + ";") and sent == [1, 1]           # 사용자 환경 변수 한 번, PATH 한 번
    kept = _backups(cli, "system_path.txt")
    assert len(kept) == 1 and base in kept[0].read_text(encoding="utf-8")       # 쓰기 전 원래 값
    assert "시스템 PATH에 추가" in capsys.readouterr().out
    assert su.main([]) == 0                         # 두 번째: 이미 있음, 쓰지 않음
    assert len(reg.writes) == 1 and "이미 있음" in capsys.readouterr().out
    assert su.missing_path_entries(reg) == []


def test_path_keeps_reg_sz_and_creates_missing_value(monkeypatch, tmp_path):
    reg = FakeReg(r"C:\Windows", FakeReg.REG_SZ)
    cli, _old = _iso(monkeypatch, tmp_path, reg)
    assert su.main([]) == 0
    assert reg.writes[0][1] == FakeReg.REG_SZ                               # 원래 형식 유지
    reg2 = FakeReg(None)
    monkeypatch.setattr(su, "_winreg", lambda: reg2)
    assert su.main([]) == 0
    assert reg2.writes[0][1] == FakeReg.REG_EXPAND_SZ and reg2.value == str(cli / "python") + ";" + str(cli / "python" / "Scripts")


def test_path_check_and_no_path_do_not_write(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows")
    cli, _old = _iso(monkeypatch, tmp_path, reg)
    su.main(["--no-path"])
    assert reg.writes == []
    capsys.readouterr()
    assert su.main(["--check"]) == 1
    assert reg.writes == [] and "시스템 PATH에 없음" in capsys.readouterr().out
    assert su.missing_path_entries(reg) == [str(cli / "python"), str(cli / "python" / "Scripts")]


def test_path_permission_denied_is_fail_without_write(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows", writable=False)
    cli, _old = _iso(monkeypatch, tmp_path, reg)
    assert su.main([]) == 1
    out = capsys.readouterr().out
    assert reg.writes == [] and reg.value == r"C:\Windows"
    assert "[FAIL] 시스템 PATH를 쓰지 못함 (권한 없음)" in out and "관리자 권한으로 실행" in out


def test_not_windows_path_is_warning_only(monkeypatch, tmp_path, capsys):
    cli, _old = _iso(monkeypatch, tmp_path, None)
    assert su.main([]) == 0
    assert "[WARN] Windows가 아니라 시스템 PATH" in capsys.readouterr().out
    assert su.missing_path_entries(None) is None


# ---- 동작 확인, setup.bat ----
def test_verify_runs_tool_help_and_imports(monkeypatch, tmp_path):
    cli = tmp_path / "AIKIT"
    for tool in ("ml-kit", "office-cli"):
        shutil.copytree(os.path.join(CLI, tool), cli / tool, ignore=shutil.ignore_patterns("data", "__pycache__"))
    (cli / "opencode.exe").write_bytes(b"")
    monkeypatch.setattr(su, "HERE", str(cli))
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "mldata"))
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path / "officedata"))
    rep = su.Report()
    su.step_verify(rep)
    assert rep.fails == [], rep.fails
    assert rep.ok >= 4                                                        # 라이브러리, 도구 2개, opencode.exe
    os.remove(cli / "opencode.exe")
    rep = su.Report()
    su.step_verify(rep)
    assert any("opencode.exe: 없음" in f for f in rep.fails)


def test_parse_jsonc_comments_and_trailing_commas():
    text = '{\n  // 사내 LLM\n  "model": "a//b", /* 막음 */ "x": [1, 2,],\n  "s": "q\\"//",\n}'
    assert su.parse_jsonc(text) == {"model": "a//b", "x": [1, 2], "s": 'q"//'}
    assert su.parse_jsonc('{"a": "x, ]", "b": [1,],}') == {"a": "x, ]", "b": [1]}          # 문자열 안의 쉼표는 그대로


def test_opencode_config_warns_on_shell_instructions_and_permission(tmp_path, capsys):
    """집 PC opencode 시험(2026-11-07)에서 결과를 흐린 설정: 셸 cmd, 옛 지침 instructions, 작업 폴더 밖 쓰기 거절."""
    (tmp_path / "opencode.jsonc").write_text(
        '{\n // 설정\n "shell": "cmd", "instructions": ["C:/old/AGENTS.md"],\n "permission": {"external_directory": "deny"},\n}',
        encoding="utf-8")
    rep = su.Report()
    su.step_opencode_config(rep, str(tmp_path))
    out = capsys.readouterr().out
    assert rep.warn == 3 and rep.fails == [], out
    assert '"shell"이 PowerShell이 아닙니다' in out and "C:/old/AGENTS.md" in out and "external_directory" in out
    (tmp_path / "opencode.jsonc").write_text('{"shell": "powershell", "permission": {"external_directory": "allow"}}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_config(rep, str(tmp_path))
    assert rep.warn == 0
    for perm in ('"allow"', '{"external_directory": {"*": "allow"}}'):                 # 전체 허용, 경로 패턴별 허용
        (tmp_path / "opencode.jsonc").write_text('{"permission": %s}' % perm, encoding="utf-8")
        rep = su.Report()
        su.step_opencode_config(rep, str(tmp_path))
        assert rep.warn == 0, perm
    (tmp_path / "opencode.jsonc").write_text('{"model": "x"}', encoding="utf-8")       # 권한이 없으면 INFO만 (세지 않음)
    rep = su.Report()
    su.step_opencode_config(rep, str(tmp_path))
    assert rep.warn == 0 and "[INFO] 설정에 permission.external_directory가 없습니다" in capsys.readouterr().out
    (tmp_path / "opencode.jsonc").write_text('{ 깨진', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_config(rep, str(tmp_path))
    assert rep.warn == 1 and rep.fails == []                                       # 못 읽어도 설치를 실패로 보지 않음


def test_opencode_providers_suggests_enabled_providers_line(tmp_path, capsys):
    """폐쇄망에서 못 쓰는 기본 모델 숨기기: enabled_providers가 없으면 사내 provider 이름을 넣은 줄을 그대로 알려 줌."""
    cfg = tmp_path / "opencode.jsonc"
    cfg.write_text('{"model": "corp/gemma", "provider": {"corp": {"npm": "@ai-sdk/openai-compatible"}}}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_providers(rep, str(tmp_path))
    out = capsys.readouterr().out
    assert rep.warn == 1 and '"enabled_providers": ["corp"]' in out and '"autoupdate": false' in out, out
    cfg.write_text('{"model": "corp/gemma", "provider": {"corp": {}}, "enabled_providers": ["corp"], "autoupdate": false}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_providers(rep, str(tmp_path))
    out = capsys.readouterr().out                             # 웹 도구를 막지 않았으면 INFO (폐쇄망에서 모델이 인터넷 검색을 시도함, 2026-11-11)
    assert rep.warn == 0 and '"webfetch": "deny", "websearch": "deny"' in out, out
    cfg.write_text('{"model": "corp/gemma", "provider": {"corp": {}}, "enabled_providers": ["corp"], "autoupdate": false, '
                   '"permission": {"external_directory": "allow", "webfetch": "deny", "websearch": "deny"}}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_providers(rep, str(tmp_path))
    assert rep.warn == 0 and capsys.readouterr().out == ""
    cfg.write_text('{"model": "corp/gemma", "provider": {"corp": {}}, "enabled_providers": ["crop"], "autoupdate": false}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_providers(rep, str(tmp_path))
    out = capsys.readouterr().out
    assert rep.warn == 2 and "provider\"에 없는 이름" in out and "그 모델을 쓸 수 없습니다" in out, out
    cfg.write_text('{"enabled_providers": []}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_providers(rep, str(tmp_path))
    assert rep.warn == 1 and "모든 모델이 숨겨집니다" in capsys.readouterr().out


def test_setup_bat_is_ascii_crlf_and_uses_bundled_python():
    data = open(os.path.join(CLI, "setup.bat"), "rb").read()
    assert data.isascii() and b"chcp" not in data.lower()
    assert data.count(b"\r\n") == data.count(b"\n") and data.endswith(b"\r\n")
    assert b'"%~dp0python\\python.exe" "%~dp0setup.py" %*' in data and b"pause" in data


# ---- run_all_tests.py diag ----
def test_diag_fails_when_env_path_or_old_copies_differ(tmp_path, monkeypatch):
    """run_all_tests.py diag: setup.py --check처럼 읽기만. PATH, OPENCODE_CONFIG_DIR, 옛 전역 설정의 우리 스킬·지침 복사본."""
    rat = _load("run_all_tests_s", os.path.join(DEV, "run_all_tests.py"))
    old = tmp_path / "old_cfg"
    lines = []
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(old), reg=None)
    if os.name != "nt":                                                   # Windows가 아니면 참고만
        assert st.failed == 0 and st.passed == 0 and any("Windows가 아니라" in x for x in lines)
    paths = ";".join(os.path.join(CLI, *p) for p in (("python",), ("python", "Scripts")))
    want = CLI
    reg = FakeReg(r"C:\Windows;" + paths)
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(old), reg=reg)
    assert st.failed == 1 and "OPENCODE_CONFIG_DIR가 없음" in st.failures[0] and "setup.bat" in st.failures[0]
    reg.user[su.CONFIG_ENV] = want
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(old), reg=reg)
    assert st.failed == 0 and st.passed == 2 and reg.writes == [] and reg.user == {su.CONFIG_ENV: want}   # 읽기만
    _skill(old, "excel-kit")
    _skill(old, "weekly-report")                                          # 사용자 스킬만 있으면 실패가 아님 (setup이 옮김)
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(old), reg=reg)
    assert st.failed == 1 and "excel-kit" in st.failures[0] and "weekly-report" not in st.failures[0]


def test_service_is_stopped_so_new_path_and_skills_apply(tmp_path, monkeypatch, capsys):
    """팀원 배포 2026-11-11: 설치 전에 뜬 OpenCode 백그라운드 서비스가 옛 PATH(python 없음)로 계속 돌아, web.bat을 다시 실행해도
    재부팅 전까지 python을 못 찾았다. setup은 서비스를 끄고(다음 실행 때 새 설정으로 뜸), opencode.exe가 없으면 지나간다."""
    calls = []
    monkeypatch.setattr(su, "HERE", str(tmp_path))
    rep = su.Report()
    su.step_service(rep, run=lambda cmd, t: calls.append(cmd) or (0, "", ""))
    assert calls == [] and rep.warn == 0                                    # opencode.exe 없음
    (tmp_path / "opencode.exe").write_bytes(b"")
    su.step_service(rep, run=lambda cmd, t: calls.append(cmd) or (0, "", ""))
    assert calls == [[str(tmp_path / "opencode.exe"), "service", "stop"]] and "껐습니다" in capsys.readouterr().out
    rep = su.Report()
    su.step_service(rep, run=lambda cmd, t: (_ for _ in ()).throw(OSError("x")))
    assert rep.warn == 1 and rep.fails == [] and "web.bat stop" in capsys.readouterr().out   # 못 꺼도 설치 실패가 아님


def test_old_skills_path_in_jsonc_is_warned(tmp_path, capsys):
    """1.6.6: 예전 opencode.jsonc의 "skills"(옛 전역 폴더 등)가 남으면 옛 스킬이 섞임 -> WARN (스킬은 C:\\AIKIT\\skills에 두면 됨)."""
    (tmp_path / "opencode.jsonc").write_text('{"skills": ["C:/Users/x/.config/opencode/skills"], "model": "x"}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_config(rep, str(tmp_path))
    assert rep.warn == 1 and '"skills"(스킬 경로)' in capsys.readouterr().out