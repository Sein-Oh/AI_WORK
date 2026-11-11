# -*- coding: utf-8 -*-
"""AIKIT/setup.py (배포 설치): 스킬(ml-kit, office-cli 작업별 5개)·전역 AGENTS.md·opencode.jsonc를 OpenCode 설정 폴더로 복사
(바뀐 것만, 덮어쓰기 전 AIKIT\\data\\setup_backup_<시각>에 보관), 옛 통합 스킬 폴더 정리, 시스템 PATH 등록(가짜 winreg),
동작 확인, --check, setup.bat, run_all_tests.py diag의 스킬·PATH 확인."""
import importlib.util
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
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


class FakeReg:
    """winreg 대신: HKLM 시스템 환경 변수 Path 하나."""
    HKEY_LOCAL_MACHINE, KEY_READ, KEY_SET_VALUE, REG_SZ, REG_EXPAND_SZ = "HKLM", 1, 2, 1, 2

    def __init__(self, value=None, kind=2, writable=True):
        self.value, self.kind, self.writable, self.writes = value, kind, writable, []

    def OpenKey(self, root, sub, reserved=0, access=1):
        assert root == "HKLM" and sub == su.ENV_KEY
        if access & self.KEY_SET_VALUE and not self.writable:
            raise PermissionError(5, "액세스가 거부되었습니다")
        return _Key()

    def QueryValueEx(self, key, name):
        assert name == "Path"
        if self.value is None:
            raise FileNotFoundError(2, "없음")
        return self.value, self.kind

    def SetValueEx(self, key, name, reserved, kind, value):
        self.writes.append((name, kind, value))
        self.value, self.kind = value, kind


def _iso(monkeypatch, tmp_path, reg=None):
    """가짜 C:\\AIKIT(스킬·전역 지침만 저장소에서 복사)와 가짜 설정 폴더. 동작 확인은 끔, PATH는 가짜 winreg."""
    cli = tmp_path / "AIKIT"
    for _n, src, _rel in su.targets():
        dst = cli / os.path.relpath(src, CLI)
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(open(src, "rb").read())
    monkeypatch.setattr(su, "HERE", str(cli))
    monkeypatch.setattr(su, "step_verify", lambda rep: None)
    monkeypatch.setattr(su, "_winreg", lambda: reg)
    monkeypatch.setattr(su, "broadcast_env_change", lambda: True)
    return cli, tmp_path / "opencode_cfg"


def _backups(cli, pattern):
    return list((cli / "data").glob("setup_backup_*/" + pattern))


def test_sources_exist_in_repo_layout():
    for name, src, rel in su.targets():
        assert os.path.isfile(src), (name, src)
    assert su.default_config_dir().endswith(os.path.join(".config", "opencode"))
    on_disk = sorted(os.listdir(os.path.join(CLI, "office-cli", "skills")))
    assert sorted(su.OFFICE_SKILLS) == on_disk                            # 새 스킬 폴더를 만들면 setup.py에도 올림
    assert {os.path.dirname(rel) for _n, _s, rel in su.targets()} >= {os.path.join("skills", n) for n in on_disk}
    assert su.jsonc_source() == os.path.join(CLI, "opencode", "opencode.jsonc")
    gitignore = os.path.join(REPO, ".gitignore")
    if os.path.isfile(gitignore):        # 저장소에서만 (회사 PC는 C:\AIKIT\dev만 반입해 .gitignore가 없음)
        ignore = open(gitignore, encoding="utf-8").read().splitlines()
        assert "AIKIT/opencode/opencode.jsonc" in ignore and "AIKIT/data/" in ignore     # 사내 LLM 주소, 보관 폴더는 git 제외
    assert not os.path.exists(os.path.join(DEV, "apply_skills.py"))             # 설치는 setup.py 하나
    for dist in ("pandas", "scikit-learn", "openpyxl", "python-docx", "python-pptx", "pymupdf", "pywin32"):
        assert dist in dict(su.LIBS)


def test_apply_copies_new_then_reports_same(monkeypatch, tmp_path, capsys):
    cli, cfg = _iso(monkeypatch, tmp_path)
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0
    for name in su.ML_SKILLS:
        assert (cfg / "skills" / name / "SKILL.md").read_bytes() == (cli / "ml-kit" / "skills" / name / "SKILL.md").read_bytes()
    assert (cfg / "AGENTS.md").read_bytes() == (cli / "opencode" / "AGENTS.md").read_bytes()
    assert not (cfg / "skills" / "office-cli-test").exists()
    assert not list((cli / "data").glob("setup_backup_*"))                # 덮어쓴 것이 없으면 보관도 없음
    assert (cli / "data" / "scripts").is_dir()                            # 전역 지침 8장의 스크립트 폴더
    assert (cli / "work").is_dir()                                        # Web UI의 작업 폴더 (팀원 배포 2026-11-11)
    out = capsys.readouterr().out
    assert "[STATUS] 끝." in out and "새 터미널" in out
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 0
    out = capsys.readouterr().out
    assert "다름" not in out.replace("다름 0", "") and "[FAIL]" not in out


def test_changed_file_is_backed_up_and_check_fails_before_apply(monkeypatch, tmp_path, capsys):
    cli, cfg = _iso(monkeypatch, tmp_path)
    su.main(["--config-dir", str(cfg), "--no-path"])
    old = cfg / "skills" / "office-excel" / "SKILL.md"
    old.write_text('---\nname: office-excel\nmetadata:\n  version: "1.0.0"\n---\n옛 내용', encoding="utf-8")
    (cfg / "skills" / "ml-train" / "ml_train.py").write_text("x", encoding="utf-8")
    capsys.readouterr()
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 1
    out = capsys.readouterr().out
    assert "office-excel 스킬: 다름 (1.0.0 -> " in out and "옛 내용" in old.read_text(encoding="utf-8")   # --check는 바꾸지 않음
    assert "ml_train.py" not in out                     # 스킬 폴더의 .py는 알리지 않음
    test_skill = os.path.join(DEV, "office-cli", "tools", "office-cli-test", "SKILL.md")
    assert su.main(["--config-dir", str(cfg), "--no-path", "--test-skill", test_skill]) == 0
    # 2026-11-07 집 PC: 복사한 뒤에 버전을 읽어 '다름 (1.4.82 -> 1.4.82)'로 보였음. 옛 버전은 복사 전 것
    assert "office-excel 스킬: 다름 (1.0.0 -> " in capsys.readouterr().out
    kept = _backups(cli, "config/skills/office-excel/SKILL.md")
    assert len(kept) == 1 and "옛 내용" in kept[0].read_text(encoding="utf-8")
    assert old.read_bytes() == (cli / "office-cli" / "skills" / "office-excel" / "SKILL.md").read_bytes()
    assert (cfg / "skills" / "office-cli-test" / "SKILL.md").read_bytes() == open(test_skill, "rb").read()   # 개발용 보조 스킬


def test_old_office_cli_skill_is_backed_up_and_removed(monkeypatch, tmp_path, capsys):
    """1.4.61까지의 통합 스킬 폴더가 남으면 OpenCode가 옛 스킬도 고름: --check는 다름, 적용하면 보관 후 지움."""
    cli, cfg = _iso(monkeypatch, tmp_path)
    su.main(["--config-dir", str(cfg), "--no-path"])
    old = cfg / "skills" / "office-cli" / "SKILL.md"
    old.parent.mkdir(parents=True)
    old.write_text('---\nname: office-cli\nmetadata:\n  version: "1.4.61"\n---\n옛 통합 스킬', encoding="utf-8")
    capsys.readouterr()
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 1
    assert "옛 office-cli 스킬: 지울 것" in capsys.readouterr().out and old.is_file()
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0
    assert not old.parent.exists()
    kept = _backups(cli, "config/skills/office-cli/SKILL.md")
    assert len(kept) == 1 and "옛 통합 스킬" in kept[0].read_text(encoding="utf-8")
    assert "보관 후 지움" in capsys.readouterr().out
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 0


def test_old_ml_kit_skill_is_backed_up_and_removed(monkeypatch, tmp_path, capsys):
    """2.9.71: ml-kit 스킬을 넷(ml-train, ml-predict, ml-report, ml-map)으로 나눔. 옛 skills\\ml-kit이 남으면 OpenCode가 옛 스킬도 고름."""
    cli, cfg = _iso(monkeypatch, tmp_path)
    su.main(["--config-dir", str(cfg), "--no-path"])
    old = cfg / "skills" / "ml-kit" / "SKILL.md"
    old.parent.mkdir(parents=True)
    old.write_text('---\nname: ml-kit\nmetadata:\n  version: "2.9.70"\n---\n옛 통합 스킬', encoding="utf-8")
    capsys.readouterr()
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 1
    assert "옛 ml-kit 스킬: 지울 것" in capsys.readouterr().out
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0
    assert not old.parent.exists() and len(_backups(cli, "config/skills/ml-kit/SKILL.md")) == 1
    for name in su.ML_SKILLS:
        assert (cfg / "skills" / name / "SKILL.md").is_file()


def test_old_office_cli_skill_link_removes_only_link(monkeypatch, tmp_path, capsys):
    """옛 skills\\office-cli가 링크(정션)이면 링크만 지우고 가리키는 폴더(사용자 것일 수 있음)는 그대로."""
    cli, cfg = _iso(monkeypatch, tmp_path)
    su.main(["--config-dir", str(cfg), "--no-path"])
    real = tmp_path / "내 스킬 원본"
    real.mkdir()
    (real / "SKILL.md").write_text("---\nname: office-cli\n---\n원본", encoding="utf-8")
    link = cfg / "skills" / "office-cli"
    try:
        os.symlink(real, link, target_is_directory=True)
    except (OSError, NotImplementedError):
        pytest.skip("이 환경에서는 링크를 만들 수 없음")
    capsys.readouterr()
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0
    out = capsys.readouterr().out
    assert not os.path.lexists(link) and (real / "SKILL.md").is_file(), out
    assert "링크만 지움" in out and str(real) in out
    assert not _backups(cli, "config/skills/office-cli")


def test_missing_source_is_error(monkeypatch, tmp_path, capsys):
    cli, cfg = _iso(monkeypatch, tmp_path)
    os.remove(cli / "opencode" / "AGENTS.md")
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 1
    out = capsys.readouterr().out
    assert "[FAIL] 전역 지침: 원본이 없습니다" in out and not (cfg / "AGENTS.md").exists()
    assert "[STATUS] 끝." not in out


def test_opencode_jsonc_copied_backed_up_or_warned(monkeypatch, tmp_path, capsys):
    cli, cfg = _iso(monkeypatch, tmp_path)
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0          # 원본도 설정 폴더 것도 없음: 경고만
    assert "[WARN] opencode.jsonc:" in capsys.readouterr().out and not (cfg / "opencode.jsonc").exists()
    (cfg / "opencode.jsonc").write_text('{"model": "사용자 것"}', encoding="utf-8")
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0          # 원본 없음, 설정 폴더 것 사용
    assert "그대로 사용" in capsys.readouterr().out
    (cli / "opencode" / "opencode.jsonc").write_text('{"model": "사내 LLM"}', encoding="utf-8")
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 1
    assert "사용자 것" in (cfg / "opencode.jsonc").read_text(encoding="utf-8")
    assert su.main(["--config-dir", str(cfg), "--no-path"]) == 0
    assert "사내 LLM" in (cfg / "opencode.jsonc").read_text(encoding="utf-8")
    kept = _backups(cli, "config/opencode.jsonc")
    assert len(kept) == 1 and "사용자 것" in kept[0].read_text(encoding="utf-8")
    assert su.main(["--config-dir", str(cfg), "--no-path", "--check"]) == 0


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
    cli, cfg = _iso(monkeypatch, tmp_path, reg)
    sent = []
    monkeypatch.setattr(su, "broadcast_env_change", lambda: sent.append(1) or True)
    assert su.main(["--config-dir", str(cfg)]) == 0
    want = [str(cli / "python"), str(cli / "python" / "Scripts")]
    assert reg.writes == [("Path", FakeReg.REG_EXPAND_SZ, base + ";" + ";".join(want))]
    assert reg.value.startswith(base + ";") and sent == [1]
    kept = _backups(cli, "system_path.txt")
    assert len(kept) == 1 and base in kept[0].read_text(encoding="utf-8")       # 쓰기 전 원래 값
    assert "시스템 PATH에 추가" in capsys.readouterr().out
    assert su.main(["--config-dir", str(cfg)]) == 0                         # 두 번째: 이미 있음, 쓰지 않음
    assert len(reg.writes) == 1 and "이미 있음" in capsys.readouterr().out
    assert su.missing_path_entries(reg) == []


def test_path_keeps_reg_sz_and_creates_missing_value(monkeypatch, tmp_path):
    reg = FakeReg(r"C:\Windows", FakeReg.REG_SZ)
    cli, cfg = _iso(monkeypatch, tmp_path, reg)
    assert su.main(["--config-dir", str(cfg)]) == 0
    assert reg.writes[0][1] == FakeReg.REG_SZ                               # 원래 형식 유지
    reg2 = FakeReg(None)
    monkeypatch.setattr(su, "_winreg", lambda: reg2)
    assert su.main(["--config-dir", str(cfg)]) == 0
    assert reg2.writes[0][1] == FakeReg.REG_EXPAND_SZ and reg2.value == str(cli / "python") + ";" + str(cli / "python" / "Scripts")


def test_path_check_and_no_path_do_not_write(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows")
    cli, cfg = _iso(monkeypatch, tmp_path, reg)
    su.main(["--config-dir", str(cfg), "--no-path"])
    assert reg.writes == []
    capsys.readouterr()
    assert su.main(["--config-dir", str(cfg), "--check"]) == 1
    assert reg.writes == [] and "시스템 PATH에 없음" in capsys.readouterr().out
    assert su.missing_path_entries(reg) == [str(cli / "python"), str(cli / "python" / "Scripts")]


def test_path_permission_denied_is_fail_without_write(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows", writable=False)
    cli, cfg = _iso(monkeypatch, tmp_path, reg)
    assert su.main(["--config-dir", str(cfg)]) == 1
    out = capsys.readouterr().out
    assert reg.writes == [] and reg.value == r"C:\Windows"
    assert "[FAIL] 시스템 PATH를 쓰지 못함 (권한 없음)" in out and "관리자 권한으로 실행" in out


def test_not_windows_path_is_warning_only(monkeypatch, tmp_path, capsys):
    cli, cfg = _iso(monkeypatch, tmp_path, None)
    assert su.main(["--config-dir", str(cfg)]) == 0
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
    assert rep.warn == 0 and "[INFO] opencode.jsonc에 permission.external_directory가 없습니다" in capsys.readouterr().out
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
def test_diag_fails_when_installed_skill_or_path_differs(tmp_path, monkeypatch):
    rat = _load("run_all_tests_s", os.path.join(DEV, "run_all_tests.py"))
    cfg = tmp_path / "cfg"
    lines = []
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(cfg), reg=None)                # 설정 폴더 없음: 확인 안 함
    if os.name != "nt":
        assert st.failed == 0 and st.passed == 0 and "확인 안 함" in lines[-1]
        assert any("Windows가 아니라" in x for x in lines)
    st = rat.Stage("diag")
    monkeypatch.setattr(su, "HERE", CLI)
    monkeypatch.setattr(su, "_winreg", lambda: None)
    rep = su.Report()
    su.step_files(rep, str(cfg), str(tmp_path / "bk"))
    jsonc = os.path.isfile(su.jsonc_source())                               # 개발 PC에 실제 opencode.jsonc가 있으면 그것도 복사됨
    rat.diag_skills(st, lines, config_dir=str(cfg), reg=FakeReg(r"C:\Windows"))
    assert st.failed == 1 and "시스템 PATH에 없음" in st.failures[0] and st.passed == 1   # PATH는 읽기만
    paths = ";".join(os.path.join(CLI, *p) for p in (("python",), ("python", "Scripts")))
    reg = FakeReg(r"C:\Windows;" + paths)
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, config_dir=str(cfg), reg=reg)
    assert st.failed == 0 and st.passed == 2 and reg.writes == []
    (cfg / "AGENTS.md").write_text("옛 지침", encoding="utf-8")
    rat.diag_skills(st, lines, config_dir=str(cfg), reg=reg)
    assert st.failed == 1 and "setup.bat" in st.failures[0] and "전역 지침 다름" in st.failures[0]
    cfg2 = tmp_path / "cfg2"
    su.step_files(su.Report(), str(cfg2), str(tmp_path / "bk2"))
    (cfg2 / "skills" / "office-cli").mkdir()                                 # 옛 통합 스킬 폴더가 남음
    (cfg2 / "skills" / "office-cli" / "SKILL.md").write_text("옛 스킬", encoding="utf-8")
    rat.diag_skills(st, lines, config_dir=str(cfg2), reg=reg)
    assert st.failed == 2 and "옛 office-cli 스킬 지울 것" in st.failures[1]
    assert jsonc or not (cfg / "opencode.jsonc").exists()


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
