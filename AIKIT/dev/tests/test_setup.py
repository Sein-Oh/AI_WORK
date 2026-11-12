# -*- coding: utf-8 -*-
"""AIKIT/setup.py (배포 설치): OpenCode 설정 폴더를 C:\\AIKIT으로 (사용자 환경 변수 OPENCODE_CONFIG_DIR, 가짜 winreg),
시스템 PATH 등록(가짜 winreg), 동작 확인, opencode.json + opencode.jsonc 합친 설정 점검, --check, setup.bat,
run_all_tests.py diag의 환경 변수·PATH 확인."""
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
    """가짜 C:\\AIKIT(OpenCode가 읽는 파일만 저장소에서 복사). 동작 확인·opencode 확인은 끔."""
    cli = tmp_path / "AIKIT"
    for _n, src in su.sources():
        dst = cli / os.path.relpath(src, CLI)
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(open(src, "rb").read())
    monkeypatch.setattr(su, "HERE", str(cli))
    monkeypatch.setattr(su, "step_verify", lambda rep: None)
    monkeypatch.setattr(su, "step_opencode_paths", lambda rep: None)
    monkeypatch.setattr(su, "_winreg", lambda: reg)
    monkeypatch.setattr(su, "broadcast_env_change", lambda: True)
    monkeypatch.delenv(su.CONFIG_ENV, raising=False)                    # step_env가 이 프로세스에 넣는 값을 시험 뒤 되돌림
    return cli


def _backups(cli, pattern):
    return list((cli / "data").glob("setup_backup_*/" + pattern))


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
    assert cfg["shell"] == "powershell"
    # 폐쇄망에서 되지 않는 기능 잠금 (setup이 점검하는 값과 같음)
    assert all(cfg.get(k) == v for k, v in su.OFFLINE_SET)
    assert cfg["permission"] == {"external_directory": "allow", **{k: "deny" for k in su.OFFLINE_DENY}}
    gitignore = os.path.join(REPO, ".gitignore")
    if os.path.isfile(gitignore):        # 저장소에서만 (반입한 PC에는 .gitignore가 없음)
        ignore = open(gitignore, encoding="utf-8").read().splitlines()
        assert {"AIKIT/opencode.jsonc", "AIKIT/skills/*", "AIKIT/data/", "AIKIT/service.json"} <= set(ignore)
        assert {f"!AIKIT/skills/{n}/" for n in su.ML_SKILLS + su.OFFICE_SKILLS} <= set(ignore)    # 기본 스킬만 저장소에
        assert "AIKIT/opencode.json" not in ignore
    assert su.config_dir() == su.HERE
    for dist in ("pandas", "scikit-learn", "openpyxl", "python-docx", "python-pptx", "pymupdf", "pywin32"):
        assert dist in dict(su.LIBS)


def test_apply_sets_user_env_then_check_passes(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows", user={su.CONFIG_ENV: r"D:\옛 설정"})
    cli = _iso(monkeypatch, tmp_path, reg)
    assert su.main(["--no-path"]) == 0
    assert reg.user[su.CONFIG_ENV] == str(cli) and os.environ[su.CONFIG_ENV] == str(cli)
    kept = _backups(cli, "user_env.txt")
    assert len(kept) == 1 and r"D:\옛 설정" in kept[0].read_text(encoding="utf-8")      # 쓰기 전 원래 값
    assert reg.writes == []                                               # --no-path: 시스템 PATH는 그대로
    assert (cli / "data" / "scripts").is_dir() and not (cli / "work").exists()     # Web UI 프로젝트 폴더는 C:\\AIKIT 자체 (1.6.7)
    out = capsys.readouterr().out
    assert "[STATUS] 끝." in out
    assert su.main(["--no-path", "--check"]) == 0
    out = capsys.readouterr().out
    assert "다름" not in out.replace("다름 0", "") and "[FAIL]" not in out and "사용자 환경 변수 OPENCODE_CONFIG_DIR = " in out


def test_missing_source_is_error(monkeypatch, tmp_path, capsys):
    cli = _iso(monkeypatch, tmp_path, FakeReg(r"C:\Windows"))
    os.remove(cli / "AGENTS.md")
    shutil.rmtree(cli / "skills" / "excel-kit")
    assert su.main(["--no-path"]) == 1
    out = capsys.readouterr().out
    assert "[FAIL] 전역 지침: 파일이 없습니다" in out and "[FAIL] excel-kit 스킬: 파일이 없습니다" in out and "[STATUS] 끝." not in out


def test_not_windows_env_is_warning_only(monkeypatch, tmp_path, capsys):
    cli = _iso(monkeypatch, tmp_path, None)
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
    cli = _iso(monkeypatch, tmp_path, reg)
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
    cli = _iso(monkeypatch, tmp_path, reg)
    assert su.main([]) == 0
    assert reg.writes[0][1] == FakeReg.REG_SZ                               # 원래 형식 유지
    reg2 = FakeReg(None)
    monkeypatch.setattr(su, "_winreg", lambda: reg2)
    assert su.main([]) == 0
    assert reg2.writes[0][1] == FakeReg.REG_EXPAND_SZ and reg2.value == str(cli / "python") + ";" + str(cli / "python" / "Scripts")


def test_path_check_and_no_path_do_not_write(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows")
    cli = _iso(monkeypatch, tmp_path, reg)
    su.main(["--no-path"])
    assert reg.writes == []
    capsys.readouterr()
    assert su.main(["--check"]) == 1
    assert reg.writes == [] and "시스템 PATH에 없음" in capsys.readouterr().out
    assert su.missing_path_entries(reg) == [str(cli / "python"), str(cli / "python" / "Scripts")]


def test_path_permission_denied_is_fail_without_write(monkeypatch, tmp_path, capsys):
    reg = FakeReg(r"C:\Windows", writable=False)
    cli = _iso(monkeypatch, tmp_path, reg)
    assert su.main([]) == 1
    out = capsys.readouterr().out
    assert reg.writes == [] and reg.value == r"C:\Windows"
    assert "[FAIL] 시스템 PATH를 쓰지 못함 (권한 없음)" in out and "관리자 권한으로 실행" in out


def test_not_windows_path_is_warning_only(monkeypatch, tmp_path, capsys):
    cli = _iso(monkeypatch, tmp_path, None)
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


def test_opencode_providers_checks_enabled_providers_only_when_given(tmp_path, capsys):
    """모델 목록 설정은 PC마다 사용자가 정함 (2026-11-13: 회사 PC는 disabled_providers로 opencode만 숨기고 집 PC는 보이게).
    enabled_providers가 없으면 알리지 않고, 적었을 때만 이름이 provider·model과 맞는지 봄."""
    cfg = tmp_path / "opencode.jsonc"
    for text in ('{"model": "corp/gemma", "provider": {"corp": {"npm": "@ai-sdk/openai-compatible"}}}',
                 '{"model": "corp/gemma", "provider": {"corp": {}}, "disabled_providers": ["opencode"]}'):
        cfg.write_text(text, encoding="utf-8")
        rep = su.Report()
        su.step_opencode_providers(rep, str(tmp_path))
        out = capsys.readouterr().out
        assert rep.warn == 0 and "enabled_providers" not in out and '"autoupdate": false' in out, out
    cfg.write_text('{"model": "corp/gemma", "provider": {"corp": {}}, "enabled_providers": ["corp"], "autoupdate": false}', encoding="utf-8")
    rep = su.Report()
    su.step_opencode_providers(rep, str(tmp_path))
    out = capsys.readouterr().out                             # 웹 도구를 막지 않았으면 INFO (폐쇄망에서 모델이 인터넷 검색을 시도함, 2026-11-11)
    assert rep.warn == 0 and '"webfetch": "deny", "websearch": "deny"' in out, out
    assert '"grep": "deny", "glob": "deny"' in out and '"share": "disabled"' in out, out   # 폐쇄망 잠금 (2026-11-12)
    cfg.write_text('{"model": "corp/gemma", "provider": {"corp": {}}, "enabled_providers": ["corp"], "autoupdate": false, '
                   '"share": "disabled", "lsp": false, "formatter": false, '
                   '"permission": {"external_directory": "allow", "webfetch": "deny", "websearch": "deny", "browser": "deny", '
                   '"grep": "deny", "glob": "deny"}}', encoding="utf-8")
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


def test_web_bat_is_ascii_crlf_and_opens_the_aikit_folder():
    """1.6.7 (사용자 결정): Web UI의 기본 프로젝트 폴더 = C:\\AIKIT (설정 폴더와 같음). 서비스를 이 폴더에서 띄우고 처음 열 폴더로 안내."""
    data = open(os.path.join(CLI, "web.bat"), "rb").read()
    assert data.isascii() and b"chcp" not in data.lower()
    assert data.count(b"\r\n") == data.count(b"\n") and data.endswith(b"\r\n")
    assert b'set "WORK=%OPENCODE_CONFIG_DIR%"' in data and b'cd /d "%WORK%"' in data
    assert b"work\"" not in data and b"mkdir \"%WORK%\"" not in data
    assert data.index(b'set "OPENCODE_CONFIG_DIR=%OPENCODE_CONFIG_DIR:~0,-1%"') < data.index(b'set "WORK=')


def test_web_bat_stops_service_when_web_ui_window_closes():
    """2026-11-13 (사용자 결정): Web UI 창을 닫으면 백그라운드 서비스를 끔. 화면·트레이에 아무것도 남기지 않는 숨은 PowerShell이
    web_profile을 쓰는 브라우저(사용자의 평소 브라우저는 프로필이 달라 세지 않음)가 없어지면 'opencode service stop'.
    cmd가 PowerShell 명령을 자르지 않게 큰따옴표 밖에 | & < > 가 없어야 함."""
    data = open(os.path.join(CLI, "web.bat"), "rb").read().decode("ascii")
    line = next(ln for ln in data.splitlines() if ln.startswith('start "" /min powershell'))
    assert "-WindowStyle Hidden" in line and "service stop" in line and "$p='%PROFILE%'" in line and "msedge.exe" in line
    assert data.index('--user-data-dir="%PROFILE%"') < data.index(line)          # 브라우저를 연 뒤에 감시
    inside, outside = False, []
    for ch in line:
        if ch == '"':
            inside = not inside
        elif not inside and ch in "|&<>":
            outside.append(ch)
    assert not inside and outside == [], outside


# ---- run_all_tests.py diag ----
def test_diag_fails_when_env_or_path_differ(tmp_path, monkeypatch):
    """run_all_tests.py diag: setup.py --check처럼 읽기만. PATH, OPENCODE_CONFIG_DIR."""
    rat = _load("run_all_tests_s", os.path.join(DEV, "run_all_tests.py"))
    lines = []
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, reg=None)
    if os.name != "nt":                                                   # Windows가 아니면 참고만
        assert st.failed == 0 and st.passed == 0 and any("Windows가 아니라" in x for x in lines)
    paths = ";".join(os.path.join(CLI, *p) for p in (("python",), ("python", "Scripts")))
    want = CLI
    reg = FakeReg(r"C:\Windows;" + paths)
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, reg=reg)
    assert st.failed == 1 and "OPENCODE_CONFIG_DIR가 없음" in st.failures[0] and "setup.bat" in st.failures[0]
    reg.user[su.CONFIG_ENV] = want
    st = rat.Stage("diag")
    rat.diag_skills(st, lines, reg=reg)
    assert st.failed == 0 and st.passed == 2 and reg.writes == [] and reg.user == {su.CONFIG_ENV: want}   # 읽기만


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
