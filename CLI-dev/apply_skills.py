r"""OpenCode 스킬과 전역 지침 적용 (사람이 실행하는 개발 도구, 반입 후 한 번)

사용법 (C:\CLI-dev에서):
    python apply_skills.py            # 바뀐 파일만 복사 (이전 파일은 work\skill_backup_<시각>에 보관)
    python apply_skills.py --check    # 복사하지 않고 무엇이 다른지만 보여 줌 (다르면 exit 1)
    python apply_skills.py --test     # 실기 테스트용 보조 스킬(office-cli-test)도 함께

복사하는 파일 (원본 -> OpenCode 설정 폴더 %USERPROFILE%\.config\opencode):
    C:\CLI\ml-kit\SKILL.md                              -> skills\ml-kit\SKILL.md
    C:\CLI\office-cli\SKILL.md                          -> skills\office-cli\SKILL.md
    C:\CLI-dev\opencode\AGENTS.md                       -> AGENTS.md (전역 지침)
    C:\CLI-dev\office-cli\tools\office-cli-test\SKILL.md -> skills\office-cli-test\SKILL.md (--test일 때만)
설정 폴더가 다르면 --config-dir. 끝나면 OpenCode를 다시 시작해야 새 지침이 적용됩니다.

함께 알려 주는 것: 스킬 폴더에 남은 예전 도구 파일(.py), rg.exe(ripgrep)가 PATH에 있는지.
이 스크립트는 에이전트용 출력 계약(stdout 1줄 JSON) 대상이 아닙니다. exit 0 = 완료(또는 --check에서 모두 같음), 1 = 원본 없음 또는 다름.
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))                   # CLI-dev
CLI_DIR = os.path.join(os.path.dirname(HERE), "CLI")               # CLI-dev 옆의 CLI


def default_config_dir():
    return os.path.join(os.path.expanduser("~"), ".config", "opencode")


def targets(include_test=False):
    """(이름, 원본, 설정 폴더 기준 대상 상대 경로) 목록."""
    items = [("ml-kit 스킬", os.path.join(CLI_DIR, "ml-kit", "SKILL.md"), os.path.join("skills", "ml-kit", "SKILL.md")),
             ("office-cli 스킬", os.path.join(CLI_DIR, "office-cli", "SKILL.md"), os.path.join("skills", "office-cli", "SKILL.md")),
             ("전역 지침", os.path.join(HERE, "opencode", "AGENTS.md"), "AGENTS.md")]
    if include_test:
        items.append(("office-cli-test 스킬", os.path.join(HERE, "office-cli", "tools", "office-cli-test", "SKILL.md"),
                      os.path.join("skills", "office-cli-test", "SKILL.md")))
    return items


def _read(path):
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return None


def version_of(data):
    """SKILL.md 머리말의 version 값 (없으면 빈 글)."""
    m = re.search(rb'^\s*version:\s*"?([^"\r\n]+)"?', data or b"", re.M)
    return m.group(1).decode("utf-8", "replace").strip() if m else ""


def compare(config_dir, include_test=False):
    """[(이름, 원본, 대상, 상태)]. 상태: 같음 / 다름 / 새 파일(대상 없음) / 원본 없음."""
    out = []
    for name, src, rel in targets(include_test):
        dst = os.path.join(config_dir, rel)
        a, b = _read(src), _read(dst)
        state = "원본 없음" if a is None else "새 파일" if b is None else "같음" if a == b else "다름"
        out.append((name, src, dst, state))
    return out


def leftovers(config_dir):
    """스킬 폴더에 남은 예전 도구 파일 (.py). 두 곳에 도구가 있으면 어느 쪽이 쓰였는지 헷갈림."""
    found = []
    for tool in ("ml-kit", "office-cli"):
        folder = os.path.join(config_dir, "skills", tool)
        for root, _, files in os.walk(folder):
            found += [os.path.join(root, f) for f in files if f.lower().endswith(".py")]
    return found


def label(name, src, dst, state):
    a, b = version_of(_read(src)), version_of(_read(dst))
    ver = f" ({b or '?'} -> {a})" if state == "다름" and a else (f" ({a})" if a else "")
    return f"{name}: {state}{ver}  {dst}"


def apply(config_dir, include_test=False, now=None):
    """바뀐 파일만 복사합니다. 덮어쓰는 파일은 work\\skill_backup_<시각>에 먼저 보관합니다. (줄 목록, 원본 없음 수)"""
    stamp = f"{(now or dt.datetime.now()):%Y%m%d_%H%M%S}"
    backup = os.path.join(HERE, "work", f"skill_backup_{stamp}")
    lines, missing = [], 0
    for name, src, dst, state in compare(config_dir, include_test):
        if state == "원본 없음":
            missing += 1
            lines.append(f"[ERROR] {name}: 원본이 없습니다: {src}")
            continue
        if state == "같음":
            lines.append(f"[INFO] {label(name, src, dst, state)}")
            continue
        if state == "다름":
            keep = os.path.join(backup, os.path.relpath(dst, config_dir))
            os.makedirs(os.path.dirname(keep), exist_ok=True)
            shutil.copy2(dst, keep)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        tmp = dst + ".tmp"
        shutil.copyfile(src, tmp)
        os.replace(tmp, dst)
        ok = _read(dst) == _read(src)
        lines.append(f"[{'OK' if ok else 'FAIL'}] {label(name, src, dst, state)}" + (" -> 복사함" if ok else " -> 복사 후 내용이 다름"))
        missing += 0 if ok else 1
    if os.path.isdir(backup):
        lines.append(f"[INFO] 이전 파일 보관: {backup}")
    return lines, missing


def main(argv=None):
    p = argparse.ArgumentParser(description="OpenCode 스킬(SKILL.md)과 전역 지침(AGENTS.md)을 C:\\CLI, C:\\CLI-dev에서 설정 폴더로 복사")
    p.add_argument("--check", action="store_true", help="복사하지 않고 다른 것만 보여 줌 (다르면 exit 1)")
    p.add_argument("--test", action="store_true", help="실기 테스트용 보조 스킬 office-cli-test도 함께")
    p.add_argument("--config-dir", default=default_config_dir(), help=r"OpenCode 설정 폴더 (기본 %%USERPROFILE%%\.config\opencode)")
    args = p.parse_args(argv)
    cfg = os.path.abspath(args.config_dir)
    print(f"[INFO] OpenCode 설정 폴더: {cfg}")
    if args.check:
        rows = compare(cfg, args.test)
        for row in rows:
            print(("[OK] " if row[3] == "같음" else "[WARN] ") + label(*row))
        bad = [r for r in rows if r[3] != "같음"]
        if bad:
            print("[STATUS] 다른 파일이 있습니다. 적용: python apply_skills.py" + (" --test" if args.test else ""))
        code = 1 if bad else 0
    else:
        if not os.path.isdir(cfg):
            print("[WARN] 설정 폴더가 없어 새로 만듭니다 (OpenCode를 한 번도 실행하지 않았거나 다른 위치일 수 있음: --config-dir)")
        lines, missing = apply(cfg, args.test)
        print("\n".join(lines))
        code = 1 if missing else 0
        if not missing:
            print("[STATUS] 끝. OpenCode(TUI, Desktop)를 모두 닫고 다시 시작하세요.")
    for path in leftovers(cfg):
        print(f"[WARN] 스킬 폴더에 예전 도구 파일이 있습니다 (도구는 C:\\CLI만 쓰므로 지우세요): {path}")
    if not shutil.which("rg"):
        print("[WARN] rg(ripgrep)가 PATH에 없습니다. OpenCode가 스킬을 못 불러올 수 있습니다 (docs/opencode/README.md 1장).")
    return code


if __name__ == "__main__":
    sys.exit(main())
