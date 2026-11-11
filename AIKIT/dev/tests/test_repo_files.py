# -*- coding: utf-8 -*-
"""반입 확장자 규칙 (AGENTS.md "폐쇄망 / Windows 제약"): AIKIT/·AIKIT/dev/에 저장소로 두는 파일은 회사 반입 시스템이 통과시키는
확장자(.py .md .txt .json .tsv .bat)만. .yaml/.yml/.jsonl/.toml 등은 반입에서 막혀 PC에 없게 된다 (2026-11-05: agent_scenarios.yaml).
실행 중 PC에서 생기는 파일(history.jsonl 등)은 git에 없으므로 대상이 아니다."""
import os
import subprocess

import pytest

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
ALLOWED = {".py", ".md", ".txt", ".json", ".tsv", ".bat"}
GUIDE_ALLOWED = {".html"}        # AIKIT/guide/의 사람용 사용 가이드만 (2026-11-12 사용자 결정: 회사 반입 통과)


def _tracked():
    try:
        proc = subprocess.run(["git", "ls-files", "-z", "--", "AIKIT"], cwd=REPO, capture_output=True, timeout=60)
    except (OSError, subprocess.SubprocessError) as e:
        pytest.skip(f"git 없음 ({type(e).__name__}): 저장소에서만 확인")
    if proc.returncode:
        pytest.skip("git 저장소가 아님 (반입된 PC): 저장소에서만 확인")
    return [p for p in proc.stdout.decode("utf-8").split("\0") if p]


def test_tracked_files_use_only_importable_extensions():
    files = _tracked()
    assert any(p.endswith("agent_scenarios.py") for p in files)
    bad = [p for p in files if os.path.splitext(p)[1].lower() not in ALLOWED
           and not (p.startswith("AIKIT/guide/") and os.path.splitext(p)[1].lower() in GUIDE_ALLOWED)]
    assert not bad, f"회사 반입에서 막히는 확장자 (AGENTS.md 반입 규칙): {bad}"


def test_user_guides_are_in_aikit_guide():
    """사람용 사용 가이드는 C:\\AIKIT\\guide에 함께 배포 (2026-11-12): 도구별 HTML 하나씩, 옛 스킬 이름·옛 배포 폴더가 없음."""
    guide = os.path.join(REPO, "AIKIT", "guide")
    names = sorted(os.listdir(guide))
    assert names == ["ml-kit_사용가이드.html", "office-cli_사용가이드.html"], names
    for name in names:
        with open(os.path.join(guide, name), encoding="utf-8") as f:
            text = f.read()
        for old in ("office-doc", "office-excel", "skill-maker", "skill-kit", "C:\\CLI", "docs/guide"):
            assert old not in text, (name, old)
