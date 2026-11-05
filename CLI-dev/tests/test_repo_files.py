# -*- coding: utf-8 -*-
"""반입 확장자 규칙 (AGENTS.md "폐쇄망 / Windows 제약"): CLI/·CLI-dev/에 저장소로 두는 파일은 회사 반입 시스템이 통과시키는
확장자(.py .md .txt .json .tsv .bat)만. .yaml/.yml/.jsonl/.toml 등은 반입에서 막혀 PC에 없게 된다 (2026-11-05: agent_scenarios.yaml).
실행 중 PC에서 생기는 파일(history.jsonl 등)은 git에 없으므로 대상이 아니다."""
import os
import subprocess

import pytest

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ALLOWED = {".py", ".md", ".txt", ".json", ".tsv", ".bat"}


def _tracked():
    try:
        proc = subprocess.run(["git", "ls-files", "-z", "--", "CLI", "CLI-dev"], cwd=REPO, capture_output=True, timeout=60)
    except (OSError, subprocess.SubprocessError) as e:
        pytest.skip(f"git 없음 ({type(e).__name__}): 저장소에서만 확인")
    if proc.returncode:
        pytest.skip("git 저장소가 아님 (반입된 PC): 저장소에서만 확인")
    return [p for p in proc.stdout.decode("utf-8").split("\0") if p]


def test_tracked_files_use_only_importable_extensions():
    files = _tracked()
    assert any(p.endswith("agent_scenarios.py") for p in files)
    bad = [p for p in files if os.path.splitext(p)[1].lower() not in ALLOWED]
    assert not bad, f"회사 반입에서 막히는 확장자 (AGENTS.md 반입 규칙): {bad}"
