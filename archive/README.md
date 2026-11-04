# archive – 보관 폴더

더 이상 이 저장소에서 개발하지 않는 도구를 보관합니다. 기록용이며 다음 규칙을 따릅니다.

- 에이전트는 이 폴더의 파일을 **수정하지 않습니다**. `AGENTS.md` 지침 반영, 준수 검토, 테스트(`pytest.ini`의 `testpaths = CLI`) 대상이 아닙니다.
- 폴더 안 문서의 경로(`CLI/<tool>/...`)는 보관 당시 위치 기준이라 지금 위치와 다릅니다.
- 다시 개발하려면 `CLI/`로 옮기고 `AGENTS.md`의 "도구별 규칙"과 `CLI/README.md` 도구 목록에 추가합니다.

| 도구 | 보관일 | 마지막 버전 | 사유 |
|---|---|---|---|
| [`engenius/`](engenius/) | 2026-11-04 | office-cli 1.4.36, ml-kit 2.9.38 기준 | 사내 에이전트 En-Genius 스킬 패키지. 업데이트를 멈추고 보관 (필요할 때 다시 꺼냄). `build/`(빌드 스크립트·도구별 run.py, 원래 `CLI-dev/engenius`), `package/`(zip·에이전트 XML·설치 안내, 원래 `docs/engenius`), `test_engenius.py`(원래 `CLI-dev/tests`). 다시 쓰려면 원래 자리로 옮기고 `python CLI-dev/engenius/build.py`로 zip을 새로 만든 뒤 테스트 |
| [`doc-classifier/`](doc-classifier/) | 2026-10-28 | 0.9.2 | 다른 개발자가 따로 개발하기로 함. 0.9.3 작업(PR #12, 병합 안 함)은 브랜치를 지웠고, 닫힌 PR #12 화면에서 변경 내용을 볼 수 있음 |
