# 작업 인계 (2026-11-12)

지금 상태와 할 일만 둔다. 원칙 결정은 [`DECISIONS.md`](DECISIONS.md), 지난 경위(이 파일의 옛 3-x장, 실기 결과 이력)는
[`history/HANDOFF_2026-11-12.md`](history/HANDOFF_2026-11-12.md) (CHANGELOG의 "HANDOFF 3-25" 같은 참조는 그 파일의 절).

## 1. 지금 상태
- 버전: ml-kit 2.9.92, office-cli 1.6.7. 작업 브랜치 `ccr-ecab210c-awib23` (머지는 사용자가 정함).
- 배포: `C:\AIKIT` 한 폴더 (도구, `skills\`, `dev\`, `guide\`, `AGENTS.md`, `opencode.json`, `setup.bat`, `web.bat`). `C:\AIKIT`이 OpenCode 설정 폴더.
- 기본 스킬 12개: excel-kit, excel-format-kit, word-kit, ppt-kit, outlook-kit, search-kit, stats-kit, ml-train, ml-dataset, ml-predict, ml-report, ml-map.
- 기준 모델 Gemma (합격 기준), BEAM도 씀. 시험: 클라우드 pytest, 집·회사 PC 실기 러너와 눈 확인, Gemma 시나리오.
- 1.6.7 새 기반 기능(엑셀 서식 칸·이름 정의·자동 필터·그림·링크, 워드 머리글·링크·책갈피, PPT 표, Outlook 연락처, `office.py open`)은
  실기 러너 P 그룹으로 실제 Office에서 확인할 차례 (집 PC Office 2019 회귀는 통과, 3장).

## 2. 회사 PC에서 할 일
상세: [`opencode/실기테스트.md`](opencode/실기테스트.md).
1. 준비: `C:\AIKIT`의 도구 폴더(`ml-kit`, `office-cli`), `dev`, 기본 스킬 12개 폴더, `guide`, `AGENTS.md`, `opencode.json`, `setup.py`·`setup.bat`·`web.bat`을
   새로 붙여넣음 (`python\`, `opencode.exe`, `opencode.jsonc`, 사용자 스킬 폴더, `data\`는 둠). `setup.bat`(관리자) -> 새 터미널·OpenCode 다시 시작.
2. 자동 점검 (질문 없음): `cd C:\AIKIT\dev` -> `python run_all_tests.py --label 회사PC --outlook`
   (환경 진단 -> ml-kit 베타 점검 -> office-cli 실기 러너(P 그룹 포함, `--outlook`이면 P10 연락처까지) -> 도구 오류 로그). 결과 `result_회사PC_<시각>.md`.
3. Gemma 종합 테스트(선택): 2의 명령에 `--gemma` (일반 권한 창).
4. 눈 확인: 실기테스트.md 4장 (집 PC가 못 하는 2번 문서 보안, 9·10번 Outlook).
5. P 그룹이 만들지 않는 상황은 회사 문서로 볼 것: 워드 머리글 구역 연결, Exchange 주소록 이름 확인, 보호된 보기로 열린 파일의 `open`.
6. 실기 러너 종료 정리(남은 Excel 끝내기)가 회사 PC에서도 되는지: 결과 파일의 정리 기록과 끝난 뒤 남은 Excel이 없는지.
7. 결과 파일과 눈 확인 O/X를 GitHub 웹으로 `docs/test_logs/`에 올림.

## 3. 최근 실기 결과
- 집 PC (2026-11-12 저녁, ml-kit 2.9.92 / office-cli 1.6.7, Office 2019): diag 17/17, beta 29/29, ml 14/14, office 99/100(L3 SORT는 2019에 없는 함수라
  러너를 ENV-SKIP으로 고침), 단순성 점검 34/34, 눈 확인 1·3~8번 통과. 스킬 12개 보임, `C:\AIKIT`을 프로젝트로 열어도 `AGENTS.md`는 한 번만 들어감.
- 회사 PC (2026-11-11 저녁, ml-kit 2.9.90 / office-cli 1.6.3): office 실기 PASS 102, 사내 LLM 시험 Gemma 15/21·BEAM 18/21 (채점 기준 수정 뒤 16/21·20/21).
- 그 전 결과는 옛 인계 문서 4장.

## 4. 남은 일
- 집 PC BOM 비교 시험 (사용자 조합 스킬 bom-compare, 정답 아는 샘플): [`opencode/BOM비교_시험.md`](opencode/BOM비교_시험.md).
- 각 도구 `AIKIT/dev/<tool>/CHANGELOG.md` 하단 "남은 할 일".
- 사용 가이드(`AIKIT/guide/`)에 1.6.7 기반 기능 사용 예 넣기.
- 프로젝트 마무리 때 `docs/test_logs/` 등 로그 이력 정리 (저장소를 다시 만드는 방식).
- 보류 항목은 `DECISIONS.md` 끝.

## 5. 붙여넣을 첫 메시지
클라우드용:
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md, docs/DECISIONS.md, docs/HANDOFF.md를 읽어. 작업 브랜치는 ccr-ecab210c-awib23.
답변은 한국어로, 이모지 없이.
```

회사 PC용:
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md를 읽고 지금 상태를 파악해.
지금은 회사 PC(Windows 11, WinPython 3.13.13, Office 16, 문서 보안)에서 실기 테스트를 하려고 해.
1. HANDOFF 2장대로 준비를 같이 하고, 자동 점검(C:\AIKIT\dev에서 python run_all_tests.py --label 회사PC --outlook)을 실행해 줘. 명령은 python만 쓰고, 실행 전에 무엇을 할지 짧게 알려 줘.
2. 결과 파일 맨 위의 요약 표와 실패·ENV-SKIP·UNVERIFIED 목록을 표로 정리하고 원인 추정과 고칠 곳(파일:줄)을 찾아 줘.
3. 고칠 것은 바로 고치지 말고 계획을 먼저 보여 줘.
답변은 한국어로, 이모지 없이.
```
