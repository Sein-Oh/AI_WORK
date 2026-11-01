# 작업 인계 (2026-10-30, 클라우드 세션 -> 집 PC)

집 PC에서 저장소를 받은 뒤, 새 Claude Code 대화를 열고 맨 아래 "붙여넣을 첫 메시지"를 그대로 보내면 이어서 진행할 수 있습니다.
이 파일은 작업이 이어지면 최신 상태로 고치거나 지웁니다.

## 1. 지금 상태
- main 기준 도구 버전: **ml-kit 2.9.11**, **office-cli 1.4.10** (집 PC 실기 수정 포함)
- 저장소 규칙은 `AGENTS.md` 한 곳 (Claude Code는 `CLAUDE.md`가 불러옴). 특히:
  - 답변은 한국어, 이모지·장식 기호 금지 (CP949 안전 문자만)
  - 베타 단계: 이전 버전 호환 코드를 넣지 않음. CLI 인자·JSON 키는 바꿔도 되고, 바꾸면 SKILL.md·README·CHANGELOG·테스트를 같은 커밋에서
  - 리팩터링과 기능 변경은 별도 커밋/PR, 버그 수정에는 회귀 테스트
  - `archive/`는 수정하지 않음
- 배포 파일에 사용자 이름이 든 전체 경로를 넣지 않음 (성능예측표 등은 `<데이터>_ml_data\vN`처럼 끝부분만)
- 실기 로그는 사용자가 GitHub 웹으로 `docs/test_logs/`에 올림 (git 명령으로는 올리지 않음, `.gitignore` 처리). 프로젝트 마무리 때 저장소 이력에서 정리 예정

## 2. 최근 바뀐 것 (실기에서 볼 것)
| 버전 | 내용 |
|---|---|
| ml-kit 2.9.6 | 성능예측표 학습 범위를 학습에 쓴 행 기준으로, 값이 하나뿐인 입력 제외, 입력 8개 이상 수식(범위 합계), `제외` 열 0/FALSE는 표시 없음 |
| ml-kit 2.9.7 | 이전 버전 호환 코드 삭제: 2.9.1 이전 모델은 "다시 학습" 안내 (STOP_RETRY) |
| ml-kit 2.9.8 | 예측 결과 파일·보고서를 다시 읽어 확인 (R1) |
| office-cli 1.4.5 | `excel write`: 모든 검사 후 쓰기, **값이 있는 칸은 사용자 확인 후 `--overwrite`**, 인자 오류에 `[추천]` |
| office-cli 1.4.6~1.4.8 | 모든 쓰기 명령이 결과를 다시 읽어 확인한 경우에만 `VERIFIED:`, 아니면 `UNVERIFIED:` + `unverified`. `_runtime.ok()`는 VERIFIED를 대신 붙이지 않음 |

## 3. 집 PC 실기 테스트 (Windows 11, Python 3.13.13, Office 2007)
상세: `docs/opencode/실기테스트_2026-10-30.md` 부록 "집 PC(Office 2007)로 먼저 할 때".

1. `CLI` -> `C:\CLI`, `CLI-dev` -> `C:\CLI-dev` 복사 후 설치:
   `python -m pip install -r C:\CLI-dev\ml-kit\requirements.txt -r C:\CLI-dev\office-cli\requirements.txt`
2. 자동 점검:
   ```powershell
   cd C:\CLI-dev
   python run_beta_check.py --label 집PC
   python run_all_tests.py --label 집PC --auto
   ```
3. 눈 확인: 실기테스트 문서 4단계 V1~V7, V9~V10
4. 결과 `beta_check_집PC_*.md`, `test_result_집PC_*.txt`, V 항목 O/X

판정에서 특히 볼 것:
- `test_result_*.txt`의 office 단계 끝 **"UNVERIFIED 결과"** 목록: 비어 있어야 정상. 있으면 해당 명령의 다시 읽기 비교 규칙(`excel.py _verify_written/_verify_style/_same_cell`, `word.py _verify_inserted`, `ppt.py _verify_slides` 등)을 Office 2007 동작에 맞게 고침
- `excel range`의 너비·숨김·틀 고정 확인은 클라우드(가짜 Excel)에서 못 해 봄
- `run_beta_check.py` F7(엑셀 재계산): Excel 2007에서 성능예측표 수식(INDEX/MATCH, IFERROR) 계산
- PDF 내보내기는 Office 2007 추가 기능이 없으면 ENV-SKIP (오류 아님)

집 PC에서 확인 못 하는 것 (회사 PC 몫): 문서 보안(DRM), Office 16 차트 전달 방식, Gemma 종합 테스트(`docs/opencode/자동테스트/` + `CLI-dev/grade_opencode.py`), 파이썬 없는 PC에서 성능예측표 열기(V8)

## 3-1. 집 PC 자동 점검 결과 (2026-10-31, Office 2007)
- `run_beta_check.py`: 27개 항목 모두 OK (F7 Excel 2007 수식 재계산 포함)
- `run_all_tests.py --auto`: pytest 통과, UNVERIFIED 결과 0건, ENV-SKIP은 D1·D4 (PDF 추가 기능 없음, 오류 아님)
- 수정 후 실제 Office로 다시 통과: M6·M7 (보고서가 Word 2007에서 열림), M10 (2만 행 예측 900초 초과 -> 40.7초), A3 (닫힌 Word 파일 쓰기)
- 2.9.10으로 전체 재실행: pytest 516 passed·4 skipped, ml-kit 실기 14/14, office-cli 실기 PASS 69·ENV-SKIP 2, UNVERIFIED 0건
- 눈 확인: V1~V7, V9~V10 통과 (`docs/test_logs/눈확인_집PC_20261031.md`)

## 4. 남은 일
- 집 PC·회사 PC 실기 결과 분석과 수정
- 회사 PC: Gemma 종합 테스트 다시 (X2 시나리오가 "이상치 기준을 바꾸는 옵션 있어?"로 바뀜), F2(보고서·성능예측표 같은 폴더) 재확인
- 각 도구 `CLI-dev/<tool>/CHANGELOG.md` 하단 "남은 할 일" 참고 (HTML 성능예측표는 하지 않기로 결정해 목록에서 뺌)
- 프로젝트 마무리 때 `docs/test_logs/` 등 로그 이력 정리 (저장소를 다시 만드는 방식)

## 5. 붙여넣을 첫 메시지
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md를 읽고 지금 상태를 파악해.
지금은 집 PC(Windows 11, Python 3.13.13, Office 2007)에서 실기 테스트를 하려고 해.
1. HANDOFF.md 3장대로 설치와 자동 점검(run_beta_check.py, run_all_tests.py --auto)을 같이 진행해 줘. 명령은 python만 쓰고, 실행 전에 무엇을 할지 짧게 알려 줘.
2. 끝나면 결과 파일을 읽고 실패·ENV-SKIP·"UNVERIFIED 결과"를 표로 정리해 줘. 원인 추정과 고칠 곳(파일:줄)도.
3. 고칠 것은 바로 고치지 말고 계획을 먼저 보여 줘. 승인하면 브랜치를 만들어 수정·테스트(python -m pytest -q) 후 PR을 올려 줘.
답변은 한국어로, 이모지 없이.
```
