# 작업 인계 (2026-11-02, 클라우드 세션 -> 회사 PC)

회사 PC에서 저장소 파일을 받은 뒤, 새 Claude Code(또는 OpenCode) 대화를 열고 맨 아래 "붙여넣을 첫 메시지"를 보내면 이어서 진행할 수 있습니다.
이 파일은 작업이 이어지면 최신 상태로 고치거나 지웁니다.

## 1. 지금 상태
- main 기준 도구 버전: **ml-kit 2.9.14**, **office-cli 1.4.12**
- 저장소 규칙은 `AGENTS.md` 한 곳 (Claude Code는 `CLAUDE.md`가 불러옴). 특히:
  - 답변은 한국어, 이모지·장식 기호 금지 (CP949 안전 문자만)
  - 베타 단계: 이전 버전 호환 코드를 넣지 않음. CLI 인자·JSON 키는 바꿔도 되고, 바꾸면 SKILL.md·README·CHANGELOG·테스트를 같은 커밋에서
  - 리팩터링과 기능 변경은 별도 커밋/PR, 버그 수정에는 회귀 테스트
  - `archive/`는 수정하지 않음
- 배포 파일에 사용자 이름이 든 전체 경로를 넣지 않음 (성능예측표 등은 `<데이터>_ml_data\vN`처럼 끝부분만)
- 실기 로그는 사용자가 GitHub 웹으로 `docs/test_logs/`에 올림 (git 명령으로는 올리지 않음, `.gitignore` 처리). 프로젝트 마무리 때 저장소 이력에서 정리 예정
- 클라우드 테스트: 개발 도구 테스트 통과 (2026-11-02, 3차 점검). 전체 수는 회사 PC `run_all_tests.py`의 pytest 결과로 확인

## 2. 최근 바뀐 것 (클라우드 전체 점검 반영, 실기에서 볼 것)
| 버전 | 내용 |
|---|---|
| office-cli 1.4.10 | `ok()`가 `verified=False`를 거부하고 저장 파일은 형식별로 다시 열어 확인. `excel replace`가 `--range`의 시트에 씀. 결과가 ""인 수식 칸도 덮어쓰기 확인. `ppt write --title`로 다른 제목을 바꾸려면 `--overwrite`. 파일 이름만 준 쓰기 대상은 새로 만들지 않음(전체 경로만). reader 잘림·글자 없는 PDF 쪽을 알림. 오류 로그에서 본문·받는 사람 등 내용 값은 `<N자 생략>` |
| office-cli 1.4.11 | 문서 색인 형식이 바뀜(PARSER_VERSION 6): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행 |
| office-cli 1.4.12 | 문서 색인 형식이 다시 바뀜(PARSER_VERSION 7, PDF 스캔 쪽 판정): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행. 동시에 실행한 `index`도 색인 항목을 잃지 않음 |
| ml-kit 2.9.11 | 날짜 열을 학습·예측에서 같은 방식으로 해석(YYYYMMDD 포함). `-i` 값은 글자로 받아 범주 코드 `007` 유지. 머리글·단위 행 옵션이 예측 `next_commands`에도 들어감. `-i` 6건 이상은 결과 CSV 저장 |
| ml-kit 2.9.12 | 성능예측표: 작은 값 출력의 자리수, "실측 없는 영역" 기준을 표 해상도에 맞춤, **입력은 타겟마다 최대 8개**, 타겟별 `--inputs "T1=a,b;T2=c"`. `ml_export`·`ml_map`에 `--overwrite`. `ml_map`은 열마다 그 열에 값이 있는 행으로 계산. 예측값은 작은 값이면 유효숫자 4자리 |
| ml-kit 2.9.13 | 타겟 값이 모두 같으면 `STOP_RETRY`(학습 안 함). 완전히 같은 행은 교차검증에서 같은 fold로 묶고 `warnings`에 중복 개수. **`[추천]` 뒤가 `python`으로 시작하는 전체 명령이면 그대로 1회 실행**(레이아웃 옵션을 뺀 예측, 성능예측표 `--inputs` 전체. 전역 지침 `docs/opencode/AGENTS.md`도 바뀜). `ml_map` 여러 구할 열은 `basis_by_column`. 작은 입력값(1e-6 단위)의 학습 범위·`--find` 결과 |
| ml-kit 2.9.14 | 숫자 입력에 쉼표(`-i "압력=500,000"`)가 있으면 `STOP_RETRY` + 쉼표를 뺀 전체 명령 `[추천]`. 예측 JSON의 `drift_warnings` 키 삭제(`extrapolation_warnings`는 20개까지, 전체 수 `extrapolation_warning_count`). 한 버전이 모든 타겟의 최고면 `ml_report -f`가 다시 학습하지 않고 그 버전을 씀(`made_best: false`). 머리글 행 안내는 열 이름 `[추천]`/`[후보]` 또는 `--header-row`. 입력+타겟이 다른 행이 10개 미만이면 `STOP_RETRY` |
| 개발 도구 | 내보내기는 `cmd /c "opencode export <세션ID> > chat.json"` (PowerShell 5.1의 `\| Out-File`이나 `>`는 한글을 깨뜨림, 채점기가 깨진 대화를 찾으면 [ERROR]). "한 번에 보내기"는 시나리오별로 나눠 채점(참고). 최종 답변은 시나리오 명령 뒤의 글. 러너 `--agent` 모드 JSON 계약(도움말 포함 `verified`), office 러너 `unverified_count`. 3차: 한 번에 보내기에서 가운데 시나리오의 재실행이 다음 시나리오로 넘어가지 않음, X2 답변 정규식이 틀린 답("바꿀 수 있습니다")을 받지 않음, `C:\CLI-dev`에서 이름만 쓴 러너 실행·glob 파일 목록·`*.md` grep은 위반 아님 |
| 원격 에이전트 | `docs/remote-agent/run.py`가 바뀜: 사용하는 팀원은 한 번 다시 올림 |

## 3. 회사 PC에서 할 일
상세: `docs/opencode/실기테스트_2026-10-30.md` (1. 준비 -> 2. 자동 점검 -> 3. Gemma 종합 테스트 -> 4. 눈으로 확인).

1. 준비: `C:\CLI`, `C:\CLI-dev`, OpenCode 스킬(`SKILL.md` 2개), 전역 지침(`docs/opencode/AGENTS.md`)을 새 파일로 바꾸고 OpenCode 다시 시작
2. 자동 점검:
   ```powershell
   cd C:\CLI-dev
   python run_beta_check.py --label 회사PC
   python run_all_tests.py --label 회사PC --only diag,pytest
   ```
   - F2(보고서·성능예측표 같은 폴더)와 F7(엑셀 재계산)을 꼭 확인
3. Gemma 종합 테스트: `docs/opencode/자동테스트/` 메시지를 보내고, 내보낸 대화를 `python C:\CLI-dev\grade_opencode.py --export chat.json --label 회사PC`로 채점
4. 눈 확인: V1~V10. 집 PC에서 못 한 V8(파이썬 없는 PC에서 성능예측표 열기, 문서 보안 아래)이 핵심
5. 결과 파일과 V 항목 O/X를 GitHub 웹으로 `docs/test_logs/`에 올림

회사 PC에서만 확인할 수 있는 것: 문서 보안(DRM) 영향, Office 16 차트 전달 방식, Gemma의 요청 처리, 사내 API 채점(`--judge`).

## 4. 지난 실기 결과 요약
- 집 PC(2026-10-31, Office 2007, ml-kit 2.9.10 / office-cli 1.4.9): beta check 27/27(F7 포함), pytest 통과, ml-kit 실기 14/14, office-cli 실기 PASS 69·ENV-SKIP 2(PDF 추가 기능 없음), UNVERIFIED 0건, 눈 확인 V1~V7·V9~V10 통과
- 회사 PC(2026-10-29~30): beta check 26/27(F2), Gemma 12/14(X1·X2에서 도구 코드 읽기 -> 지침 보강, X2 시나리오 변경)

## 5. 남은 일
- 회사 PC 실기 결과 분석과 수정
- 각 도구 `CLI-dev/<tool>/CHANGELOG.md` 하단 "남은 할 일"
- 프로젝트 마무리 때 `docs/test_logs/` 등 로그 이력 정리 (저장소를 다시 만드는 방식)
- 도구 개발을 마친 뒤 다시 판단하기로 한 검토 (2026-11-01, 지금은 하지 않음):
  - MCP 전환: CLI 유지. 다시 본다면 CLI를 하위 프로세스로 부르는 얇은 MCP 층부터
  - 전용 에이전트(직접 만든 도구 호출 루프): 회사 Gemma로 OpenCode와 같은 시나리오를 비교한 뒤 결정
  - ml-kit 웹 서버: 동료 수요(예측이면 성능예측표로 충분)와 IT 정책·문서 보안 확인 후
  - `opencode run "<요청>"`으로 Gemma 종합 테스트 자동화: 회사 PC OpenCode(Desktop)에서 `opencode --version`, `opencode run "안녕"`이 되는지 먼저 확인

## 6. 붙여넣을 첫 메시지
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md를 읽고 지금 상태를 파악해.
지금은 회사 PC(Windows 11, WinPython 3.13.13, Office 16, 문서 보안)에서 실기 테스트를 하려고 해.
1. HANDOFF.md 3장대로 준비와 자동 점검(run_beta_check.py, run_all_tests.py)을 같이 진행해 줘. 명령은 python만 쓰고, 실행 전에 무엇을 할지 짧게 알려 줘.
2. 끝나면 결과 파일을 읽고 실패·ENV-SKIP·"UNVERIFIED 결과"를 표로 정리해 줘. 원인 추정과 고칠 곳(파일:줄)도.
3. 고칠 것은 바로 고치지 말고 계획을 먼저 보여 줘.
답변은 한국어로, 이모지 없이.
```
