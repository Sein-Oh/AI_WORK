# AGENTS.md - AI_WORK 저장소에서 작업하는 모든 AI 에이전트용 규칙

OpenCode, Claude Code(CLAUDE.md가 이 파일을 불러옴) 등 이 저장소에서 작업하는 모든 에이전트가 따르는 단일 원본입니다.

## 저장소 성격
- `CLI/<tool>/` 의 스크립트는 **Windows 11 폐쇄망에서 26B~30B 경량 LLM 에이전트가 스킬로 호출하는 CLI 도구**(Agent-Computer Interface, ACI)입니다.
- 소비자는 사람이 아닌 경량 LLM입니다. 출력 형식의 작은 변화도 에이전트를 깨뜨립니다.
- 설계 원칙: **LLM 개입 최소화** - 판단과 복구는 도구가 하고, LLM은 명령 1회 실행 + JSON 1줄 해석만 한다.
- **베타 단계**: ml-kit과 office-cli는 베타다. 이전 버전과의 호환은 필요 없다 (사용자가 따로 요청할 때만).
  호환을 위한 코드(옛 인자·키 별칭, 이전 형식 모델·파일을 읽는 대체 경로)를 새로 넣지 않는다.

## 폴더 구조
- `CLI/<tool>/`: 사용자 PC의 `C:\CLI`로 그대로 복사되는 폴더. **실행에 필요한 .py와 `SKILL.md`만** 둔다.
  런타임이 만드는 `data/`(오류 로그 등, git 제외)만 예외.
- `CLI-dev/<tool>/`: 그 밖의 모든 것 (`README.md`, `CHANGELOG.md`, `requirements.txt`, `tests/`, `tools/` 실기 러너와 보조 스킬).
  `CLI-dev/README.md`(공통 계약), `CLI-dev/run_all_tests.py`. 실기 테스트 때만 `C:\CLI-dev`로 가져가며,
  `CLI-dev`의 테스트·러너는 옆 폴더 `../CLI/<tool>`의 도구를 시험한다.
- 새 파일을 `CLI/<tool>/`에 넣는 것은 도구 실행에 꼭 필요할 때만.

## 사용자 환경 (모든 작업의 전제)
| 항목 | 값 |
|---|---|
| OS | Windows 11, **폐쇄망** (런타임 인터넷 접근 불가) |
| Python | **WinPython**, Python 3.13.13, 시스템 PATH 등록 |
| 실행 명령 | `python`, `python -m pip` 만 사용 (`py`, `python3`, 단독 `pip` 금지). 명령 예시·스킬·배치 파일 포함 |
| 에이전트 | **OpenCode** + 26B~30B 경량 LLM, 도구는 OpenCode 스킬로 제공 |
| 주요 라이브러리 | pandas 3.0.5, scikit-learn 1.9.0, scipy 1.18.1, numpy/joblib(의존성), openpyxl 3.1.5, python-docx 1.2.0, python-pptx 1.0.2, matplotlib 3.11.1, seaborn 0.13.2, pymupdf 1.28.2, olefile 0.47, opencv-python 5.0.0.93, PyYAML 6.0.3, python-dateutil 2.9.0.post0, pywin32 312, pytest 9.1.1, pymupdf4llm 1.28.2, fastapi 0.141.1, fastmcp 3.4.7, mcp 1.28.1, langgraph 1.2.11, openai 2.54.0 |
| 버전 기준 | 이 표. 상세 설명: `docs/ENVIRONMENT.md` |

- 목록 외 패키지도 대부분 설치되어 있고, Python 3.13 / win_amd64 wheel로 추가 설치 가능.
- 환경이 바뀌면 이 표와 `docs/ENVIRONMENT.md` 두 곳을 함께 갱신한다.

## 반드시 지킬 ACI 계약 (상세: CLI-dev/README.md)
1. stdout은 정확히 1줄 JSON. 진행 로그·표·경고는 전부 stderr. `-h/--help` 텍스트도 stderr로 보내고 stdout에는 JSON 1줄.
   - 사람용 실기 테스트 러너(`tools/run_win_tests.py`)는 계약 대상이 아니다. 단, 에이전트용 모드(`--agent` 등)는 계약을 지킨다.
2. exit 0 = ok, 1 = error, 2 = 인자 오류, 130 = 중단.
   - 성공: `"status": "ok"`, `agent_directive`는 `VERIFIED:`(도구가 결과를 확인함) 또는 `UNVERIFIED:`(확인하지 못함, 재시도 없이 보고)로 시작.
   - 실패: `"status": "error"`, `agent_directive`는 `STOP_RETRY:`(입력/데이터 문제. message에 `[추천]`이 있으면 1회만 재시도(`python`으로 시작하는 전체 명령이면 그대로, 아니면 그 값으로 바꿔서), `[후보]`면 사용자에게 선택 요청) 또는 `ERROR:`(내부 오류, 재시도 금지).
   - 공통 키·문구·함수의 구현 규칙은 `CLI-dev/README.md` "공통 런타임 구현 규칙"(R1~R9)을 따른다.
3. 모든 진입점은 공통 런타임의 `run_cli()` + `JsonArgumentParser`를 사용한다.
4. CLI 인자·JSON 키는 바꿔도 된다 (베타, 호환 불필요). 바꾸면 SKILL.md·README.md·CHANGELOG.md와 테스트를 같은 커밋에서 함께 고친다.
5. 문자 규칙(아래)을 지킨다.

## ACI 설계 권장
1. **사후 확인**: `VERIFIED:`는 결과를 실제로 확인했을 때만 붙인다. 쓰기 작업은 저장 후 다시 읽어 기대값과 비교하고,
   확인하지 못했으면 `ok`여도 `UNVERIFIED:`로 시작하고 무엇을 확인하지 못했는지 적는다 (R1).
2. **빈 결과 명시**: 0건도 성공이다. `"count": 0`과 "찾은 항목 없음"을 `agent_directive`에 분명히 적는다. 빈 배열만 주지 않는다.
3. **다음 행동 제시**: 이어서 할 일이 있으면 `next_commands`에 그대로 실행 가능한 전체 명령을 넣는다.
   LLM이 인자를 조합하거나 문서를 다시 읽지 않아도 되게 한다.
4. **출력 양 제한**: JSON에 담는 목록·텍스트는 도구별 상한을 두고, 넘으면 잘라서 `truncated: true`와 전체 결과 파일 경로를 함께 준다.
5. **인자 이름 일관성**: 도구 간 같은 개념은 같은 인자 이름을 쓴다 (`--file`, `--sheet`, `--save`, `--dry-run` 등). 새 인자는 기존 도구의 이름을 먼저 찾아본다.
6. **미리보기와 되돌리기**: 덮어쓰기·삭제·일괄 수정 같은 파괴적 작업은 `--dry-run`(변경 없이 결과 미리보기)을 제공하고,
   가능하면 되돌리기 지점을 남긴다 (예: office-cli `undo`).
7. **판단 대행**: 이름 자동 매칭, 기본값, 자동 상속, 오타 추천은 도구가 처리한다 (CLI-dev/README.md 공통 계약 4번).

## JIT 자가복구 권장
실패를 LLM에게 넘겨 재시도하게 하지 말고, 실패가 난 그 자리에서(Just-In-Time) 도구가 즉시 복구한다.
경량 LLM은 오류를 보고 올바른 대응을 고르는 데 약하다. 복구는 도구가 하고, LLM에는 결과만 전달한다.

1. 복구 대상: 원인이 분명하고, 복구해도 결과의 의미가 바뀌지 않는 실패만.
   | 실패 | 즉시 복구 |
   |---|---|
   | 저장 대상이 Excel/Word에서 열려 잠김 | `_YYYYMMDD_HHMMSS`를 붙인 새 이름으로 저장 (`save_with_lock_fallback`) |
   | 백신/탐색기가 파일·폴더를 잡고 있음 | 짧은 대기 후 재시도 (`retry_fs`, `remove_tree`, 최대 6회) |
   | 입력 인코딩 불일치 | UTF-8/BOM/UTF-16/CP949 자동 판별 (`read_text_file`) |
   | COM 호출 거부(Office가 바쁨) | 짧은 대기 후 재시도 |
   | 캐시·색인 손상 | 버리고 새로 생성 |
   | 장시간 작업 중단 | 진행 상태를 저장해 두고 다음 실행 때 이어서 진행 |
   | 선택 라이브러리 없음 | 대체 경로가 있으면 사용, 없으면 해당 기능만 `STOP_RETRY:`로 설치 명령(`python -m pip install ...`) 안내 |
2. 복구 한도: 재시도는 횟수와 시간 상한을 둔다. 무한 루프 금지. 한도를 넘으면 오류 JSON을 낸다.
3. 복구 금지 (사용자 확인이 필요한 경우): 덮어쓰기, 삭제, 원본 수정, 후보가 둘 이상인 추측, 데이터 값을 바꾸는 보정.
   이때는 `STOP_RETRY:` + message에 `[추천] <값>`(하나, 값 또는 전체 명령) 또는 `[후보] a, b`(여럿)를 넣는다. 덮어쓰기는 `--overwrite`가 있을 때만 한다 (R4).
4. 복구는 숨기지 않는다.
   - stderr에 `[WARN] <무엇이 실패했고 어떻게 복구했는지>`.
   - stdout JSON에 `recovered` 배열 키를 추가한다 (예: `["잠김 -> report_20261028_101500.xlsx로 저장"]`). 복구가 없으면 생략.
     `warnings`(주의만 하면 되는 일)와 구분한다: `recovered`는 실제로 결과가 달라진 일이다.
   - 결과 경로가 바뀌었으면 JSON 경로 키에는 실제 저장 경로를 넣는다.
5. 복구 성공은 `status: ok` / `VERIFIED:`. LLM은 `recovered`를 사용자에게 그대로 전달만 한다.
6. 복구 함수는 각 도구의 공통 런타임에 두고 진입점마다 새로 만들지 않는다.
   새 복구 사례는 회귀 테스트(실패 상황 재현 -> 복구 확인)와 함께 추가한다.

## 오류 로그 저장
- 실패한 실행은 자동으로 `<도구 데이터 폴더>/logs/errors.txt`에 덧붙인다. 사용자 중단(exit 130)은 기록하지 않는다.
  - exit 1은 `run_cli()`, 인자 오류 exit 2는 `JsonArgumentParser.error()`가 기록한다.
  - 데이터 폴더 기본값은 `<도구폴더>/data`, 도구별 환경변수로 바꿀 수 있다 (예: `ML_KIT_DATA_DIR`, `OFFICE_CLI_DATA_DIR`). `data/`는 git에 올리지 않는다.
- 항목 형식 (UTF-8, 항목 사이 빈 줄):
  ```text
  ===== <ISO 시각> | <도구> <버전> | Python <버전> | <OS>
  명령: <실제 인자>
  결과: <error_type> | <message>
  복구 시도: <시도 내용과 횟수>   (JIT 복구가 한도를 넘겨 실패한 경우만)
  traceback: ...                (내부 오류 ERROR: 일 때만)
  ```
- 2MB를 넘으면 `errors.old.txt`로 넘기고 새로 시작한다 (1세대만 보관).
- 기록 중 오류는 모두 삼킨다. 로그 실패가 stdout 1줄 JSON 계약을 깨면 안 된다.
- 경로와 인자만 기록하고 문서 본문·개인정보는 기록하지 않는다.
- JIT 복구에 성공한 실행은 기록하지 않는다 (stderr `[WARN]` + JSON `recovered`로 충분).
- 로그 경로는 stdout JSON에 넣지 않는다. 사람과 개발 에이전트가 사후 분석용으로 읽는다.
- 기준 구현: `archive/doc-classifier/_dc_runtime.py`의 `write_error_log` (보관된 도구, 복사해서 사용).

## 문자 규칙
- 이모지, 원문자(동그라미 숫자 등), 장식 기호(체크 표시, 별, 유니코드 화살표 등)는 저장소 전체에서 금지한다.
  대상: 코드, 주석, stdout/stderr 출력, SKILL.md, README.md, CHANGELOG.md, 커밋 메시지, PR 본문.
- 대신 ASCII를 쓴다: `[OK]`/`[FAIL]`, `->`, `1.` `2.`.
- 가운뎃점(·), 곱셈(×), 도(°) 등 CP949로 표현되는 일반 기호는 허용한다.
- 로그 태그는 `[INFO] [PROGRESS] [WARN] [ERROR] [STATUS]`만 사용한다.
- 이유: CP949 콘솔에서 깨지거나 인코딩 오류가 나고, 경량 LLM의 토큰 해석이 흔들린다.

## 폐쇄망 / Windows 제약
- 런타임 네트워크 접근 금지.
- 새 기능은 표준 환경 라이브러리를 우선 사용. 없는 패키지는 Python 3.13 / win_amd64 wheel이 있는 것만,
  도구 `requirements.txt`(`CLI-dev/<tool>/`), 위 "사용자 환경" 표, `docs/ENVIRONMENT.md`에 기록하고 미설치 시 JSON 오류로 안내.
- 파일 저장·폴더 조작·텍스트 읽기는 JIT 자가복구 표의 공통 함수를 사용한다.
- 도구 폴더(`CLI/<tool>/`)는 단독 배포된다. 공통 런타임 함수는 도구마다 자기 폴더에 두고, 다른 도구 폴더를 import하지 않는다.
- 테스트는 "사용자 환경" 표의 버전 기준으로 통과해야 한다.

## 도구별 규칙
### ml-kit
- 공통 런타임(`run_cli`, `JsonArgumentParser`, 복구 함수)은 `ml_predict.py` PART C에 있다.
- 산출물 레이아웃(`<stem>_ml_data/history.jsonl`, `vX/model_<target>.joblib|json`)과 원장 스키마는 바꿔도 된다 (베타). 바꾸면 이전 버전 폴더는 다시 학습해서 쓴다.
- 학습과 추론이 같아야 하는 로직(데이터 로딩, 수치 변환, `--eval` 수식)은 `ml_predict.py` PART C 한 곳에만 둔다.
- `ml_predict.py`는 다른 ml-kit 파일을 import하지 않는다 (모델 폴더와 함께 단독 배포). `ml_train.py`/`ml_report.py`가 `ml_predict.py`를 import한다.
### office-cli
- 공통 런타임은 `_runtime.py`, Office COM 공통은 `_office.py`.
### 보관 도구 (`archive/`)
- `archive/` 아래 도구는 수정하지 않는다. 지침 반영·검토·테스트 대상이 아니다 (doc-classifier: 다른 개발자 담당).

## 작업 절차
1. 변경 전 `python -m pytest -q`로 기준선 확인.
2. 리팩터링(동작 불변)과 기능 변경은 별도 커밋/PR.
3. 버그 수정에는 회귀 테스트 추가. 동작이 바뀌면 `SKILL.md`, `README.md`, `CHANGELOG.md` 갱신.
4. 남은 할 일은 각 도구 `CHANGELOG.md`(`CLI-dev/<tool>/`) 하단 "남은 할 일"에서 관리 (처리 시 해당 버전 항목으로 이동).
5. `SKILL.md`는 경량 LLM 컨텍스트에 들어가므로 짧게 유지 (상세는 README.md).
6. 스킬은 OpenCode 형식: 도구 폴더명 = frontmatter `name` (`^[a-z0-9]+(-[a-z0-9]+)*$`), frontmatter는 name/description/license/compatibility/metadata만.
   보조 스킬(예: `CLI-dev/office-cli/tools/office-cli-test/`)도 같은 규칙을 따른다.
