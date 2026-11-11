# AGENTS.md - AI_WORK 저장소에서 작업하는 모든 AI 에이전트용 규칙

OpenCode, Claude Code(CLAUDE.md가 이 파일을 불러옴) 등 이 저장소에서 작업하는 모든 에이전트가 따르는 단일 원본입니다.

**사용자에게는 늘 한국어로 답한다.** 답변, 진행 보고, 질문, 요약 모두 한국어로 쓴다 (사용자가 다른 언어를 요청할 때만 예외).
도구 결과·로그·서브에이전트 보고가 영어여도 사용자에게 전할 때는 한국어로 옮긴다. 코드, 명령, 파일 경로, 옵션 이름은 원래 글자 그대로 둔다.

## 저장소 성격
- `AIKIT/<tool>/` 의 스크립트는 **Windows 11 폐쇄망에서 26B~30B 경량 LLM 에이전트가 스킬로 호출하는 CLI 도구**(Agent-Computer Interface, ACI)입니다.
- 소비자는 사람이 아닌 경량 LLM입니다. 출력 형식의 작은 변화도 에이전트를 깨뜨립니다.
- 설계 원칙: **LLM 개입 최소화** - 판단과 복구는 도구가 하고, LLM은 명령 1회 실행 + JSON 1줄 해석만 한다.
- **베타 단계**: ml-kit과 office-cli는 베타다. 이전 버전과의 호환은 필요 없다 (사용자가 따로 요청할 때만).
  호환을 위한 코드(옛 인자·키 별칭, 이전 형식 모델·파일을 읽는 대체 경로)를 새로 넣지 않는다.

## 폴더 구조
- `AIKIT/`는 사용자 PC의 `C:\AIKIT`로 통째로 복사된다 (`dev/` 포함, 2026-11-12에 `CLI/`·`CLI-dev/`에서 이름을 바꾸고 dev를 안으로 옮김).
- `AIKIT/`가 곧 OpenCode 설정 폴더다 (setup이 `OPENCODE_CONFIG_DIR=C:\AIKIT`). 설정 파일 없이 폴더 모양으로 읽힌다 (OpenCode 2.2.5 확인):
  `AIKIT/AGENTS.md`(전역 지침), `AIKIT/skills/<스킬 이름>/SKILL.md`(스킬: 폴더 하나 = 스킬 하나).
- `AIKIT/skills/`: 기본 스킬 12개 (ml-kit: `ml-train`, `ml-predict`, `ml-report`, `ml-map`, `ml-dataset`, office-cli: `excel-kit`,
  `word-kit`, `ppt-kit`, `outlook-kit`, `search-kit`, `stats-kit`, `skill-kit`). 그 밖의 폴더는 사용자가 만든 스킬(git 제외).
  스킬이 여러 개를 함께 쓰는 도구는 스킬 폴더가 아니라 도구 폴더에 두고, SKILL.md에는 `python "C:\AIKIT\<tool>\<파일>.py" ...` 전체 명령을 적는다.
  사용자가 만든 스킬의 스크립트는 그 스킬 폴더의 `scripts/`에 둔다 (스킬 폴더만 복사하면 동료 PC에서도 됨).
- `AIKIT/<tool>/`: **실행에 필요한 .py만** 둔다. 런타임이 만드는 `data/`(오류 로그 등, git 제외)만 예외.
- `AIKIT/dev/<tool>/`: 그 밖의 모든 것 (`README.md`, `CHANGELOG.md`, `requirements.txt`, `tests/`, `tools/` 실기 러너와 보조 스킬).
  `AIKIT/dev/README.md`(공통 계약), `AIKIT/dev/run_all_tests.py`. 도구는 dev를 읽지 않는다 (배포에 함께 가도 실행에 영향 없음).
  `AIKIT/dev`의 테스트·러너는 부모 폴더 `AIKIT/<tool>`의 도구를 시험한다.
- 새 파일을 `AIKIT/<tool>/`에 넣는 것은 도구 실행에 꼭 필요할 때만.
- `AIKIT/` 맨 위에는 도구 폴더, `skills/`, `dev/` 외에 `AGENTS.md`(전역 지침), `opencode.json`(셸·권한 기본 설정, 스킬 경로는 적지 않음),
  배포 설치 파일 `setup.py`(사람이 실행, 에이전트 계약 대상 아님)·`setup.bat`, Web UI 실행 파일 `web.bat`만 둔다.
  사용자가 배포 전에 넣는 `opencode.jsonc`(사내 LLM 주소), OpenCode가 만드는 `service.json`, 설치가 만드는 `data/`(보관 폴더)와
  `work/`(Web UI 작업 폴더)는 git 제외. 배포 PC에 함께 두는 `python/`, `opencode.exe`도 저장소에 넣지 않는다.

## 사용자 환경 (모든 작업의 전제)
| 항목 | 값 |
|---|---|
| OS | Windows 11, **폐쇄망** (런타임 인터넷 접근 불가) |
| Python | **WinPython**, Python 3.13.13, `C:\AIKIT\python`에 두고 시스템 PATH 등록 (`python`, `python\Scripts`) |
| 실행 명령 | `python`, `python -m pip` 만 사용 (`py`, `python3`, 단독 `pip` 금지). 명령 예시·스킬·배치 파일 포함 |
| 에이전트 | **OpenCode** CLI 2.2.5 + 26B~30B 경량 LLM(회사 PC: Gemma), 도구는 OpenCode 스킬로 제공. 명령 실행 도구의 시간 제한 120초 |
| 셸 | Windows PowerShell 5.1 (회사 PC에 pwsh 없음). 도구가 내는 명령의 따옴표는 이 셸 기준 (`quote_arg`) |
| Office | Office 16만 대상 (회사 PC 2021, 집 PC 2019 실기 보조). 그보다 낮은 버전은 지원하지 않음 |
| 주요 라이브러리 | pandas 3.0.5, scikit-learn 1.9.0, scipy 1.18.1, numpy/joblib(의존성), openpyxl 3.1.5, python-docx 1.2.0, python-pptx 1.0.2, matplotlib 3.11.1, seaborn 0.13.2, pymupdf 1.28.2, olefile 0.47, opencv-python 5.0.0.93, PyYAML 6.0.3, python-dateutil 2.9.0.post0, pywin32 312, pytest 9.1.1, pymupdf4llm 1.28.2, fastapi 0.141.1, fastmcp 3.4.7, mcp 1.28.1, langgraph 1.2.11, openai 2.54.0 |
| 버전 기준 | 이 표. 상세 설명: `docs/ENVIRONMENT.md` |

- 목록 외 패키지도 대부분 설치되어 있고, Python 3.13 / win_amd64 wheel로 추가 설치 가능.
- 환경이 바뀌면 이 표와 `docs/ENVIRONMENT.md` 두 곳을 함께 갱신한다.

## 반드시 지킬 ACI 계약 (상세: AIKIT/dev/README.md)
1. stdout은 정확히 1줄 JSON. 진행 로그·표·경고는 전부 stderr. `-h/--help` 텍스트도 stderr로 보내고 stdout에는 JSON 1줄.
   - 사람용 실기 테스트 러너(`tools/run_win_tests.py`)는 계약 대상이 아니다. 단, 에이전트용 모드(`--agent` 등)는 계약을 지킨다.
2. exit 0 = ok, 1 = error, 2 = 인자 오류, 130 = 중단.
   - 성공: `"status": "ok"`, `agent_directive`는 `VERIFIED:`(도구가 결과를 확인함) 또는 `UNVERIFIED:`(확인하지 못함, 재시도 없이 보고)로 시작.
     긴 작업이 명령 시간 제한(에이전트 실행 도구 120초) 전에 멈추면 `CONTINUE:` + `"continue": true` + `next_commands`(같은 명령): 한 일은 저장, 같은 명령을 다시 실행하면 이어서 진행 (R10).
   - 실패: `"status": "error"`, `agent_directive`는 `STOP_RETRY:`(입력/데이터 문제. message에 `[추천]`이 있으면 그 뒤의 전체 명령을 그대로 1회만 실행(아래 "사용법 단순성" 2번), `[후보]`면 사용자에게 선택 요청) 또는 `ERROR:`(내부 오류, 재시도 금지).
   - 공통 키·문구·함수의 구현 규칙은 `AIKIT/dev/README.md` "공통 런타임 구현 규칙"(R1~R10)을 따른다.
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
7. **판단 대행**: 이름 자동 매칭, 기본값, 자동 상속은 도구가 처리한다 (AIKIT/dev/README.md 공통 계약 4번).
   오타는 같은 뜻 표(`SAME_MEANING`)로 고칠 수 있을 때만 `[추천]`, 철자만 비슷하면 `[후보]`(사용자에게 물음). 철자 비교만으로 `[추천]`하지 않는다
   (예외: 가장 가까운 것이 읽기만 하는 명령이면 `[추천]`, 잘못 골라도 바뀌는 것이 없음: office-cli `READ_ONLY_CHOICES`)
   (회사 PC 감사 2026-11-06: `--cell` -> `--all`, `--text` -> `--target`처럼 뜻이 다른 옵션을 LLM이 묻지 않고 실행할 뻔했음).

## 사용법 단순성 (반드시 지킬 것)
경량 LLM은 규칙이 많거나 예외가 있거나 고를 것이 있으면 틀린다. 사용법이 단순한 쪽을 기능보다 앞에 둔다.
새 옵션·새 키·새 스킬 문장을 넣을 때와 리뷰할 때 아래를 확인한다.
1. **한 가지 일에 한 가지 사용법**: 같은 결과를 내는 두 번째 사용법(옵션 별칭, 줄임말, 옛 형식)을 두지 않는다.
   잘못 적은 옵션은 받아 주지 않고 고친 전체 명령을 `[추천]`으로 준다. 스킬에는 한 가지 형태만 적는다.
2. **도구가 내는 명령은 늘 그대로 실행할 전체 명령**: `[추천]`, `next_commands`, `confirm_commands`는 `python "<스크립트>" ...` 전체.
   값만 주어 LLM이 명령에 끼워 넣게 하지 않는다 (`[추천] --range A1:D20` 금지). 고를 것이 여럿이면 `[후보] 1) a 2) b`(사용자가 읽을 값에 번호)와 같은 번호의 `confirm_commands`(그 값을 넣은 전체 명령).
3. **구분자 규칙은 저장소 전체에 하나**: 이름 목록과 `이름=값` 여러 개는 `,`, 조건·식·칸 여러 개는 `;`, 조건의 "또는"은 `|`.
   새 옵션은 이 규칙을 따른다. 어긋나게 쓰면 도구가 거절하고 고친 전체 명령을 `[추천]`.
4. **같은 옵션은 어느 명령에서나 같은 뜻, 같은 규칙**: 한 명령에서는 붙이고 다른 명령에서는 붙이지 말라는 식의 예외를 두지 않는다.
   필요 없는 명령에서는 받고 무시하거나(그 사실을 `warnings`) 도구가 스스로 정한다.
5. **LLM에게 고르기·세기·조합을 맡기지 않는다**: 어느 버전·파일을 쓸지, 몇 번 반복했는지, 다음에 무엇을 할지는 도구가 정해
   기본값·`next_commands`로 준다. 반복 상한은 도구가 세고, 끝나면 `next_commands`를 비운다.
6. **사용자에게 알릴 글은 도구가 만든다**: 결과마다 `preview`(그대로 전할 짧은 한국어)를 주고, 스킬은 "preview와 warnings를 알린다"를 기본으로 한다.
   스킬에 "알릴 키" 목록을 길게 적지 않는다 (키 이름을 많이 외우게 하면 빠뜨리거나 지어낸다).
7. **스킬은 짧고 모순이 없게**: 스킬 하나 8KB(8192바이트, LF 기준, 2026-11-10 사용자 결정) 이하. 넘으면 작업별로 나눈다. 규칙에 예외가 있으면 그 규칙 문장 안에 적는다
   (다른 줄의 예외는 놓친다). 전역 지침(`AIKIT/AGENTS.md`)과 스킬이 같은 일을 다르게 말하지 않는다.
8. **옵션을 늘리기 전에**: 기본값, 자동 판단, 기존 옵션으로 안 되는지 먼저 본다. 꼭 필요한 옵션만 스킬에 적고 나머지는 README에만.

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
   이때는 `STOP_RETRY:` + message에 `[추천] <전체 명령>`(하나) 또는 `[후보]`(여럿, 번호를 붙인 값과 같은 번호의 `confirm_commands` 전체 명령. 사용자가 고른 번호의 명령을 실행)를 넣는다. 덮어쓰기는 `--overwrite`가 있을 때만 한다 (R4).
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
  도구 `requirements.txt`(`AIKIT/dev/<tool>/`), 위 "사용자 환경" 표, `docs/ENVIRONMENT.md`에 기록하고 미설치 시 JSON 오류로 안내.
- 파일 저장·폴더 조작·텍스트 읽기는 JIT 자가복구 표의 공통 함수를 사용한다.
- AIKIT/·AIKIT/dev/에 반입할 파일은 .py .md .txt .json .tsv .bat만 쓴다. .yaml/.yml/.jsonl/.toml 등은 회사 반입 시스템이 막으므로 쓰지 않는다(실행 중 PC에서 생기는 파일은 상관없음). 검사: `AIKIT/dev/tests/test_repo_files.py`.
- 도구 폴더(`AIKIT/<tool>/`)는 단독 배포된다. 공통 런타임 함수는 도구마다 자기 폴더에 두고, 다른 도구 폴더를 import하지 않는다.
- 테스트는 "사용자 환경" 표의 버전 기준으로 통과해야 하고, 리눅스(클라우드)와 Windows(회사·집 PC)에서 같은 결과여야 한다. 실기에서 깨졌던 것:
  - 도구가 낸 명령을 나눌 때 `shlex.split`(posix)은 Windows 경로의 `\`를 지운다 -> 테스트 공용 `split_command`(Windows 명령줄 규칙)를 쓴다.
  - 파일 크기·내용 비교는 줄 끝과 무관하게 (`core.autocrlf=true`로 받으면 CRLF가 된다).
  - 라이브러리 기본 템플릿(`pptx.Presentation()`, `docx.Document()`)에 기대지 않는다 (문서 보안 프로그램이 지우거나 바꾼다).
  - 긴 매개변수 테스트 id는 자른다 (Windows 환경변수 32767자 한도). 경로는 260자를 넘을 수 있다고 본다.

## 도구별 규칙
### ml-kit
- 공통 런타임(`run_cli`, `JsonArgumentParser`, 복구 함수, 데이터 로딩, 버전 고르기, `--input` 해석)은 `_ml_runtime.py`에 있다 (진입점 아님, 스킬에 적지 않음).
- 산출물 레이아웃(`<stem>_ml_data/history.jsonl`, `vX/model_<target>.joblib|json`)과 원장 스키마는 바꿔도 된다 (베타). 바꾸면 이전 버전 폴더는 다시 학습해서 쓴다.
- 학습과 추론이 같아야 하는 로직(데이터 로딩, 수치 변환, `--eval` 수식)은 `_ml_runtime.py` 한 곳에만 둔다.
- 진입점(`ml_train.py`, `ml_predict.py`, `ml_report.py`, `ml_export.py`, `ml_map.py`, `ml_dataset.py`)은 공통으로 `_ml_runtime.py`만 import한다. 아무도 `ml_predict.py`를
  import하지 않고, `_ml_runtime.py`는 진입점을 import하지 않는다 (예외: 보고서와 함께 성능예측표를 만드는 `ml_report` -> `ml_export`).
  예측만 하는 PC에는 `ml_predict.py` + `_ml_runtime.py` + 버전 폴더를 복사한다. 검사: `tests/test_contract.py`.
### office-cli
- 공통 런타임은 `_runtime.py`, Office COM 공통은 `_office.py`. 내부 모듈(진입점 아님, 스킬에 적지 않음): `_markdown.py`(마크다운 해석),
  `_xlsx_formula.py`(닫힌 xlsx의 계산 안 된 수식 칸), `_spec_check.py`(excel check).
  `skill.py`(진입점, 스킬 `skill-kit`)는 사용자가 만든 OpenCode 스킬을 등록·점검한다 (Office 불필요, 공통 런타임을 쓰려고 office-cli에 둠).
  `_runtime.py`는 다른 office-cli 파일을 import하지 않는다. `_spec_check.py`는 `excel.py`의 주소·읽기 도우미를 쓴다.
- 세션(`open_target`/`write_session` with 블록) 밖에서 그 안의 COM 객체(시트, 범위, 문서)를 읽지 않는다. 닫힌 파일은 블록이 끝나면 문서와
  백그라운드 Office가 닫혀 OLE 0x800a01a8로 실패하고, 저장은 끝났는데 실패를 보고한다 (회사 PC 2026-11-06 O2). 결과에 쓸 값은 블록 안에서
  파이썬 값으로 꺼내 둔다. 가짜 COM은 이것을 흉내 내지 않으므로 `tests/test_session_scope.py`가 코드 모양으로 검사한다.
### 보관 도구 (`archive/`)
- `archive/` 아래 도구는 수정하지 않는다. 지침 반영·검토·테스트 대상이 아니다 (doc-classifier: 다른 개발자 담당).

## 작업 절차
1. 변경 전 기준선: 바꿀 도구의 테스트(`python -m pytest -q AIKIT/dev/<tool>/tests`)와 공통 테스트(`AIKIT/dev/tests`). 전체(`python -m pytest -q`)는
   ml-kit만 클라우드 30분·회사 PC 1시간 이상 걸리므로 PR 전 1번. 오래 걸리는 파일(`test_layout_matrix`, `test_loading_*`)은 그 영역을 바꿨을 때 꼭 돌린다.
2. 리팩터링(동작 불변)과 기능 변경은 별도 커밋/PR.
3. 버그 수정에는 회귀 테스트 추가. 동작이 바뀌면 `SKILL.md`, `README.md`, `CHANGELOG.md` 갱신.
4. 실기 결과(`docs/test_logs/`)에는 회사 경로·사용자 이름이 들어 있다. 그 내용을 코드·테스트·문서·커밋 메시지에 옮기지 않는다
   (요약은 `docs/HANDOFF.md` 4장에 경로 없이). 배포 파일에 사용자 폴더 전체 경로를 넣지 않는다.
5. 남은 할 일은 각 도구 `CHANGELOG.md`(`AIKIT/dev/<tool>/`) 하단 "남은 할 일"에서 관리 (처리 시 해당 버전 항목으로 이동).
6. `SKILL.md`는 경량 LLM 컨텍스트에 들어가므로 짧게 유지 (8KB, 상세는 README.md, 위 "사용법 단순성" 7번).
7. 스킬은 OpenCode 형식(Claude 스킬과 같은 모양): `AIKIT/skills/<스킬 이름>/SKILL.md`, 폴더명 = frontmatter `name`
   (`^[a-z0-9]+(-[a-z0-9]+)*$`; `AIKIT/skills/ml-train/SKILL.md`는 `ml-train`), frontmatter는 name/description/license/compatibility/metadata만.
   `skills/` 바로 아래에 둔다 (묶음 폴더를 한 단계 더 두는 것은 2.2.5에서 되지만 표준이 아님).
   보조 스킬(예: `AIKIT/dev/office-cli/tools/office-cli-test/`)도 같은 규칙을 따른다. 스킬을 더하거나 나누면 `AIKIT/setup.py`
   (설치·옛 폴더 정리), `office-cli/skill.py`의 `BASE_SKILLS`, `.gitignore`의 기본 스킬 줄, 전역 지침 `AIKIT/AGENTS.md` 3장 표를 함께 고친다.
