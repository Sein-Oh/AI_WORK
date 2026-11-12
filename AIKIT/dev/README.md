# CLI 도구 모음

AI 에이전트(26B~30B 경량 LLM)가 **스킬로 호출하는 CLI 도구**의 개발·테스트 폴더입니다. 도구 1개 = 폴더 1개.

- 실행 파일은 `AIKIT/<tool>/`, 스킬은 `AIKIT/skills/<스킬 이름>/SKILL.md`에 있습니다 (사내 PC `C:\AIKIT`로 그대로 복사, `C:\AIKIT`이 OpenCode 설정 폴더).
- 이 `AIKIT/dev/<tool>/`에는 설명서, 변경 이력, 의존성 목록, 테스트, 실기 러너가 있습니다.
  테스트와 러너는 부모 폴더 `AIKIT/<tool>`(PC에서는 `C:\AIKIT\<tool>`)의 도구를 시험합니다. 이 폴더는 `C:\AIKIT\dev`로 도구와 함께 배포됩니다.

| 도구 | 설명 | 스킬 |
|---|---|---|
| [`ml-kit/`](ml-kit/README.md) | 회귀·분류 ML 학습 / 예측(단독 배포 가능) / Word 보고서 / 표 보간·반대 찾기 | [`AIKIT/skills/`](../skills/) (ml-train, ml-predict, ml-report, ml-map, ml-dataset) |
| [`office-cli/`](office-cli/README.md) | Excel/Word/PowerPoint/Outlook 자동화(열린 문서 읽기·쓰기, 앱 간 전달, 되돌리기), PDF/HWP/Office 파일 읽기와 검색 | [`excel-kit`](../skills/excel-kit/SKILL.md), [`excel-format-kit`](../skills/excel-format-kit/SKILL.md), [`word-kit`](../skills/word-kit/SKILL.md), [`ppt-kit`](../skills/ppt-kit/SKILL.md), [`outlook-kit`](../skills/outlook-kit/SKILL.md), [`search-kit`](../skills/search-kit/SKILL.md), [`stats-kit`](../skills/stats-kit/SKILL.md) |

에이전트 런타임은 **OpenCode** 입니다. 스킬 폴더 이름 = 스킬 이름
(소문자, 숫자, 하이픈만: `^[a-z0-9]+(-[a-z0-9]+)*$`) 이어야 합니다. 사내 PC 배치: 도구는 `C:\AIKIT\<tool>\`에 두고, OpenCode는 설정 폴더 `C:\AIKIT`의 `skills\<스킬 이름>\SKILL.md`를 그 자리에서 읽습니다 ([`docs/opencode/README.md`](../../docs/opencode/README.md)). SKILL.md의 명령 경로는 `C:\AIKIT\<tool>` 기준입니다. 한 도구의 명령이 많으면 작업별 스킬로 나눠 `AIKIT/<tool>/skills/<스킬 이름>/SKILL.md`에 둡니다 (office-cli). 이때 폴더 이름 = 스킬 이름이고, 스킬 하나는 8KB(8192바이트) 이하를 목표로 합니다.

## 실기 자동 점검 (`run_all_tests.py`)

`C:\AIKIT`(도구와 이 폴더 `dev`)가 있는 PC에서 자동 점검을 **명령 하나**로 돌리고 결과를 **파일 하나**로 모읍니다. 사람이 실행하는 도구입니다.
회사 PC 실기 절차 전체(준비, Gemma, 눈 확인, 올릴 파일)는 [`docs/opencode/실기테스트.md`](../docs/opencode/실기테스트.md).

```bat
cd /d C:\AIKIT\dev
python run_all_tests.py --label 회사PC            :: 자동 점검 전부 (질문 없음, 자리 비워도 됨)
python run_all_tests.py --label 회사PC --outlook  :: office 러너에 Outlook(G그룹)도
python run_all_tests.py --label 회사PC --gemma    :: 끝에 Gemma 종합 테스트 (opencode run, 일반 권한 창에서)
python run_all_tests.py --only diag              :: 환경 진단과 도구 오류 로그만 (1분)
```

- 단계(이 순서): `diag`(Python, 라이브러리, 문서 보안(DRM)이 python-docx/pptx 템플릿과 새로 쓴 파일을 바꾸는지) -> `beta`(`run_beta_check.py --agent`, F7~F9 엑셀 재계산 포함: F8 파생변수, F9 글자 입력) -> `office`(office-cli 실기 러너, 실제 Office). 끝에 각 도구의 `data\logs\errors.txt`(실제 사용 중 오류)를 늘 붙입니다.
  `ml`(ml-kit 실기 러너)은 기본 단계가 아니며 `--only`에 넣을 때만. `--only`의 순서와 상관없이 위 순서로 실행합니다.
- pytest(오프라인 테스트)는 이 러너에 없습니다. 클라우드에서만 돌립니다 (2026-11-12 사용자 결정: 회사·집 PC에서 날마다 몇 시간씩은 무리.
  가짜 Office로 보는 시험이라 Windows에서도 결과가 같아야 하고, Windows에서만 보이는 것은 아래 실제 환경 단계가 봅니다).
- `--gemma`(또는 `--only gemma`): 마지막에 `gemma` 단계. `agent_scenarios.py`의 요청을 `opencode run --format json`으로 한 세션에 차례로 보내고
  (첫 요청 뒤로는 `--session <ID>`) `grade_opencode.py`로 채점해 결과 파일에 넣습니다. opencode는 `--opencode`, PATH, `C:\AIKIT\opencode.exe` 순서로 찾고 없으면 ENV-SKIP.
  시험 폴더(기본 `C:\AIKIT\dev\work\gemma`, `--gemma-dir`)가 있으면 `_old_<시각>`으로 옮기고, 요청별 기록은 `C:\AIKIT\dev\work\gemma_runs_<시각>`에 남깁니다.
  첫 요청에서 세션이 안 생기면 나머지를 보내지 않고 FAIL. 요청 하나의 시간 상한 `--gemma-timeout`(기본 900초). 관리자 권한 창이면 UNVERIFIED에 경고.
  시험 자료 준비 요청(`setup: true`, S0·S1)은 모델에 보내지 않고 러너가 그 python 명령을 직접 실행합니다 (모델에 보내면 `C:\AIKIT\dev\ml-kit\tools` 경로를
  도구 위치로 따라 해 다음 요청이 `C:\AIKIT\dev\ml-kit\ml_train.py`를 실행함, 집 PC 2026-11-07). 메시지 팩에는 그대로 있습니다.
  `--gemma-only A1,P1`: 그 요청만 시나리오 순서대로 (준비는 늘 러너가 먼저, 고르지 않은 요청은 채점에서 세지 않음, 모르는 ID면 아무것도 보내지 않고 FAIL).
  `--gemma-model <모델>[,<예비 모델>...]`: opencode `--model` 값(예: `openrouter/google/gemma-4-26b-a4b-it`). 공급자 오류가 사용량 제한(429, rate limit, quota)이면
  제한 기록을 `<ID>.limited<N>.jsonl`로 남기고 다음 모델로 같은 요청을 다시 보내며, 그 뒤 요청도 그 모델로 보냄. 예비가 없으면 UNVERIFIED.
  요청별로 답한 모델은 기록 폴더 `models.json`, 실행 기록의 "요청별 모델", 채점 보고서 머리에 남음.
  합격 기준은 Gemma 4 26B 결과 (`docs/opencode/실기테스트.md` 3장).
  시나리오에 `manual`(사람이 준비할 것)이 있는 요청(O3, Outlook 메일)은 보내지 않고 건너뜀 "수동 전용"으로 남깁니다.
- `diag` 등급: 라이브러리 기본 템플릿 손상은 도구가 그 템플릿을 쓸 때만 FAIL. python-pptx 기본 템플릿은 어느 도구도 쓰지 않아 참고(WARN),
  python-docx 기본 템플릿은 ml-kit 보고서(`ml_report.new_report_document`, 폴더 템플릿으로 복구)가 빈 문서를 실제로 만들 수 있으면 참고, 못 만들면 FAIL.
  참고(WARN)는 결과 파일 맨 위 "참고" 목록에 나오고 실패로 세지 않습니다. 오프라인 테스트의 샘플 문서는 기본 템플릿 없이도 만듭니다
  (`office-cli/tests/sample_docs.py`: pptx는 코드로 만든 최소 OOXML, docx는 폴더 템플릿).
- `--ask`: office 러너의 눈 확인 질문(y/n)에 답하며 진행 (기본은 질문 없이 자동 판정 항목만, 질문 항목은 MANUAL-SKIP).
- Windows + pywin32가 아니면 `ml`, `office`는 ENV-SKIP (클라우드에서도 나머지는 그대로 돎).
- 결과: `C:\AIKIT\dev\result_<라벨>_<시각>.md` 하나. 맨 위 요약 표(단계별 PASS/FAIL/SKIP, 시간, 비고)와 실패·건너뜀(ENV-SKIP)·UNVERIFIED 목록, 그 아래 단계별 상세
  (베타 점검 보고서, 실기 러너의 errors/report 파일), 맨 끝에 도구 오류 로그. 단계가 끝날 때마다 다시 쓰므로 Ctrl+C로 끊어도 남습니다.
  베타 점검이 만든 `beta_check_*.md/.xlsx`는 내용을 결과 파일에 넣고 시험 폴더(`C:\AIKIT\dev\work\beta_<시각>`)로 옮기며, 점검이 일부러 낸 도구 오류는 그 폴더의 `_data`에 기록합니다(실제 오류 로그와 섞이지 않음).
- 종료 코드: 0 = 실패 없음, 1 = FAIL인 단계가 있음, 130 = Ctrl+C. 결과 파일은 GitHub 웹에서 `docs/test_logs/`에 업로드합니다(회사 정보가 있어 프로젝트 마무리 때 이력에서 정리).
- 시험 폴더는 모두 `C:\AIKIT\dev\work\` 아래입니다 (`beta_<시각>`, `office`, `ml 테스트`, `gemma`, `gemma_runs_<시각>`, 눈 확인용 `눈확인`). git에는 올리지 않으며, 반입 때 `C:\AIKIT\dev`를 통째로 지우면 함께 사라집니다.
- 시작 전에 열린 Office 문서를 모두 저장하고 닫고, 점검 중에는 Office 창을 만지지 마세요. `run_all_tests.bat`를 더블클릭해도 같습니다(라벨 `PC`, 끝나면 창이 멈춰 요약을 보여 줌).

## 베타 기능 점검과 Gemma 종합 테스트

| 파일 | 하는 일 |
|---|---|
| `run_beta_check.py` | 정답을 아는 샘플을 새로 만들어 ml-kit 베타 기능(이상치·제외 열, ml_map, 실측 비교값, 모델 반대 찾기, 보고서+성능예측표, 일치도, 엑셀 잠금·숨김·재계산)을 자동 판정. 결과 `beta_check_<라벨>_<시각>.md/.xlsx` (단독 실행 때. `run_all_tests.py`의 `beta` 단계는 이 내용을 결과 파일에 넣음). `--agent`면 stdout에 JSON 1줄 (Gemma가 실행할 때. `-h`·인자 오류·Ctrl+C(exit 130)도 JSON 1줄, `verified` 포함). `--only`에 없는 묶음은 인자 오류 |
| `run_simplify_check.py` | 사용법 단순화(ml-kit 2.9.71, office-cli 1.4.79) 실기 점검 (사람이 실행, Gemma 없이). 틀린 명령을 일부러 실행하고 도구가 준 `[추천]` 전체 명령이나 `[후보]` 번호의 `confirm_commands`를 PowerShell 5.1(리눅스는 bash)에 그대로 넣어 성공하는지, 설정 폴더의 기본 스킬 11개 배치. 닫힌 Office 파일 쓰기 항목은 Windows + Office에서만. 결과 `result_simplify_<라벨>_<시각>.md` |
| `agent_scenarios.py` | Gemma 종합 테스트 시나리오 (요청 문장과 채점 기준) 단일 원본. 데이터만 담은 파이썬 모듈(상수 `SCENARIOS`, 정규식은 raw 문자열): 회사 반입 시스템이 .yaml을 막아 2026-11-05에 `agent_scenarios.yaml`에서 옮김. 없으면 채점 테스트는 "반입 누락"으로 건너뛰고 gemma 단계는 ENV-SKIP. ml-kit 묶음(S0 샘플, A/M/P/R/X)과 office-cli 묶음(S1 샘플 `office-cli/tools/make_office_samples.py`, O1~O5: 규격 대조, 여러 칸 쓰기, 메일 첨부 읽기(수동 전용), 워드 문단 덧붙이기, 양식 채우기). 판정은 스킬 이름이 아니라 실행한 명령으로 하고, 회사 PC 오류 로그에서 나온 지어낸 office-cli 옵션은 공통 규칙 `no_fake_office`로 모든 요청에서 실패 |
| `../setup.py` (`setup.bat`) | 배포 설치 (`C:\AIKIT`에 있음, 사람이 실행). 사용자 환경 변수 `OPENCODE_CONFIG_DIR=C:\AIKIT`(HKCU, 원래 값 보관, 다시 읽어 확인)로 OpenCode가 `C:\AIKIT\AGENTS.md`와 `C:\AIKIT\skills`를 설정 파일·복사 없이 읽게 함. 시스템 PATH(HKLM)에 `C:\AIKIT\python`, `C:\AIKIT\python\Scripts`를 없을 때만 추가(`winreg`, 형식 유지, 원래 값 보관, `WM_SETTINGCHANGE`), 동작 확인(라이브러리 import, 도구 `-h` 1줄 JSON, `opencode.exe`, rg, `opencode debug paths`의 config). `--check`(바꾸지 않고 확인만, 다르면 exit 1), `--no-path`, 개발용 `--test-skill C:\AIKIT\dev\office-cli\tools\office-cli-test\SKILL.md`(`C:\AIKIT\skills`에 둠). `run_all_tests.py`의 diag가 PATH·환경 변수를 읽기만 해서 확인하고 다르면 FAIL |
| `grade_opencode.py` | OpenCode 대화 내보내기 JSON 채점 (필요한 명령, 도구 결과, 답변 값, 금지 행동). 결과 `grade_<라벨>_<시각>.md/.xlsx`. `--judge`는 사내 OpenAI 호환 API로 참고 점수. `--make-pack`은 `agent_scenarios.py`로 메시지 팩을 다시 만듦. 답변 검사는 최종 답변(시나리오 명령 뒤의 글)만 보고, 한 번에 보낸 요청은 명령 순서로 나눠 "참고"로 채점. 시나리오에 `answer_reference: true`인 것(X2, 자유 문장 답)는 답변 정규식 판정을 "통과(참고)"/"실패(참고)"로만 보이고 최종 판정은 `--judge` 점수나 사람이 답변을 읽고 정함 (명령·금지 규칙은 그대로). 내보내기: `cmd /c "opencode export <세션ID> > chat.json"` (PowerShell의 `| Out-File`이나 `>`는 한글을 깨뜨림. 깨진 대화는 [ERROR]로 멈춤). `--runs <폴더>`는 `run_all_tests.py --gemma`가 남긴 `opencode run` 기록(`<ID>.jsonl`, `messages.json`)을 채점. OpenCode 2.x의 셸 도구 이름 `shell`과 파일 경로 키 `path`도 읽음 |

진행 방법과 붙여넣을 메시지: [`docs/opencode/자동테스트/README.md`](../docs/opencode/자동테스트/README.md). 결과 파일은 git에 올리지 않습니다(.gitignore).

## 새 도구 폴더 구성

```text
AIKIT/<tool>/              # 사내 PC C:\AIKIT\<tool>로 그대로 복사
└── *.py                # CLI 진입점과 공통 런타임 (예: excel.py, _runtime.py. 실행에 필요한 파일만, 폴더째 복사해도 단독 실행 가능)

AIKIT/skills/<스킬 이름>/   # 작업별 스킬 (폴더 하나 = 스킬 하나, 예: AIKIT/skills/excel-kit/). OpenCode가 설정 파일 없이 읽음
└── SKILL.md             # frontmatter: name(=폴더명), description(1024자 이하), license, compatibility, metadata
                         # 명령은 도구의 전체 경로 (python "C:\AIKIT\<tool>\<파일>.py" ...)

AIKIT/dev/<tool>/
├── README.md            # 사람용 상세 설명서
├── CHANGELOG.md         # 변경 이력 + 남은 할 일
├── requirements.txt     # 런타임 의존성 (폐쇄망 wheel 반입 기준)
├── tools/               # 실기 테스트 러너 (선택)
└── tests/               # CLI 계약 테스트 (저장소 루트에서 python -m pytest 로 전체 실행). 도구는 ../../AIKIT/<tool>에서 불러옴
```

## 모든 도구가 지키는 ACI 공통 계약

1. **stdout = 정확히 1줄 JSON.** 진행 로그·표·경고는 모두 stderr.
2. 성공: exit 0, `"status": "ok"`, `agent_directive`가 `VERIFIED:`(확인함) 또는 `UNVERIFIED:`(확인 못 함, R1)로 시작.
   긴 작업이 시간 제한 전에 멈추면 `CONTINUE:` + `"continue": true`, 같은 명령을 다시 실행하면 이어서 진행 (R10).
3. 실패: exit 1(인자 오류 2, 중단 130), `"status": "error"`, `agent_directive`가
   - `STOP_RETRY:` - 입력/데이터 문제. message에 `[추천]`이 있으면 그 뒤의 전체 명령을 그대로 1회만 실행 (R8).
   - `ERROR:` - 내부 오류. 재시도 금지.
4. 에이전트가 판단할 일을 줄인다: 이름 자동 매칭, 기본값, 자동 상속을 도구가 처리. 오타는 같은 뜻 표로 고칠 수 있을 때만 `[추천]`, 철자만 비슷하면 `[후보]` (R8).
5. 문자 규칙: 이모지·원문자·장식 기호 금지, 로그 태그는 `[INFO] [PROGRESS] [WARN] [ERROR] [STATUS]` ([`AGENTS.md`](../AGENTS.md) "문자 규칙").
6. 폐쇄망: 실행 중 네트워크 접근 금지. [표준 환경](../docs/ENVIRONMENT.md)에 있는 라이브러리만 사용 ([`docs/WINDOWS_OFFLINE.md`](../docs/WINDOWS_OFFLINE.md)).
7. 공통 결과 키 (해당할 때만 넣음):
   - `next_commands` - 이어서 그대로 실행할 전체 명령 목록 (LLM이 묻지 않고 실행해도 안전한 것만).
   - `confirm_commands` - 사용자에게 보여 주고 "예"를 받은 뒤에만 실행할 전체 명령 목록 (메일 발송, 미리보기 뒤 실제 변경, 읽기·저장 뒤 다른 문서에 넣기 같은 문서를 바꾸는 이어 하기).
     경량 LLM이 `next_commands[0]`을 기계적으로 실행해도 메일이 나가거나 문서가 바뀌지 않게 키를 나눈다.
     덮어쓰기·삭제·메일 발송 항목에는 `"destructive": true`를 붙인다: 사용자가 "묻지 말고 끝까지"라고 했어도 늘 묻는 항목 (전역 지침 3장 6번).
   - 명령 오타 추천은 되돌리기 어려운 명령(메일 발송 등)이면 `[추천]`이 아니라 `[후보]`로 준다 (`[추천]`은 LLM이 묻지 않고 실행).
   - 명령 문자열(`next_commands`·`confirm_commands`·`[추천]` 전체 명령)의 인자는 공통 런타임 `quote_arg`로만 따옴표 처리한다
     (office-cli 1.4.76, ml-kit 2.9.68). 실행 셸은 OpenCode 셸 도구와 같은 순서로 정한다: `pwsh`가 PATH에 있으면 PowerShell 7, 없으면
     Windows PowerShell 5.1 (`OFFICE_CLI_SHELL`/`ML_KIT_SHELL`=`pwsh|powershell|cmd|bash`로 강제). 셸 특수문자가 없는 값은 그대로,
     나머지는 PowerShell 작은따옴표 `'...'`(안의 `'`는 `''`; `$`·백틱·`{}`·`@`·`#`·`;`·`(`가 글자 그대로). 5.1은 파이썬에 넘길 때 값 안의 `"`를
     이스케이프하지 않으므로 `"` -> `\"`, 공백이 있는 값의 끝 백슬래시 두 배, 빈 값 `'""'`를 미리 한다(7.3+는 스스로 하므로 하지 않음).
     cmd(`"..."`, `%`는 `"^%"`)와 bash(`'...'`)는 최선 노력. 검증: `AIKIT/dev/tests/shell_sim.py`(PowerShell 토크나이저, 5.1·7.3 명령줄 조립,
     Windows argv 해석, cmd, bash)와 `test_shell_quote.py`. 손으로 `"..."`를 붙여 명령을 만들지 않는다.
   - `truncated` - 결과를 상한에서 잘랐으면 `true`, 전체 결과 파일 경로를 함께 제공.
   - `recovered` - 도구가 실패를 즉시 복구해 결과가 달라진 일 (예: 잠긴 파일 -> 새 이름 저장). `warnings`와 구분.
8. 실패한 실행(exit 1, 2)은 `<도구 데이터 폴더>/logs/errors.txt`에 자동 기록. 로그 경로는 stdout JSON에 넣지 않음.

설계 권장(사후 확인, 빈 결과 명시, 미리보기와 되돌리기, 인자 이름 일관성)과 JIT 자가복구 규칙은 [`AGENTS.md`](../AGENTS.md)에 있습니다.

## 공통 런타임 구현 규칙

모든 도구가 같은 문제를 같은 모양으로 풀기 위한 규칙입니다. 도구 폴더는 단독 배포되므로 코드는 공유하지 않고,
각 도구의 공통 런타임(ml-kit `_ml_runtime.py`, office-cli `_runtime.py`)에 **같은 이름·같은 동작**으로 구현합니다.
경량 LLM이 도구마다 다른 키나 문구를 만나지 않게 하는 것이 목적입니다. 키 이름과 directive 첫 단어는 바꾸지 마십시오.

### R1. 성공 응답과 사후 확인 (`ok()`)
- 성공 응답은 항상 `"status": "ok"`와 `"verified": true|false`를 가진다.
- `verified: true` -> `agent_directive`는 `VERIFIED:`로 시작. 도구가 결과를 실제로 다시 읽어 확인한 경우만.
- `verified: false` -> `agent_directive`는 `UNVERIFIED:`로 시작하고, 확인하지 못한 내용을 적는다. `unverified` 키에 문구 목록.
  LLM 행동: 재시도하지 말고 결과와 확인하지 못한 내용을 사용자에게 그대로 보고.
- `ok()`는 directive 앞에 `VERIFIED:`를 강제로 붙이지 않는다. 호출하는 쪽이 `verified=`를 넘긴다 (기본값 없음).
- 확인 방법 (최소 기준):

  | 결과 | 확인 |
  |---|---|
  | 파일 저장 | 존재, 크기 > 0, 형식별로 다시 열기 (xlsx: openpyxl read_only, docx: python-docx, csv: 행 수 == 쓴 행 수) |
  | 열린 Office 문서 쓰기 | 쓴 범위 값·단락·도형 수를 COM으로 다시 읽어 기대값과 비교 |
  | DB/원장 기록 | 커밋 후 다시 조회한 건수 |
  | 모델 저장 | 저장된 파일을 다시 load |
- 확인을 마친 경우에만 "다시 읽어서 확인하지 마세요" 같은 문구를 쓸 수 있다.
- 새 첫 단어 `UNVERIFIED:`와 `[후보]`는 각 SKILL.md의 "agent_directive 첫 단어별 행동" 표에 같은 PR에서 추가한다.

### R2. 빈 결과와 잘린 결과
- 목록·검색·읽기 결과는 `count`(잘리기 전 전체 개수)를 가진다. 기존 개수 키(`total_matches` 등)는 그대로 두고 `count`를 추가한다.
- 0건: `status: ok`, `count: 0`, directive는 `VERIFIED: 0건. <조건>에 맞는 항목이 없습니다. 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요.`
  0건일 때 "다음 페이지", "다음 단계 실행" 같은 진행 안내를 넣지 않는다.
- 잘림: `truncated: true`, `shown`(JSON에 담은 개수), `count`, 그리고 전체 결과를 볼 방법 하나:
  `full_result_file`(도구가 `<데이터 폴더>/results/`에 자동 저장한 경로) 또는 `next_commands`(다음 페이지 명령). 안 잘렸으면 `truncated: false`.

### R3. JIT 복구 보고 (`note_recovered`)
- 공통 런타임에 `note_recovered(msg: str)`를 둔다. stderr에 `[WARN] <msg>`를 쓰고 실행 중 목록에 쌓는다.
- `ok()`와 오류 응답 모두 목록이 비어 있지 않으면 `recovered: [msg, ...]`를 붙인다. 비어 있으면 키를 생략.
- 문구: `<무엇이 실패> -> <어떻게 복구>` 한 줄, 200자 이하, 경로는 파일 이름만 (전체 경로는 결과의 경로 키에).
  예: `"report.xlsx 잠김 -> report_20261028_101500.xlsx로 저장"`
- 복구한 공통 함수(`save_with_lock_fallback`, 재시도 성공한 COM 호출 등)가 직접 `note_recovered`를 부른다. `warnings`는 주의 사항 전용으로 남긴다.

### R4. 파일 함수 (`retry_fs`, `remove_tree`, `save_with_lock_fallback`)
- `retry_fs(fn, attempts=6, delay=0.5)`: 대기 `delay * 시도번호`(합계 약 10.5초). 일시적 오류만 재시도:
  `PermissionError`, `winerror` 5/32/33(접근 거부, 공유 위반, 잠금 위반). 그 밖의 오류는 즉시 올린다.
  한도를 넘기면 `record_recovery_attempt("<작업> <파일명> <횟수>회 재시도 실패")`로 오류 로그(R6)의 `복구 시도:` 줄에 남기고 예외를 올린다.
- `remove_tree(path)`: 없는 경로는 성공. `shutil.rmtree(onexc=...)`에서 읽기 전용 속성을 풀고 다시 시도하며 전체를 `retry_fs`로 감싼다.
  **사용자가 준 경로를 지우지 않는다.** 도구가 만든 폴더(표식 파일 `.created_by_<tool>` 또는 도구 데이터 폴더 아래)만 지운다.
- `save_with_lock_fallback(path, writer, what)`: `retry_fs`로 먼저 재시도하고, 그래도 잠겨 있으면 `<이름>_<YYYYMMDD_HHMMSS><확장자>`로 저장하고 `note_recovered`.
  이름이 이미 `_YYYYMMDD_HHMMSS`로 끝나면 시각을 바꿔 붙인다 (두 번 붙이지 않는다). 실제 저장 경로를 반환한다.
- 덮어쓰기: 결과 파일이 이미 있으면 기본은 시각을 붙인 새 이름으로 저장하고 `note_recovered`. 덮어쓰기는 `--overwrite`가 있을 때만.
  (도구가 직접 관리하는 캐시·색인·스냅샷은 예외)

### R5. 오류 분류와 라이브러리 확인 (`require_lib`)
- `STOP_RETRY:` 대상: `InputError`(와 하위 클래스), `FileNotFoundError`, `NotADirectoryError`, `IsADirectoryError`, 복구 한도를 넘긴 `PermissionError`.
  일반 `ValueError`/`KeyError`는 `ERROR:`(내부 오류). 입력 때문에 나는 예외는 경계에서 `InputError`로 바꿔 올린다.
- `require_lib(module, pip_name, feature)`: import 실패 시 `InputError` 하위 클래스 `MissingLibraryError`를 올린다. 응답:
  - `error_type: "missing_library"`, `missing_package: <pip_name>`, `install_command: "python -m pip install <pip_name>"`
  - directive: `STOP_RETRY: <feature>에 필요한 <pip_name>이(가) 설치되어 있지 않습니다. install_command를 사용자에게 전달하세요. 직접 설치하지 마세요.`
- 공통 런타임 모듈은 최상단에서 표준 라이브러리만 import한다. 핵심 라이브러리(ml-kit의 pandas, sklearn 등)도 실패하면 같은 JSON을 내야 한다.

### R6. 오류 로그 (`write_error_log`)
- 기준 구현: `archive/doc-classifier/_dc_runtime.py`의 `write_error_log` (보관된 도구에서 복사). 형식과 규칙은 [`AGENTS.md`](../AGENTS.md) "오류 로그 저장".
- 호출 위치: `run_cli()`의 exit 1 경로와 `JsonArgumentParser.error()`(exit 2). 회전 실패와 기록 실패를 따로 감싸 회전이 실패해도 항목은 남긴다.
- 데이터 폴더 환경변수: `ML_KIT_DATA_DIR`, `OFFICE_CLI_DATA_DIR`. 머리줄의 도구 이름은 폴더 이름(`ml-kit` 등).
- 공통 런타임에 `VERSION` 상수를 두고 SKILL.md `metadata.version`과 같은지 테스트로 확인한다.

### R7. 도움말 (`-h/--help`)
- `JsonArgumentParser`는 도움말 텍스트를 stderr로 쓰고, stdout에는
  `{"status": "ok", "command": "help", "verified": true, "skill_file": "<SKILL.md 경로>", "agent_directive": "VERIFIED: 도움말을 stderr에 출력했습니다. SKILL.md의 명령 템플릿을 그대로 사용하세요. 도구 코드(.py)는 읽지 마세요. ..."}` 1줄, exit 0.
  에이전트가 도움말 다음에 코드를 읽지 않고 SKILL.md로 돌아가게 하려는 것.
- 서브명령 파서도 같은 클래스로 만든다.

### R8. 미리보기와 후보 (`--dry-run`, `[추천]`, `[후보]`)
- `--dry-run`: 검증까지 모두 하고 쓰기는 하지 않는다. 응답 `dry_run: true`, `planned: ["삭제: a.csv", "덮어쓰기: b.xlsx", "변경: Sheet1!A1:C10 (30셀)"]`,
  `confirm_commands[0]` = `--dry-run`을 뺀 같은 명령(`next_commands`에 넣지 않음), directive `VERIFIED: 미리보기만 했습니다(변경 없음). planned를 사용자에게 보여주고 사용자가 '예'라고 한 뒤에만 confirm_commands[0]을 실행하세요.`
- 실제 실행은 수행한 파괴적 작업을 같은 형식의 `applied` 목록으로 보고한다.
- `[추천] <전체 명령>`: 바로 다시 실행할 명령 **하나**. 늘 `python "<스크립트>" ...` 전체 명령이다 (값만 주지 않음, AGENTS.md "사용법 단순성" 2번).
  LLM은 `[추천]` 뒤의 명령을 그대로 1회 실행한다. 만드는 함수: office-cli `retry_command`·`retry_replacing`(`_runtime.py`), ml-kit `option_retry`·`command_text`·`command_with_value`(`_ml_runtime.py`).
- `[후보] 1) a 2) b`(최대 10개): 후보가 둘 이상이라 도구가 고를 수 없을 때. LLM은 재시도하지 않고 사용자에게 고르게 한다. 후보 값을 넣은 전체 명령은
  같은 번호의 `confirm_commands`(사용자가 고른 뒤에만 실행하므로 `next_commands`가 아님, 덮어쓰기·삭제·발송이면 `destructive`). 만드는 함수: office-cli `ambiguous(replace=)`,
  ml-kit `numbered`·`ask`(글만 만드는 도우미도 `ask`로 만든 `[후보]`가 든 오류가 같은 번호의 명령을 가짐). `[후보]`는 늘 번호를 붙인다.
  명령을 만들 수 없는 후보(파일 머리글의 이름, 수식 안의 이름 등)는 번호만 있고 `confirm_commands`가 없으며, LLM은 고른 값으로 1번 실행한다 (줄여 갈 대상).
- 인자 오류(없는 옵션, 필수 옵션 누락)에서 지어낸 옵션이 같은 뜻의 옵션으로 바꿀 수 있는 것(도구별 표 `SAME_MEANING`)이면 고친 전체 명령을 `[추천]`.
  철자만 비슷한 옵션은 뜻이 다를 수 있어 `[후보]`로만 주고, 되돌리기 어렵거나 뜻이 크게 다른 옵션(`--all`, `--overwrite`, `--to` 등)은 추천하지 않는다.
- 여러 항목을 받는 인자(`--cells`, `--slots`, `--where`)를 정해진 구분자가 아닌 쉼표·공백 등으로 나눈 것이 분명하면 쓰지 않고 고친 값·명령을 `[추천]`
  (값 안의 쉼표는 그대로: 숫자 `1,000`, 따옴표로 묶은 글, `R1=10k,R2=4.7k`처럼 앞에 `=`가 있는 글).

### R9. 테스트
각 도구는 위 규칙마다 회귀 테스트를 둔다: `verified`/`UNVERIFIED:`, 0건 응답, `truncated` 키, `recovered`(잠금 모의),
`retry_fs`의 재시도 대상 구분, `remove_tree`의 사용자 경로 거부, `missing_library` 응답, errors.txt(exit 1/2 기록, 130 미기록, 회전),
`-h` stdout 1줄, `--dry-run`이 파일을 바꾸지 않음.

### R10. 긴 작업은 나눠서 진행 (`CONTINUE:`)
- 에이전트의 명령 실행 도구는 명령을 시간 제한(OpenCode 120초)에 끊는다. 끊기면 결과 JSON이 없어 LLM이 같은 명령을 반복한다 (2026-11-04 회사 PC Gemma 시험 R1, R2).
- 오래 걸리는 단계(재학습 등)는 끝낸 결과를 디스크에 저장하고, 다음 단계 전에 시계를 본다 (`check_time`). 제한(`ML_KIT_TIME_LIMIT`, 기본 100초)에서 여유 30초를 뺀 시각을 넘을 것 같으면
  `Unfinished`로 멈추고 `run_cli`가 `"status": "ok"`, `"verified": false`, `"continue": true`, `"progress"`, `"next_commands": [{"purpose", "command": 같은 명령, "needs_input": false}]`(다른 결과와 같은 모양),
  `agent_directive` = `CONTINUE: ...`(exit 0)을 낸다. 오류 로그에는 남기지 않는다.
- 같은 명령을 다시 실행하면 저장된 결과를 다시 쓰고 이어서 진행한다. 실행마다 새 단계를 최소 1개는 하므로 반복하면 반드시 끝난다.
- LLM 행동: 사용자에게 묻지 않고 `next_commands`를 그대로 다시 실행, `CONTINUE:`가 아닌 결과가 나올 때까지. 횟수는 LLM이 세지 않는다.
- 도구가 센다 (ml-kit 2.9.71, office-cli 1.4.79): `<데이터 폴더>/continue_state.json`에 명령마다 횟수와 마지막 `progress`. 300번을 넘거나 바로 앞과 같은
  `progress`에서 멈췄으면(ml-kit은 이번 실행에서 새로 저장한 것도 없을 때) `CONTINUE:` 대신 `STOP_RETRY:`(`error_type: unfinished_limit`, exit 1).
  명령이 다른 결과로 끝나면 그 기록을 지운다. 회사 PC에서 타겟 2개에 약 55번이 걸렸으므로(2.9.69) 상한은 300. `progress`는 실행마다 달라지게(끝낸 단계 수, `약 N단계 남음`) 쓴다.
- 사람이 돌리는 점검 러너(`run_beta_check.py`, 실기 러너)와 테스트는 `ML_KIT_TIME_LIMIT=0`(제한 없음).
- 적용: ml-kit(`ML_KIT_TIME_LIMIT`, `ml_train.py` 여러 타겟 학습), office-cli(`OFFICE_CLI_TIME_LIMIT`, `reader.py index`와 `search --file` 폴더 등록). 새 긴 작업도 같은 이름·같은 키로 구현한다.
- 나누는 단위는 한 실행에 반드시 끝낼 수 있을 만큼 작게: ml-kit 학습은 학습 전 준비(2.9.67) -> 타겟 -> 후보 모델 -> 교차검증 fold(2.9.66) -> 1위 전체 적합 -> 순열 중요도 열 단위로 저장한다.
  실행마다 다시 하는 고정 비용(파일 읽기, 검증 등)이 크면 그것도 저장해 이어 실행이 다시 쓴다 (고정 비용 + 첫 단계가 제한을 넘으면 진행이 없음).
- 단계 하나가 제한보다 긴 경우 (ml-kit 2.9.67): 실행의 첫 단계(시계 없이 함)는 시작 표시(`<단계>.start`, 시도 횟수와 제한)를 남기고 저장하면 지운다.
  표시만 남은 단계를 다시 만나면(이전 실행이 그 단계 도중 끊김) 2번째까지 다시 시도하고, 2번 끊긴 단계는 시작하지 않고 `STOP_RETRY:` +
  `[추천] <같은 명령> --time-limit 570`(명령 실행 도구 timeout 600000)을 낸다. 늘린 제한으로도 끊겼으면 `[추천]` 없이 멈춘다.
- 안전망 (ml-kit 2.9.67): `run_cli`가 감시 스레드를 둔다. 명령 시작부터 제한 + 15초에 아직 안 끝났으면 1줄 JSON을 직접 내고 끝낸다
  (이 실행이 저장한 단계가 있거나 표시한 단계의 첫 끊김이면 `CONTINUE:`, 아니면 `STOP_RETRY:`, 오류 로그 기록). stdout은 락과 표시로 한 번만 쓴다.

