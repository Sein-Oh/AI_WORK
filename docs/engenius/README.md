# En-Genius 스킬 패키지 (ml-kit, office-cli)

사내 에이전트 En-Genius에 `ml-kit`, `office-cli`를 스킬로 붙이는 파일과 순서입니다.
스크립트는 En-Genius 서버가 아니라 **사용자 PC에서** En-Genius® MCP가 실행합니다 (Office 조작, PC 파일 사용 가능).

## 파일
| 파일 | 용도 |
|---|---|
| `ml-kit.zip` | 스킬 가져오기용. `ml-kit/SKILL.md`(En-Genius용 지시문) + `run.py`(실행 입구) + 도구 .py |
| `office-cli.zip` | 같은 구성의 office-cli |
| `에이전트_프롬프트.md` | 전용 에이전트의 역할 프롬프트 (OpenCode 전역 지침을 En-Genius에 맞게 고친 것) |

`run.py`는 도구마다 따로 있습니다 (`CLI-dev/engenius/ml-kit/run.py`, `CLI-dev/engenius/office-cli/run.py`). 각자 자기 도구의 스크립트만 실행하고,
다른 도구의 명령이 오면 그 스킬로 실행하라고 안내합니다.
zip은 `CLI/<도구>`의 도구 파일과 SKILL.md로 만듭니다. 도구나 SKILL.md를 고치면 저장소 루트에서
`python CLI-dev/engenius/build.py`를 실행해 다시 만듭니다 (`CLI-dev/tests/test_engenius.py`가 최신인지 확인).

## 사용자 PC 준비 (한 번)
1. En-Genius® MCP 설치·로그인 (게시판 > 자료실). 좌측 아래 "내 PC 연결됨" 확인.
2. En-Genius® MCP 설정
   - "이 PC에서 스킬 스크립트 실행 허용": 켜기 (기본 켜짐)
   - 「스킬 실행 파이썬」: **자동**이 WinPython(`python`)을 찾는지 설정 창의 경로·버전으로 확인. 다르면 **직접 지정**으로 WinPython의 `python.exe`
3. 라이브러리는 그 파이썬에 이미 설치된 것을 씁니다 (`docs/ENVIRONMENT.md` 표). 트레이 내장 파이썬에는 pandas 등이 없으므로 **트레이 내장으로 두지 않습니다**.

## En-Genius에서 설정
1. 나의 워크스테이션 -> **스킬** 탭 -> [가져오기] -> `ml-kit.zip`, 다시 [가져오기] -> `office-cli.zip`.
   새 버전은 같은 방법으로 가져오고 "덮어쓸지" 물으면 덮어씁니다.
2. 에이전트 만들기 (이름 예: `오피스ML`)
   - 에이전트 프롬프트: `에이전트_프롬프트.md`의 내용
   - LLM: 시험할 모델 (Gemma4 26B / BEAM 1.0 / Nemotron 3 Ultra)
   - 붙은 스킬: `/ml-kit`, `/office-cli`
   - 도구 서버: 고르지 않음. 내장 문서(Office)·Outlook 도구와 겹치면 모델이 그쪽을 고를 수 있습니다.
   - 도구 제한을 건다면 목록에 `run_skill_script`를 꼭 넣습니다 (없으면 스크립트를 실행하지 못함).
3. 그 에이전트 대화(`@오피스ML`)에서 요청합니다.

## 실행 방식 (모델이 하는 일)
- `run_skill_script(name="office-cli", path="run.py", args=["excel.py read --range A1:D20"])`
- 명령은 args 첫 칸에 한 줄로. 결과의 `next_commands`, `[추천]` 전체 명령도 그대로 넣습니다 (`python`, `C:\CLI\...` 경로가 붙어 있어도 됨).
- 표·여러 줄 글은 `@input1.csv`, `@input2.md` 자리를 쓰고 내용을 args 다음 칸에 넣습니다 (한 칸 2000자).
  예: `args=["word.py write --file @input1.md", "## 제목\n- 항목"]`
- 시간 제한: 기본 120초, 최대 600초. 학습·보고서·성능예측표는 `timeout=600`.
- 실행할 때마다 승인 창이 뜹니다 (끌 수 없음).

## 데이터가 남는 곳
트레이는 스킬 파일을 실행마다 `%TEMP%\engenius-skill\<id>\`에 풀고 끝나면 지웁니다. 그래서 도구가 오래 쓰는 파일은
`%USERPROFILE%\Documents\En-Genius\CLI-data\<도구>\`에 둡니다 (office-cli 1.4.33, ml-kit 2.9.37).
| 도구 | 남는 것 |
|---|---|
| office-cli | `config.json`(차트 전달 방식), `data\`(오류 로그, `_work` 결과 파일, 메일 첨부), `cache\`(문서 색인, 되돌리기 스냅샷) |
| ml-kit | `data\logs\errors.txt`(오류 로그). 모델과 보고서는 지금처럼 데이터 파일 옆 `<데이터>_ml_data\` |

OpenCode에서 쓰던 `C:\CLI\...\data`, `cache`와는 따로입니다. 옮기려면 그 폴더를 위 위치로 복사합니다.
환경변수(`OFFICE_CLI_DATA_DIR`, `OFFICE_CLI_CACHE_DIR`, `OFFICE_CLI_CONFIG`, `ML_KIT_DATA_DIR`)를 PC에 지정하면 그것이 먼저입니다.

## 실기에서 먼저 확인할 것
| # | 확인 | 방법 | 통과 기준 |
|---|---|---|---|
| E1 | 연결·파이썬 | `@오피스ML`에게 "office-cli 사용법 보여 줘" (run.py를 인자 없이 실행) | 승인 창 -> 결과 JSON에 `scripts` 목록 |
| E2 | 한글 출력 | "바탕화면의 <엑셀 파일> 앞부분 읽어 줘" | 결과의 한글 값이 깨지지 않음 |
| E3 | Office 조작 | 열린 엑셀에 "A1에 시험 이라고 써 줘" | 실제로 써지고 `VERIFIED` |
| E4 | 여러 줄 글 | 열린 워드에 "제목과 항목 2개로 회의 요약 써 줘" | `@input1.md`로 실행, 제목·목록이 들어감 |
| E5 | 긴 작업 | 샘플 데이터로 "생산효율 예측 모델 만들어 줘" | `timeout=600`으로 실행, 끊기지 않음 |
| E6 | 데이터 유지 | E3 뒤 "방금 것 되돌려 줘" (`office.py undo`) | 되돌려짐 (스냅샷이 남아 있음) |
| E7 | `[후보]` 처리 | 오타 열 이름으로 학습 요청 | 다시 실행하지 않고 사용자에게 물음 |

그 뒤 `docs/opencode/자동테스트/`의 S·A·M·P·R·X 메시지를 같은 에이전트에 보내고, 대화를 [대화 내보내기](.md)로 저장해 `docs/test_logs/`에 올립니다
(채점기 `grade_opencode.py`는 아직 OpenCode JSON만 읽습니다. En-Genius .md 읽기는 샘플을 받은 뒤 추가).
