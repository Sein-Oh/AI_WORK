# En-Genius 스킬 패키지 (ml-kit, office-cli)

사내 에이전트 En-Genius에 `ml-kit`, `office-cli`를 스킬로 붙이는 파일과 순서입니다.
스크립트는 En-Genius 서버가 아니라 **사용자 PC에서** En-Genius® MCP(트레이)가 실행합니다 (Office 조작, PC 파일 사용 가능).

## 구조
- 도구는 사용자 PC의 `C:\CLI\ml-kit`, `C:\CLI\office-cli`에 있습니다. **OpenCode와 같은 폴더, 같은 데이터**(학습 이력, 되돌리기, 문서 색인)를 씁니다.
- 스킬 zip에는 지시문(`SKILL.md`)과 실행 입구(`run.py`)만 들어 있습니다. `run.py`가 `C:\CLI\<도구>`의 도구를 실행합니다.
- `run.py`는 도구마다 따로 있습니다 (`CLI-dev/engenius/ml-kit/run.py`, `CLI-dev/engenius/office-cli/run.py`).
  - 자기 도구의 스크립트만 실행합니다. 다른 도구의 명령이 오면 그 스킬로 실행하라고 안내합니다.
  - `C:\CLI\<도구>\SKILL.md`의 명령 안내(버전 줄 제외)가 zip을 만들 때와 달라졌으면 실행하지 않고 zip을 다시 가져오라고 안내합니다
    (`error_type: "setup"`). 도구만 고친 업데이트는 zip을 그대로 둬도 됩니다.
  - 트레이가 고른 파이썬에 도구 라이브러리가 없으면 `C:\WPy64-*`의 WinPython을 찾아 그것으로 실행합니다.

| 파일 | 용도 |
|---|---|
| `ml-kit.zip`, `office-cli.zip` | 스킬 가져오기용 (`<도구>/SKILL.md` + `<도구>/run.py`) |
| `에이전트_프롬프트.md` | 전용 에이전트의 역할 프롬프트 (OpenCode 전역 지침을 En-Genius에 맞게 고친 것) |

zip은 `CLI/<도구>/SKILL.md`와 `CLI-dev/engenius/<도구>/run.py`로 만듭니다. 도구 버전이나 SKILL.md를 고치면 저장소 루트에서
`python CLI-dev/engenius/build.py`를 실행해 다시 만듭니다 (`CLI-dev/tests/test_engenius.py`가 최신인지 확인).

## 배포 (사용자 PC 한 대)
1. 배포 묶음을 복사합니다.
   - WinPython 폴더 (예: `C:\WPy64-313130`) - 라이브러리가 설치된 그대로
   - `CLI` 폴더 -> `C:\CLI` (`C:\CLI\ml-kit`, `C:\CLI\office-cli`)
2. En-Genius® MCP 설치·로그인 (게시판 > 자료실). 좌측 아래 "내 PC 연결됨" 확인.
3. En-Genius® MCP 설정
   - "이 PC에서 스킬 스크립트 실행 허용": 켜기 (기본 켜짐)
   - 「스킬 실행 파이썬」: **자동**이면 됩니다. 트레이가 찾은 파이썬에 도구 라이브러리가 없으면 `run.py`가 `C:\WPy64-*`의 WinPython으로 실행합니다.
     WinPython을 `C:\` 바로 아래가 아닌 곳에 두었다면 **직접 지정**으로 그 `python.exe`를 고릅니다.
4. En-Genius 나의 워크스테이션 -> **스킬** 탭 -> [가져오기] -> `ml-kit.zip`, 다시 [가져오기] -> `office-cli.zip`.
5. 에이전트 만들기 (이름 예: `오피스ML`)
   - 에이전트 프롬프트: `에이전트_프롬프트.md`의 내용
   - LLM: 쓸 모델 (Gemma4 26B / BEAM 1.0 / Nemotron 3 Ultra)
   - 붙은 스킬: `/ml-kit`, `/office-cli`
   - 도구 서버: 고르지 않음. 내장 문서(Office)·Outlook 도구와 겹치면 모델이 그쪽을 고를 수 있습니다.
   - 도구 제한을 건다면 목록에 `run_skill_script`를 꼭 넣습니다 (없으면 스크립트를 실행하지 못함).
6. 그 에이전트 대화(`@오피스ML`)에서 요청합니다. 다른 사람에게는 에이전트를 조직에 공유하면 붙은 스킬도 함께 갑니다
   (그 사람 PC에도 1~3이 필요).

## 업데이트
1. 언제나: `C:\CLI` 폴더를 새 것으로 바꿉니다. OpenCode와 En-Genius가 함께 새 도구를 씁니다.
2. 명령 안내가 바뀐 업데이트만(새 명령·옵션, CHANGELOG에 "SKILL" 변경이 있을 때): 스킬 zip을 다시 [가져오기] -> "덮어쓸지" 물으면 덮어씁니다.
   - En-Genius의 지시문은 zip을 가져올 때 서버에 복사되므로, PC의 `C:\CLI`만 바꿔서는 모델이 새 명령을 모릅니다.
   - 잊어도 괜찮습니다: `run.py`가 `C:\CLI\<도구>\SKILL.md`(버전 줄 제외)와 zip을 만들 때의 지시문 지문을 비교해,
     다르면 실행하지 않고 "스킬 zip을 다시 가져오세요"라고 알려 줍니다. 버그 수정처럼 지시문이 그대로인 업데이트는 zip이 그대로여도 실행됩니다.

## 실행 방식 (모델이 하는 일)
- `run_skill_script(name="office-cli", path="run.py", args=["excel.py read --range A1:D20"])`
- 명령은 args 첫 칸에 한 줄로. 결과의 `next_commands`, `[추천]` 전체 명령도 그대로 넣습니다 (`python`, `C:\CLI\...` 경로가 붙어 있어도 됨).
- 표·여러 줄 글은 `@input1.csv`, `@input2.md` 자리를 쓰고 내용을 args 다음 칸에 넣습니다 (한 칸 2000자).
  예: `args=["word.py write --file @input1.md", "## 제목\n- 항목"]`
- 시간 제한: 기본 120초, 최대 600초. 학습·보고서·성능예측표는 `timeout=600`.
- 실행할 때마다 승인 창이 뜹니다 (끌 수 없음).

## 데이터가 남는 곳
OpenCode와 같습니다: `C:\CLI\office-cli\data`, `cache`, `config.json`, `C:\CLI\ml-kit\data`. 모델과 보고서는 데이터 파일 옆 `<데이터>_ml_data\`,
메일 첨부는 `<다운로드>\메일 첨부\`. 환경변수(`OFFICE_CLI_DATA_DIR` 등)를 PC에 지정하면 그것이 먼저입니다.

## 실기에서 먼저 확인할 것
| # | 확인 | 방법 | 통과 기준 |
|---|---|---|---|
| E1 | 연결·파이썬 | `@오피스ML`에게 "office-cli 사용법 보여 줘" (run.py를 인자 없이 실행) | 승인 창 -> 결과 JSON에 `tool_version`, `python`(WinPython 경로), `skill_current: true` |
| E2 | 한글 출력 | "바탕화면의 <엑셀 파일> 앞부분 읽어 줘" | 결과의 한글 값이 깨지지 않음 |
| E3 | Office 조작 | 열린 엑셀에 "A1에 시험 이라고 써 줘" | 실제로 써지고 `VERIFIED` |
| E4 | 여러 줄 글 | 열린 워드에 "제목과 항목 2개로 회의 요약 써 줘" | `@input1.md`로 실행, 제목·목록이 들어감 |
| E5 | 긴 작업 | 샘플 데이터로 "생산효율 예측 모델 만들어 줘" | `timeout=600`으로 실행, 끊기지 않음 |
| E6 | 데이터 공유 | E3 뒤 OpenCode에서 `office.py undo --list` | En-Genius에서 한 작업의 되돌리기 지점이 보임 |
| E7 | `[후보]` 처리 | 오타 열 이름으로 학습 요청 | 다시 실행하지 않고 사용자에게 물음 |
| E8 | 메일 첨부 | "지금 보고 있는 메일 첨부 내려받아 줘" / "이 파일 붙여서 답장 써 줘" | 다운로드 폴더에 저장 / 답장 창에 첨부 |

그 뒤 `docs/opencode/자동테스트/`의 S·A·M·P·R·X 메시지를 같은 에이전트에 보내고, 대화를 [대화 내보내기](.md)로 저장해 `docs/test_logs/`에 올립니다
(채점기 `grade_opencode.py`는 아직 OpenCode JSON만 읽습니다. En-Genius .md 읽기는 샘플을 받은 뒤 추가).
