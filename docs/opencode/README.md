# 사내 PC OpenCode 배치 안내

사내 PC(Windows 11, 폐쇄망, OpenCode + 26B~30B 모델)에 도구 스킬과 전역 지침을 놓는 방법입니다.
저장소는 반입 후 `C:\CLI`에 `ml-kit`, `office-cli`가 있는 상태를 기준으로 합니다.

## 1. 도구와 스킬 배치

도구는 **`C:\CLI\ml-kit`, `C:\CLI\office-cli`** 에 둡니다. 모든 SKILL.md와 전역 AGENTS.md의 명령 경로가 이 위치를 기준으로 적혀 있습니다.
반입할 때는 `C:\CLI`만 바꾸면 됩니다. `C:\CLI`에는 실행 파일과 스킬(SKILL.md)만 있습니다.
실기 테스트가 필요할 때만 저장소의 `CLI-dev` 폴더를 `C:\CLI-dev`로 함께 가져옵니다 (`C:\CLI`와 나란히 두어야 `C:\CLI`의 도구를 찾습니다).

OpenCode 스킬 폴더(`%USERPROFILE%\.config\opencode\skills`)에는 SKILL.md가 든 폴더만 복사합니다. 도구 파일은 넣지 않습니다.

| 스킬 폴더 | 넣을 파일 |
|---|---|
| `skills\ml-kit\` | `C:\CLI\ml-kit\SKILL.md` |
| `skills\office-excel\` | `C:\CLI\office-cli\skills\office-excel\SKILL.md` (엑셀) |
| `skills\office-doc\` | `C:\CLI\office-cli\skills\office-doc\SKILL.md` (워드, PPT, 앱 사이 옮기기) |
| `skills\office-mail\` | `C:\CLI\office-cli\skills\office-mail\SKILL.md` (Outlook) |
| `skills\office-spec\` | `C:\CLI\office-cli\skills\office-spec\SKILL.md` (파일 읽기·근거 검색, 규격 대조) |
| `skills\office-cli-test\` | `C:\CLI-dev\office-cli\tools\office-cli-test\SKILL.md` (실기 테스트를 맡길 때만) |

SKILL.md는 반입 때마다 바뀔 수 있으므로, `C:\CLI`를 갱신한 뒤 `C:\CLI-dev`에서 `python apply_skills.py`로 복사합니다
(전역 지침까지 바뀐 것만 복사, `--test`면 office-cli-test도, `--check`는 확인만). 그 뒤 OpenCode를 다시 시작합니다.
폴더 이름은 표와 같아야 합니다 (SKILL.md 맨 위 `name`과 같은 이름).
office-cli 1.4.61까지 쓰던 통합 스킬 폴더 `skills\office-cli\`는 `apply_skills.py`가 `work\skill_backup_<시각>`에 보관한 뒤 지웁니다
(남아 있으면 OpenCode가 옛 스킬도 고릅니다).

**ripgrep(rg.exe)이 필요합니다.** OpenCode는 스킬을 불러올 때 ripgrep을 쓰고, 없으면 인터넷에서 받으려 합니다.
폐쇄망에서는 받지 못해 스킬 불러오기가 `ripgrep execution failed`로 실패합니다(2026-10-29 실기: 대화 6개 모두).
ripgrep Windows용(`ripgrep-<버전>-x86_64-pc-windows-msvc.zip`)을 반입해 `rg.exe`를 PATH에 있는 폴더(예: `C:\CLI\python`)에 두고,
PowerShell에서 `rg --version`이 나오는지 확인한 뒤 OpenCode를 다시 시작합니다. 그 전까지는 전역 AGENTS.md 3장 3번대로 SKILL.md를 직접 읽게 됩니다.

- 도구 오류 로그는 `C:\CLI\<도구>\data\logs\errors.txt`에 쌓이고, `python C:\CLI-dev\run_all_tests.py`가 결과 파일 끝에 붙여 줍니다 (오류 로그만 빨리 모으려면 `--only diag`, 1분). 실기 절차는 [`실기테스트.md`](실기테스트.md).
- 이전처럼 스킬 폴더에 도구 전체를 복사해 두었다면 그 폴더는 SKILL.md만 남기고 `C:\CLI`만 쓰세요. 두 곳에 도구가 있으면 어느 쪽이 쓰였는지 헷갈립니다.

## 2. 전역 지침 (AGENTS.md)

1. `CLI-dev/opencode/AGENTS.md`(PC에서는 `C:\CLI-dev\opencode\AGENTS.md`)를 `%USERPROFILE%\.config\opencode\AGENTS.md`로 복사합니다.
   `python apply_skills.py`가 스킬과 함께 복사합니다. 원본은 CLI-dev에 둡니다 (반입 때 `C:\CLI-dev`가 함께 가므로).
2. 고쳐 쓸 곳은 없습니다. 도구 경로는 `C:\CLI` 기준으로 적혀 있습니다.

이미 넣어 두신 내용은 아래처럼 반영했습니다.

| 기존 항목 | 반영 위치 |
|---|---|
| 한국어 우선 (답변, 설명, 주석) | 1장 1번. 명령어와 경로는 원문 그대로 쓰도록 구분 |
| 이모지 금지 | 1장 2번. 대신 쓸 글자 표시(`[OK]`, `->`)를 함께 줌 |
| 파일 삭제 전 명시적 승인 | 5장. 덮어쓰기, 메일 발송, 시트·슬라이드 삭제도 같은 규칙으로 묶음 |
| 폐쇄망, pip·인터넷 불가 | 2장. 도구가 설치 명령을 안내해도 직접 실행하지 않음(4장) |
| ml-kit, office-cli 적극 활용 | 3장. 요청별 스킬 표, 직접 파이썬 코드를 짜지 않는 규칙 |
| PowerShell, 한글 깨짐 방지 | 2장. 두 환경변수에 `[Console]::OutputEncoding`을 더함 (아래 참고) |

한글 깨짐: `PYTHONIOENCODING`, `PYTHONUTF8`은 파이썬이 UTF-8로 **내보내게** 하는 설정이고,
`[Console]::OutputEncoding=[Text.Encoding]::UTF8`은 PowerShell이 그 출력을 UTF-8로 **읽게** 하는 설정입니다. 둘 다 있어야 확실합니다.
두 환경변수는 cmd에서 한 번 `setx PYTHONIOENCODING utf-8`, `setx PYTHONUTF8 1`로 영구 설정해 두면 더 안전합니다 (새 창부터 적용).

## 3. 26B~30B 모델에 맞춘 작성 원칙

AGENTS.md와 각 SKILL.md는 아래 원칙으로 썼습니다. 고칠 때도 지켜 주세요.

1. 규칙은 번호 목록, 선택은 표로 씁니다. 긴 문단은 작은 모델이 놓칩니다.
2. 명령은 그대로 복사할 수 있는 템플릿으로 주고 `<...>`만 바꾸게 합니다. 옵션을 짐작하게 두지 않습니다.
3. 결과 해석은 `agent_directive` 첫 단어 하나로 정합니다 (`VERIFIED:`, `UNVERIFIED:`, `CONTINUE:`, `STOP_RETRY:` + `[추천]`/`[후보]`, `ERROR:`).
   전역 지침과 각 SKILL.md의 표가 같은 말을 하도록 맞춰 두었습니다.
4. "정확하게", "친절하게" 같은 일반론은 쓰지 않습니다. 할 행동과 하지 말 행동만 씁니다.
5. PowerShell에서 깨지기 쉬운 것(큰따옴표가 든 JSON 인자, `$`, 여러 줄 글)은 파일로 넘기게 합니다.
   예전 office-cli SKILL.md의 `--values "[[\"이름\",\"값\"]]"` 예시는 PowerShell에서 따옴표가 사라져 실패하므로 `--file "표.csv"`로 바꿨습니다.
6. SKILL.md 상세 설명은 README.md로 보냅니다. SKILL.md는 스킬을 쓸 때마다 컨텍스트에 들어갑니다.

## 4. 지침이 잘 듣는지 시험

`SCENARIOS.md`의 요청 27개를 OpenCode에 입력하고 O/X를 기록합니다 (40~50분). 화면에 보이는 명령과 답변만으로 판정할 수 있습니다.

## 5. 사내 LLM이 지침을 다듬게 할 때

처음부터 다시 쓰게 하지 말고, 실제로 틀린 대화를 보여 주고 고칠 곳만 제안하게 합니다.

```text
아래는 네가 office-cli 스킬을 쓰다가 틀린 대화와, 지금의 전역 지침(AGENTS.md)이다.
같은 실수를 다시 하지 않게 하려면 AGENTS.md에서 무엇을 바꾸면 되는지 제안해라.
- 최대 3줄만 추가하거나 고친다. 전체를 다시 쓰지 않는다.
- 각 SKILL.md에 이미 있는 내용은 반복하지 않는다.
- 형식: "바꿀 줄: ... / 새 문장: ... / 이유: ..."
```

제안은 사용자가 보고 반영한 뒤, 같은 요청을 다시 해서 고쳐졌는지 확인합니다. 바뀐 AGENTS.md는 이 폴더에도 올려 두면 다음 검토 때 함께 봅니다.
