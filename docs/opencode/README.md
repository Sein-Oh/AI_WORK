# 사내 PC OpenCode 배치 안내

사내 PC(Windows 11, 폐쇄망, OpenCode + 26B~30B 모델)에 도구 스킬과 전역 지침을 놓는 방법입니다.
저장소는 반입 후 `C:\AIKIT`에 `ml-kit`, `office-cli`가 있는 상태를 기준으로 합니다.

## 1. 도구와 스킬 배치

`C:\AIKIT`이 곧 OpenCode 설정 폴더입니다. `C:\AIKIT\setup.bat`이 사용자 환경 변수 `OPENCODE_CONFIG_DIR=C:\AIKIT`을 설정합니다
(관리자 권한 필요 없음. 시스템 PATH 등록에만 관리자 권한). OpenCode는 **설정 파일 없이 폴더 모양만으로** 지침과 스킬을 읽고,
복사하지 않고 그 자리에서 읽습니다 (2026-11-12, OpenCode 2.2.5에서 확인).

```text
C:\AIKIT\
├── AGENTS.md              전역 지침: 모든 대화에 적용
├── skills\                스킬: 폴더 하나 = 스킬 하나 (OpenCode가 이 아래를 찾음)
│   ├── ml-train\SKILL.md  기본 스킬 11개 (명령은 C:\AIKIT\ml-kit\..., C:\AIKIT\office-cli\... 전체 경로)
│   ├── ...
│   └── <사용자 스킬>\      SKILL.md + 그 스킬의 scripts\
├── ml-kit\, office-cli\    여러 스킬이 함께 쓰는 도구 (.py)
├── opencode.json          셸·권한 기본 설정 (배포 파일, 스킬 경로는 적지 않음)
├── opencode.jsonc         사내 LLM 연결 (PC마다, 사용자가 넣음, git 제외)
└── opencode.exe, python\, setup.bat, web.bat, dev\, guide\, data\
```

| 무엇 | 파일 |
|---|---|
| 전역 지침 | `C:\AIKIT\AGENTS.md` |
| 스킬 (기본 11개, 사용자가 만든 것) | `C:\AIKIT\skills\<이름>\SKILL.md` (스킬이 쓰는 스크립트는 그 폴더의 `scripts\`. 불러오면 OpenCode가 그 폴더 위치를 모델에 알려 줌) |
| 한 프로젝트에서만 쓰는 지침·스킬 | `<프로젝트 폴더>\AGENTS.md`, `<프로젝트 폴더>\.opencode\skills\<이름>\SKILL.md` (그 폴더에서 연 세션에만 더해짐, `skill.py save --project`) |
| 셸·권한·웹 도구 막기·업데이트 끄기 | `C:\AIKIT\opencode.json` (배포 파일) |
| 사내 LLM 설정 (provider, model, `enabled_providers`) | `C:\AIKIT\opencode.jsonc` (사용자가 넣음) |

- 반입: `C:\AIKIT`의 도구 폴더(`ml-kit`, `office-cli`), `dev`, `guide`(사용 가이드), 기본 스킬 11개 폴더(`skills\ml-train` 등), `AGENTS.md`, `opencode.json`,
  `setup.py`·`setup.bat`·`web.bat`을 바꿉니다. `skills` 폴더를 통째로 지우지 않습니다 (사용자 스킬이 함께 지워짐).
  `C:\AIKIT\python`, `opencode.exe`, `opencode.jsonc`, 사용자 스킬 폴더, `data\`는 그대로 둡니다.
  처음 한 번 `C:\AIKIT\setup.bat`을 실행하고(환경 변수, PATH), 그 뒤 반입은 파일만 바꾸고 OpenCode를 다시 시작하면 됩니다
  (setup은 다시 실행해도 됩니다: 확인과 서비스 끄기).
- 예전 구성에서 바꿀 때 setup이 옮깁니다 (원본은 `C:\AIKIT\data\setup_backup_<시각>`에 보관):
  - 2026-11-12 전 (`%USERPROFILE%\.config\opencode`로 복사): 그 폴더의 우리 스킬·`AGENTS.md`는 지우고, 사용자 스킬과 `opencode.jsonc`는 `C:\AIKIT`으로.
  - 2026-11-12 낮 (`C:\AIKIT\opencode`가 설정 폴더, 스킬은 도구 폴더 안): `opencode.jsonc`와 사용자 스킬은 맨 위로, 나머지와 도구 폴더 안의 `skills`는 지움.
- office-cli-test(실기 보조 스킬)는 개발용이라 `python C:\AIKIT\setup.py --test-skill C:\AIKIT\dev\office-cli\tools\office-cli-test\SKILL.md`로
  `C:\AIKIT\skills`에 둡니다.
- 바뀐 설정이 안 보이면: OpenCode 백그라운드 서비스는 처음 뜰 때의 환경 변수를 계속 씁니다. `C:\AIKIT\web.bat stop` 뒤 다시 실행하고,
  `C:\AIKIT\opencode.exe debug paths`의 `config`가 `C:\AIKIT`인지 봅니다 (setup도 끝에 확인).
- Web UI의 기본 프로젝트 폴더는 `C:\AIKIT` 자체입니다 (`web.bat`이 서비스를 여기서 띄움, 2026-11-12 사용자 결정). 처음 한 번 화면에서
  프로젝트 열기로 `C:\AIKIT`을 고릅니다. 이 폴더의 `AGENTS.md`는 전역 지침이면서 프로젝트 지침 자리이기도 해서, 두 번 들어가는지 집 PC에서
  확인합니다 (HANDOFF 3-17). 프로젝트별 지침·스킬이 필요한 일은 따로 만든 폴더(예: `D:\프로젝트A`)를 프로젝트로 엽니다.

**ripgrep(rg.exe)이 필요합니다.** OpenCode는 스킬을 불러올 때 ripgrep을 쓰고, 없으면 인터넷에서 받으려 합니다.
폐쇄망에서는 받지 못해 스킬 불러오기가 `ripgrep execution failed`로 실패합니다(2026-10-29 실기: 대화 6개 모두).
ripgrep Windows용(`ripgrep-<버전>-x86_64-pc-windows-msvc.zip`)을 반입해 `rg.exe`를 PATH에 있는 폴더(예: `C:\AIKIT\python`)에 두고,
PowerShell에서 `rg --version`이 나오는지 확인한 뒤 OpenCode를 다시 시작합니다. 그 전까지는 전역 AGENTS.md 3장 3번대로 SKILL.md를 직접 읽게 됩니다.

**설정에서 확인할 것** (`setup.py`가 `opencode.json`과 `opencode.jsonc`를 합쳐 보고 WARN으로 알림, 고치지는 않음).
`"shell"`, `"permission"`, `"autoupdate"`, `"skills"`는 배포 파일 `opencode.json`에 이미 있으니 `opencode.jsonc`에는 적지 않습니다
(두 파일에 같은 키가 있으면 setup이 WARN). 집 PC opencode 시험(2026-11-07)에서 앞의 셋이 결과를 흐렸습니다.
- `"shell": "powershell"` (`opencode.json`): `"cmd"`면 스킬·전역 지침의 PowerShell 명령(인코딩 줄, 작은따옴표)이 맞지 않습니다.
- `"instructions"`: 비워 둡니다. 전역 지침 `AGENTS.md`는 설정 폴더에서 따로 읽히고, 여기에 옛 지침을 넣으면 지금 없는 명령을 씁니다.
- `"permission": {"external_directory": "allow"}` (`opencode.json`): opencode 1.19.5 `run`은 작업 폴더 밖 쓰기를 묻지 않고 거절해 모델이 멈췄습니다.
- `"enabled_providers"`: 폐쇄망에서 쓸 수 없는 기본 모델(opencode Zen 무료 모델, models.dev 목록의 인터넷 모델)이 `opencode models`와
  모델 선택에 나오지 않게, 사내 LLM provider 이름만 적습니다 (`"provider"` 블록의 이름, `"model"`의 `/` 앞). 적은 provider만 남고 나머지는 모두 숨겨집니다.
  특정 provider만 숨기려면 `"disabled_providers": ["opencode"]`(둘 다 있으면 숨기기가 이김). opencode 2.2.7 소스로 확인(`config/normalize.ts`: 모두 막고 적은 것만 허용).
  이름이 틀리면 사내 LLM도 숨겨지므로 `setup.py`가 `"provider"`·`"model"`과 맞는지 함께 봅니다.
- `"autoupdate": false` (`opencode.json`): 폐쇄망에서는 업데이트를 받을 수 없으니 확인을 시도하지 않게 합니다. 원격 모델 목록 받기도 끄려면 환경 변수
  `OPENCODE_DISABLE_MODELS_FETCH=1`(설정 파일 키는 없음, `enabled_providers`만으로도 목록에는 나오지 않음).
- `"permission"`의 `"webfetch": "deny"`, `"websearch": "deny"` (`opencode.json`): 폐쇄망에서 모델이 인터넷 검색·웹 페이지 열기를 시도하지 못하게 막습니다
  (2026-11-11 회사 PC: 스킬 만드는 법을 묻자 인터넷 검색을 시도함. 전역 지침 2장에도 금지를 적었지만 설정으로 막는 것이 확실함. opencode 2.2.7 소스의
  도구 이름 기준이고 회사 PC 2.2.5에서는 확인 전. `setup.py`가 없으면 INFO로 알림).
- 사내 PC `opencode.jsonc` 예 (provider 이름·주소·모델은 PC마다 다름, 키는 파일에 적지 말고 opencode 로그인·환경 변수로):
  ```jsonc
  {
    "model": "corp/gemma",
    "enabled_providers": ["corp"],
    "provider": {
      "corp": {
        "npm": "@ai-sdk/openai-compatible",
        "options": { "baseURL": "http://<사내 LLM 주소>/v1" },
        "models": { "gemma": {} }
      }
    }
  }
  ```

- 도구 오류 로그는 `C:\AIKIT\<도구>\data\logs\errors.txt`에 쌓이고, `python C:\AIKIT\dev\run_all_tests.py`가 결과 파일 끝에 붙여 줍니다 (오류 로그만 빨리 모으려면 `--only diag`, 1분). 실기 절차는 [`실기테스트.md`](실기테스트.md).
- 이전처럼 스킬 폴더에 도구 전체를 복사해 두었다면 그 폴더는 SKILL.md만 남기고 `C:\AIKIT`만 쓰세요. 두 곳에 도구가 있으면 어느 쪽이 쓰였는지 헷갈립니다.

## 2. 전역 지침 (AGENTS.md)

1. `AIKIT/AGENTS.md`(PC에서는 `C:\AIKIT\AGENTS.md`)를 OpenCode가 그 자리에서 읽습니다 (1장 표, 복사하지 않음).
2. 고쳐 쓸 곳은 없습니다. 도구 경로는 `C:\AIKIT` 기준으로 적혀 있습니다.
3. 팀원이 "지침·스킬 경로"를 물으면 에이전트가 답하도록 9장에 위치 표가 있습니다 (지침·스킬·설정 파일, 프로젝트별 지침·스킬, 오류 기록).

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
