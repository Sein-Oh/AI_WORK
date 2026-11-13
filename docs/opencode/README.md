# 사내 PC OpenCode 배치 안내

사내 PC(Windows 11, 폐쇄망, OpenCode + 사내 LLM)에 도구 스킬과 전역 지침을 놓는 방법입니다.
저장소는 반입 후 `C:\AIKIT`에 `ml-kit`, `office-cli`가 있는 상태를 기준으로 합니다.

## 1. 도구와 스킬 배치

`C:\AIKIT`은 **OpenCode 프로젝트 폴더**입니다 (2026-11-13 사용자 결정). 우리 지침·스킬·도구는 `C:\AIKIT`을 프로젝트로 연 대화에서만 쓰이고,
다른 프로젝트의 대화에는 들어가지 않습니다. OpenCode는 **설정 파일 없이 폴더 모양만으로** 프로젝트의 지침(`AGENTS.md`)과 스킬(`.opencode\skills`)을
그 자리에서 읽습니다. `C:\AIKIT`을 프로젝트로 여는 방법: Web UI는 `C:\AIKIT\web.bat`(이 폴더에서 서비스를 띄움), TUI는 `C:\AIKIT`에서 `opencode`.

사내 LLM 설정과 셸·권한·폐쇄망 잠금은 모든 프로젝트에 적용되도록 **사용자 전역 설정** `%USERPROFILE%\.config\opencode`에 둡니다.
`C:\AIKIT\setup.bat`이 `C:\AIKIT\opencode.json`(셸·권한·폐쇄망 잠금)을 전역 `opencode.json`에 합칩니다 (원래 파일 보관, 관리자 권한 필요 없음.
시스템 PATH 등록에만 관리자 권한). 사내 LLM 주소가 든 `opencode.jsonc`는 사용자가 그 전역 폴더에 넣습니다 (`C:\AIKIT\opencode.jsonc`에 두면 setup이 전역 폴더로 옮김).

```text
C:\AIKIT\                  OpenCode 프로젝트
├── AGENTS.md              전역 지침 (이 프로젝트의 모든 대화에 적용)
├── .opencode\skills\      스킬: 폴더 하나 = 스킬 하나
│   ├── ml-train\SKILL.md  기본 스킬 13개 (명령은 C:\AIKIT\ml-kit\..., C:\AIKIT\office-cli\..., C:\AIKIT\pandas-kit\... 전체 경로)
│   ├── ...
│   └── <사용자 스킬>\      SKILL.md + 그 스킬의 scripts\ (OpenCode에서 스킬을 만들면 여기에 생김)
├── ml-kit\, office-cli\    여러 스킬이 함께 쓰는 도구 (.py)
├── opencode.json          셸·권한·폐쇄망 잠금 (배포 파일, setup이 전역 설정에도 합침)
└── opencode.exe, python\, setup.bat, web.bat, dev\, guide\, data\

%USERPROFILE%\.config\opencode\   OpenCode 전역 설정 (모든 프로젝트)
├── opencode.json          setup이 합친 셸·권한·폐쇄망 잠금 (원래 있던 다른 설정은 그대로)
└── opencode.jsonc         사내 LLM 연결 (PC마다, 사용자가 넣음)
```

| 무엇 | 파일 |
|---|---|
| 전역 지침 (`C:\AIKIT` 프로젝트) | `C:\AIKIT\AGENTS.md` |
| 스킬 (기본 13개, 사용자가 만든 것) | `C:\AIKIT\.opencode\skills\<이름>\SKILL.md` (스킬이 쓰는 스크립트는 그 폴더의 `scripts\`. 불러오면 OpenCode가 그 폴더 위치를 모델에 알려 줌) |
| 셸·권한·폐쇄망 잠금 | `C:\AIKIT\opencode.json` (배포 파일) -> setup이 `%USERPROFILE%\.config\opencode\opencode.json`에 합침 |
| 사내 LLM 설정 (provider, model, 모델 목록 숨기기 `disabled_providers`) | `%USERPROFILE%\.config\opencode\opencode.jsonc` (사용자가 넣음) |

- 반입: `C:\AIKIT`의 도구 폴더(`ml-kit`, `office-cli`, `pandas-kit`), `dev`, `guide`(사용 가이드), `.opencode\skills`의 기본 스킬 13개 폴더, `AGENTS.md`, `opencode.json`,
  `setup.py`·`setup.bat`·`web.bat`을 바꿉니다. `.opencode\skills` 폴더를 통째로 지우지 않습니다 (사용자 스킬이 함께 지워짐).
  `C:\AIKIT\python`, `opencode.exe`, 사용자 스킬 폴더, `data\`는 그대로 둡니다.
  처음 한 번 `C:\AIKIT\setup.bat`을 실행하고, 그 뒤 반입은 파일만 바꾸고 OpenCode를 다시 시작하면 됩니다
  (`opencode.json`을 바꾼 반입이면 setup을 다시 실행해 전역 설정에 합침. setup은 여러 번 실행해도 됩니다).
- 예전 구성에서 옮길 때 (2026-11-13 이전 배포): setup이 사용자 환경 변수 `OPENCODE_CONFIG_DIR=C:\AIKIT`를 지웁니다. `C:\AIKIT\opencode.jsonc`는
  setup이 `%USERPROFILE%\.config\opencode\opencode.jsonc`로 옮깁니다 (원래 파일은 보관 폴더로. 전역에 내용이 다른 파일이 이미 있으면 둘 다 두고 알림). `C:\AIKIT\skills`의 사용자 스킬은
  `C:\AIKIT\.opencode\skills`로 옮긴 뒤 `C:\AIKIT\skills` 폴더를 지웁니다.
- 바뀐 설정이 안 보이면: OpenCode 백그라운드 서비스는 처음 뜰 때의 환경 변수를 계속 씁니다. `C:\AIKIT\web.bat stop` 뒤 다시 실행하고,
  `C:\AIKIT\opencode.exe debug paths`의 `config`가 전역 설정 폴더(`%USERPROFILE%\.config\opencode`)인지 봅니다 (setup도 끝에 확인).
- Web UI 창(web.bat이 Edge·Chrome 앱 창으로 엶)을 닫으면 백그라운드 서비스도 꺼집니다 (2026-11-13 사용자 결정). 화면·트레이에 보이지 않는 PowerShell이
  `C:\AIKIT\data\web_profile`을 쓰는 브라우저가 없어지면 `opencode service stop`을 실행합니다. Edge·Chrome이 없어 기본 브라우저로 연 경우는
  예전처럼 서비스가 계속 돌고 `web.bat stop`으로 끕니다. 같은 서비스에 붙은 다른 OpenCode 창(TUI 등)이 있으면 함께 끊길 수 있습니다.
- Web UI의 기본 프로젝트 폴더는 `C:\AIKIT` 자체입니다 (`web.bat`이 서비스를 여기서 띄움). 처음 한 번 화면에서 프로젝트 열기로 `C:\AIKIT`을 고릅니다.
  다른 폴더를 프로젝트로 열면 우리 지침·스킬은 들어가지 않습니다 (사내 LLM과 폐쇄망 잠금은 전역이라 그대로).

**ripgrep(rg.exe)은 배포하지 않습니다** (2026-11-12 사용자 결정, 집 PC에서 rg를 치우고 인터넷을 막은 채 확인).
OpenCode 2.2.5는 rg 없이 스킬을 불러옵니다(`scripts\` 폴더가 있는 사용자 스킬 포함). rg가 필요한 것은 OpenCode의 파일 찾기 도구
grep·glob과 Web UI의 `@` 파일 자동 완성뿐이고, rg가 없으면 인터넷에서 받으려다 `ripgrep execution failed`로 실패합니다.
그래서 `opencode.json`의 `"permission"`에서 `"grep": "deny"`, `"glob": "deny"`로 두 도구를 꺼 둡니다. 모델은 처음부터 셸 명령
(`Get-ChildItem`, `Select-String`)으로 파일을 찾습니다. Web UI에서는 `@` 대신 파일 경로를 직접 적습니다.
(OpenCode 1.x는 스킬 불러오기에도 rg가 필요했습니다. 2026-10-29 실기: 대화 6개 모두 실패.)

**설정에서 확인할 것** (`setup.py`가 전역 설정 폴더의 `opencode.json`과 `opencode.jsonc`를 합쳐 보고 WARN으로 알림, `opencode.jsonc`는 고치지 않음).
`"shell"`, `"permission"`, `"autoupdate"`, `"share"`, `"lsp"`, `"formatter"`는 setup이 전역 `opencode.json`에 합치니 `opencode.jsonc`에는 적지 않습니다
(두 파일에 같은 키가 있으면 setup이 WARN). 집 PC opencode 시험(2026-11-07)에서 앞의 셋이 결과를 흐렸습니다.
- `"shell": "powershell"` (`opencode.json`): `"cmd"`면 스킬·전역 지침의 PowerShell 명령(인코딩 줄, 작은따옴표)이 맞지 않습니다.
- `"instructions"`: 비워 둡니다. 전역 지침 `AGENTS.md`는 `C:\AIKIT` 프로젝트에서 따로 읽히고, 여기에 옛 지침을 넣으면 지금 없는 명령을 씁니다.
- `"permission": {"external_directory": "allow"}` (`opencode.json`): opencode 1.19.5 `run`은 작업 폴더 밖 쓰기를 묻지 않고 거절해 모델이 멈췄습니다.
- 모델 목록 (PC마다 사용자가 정함, setup은 없어도 알리지 않음): 폐쇄망에서 쓸 수 없는 기본 모델(opencode Zen 무료 모델, models.dev 목록의 인터넷 모델)을
  `opencode models`와 모델 선택에서 숨기려면 `"disabled_providers": ["opencode"]`(적은 것만 숨김) 또는 `"enabled_providers": ["<사내 provider>"]`
  (적은 것만 남기고 모두 숨김, 둘 다 있으면 숨기기가 이김). opencode 2.2.7 소스로 확인(`config/normalize.ts`). 회사 PC는 `disabled_providers`로
  opencode를 숨기고 집 PC는 보이게 둠 (2026-11-13). `enabled_providers`를 적으면 이름이 틀릴 때 사내 LLM도 숨겨지므로 `setup.py`가 `"provider"`·`"model"`과 맞는지 봅니다.
- `"autoupdate": false` (`opencode.json`): 폐쇄망에서는 업데이트를 받을 수 없으니 확인을 시도하지 않게 합니다. 원격 모델 목록 받기도 끄려면 환경 변수
  `OPENCODE_DISABLE_MODELS_FETCH=1`(설정 파일 키는 없음, `enabled_providers`만으로도 목록에는 나오지 않음).
- `"permission"`의 `"webfetch": "deny"`, `"websearch": "deny"` (`opencode.json`): 폐쇄망에서 모델이 인터넷 검색·웹 페이지 열기를 시도하지 못하게 막습니다
  (2026-11-11 회사 PC: 스킬 만드는 법을 묻자 인터넷 검색을 시도함. 전역 지침 2장에도 금지를 적었지만 설정으로 막는 것이 확실함. opencode 2.2.7 소스의
  도구 이름 기준이고 회사 PC 2.2.5에서는 확인 전. `setup.py`가 없으면 INFO로 알림).
- 폐쇄망 잠금 (`opencode.json`, 2026-11-12 사용자 결정): 폐쇄망에서 되지 않는 기능은 모두 잠급니다. `setup.py`가 하나라도 빠지면 INFO로 알림
  (`OFFLINE_DENY`, `OFFLINE_SET`).
  - `"permission"`의 `"browser": "deny"`(브라우저 도구), `"grep": "deny"`, `"glob": "deny"`(파일 찾기, rg.exe가 필요해 위 rg 설명 참고).
  - `"share": "disabled"`: 세션 공유 끔 (공유하면 대화가 opencode.ai로 보내짐).
  - `"lsp": false`, `"formatter": false`: 언어 서버·서식 도구 끔 (파일 종류에 따라 인터넷에서 내려받으려 함. 코드 편집용 기능이라 우리 일에는 쓰지 않음).
  - 집 PC 2.2.5에서 LLM 주소 말고 모든 바깥 연결을 막고 기록한 결과: 잠근 뒤 스킬 불러오기·파일 만들기·python 실행이 되고, 대화·서버 시작 모두
    LLM 말고는 바깥 연결 시도가 없음 (모델 목록 캐시를 지워도 models.dev에 접속하지 않음).
- 사내 PC `opencode.jsonc` 예 (provider 이름·주소·모델은 PC마다 다름, 키는 파일에 적지 말고 opencode 로그인·환경 변수로):
  ```jsonc
  {
    "model": "corp/gemma",
    "disabled_providers": ["opencode"],
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

1. `AIKIT/AGENTS.md`(PC에서는 `C:\AIKIT\AGENTS.md`)를 OpenCode가 `C:\AIKIT` 프로젝트에서 그 자리에서 읽습니다 (1장 표, 복사하지 않음).
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

`C:\AIKIT\dev`에서 `python run_all_tests.py --label 회사PC --only gemma`로 시험 요청(`AIKIT/dev/agent_scenarios.py`)을 OpenCode에 보내고 채점합니다 (20~60분, [`실기테스트.md`](실기테스트.md) 3장).

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
