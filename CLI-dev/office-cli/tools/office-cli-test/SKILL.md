---
name: office-cli-test
description: office-cli 도구의 Windows 실기 테스트(실제 Excel/Word/PowerPoint로 수십 개 항목 자동 확인)를 실행하고 결과를 보고한다. 사용자가 "office-cli 테스트", "실기 테스트 돌려줘", "오피스 도구 점검해줘"라고 요청할 때만 사용. 일반 문서 작업에는 쓰지 않는다(그때는 office-cli 스킬).
license: Proprietary
compatibility: opencode
metadata:
  version: "1.4.26"
  platform: windows
---

# office-cli-test 스킬 (Windows)

테스트 러너 위치는 `C:\CLI-dev\office-cli\tools\run_win_tests.py`이다. 아래 명령의 경로를 그대로 쓴다.
테스트는 실제 Office 창을 열고 닫으며 10~30분 걸린다. 명령은 **1회만** 실행한다.

## 1. 실행 전 (반드시)
사용자에게 한 번 묻고 "예"를 받은 뒤 실행한다:
"테스트 중 Excel/Word/PowerPoint 창이 열리고 닫힙니다. 열려 있는 Office 문서를 저장하고 모두 닫았나요? 테스트가 끝날 때까지 Office 창을 만지지 마세요."

## 2. 명령
명령 실행 도구의 timeout 은 **3600000**(60분)으로 지정한다.
PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 붙인다.

| 요청 | 명령 |
|---|---|
| 기본 (Outlook 제외) | `python "C:\CLI-dev\office-cli\tools\run_win_tests.py" --agent --skip-outlook` |
| 일부 그룹만 (예: A, C) | `python "C:\CLI-dev\office-cli\tools\run_win_tests.py" --agent --skip-outlook --only A,C` |
| Outlook 포함 (사용자가 요청한 경우만) | `python "C:\CLI-dev\office-cli\tools\run_win_tests.py" --agent` |
| 테스트 폴더 지정 | 위 명령 끝에 `--dir "D:\office_test"` |

- `--self-email` 은 실제 메일을 보낸다. 사용자가 주소를 주며 **발송 테스트를 분명히 요청한 경우에만** 붙인다.
- 그룹: A 창 유지, B 읽기 전용, C 되돌리기, D 내보내기, E 엑셀 편집, F PPT, G Outlook, H 연계, I 기능 확대, J 배치 점검, K 검토 기능(목차·메모·변경 추적·노트).

## 3. 결과 처리
stdout 의 **마지막 1줄 JSON** 만 읽는다 (stderr 진행 로그는 무시).

| 결과 | 사용자에게 보고할 내용 |
|---|---|
| `status: ok`, `failed` 가 빈 목록, `unverified_count` 0 | "테스트 통과" + `summary` (예: PASS=63, ENV-SKIP=2) |
| `status: ok`, `unverified_count` 1 이상 | "통과"라고 하지 않는다. `summary`, `failed`, "도구가 결과를 확인하지 못한 호출 N건", 그리고 "`error_log` 파일을 Claude에게 그대로 전달해 주세요" + `error_log` 경로 |
| `status: ok`, `failed` 가 있음 | `summary`, `failed` 목록, 그리고 "`error_log` 파일을 Claude에게 그대로 전달해 주세요" + `error_log` 경로 |
| `status: error` | `message` 와 `error_log` 경로를 그대로 전달 |

- `agent_directive` 가 `VERIFIED:`, `UNVERIFIED:`, `STOP_RETRY:`, `ERROR:` 중 무엇으로 시작해도 **다시 실행하지 않는다**.
- 실패 원인을 추측하거나 코드를 고치지 않는다. `error_log` 파일을 열어 요약하거나 고치지 않는다(분석은 Claude가 한다).
- `MANUAL-SKIP`(눈 확인 생략), `ENV-SKIP`(이 PC에 없는 기능) 은 실패가 아니다.
- 파일 경로는 결과 JSON의 값을 그대로 1번만 쓴다 (`file:///` 링크로 바꾸지 않는다). 이모지 금지.
