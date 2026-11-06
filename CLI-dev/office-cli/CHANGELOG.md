# CHANGELOG - office-cli

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`. 1.0.0부터 관리합니다.

## 테스트 정리

- `tests/`의 차수별 파일(`test_review_v14xx.py` 등 35개)을 기능별 11개 파일(`test_contract`, `test_excel_read`, `test_excel_write`, `test_word`, `test_ppt`, `test_outlook`, `test_reader`, `test_office`, `test_markdown`, `test_skill_docs`, `test_win_runner`)과 공용 도우미 `office_helpers.py`로 합침 (동작 변경 없음). 아래 기록에 나오는 옛 테스트 파일 이름은 정리 전 배치를 가리킴.
- 회사 PC 실기(2026-11-05, Windows) 테스트 수정 (도구 동작 변경 없음): 도구가 낸 명령을 `shlex.split`(posix)으로 나눠 Windows 경로의 백슬래시가
  지워지던 것을 `office_helpers.split_command`(Windows 명령줄 규칙)로. 매개변수 id가 긴 테스트(`test_number_search_linear_time`)는 `ids=`, conftest의
  `pytest_make_parametrize_id`로 긴 글자 id를 60자로 자름(Windows 환경변수 `PYTEST_CURRENT_TEST` 32767자 한도 ERROR). DRM이 python-pptx 기본 템플릿을
  지우고 python-docx `default.docx`를 바꾼 PC에서도 샘플을 만들도록 `sample_docs.new_pptx`는 코드로 만든 최소 OOXML, `new_docx`는 폴더 템플릿을 씀
  (건너뛰지 않고 시험). 테스트는 `pptx.Presentation()`/`docx.Document()`를 직접 부르지 않음.

## [1.4.77] - 2026-11-06

회사 PC 실기(2026-11-06, Gemma 종합 테스트 O1·O2·O5) 대응.

### 수정
- 닫힌 파일에 쓴 뒤 세션(문서를 닫고 백그라운드 Office를 끝낸 뒤)에서 시트·범위 객체를 읽어 OLE 0x800a01a8(개체 필요)로 실패하던 문제.
  저장은 이미 끝나 실제로는 바뀌었는데 도구는 실패를 보고했음: `excel write --cells`(결과의 칸 목록 `ws.Name`), `excel chart`(transfer 명령의
  `--sheet`), `excel comment --action list --range`(`rng.Address`). 회귀 테스트 `tests/test_session_scope.py`: with 세션 블록 밖에서 그 안의
  COM 객체 속성을 읽는 코드를 코드 모양으로 찾음 (가짜 COM은 닫힌 뒤의 객체를 흉내 내지 않음).
- `excel write --cells "C6=ft,C7=℃"`: 칸 구분은 `;`인데 쉼표로 나누면 C6에 `ft,C7=℃`가 통째로 들어감. 값 안에 `,<칸>=`(시트 붙은 칸 포함)가
  있으면 쓰지 않고 `;`로 바꾼 전체 명령을 `[추천]` (값 안의 쉼표 `1,234`·글·수식·따옴표로 묶은 값은 그대로).
  첫 `,<칸>=` 앞에 `=`가 있으면 값 자체가 `R1=10k,R2=4.7k` 같은 글이라 나누지 않음(검토: 그대로 실행하면 R2 칸에 쓰는 [추천]이었음),
  따옴표는 값 전체를 묶었을 때만 제외(`'ft',C7='x'`는 거절). `word fill --slots`도 같은 규칙, JSON의 객체·목록·참거짓 값은 거절.
  공백·`|`로 나눈 것(`C6=ft C7=20`: C6에 `ft C7=20`이 들어갔음)도 같은 방식으로 거절.
- `excel read --where "금액 >= 100, 부서 = 영업"`: 조건을 쉼표로 나누면 값 `100, 부서 = 영업` 하나로 읽어 0건이었음. 쉼표 뒤가 `<열> <연산자>`이면
  `;`로 고친 식을 `[추천]` (값 안의 쉼표 `1,000`, 따옴표로 묶은 글은 그대로).
- 인자 오류의 옵션 추천 (Gemma 실패 유형 감사): 철자 비교(difflib)로 고른 옵션을 `[추천]`(LLM이 묻지 않고 실행)으로 줘
  `--cell` -> `--all`(reader clear 캐시 전부 삭제, outlook reply 전체 답장), `--at` -> `--to`, `--text` -> `--target`, `--append` -> `--app`처럼 뜻이 다른
  추천이 30종쯤 있었음. 같은 뜻 표 `SAME_MEANING`(도구·명령별: `--file`/`--path`/`--input` -> `--target`, reader는 `--target` -> `--file`,
  `--output` -> `--save`, `--cell` -> `--range`(엑셀), `--rows` -> `--max-rows`, word write `--append` -> `--pos end`, excel read `--text` -> `--search`,
  outlook `--text` -> `--body`, `--max` -> `--limit`, `--json`·`--encoding`은 뺌)로 고친 전체 명령을 `[추천]`. 필수 옵션 누락 오류(`reader read --target`)에도.
  표에 없으면 철자 비교는 cutoff 0.85·앞 글자가 같을 때만 `[후보]`(사용자에게 물음), `--all`·`--overwrite`·`--to`·`--cc`·`--attach`·`--target` 등은 추천하지 않음.

### 기능
- `excel check --unit-col <열>`: 항목·값·단위 3열 표에서 행마다 단위 열을 읽음. 생략하면 값 열 바로 오른쪽 열의 머리글이 `단위`(`Unit`)일 때
  자동으로 씀(JSON `unit_col`). 전에는 값 칸·머리글에 단위가 없으면 모든 행이 `판단필요`(`칸 단위 없음`)라 Gemma가 메모 미리보기 없이 사용자에게
  판정을 되물었음. 실기 샘플 `규격대조.xlsx`에 단위 열(D)을 넣음 (기대: 일치 2, 불일치 2, 근거없음 1).
- `word fill --slots "업체명=한빛전자; 시험일=2026-11-04; 표2!R3C4=25"`: 값 파일 없이 바로 채움. Gemma는 값 파일이 없어 `--slots` JSON을
  지어냈고(PowerShell 5.1에서 따옴표가 깨짐), 이어 양식 문서를 `--file`로 줘 '인코딩 인식 실패'로 끝났음. `--file`과 `--slots`는 둘 중 하나,
  양식 문서(.docx 등)를 `--file`로 주면 "양식은 --target" 안내, 값이 없으면 두 방법을 안내(예시 값 없이 `<자리>=<값>`).

## [1.4.76] - 2026-11-05

도구가 내는 명령(`next_commands`, `confirm_commands`, `[추천]` 전체 명령)의 인자 따옴표를 실제 실행 셸에 맞춤.

### 수정
- `quote_arg`가 `"..."` + `\"`(cmd 기준)라 OpenCode가 쓰는 PowerShell에서 깨지던 값: 끝이 `\`인 경로(`"C:\a b\"`), 값 안의 `"`,
  `$`·백틱 확장(`=$A$1*2`), 따옴표 없던 `{{업체명}}`·`@a`·맨 앞 `#`·`'`·`;`·`(`, cmd의 `%NAME%`.
  - 실행 셸: OpenCode 셸 도구(v2.2.5 `shell/select.ts`)와 같은 순서로 `pwsh`가 PATH에 있으면 PowerShell 7, 없으면 Windows PowerShell 5.1
    (`rt.command_shell()`). `OFFICE_CLI_SHELL=pwsh|powershell|cmd|bash`로 강제.
  - 셸 특수문자가 없는 값(영숫자·한글과 `_-.:\/,=+@!`, 숫자로 읽힐 모양·맨 앞 `-`(옵션 제외)·`=`·`@`·`,` 등 제외)은 그대로.
    나머지는 PowerShell 작은따옴표 `'...'`(안의 `'`와 둥근 작은따옴표는 두 번).
  - Windows PowerShell 5.1은 값을 파이썬에 넘길 때 공백이 있으면 바깥만 `"..."`로 감싸고 안의 `"`는 그대로 두며 빈 값은 버리므로,
    `"` -> `\"`(앞 백슬래시 두 배), 감쌀 값의 끝 백슬래시 두 배, 빈 값 `'""'`를 미리 함. PowerShell 7.3+는 미리 바꾸지 않음(이중 이스케이프 방지).
  - cmd(`"..."`, 안의 `"`는 `""`, `%`는 `"^%"`)와 bash(`'...'`)는 최선 노력. ml-kit 2.9.68에 같은 코드(복사).
- 테스트: 셸 시뮬레이터 `CLI-dev/tests/shell_sim.py`(PowerShell 토크나이저, 5.1 레거시·7.3 Standard 명령줄 조립, Windows argv(UCRT) 해석,
  cmd, bash)와 `CLI-dev/tests/test_shell_quote.py`(까다로운 값 100여 개 x 셸 4종 x 두 도구, 셸 판별, 채점기 `command_segments`).
  `office_helpers.split_command`는 시뮬레이터를 씀(작은따옴표를 읽음). excel `write --cells`·word `fill`의 `confirm_commands`를 셸마다 끝까지 실행.

## [1.4.75] - 2026-11-05

독립 검토(11차) 수정. 확인 없이 실행되는 길(파괴적 confirm, 문서를 바꾸는 next_commands, 파괴적 오타 추천)과 숫자 검색 층 A 허용 목록.

### 기능
- `confirm_commands` 항목에 `"destructive": true`(덮어쓰기·삭제·메일 발송): `--overwrite`가 든 excel `write --cells`·`range`·`comment`, word `fill`,
  excel `range --action delete`, `comment --action delete`, outlook `send`. 전역 지침 3장 6번의 "묻지 말고 끝까지" 예외가 이 항목에는 적용되지 않음
  (늘 묻는다, 5장과 같음). 그 밖의 confirm_commands(빈 칸 쓰기, 새 메모, 다른 문서에 넣기)만 예외. `rt.next_command(..., destructive=True)`.
  전역 지침 3장 6번·5장, 네 스킬 규칙 3, `CLI-dev/README.md` 공통 키, 실기 러너 G1, 시나리오 O5 judge(빈 자리라 destructive 아님)를 함께 고침.
- 문서를 바꾸는 이어 하기 명령을 `next_commands`에서 `confirm_commands`(destructive 아님)로 옮기고 directive는 "사용자가 원하면": excel `read --save`(조건 읽기 포함)
  뒤 word/ppt `write`, excel `write`(표·수식) 뒤 `chart`, excel `chart` 뒤 `transfer`, word·ppt `read --save` 뒤 `write`·`transfer`,
  reader `read --save`·`search --save` 뒤 word/ppt/excel `write`(`_write_confirm_commands`), outlook `read`/`current --save` 뒤 `word.py write`(대상 생략 = 열린 활성 문서).
  `next_commands`에는 읽기·미리보기·파일 저장(문서를 바꾸지 않음)만 남음. 예외: `CONTINUE:`와 `drm_or_dialog`의 같은 명령(사용자가 시킨 그 명령).
  회귀: `tests/conftest.py`의 자동 점검(모든 테스트의 stdout JSON에서 `next_commands`에 문서를 바꾸는 명령이 있으면 실패).
- 명령 오타 추천: 가장 비슷한 명령이 되돌리기 어려운 명령(`send`, `clear`, `undo`)이면 `[추천]`(LLM이 묻지 않고 실행)이 아니라 `[후보]`(사용자에게 물음).
  `sent`, `sned`, `end` -> `[후보] send`, office `redo` -> `[후보] undo`.

### 수정
- reader 층 A 허용 목록 완화 (글자가 같은데 형식 때문에 층 B로 떨어지던 것): 숫자 앞의 여는 따옴표·괄호(U+201C, U+2018, U+300C, U+300E, U+3010, U+3008,
  U+300A, 앞이 숫자면 아님), 짝이 맞는 ASCII `"` `'`, 전각 콜론 U+FF1A, 목록 머리 U+2022(줄 시작·공백 뒤). 단위 뒤의 닫는 짝(U+201D, U+2019, U+300D,
  U+300F, U+3011, U+3009, U+300B), U+3001, U+00D7(뒤만, 숫자 앞 U+00D7은 계속 층 B), `±` 뒤 숫자(`220 V±10 %`의 `220 V`), 각주 `*1`, 괄호 `(`
  (막는 꼬리 `(rms)` 등과 `(-)`, `(1)`, `(실효값)`류 밖: `220 V(60 Hz)`, `1 MPa(10 kgf/cm²)`, `2 h(1회)`). 10차 `cases.txt`와 `number_literal_cases.tsv`의
  notA는 그대로 notA. 11차 검토 모음 층 A 비율 6/21 -> 20/21(`100 mm`+U+00D7+`50 mm`의 `50 mm`만 층 B), 8차 `pos.tsv` 37/60 -> 41/60.
  회귀 표 `number_literal_cases.tsv` "11차 검토" 묶음(A 24, notA 13).
- outlook `send`: `Send()` 예외 뒤 발송 원장에 실패 시각(`failed`)을 남기고, 그 뒤 send 거부 문구는 "발송이 실패한 초안입니다 ... 작성 창에서 직접 [보내기]"
  ("이미 send로 보냈습니다"와 구분). 여전히 보내지 않음(fail-closed, '보냄' 기록 유지). 회귀 `test_send_failure_is_recorded_and_refused_with_own_message`.
- 없앤 `--send`의 `[추천]` 명령에서 받지 않는 다른 옵션(`new ... --send --dry-run`의 `--dry-run`)도 뺌.

## [1.4.74] - 2026-11-05

### 수정
- 관리자 권한 판단(`_runtime.is_elevated`)을 `IsUserAnAdmin`에서 토큰 상승 종류(`TokenElevationTypeFull`)로 바꿈. UAC가 꺼진 회사 PC
  (`EnableLUA`=0, 모든 프로그램이 High 무결성)에서는 Office도 같은 권한이라 열린 문서가 보이는데, 문서가 없을 때마다 "일반 PowerShell에서
  다시 실행" 안내(`elevated_process`)를 내서 사용자가 따를 수 없는 행동을 시켰음. `CLI-dev/run_all_tests.py`의 Gemma 단계 경고도 같은 판단.
  회귀 테스트: `test_office.py` `test_is_elevated_only_for_uac_split_token`.

## [1.4.73] - 2026-11-04

독립 검토(10차) 수정. 발송 안전과 숫자 검색 층 A 경계.

### 기능
- 사용자 확인이 필요한 명령은 `next_commands`가 아니라 새 키 `confirm_commands`로 줌 (발송 안전): 전역 지침 3장 6번 "next_commands는
  글자 그대로 실행"을 경량 LLM이 기계적으로 따라도 메일이 나가거나 문서가 바뀌지 않게. `needs_user_confirm` 표시는 없앰.
  대상: outlook `new`·`reply`·`forward`의 `send --id <draft_id>`, 미리보기(`--dry-run`) 뒤 실제 실행 명령 전부(excel `write`·`range`·`comment`,
  word `fill`, reader `clear`). excel `check`의 `next_commands[0]`(메모 미리보기, 변경 없음)은 그대로, 그 미리보기 결과가 `confirm_commands`를 줌.
  전역 지침(`CLI-dev/opencode/AGENTS.md` 3장 6번 예외, 4장, 5장), 네 스킬 규칙 3, `CLI-dev/README.md` 공통 키·R8, 실기 러너 G1·G3,
  시나리오(`agent_scenarios.yaml` O1·O3·O5, `SCENARIOS.md` 13-1)를 함께 고침.
- outlook `send` 발송 원장(`<데이터 폴더>/outlook_drafts.json`, 번호 기억 파일과 같은 잠금): `new`·`reply`·`forward`가 만든 초안의 EntryID를 기록하고,
  `send`는 원장에 있고 아직 보내지 않은 초안만 보냄. 보내기 전에 잠금 아래 '보냄'을 먼저 기록. 원장을 읽지 못하면 보내지 않음(fail-closed),
  기록하지 못한 초안에는 send 명령을 주지 않음(`UNVERIFIED:`, 사용자가 작성 창에서 보냄). 7일(번호 기억과 같음)이 지난 초안은 send로 보내지 않음.
- 스니펫 강조: 층 A 자리(글자 그대로)와 숫자 없는 검색어는 `**`, 숫자가 든 검색어의 층 B 자리(표기가 다른 곳)는 `[[ ]]`.

### 수정
- [높음, 발송 안전] 전역 지침·office-mail 규칙 3의 "next_commands는 그대로 실행"이 send의 사용자 확인을 덮음: 위 `confirm_commands`.
- [중간, 발송 안전] 같은 초안을 두 번 보낼 수 있던 것(동시 실행, 보낼 편지함의 `Submitted=True, Sent=False`): `Submitted`도 확인하고 원장 잠금.
- [중간, 발송 안전] 지운 초안(PST에서 EntryID 그대로)과 사용자가 쓰다 둔 다른 초안도 send로 나가던 것: 메일의 폴더가 임시 보관함(기본 계정 또는
  그 메일 저장소의 임시 보관함)이 아니거나 원장에 없는 초안이면 거부.
- [중간] send의 id 해석 오류가 "목록을 다시 보고 다시 실행"(`next_commands` = 목록)을 주던 것: 발송 거부 안내(`STOP_RETRY: 발송하지 않았습니다`)만, `next_commands` 없음.
- [낮음] 없앤 `--send`를 `new`·`reply`·`forward`에 주면 exit 2 + "--send는 없습니다. 초안을 만든 뒤 사용자 확인을 받아 send를 실행" + `[추천] <--send를 뺀 전체 명령>`.
- [치명] 숫자 앞·단위 뒤의 글자를 금지 목록에서 허용 목록으로 (검토마다 새 글자가 나왔음): 숫자 앞은 공백류, 줄 시작, `(` `[` `~` U+FF5E U+301C `=`,
  `:`(앞이 숫자면 아님: `1:5 V`), 한글, `DC`/`AC`, 값 구분 `/`, 붙여 쓴 범위 대시만 층 A. 단위 뒤는 공백류, 줄 끝, `)` `]` `;` `~` U+FF5E U+301C,
  `,` `.`(뒤가 숫자·영문이 아닐 때), `:`(뒤가 공백류·줄 끝), 한글, 값 구분 `/`, `DC`/`AC`, 붙여 쓴 괄호 허용 목록(`(DC)` `(AC)` `(Max)` `(Min)` `(Typ.)` 등),
  범위 대시(`10 m-20 m`)만 층 A. 그 밖 글자는 층 B: 10차 사례 `1`+U+FF0C+`500 V`, `1`+U+FF0E+`5 V`, `1·5 V`, `△40 ℃`, `▲40 ℃`, `1`+U+2018+`500 V`,
  ``1`500 V``, `1_500 V`, `1′500 V`, `1:5 V`, `1÷5 V`, `1\5 V`, `½5 A`, `5 V`+U+3383.
- [치명] 앞이 숫자뿐인 `/` 값 구분에서 분수(`15/16 V`, 분모 2·4·8·16·32·64)와 비율(`10/100 V`, 1·10^n)을 뺌. 위 사례 모두 `number_literal_cases.tsv`에 notA로.

### 회사 PC 확인
- 실기 러너 G1(`confirm_commands`, `next_commands` 없음), G3(두 번째 send 거부). 임시 보관함이 기본 계정이 아닌 계정(공유 사서함 답장 등)의 초안도
  send가 보내는지(메일 저장소의 임시 보관함 비교).
- 미리 저장한 초안은 사용자가 작성 창에서 보내거나 지우기 전까지 임시 보관함에 남음 (README outlook 절).

## [1.4.72] - 2026-11-04

독립 검토(9차) 수정. search `VERIFIED`(층 A)의 뜻(글자 그대로 + 숫자·단위 낱말 경계, 문맥 의미는 스니펫으로)은 그대로.

### 기능
- outlook 발송 흐름 변경 (발송 안전): `new`·`reply`·`forward`에서 `--send`를 없앰. 세 명령은 초안을 임시 보관함에 저장하고 작성 창만 띄우며,
  결과에 `draft_id`와 그 초안을 보내는 `next_commands[0]` = `outlook.py send --id <draft_id>`(`needs_user_confirm: true`)를 줌.
  새 명령 `send --id`는 그 초안(같은 항목)만 보냄: 보내지 않은 초안(`Sent == False`)이 아니면(이미 보냄, 지움, 받은 메일) 보내지 않고 `STOP_RETRY:`.
  첨부를 확인하지 못했거나 초안을 저장하지 못했으면 send 명령을 주지 않음(`UNVERIFIED:`, 사용자가 작성 창에서 직접 보냄).
  office-mail 스킬, 전역 지침(`CLI-dev/opencode/AGENTS.md` 5장), 실기 러너 G1~G3, 시나리오 O3 금지 명령을 함께 고침.

### 수정
- [치명] 1.4.71의 `/` 완화가 넓어 비율·서지 파형·분수가 층 A: `10/350 μs`의 `350 μs`, `15/16 inch`의 `16 inch`, `0.1 mm/1000 mm`의 `0.1 mm`·`1000 mm`,
  `2 ℃/1 min`, `5 %/1000 h`. `/` 값 구분은 양쪽이 숫자+단위(`100 V/220 V`, `85 ℃/85 %RH`, `10 kHz/50 kHz`)이고 `/` 뒤 값이 1·10^n(비율의 '당')이 아닐 때만,
  앞이 숫자뿐이면 주파수·전압(Hz, kHz, MHz, GHz, V, kV, mV)과 두 자리 이상 정수만(`50/60 Hz`, `10/50 kHz`, `110/220 V`).
- [치명] AND 검색 스니펫이 첫 검색어 자리만 창으로 잡아 멀리 있는 숫자 검색어의 층 A 자리가 빠짐: 검색어마다 창(앞 최소 15자), 겹치지 않으면
  조각을 ` / `로 이음(`...**온도**... / ...범위 - **40 ℃**에서...`).
- [중요, 발송 안전] 1단계(`--send` 없이) 초안 창이 열린 채 `--id` 없는 `reply/forward --send`가 열린 창의 초안(`ActiveInspector().CurrentItem`)을
  원본으로 잡아 초안에 답장·전달을 보내거나, 같은 메일이 두 번 나갈 수 있던 것: 위 발송 흐름 변경 + 답장·전달은 보내지 않은 메일(`Sent == False`)을
  원본으로 쓰지 않음(`STOP_RETRY:` "원본 메일을 --id로 지정하세요" + 목록 `next_commands`).
- [중간] 단위 없는 검색어에 DC/AC 예외가 적용돼 `12`가 `AC12 단자대`·`DC12 커넥터`, `100`이 `AC100`에 층 A: 검색어에 단위가 있을 때만(`DC24 V`의 `24 V`).
- [중간] 붙은 기호를 층 A에서 막음: 천 단위 따옴표(`1'500 W`, U+2019), 분수 빗금 U+2044·전각 U+FF0F·나눗셈 U+2215(`1`+U+2044+`4 W`, `10 V`+U+FF0F+`m`),
  곱하기(`1×2 mm`, `1*2 mm`), 대시로 쓴 한글 U+3161, 폭 없는 공백(U+200B 등)·전각 공백 U+3000을 사이에 둔 빼기, 단위 지수(`10 m-1`), `5 % R.H.`,
  `3 dB[A]`, `10 kg f`. 붙여 쓴 범위 대시(`40 ℃`+U+2013+`85 ℃`의 `85 ℃`)는 층 A(기호 단위 뒤만).
- 9차 검토 쌍 118개를 `number_literal_cases.tsv`에 추가.

## [1.4.71] - 2026-11-04

독립 검토(8차) 수정. 숫자 검색 층 A의 뜻을 확정하고 문서화: search의 `VERIFIED`(층 A)는 "검색어 글자열이 그 자리에 글자 그대로,
숫자·단위 낱말 경계가 맞게 있다"는 뜻이다. 앞뒤 문맥의 의미(띄어 쓴 부호 `- 40 ℃`, 범위 `~`, 이하/이상, 한글 꼬리 `(실효값)`)는 판정하지 않고
(8번의 검토에서 문맥 의미 판정 규칙은 끝없이 빈틈이 생김) check(늘 UNVERIFIED)와 사용자의 몫으로 둔다. 대신 스니펫이 그 자리를 반드시 보여 준다.

### 기능
- reader `search`: 숫자가 든 검색어의 directive에 "결과를 알릴 때 snippet의 앞뒤(부호, 범위 ~, 이하/이상)를 그대로 함께 전달하세요" (`NUM_CONTEXT_NOTE`).
  office-spec 스킬과 전역 지침(`CLI-dev/opencode/AGENTS.md`)에 같은 규칙.
- 층 A 완화(글자가 같은데 형식 때문에 떨어지던 것): 단위 뒤 `/` 다음이 숫자·부호면 값 구분(`100 V/220 V`, `85 ℃/85 %RH`), `/` 앞이 같은 종류 값이면
  값 구분(`50/60 Hz`의 `60 Hz`, 분수 `1/4 W`는 아님), 단위 뒤 괄호는 막는 목록(`(rms)`, `(p-p)`, `(g)`, `(a)`·`(A)`, `(G)` 등)만 층 B(`(DC)`, `(AC)`, `(Max)`,
  `(Min)`, `(Typ.)`은 층 A), 단위에 붙은 대문자 `DC`/`AC`(`24 VDC`, `DC5V`, `AC220V`), 영문 접두와 숫자 사이 띄어쓰기 0~1칸 차이(`IP 67`, `M 10`, `IEC61000-4-2`).
  8차 정상 문장 60개 중 층 A 28개 -> 41개 (나머지는 표기 자체가 다른 `°C`/`℃`, `㎜`, `µ`/`μ`, `KΩ`, `1,000`/`1000`, `+/-`/`±` 등으로 층 B 그대로).
- excel `check --header-row 0` = 머리글 없음.

### 수정
- [중간] 스니펫·근거 문장(`make_snippet`, `match_sentence`)과 check의 조건 읽기·`spec_text`(`near_conditions`)가 층 B 매처 자리를 써서, 층 A로만 맞은 단위
  (`- 40 ℃`의 `40 ℃` 등)에서 스니펫에 값이 빠지고 글 앞부분만 나오던 것: 층 A 자리를 먼저(`term_span`), 층 A 자리가 든 문장을 먼저, 일치 앞 최소 15자.
- 낱말 수준 오일치를 층 A에서 막음: 띄어 쓴 천 단위(`1 500 W`, `10 000 h`, NBSP·좁은 공백 포함), 곱 단위 연산자(U+22C5, U+2219, U+2022, U+2217)와 띄어 쓴 단위 이어짐
  (`5 N m`, `10 Ω cm`, `10 kV A`, `10 mm Hg`, `10 mm H2O`, `2.5 mm SQ`, `5 % RH`, `0.5 % FS`, `3 dB μV`), 띄어 쓴 `/단위`·위첨자(`25 ℃ /W`, `10 m ³`),
  숫자로 끝나는 검색어 뒤의 `-숫자`(`61000-4`와 `IEC 61000-4-2`, `1234`와 `1234-5`; 회귀 표의 `10`/`10-3`도 층 B로 바뀜),
  같은 칸의 붙은 빼기(`-\u00a040 ℃`, `온도 −\t40 ℃`; 표 빈 칸 `\t-\t40`은 층 A 그대로).
- [중간] check 남은 행 명령: 원래 범위가 1행부터(머리글 없음)면 `--header-row 0`을 넘기고 0을 "머리글 없음"으로 처리
  (전에는 0이 조용히 무시되어 남은 범위 바로 위 데이터 행을 값 열 단위 머리글로 읽어, 이어서 대조한 판정이 처음 대조와 달랐음).
- 회귀 표 `number_literal_cases.tsv`에 8차 낱말 수준 쌍(층 A 금지)과 완화 대상(층 A) 추가, `- 40 ℃` 유형은 "층 A + 스니펫에 `- 40` 포함" 별도 테스트.

## [1.4.70] - 2026-11-04

독립 검토(7차) 수정. reader 숫자 정확 일치는 일곱 번의 검토마다 부호·대시·단위 정규화의 새 오일치가 나와(7차: 탭 너머 `-`가 부호로
읽혀 `-40 ℃` exact, 전각 빼기 `－40 ℃`가 `40 ℃` exact, `Tstg-40℃`, `10 mm²`, `5 V rms`, `100 pA`/`100 Pa` 등) 정규화 규칙을 더 땜질하지 않고
두 층 구조로 바꿈. 회귀 표(6·7차 검토의 오일치 쌍 전부 + 기본 표기)를 먼저 만들어 재현한 뒤 고침.

### 구조
- 숫자가 든 검색어의 두 층 판정: 층 A(글자 그대로, `_Literal`: 검색어 글자열 + 엄격한 앞뒤 경계)와 층 B(지금까지의 정규화 매처 `_single_pattern`).
  `build_matcher`는 층 A 또는 층 B로 맞으면 일치이고(층 A로 다 맞은 검색어 묶음을 먼저 고름), `literal_hit(text, terms)`가 층 A인지 알려 줌.
  숫자가 없는 검색어는 그대로.

### 기능
- `reader.py search`: 층 B 일치도 결과(`matches`)에 넣되 항목마다 `literal`(true/false, 숫자가 든 검색어일 때), 층 A를 먼저. 새 키 `literal_count`(층 A 수,
  `count`는 층 A+B). 층 B가 섞이면 `match_mode: exact_mixed`(새 값) + `UNVERIFIED: 글자 그대로 일치 N건, 표기가 다른 곳 M건(literal: false, 부호·단위를
  확인하세요)`. 모두 층 A면 지금처럼 `exact` VERIFIED. `--save`는 `exact_mixed`일 때 `표기` 열(.md는 제목 표시).
- 층 A 규칙: 띄어쓰기·탭 여러 칸 = 한 칸, 숫자와 단위 사이·`±` 앞뒤의 띄어쓰기 0~1칸 차이만 허용. 대소문자·전각/반각·부호 글자는 그대로.
  앞 경계(숫자, `.`, `,`, 영문, 대시·부호류, `/`, `^`가 아님, 띄어 쓴 `±`·`+`, `영하`, `(-)`도 아님, 빼기로 시작하는 검색어는 앞에 띄어 쓴 다른 값·`to`가 없음),
  뒤 경계(숫자, `.`·`,`+숫자, 영문, `²`, `³`, 위첨자, `^`, `/`, `%`가 아님, 단위 끝이면 `(rms)` 같은 괄호 꼬리, 띄운 `rms`, `p-p`, `peak`, `pk`, `abs`, `g`,
  `-`·`·`+영문도 아님). 기본 표기(`5 V`, `PLATE 10 mm`, `25 ℃`, `±0.5 mm`, `1,500 rpm`, `50 Hz`, `-40 ℃`, `10 MM`, `5V` = `5 V`)는 층 A.
- `excel.py check`: 숫자가 든 항목 글을 층 B로만 찾으면 `reason` 끝에 `표기가 다른 곳에서 찾음 (...)`, `match_mode` `exact_mixed`(판정은 그대로).
- `excel.py check --header-row N`(새 인자): 값 열 단위를 읽을 머리글 행 (기본 `--range` 바로 위 행).

### 수정
- [치명] 6·7차 검토의 오일치 쌍(review6 q1·q2, review7 q1~q6의 기대 0 전부, 264쌍)이 층 A(VERIFIED)로 나오지 않음 (회귀 `test_number_literal_layer_review7`,
  표 `tests/number_literal_cases.tsv`).
- [보통] 층 B에서 모두 영문 대문자로 쓴 단위(`10 MM`, `10UF`, `20MA`)는 접두 대소문자를 가리지 않음 (후보로만 나오던 회귀. `10UF` = `10 uF` = `10 μF`).
  소문자·섞인 단위(`10 mA`와 `10 MA`, `10 MΩ`과 `10 mΩ`)는 그대로 가림.
- [보통] 숫자 안전망의 20초 상한을 50단위마다가 아니라 단위마다 확인 (1MB 단위가 이어지면 상한을 크게 넘었음. 1MB 단위 55개에서 상한 2초·5초로
  재면 2.04초·5.13초에 멈춤).
- [보통] check 시간 제한으로 대조하지 못한 행: `reason` `대조하지 못함 (명령 시간 제한)`, `match_mode` null, 메모 파일에서 뺌, 남은 범위(대조하지 못한 첫 행~끝 행)만
  다시 대조하는 전체 check 명령을 `next_commands`에 넣고(머리글은 `--header-row`로 원래 위치) directive는 "사용자에게 확인을 받은 뒤 next_commands[N]을 그대로 실행".
- [경미] 규격 수치(`_NUM_RE`): 쉼표 뒤가 3자리 묶음이 아니면(`1,2345`, `1,5`, `10,20`) 수치로 보지 않음.

## [1.4.69] - 2026-11-04

독립 검토(6차) 수정. reader 숫자 정확 일치는 다섯 번의 검토에서 매번 부호·대시 경계 오일치가 나와, 경계를 더 땜질하지 않고
보수적 원칙으로 정리함: 정확 일치(exact, VERIFIED)는 모호하지 않을 때만, 애매하면 숫자 안전망 후보(number, UNVERIFIED).
회귀 테스트(검토 보고 q1·q2 전부, 6번 ± 빈틈, 반대로 맞아야 할 것)를 표로 먼저 만들어 재현한 뒤 고침.

### 구조
- 검색어 정규식의 숫자 앞 부호 판정을 고정 길이 뒤보기 묶음(`_NO_MINUS`, `_RANGE_BEFORE`, `_SIGN_PAT`)에서 일치 뒤 확인으로 바꿈:
  숫자 시작 자리에 빈 표시 그룹을 두고, 일치마다 글의 숫자 앞 부호(`_text_sign`: none / range / hyphen / minus / ambig / pm / plus)를
  검색어 숫자의 부호와 표(`_SIGN_OK`)로 비교해 맞지 않으면 버림(`_SignChecked`, re.Pattern의 search·finditer·pattern과 같게 씀).
  띄어쓰기 여러 칸·`/` 건너의 부호를 길이 제한 없이(12자까지) 보게 됨.
- 숫자 안전망을 하나의 큰 정규식에서 숫자별 `finditer` 위치 목록 + 40자 창 선형 탐색(`_LooseNumbers`)으로 바꿈.

### 수정
- [치명] 숫자 안전망 정규식(`_loose_number_pattern`)의 지수 백트래킹: `REGA R/W 0 0 0 ...` + `-q "0 0 0 0 0 0 0 1"`,
  `'1 '*600` + `"1 1 1 1 1 3"`, `'1 '*2500` + `"1 1 1 1 3"`이 끝나지 않았음. 선형 탐색으로 1초 안(회귀 `test_number_search_linear_time`).
  안전망 전체에 20초 상한(`NUMBER_SEARCH_SECONDS`)을 두고, 넘으면 남은 단위는 보지 않고 `not_searched`에 알려 0건이어도 VERIFIED로 단정하지 않음
  (`check`는 `판단필요`). search는 읽기 전용이라 저장할 중간 결과가 없어 CONTINUE 대신 이 상한을 씀.
- [치명] 보수적 원칙(부호·범위 대시): 글의 숫자 바로 앞(띄어쓰기 여러 칸·`/` 건너)이 `-`, U+2212, U+2013, `±`, `+/-`, `+ / -`, `+-`, `+` 중 하나인데
  검색어의 그 숫자에 같은 부호가 없거나, 검색어에 부호가 있는데 글이 다른 부호이거나 범위 대시인지 부호인지 애매하면 exact가 아님(숫자 안전망 후보).
  범위 대시는 양옆을 띄운 `0 - 20`만: 부호 없는 `20`과는 exact, `-20`과는 아님. `40 ℃`가 `Ta 25 -40 ℃`·`25 to -40 ℃`·`25,-40 ℃`·`Min-40℃`에,
  `12 V`가 `전원 +5 V -12 V`에, `+10 %`가 `+5 % -10 %`에, `0.5 mm`가 `±  0.5 mm`·`+ / - 0.5 mm`에, `-5 V`가 `+ / - 5 V`에, `08:30`이 `UTC-08:30`에 맞던 것
  모두 후보로. `-40 ℃`는 `25 to -40 ℃`·`25,-40 ℃`에 exact(1.4.68은 `to`를 단위로 보아 0건). 동작 변경: `20`과 `1-20`·`10 -20`, `+5V`와 `5V`,
  `20 ℃`와 `+20 ℃`도 이제 후보(exact 아님). 다른 낱말 뒤 숫자의 붙임표는 구분자(`SOI 1` = `SOI-1`, `ABC-1234`).
- [높음] 단위 접두 대소문자: 숫자 바로 뒤 `m`/`M`, `p`/`P`는 가림(`10 mohm`/`10 MΩ`, `10 mHz`/`10 MHz`, `1 mW`/`1 MW`, `5 mV`/`5 MV`, `10 MPa`/`10 mPa·s`가 맞았음).
  소문자 `mohm`은 mΩ로 보아 `MΩ`과 맞지 않음. `μ` = `u` = U+00B5. 접두가 아닌 단위(`min`, `max`, `ppm`, `psi`, `Pa`, `pH`)와 그 밖의 글자는 대소문자 무관(`5 v` = `5 V`, `10 KΩ` = `10 kΩ`).
- [높음] `spec_conditions` 긴 숫자열 제곱 시간: 수치 패턴이 숫자열의 모든 자리를 시작점으로 다시 읽었음(`"1"*5000` 78초). 숫자는 다른 숫자의 가운데서 시작하지 않고,
  구분자 없이 30자를 넘는 숫자열은 수치가 아님. `near_conditions`의 낱말 앞 절(`pre`)은 끝 200자만. `"1"*20000`, `시리얼 <숫자 5500자> 정격 전압`이 1초 안
  (회귀 `test_spec_conditions_long_digit_runs`, `test_check_long_serial_number_is_fast`).
  `check` 행 대조에 시간 확인(`rt.time_up`): 명령 시간 제한이 다가오면 남은 행은 `판단필요`(명령 시간 제한으로 대조하지 못함) + directive에 행 수와
  더 작은 `--range`로 다시 대조할지 물으라는 안내 (회귀 `test_check_time_limit_marks_rest_needs_review`).
- [중간] 숫자 안전망이 숫자 아닌 낱말을 보지 않았음(`1 퀀텀`이 아무 `1`에 후보). 숫자가 든 검색어의 숫자 아닌 낱말(단위 빼고, 조사 처리 포함)도 같은 단위 안에
  있어야 후보, 없으면 number를 건너뛰고 partial·fuzzy, 그래도 없으면 0건 VERIFIED (회귀 `test_number_safety_net_needs_other_words`).
- [낮음] 뜻이 다른 꼬리를 같은 단위로 보던 것: `Vp-p`, `Vpp`, `Vp`, `Apk`(피크), `Vrms`(실효), `barg`, `bara`, `MPaG`(게이지·절대), `dBA`(가중)는 exact 아님(후보).
  `1x10`이 `1×10^-3`·위첨자 지수에 맞던 것(숫자 뒤 지수면 다른 값), `φ`·`Φ`를 `Ø`와 같게 보던 것(상 `3φ`)도 exact 아님. `Vdc`, `VAC`, `sec`, `mins`, `hrs`는 그대로.
- [낮음] README `check` 설명의 찾기 순서와 판단필요 조건에 number 단계를 넣음.
- 회귀 표: `tests/test_reader.py` `NUMBER_KIND_CASES_6`(exact / 후보 / 없음). 1.4.67~1.4.68 표의 바뀐 기대값(`+20 ℃`, `1-20`, `+5V`/`5V`, `10 -20`, `Vrms`,
  `Vp-p`, `Vpp`, `dBA`, `barg`, `MPaG`, `Φ10`/`Ø10`)은 False로 고침.

## [1.4.68] - 2026-11-04

독립 검토(5차)에서 나온 reader 숫자 검색 경계 회귀와 `check` 안내 수정. 숫자 검색은 검토마다 경계 회귀가 나서 땜질에 더해 구조적 안전망을 둠.
회귀 테스트(보고서 사례 전부)를 먼저 만들어 재현한 뒤 고침.

### 기능
- 숫자 안전망 (`reader.py search`, `excel.py check`의 항목 찾기): 정확히 일치한 곳이 0건이고 검색어에 숫자가 있으면 숫자 값만 같은 곳을 느슨하게 다시 찾음
  (부호·단위 경계·붙여쓰기·기호 무시, 숫자 경계와 값은 지킴: `5`는 `15`·`5.5`와 다름, 숫자 사이 40자까지, 숫자가 없는 검색어는 정확히). 결과는 넓혀 찾기와 같은 형식
  (`count: 0`, `candidates`, `match_mode: number`, 항목마다 `number_words`(검색어: 앞 부호·뒤 단위를 포함한 실제 글)), directive
  `UNVERIFIED: 정확히 일치한 곳은 0건, 숫자가 같은 곳 N건 - 표기를 확인하세요 ...`. 경계 규칙이 너무 엄격해 맞는 표기를 놓쳐도 "0건 VERIFIED -> 규정 없음 단정"이 나지 않음:
  숫자가 든 검색어의 0건 VERIFIED는 느슨한 검색도 0건일 때만. 순서는 number -> partial -> fuzzy. `check`는 number면 `판단필요`(숫자 값만 같은 곳에서 찾음).
  회귀 테스트: `tests/test_reader.py` `test_number_safety_net_candidates`.

### 수정
- [치명] `±` 앞 글자 제한(`(?<![\w.])`) 때문에 `100±0.5 mm`, `Φ10±0.5mm`, `5V±5%`, `AC220V±10%`, `24VDC±10%`, `공차±0.1`, `R5±0.1`, `3 dB±1 dB`, `1000 ppm±2%`,
  `정격 전압: 5V±5%`가 0건 VERIFIED였음. `±`(`+/-`, `+-`)는 앞 글자와 무관하게 맞음. 검색어 쪽도 `5V±5%`처럼 붙여 쓴 `±`를 부호로 읽음
  (첫 숫자에 부호가 있으면 앞 경계는 부호 패턴이 봄). 글의 빼기 부호는 앞이 `[\dA-Za-z.]`가 아니면 부호(`온도-20℃`, `Ta=-20℃`), 부호 없는 숫자도 같은 규칙.
- [높음] (a) `+10%`가 `±10%`에, `+5V`가 `±5V`에 맞았음: 부호 없는/`+` 숫자 앞에 `±`, `+/-`, `+-`(띄어쓰기 하나 포함)가 오면 아님. `-5`도 `+/-5`와 다름.
  (b) `-20 ℃`가 `0 - 20 ℃`(범위)에 맞았음: 빼기 앞을 띄어쓰기 건너 보아 숫자 또는 숫자+단위(단위 글자 1~3개)면 범위의 대시라 부호가 아님
  (`0 - 20 ℃`, `0 -20 ℃`, `5 V - 20 ℃`). 부호 없는 숫자는 그 대시 뒤에서 맞음(`20` = `10 -20`, `100 Hz` = `10 -100 Hz`(U+2212 빼기 포함)). `0 ~ -20℃`는 그대로 -20.
- [중간] 끝 단위 뒤 경계가 같은 단위 표기까지 막았음(`5 V`가 `5 Vrms/Vp-p/Vpp/Vin/Vcc/Vmax`, `3 s`가 `3 sec`, `10 min`이 `10 mins/minutes`, `10 h`가 `10 hr/hrs`,
  `30 dB`가 `30 dBA`, `1 bar`가 `1 barg`에 0건). 단위별 허용 꼬리(`_UNIT_TAILS`: V `rms p-p pp p pk peak in out o cc dd ee ss max min typ nom ac dc`, A는 같은 목록에서
  `p o cc dd ee ss` 뺌, W `max min typ nom rms ac dc`, s `sec secs second seconds`, min `s ute utes`, h `r rs our ours`, dB `A`, 압력 `g a`)만 허용하고
  다른 단위로 이어지면(`10 mm`, `10 MΩ`, `1 Ah`, `30 dBm`, `30 dBi`, `1 kgf`, `100 gf`, `1 Wh`, `5 mmH2O`) 아님. `test_number_search_boundaries`의 `2 s`/`2 sec`·`2 seconds`는 맞음으로 바꿈.
- [중간] `excel.py check`의 UNVERIFIED 안내가 기본 문구("다시 실행하지 말고 그대로 알리세요")라 메모 미리보기 행동이 빠져 채점 O1(`CLI-dev/agent_scenarios.yaml`)과 어긋났음.
  directive를 "사용자가 메모를 원하면 next_commands[0](미리보기, 변경 없음)을 실행해 보여 주고, 사용자 확인 뒤에만 그 명령에서 --dry-run을 뺀 명령을 실행하세요.
  같은 check는 다시 실행하지 마세요"로. O1 요청("판정은 값 칸에 메모로 달아줘")과 판정(미리보기 `--dry-run` 필수, 실제 달기 금지)이 이 흐름과 맞아 yaml·팩은 그대로.
  규격서 등록이 남으면(`remaining`) 대조가 끝나지 않았으므로 `next_commands`는 같은 check 명령 하나뿐이고 directive는 "그대로 다시 실행해 이어서 대조"(CONTINUE 흐름,
  "다시 실행하지 말고"와 부딪히지 않음). 회귀 테스트: `tests/test_excel_read.py` `test_check_remaining_specs_reruns_same_command`.
- [낮음] `spec_norm`의 지수 상한이 3자리라 `1E999`, `1×10^999`가 1000자리 숫자로 바뀌어 조건(`999 Ω 이상` 포함)으로 읽혔음. 지수는 2자리까지 계산하고,
  3자리 이상은 `(지수 범위 밖)`으로 바꿔 조건으로 읽지 않음(check `판단필요`).
- [여유] 숫자 검색 표기 차이: `10 kΩ` = `10 kohm` = `10 kΩ`(U+2126), `100mm` = `100㎜`, `10 mA` = `10 ㎃`, `M6x1.0` = `M6×1.0`(숫자 사이 `x × *`), `Φ10` = `Ø10`(`Φ φ Ø ø`, U+2300),
  `08:30` = `8:30`(`18:30`, `8:30:15`와 다름), 영문으로 시작하고 `-`·`_`로 이은 코드(`ABC-1234`)는 `XABC-1234`, `ABC-1234-5`, `ABC-12345`와 맞지 않음.
- 회귀 테스트: `tests/test_reader.py` `test_number_search_boundaries_review5`(보고서 사례 전부, 93개), `test_spec_norm_exponent_over_two_digits_is_not_a_condition`.
  5차 재현 자료 p1~p6(248쌍)을 1.4.67과 비교해 새로 0건이 된 것은 모두 다른 값·부호·단위(`±5V`의 `5V`, 범위 `0 - 20 ℃`의 `-20 ℃`, `5V±5%`의 `5V 5%`)뿐.

## [1.4.67] - 2026-11-04

독립 검토(4차)에서 나온 office-cli 버그 수정과 `excel.py check`의 계약 변경. 회귀 테스트를 먼저 만들어 재현한 뒤 고침.

### 기능 (설계 변경)
- `excel.py check` 결과는 늘 `UNVERIFIED:`(대조할 항목이 없는 0건도). 이유: 판정은 자연어 규격 문장을 기계로 해석한 것이고, 1~4차 검토마다 새로운 해석 오판정
  (유니코드 빼기, 지수 표기, 최대·최소 낱말, 단서 조항, 표 평문 행 등)이 나왔음. 규칙을 더해도 휴리스틱이라 모든 문장을 맞힌다고 보장할 수 없으므로
  "도구가 결과를 확인함(VERIFIED)"이라고 하지 않고 계약을 정직하게 바꿈. `unverified[0]`(`CHECK_ADVISORY`)은 "자동 판정은 참고용입니다 ... 행마다 읽은 조건(condition)과
  근거 문장(spec_text)을 사용자에게 보여 주고 확인을 받은 뒤 메모를 다세요". 행 판정 값(일치/불일치/판단필요/근거없음)과 보수적 판정 원칙은 그대로.
  메모 CSV(`memo_file`)와 `next_commands[0]`로 다는 메모는 `자동 판정(참고): <판정> (이유) - 읽은 조건: <condition> - 근거: <장절>: <규격 문장 앞 60자>`.
  office-spec 스킬·전역 지침(`CLI-dev/opencode/AGENTS.md` 4장)에 "판정은 참고용, 조건을 보여 주고 사용자 확인 뒤 메모" 반영.

### 수정
- [높음] `excel.py check` 해석 버그 (말뭉치 `tests/check_corpus.tsv`에 4차 묶음 129건 추가, 모두 261건: 일치 110, 불일치 66, 판단필요 83, 근거없음 2):
  - C1 유니코드 빼기(U+2212, U+2010~2012, 앞이 숫자·단위가 아닌 숫자 앞 U+2013)를 `-`로(`spec_norm`: 규격 글, 칸 값, 다음 항목 찾기). `-20 ~ 60 ℃`(U+2212 빼기)가 `20~60 °C`로 읽혀 -10 ℃를 불일치였음.
  - C2 `±N 이하/이내`는 -N~N(`±5 % 이내`가 `5 % 이하`라 -6 %를 일치였음).
  - C3 ohm 계열 대소문자 구분(`mOhm` = mΩ, `Mohm`·`MOhm` = MΩ). 소문자 `mohm`은 밀리인지 메가인지 몰라 단위 모름 -> 판단필요(대소문자 무시로 MΩ이 되어 200 mΩ을 일치였음).
  - C4 지수 표기 `1×10`+위첨자 9, `1×10^9`, `1E9`, `10`+위첨자 6은 계산(NFKC가 위첨자를 `109`로 바꾸기 전에). 묶음 단위(`Ω·cm`, `V/m`, `mm²`)는 그 단위로 읽지 않음.
  - C5 단위 뒤 괄호 꼬리(`dB(A)`, `mV(p-p)`, `V(DC)`) 뒤의 이하·이상을 읽음.
  - C6 기준 값만 있는데 항목·절에 `최소 최대 최고 최저 상한 하한 Max Min Typ 약 내외 정도 approx` 등이 있으면 판단필요. 상황마다 다른 조건(`대기 시 ..., 동작 시 ...`,
    `초과 시 동작`), 단서 조항(`다만,`, `(단, ...)`, 다음 문장이 `다만`)도 판단필요.
  - C7 단위 있는 조건 곁의 단위 없는 수치(표 평문 `출력 전압 4.75 5.0 5.25 V`, `외경 100 mm, ±0.5`, `100 mm ±0.5`), 비대칭 공차(`+0.5/-0.2`, `(+10 %, -15 %)`),
    치수 나열(`100×200×50 mm`, `W 400 × H 300`), 소수점 쉼표일 수 있는 수치(`1,5 mm`) -> 판단필요.
  - C8 (c) 합치기는 범위 앞에 한계 낱말이 있거나 기준 값 바로 뒤 괄호 안에 범위만 있을 때만. 범위 뒤에 `까지 측정 가능`·`입력 가능` 등 쓰임새가 이어지면 합치지 않음.
    `A에서 B로(를)`(`220 V에서 110 V로 강압`)는 범위가 아님(`A에서 B까지`, 뒤에 `로`가 없는 `A에서 B`는 범위). `DC 24 V(허용 범위 DC 20.4 V ~ 26.4 V)`의 `DC`는 다른 항목 낱말이 아님.
  - C9 부정문(`되어서는 안 된다`, `아니 된다`, `이어서는 안`) -> 판단필요.
  - 4차 자료 c2의 `소음 45 dB(A) 이하` / 44 dB는 판단필요, c3의 같은 문장 / 40 dB는 일치로 기대가 엇갈려 c3(일치)을 따름.
- [치명] reader 숫자 검색의 오일치가 exact VERIFIED였음: 검색어의 부호를 버려 `-20 ℃`가 `20 ℃`·`+20 ℃`·`1-20 ℃`에, `-5 V`가 `+5 V`에 맞았고, 단위 뒤 경계가 없어
  `10 m`가 `10 mm`·`10 mA`·`10 MΩ`에, `5 V`가 `5 VA`에, `5 A`가 `5 Ah`에, `100 W`가 `100 Wh`에 맞았고, 기호 분기에 소수 경계가 없어 `±5`가 `±5.5`에, `1/2`가 `1/2.5`에 맞았음.
  부호(`-`, `+`, `±`, `+/-`, U+2212)는 패턴에 넣고 앞쪽 경계를 지킴(글의 U+2212·U+2010~2013도 `-`). 부호 없는 숫자도 음수(`-20`)의 끝과는 맞지 않음(`1-20`, `A-20`은 맞음).
  끝이 단위면 뒤 경계 `(?![A-Za-z...])`(Ω·μ·℃ 등 비ASCII 포함, `Vdc`·`VAC`는 맞음). 기호 분기 소수 경계 `(?![.,]?\d)`.
  맞아야 하는 것: `10%` = `10.0 %`, `0.5` = `0.500`(끝 0 둘 이상, 끝 0 하나는 장절 번호일 수 있어 단위가 붙을 때만), 날짜 모양 `2026.11.04` = `2026.11.4` = `2026-11-04` = `2026/11/04`
  = `2026년 11월 4일`, 천 단위 쉼표는 3자리 묶음일 때만(`1,500`이 `15,00`, `2,1,500`에 맞지 않음).
  회귀 테스트: `tests/test_reader.py` `test_number_search_boundaries`(80개). `test_search_recall_variants`의 `20°C` 사례는 `-20°C`로 바꿈(°C = ℃ 확인은 그대로).
- [중간] outlook 짧은 번호 잠금 경합: 남은 잠금 치우기(`_take_stale`)가 다른 실행이 막 만든 잠금을 이름 바꿔 가져가고 되돌리지 못해 두 실행이 함께 잠금을 쥐어
  번호가 겹쳤음(8프로세스 21회 중 2회). 잠금 파일을 연 채 운영체제 잠금(Windows `msvcrt.locking`, 그 밖 `fcntl.flock`)을 걸고, 프로세스가 끝나면(강제 종료 포함) 운영체제가 풂.
  pid 판정·나이로 치우기·이름 바꾸기를 모두 없앰(`LOCK_STALE`, `LOCK_STALE_ALIVE`, `_pid_alive` 삭제). 잠금 파일은 지우지 않고 남김.
  못 잡은 `warnings`: `다른 실행이 번호 파일을 쓰는 중 (5초 기다림)` / `잠금 파일 outlook_ids.lock을 열지 못함 (...)`.
  `_backup_broken`의 `_read_next_file()`에 `deadline`을 넘김(잠금을 쥔 동안 재시도 상한).
  회귀 테스트: `tests/test_outlook.py` `test_short_ids_many_processes_never_share_numbers`(8프로세스 4회), `test_short_ids_lock_held_by_other_process_then_killed`,
  `test_short_ids_leftover_lock_file_does_not_block`.
- [낮음] `spec_text`·`condition`·`evidence`에 실리는 규격 원문 인용의 CP949 밖 글자: 도구가 만든 문구가 아니라 사용자 문서의 인용이라 그대로 둠
  (AGENTS 문자 규칙은 도구가 만드는 문구에 적용). README `check`에 한 줄.

## [1.4.66] - 2026-11-04

독립 검토(3차)에서 나온 1.4.65 회귀와 남은 버그 수정. `excel.py check`는 검토마다 판정 회귀가 나서 땜질 대신 보수적 판정 원칙으로 정리하고
표준 판정 말뭉치로 회귀를 막음.

### 기능
- `excel.py check` 보수적 판정: 일치·불일치는 항목 절(항목 낱말부터 다음 항목 직전)의 조건이 모호하지 않을 때만. 조건을 하나로 합치는 규칙은 셋뿐:
  (a) 같은 단위 아래·위 한계 -> 범위, (b) 기준 값 바로 뒤의 `±N 단위`·`±N %`(같은 절이거나 `허용 오차`·`허용차`·`공차`·`편차`·`허용 범위` 절) -> 기준 값 ± N,
  (c) 기준 값 바로 뒤의 양쪽 한계 범위가 기준 값을 품으면 그 범위. 그 밖에 조건이 둘 이상 남으면 `판단필요`(`reason`·`condition`에 후보 조건).
  한쪽 한계만 있는 허용 절(`허용 범위 4.75 V 이상`, `허용 150 mA 이하`, `허용치 10 A 이하`, `오차 0.1 mm 이내`)과 기준 값, `±` 없는 % 조건(`허용 오차 5 %`),
  기준 값을 품지 않는 범위, 기준 값 없는 `±` 공차, 불합격 기준 문장(`0.5 mA 초과 시 불합격`), 값이 여럿인 문장(`50/60 Hz`)은 `판단필요`.
  기준 값을 범위 안이면 빼던 1.4.57 규칙(한쪽 한계 포함)은 (c)로 대신함. 약어 점(`Max. 3A`, `min. 100 MΩ`)에서 문장을 나누지 않음.
- 표준 판정 말뭉치 `tests/check_corpus.tsv` 132건(일치 71, 불일치 36, 판단필요 23, 근거없음 2): B1 33문장, 1.4.63 두께/길이·외경/폭, 1차·2차·3차 검토 사례 전부.
  기대 판정은 "틀린 VERIFIED 금지"가 최우선(모호하면 판단필요). `test_check_corpus`가 행마다 `_judge_row`로 판정.

### 수정
- [치명] `excel.py check` 회귀(1.4.65): `사용`·`동작`을 한계 낱말에 넣어 "시험 전압 500 V, 사용 범위 0~1000 V"의 시험 전압 조건이 `0~1000 V`가 되어 800 V를 일치.
  `사용 범위`·`동작 범위`·`측정 범위`는 별개 항목(한계 낱말에서 뺌). 뒤 항목의 범위로 판단필요를 거는 안전장치도 이름에 `범위`가 든 뒤 항목은 보지 않음
  (그래서 800 V는 불일치, 500 V는 일치). 회귀 테스트: `check_corpus.tsv` 3차 묶음, `test_near_conditions_stop_at_next_item`.
- [치명] 숫자 검색어 `2.0`이 `2.5 V`, `1.0`이 `1.03 mm`, `0.0`이 `0.5`에 맞았음(정수 뒤 `(?:\.0+)?`를 건너뛰고 경계가 소수점을 허용).
  소수 끝 0 무시는 글의 숫자가 수학적으로 같은 값일 때만(`2.0` = `2`, `2.00`; `2.0` != `2.5`, `2.05`, `2.0.3`). 끝 0을 적은 검색어 `2.10`은 글의 `2.10`과 단위 붙은 `2.1 V`에 맞음.
  회귀 테스트: `test_number_only_query_normalized`(표 20개 추가).
- [높음] 조사 `로` 이중 검색이 뗀 낱말 앞부분에 맞았음(`전기회로` -> `전기회사`, `냉각수로` -> `냉각수 온도`, `보호회로` -> `보호회의`, `고속도로` -> `고속도`).
  받침 없는 글자 뒤 `로`는 떼지 않고, ㄹ 받침 뒤 `로`만 뗌(`케이블로` -> `케이블`). 뗀 낱말 뒤에는 한글이 이어지지 않아야 맞음(조사는 허용: `케이블은`은 맞고 `케이블카`는 아님).
  회귀 테스트: `test_particle_ro`.
- [낮음] 전각 단위 글자(`㎜ ㎝ ㎏ ㎖ ㎃` 등 U+3380~U+33DF, `ℓ`)를 숫자 끝 0 판정의 단위로 봄(`1.5`가 `1.50㎜`에 맞음).
- [낮음] 메일 번호 잠금: 잠금을 쥔 실행이 `retry_fs`로 30초 넘게 기다려 다른 실행이 쓰는 중인 잠금을 지울 수 있었음. 잠금을 쥔 동안 재시도는 합쳐 10초(`LOCK_HOLD`) 안에서
  끝냄(`rt.retry_fs`·`safe_write_text`의 `deadline`). 잠금 파일에 `<pid> <표식>`을 쓰고, 남은 잠금은 만든 실행이 끝났으면 바로(Windows `OpenProcess`·`GetExitCodeProcess`),
  살았는지 모르면 30초, 살아 있어도 120초가 지나면 지움(이름을 바꿔 가져와 내용이 같을 때만, 다르면 되돌림). 풀 때는 자기 표식일 때만 지움.
  못 잡은 `warnings` 문구를 까닭에 맞춤(`다른 실행(pid N)이 번호 파일을 쓰는 중` / `잠금 파일 outlook_ids.lock이 남아 있음(...)`).
  회귀 테스트: `tests/test_outlook.py` `test_short_ids_lock_of_dead_process_is_cleared_at_once`, `test_short_ids_lock_of_live_process_is_kept`,
  `test_short_ids_lock_hold_time_is_bounded`, `test_retry_fs_deadline_stops_waiting`, `test_short_ids_lock_timeout_and_stale_lock`.
- [낮음~중간] 숨은 Word·PowerPoint 종료: `Documents.Count`가 예외(COM 바쁨)면 끄지 않고 남겼고, 자기 문서 닫기가 실패해 남아도 끄지 않았음.
  남은 문서 수는 0.3초 간격으로 5번까지 다시 읽고, 남은 문서가 자기 문서뿐이면 끔. 끝내 못 읽으면 끄지 않고 stderr `[INFO]`.
  회귀 테스트: `tests/test_office.py` `test_word_hidden_instance_quits_only_when_no_documents_left`(4개 추가).

## [1.4.65] - 2026-11-04

독립 검토(2차 재검토)에서 나온 office-cli 버그 수정. 각 항목은 회귀 테스트로 먼저 재현한 뒤 고침.

### 수정
- [치명] outlook 짧은 `id` 재사용: 번호 기억 파일을 읽지 못하면(잠김·손상) `{"next": 1}`로 덮어써 번호를 1부터 다시 매겨 옛 `--id 3`이 다른 메일을 가리켰고,
  동시에 실행한 두 목록이 같은 번호를 받았음. 이제 읽기는 `retry_fs`로 재시도하고, 파일이 있는데 끝내 못 읽거나 깨졌으면 번호를 새로 매기지 않고 그 결과는
  `id` = 긴 EntryID + `warnings`, 짧은 번호 해석은 STOP_RETRY "목록을 다시 보세요"(+ `next_commands` list). 깨진 파일은 `outlook_ids.broken_<시각>.json`으로 옮겨 보존.
  다음 번호를 작은 파일 `outlook_ids.next`에도 두고 큰 값을 써서 줄어들지 않음. 9999에서 1로 돌아가던 것도 없앰(번호는 9자리까지).
  불러오기~저장은 잠금 파일 `outlook_ids.lock`(O_EXCL, 5초 대기 상한, 30초 넘은 잠금은 지움)으로 묶음, 못 잡으면 긴 EntryID + 경고.
  회귀 테스트: `tests/test_outlook.py` `test_short_ids_corrupt_file_never_restarts_numbering`, `test_short_ids_read_failure_keeps_file_and_numbering`,
  `test_short_ids_concurrent_listings_get_distinct_numbers`, `test_short_ids_corrupt_file_and_next_file_keeps_largest`, `test_short_ids_lock_timeout_and_stale_lock`.
- [치명] `excel.py check` 회귀(1.4.63): "정격 전압은 5 V이며, 허용 범위는 4.75 V 이상 5.25 V 이하로 한다."에서 조건이 `5 V`만 남아 5.1 V를 불일치 VERIFIED.
  허용·변동·범위·공차·오차·편차로 시작하는 절은 새 항목이 아니라 앞 항목의 한계로 봄(`_NEXT_ITEM`, 기준 값 지우기의 항목 낱말 판정). `허용 전류`처럼 뒤에 다른 낱말이
  오면 새 항목. 안전장치: 잘라 낸 쪽에 기준 값만 남고 같은 문장 뒤에 같은 단위의 범위가 그 값을 품으면, 칸 값이 기준 값과 다르지만 그 범위 안일 때 `판단필요`.
  회귀 테스트: `tests/test_reader.py` `test_near_conditions_stop_at_next_item`(6개 추가), `test_near_conditions_plain_with_later_range_is_marked`,
  `tests/test_excel_read.py` `test_check_limit_clause_belongs_to_item`.
- [높음] 닫힌 Word 파일 작업 뒤 숨은 Word를 `Quit(SaveChanges=0)`로 무조건 끝내 사용자가 그 인스턴스에서 연 문서의 변경을 버렸음. 자기 문서를 닫은 뒤
  `Documents.Count == 0`일 때만 끝내고(PowerPoint와 같음), 남겨 두면 stderr `[INFO]`. 회귀 테스트: `tests/test_office.py` `test_word_hidden_instance_quits_only_when_no_documents_left`.
- [중간] 열기 감시: 사용자가 이미 띄운 PowerPoint의 `DisplayAlerts`를 바꾼 채 `_abandon`(os._exit)하면 되돌리지 못했음. 이미 떠 있던 PowerPoint는 바꾸지 않음.
  끄지 않고 둔 프로세스(`POWERPNT.EXE PID N` 또는 "이미 떠 있던 것")를 message에 적음. 회귀 테스트: `tests/test_office.py` `test_open_watchdog_leaves_running_ppt_alerts`.
- [중간] 2열 표 파일 셋째 열 오류의 `[추천] C7,"1,500"`이 인자 값처럼 읽혔음. `[추천]`을 쓰지 않고 "고친 줄: 3번째 줄 C7,"1,500"" (한 줄이어도).
  회귀 테스트: `tests/test_excel_write.py` `test_write_cells_file_third_column_stops_before_writing`.
- [중간] 숫자 검색어: `2.1`이 `2.10 전압`에 맞고 `2.1.3`에는 안 맞았음. 소수 한 자리 맨 숫자는 장절 번호 경계(`section_matches`와 같음: `2.1.3`은 맞고 `2.10`, `12.1`은 아님),
  소수 끝 0은 소수 둘째 자리 이상이거나 글에서 단위가 붙을 때만 무시(`0.5` = `0.50 mm`, `0.5` != `0.50 이다`). 회귀 테스트: `test_number_only_query_normalized`(8개 추가).
- [낮음] 넓혀 찾은 검색의 directive가 stale·indexed_now·상한·text 잘림 안내를 지웠음 -> 덧붙임(stale이면 "다시 실행하지 마세요"를 빼고 그 안내를 따름).
  조사 `로`: ㄹ 받침 뒤는 뗌(`케이블로` -> `케이블`), 받침 없는 글자 뒤는 원래 낱말과 뗀 낱말을 둘 다 찾음(`서버로`, `보호회로`는 그대로 유지).
  회귀 테스트: `tests/test_reader.py` `test_search_widened_directive_keeps_notes`, `test_particle_ro`.
- office-mail 스킬: 규칙 4를 사실에 맞춤(옛 번호는 같은 메일이거나 "없는 번호", 긴 EntryID `id`도 그대로), `--send` 순서(먼저 `--send` 없이 또는 `read --id N` ->
  `original`·받는 사람을 사용자에게 확인 -> 확인 뒤에만 `--send`).

## [1.4.64] - 2026-11-04

독립 검토(2차)에서 나온 버그 수정. 각 항목은 회귀 테스트로 먼저 재현한 뒤 고침.

### 수정
- [치명] outlook 짧은 `id`: 마지막 목록 하나만 기억해 `search` 뒤 `thread`(또는 다른 `list`)를 하면 같은 `--id 3`이 다른 메일을 가리켰고,
  `reply/forward --id 3 --send`가 엉뚱한 메일로 나갈 수 있었음. 번호를 목록마다 1부터 다시 매기지 않고 이어서 늘어나게 주고(이미 번호가 있는 메일은 같은 번호),
  번호 -> EntryID·제목·보낸사람·받은 시각을 최근 500개(7일)까지 `<데이터 폴더>/outlook_ids.json`(옛 `outlook_last_list.json` 대신)에 기억. 옛 번호는 같은 메일이거나
  "없는 번호" STOP_RETRY(+ `next_commands` list). 9999를 넘으면 1부터 다시 매기며 기억한 번호를 모두 버림.
  `reply`/`forward` 결과에 `original`(subject·from·received), `attachments`의 `mail`에 from·received를 더하고 preview·directive가 어느 메일인지 알리게 함.
  회귀 테스트: `tests/test_outlook.py` `test_short_ids_never_point_to_another_mail_after_new_listing` 외 3개.
- office-mail 스킬: `reply`·`forward` 행에 `--id 3`, `--send`는 사용자가 받는 사람·원본 메일을 확인한 뒤에만.
- 스킬 4개에 옛 통합 SKILL의 안전 규칙 보충: 새 파일은 사용자가 준 전체 경로로만, 같은 이름이 있으면 덮어쓰지 않고 새 이름으로 저장되니 `output_file`·`recovered`를 알림,
  office-spec의 메모 `--overwrite`는 사용자가 요청했을 때만 (각 6KB 이하).
- [치명] reader 등록·검색: 읽기에 실패하는 느린 파일이 있으면 매 실행 같은 파일을 다시 읽고 다음 파일 앞에서 `CONTINUE:`가 나 끝나지 않았음
  (`_index_files`, 검색 때 원본이 바뀐 캐시 다시 읽기). 실패를 (경로, 수정 시각, 크기, 파서 버전)으로 `<캐시 폴더>/index_failures.json`에 기록해 다음 실행에서
  건너뛰고 `failures`/`skipped`/새 키 `refresh_failures`로 보고. 파일 하나만 주면 바로 다시 시도, `clear --all`은 기록도 지움. `progress`의 등록 수는 실제 등록 수.
  회귀 테스트: `tests/test_reader.py` `test_failing_file_is_remembered_so_continue_ends`(R10 실행마다 새 단계 1개), `test_stale_refresh_failure_is_remembered`.
- [중요] `--section "2.1"`이 `2.10`, `12.1`에도 맞던 문제 (`section_matches`, read/search/diff/excel check 공통). 장절 번호로 보이는 값은 번호 경계를 지키고
  하위 절(`2.1.3`)과 `제2장`은 포함, 제목 글자는 지금처럼 부분 일치. 회귀 테스트: `tests/test_reader.py` `test_section_number_keeps_boundaries`.
- [중요] `search` 넓혀 찾기: "대기 전류"가 "대기 전력"에 `VERIFIED` count 2로 맞았음. 한글 fuzzy는 4음절 이상, 첫·끝 음절이 같고 바뀐 음절은 자모 하나만 다를 때만
  (전압/전류, 최대/최소, 상한/하한이 서로 맞지 않음). 넓혀 찾은 결과는 `count: 0`·`total_matches: 0`, 목록은 `candidates`(`matches` 아님)와 `candidate_count`,
  directive `UNVERIFIED: 정확히 일치한 곳은 0건, 비슷한 곳 N건 ...`(옛 `recovered` "넓혀 찾음"은 없앰). `matched_terms`는 글자 그대로 맞은 검색어만.
  `excel.py check`의 `판단필요`는 그대로. office-spec 스킬과 전역 지침 `CLI-dev/opencode/AGENTS.md`의 match_mode 문구를 맞춤.
  회귀 테스트: `tests/test_reader.py` `test_search_fuzzy_daegi_current_is_not_power`, `test_fuzzy_never_matches_different_korean_terms`.
- 경미: 숫자만 든 검색어 정규화(`5,000` = `5000`, `0.5` = `0.50`, 숫자 경계는 지킴). 검색어 낱말 끝 조사는 받침이 맞을 때만 떼고 `로`는 떼지 않음
  (`와이파이` -> `와이파`, `보호회로` -> `보호회`로 떼던 문제). 캐시의 수식 칸 다시 확인(`recheck_uncomputed`)은 한 실행에 3개·시간 여유 안에서만.
  `diff --save x.md`가 말없이 `.xlsx`로 저장하던 것을 `warnings`로 알림. 회귀 테스트: `tests/test_reader.py` `test_number_only_query_normalized`,
  `test_word_final_syllable_is_not_particle`, `test_recheck_uncomputed_is_capped_per_run`, `test_diff_save_md_warns_xlsx`.
- 경미(outlook): `calendar`에서 Outlook이 기간 조건을 거부하면 일정마다 시각을 비교(`recovered`, 5,000개 상한), 읽지 못한 일정이 있으면 0건이어도
  `UNVERIFIED:`(`unread_events`, 예외를 삼켜 0건 VERIFIED이던 문제). `thread` 대체 경로가 폴더마다 2,000통 상한에 닿으면 `UNVERIFIED:`.
  회귀 테스트: `tests/test_outlook.py` `test_calendar_restrict_refused_compares_each_event`, `test_calendar_unreadable_events_are_not_verified_zero`,
  `test_thread_fallback_scan_limit_is_unverified`.

### 개발 도구
- Gemma 종합 테스트 채점(`CLI-dev/grade_opencode.py`, `agent_scenarios.yaml` 묶음 O): 채점기가 `;`·`&&`·`||`·`|`로 이은 명령을 나눠(따옴표 안은 그대로)
  `commands`·`forbid`를 명령마다 봄 (O1 금지 규칙 `(?![^\n]*--dry-run)`이 이어 붙인 다른 명령의 `--dry-run`을 보고 실제 메모 달기를 놓치던 문제).
  O1: 메모 글(`--text`) 안의 "--dry-run" 글자는 옵션으로 보지 않음(따옴표 밖만). O4: `--anchor`(문서 중간 삽입) 금지, 문서는 `--target`으로.
  O5: 채운 문서가 `--target`의 성적서양식.docx여야 함(값 파일 이름에 든 양식 이름으로 통과하던 문제). 회귀 사례: `CLI-dev/tests/test_grade_opencode.py`.
- `CLI-dev/apply_skills.py`: 옛 `skills\office-cli`가 링크·정션이면 링크만 지우고 가리키던 폴더는 그대로 둠(보관도 하지 않음, 안내에 대상 경로).
  전에는 링크를 따라 복사한 뒤 지우다 실패하고 "OpenCode를 닫고 다시 실행"이라고 잘못 안내함. 잠김(PermissionError)이 아닌 실패는 다른 안내.
  회귀 테스트: `CLI-dev/tests/test_apply_skills.py` `test_old_office_cli_skill_link_removes_only_link`.

## [1.4.63] - 2026-11-04

독립 검토에서 나온 버그 수정. 회귀 테스트는 항목마다 고치기 전 코드에서 실패를 확인함.

### 수정
- 2열 표 파일의 셋째 열 (`excel.py write --cells 칸값.csv`, `excel.py comment --file`, `word.py fill --file`): 값의 쉼표가 열을 나누면
  (`C7,1,500`, `B5,규격 3.1과 다름, 재확인 필요`) 둘째 열만 쓰고 `VERIFIED`였음. 셋째 열 이후에 글이 있으면 아무것도 쓰기 전에 STOP_RETRY:
  줄 번호와 "값을 큰따옴표로 감싸세요", 한 줄이면 고친 줄(`C7,"1,500"`)을 `[추천]`(오류 로그에는 줄 번호만). 공통 함수 `_runtime.check_two_columns`.
  회귀 테스트: `test_excel_write.py` `test_write_cells_file_third_column_*`, `test_comment_file_third_column_*`, `test_word.py` `test_fill_third_column_*`.
- `write --cells` 항목 나누기: 닫히지 않은 따옴표가 뒤 항목을 삼킴(`C6='007; C7=5` -> 한 항목). 따옴표 묶음은 같은 따옴표로 닫히고 바로 뒤가
  `;`/끝일 때만 인정, 아니면 따옴표 없는 값으로 나눔(`'007`은 엑셀 글자 표시로 글자 007). `=`로 시작하는 값은 큰따옴표·괄호 안의 `;`를 나누지 않음
  (`=IF(A1="a;b",1,2)`, `=SUM(A1;A2)`는 엑셀이 받지 않으면 `,`로 고친 `[추천]`까지). 회귀 테스트: `test_write_cells_quote_must_close_*`, `test_write_cells_formula_keeps_*`.
- `excel.py comment --append` 재실행: 같은 메모가 두 번 붙음. 기존 메모 끝 줄이 붙일 메모와 같거나 기존 메모가 새 메모와 같으면 그 칸은 그대로
  (`unchanged`, 모두 그대로면 되돌리기 지점 없음). 회귀 테스트: `test_comment_append_rerun_does_not_duplicate`.
- 엑셀 쓰기 실패 후 되돌리기(`apply_writes`): `rng.Formula = prior`가 글자 `007`, `1-2`, 긴 숫자 글자, `'=abc`를 숫자·날짜·수식으로 바꾸고도
  되돌린 것으로 보아 되돌리기 지점까지 지웠음. 수식 칸(`HasFormula`)만 수식 글, 상수 칸은 `Value2`(글자는 앞에 `'`)로 저장·복원하고, 복원 뒤 다시 읽어
  비교해 다르면 그 칸을 message에 적고 되돌리기 지점을 남김. 가짜 Excel에 `typing`(대입한 글자를 Excel처럼 해석), `Value2`, `PrefixCharacter`.
  회귀 테스트: `test_rejected_formula_rollback_keeps_text_constants`, `test_rejected_formula_rollback_mismatch_keeps_undo_point`.
- `excel.py check` 규격 조건이 같은 문장의 다른 항목 조건을 섞음: `두께 5 mm 이상, 길이 100 mm 이하로 한다.` -> 두께 `5~100 mm`,
  `외경 3 mm, 폭 5 mm 이하로 한다.` -> 외경 `5 mm 이하`인데 `VERIFIED`. `reader.near_conditions`는 낱말부터 다음 항목(절 경계 뒤 낱말 + 수치) 직전까지만,
  없으면 낱말이 든 절에서 찾음. `reader.spec_conditions`는 두 조건 사이에 다른 항목 낱말이 있으면 한계를 합치지 않고 기준 값(plain)도 지우지 않음.
  기존 규격 문장 33가지(B1)는 그대로. 회귀 테스트: `test_reader.py` `test_spec_conditions_do_not_mix_items`, `test_near_conditions_stop_at_next_item`,
  `test_excel_read.py` `test_check_does_not_mix_items_of_one_sentence`.
- `excel.py check` 칸·머리글·항목에 단위가 없으면 규격 단위로 추정해 판정하던 것을 `판단필요`(`칸 단위 없음 (규격은 cm): 값 열 머리글에 단위를 적으세요`)로.
  머리글 단위가 규격과 같은 종류면 접두어를 맞춰 판정(그대로). 회귀 테스트: `test_check_cell_without_unit_is_not_guessed`.
- `word.py fill`: 표 칸 `표2!R3C4`가 병합 표에서 다른 칸일 수 있는데 미리보기로 알아볼 수 없었음. `planned`·`applied`·덮어쓰기 확인 message에
  그 행 첫 칸과 그 열 첫 행 글(12자)을 붙임(`표2!R3C3 [행 '전류', 열 '측정']`). 회귀 테스트: `test_fill_preview_shows_row_and_column_heads`.
- `word.py fill`: 문서 보호(`ProtectionType`)를 쓰기 전에 확인해 STOP_RETRY, 쓰는 중 실패하면 "앞의 N곳은 이미 바뀌었습니다(office.py undo)"
  (닫힌 파일은 저장하지 않아 그대로). 회귀 테스트: `test_fill_protected_document_changes_nothing`, `test_fill_failure_midway_says_how_many_changed`.
- 닫힌 파일 열기 감시: PowerPoint는 한 프로세스라 감시 중에 사용자가 연 발표 파일까지 같이 끝낼 수 있었음. Word·PowerPoint는 끝내기 직전에
  그 프로세스에 보이는 문서 창이 없다고 창 목록(COM 없이)으로 확인될 때만 끝냄. 사용자 문서 창이 있거나 확인하지 못하거나, 끝낼 프로세스를 모르면
  (사용자가 이미 띄운 PowerPoint, 번호를 못 찾음: 전에는 감시하지 않아 멈춤) 끝내지 않고 `drm_or_dialog` 오류 1줄을 낸 뒤 도구만 끝냄.
  시간 초과 message에 "열기가 오래 걸린 경우일 수도 있습니다(OFFICE_CLI_OPEN_TIMEOUT으로 늘릴 수 있음)".
  회귀 테스트: `test_office.py` `test_open_watchdog_kills_only_without_user_documents`, `test_open_watchdog_abandon_emits_one_json_line`.

## [1.4.62] - 2026-11-04

작업별 스킬로 나누기 (C3). 통합 `SKILL.md`(16KB, 명령 표 한 줄 300~600자)에서 Gemma 26B가 옵션을 잘못 기억하던 문제.
회귀 테스트: `tests/test_skill_docs.py`(배치·크기·frontmatter·전체 명령·모든 하위 명령 포함), `CLI-dev/tests/test_apply_skills.py`
(옛 스킬 폴더 보관 후 지움, diag FAIL), `CLI-dev/tests/test_grade_opencode.py`(새 스킬 경로 읽기는 코드 읽기 아님).

### 구조
- `CLI/office-cli/SKILL.md`를 지우고 `CLI/office-cli/skills/<이름>/SKILL.md` 4개로 나눔 (각 6KB 이하, 명령은 `python "C:\CLI\office-cli\<파일>.py" ...` 전체 명령):
  `office-excel`(excel.py, 엑셀 차트·범위 transfer), `office-doc`(word.py, ppt.py, office.py transfer·undo·windows·focus),
  `office-mail`(outlook.py), `office-spec`(reader.py 읽기·검색·diff·캐시, excel.py check·comment).
  undo·windows는 excel·doc에, excel comment는 excel·spec에 함께 둠. 공통 결과 처리(첫 단어별 행동)는 전역 지침 4장, 스킬에는 작업별 추가 규칙만.
- 배치 이유: 스킬 폴더 이름 = `name`이어야 해서(OpenCode) 도구 폴더 바로 아래에는 스킬 하나만 둘 수 있음. `skills/` 아래 두면 `C:\CLI\office-cli`를
  그대로 복사하는 배포가 그대로이고, 도구 VERSION과 스킬 버전이 한 폴더에서 함께 바뀜. 보관된 En-Genius 패키지(`archive/engenius`)는 옛 경로를 읽지만
  업데이트 중지 상태라 고치지 않음(다시 꺼낼 때 새 경로로).
- `CLI-dev/apply_skills.py`: 4개 스킬을 `skills\<이름>\`에 설치하고, 옛 `skills\office-cli` 폴더는 `work\skill_backup_<시각>`에 보관한 뒤 지움
  (남으면 OpenCode가 옛 스킬도 고름). `--check`와 `run_all_tests.py` diag는 옛 폴더가 남아 있으면 "지울 것"으로 FAIL.
- `-h` 결과의 `skill_file`이 그 도구의 작업 스킬(예: excel.py -> `skills\office-excel\SKILL.md`), directive와 인자 오류 안내가 스킬 이름을 말함.
- 실기 러너·`run_all_tests.py`의 버전 표시는 `_runtime.py` VERSION / 첫 작업 스킬에서 읽음. 채점 규칙(no_read_code)은 `C:\CLI\office-cli\skills` 읽기·검색을 코드 읽기로 보지 않음.
- 전역 지침 3장 표(요청 -> 스킬)를 4개로, 저장소 AGENTS.md "폴더 구조"·"작업 절차 6"에 새 배치.

## [1.4.61] - 2026-11-04

워드 양식 채우기 (C4). 회귀 테스트: `tests/test_word.py` "word.py fill"(1.4.60에서 6개 중 6개 실패, 명령 없음).
가짜 Word에 칸 글이 있는 표(`FakeGridTable`, `doc.Tables(i).Cell(r, c).Range.Text`)를 넣음.

### 기능
- `word.py fill --file 값.csv`(열: 자리, 값): `{{이름}}` 자리는 문서 모든 영역에서 바꾸고, `표2!R3C4`는 그 표 칸에 씀. 모든 줄을 먼저 확인(모양, 문서에 없는 자리,
  없는 표·칸)하고 하나라도 틀리면 아무것도 바꾸지 않음. 값이 있는 표 칸은 `needs_user_decision` 하나와 `--overwrite`, `--dry-run`, 되돌리기 지점 하나,
  다시 읽어 확인.

## [1.4.60] - 2026-11-04

Outlook 이어 하기 (C2). 회귀 테스트: `tests/test_outlook.py` "짧은 번호, 첨부 저장 뒤 읽기 명령, calendar, thread"(1.4.59에서 9개 중 9개 실패).

### 기능
- 짧은 번호: `list`/`search`/`thread` 결과의 `id`가 1, 2, 3 ...(긴 EntryID는 `entry_id`). 마지막 목록을 데이터 폴더(`outlook_last_list.json`)에 기억해
  `--id 3`으로 읽기·첨부·답장·전달·대화. 긴 EntryID도 그대로 받음. `read`/`current`/`attachments`의 `mail.id`는 `mail.entry_id`로 바꿈(베타).
- `attachments --save`: 저장한 문서마다 `reader.py read --file`(3개까지), 둘 이상이면 `reader.py index --folder` 명령을 `next_commands`로.
- `outlook.py thread --id N`(읽기 전용): 같은 대화의 받은·보낸 메일을 오래된 순으로(20통, 본문 400자).
- `outlook.py calendar [--days 7]`(읽기 전용): 오늘부터 N일의 일정(반복 일정 펼침, 50개).

## [1.4.59] - 2026-11-04

두 판의 문서 비교 (B5). 회귀 테스트: `tests/test_reader.py` "reader.py diff"(1.4.58에서 5개 중 5개 실패, 명령 없음).

### 기능
- `reader.py diff --file "<이전 판>" --file2 "<새 판>" [--section "3"] [--save "차이.xlsx"]`: 장절 경로로 단락 묶음을 짝짓고(경로가 다르면 제목 유사도,
  짝 없는 장절은 `unpaired_sections`), 묶음 안에서 `삭제`/`추가`/`변경`. 숫자 변경은 말로(`5 V에서 5.5 V로`, 화살표 기호 없음) `number_changes`, 맨 앞에.
  항목마다 이전·새 근거(`old_evidence`, `new_evidence`). JSON 20개, 넘으면 `truncated`와 `full_result_file`(xlsx). `--save` xlsx/csv.

## [1.4.58] - 2026-11-04

닫힌 파일 열기 감시 (B4 1부). 문서 보안(DRM) 인증 창이 보이지 않는 인스턴스에 뜨면 명령이 끝나지 않아(에이전트 120초에 끊김)
LLM이 같은 명령을 반복함. 회귀 테스트: `tests/test_office.py` "닫힌 파일 열기 감시"(1.4.57에서 3개 중 3개 실패).

### 기능
- `_office.open_target`: 닫힌 파일을 여는 동안 감시(`OFFICE_CLI_OPEN_TIMEOUT`, 기본 30초, 0 = 없음). 시간이 지나면 도구가 띄운 인스턴스의 프로세스만 끝내고
  STOP_RETRY `error_type: drm_or_dialog`("문서 보안 인증 창 등으로 열리지 않음. 사용자가 문서를 직접 열어 두면 열린 문서로 읽음")와
  `next_commands[0]`(경로 대신 문서 이름을 넣은 같은 명령). 프로세스 번호: Excel은 창 핸들, Word·PowerPoint는 띄우기 전후 `tasklist` 차이.
- 하지 않은 것: DRM 문서를 reader 캐시에 평문으로 등록(회사 보안 정책 결정 필요).

## [1.4.57] - 2026-11-04

표 전체를 규격서와 한 번에 대조 (B1). 전에는 칸마다 `read` -> `search` -> 판단 -> 메모.csv를 LLM이 이어 붙여야 했음.
회귀 테스트: `tests/test_excel_read.py` "excel.py check"(1.4.56에서 5개 중 5개 실패, 명령 없음),
`tests/test_reader.py` "규격 글의 수치 조건"(규격 문장 33개 말뭉치, 단위·한계·후보 판정).

### 기능
- `excel.py check --range B2:D50 --item-col B --value-col D [--section-col E] --spec "<파일 또는 폴더>" [--section "3.2"] [--save "대조결과.xlsx"]` (읽기 전용).
  행마다 항목 글을 등록된 규격서에서 찾고(`reader.py search`와 같은 규칙: 동의어, partial/fuzzy), 근거 문장의 수치 조건(`±`, 범위, 이하·이상·미만·초과,
  영어 표기, 단위 접두어)을 칸 값(칸 글·머리글·항목 글의 단위)과 비교해 `일치`/`불일치`/`근거없음`/`판단필요`.
  판단필요: partial/fuzzy, 수치 후보 둘 이상, 단위 종류가 다름, 수치 조건 없음, 숫자가 아닌 값.
- 결과 `counts`, `results`(20개), `needs_review`, `not_decided`, `memo_file`(CSV 칸,메모 = "판정 | 근거"), `next_commands[0]` = `excel.py comment --action add --file <memo_file> --dry-run`.
  모든 행을 판정했을 때만 `VERIFIED:`. `--save`는 xlsx로 전체 결과.
- `reader.py`에 공용 도우미 `find_item`, `spec_conditions`, `near_conditions`, `quantity_of`, `unit_of`, `cond_holds`, `load_spec_docs`.

## [1.4.56] - 2026-11-04

긴 작업 나눠서 진행 (B2, 공통 계약 R10, ml-kit과 같은 이름·키). 회사 PC Gemma 시험에서 명령이 120초에 끊기면 결과 JSON이 없어
LLM이 같은 명령을 반복함. 회귀 테스트: `tests/test_reader.py` "R10 CONTINUE" (1.4.55에서 4개 중 4개 실패).

### 기능
- `_runtime.py`: `OFFICE_CLI_TIME_LIMIT`(초, 기본 100, 0 = 없음), `TIME_RESERVE` 30초, `check_time`, `Unfinished`(BaseException),
  `run_cli`가 `status: ok`, `verified: false`, `continue: true`, `progress`, `next_commands`(같은 명령 하나), `CONTINUE:` directive를 냄(exit 0, 오류 로그 없음).
  `next_commands` 항목은 office-cli 모양(`{purpose, command}`).
- `reader.py index`, `search --file 폴더`: 새 파일을 읽기 전마다 시계를 봄(이번 실행에서 한 파일은 읽은 뒤에만, 읽은 파일은 캐시에 저장돼 다시 실행하면 건너뜀).
  `search`(캐시만)의 바뀐 캐시 다시 읽기도 같음.
- SKILL.md: 결과 처리 표에 `CONTINUE:` 줄, 규칙 3에 예외. 테스트(conftest)와 실기 러너는 `OFFICE_CLI_TIME_LIMIT=0`.

## [1.4.55] - 2026-11-04

보호된 시트와 엑셀이 받지 않는 수식 (A6). 회귀 테스트: `tests/test_excel_write.py` "시트 보호 사전 확인과 엑셀이 받지 않는 수식"
(1.4.54에서 17개 중 13개 실패: 보호된 시트에 쓰다 COM 오류(내부 오류) 또는 일부만 바뀜, 거부된 수식이 내부 오류이고 먼저 쓴 칸·새 시트가 남음).
가짜 COM에 시트 보호(`ProtectContents`, `ProtectDrawingObjects`, `Protection.Allow*`, 칸 `Locked`)와 수식 거부(`reject_formula`)를 넣음.

### 기능
- 공통 `check_protection`: `write`(`--cells` 포함), `replace`, `style`, `range`(모든 동작, 행·열 넣기·지우기, 정렬), `comment`, `chart`가
  바꾸기 전에 시트 보호를 확인. 막히면 STOP_RETRY `error_type: sheet_protected`, "검토 > 시트 보호 해제는 사용자가 직접 (도구는 암호를 풀지 않음)",
  값 쓰기는 잠긴 칸과 잠기지 않은 칸 `[후보]`. 보호 대화상자에서 허용한 서식·행 넣기·정렬은 그대로 함.
- 공통 `apply_writes`: 엑셀이 받지 않는 수식은 어느 칸의 어떤 수식인지와 규칙(영어 함수 이름, 쉼표, 큰따옴표)을 STOP_RETRY로.
  따옴표 밖 `;`, 전각, 둥근 따옴표는 고친 전체 명령 `[추천]`. 먼저 쓴 칸을 되돌리고 같은 명령의 `--new-sheet` 시트를 지움(되돌리기 지점 없음).
  오류 로그에는 수식 글을 남기지 않음.

## [1.4.54] - 2026-11-04

흩어진 칸 여러 개 한 번에 쓰기 (A5). 회사 PC 기록에서 C6..C16을 채우려고 `write`를 16번, 덮어쓰기 확인을 8번 따로 받음.
회귀 테스트: `tests/test_excel_write.py` "write --cells" (1.4.53에서 9개 중 9개 실패, 옵션 없음). 실기 러너 E8.

### 기능
- `excel.py write --cells "C6=ft; C7=℃; 계산기!C11=20"` 또는 `--cells "칸값.csv"`(열: 칸, 값, `comment --file`과 같은 모양).
  `=`로 시작하는 값은 수식. 값이 있는 칸은 모두 모아 `needs_user_decision` 하나, `--overwrite` 한 번, 되돌리기 지점 하나, 쓴 칸을 모두 다시 읽어 확인.
  이미 같은 값인 칸은 그대로(`unchanged`), `--dry-run`(`planned`, `next_commands`). 하나라도 잘못되면 하나도 쓰지 않음.
- `write --file`에 머리글이 `칸, 값`인 파일을 주면 표로 쓰지 않고 `--cells` 전체 명령을 `[추천]`(표의 `--file`과 구분은 옵션 이름으로).

## [1.4.53] - 2026-11-04

문서에 없던 엑셀 서식 옵션 점검 (사용자 결정 "구현"). 회귀 테스트: `tests/test_excel_write.py` "문서에 없던 서식 옵션 점검"
(1.4.52에서 12개 중 4개 실패: 줄바꿈·테두리를 다시 읽지 않음, cp949로 쓸 수 없는 글자가 내부 오류).

### 수정
- `style --wrap`(`replace`도): 적용 뒤 `WrapText`를 다시 읽어 확인(R1). 전에는 확인 없이 `VERIFIED:`.
- `range --action border --weight/--box`: 선 모양·굵기를 다시 읽어 비교(R1, 한 줄·한 칸 범위의 안쪽 선은 건너뜀). 다르면 `UNVERIFIED:`.
- `export --encoding cp949`: 쓸 수 없는 글자(이모지 등)가 있으면 내부 오류 대신 쓰기 전에 STOP_RETRY(칸 주소, `[추천] utf-8-sig`).
  CSV 내보내기는 다시 읽어 행 수까지 확인.

### 문서
- 점검 결과 모두 동작해 남김(뺀 옵션 없음): `style --italic --wrap --no-zebra --size`, `range --action border --weight --box`,
  `write --no-autofit`, `export --encoding`. README `excel.py` 표에 설명, SKILL.md 서식 줄에 `--no-zebra`, `--italic`, `--wrap`, 외곽선(`--box --weight`).

## [1.4.52] - 2026-11-04

쓰지 않는 옵션 정리 (사용자 결정). 회귀 테스트: `tests/test_office.py` `test_transfer_method_option_removed`,
`tests/test_reader.py` `test_search_case_sensitive_option_removed` (1.4.51에서 2개 중 2개 실패).

### 구조
- `office.py transfer --method` 삭제. 차트 그림 방식은 자동 순서(`--probe`가 저장한 순서 먼저)만. 실기 러너 I24는 임시 `OFFICE_CLI_CONFIG`로 redraw를 먼저 시도.
- `reader.py search --case-sensitive`/`-c` 삭제. 검색은 늘 대소문자를 무시.

## [1.4.51] - 2026-11-04

`reader.py search` 재현율 (경량 LLM의 검색어가 규격서 글과 조금 달라도 근거를 찾게). 회귀 테스트: `tests/test_reader.py` "검색 재현율"
(1.4.50에서 12개 중 11개 실패). 재현율 시험(단락 7개, 검색어 8개, 기대 일치 10건): 1.4.50 2건 -> 1.4.51 10건.

### 기능
- 검색어마다 도구가 정규식을 만듦(`_single_pattern`): 검색어 끝 조사를 떼고 글의 조사를 허용, 전각 영문·숫자, `5V` = `5 V` = `5[V]` = `5.0 V`,
  `5000` = `5,000`, `°C` = `℃`. 숫자 경계(`15V`, `5.5V`)와 기존 규칙(`SOI 1` = `SOI #1`, `#4`, `1.2`)은 그대로.
- 0건이면 넓혀 찾기: `partial`(검색어 절반 이상, 가장 많이 든 단위) -> `fuzzy`(한 글자 다른 낱말, 영문 4자·한글 3음절 이상, 숫자 없는 검색어).
  `match_mode`, 항목의 `matched_terms`·`missing_terms`·`fuzzy_words`, `recovered`, directive에 "정확히 일치한 근거가 아님".
- 결과를 점수 순으로(`score`, 저장 파일 `점수` 열): 맞은 검색어 수, 나온 횟수, 검색어끼리 가까움, 장절 제목.
- 동의어 파일 `<데이터 폴더>/synonyms.txt`(한 줄에 한 묶음, 쉼표 구분). 결과 `synonyms_used`.
- `--regex` 오류: 몇 번째 글자와 이유. 정규식 문법이 없는 글이면 `--regex` 없이 찾는 전체 명령을 `[추천]`(오류 로그에는 검색어를 남기지 않음).

## [1.4.50] - 2026-11-04

규격 대조 흐름(판단 결과를 엑셀 칸 메모로). 회귀 테스트: `tests/test_excel_write.py` "comment" (1.4.49에서 9개 중 9개 실패, 명령 없음).
가짜 COM(`tests/fake_excel.py`)에 칸 메모(`Range.Comment`/`AddComment`, `Comment.Text()`/`Delete()`, `Worksheet.Comments`)를 넣음.
실기 러너 `tools/run_win_tests.py`에 E7(달기·확인 요청·덧붙이기·목록·지우기).

### 기능
- `excel.py comment`: 칸 메모(이전 방식 메모, Office 2007+와 365). `word.py comment`와 같은 모양.
  - 목록(기본): 시트 또는 `--range`의 메모 `cell`, `author`, `text`. 상한을 넘으면 `truncated` + `--start` 이어 보기 명령, 0개면 `VERIFIED: 0건`.
  - `--action add --range B5 --text "의견"`, 여러 개는 `--file "메모.csv"`(열: 칸, 메모, `load_table_file(raw=True)`와 `resolve_input_path`).
    한 줄이라도 잘못되면 하나도 달지 않음. 이미 메모가 있는 칸은 `needs_user_decision`(오류 로그에는 칸 주소만), `--overwrite` 바꾸기 / `--append` 덧붙이기.
  - `--action delete --range ...`(`--range` 필수). add/delete 모두 `--dry-run`(`planned`, `next_commands`), 되돌리기 지점, 다시 읽어 확인(`applied`).

## [1.4.49] - 2026-11-04

규격 대조 흐름(엑셀 칸 내용 -> 등록된 규격서의 장절 근거). 회귀 테스트: `tests/test_reader.py` "read --section" (1.4.48에서 3개 중 3개 실패).

### 기능
- `reader.py read --file 파일 --section "장절 일부"`: 장절 경로에 그 글이 든 단위만 읽음(`search --section`과 같은 규칙, 공통 함수 `section_key`·`section_matches`).
  `--pages`/`--units`와 함께 쓰면 두 조건을 모두 만족하는 단위. 결과 `section_filter`, 단위마다 `evidence`(search와 같은 근거 문구).
  읽기 상한·잘림·`--save`는 그대로이고, 이어 읽기와 전체 저장 명령에 `--section`(과 고른 범위)이 붙음.
  맞는 단위가 없으면 STOP_RETRY + `[후보]`(`suggest`로 찾은 비슷한 장절 이름, 없으면 문서의 장절 이름. `ambiguous`라 오류 로그에는 개수만).
  장절 정보가 없는 문서는 `--pages`로 읽으라는 STOP_RETRY. 스캔 PDF는 전처럼 OCR 안내(`UNVERIFIED:`).
  전에는 `search --regex --query "." --section ... --full`(10건 상한)뿐이었음.
- 1.4.39의 "reader.py read는 장절을 고르지 않습니다" 인자 오류 안내(`OTHER_COMMAND`)를 뺌.

## [1.4.48] - 2026-11-04

저장소 지침 반영(next_commands). 회귀 테스트: `tests/test_excel_write.py`(쓰기 뒤 차트 명령), `tests/test_word.py`·`tests/test_ppt.py`(닫힌 파일 읽기 `--save`)
(1.4.47에서 3개 중 3개 실패).

### 수정
- `excel.py write`(`--file`, `--formula`) 뒤 차트 명령에서 자리표시 `--title 차트`를 뺌(LLM이 그대로 실행해 제목이 '차트'인 차트가 생김).
  purpose에 "제목은 --title \"제목\"을 붙임".
- `word.py read --save`, `ppt.py read --save`의 "엑셀 표로 옮기기" transfer 명령: `--target`으로 다른 문서(열린 문서 이름, 닫힌 파일)를 읽었으면
  `--from-target`(전체 경로가 있으면 그것)을 붙임. 전에는 보고 있는 문서에서 옮겼음.

## [1.4.47] - 2026-11-04

저장소 지침 반영(R5): `require_lib()`. 회귀 테스트: `tests/test_contract.py` "R5 require_lib" (1.4.46에서 5개 중 5개 실패, 1개는 함수 없음).

### 수정
- `_runtime.require_lib(module, pip_name, feature)`와 `MissingLibraryError`: 라이브러리가 없으면 `STOP_RETRY:`, `error_type: "missing_library"`,
  `missing_package`, `install_command: "python -m pip install <패키지>"`, directive "install_command를 사용자에게 전달하세요. 직접 설치하지 마세요."
- 적용: reader의 DOCX(python-docx)·PPTX(python-pptx)·XLSX(openpyxl)·PDF(pymupdf)·HWP(olefile) 읽기와 xlsx 저장, excel.py 닫힌 xlsx 읽기와 수식 읽기
  (전에는 `import` 실패가 `ERROR:` 내부 오류), office.py transfer의 닫힌 워드 표 읽기. 전에는 reader는 `input` 오류 글에 설치 명령만 있었음.
- pywin32(`require_com`)는 Office 없이 읽는 다른 명령 안내가 있어 그대로 둠.

## [1.4.46] - 2026-11-04

1.4.37에서 남긴 것: range sort 기준 4개 이상. 회귀 테스트: `tests/test_excel_write.py` `test_sort_more_than_three_keys_uses_sort_fields`, `test_sort_too_many_keys`
(1.4.45에서 2개 중 2개 실패).

### 기능
- `excel.py range --action sort --by "팀,부서,등급,지역,이름"`: 기준 4개부터 64개(엑셀 상한)까지 `Worksheet.Sort.SortFields`로 정렬
  (기준마다 데이터 행의 열, 값으로 비교, `--desc` 그대로). 정렬한 뒤 `SortFields.Clear()`로 시트에 정렬 기준을 남기지 않음.
  3개까지는 전처럼 Range.Sort. 확인(R1)은 같음(모든 기준의 순서). 전에는 4개부터 STOP_RETRY.

### 테스트
- 가짜 Excel(`tests/fake_excel.py`): `Worksheet.Sort`(`SortFields.Add/Clear`, `SetRange`, `Header`, `Apply`).

## [1.4.45] - 2026-11-04

1.4.37에서 남긴 것: `read --where`의 또는(OR) 조건과 값 안의 `;`. 회귀 테스트: `tests/test_excel_read.py` `test_where_or*`, `test_where_quoted_value_with_separators`
(1.4.44에서 9개 중 6개 실패).

### 기능
- `excel.py read --where`: `|`는 또는. `;`(그리고)가 먼저 묶임: `부서 = 영업; 금액 > 250 | 부서 = 기획`. 두 묶음에 맞는 행도 한 번만.
- 따옴표로 감싼 값 안의 `;` `|`는 나누지 않음(`메모 = 'a;b'`). 따옴표 안의 `AND`·`==` 등도 고치지 않음.
- `OR`, `or`, `또는`, `||`는 `|`로 고친 식 `[추천]`(`AND`를 `;`로 고치는 것과 같은 방식). 결과 `where`와 `[추천]`은 `;` `|`가 든 값을 작은따옴표로 감쌈.

### 문서
- SKILL.md 엑셀 읽기 줄에 `|`(또는)와 값 따옴표 한 줄. README `--where` 절.

## [1.4.44] - 2026-11-04

24차 확인 검사 LOW-4. 회귀 테스트: `tests/test_excel_read.py` `test_far_uncomputed_*`, `test_near_uncomputed_row_still_kept` (1.4.43에서 4개 중 3개 실패).

### 수정
- 닫힌 xlsx에서 값(A1:C10)과 멀리 떨어진 계산되지 않은 수식 칸(Z200000)이 있으면:
  - 범위 생략 `read`: 그 칸까지 빈 행 20만 개를 채워 읽던 것을 데이터 범위(A1:C10)만. `-r A:Z`: "범위가 너무 큽니다"와 그 칸을 빼는 `[추천]` 대신 A1:Z10.
    `-r A:C`: 파일 크기 정보 때문에 20만 행을 채우고 `VERIFIED`이던 것을 A1:C10.
  - 요청한 범위 안의 먼 칸은 `UNVERIFIED:`, `far_uncomputed_formulas`, `far_uncomputed_count`와 그 칸만 읽는 명령(`next_commands` 끝). `export --save .csv`도 같음.
  - "먼 칸": 값이 있는 끝에서 1,000행(열) 넘게 떨어짐. 가까운 칸(값 바로 아래 합계 행 등)은 전처럼 범위에 넣음.
- `rt.xlsx_uncomputed_formulas(..., positions=True)`: 센 칸 위치(10만 개까지)와 값이 있는 칸의 범위(`value_bounds`).

## [1.4.43] - 2026-11-04

23차 점검 L1. 회귀 테스트: `tests/test_excel_read.py` `test_array_formula_*` (1.4.42에서 2개 중 1개 실패, 통과한 1개는 엑셀에서 저장한 배열 수식이 0개인지).

### 수정
- 닫힌 xlsx의 계산된 값이 없는 수식 칸 찾기(`excel.py read`, `export`, transfer, `reader.py`): 배열 수식 `<f t="array" ref="C1:C3">`의 `ref` 범위 모든 칸을
  수식 칸으로 셈. 다른 칸에 계산된 값이 없거나 파일에 그 칸이 없으면(openpyxl), `fullCalcOnLoad`이면 `uncomputed_formulas`에 들어감.
  전에는 왼쪽 위 칸만 세어 `uncomputed_count`가 적고, 범위를 생략한 읽기에서 그 행·열이 잘릴 수 있었음.

## [1.4.42] - 2026-11-04

24차 확인 검사 LOW-5. 회귀 테스트: `tests/test_reader.py` `test_check_failure_*` (1.4.41에서 2개 중 2개 실패).

### 수정
- reader: 닫힌 xlsx의 수식 칸 확인 실패(MemoryError 등, `meta.uncomputed_formulas.unknown`)를 캐시에서 그대로 다시 쓰던 것을, 캐시를 쓸 때마다
  (`read`, 색인 검색) 원본을 다시 확인하도록. 확인되면 meta와 캐시 파일을 고치고(`[INFO]`), 다시 실패하면 그대로 `UNVERIFIED:`.
  전에는 파일이 바뀌거나 캐시를 지울 때까지 같은 파일 읽기가 계속 "확인하지 못함"이었음.

## [1.4.41] - 2026-11-04

1.4.37에서 남긴 것: range delete가 다른 칸의 수식이 `#REF!`가 되는 것을 미리 알리지 않음.
회귀 테스트: `tests/test_excel_write.py` "지울 칸을 참조하는 다른 칸의 수식" (1.4.40에서 5개 중 5개 실패).

### 수정
- `excel.py range --action delete`: 다른 칸의 수식이 지울 행·열 안에만 있는 칸이나 범위를 참조하면(지우면 `#REF!`) 지우지 않고
  `needs_user_decision`. 지울 행·열이 비어 있어도 같음. message에 수식 칸 수와 위치(`시트!주소` 5개까지), 수식 글은 넣지 않음(오류 로그도).
- 결과 `broken_refs`, `broken_ref_cells`(20개까지)는 `--dry-run`과 실행 결과 모두에. `--dry-run`의 preview와 `next_commands[0]`(`--overwrite` 포함)도.
- 같은 통합문서의 다른 시트 수식(`Sheet1!B2`, `'내 시트'!C3`)도 봄. 범위 일부만 지워지는 참조, 다른 통합문서, 이름 정의·`INDIRECT`는 세지 않음.

## [1.4.40] - 2026-11-04

회사 PC에서 관리자 권한 PowerShell로 띄운 OpenCode가 사용자가 연 문서를 보지 못하고 "열려 있는 ... 문서가 없습니다"라고 답한 문제.
Windows는 권한 수준(무결성 수준)이 다른 프로세스 사이의 COM 연결을 막아, 관리자 권한 도구에는 일반 권한 Office가 보이지 않음.
회귀 테스트: `tests/test_office.py` "관리자 권한 실행" (1.4.39에서 11개 중 10개 실패).

### 수정
- 공용 `_runtime.is_elevated()`(`IsUserAnAdmin`) / `elevation_hint()`, `_office.no_open_doc`과 `no_open_docs_result`가 한 곳에서 판단.
  관리자 권한이면 열린 문서를 못 찾은 오류(`active`, 열린 문서 이름, transfer 받는 문서)가 `error_type: "elevated_process"`,
  message "관리자 권한으로 실행 중이라 일반 권한으로 연 문서를 볼 수 없음, 관리자가 아닌 일반 PowerShell 창에서 OpenCode를 다시 실행",
  directive는 다시 실행하지 말고 그대로 전하라는 것(경로를 물어 `--target`으로 다시 실행하라는 안내를 하지 않음).
- `excel/word/ppt.py info --all`, `office.py windows`의 0건: 관리자 권한이면 `VERIFIED: 0건` 대신 `UNVERIFIED:`와 같은 안내.
- Outlook 연결 실패 message에도 같은 안내(관리자 권한일 때만).
- 일반 권한에서는 바뀐 것 없음.

## [1.4.39] - 2026-11-04

회사 PC 오류 로그(2026-10-29~11-04, 실제 사용)의 인자 오류를 현재 버전으로 다시 실행해, 안내가 맞지 않던 것을 고침.
회귀 테스트: `tests/test_contract.py` (1.4.38에서 4개 실패).

### 수정
- 다른 파일의 명령이 맞는 경우 `[추천]`(옵션 이름만 바꿔 1회 재시도) 대신 맞는 명령을 안내 (`OTHER_COMMAND`):
  `office.py write`(없는 명령, 전에는 `[후보] windows, focus, transfer, undo`), `office.py transfer --text`(전에는 옵션 목록만),
  `reader.py read --section`(장절은 `reader.py search --section` 또는 열린 워드의 `word.py read --section`). 세 경우 모두 워드/PPT/엑셀 쓰기 명령을 그대로 적음.
- `excel.py write --cell K1`에 `[추천] --range` (전에는 옵션 목록만).

## [1.4.38] - 2026-11-04

회사 PC `opencode run` 기록(Gemma, "열려있는 엑셀 A2에 반가워 입력해줘")에서 찾은 세 가지.
회귀 테스트: `tests/test_excel_write.py` "Gemma 실행 기록(1.4.38)" (1.4.37에서 3개 중 2개 실패).

### 수정
- 인자 오류 추천: `excel.py write`에 워드/PPT 옵션 `--text`를 쓰면 철자가 비슷한 `[추천] --target`(따르면 또 실패)을 주던 것을
  같은 뜻의 `[추천] --value`로 (`OPTION_SYNONYMS`).
- 입력 표 파일을 못 찾았을 때 "사용 가능한 입력 표 파일"에 그림·PDF·문서 등 모든 파일을 나열하던 것을 표 파일(`.csv .tsv .md .txt`)만
  (`resolve_input_path(..., exts=)`). 관계없는 파일 이름이 LLM 응답에 섞이지 않음.

### 문서
- SKILL.md 명령 표에 "엑셀 칸에 값 하나 쓰기" `excel.py write --range A2 --value "반가워"` 줄 추가 (없어서 Gemma가 `--text`를 쓰고
  CSV 파일로 돌아가다 머리글 줄까지 쓰려 함).

## [1.4.37] - 2026-11-04

SKILL.md에 없던 기존 기능을 넣고(경량 LLM은 SKILL.md에 없는 기능을 "지원하지 않음"으로 답함) 빠진 기능 넷을 더함.
회귀 테스트: `tests/test_review_v1437.py` (1.4.36에서 48개 중 47개 실패·오류, 통과한 1개는 SKILL.md와 argparse 대조로 1.4.36 SKILL.md도 어긋남이 없었음).

### 기능
- `outlook.py list|search --days N`(오늘 포함 최근 N일, 1 = 오늘) / `--since YYYY-MM-DD`(그 날짜 0시부터). `--unread`, `--query`, `--from`과 함께 씀.
  - `list`: Jet `[ReceivedTime] >= 'YYYY-MM-DD HH:MM'`(현지 시각). `MM/DD/YYYY`는 일/월 순서 지역에서 뒤바뀔 수 있어 쓰지 않음.
  - `search`: DASL `"urn:schemas:httpmail:datereceived" >= '<UTC 시각>'`을 같은 `@SQL=`에 AND로 붙임(DASL 날짜 비교는 UTC. 현지 시각으로 읽혀도
    UTC보다 빠른 지역은 더 넓게 거르므로 빠지는 메일이 없음). `--days`/`--since`만으로도 검색됨.
  - 두 경우 모두 받은 시각을 메일마다 다시 확인하고(내림차순이라 기간 밖 메일을 만나면 멈춤), Outlook이 날짜 조건을 거부하면 날짜 없는 필터로
    거른 뒤 받은 시각으로 거르고 `recovered`로 알림(JIT 복구).
  - 결과 `period`(`since`, `until`, `days`)와 preview·0건 안내의 기간. 잘못된 날짜는 Outlook에 연결하기 전에 STOP_RETRY:
    고칠 수 있는 표기(`2026/10/1`, `20261001`, `2026년 10월 1일`, `2026-1-5`)는 `[추천] YYYY-MM-DD`, `--days 0`은 `[추천] 1`, `--days -7`은 `[추천] 7`.
    없는 날짜(`2026-02-30`), 미래 날짜, 읽을 수 없는 글, `--days`와 `--since`를 함께 쓰면 `[추천]` 없이 STOP_RETRY.
- `excel.py range --action insert|delete --range 5:7`(행) / `B:C`(열), `5`, `B`도 됨. `A5:C7`은 `--axis rows|cols`가 필요(없으면 `[후보]`).
  - insert는 범위 앞에 넣고 아래(오른쪽)가 밀림. delete는 지울 행·열에 값·수식이 있는 칸이 있으면 지우지 않고 `needs_user_decision`(칸 수와 주소
    5개까지. 값은 stdout·오류 로그 어디에도 넣지 않음), 사용자가 원하면 `--overwrite`.
  - `--dry-run`(insert/delete만): 바꾸지 않고 `filled_cells`, `filled_addresses`(20개까지), 실행 명령 `next_commands[0]`(값이 있으면 `--overwrite`).
    되돌리기 지점도 남기지 않음. 다른 range 동작에 `--dry-run`을 주면 STOP_RETRY.
  - 되돌리기 지점(열린 문서 스냅샷)은 다른 쓰기와 같음. 닫힌 파일은 다른 range 동작처럼 엑셀로 열어 고침.
  - 확인(R1): 사용 범위 끝 행(열)이 넣은·지운 수만큼 바뀌었는지, 옮겨진 칸 하나(수식이 아닌 값의 첫 칸, `moved`)가 기대한 자리에 같은 값인지.
- `excel.py range --action sort --range A1:E100 --by "<열>" [--desc [기준]] [--header auto|yes|no]`: Range.Sort, 기준 3개까지(`--by "부서,금액"`).
  - `--desc`만 쓰면 모든 기준 내림차순, `--desc "금액"`은 그 기준만. `--by`가 머리글 이름이면 첫 행은 머리글, 열 문자면 `--header auto`가
    판단(첫 행 칸이 모두 글자이고 굵게이거나 아래에 숫자·날짜가 있는 열이 있으면 머리글). 결과 `header`, `header_decided_by`, `keys`.
  - 없는 열은 머리글 이름 `[후보]`(오류 로그에는 남기지 않음). 병합된 칸은 STOP_RETRY, 범위 바로 옆 열의 같은 행에 값이 있으면(표를 가로지름)
    `needs_user_decision`과 표 전체 범위. 열 전체 범위는 사용 범위까지만. 이미 정렬되어 있으면 바꾸지 않음(`unchanged`).
  - 확인(R1): 다시 읽어 행 수, 머리글 행, 행 묶음이 그대로인지와 기준 순서(숫자 < 글자 < 논리값, 빈 칸은 맨 뒤). 순서를 확인하지 못하면 `UNVERIFIED:`.
- `excel.py read --where "<열> <연산자> <값>; ..."`: 첫 행을 머리글로 보고 조건에 모두 맞는 행만. 열린 문서, 닫힌 xlsx, CSV 모두 같은 읽기 결과를 씀.
  - 연산자 `>= <= > < = != ~`(포함). 숫자는 숫자로(`1,000` 포함), 날짜 칸은 `YYYY-MM-DD` 값과 날짜로, 글자는 앞뒤 공백을 빼고 대소문자 없이.
  - 결과 `values`(머리글 + 맞는 행), `row_numbers`(`values`와 같은 순서의 시트 행 번호), `count`, `data_rows`, `where`. 0건은 `VERIFIED: 0건.`(R2).
    잘림(`truncated`, `--where`가 든 저장 명령)과 `--save`(머리글과 맞는 행 CSV, 다시 읽어 확인)는 read와 같음.
  - 없는 열은 머리글 이름 `[후보]`. 고칠 수 있는 표기(`==`, `=>`, `<>`, `AND`, `contains` 등)는 고친 식 `[추천]`, 읽을 수 없는 식은 형식과 예만.
    `--search`, `--formulas`와 함께 쓸 수 없음.

### 문서
- SKILL.md에 이미 있던 기능을 넣음: 메일 `list [--unread] [--folder] [-n]`, `folders`, `search --folder --where subject`, `reply --all`, `new --cc`, `--body-file`,
  PPT `slide --action add|duplicate|delete|move --slide --to --layout`, 엑셀 `sheet --action add|rename|copy|delete --name --new-name`,
  `range --action freeze|...|border --dest --size`, `export --save .csv|.xlsx|.pdf --range`, `style --number-format --align --border`,
  `replace --all-sheets`, 차트 `--type`, 워드 `write --toc`, `--callout`, `replace --find --highlight`(서식만).
- SKILL.md에 새 기능(`--days`/`--since`, `range insert|delete|sort`, `read --where`). 행·열 삭제와 `--overwrite`(지우기)는 사용자가 요청한 경우에만.
- README excel·outlook 절, 사용 가이드 2.1(조건에 맞는 행), 2.7(행·열 넣기·지우기, 정렬), 6절(기간으로 메일 보기).

### 테스트
- SKILL.md의 백틱 구간(`<script>.py <명령> --옵션`, 표의 한 행 안에서는 앞 명령에 딸린 `--옵션`과 `명령 --옵션`)의 명령·옵션·`a|b` 선택지가
  argparse에 있는지 대조(앞으로 문서와 코드가 어긋나지 않게).
- 가짜 Excel(`tests/fake_excel.py`): `EntireRow`/`EntireColumn`의 `Insert`/`Delete`(칸 옮기기), `Range.Sort`(Key1~Key3, Header).

## [1.4.36] - 2026-11-04

엑셀 수식 읽기. 회귀 테스트: `tests/test_review_v1436.py` (1.4.35에서 18개 중 17개 실패, 통과한 1개는 `--formulas` 없는 읽기가 그대로인지).

### 기능
- `excel.py read --formulas`: 값은 그대로 두고 읽은 범위의 수식 칸 목록을 더합니다. `formulas`(`cell`, `formula`, `value`, 배열 수식은
  `array`, `array_range`), `formula_count`, `formulas_truncated`. 200개를 넘으면 앞 200개만 담고 전체 목록 CSV를 `full_result_file`로 줍니다.
  `--save`와 함께면 값 CSV 옆 `<이름>_formulas.csv`에 전체 목록을 쓰고 다시 읽어 행 수를 확인합니다.
  수식이 없으면 `formula_count: 0`과 "수식 0개: ...에는 수식이 없습니다" 안내. CSV는 늘 0개(수식을 저장하지 않음).
- 열린 문서: `HasFormula`가 False면 수식을 읽지 않고, 수식 글은 범위 전체를 한 번에 `Formula2`(없으면 `Formula`, 영어 함수 이름 A1 주소)로 읽습니다.
  값과 같은 글인 칸(`'=` 글자 상수일 수 있음)만 `HasFormula`로, 배열 수식은 범위의 `HasArray`가 False가 아닐 때만 칸마다 확인합니다(2,000칸까지).
- 닫힌 xlsx/xlsm: 같은 범위만 openpyxl `data_only=False`로 다시 읽습니다(보통·공유·배열 수식, 데이터 표). 값은 저장된 계산된 값이고
  계산되지 않은 수식은 `null`(계산되지 않은 수식 `UNVERIFIED:` 안내는 1.4.31/1.4.32 그대로).
- Office 없이 수식을 읽을 수 없는 닫힌 파일(.xls 등)은 Office(pywin32)가 없으면 `STOP_RETRY:`(엑셀에서 연 뒤 다시 실행). Office가 있으면 엑셀로 열어 읽습니다.
- `--search`와 함께면 수식 글에서도 찾고 찾은 수식 칸에 `formula`를 붙입니다.
- 수식 읽기 실패는 예외 종류만 남기는 내부 오류(`ERROR:`)라 수식 글이 오류 로그에 남지 않습니다.

### 문서
- SKILL.md 엑셀 읽기 줄에 `--formulas`(수식 보기). **SKILL.md 본문이 바뀌었으므로 En-Genius 스킬 zip(`docs/engenius/office-cli.zip`)을 다시 가져와야 합니다.**
- README excel 절, 사용 가이드 2.1 "수식 보기".

### 테스트
- 가짜 Excel(`tests/fake_excel.py`): `HasFormula`, `HasArray`, `CurrentArray`(배열 수식 `arrays`), 범위 `Formula` 읽기 기록.

## [1.4.35] - 2026-11-04

En-Genius 스킬 배포 구조 변경에 맞춤. 회귀 테스트: `tests/test_review_v1433.py` (고침).

### 바뀐 동작
- En-Genius 스킬 zip(`docs/engenius/office-cli.zip`)에는 지시문과 `run.py`만 들어가고, `run.py`가 `C:\CLI\office-cli`의 도구를 실행합니다.
  그래서 `ENGENIUS_OUT`이 있어도 도구 폴더의 `config.json`, `data`, `cache`(OpenCode와 같은 것)를 씁니다.
  `%USERPROFILE%\Documents\En-Genius\CLI-data\office-cli` 고정 폴더(1.4.33)는 도구 파일이 트레이 임시 폴더(`engenius-skill`)에서 실행될 때만 씁니다.
- 도구 업데이트 때 스킬 zip을 다시 가져올 필요는 `SKILL.md` 본문(버전 줄 제외)이 바뀐 경우뿐입니다. `run.py`가 지문을 비교해 알려 줍니다.

## [1.4.34] - 2026-11-04

메일 첨부 내려받기 명령과 답장·전달 첨부 안내. En-Genius 에이전트가 "첨부를 내려받을 수 없다"고 답한 문제에서 시작.
회귀 테스트: `tests/test_review_v1434.py` (1.4.33에서 44개 중 38개 실패. 통과한 6개는 이미 되던 답장·전달·새 메일 `--attach`).

### 기능
- 새 명령 `outlook.py attachments [--id <id>]`: 첨부 목록(`index` 1부터, `name`, `size`, `hidden_inline` 본문 삽입 이미지, `blocked` 실행 파일 형식).
  `--id`를 빼면 열린 창이나 목록에서 선택한 메일(`reply`와 같음). 첨부가 없으면 `count: 0`과 `VERIFIED: 0건. 이 메일에는 첨부 없음`.
- `attachments --save [폴더]`: 내려받기. 기본 폴더는 `<다운로드>\메일 첨부\<YYYYMMDD_HHMMSS>_<제목 40자>`(자동 정리 안 함, En-Genius 실행도 같음).
  다운로드는 Windows 셸 폴더 설정을 따르고, 읽지 못하면 `%USERPROFILE%\Downloads`. 폴더를 주면 전체 경로만 받고(없으면 만듦) 상대 경로는 STOP_RETRY.
  결과 `saved`(전체 경로), `skipped`(`name`, `reason`), `count`, `folder`. 덮어쓰지 않고 ` (2)`를 붙여 `recovered`로 알림.
  저장 뒤 파일이 있는지, 크기를 아는 첨부는 빈 파일이 아닌지 다시 확인하고 아니면 `UNVERIFIED:`(Outlook `Size`는 메일 속성 몫이 더해져 같은 크기는 비교하지 않음).
- `attachments --save --name 이름일부`(대소문자 무시, `;`·옵션 반복) / `--index 1,3`(`2-4`): 고른 첨부만 저장. 맞는 것이 없는 값이 있으면
  아무것도 저장하지 않고 첨부 이름을 `[후보]`로 STOP_RETRY. 실행 파일 형식은 골라도 저장하지 않고, 본문 삽입 이미지는 직접 고른 경우에만 저장.
- `read`/`current`: 첨부가 있으면 `attachment_names`(본문 삽입 이미지 제외, 20개까지, `attachment_names_truncated`), `inline_images`,
  `next_commands` "첨부 내려받기"(`attachments --id <id> --save`, `--save`와 함께면 워드 넣기 다음).
- `new/reply/forward --attach`: 전체 경로 파일이 없으면 Outlook에 연결하기 전에 같은 폴더의 비슷한 파일을 전체 경로 `[후보]`로 STOP_RETRY
  (폴더가 없으면 그렇게 알림). 붙인 뒤 `Attachments.Count`가 붙인 수만큼 늘었는지 다시 셈(R1, 전달은 원본 첨부를 그대로 두고 늘어난 수로 비교).
  맞지 않으면 작성 창만 띄우고 `UNVERIFIED:`, `--send`면 보내지 않고 STOP_RETRY(`attach_problems`). 답장·전달 첨부는 원래 되던 것을 테스트로 고정.
- 명령 별칭: `outlook.py download|attach|attachment` -> `[추천] attachments`.

### 바뀐 동작
- `read`/`current --save-attach`를 없앴습니다(베타, 호환 없음). 7일 뒤 정리되는 `data\_work\mail_attachments\`에 저장해 내려받기가 아니었습니다.
  이제 `attachments --save`를 씁니다. 결과 키 `saved_attachments`, `skipped_attachments`도 없앴습니다.

### 문서
- SKILL.md 메일 줄을 읽기 / 첨부 내려받기 / 쓰기·문서 첨부로 나눔(`reply --body ... --attach "C:\전체경로\파일.pdf"`, new·forward도 같음).
- README outlook 절, 파일 위치 표(메일 첨부), 사용 가이드 6절, En-Genius README 데이터 표.

## [1.4.33] - 2026-11-04

En-Genius(사내 에이전트) 스킬 실행 대응. 회귀 테스트: `tests/test_review_v1433.py` (1.4.32에서 4개 중 3개 실패).

### 기능
- En-Genius 트레이는 스킬 파일을 실행마다 `%TEMP%\engenius-skill\<id>\`에 풀고 끝나면 지웁니다. 그래서 도구 폴더 아래의 `config.json`(transfer 방식 기억),
  `data`(오류 로그, `_work` 결과, 메일 첨부), `cache`(문서 색인, 되돌리기 스냅샷)가 다음 실행에 남지 않았습니다.
  환경변수 `ENGENIUS_OUT`이 있으면(트레이가 넣음) 세 가지의 기본 위치를 `%USERPROFILE%\Documents\En-Genius\CLI-data\office-cli\`로 바꿉니다.
  OpenCode·`C:\CLI` 실행은 바뀌지 않습니다.
- 환경변수 `OFFICE_CLI_CONFIG`로 `config.json` 위치를 정할 수 있습니다. `OFFICE_CLI_DATA_DIR`, `OFFICE_CLI_CACHE_DIR`와 함께 En-Genius 기본값보다 먼저입니다.

## [1.4.32] - 2026-11-03

23차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1432.py`.

### 수정
- 닫힌 xlsx를 범위 없이 읽으면(`excel read`, `--search`, `--save`, `export --save .csv`, `office.py transfer --from excel`) 끝쪽의 값이 모두 None인
  행을 잘라, 계산된 값이 없는 수식만 있는 합계 행(`=SUM`)이 빠지고 `VERIFIED`이던 문제. 수식만 있는 시트는 머리글만 나왔고, 크기 정보(`<dimension>`)가
  없는 파일은 `_closed_bounds`가 끝쪽 수식 열도 잘랐음. 범위를 데이터로 정하는 곳(범위 생략, `B:B`·`5:7`, 200만 칸을 넘는 범위 줄이기
  `_clip_closed_range`)에서 자르기 전에 시트 전체를 `xlsx_uncomputed_formulas`로 훑어(빈 칸 조건 없이) 그런 칸의 행·열을 남기고(`bounds`),
  읽은 범위의 칸을 `UNVERIFIED:`와 `uncomputed_formulas`로 알립니다. 같은 읽기에서 훑은 결과는 `closed_uncomputed`가 다시 씁니다.
  reader는 원래 시트 전체를 확인해 해당 없음(빈 행은 단위가 아님).
- xlsxwriter가 쓴 파일: 수식 칸에 `<v>0</v>`을 쓰고 `workbook.xml`에 `<calcPr fullCalcOnLoad="1"/>`을 두어, 0을 값으로 `VERIFIED`이던 문제.
  `fullCalcOnLoad`가 `"1"`/`"true"`이면 모든 수식 칸을 계산된 값이 없는 칸으로 봅니다(excel·transfer·reader 모두, 빈 칸이 없는 범위도 확인).
  엑셀은 저장할 때 이 표시를 지우므로 엑셀에서 저장한 파일은 해당 없음(테스트의 "엑셀 저장 흉내" `set_cached_value`도 표시를 지움).
  openpyxl도 이 표시를 쓰므로 안내 문구는 "열 때 다시 계산" 파일로 나옵니다. reader 파서 버전 10(캐시 다시 읽기).
- 수식 칸 확인이 예외(MemoryError 포함)로 실패하면 0개로 보아 `VERIFIED`이던 문제. `{"unknown": true, "error": 예외 이름}`을 돌려주고
  결과는 `UNVERIFIED: ... 수식 계산 여부를 확인하지 못함`과 `uncomputed_check: "unknown"`(reader는 `meta.uncomputed_formulas.unknown`).
  범위를 주면 범위 끝 행 뒤는 해석하지 않고, 해석한 `<row>`는 지웁니다.

### 문서
- README: 범위를 데이터로 정할 때 수식 행·열을 남김, `fullCalcOnLoad`, 확인 못함(`uncomputed_check`).

## [1.4.31] - 2026-11-03

22차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1431.py`.

### 수정
- `office.py transfer --method relay --to ppt` (1.4.30 회귀): 슬라이드를 확인하려고 받는 PPT의 쓰기 세션(되돌리기 지점)을 임시 워드 단계보다
  먼저 열어, 임시 워드 단계가 실패하면 저장 안 된 활성·열린 PPT에 바꾼 것 없는 되돌리기 지점이 남고 다음 `undo`가 그 지점을 먼저 골랐음.
  열린 문서·활성 문서는 `resolve_target`으로 찾은 문서에서 `ppt._resolve_slide`로 확인만 하고, 쓰기 세션은 임시 워드 단계 뒤에 엽니다(1.4.29와 같음).
  닫힌 파일은 열어서 확인합니다(멈추면 디스크 파일이 그대로라 지점은 지움). 새로 만들 파일은 슬라이드 0장으로 확인하고 만들지 않습니다
  (이전에는 슬라이드 오류에도 빈 파일을 만들었음). 받는 문서, 앵커, 슬라이드 오류가 클립보드와 임시 워드보다 먼저 멈추는 것은 그대로입니다.
- 닫힌 xlsx의 계산된 값이 없는 수식 칸: openpyxl `data_only`로 읽어 openpyxl 같은 프로그램이 쓰고 엑셀에서 연 적 없는 파일의 수식 칸이 빈 칸인데
  `VERIFIED`이던 문제(`excel read`, `--search`, `--save`, `export --save .csv`, `office.py transfer --from excel --range`, `reader read/search`).
  `_runtime.xlsx_uncomputed_formulas`가 시트 XML에서 값(`<v>`)이 비어 있는 수식 칸을 찾고(수식 태그가 없는 시트는 해석하지 않음, 빈 글 결과
  `t="str"`는 계산된 값), 있으면 `UNVERIFIED:`와 `uncomputed_formulas`(20개까지), `uncomputed_count`, `uncomputed_truncated`로 알리고 엑셀에서
  열어 저장한 뒤 다시 실행하라고 안내합니다. excel은 읽은 범위에 빈 칸이 있을 때만 그 범위를 확인합니다. reader는 `meta.uncomputed_formulas`에
  남기고 search는 `not_searched`에 넣습니다. 범위 전체가 그런 칸이라 `transfer`가 '데이터 없음'으로 멈추면 이유를 message에 붙입니다.
- docx 글상자(`w:txbxContent`): 글상자 글을 읽지 않아 `reader search`가 `VERIFIED` 0건이던 문제(닫힌 파일의 `word.py read`, `transfer --what text`도
  같은 경로). 글상자가 든 단락·표 바로 뒤에 `글상자 N`(`글상자 N 단락 M`) 단위로 넣습니다. `mc:Fallback`(VML 사본) 안의 글상자는 빼서 두 번 세지 않고,
  표 칸 글에서도 글상자 단락은 뺍니다(따로 단위). 단락 번호는 그대로라 글상자가 없는 문서의 단위는 같습니다. 파서 버전 9.
- `office.py transfer --from word --what table` (닫힌 docx): python-docx `Document.tables`(본문 바로 아래 표만)로 표를 세어, 콘텐츠 컨트롤(`w:sdt`)
  안에 표가 있으면 read 결과·COM과 다른 표를 `VERIFIED`로 옮기던 문제. read와 같은 `reader._body_blocks`로 셉니다. 닫힌 docx의 표 번호를 쓰는
  다른 곳은 없음(ppt `--table`은 pptx 도형).
- `word write` 굵게 위치: 파이썬 글자 수로 세어 앞에 이모지(BMP 밖 글자)가 있으면 굵게가 밀리던 문제. Word처럼 UTF-16 단위로 셉니다(`_u16`,
  단락 범위 끝도 같음). word.py의 다른 Range 위치는 COM에서 받은 값이라 해당 없음.
- `excel read --search`: `count`가 없고 0건 안내가 R2 모양이 아니던 문제. `count`(= `total_matches`), `shown`, `VERIFIED: 0건. ...`.
- 마크다운 `---`(구분선): word가 페이지 나누기를 넣고, 맨 앞 YAML 머리말은 페이지 나누기 두 개와 `key: 값` 글로 들어가던 문제.
  구분선은 아래 테두리만 있는 빈 단락(`구분선`), 맨 앞 YAML 머리말(`---` 1번째 줄, `key: 값` 줄, `---`/`...`로 닫힘)은 넣지 않고 `warnings`.
  ppt `.md`도 머리말을 넣지 않고 `warnings`(구분선은 전처럼 넣지 않음). 페이지 나누기는 `word.py write --page-break`. 다시 읽기 확인은 넣은 첫 글
  (구분선 빈 단락과 머리말은 건너뜀). 머리말만 있는 글도 마크다운으로 봅니다(`looks_like_markdown`).

### 문서
- README: 계산된 값이 없는 수식 칸, `read --search`의 `count`, 구분선과 YAML 머리말, word 굵게 위치, reader 글상자, 닫힌 docx 표 번호, relay 쓰기 세션 순서.
- SKILL.md: 워드 쪽 나누기는 `--page-break`(`.md`의 `---`는 구분선), `warnings`도 알림.
- 남은 할 일: excel read --search(R2 `count`·`shown`·`truncated`·`next_commands`)를 뺐습니다.

## [1.4.30] - 2026-11-03

21차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1430.py`.

### 수정
- `outlook.py --folder` (1.4.29 보완): 찾지 못한 폴더 오류의 `(경로: <--folder 원문>)`이 오류 로그에 그대로 남아, `계정\없는폴더`처럼
  계정 이름으로 시작한 경로면 메일 주소가 로그 `결과:` 줄에 남던 문제. `not_found`에 `log_extra`(로그 쪽 extra)를 두고, 계정 이름으로
  시작한 경로는 첫 부분을 `<메일 계정>`으로 바꿔 남깁니다(기본 폴더 별칭으로 시작한 경로는 그대로). stdout message는 그대로입니다.
  - 1.4.29 가드 테스트(`test_guard_folder_log_has_no_mailbox_root`)가 계정 이름을 소문자로 넣고 대소문자를 구분해 확인해 실패할 수
    없었음. 같은 대소문자로 넣어 확인합니다.
  - 다른 `not_found(..., extra=...)` 호출(`resolve_input_path`의 찾은 위치, 문서 열기 안내, 색상, 차트의 시트, 워드 제목 안내)은
    인자·경로이거나 사용자가 준 시트 이름이라 해당 없음.
- `office.py transfer --method relay --to ppt`: `--slide`를 확인하기 전에 차트를 클립보드에 복사하고 임시 워드를 띄워, 2장짜리 PPT에
  `--slide 9`면 사용자의 클립보드를 덮어쓴 뒤 오류로 끝나던 문제. 받는 PPT를 먼저 열고(열린 문서·활성 문서·파일 모두) 슬라이드를
  확인한 뒤에 클립보드와 임시 워드를 씁니다. 받는 문서, 앵커, 슬라이드 오류는 모두 클립보드를 쓰기 전에 멈춥니다.
- `office.py transfer --method relay --to-target ""`: 빈 대상을 '넣지 않음'으로 읽어 `KeyError('word')`를 STOP_RETRY로 내던 문제.
  export·redraw처럼 활성 문서로 넣습니다.

### 문서
- README: outlook `--folder` 오류 로그의 경로, relay의 슬라이드 확인 순서와 빈 `--to-target`.

## [1.4.29] - 2026-11-03

20차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1429.py`.

### 수정
- `office.py transfer --chart`: 받는 문서 이름 오타(`[추천]`)에도 다음 방식(redraw, relay)을 시도해, relay가 사용자의 클립보드를 덮어쓰고
  마지막 오류는 `[추천]`이 `; relay: ...` 앞에 묻힌 채 `--probe`를 안내하던 문제. 대상을 찾지 못한 오류는 다음 방식으로 넘어가지 않고
  원래 오류 그대로 멈춥니다(`[추천]`이 message 끝).
  - 구조: `_office.resolve_target`(open_target의 찾기 단계, 문서를 열지 않음)이 내는 오류(`[추천]`, `[후보]`, 사용 가능한 문서, 확장자,
    열린 문서 없음)는 모두 `no_fallback`. `ppt._resolve_slide`의 슬라이드 번호 오류, `--anchor` 길이 오류도 `no_fallback`.
  - `relay_chart`는 받는 문서를 먼저 찾고, 워드는 앵커 위치를 정한 뒤에 차트를 클립보드로 복사합니다.
  - 다른 방식 전환 고리(`transfer --probe`, 범위 전달)는 방식마다 같은 대상을 다시 찾지 않거나 첫 오류에서 멈추므로 해당 없음.
- 오류 로그 (transfer --chart): 모든 방식이 실패했을 때 `log_message`가 없는 내부 오류(RuntimeError, ValueError 등)의 글이 그대로 남던 문제.
  `<방식>: <예외 종류>`만 남깁니다.
- `outlook.py --folder`: 찾지 못한 폴더를 전체 경로끼리 비교해 같은 상위 경로(메일 계정) 때문에 관계없는 폴더가 모두 `[후보]`이던 문제.
  끝 이름끼리 비교하고 보여 줄 때만 전체 경로로 바꿉니다(`not_found`의 `display`). 오류 로그에는 계정 이름(메일 주소) 대신 `<메일 계정>`
  (`not_found`의 `log_name`).
- 차트 다시 그리기(redraw): 제목·계열 이름·항목 이름의 `$...$`를 matplotlib 수식(mathtext)으로 읽어 깨진 그림을 `VERIFIED`로 넣거나
  `ValueError`로 실패하던 문제. `text.parse_math`를 끈 설정 안에서 저장까지 그립니다(전역 설정은 그대로).
- `reader search --file <파일 하나>`: 파서 내부 오류(`KeyError` 등)를 사유 글로 새 `InputError`를 만들어 `STOP_RETRY:` 입력 오류로 내고
  traceback이 남지 않던 문제. 원래 예외를 그대로 다시 내어 내부 오류는 `ERROR:`와 traceback, `InputError`는 `log_message`를 지킵니다.
  `index --file <파일 하나>`도 같습니다(이전에는 실패 1개로 `VERIFIED: 등록 완료`). 폴더의 실패 사유(`reason`)는 예외 종류와 함께
  (`KeyError: 'heading'`, 이전에는 `'heading'`).

### 개발 도구
- `tools/run_win_tests.py`: pywin32 확인을 인자 해석보다 먼저 해 pywin32가 없으면 `--agent -h`도 exit 1, 인자 오류도 exit 1이던 문제.
  인자를 먼저 해석합니다(`-h` exit 0, 인자 오류 exit 2, 그다음 pywin32가 없으면 exit 1).

### 문서
- README: transfer 차트 방식 전환이 멈추는 경우, outlook `--folder` 후보, reader `--file` 파일 하나의 오류, 러너 종료 코드.

## [1.4.28] - 2026-11-03

19차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1428.py`.

### 수정
- 오류 로그 message (1.4.27 보완): 가리기가 message의 표식(`[후보]`, `[추천]` 등)에만 기대어, 표식 없이 들어간 문서 글이 그대로 남던 문제.
  `excel write`·`range --action merge|move|copy`의 덮어쓸 칸 `A1=<칸 값>`, `ppt write`의 기존 제목, `ppt notes`의 기존 노트,
  없는 시트의 `(있는 시트: ...)`, `word comment --file`이 `[후보]`·`[추천]`을 `후보:`·`추천:`로 바꿔 남긴 메모 파일의 기준 글과 후보 단락,
  종류에 공백이 있는 `사용 가능한 메일 폴더:`·`사용 가능한 열린 Word 문서:`·`사용 가능한 입력 파일:` 목록, `word replace`의 긴 찾을 문구 앞 30자
  미리보기, 내부 오류 traceback의 가린 인자 값.
  - 구조: `InputError`·`UserDecisionError`에 `log_message`(문서 글을 뺀 글)를 두고, 문서 글을 message에 넣는 곳마다 줍니다.
    `ambiguous()`·`UserDecisionError`의 후보, `not_found()`의 `[추천]`·`사용 가능한 <종류>:` 목록은 자동으로 개수만 남깁니다(`[후보 N개 생략]`).
    결과 dict로 내는 오류(`outlook send`의 확인 못 한 수신자)는 `_log_message` 키를 쓰고, `run_cli`가 stdout에 내기 전에 뺍니다.
    `office.py transfer --chart`가 모든 방식에 실패한 오류의 `시도한 방식:` 목록도 각 오류의 `log_message`로 만듭니다.
  - 표식 자르기는 대비로 남기고 `후보:`, `추천:`, 공백이 든 종류(`사용 가능한 [^:]{1,30}:`)를 더했습니다.
  - 가린 값이 10자 이상이면 10자 이상의 앞부분(잘린 미리보기)도 가립니다. `word replace`의 긴 찾을 문구 오류는 앞부분 대신
    몇 번째 문구인지와 길이를 알려 줍니다.
  - traceback에도 같은 값 가리기를 합니다(표식 자르기는 하지 않음).
  - stdout JSON의 message는 그대로이고, 가리기에 실패해도 로그 기록은 오류를 삼킵니다.
- 짧은 가린 값: `--value 1`처럼 1~2자 값이 관계없는 글까지 가리던 문제(`Sheet<1자 생략>`). 3자 이상만 가리고, 1~2자 값은 따옴표 안에
  있을 때만 가립니다. 따옴표 없이 message에 들어간 1~2자 값은 로그에 남을 수 있습니다(남는 위험).
- `office.py undo`: 되돌릴 수 없는 작업(새로 만든 파일 등)만 있을 때 `needs_user_decision`이던 문제. 되돌릴 지점이 하나도 없으면
  `VERIFIED: 0건.`(`count: 0`, `restored: false`)입니다. 더 앞에 되돌릴 수 있는 지점이 있으면 그대로 사용자에게 묻습니다.
- `office.py undo --list`: 되돌릴 수 없는 작업이 20개 넘게 앞에 있으면 앞 20개만 보여 줘 되돌릴 수 있는 지점이 목록에서 빠지던 문제.
  항목은 되돌릴 수 있는 지점 앞 20개와 그보다 최근의 되돌릴 수 없는 작업(최대 20개)이고, `count`·`shown`(지점)과 `truncated`(빠진 항목이
  있음)가 항목과 맞습니다.

### 문서
- README 오류 로그 절: `log_message` 우선, 값 가리기 규칙(앞부분, 짧은 값), 표식 대비, traceback, 남는 것.
- README `undo`: `--list` 항목 구성과 0건 결과.

## [1.4.27] - 2026-11-03

18차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1427.py`.

### 수정
- 제목 굵게 표시 (1.4.26 회귀): `ppt write --file .md`가 3단계 이하 제목에만 굵게 표시 빼기를 적용해 `### __init__ 메서드`는 `init 메서드`,
  1·2단계 제목(슬라이드 제목)은 `**핵심** 요약`이 글자 그대로 들어가던 문제. word·ppt의 모든 수준 제목, 슬라이드 제목(`.md` 제목,
  `--title`과 함께 줘서 본문 첫 줄로 옮긴 제목, `ppt write --title`)에 같은 규칙(`_runtime.plain_heading`: `**x**`, `__x__`는 표시를 뺀 글)을 씁니다.
  제목은 굵기가 섞인 글을 담지 않으므로 굵게 구간은 본문에만 둡니다.
- 단어 안의 `__` (`_runtime._MD_BOLD`): `report__2024__final.xlsx`가 `report2024final.xlsx`로 들어가던 문제. CommonMark처럼 `__`는 앞에 글자·숫자가
  없을 때만 열고 뒤에 없을 때만 닫습니다(`파일__x__끝`은 그대로). `**`는 그대로(단어 안에서도 굵게). 양쪽이 단어 경계인 `__init__`은 CommonMark대로
  굵은 `init`이고, 제목에서는 모든 수준에서 `init`입니다(코드 이름은 코드 블록(```)이나 `--callout`으로 넣으면 글자 그대로).
- 번호 목록: `2026. 11. 3. 회의 결과`를 번호 목록으로 읽어 연도가 사라지던 문제(word·ppt). 목록 번호는 1~3자리 숫자만 봅니다(`_runtime.MD_NUMBERED`).
  1이 아닌 번호로 시작하는 목록(`3. 셋`)은 Word·PowerPoint 자동 번호가 1부터 다시 매기므로, 번호를 글로 둔 일반 단락(줄)으로 넣습니다.
  빈 줄로 띄운 `1. 가` / `2. 나`는 한 목록으로 이어 붙입니다. 다시 읽기 확인은 넣은 글 그대로(`3. 셋`)를 기대합니다.
- 오류 로그 message: `excel style --highlight`의 `입력: <규칙>`, word 앵커의 `비슷한 문구: '<문서 글>'`, 장절 `[후보]`(문서 제목) 등
  가린 인자 값과 문서에서 가져온 글이 message에 그대로 남던 문제. `write_error_log`가 `redact_argv`로 가린 값을 같은 `<N자 생략>`으로 바꾸고,
  `LOG_CUT_MARKERS`(`[후보]`, `[추천]`, `비슷한 문구:`, `사용 가능한 <종류>:`)부터 끝까지를 `<문서에서 가져온 후보·문구 N자 생략>`으로 자릅니다.
  stdout JSON의 message는 그대로이고, 가리기에 실패해도 로그 기록은 오류를 삼킵니다.
- `office.py undo --list`: 되돌릴 수 없는 작업(새로 만든 파일 등)만 있으면 preview는 "지점 0개"인데 `count: 1`과 항목의 command를 실행하라는
  안내이던 문제. `count`·`shown`은 되돌릴 수 있는 지점만 세고(되돌릴 수 없는 작업 수는 `barriers`), 0개면 `VERIFIED: 0건.` 안내입니다.
- `reader clear`(실제 지우기): `applied`를 지우지 못한 캐시까지 넣어 만들고, `clear_done_*.csv`를 지우기 전에 쓰던 문제. `applied`는 지우고 다시 읽어도
  없는 것만, CSV는 지운 뒤 `removed`(True/False) 열과 함께 씁니다.
- 0건 결과 (R2): `ppt slide --action list`(슬라이드 0장), `outlook list`(`--unread` 포함, 0통), `outlook search`(0통)에 `count`가 없거나 directive가
  0건을 말하지 않던 문제. `count`와 `VERIFIED: 0건. ...`을 줍니다(`outlook`의 `count`는 `--limit` 안에서 가져온 메일 수).

### 문서
- 남은 할 일: 1.4.26에서 고친 0건 항목(office windows, info --all, chart list, undo --list)과 이번에 고친 outlook list를 뺐습니다.

## [1.4.26] - 2026-11-03

17차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1426.py`.

### 수정
- `ppt write` 다시 읽기 (1.4.25 회귀): 기대 글에서만 `**`, `__`를 빼고 다시 읽은 글에서는 빼지 않아, 첫 본문 줄이 코드 줄
  (`def __init__(self):`, `x**2`)이거나 제목에 `__`·`**`가 있으면(`## __init__ 설명`, `--title "__main__ 실행"`) 거짓 `UNVERIFIED`이던 문제.
  기대 글은 넣은 글 그대로(굵게 표시는 이미 뺀 글)이므로 공백만 빼고 비교합니다.
- `word write` 다시 읽기: `---`(구분선, YAML 머리말)로 시작하는 `.md`가 `'--'을 찾지 못함`으로 `UNVERIFIED`, 첫 표 행의 `\|`를
  잘못 다루던 문제. 마크다운은 해석한 블록에서 넣은 첫 글을 고르고(구분선은 건너뜀, 코드 줄과 표 칸은 글자 그대로),
  `--callout`은 넣은 글 그대로, `--heading`/`--title`은 굵게 표시만 빼고 비교합니다.
- `ppt write --text`: 펜스(```` ``` ````, `~~~`) 줄을 글자 그대로 넣고 안의 `**`, `__`를 굵게 표시로 읽던 문제. `.md`와 같이 펜스 줄은 넣지 않고
  안의 코드 줄은 글자 그대로 넣습니다(목록 기호·굵게 표시 해석 없음).
- `ppt write --file .md`: 3단계 이하 제목(`### **핵심** 요약`)을 굵게 표시를 빼지 않고 넣어 `**`가 보이던 문제.
- `excel write --file .md`(`load_table_file`): 따로 만든 `|` 해석이 파일의 모든 `|` 줄(두 번째 표, 코드 블록 안 줄)을 모으고 `|--|--|`
  정렬 줄을 데이터 행으로 남기며 `\|`에서 칸을 나누던 문제. word·ppt와 같은 `_runtime.parse_markdown_blocks`로 첫 번째 표만 씁니다.
- 0건 결과 (R2): `office.py windows`(열린 문서 0개), `office.py undo --list`(되돌릴 지점 0개), `excel/word/ppt info --all`(열린 문서 0개),
  `excel chart --action list`(차트 0개)에 `count`가 없거나 directive가 0건을 말하지 않던 문제. `count`와 `VERIFIED: 0건. ...`을 줍니다.
- `reader clear --dry-run`(과 실제 지우기): 대상이 20개를 넘으면 `planned`/`documents`를 잘라 `truncated: true`인데 전체 목록이 없던 문제 (R2).
  `<데이터 폴더>/results/clear_plan_<시각>.csv`(실제 지우기는 `clear_done_...`)에 전체 목록을 쓰고 `full_result_file`로 줍니다.
- 오류 로그: `--section`(word read, reader search: 문서 제목 글)과 `excel style --highlight`(`=완료:...` 셀 값)를 가리지 않던 문제.
  값을 `<N자 생략>`으로 남깁니다(`word --highlight`는 색 이름이라 그대로).

### 문서
- README: 마크다운 블록 해석 범위를 `word write`의 `.md`·`--text`, `ppt write`의 `.md`로 고침(`ppt write --text`는 줄 단위, 펜스 안만 글자 그대로).

## [1.4.25] - 2026-11-03

16차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1424.py` (1.4.25 부분).

### 수정
- 마크다운 코드 블록 (1.4.24 회귀): 코드 줄을 보통 단락으로 넘겨 `**`, `__`를 굵게 표시로 읽어 `def __init__(self):`가 `def init(self):`(굵게),
  `x**2`가 `x2`로 들어가고 VERIFIED이던 문제. 코드 줄은 새 블록 종류 `code`로 넘기고 Word·PowerPoint 모두 글자 그대로 넣습니다.
- 줄 첫머리의 인라인 코드(```` ```x``` ````)를 코드 블록 시작으로 보아 문서 끝까지 코드가 되던 문제 (백틱 펜스의 정보 글에는 백틱이 없음).
- `word write --file`: 코드 블록으로 시작하는 `.md`는 다시 읽기에서 펜스 줄(```)을 찾다가 UNVERIFIED가 될 수 있었음. 첫 코드 줄로 확인합니다.

## [1.4.24] - 2026-11-03

15차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1424.py`.

### 수정
- `find_chart` `[후보]`: 같은 이름 차트를 후보끼리만 세어, 제목으로 찾은 차트의 이름이 후보가 아닌 같은 시트 차트와 같으면
  `시트!이름` 후보를 다시 줘도 둘로 찾아져 또 `[후보]`가 나오던 문제. 그 시트의 모든 차트로 세어 이름이 겹치면 `시트!#번호`를 줍니다
  (후보마다 차트 하나로 찾아짐).
- `reader`: FileHeader 스트림이 없는 OLE 파일(이름만 `.hwp`인 `.doc`/`.xls` 등)이 `ERROR:` 내부 오류("file not found")로 끝나던 문제.
  STOP_RETRY "HWP 5.0 문서가 아닙니다"로 답하고, FileHeader를 읽다 난 OLE 오류는 손상된 파일로 봅니다.
- `excel style --highlight`: 값의 `"`가 조건부 서식 수식(`Formula1`)을 깨뜨려 COM 오류가 나던 문제(`""`로 겹쳐 씀).
  `포함:` 규칙은 SEARCH가 `~ ? *`를 와일드카드로 읽어 더 넓게 강조하던 것을 `~`를 붙여 글자 그대로 찾습니다.
- `excel write --append --formula`: 쓸 행을 아래로 옮기고도 원래 행 기준 수식을 그대로 써, 상대 참조가 원래 행을 가리키고 `VERIFIED`이던 문제.
  Excel `ConvertFormula`로 원래 첫 수식 칸 기준 R1C1로 바꾼 뒤 실제로 쓸 칸 기준 A1로 되돌려 씁니다(`$` 참조는 그대로).
  바꾸지 못하면 아무것도 쓰지 않고 STOP_RETRY와 `[추천] --range`. 가짜 Excel(`tests/fake_excel.py`)에 `ConvertFormula`를 넣었습니다.
- 마크다운 해석(`_runtime.parse_markdown_blocks`, word/ppt `write`): 펜스 코드 블록을 몰라 안의 `# 주석`이 제목 1, `- x`가 글머리표가 되고
  펜스 줄이 본문에 들어가며 코드 줄이 한 단락으로 합쳐지던 문제. 펜스(백틱 3개 이상 또는 `~~~`) 안은 줄마다 단락 하나로 글자 그대로
  (들여쓰기 유지, 빈 줄 생략), 펜스 줄은 넣지 않고, 닫지 않으면 끝까지 코드로 봅니다. `looks_like_markdown`도 펜스를 마크다운으로 봅니다.
- 마크다운 표: 머리글 다음의 정렬 줄이 `-` 1~2개(`:-`, `-:`, GFM에서 맞는 형식)이면 데이터 행으로 남던 문제, 칸 안의 `\|`가 칸을 나누던 문제.
  `\|`가 아닌 `|`에서만 나누고 `\|`는 글자 `|`로 바꿉니다.
- `ppt write --cards`: 남은 공간을 보지 않아 긴 `--text` 아래에서 카드가 슬라이드 밖(위 518pt, 슬라이드 높이 540pt)에 놓이고 `VERIFIED`이던 문제.
  카드·그림은 80pt도 남지 않으면 새 슬라이드에 넣고 `warnings`와 stderr `[WARN]`으로 알립니다. 다시 읽기에서 이번에 넣은 도형이
  슬라이드 경계 밖이면 `UNVERIFIED`(`_verify_inside`).
- `reader.parse_pages`(`read --pages/--units`, ppt `--slides`): 범위를 문서 크기로 줄이기 전에 다 돌아 `--units 1-999999999`가 멈추던 문제. 먼저 줄입니다.

## [1.4.23] - 2026-11-03

14차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1423.py`.

### 수정
- `outlook search`: DASL 필터를 만들 때 `_`를 공백으로 바꾸고 `%`를 지워 `견적서_v2`, `10% 할인` 같은 검색어가 절대 맞지 않고
  `VERIFIED: 일치하는 메일이 없습니다`로 끝나던 문제. 작은따옴표만 겹쳐 쓰고 `%`, `_`는 그대로 둡니다(와일드카드지만 자기 자신과도 맞아
  빠지는 메일이 없음). 그런 검색어는 제목·본문·보낸사람을 다시 읽어 글자 그대로 들어 있는 메일만 남깁니다(대소문자 무시).
- `office.py transfer --chart`: 차트 그림 파일 이름 `chart_<차트 이름>`을 정리하지 않아, 제목 `1/4분기 매출`로 만든 차트(이름도 그 제목)는
  없는 폴더 경로가 되어 export·redraw가 실패하던 문제. `_runtime.safe_filename`을 새로 두고 `resolve_output_path`가 기본 이름
  (`default_stem`)을 늘 정리합니다(`\ / : * ? " < > |`, 제어 문자 -> `_`, 끝의 점·공백 제거, `CON`·`NUL`·`COM1` 같은 예약 이름 앞에 `_`).
  시트 이름으로 만드는 기본 파일 이름(`excel read --save`, `excel export`)도 같습니다. `outlook._safe_name`도 이 함수를 씁니다.
- `find_chart`: `시트!이름`을 첫 `!`에서만 나눠
  - 시트 `A`와 `A!B`가 함께 있으면 `A!B`의 차트를 찾지 못했고(`'A!B'!차트`도 마찬가지),
  - 다른 시트에 이름이 글자 그대로 `Sheet1!x`인 차트가 있으면 Sheet1의 `x`를 지우거나 Sheet1에 차트가 없으면 "먼저 만드세요"라고 하던 문제.
  `'시트'!이름`(작은따옴표 안의 `''`는 `'`) 형식은 그 시트에서만 찾고, 따옴표 없는 형식은 `!` 자리마다 나눠 찾아지는 나눔만 씁니다.
  글자 그대로의 차트 이름으로도 찾아 둘 다 찾아지면 `[후보]`. 시트·차트 이름에 `!`나 `'`가 있으면 `[후보]`는 `'시트'!이름` 형식입니다.
  `--sheet`가 앞부분의 시트와 다르면 두 시트를 말하는 STOP_RETRY, 차트가 없다는 안내도 시트 이름을 말합니다.
- `find_chart`: 한 시트에 이름이 같은 차트가 둘이면 `[후보]`가 하나로 합쳐져 첫 차트를 말없이 골랐던 문제. `[후보] 시트!#번호`를 내고
  `#번호`(그 시트의 몇 번째 차트)를 받습니다. `chart --action delete`는 같은 이름 차트가 하나 줄었는지로 확인합니다(남은 같은 이름 때문에
  `UNVERIFIED`가 되지 않게).
- `reader`: 손상된 HWPX(XML 오류), HWP(압축 해제 오류, OLE 구조 오류)가 `ERROR:` 내부 오류와 traceback으로 끝나던 문제.
  `load_document_status`에서 STOP_RETRY "손상된 파일이라 읽을 수 없습니다"로 바꿉니다(내부 오류로 기록하지 않음).
- `office.py undo --pick 0`이 `args.pick or 1` 때문에 가장 최근 지점을 되돌리던 문제. 0 이하는 아무것도 바꾸지 않고 STOP_RETRY.
- `outlook list/search --limit`에 상한이 없던 문제. 200을 넘으면 200통으로 줄이고 `warnings`, 0 이하는 STOP_RETRY.

## [1.4.22] - 2026-11-03

13차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1421.py` (1.4.22 부분).

### 수정
- `find_chart`: 차트 `[후보]`의 `시트!이름`을 `--sheet`와 함께 주면 찾지 못하던 문제(앞부분이 그 시트와 같으면 받음).
- `find_chart`: 차트 이름에 `!`가 있으면(`매출 증가!`) `시트!이름` 후보를 다시 줘도 찾지 못하던 문제. 왼쪽부터 `!` 자리마다 나눠
  앞부분이 있는 시트 이름이고 뒷부분이 비어 있지 않은 첫 경우를 씁니다.
- SKILL.md 결과 처리 표: 문서 `[후보]`를 늘 `--target`으로 안내하던 문구를, message가 말한 옵션(`--target`, `--to-target`,
  `--from-target`, `--name`)을 쓰라고 고침.

## [1.4.21] - 2026-11-03

12차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1421.py`.

### 수정
- 시트를 주지 않은 차트 찾기(`excel.find_chart`: `office.py transfer --chart`, `excel.py chart --action delete --name`)가 모든 시트에서
  이름 또는 제목이 맞는 첫 차트를 골라, 다른 시트에 같은 이름/제목 차트('매출')가 있으면 엉뚱한 차트를 옮기거나(VERIFIED) 지우던 문제.
  `excel chart`의 `next_commands`에 `--sheet`가 없어 SKILL 순서대로 해도 생겼고, 다른 차트의 제목이 이 차트의 이름과 같으면 가려졌습니다.
  - 이름이 같은 차트를 제목이 같은 차트보다 우선합니다. 그래도 둘 이상이면 하나를 고르지 않고 STOP_RETRY와 `[후보] 시트!이름`을 냅니다.
  - 차트 이름 자리에 `시트!이름`을 받습니다(앞부분이 있는 시트 이름일 때. `--sheet`와 같음).
  - `excel chart`의 `next_commands`(transfer)에 `--sheet <차트가 있는 시트>`를 넣습니다.
- 이름이 같은 열린 문서가 여럿일 때(`_office.match_open`) 안내가 늘 `--target`이던 문제. 그 명령의 옵션을 말합니다:
  transfer 받는 쪽 `--to-target`, 보내는 쪽 `--from-target`, `office.py focus`는 `--name` (`match_open(option=...)`).

## [1.4.20] - 2026-11-02

11차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1420.py`.

### 수정
- `office.py transfer --to word,ppt`(차트·범위): 워드에 먼저 넣은 뒤 PowerPoint가 열려 있지 않으면 `STOP_RETRY: 실행하지 않았습니다 ...
  같은 명령에 --target ...`으로 끝나, 워드는 이미 바뀌었는데 다시 실행하면 워드에 두 번 들어가고 transfer에 없는 `--target`을 안내하던 문제.
  `excel chart`의 `next_commands`가 늘 `--to word,ppt`라 워드만 열려 있으면 SKILL 순서대로 해도 생겼습니다.
  - 넣기 전 확인: 받는 앱이 둘 이상이면 무엇이든 넣기 전에 앱마다 받는 문서(활성 문서 또는 `--to-target`)와 앱별 인자(`--anchor`, `--slide`)를
    확인합니다. 못 찾은 문서가 있으면 아무것도 바꾸지 않고 STOP_RETRY(`error_type: target_not_open`)와 찾은 앱만으로 바꾼 같은 명령
    `[추천] python ... --to word`를 냅니다. 하나도 못 찾으면 문서를 열어 달라고 하거나 `--to`에 앱 하나와 `--to-target`을 쓰라고 안내합니다
    (`--to-target`은 모든 받는 앱에 함께 쓰임).
  - 넣은 뒤 일부 실패(늦게 드러난 대상 문제, COM 오류 등): 이미 다른 앱에 넣었으면 오류로 끝내지 않고 `UNVERIFIED`와
    `succeeded`/`failed`/`failures`로 답합니다(다시 실행하면 넣은 앱에 또 들어간다고 알림). 아무 앱에도 넣지 않았으면 지금처럼 오류.
- transfer에서 열린 문서가 없을 때 안내가 이 명령에 없는 `--target` 대신 받는 쪽은 `--to-target`, 보내는 쪽은 `--from-target`을 말합니다
  (`open_target(option=...)`, `_office.no_open_doc`).

### 기능
- `excel chart`의 `next_commands`는 문서가 열린 받는 앱만 `--to`에 넣습니다(워드만 열려 있으면 `--to word`). 둘 다 열렸거나 둘 다 없거나
  확인하지 못하면 지금처럼 `--to word,ppt`(transfer가 넣기 전에 다시 확인).

## [1.4.19] - 2026-11-02

10차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1418.py` (1.4.19 부분).

### 수정
- `word replace`: 바꿀 글이 찾을 글을 품고 255자를 넘으면, 바꾼 뒤 다시 세는 단계에서 Word 찾기가 오류 5854로 끝나 문서는 바뀌었는데
  `ERROR:`로 보고하던 문제. 255자를 넘는 글은 Word 찾기 대신 각 스토리 글에서 직접 셉니다.
- 찾을 글과 `--anchor`의 255자 한도를 1.4.18의 `^` -> `^^` 바꿈 뒤 길이로 셉니다 (`^`가 많은 글이 한도 안으로 보였다가 COM 오류가 나지 않게).
  가짜 Word도 255자를 넘는 검색어를 오류로 거절합니다.

## [1.4.18] - 2026-11-02

9차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1418.py`.

### 수정
- `ppt replace`: 1.4.17에서 찾은 곳을 파이썬 글자 위치로 PowerPoint에 넘겼는데, PowerPoint는 UTF-16 단위로 세어
  앞에 이모지(BMP 밖 글자)가 있으면 뒤쪽 일치를 건너뛰던 문제(다시 읽기에서 UNVERIFIED). PowerPoint에는 UTF-16 위치를 넘기고
  다음 찾기는 파이썬 위치로만 이어 갑니다.
- `word replace`/`--anchor`: Word 찾기는 와일드카드가 꺼져 있어도 `^`를 특수 코드로 읽어 `감사합니다^^`가 `감사합니다^`와 맞거나
  `10^6`이 COM 오류가 되던 문제. 찾는 글의 `^`를 `^^`로 바꿔 글자 그대로 찾습니다. 가짜 Word(`tests/fake_word.py`)도 `^` 코드를 흉내 냅니다.

## [1.4.17] - 2026-11-02

8차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1417.py`.

### 수정
- `excel replace --find ""`(또는 `--map {"": "X"}`, 다른 키와 섞인 경우, `--all-sheets` 포함): 빈 글을 찾으면 값이 있는 모든 칸의 글자 사이에
  바꿀 글이 끼어드는데 `VERIFIED`로 보고하던 문제. `word replace`/`ppt replace`도 같은 확인이 없었습니다. 셋 모두 찾을 글이 빈 문자열이면
  아무것도 바꾸기 전에 STOP_RETRY와 `write` 안내(빈 칸 채우기는 `excel.py write --range B2:B10 --value 0`)를 냅니다. 공백만 있는 찾을 글은 그대로 허용합니다.
- `office.py transfer --to excel`로 새 파일을 만들 때 기본 시트가 여럿이면(Office 2007/2010 기본 `Sheet1`~`Sheet3`, `SheetsInNewWorkbook`)
  `--sheet Sheet1`이 '이미 있음'으로 멈춰 빈 새 파일만 남던 문제(1.4.16은 기본 시트가 1개일 때만 고침). 방금 만든 새 파일이면 기본 시트 수와 관계없이 그 빈 시트에 받습니다.
- `ppt replace`/`word replace`에서 대소문자만 다른 치환(`--find iphone --replace iPhone`)이 이미 `iPhone`인 곳도 바꾼 곳으로 세어 `replaced` 2, `VERIFIED`,
  바뀐 것 없는 되돌리기 지점을 남기던 문제. 또 `ppt replace --find iPhone --replace iPhone`이 `iPhone and IPHONE`에서 1곳만 바꾸고 2곳으로 세던 문제.
  찾은 곳마다 글을 바꿀 글과 비교해 글자 그대로(대소문자까지) 같으면 건드리지 않고 `unchanged_same_text`로 세고, 실제로 바뀐 곳만 `replaced`로 셉니다
  (서식 옵션이 있으면 지금처럼 모두 바꿈). `ppt.replace_in_range`는 (바꾼 수, 같은 수)를 돌려줍니다.
- 다시 읽어 확인(R1): `ppt replace`/`word replace`는 바꿀 글이 대소문자까지 그대로 있는지 확인합니다(대소문자만 바꾸는 치환도 확인).

## [1.4.16] - 2026-11-02

7차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1416.py`.

### 수정
- `ppt replace --find A --replace A`(서식 없음, `--match-case` 없음): 대소문자만 다른 곳(`IPHONE`, `iphone`)이 있으면 PowerPoint가
  찾는 글 그대로 바꿔 써 대소문자가 통일되는데, 1.4.15는 "바꿀 글이 찾는 글과 같다"며 바꾸지 않고 VERIFIED로 보고하던 문제.
  대소문자가 모두 같을 때만 바뀐 것 없음으로 봅니다.
- `office.py transfer --to excel`로 새 파일을 만들 때 `--sheet`가 새 통합 문서의 기본 시트 이름(`Sheet1` 등)과 같으면 '이미 있음'으로 멈춰
  빈 새 파일만 남던 문제. 방금 만든 파일의 유일한 기본 시트에 받습니다.
- `range hide`/`unhide`(일부만 숨긴 범위에서 `Hidden`이 섞임 대신 False일 수 있음)와 `range clear` 전체·서식(사용 범위 밖 칸에도 데이터 유효성·메모가
  있을 수 있음)을 "바뀐 것 없음"으로 건너뛰면 실제 변경을 놓칠 수 있어, 이들은 늘 바꾼 것으로 봅니다(되돌리기 지점이 하나 더 남을 수 있음).
  `clear --type contents`는 값이 있는 칸으로 계속 판단합니다.

## [1.4.15] - 2026-11-02

6차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1415.py`.

### 수정 (엑셀 시트 이름)
- 엑셀이 예약한 시트 이름 `History`(대소문자 무관): 시트를 만든 뒤 이름을 붙이다 실패해 빈 시트가 남던 문제(`sheet add/rename/copy`, `write --new-sheet`,
  `office.py transfer --sheet`). 아무것도 바꾸기 전에 STOP_RETRY와 `[추천] History_`(있는 시트와 겹치면 번호)를 냅니다(`_office.check_sheet_name`).
  가짜 Excel(`tests/fake_excel.py`, `fake_office.py`)도 실제 Excel처럼 쓸 수 없는 이름(`History` 포함)을 `Name`에서 COM 오류로 거절합니다.
- `office.py transfer --sheet`: 받는 문서의 시트 이름을 모른 채 확인해, `[추천]` 이름이 이미 있는 시트와 겹치면 안내대로 1번 다시 실행해도 "이미 있음"으로 막히던 문제.
  문서를 연 뒤 시트 이름을 알고 나서, 아무것도 바꾸기 전에 확인합니다. 새로 만들 파일(없는 전체 경로)은 파일을 만들기 전에 확인합니다.

### 수정 (되돌리기 지점, 1.4.14 보완)
- 도구가 연 닫힌 파일: 휘발성 수식(NOW, RAND 등)이 있으면 열기만 해도 저장 안 된 상태라, 바꾼 것 없이 성공한 실행도 사용자 파일을 다시 저장하고
  바뀐 것 없는 되돌리기 지점을 남기던 문제. 명령이 바꾼 것(`t.changed`)이 있을 때만 저장합니다.
- 바꾼 것 없이 성공했는데 바꾼 것으로 표시되어 지점을 남기던 명령(다음 `undo`가 바뀐 것 없는 이 지점으로 되돌림):
  - `excel replace`/`word replace`/`ppt replace`에서 바꿀 글이 찾는 글과 같고(`--map {"A": "A"}` 포함) 서식이 없는 경우. "N곳 처리"로 보고하던 것을
    0곳/0칸 바뀜(`changed_cells`/`replaced` 0, 찾은 수는 `unchanged_same_text`)으로 답합니다. `excel replace`의 서식 없는 수식 칸도 바꾼 칸으로 세지 않습니다.
  - `ppt slide --action move`를 같은 자리로(`--slide 2 --to 2`), `ppt notes --overwrite`를 같은 노트로.
  - `excel range`: 숨긴 것 없는 `unhide`, 이미 숨긴 `hide`, 병합 없는 `unmerge`, 이미 그 범위로 병합된 `merge`, 고정 없는 `unfreeze`, 같은 `width`/`height`,
    사용 범위 밖(`Intersect(UsedRange)`)이거나 값이 없는(`--type contents`) `clear`, 같은 자리로 `copy`/`move`. COM 속성으로 확인하고 알 수 없으면 바꾼 것으로 봅니다.
  바꾸지 않은 결과는 "바뀐 것 없음"(`unchanged: true`)으로 알립니다.

### 수정 (PPT)
- `ppt replace --find X`만 주면(바꿀 글·서식 없음) 아무것도 바꾸지 않아야 하는데 대소문자만 다른 글을 바꾸고 "N곳 처리"로 보고하던 문제. `excel`/`word replace`와 같게 STOP_RETRY.

## [1.4.14] - 2026-11-02

5차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1414.py`.

### 수정 (되돌리기 지점)
- 바꾼 것 없이 성공한 쓰기(`excel replace` 0곳, `chart --action delete --all`에 차트 0개, 이미 있는 시트 `sheet add`, `sheet select`, 빈 슬라이드 `ppt slide clear`)가
  되돌리기 지점을 남겨, 다음 `office.py undo`가 바뀐 것 없는 이 지점으로 되돌리고 `VERIFIED`로 답하던 문제(실제로 바꾼 작업은 그대로 남음).
  성공한 실행에도 "바뀐 것 없음" 기준(도구가 연 파일은 디스크 파일 그대로, 열린 문서는 `t.changed` 없고 저장 상태 그대로)을 적용해 지점을 지웁니다.
  `OfficeTarget.set_focus(..., changed=False)`로 바꾸지 않은 명령을 표시하고, 화면 이동 없이 바꾸던 `excel sheet --action delete`는 `t.changed`를 표시합니다.

### 수정 (엑셀)
- 쓸 수 없는 시트 이름(`/ \ ? * [ ] :`, 31자 초과, 빈 이름, 앞뒤 `'`): 시트를 만들거나 복사한 뒤 이름을 붙이다 실패해 빈 시트·복사본이 남던 문제
  (`write --new-sheet`, `sheet add/rename/copy`, `office.py transfer --sheet`. `transfer`는 31자에서 말없이 자르던 것도).
  아무것도 바꾸기 전에 STOP_RETRY와 `[추천]`(쓸 수 없는 글자를 `_`로 바꾸고 31자로 자른 이름, 있는 시트와 겹치면 번호)을 냅니다(`_office.check_sheet_name`).
  `sheet copy`의 기본 이름은 31자를 넘지 않게 `<시트 28자> 복사`.
- 닫힌 xlsx의 큰 범위 읽기: 파일에 적힌 크기가 서식만 있는 먼 칸(`Z300000`)까지면 데이터가 작아도 STOP_RETRY이던 문제. 그 크기로도 상한을 넘으면 값을 훑어 실제 끝까지 줄입니다.

### 수정 (reader)
- 스캔 PDF `read --save`: 모든 쪽이 스캔이면 `VERIFIED`와 빈 파일을 다른 문서에 넣는 `next_commands`를 주던 문제. `UNVERIFIED`(OCR 필요), 0개 저장이거나 스캔 쪽 글뿐이면 `next_commands` 없음.
- `read --pages`로 도장·Bates 번호만 남은 스캔 쪽만 고르면 `VERIFIED`이던 문제. 고른 쪽이 모두 `pages_without_text`이면 읽기·저장 모두 `UNVERIFIED`.
- 색인(`index.json`)이 UTF-8로 풀리지 않으면(1.4.13 회귀) `read_text_file`의 입력 오류가 그대로 나가 `clear --all`까지 모든 reader 명령이 실패하던 문제.
  바이트로 읽어(`retry_fs`) UTF-8로 풀고, 풀리지 않으면 손상으로 보고 캐시 문서로 다시 만듭니다. `IndexReadError`는 끝내 읽지 못한 `OSError`일 때만.
- `search` directive가 늘 `--save "근거.xlsx"`를 안내해 `next_commands`의 `.md` 저장 명령과 다르던 문제. 결과가 잘렸을 때만 `next_commands`의 저장 명령을 가리킵니다.

## [1.4.13] - 2026-11-02

4차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1413.py`.

### 수정 (되돌리기 지점, 1.4.12 정리 보완)
- `excel write --new-sheet`: 시트를 먼저 만들고 수식·범위를 확인해, 입력 오류로 멈추면 새 시트만 남고 되돌리기 지점은 지워지던 문제
  (안내대로 다시 실행하면 "이미 있음"으로 막힘). 칸 주소·수식 확인과 쓸 내용 계획을 모두 마친 뒤 쓰기 직전에 시트를 만듭니다.
- `excel style --clear --highlight 잘못된값`: 서식을 지우고 테마를 적용한 뒤에 `--highlight`를 확인해, 지운 서식이 되돌리기 지점 없이 남던 문제.
  `--highlight`는 문서를 열기 전에 확인합니다. 같은 점검으로 `excel replace --all-sheets`는 모든 시트의 범위를 바꾸기 전에 확인하고,
  `outlook new/reply/forward`는 본문 파일·첨부 파일을 메일을 만들기 전에 확인합니다.
- 이번 실행의 지점을 지우는 조건: 입력 오류이기만 하면 지우던 것을, 문서를 바꾸지 않았을 때만 지웁니다.
  열린 문서는 바꾼 표시(`t.changed`)가 없고 저장 상태가 그대로일 때(원래 저장 안 된 문서는 입력 오류일 때만),
  도구가 열거나 만든 파일은 디스크 파일이 그대로일 때(오류나 `Save()` 실패로 저장하지 않고 닫은 경우).
- 오래된 지점 정리(문서별 5개)를 명령 실행 전에 해, 거절된 실행이 진짜 옛 지점 하나를 밀어내던 문제. 명령이 문서를 바꾼 뒤에만 정리합니다
  (`save_snapshot`/`save_barrier`의 `rotate=False`, `rotate_point`).

### 수정 (reader)
- 색인(`index.json`)을 읽다 일시적 오류(공유 위반, 백신)가 나면 빈 색인으로 보고 그 위에 저장해 항목이 사라지던 문제(120개 -> 1개).
  `retry_fs`로 다시 읽고, 끝내 못 읽으면 색인을 바꾸지 않고 STOP_RETRY(`index`, `search` 등. 파일 읽기는 계속하고 색인 넣기만 건너뜀).
  내용이 손상(JSON·인코딩 오류)된 경우에만 캐시 문서 파일로 색인을 다시 만들고 `recovered`로 알립니다(이전에는 빈 색인에서 다시 시작).
- PDF 스캔 쪽: 스캔 위에 찍은 도장("RECEIVED"), 한 쪽짜리 스캔의 Bates 번호(머리글 빼기는 2쪽 이상)처럼 짧은 글이 남은 쪽을 놓치던 문제.
  그림이 쪽 면적의 40% 이상을 덮고 남은 글이 40자보다 적은 쪽을 `pages_without_text`로 표시합니다. 보이지 않는 글자 층(OCR 결과)이 있는 쪽은 제외,
  모든 쪽이 스캔 쪽이면 `ocr_needed`(`PARSER_VERSION` 8).
- 스캔 이미지 PDF의 `read`: "읽을 글이 없습니다 ... 다시 실행하지 마세요"(`VERIFIED`)로 답하던 문제. `ocr_needed`이거나 고른 쪽이 스캔 쪽뿐이면
  `UNVERIFIED`(스캔 이미지라 글자 층이 없음, OCR 필요). Bates 번호만 읽힌 경우도 같습니다.
- `search --full`의 저장 안내가 엑셀 칸 상한(32,000자)보다 긴 단락이 있어도 `--save 근거.xlsx`(반드시 `UNVERIFIED`)였던 문제. 그런 단락이 있으면 `근거.md`.
- 오래된 색인 잠금 지우기의 경쟁: 두 실행이 함께 지우면 늦은 쪽이 먼저 쪽의 새 잠금을 지울 수 있던 문제. 고유 이름(`index.lock.stale.<pid>`)으로 옮긴 뒤 지우고,
  옮긴 것이 새 잠금이면 되돌려 놓습니다. 잠금 대기를 15초로(잠근 쪽의 `retry_fs` 재시도 약 10.5초보다 길게).

### 수정 (outlook·엑셀·공통)
- `outlook read/current --save-attach`: 첨부 저장에 실패해도 `VERIFIED`였던 문제. 저장 실패가 있거나 저장한 파일이 없으면 `UNVERIFIED`(저장 뒤 파일이 있는지 다시 확인).
- 닫힌 xlsx를 아주 큰 범위(`A1:XFD1048576`)로 읽으면 빈 칸을 채우다 메모리 부족(`ERROR`)이던 문제. 200만 칸을 넘으면 데이터가 있는 끝까지만 읽고 `recovered`로 알리며,
  그래도 넘으면 STOP_RETRY와 `[추천] --range`.
- 기본 저장 이름(`<이름>_<시각>`)이 같은 초에 겹치면 앞 결과를 말없이 덮어쓰던 문제. 번호를 붙인 새 이름(`_2`)으로 저장하고 `recovered`로 알립니다.
- 엑셀 읽기 미리보기가 첫 행이 표시 상한을 넘어 뺀 열까지 "긴 값 N칸 잘림"으로 세던 문제. 긴 값과 뺀 열을 따로 알립니다("뒤쪽 N열 생략").
  `_runtime.clamp_rows`는 (표, 행 잘림, 잘린 칸 수, 뺀 열 수)를 돌려줍니다.

## [1.4.12] - 2026-11-02

3차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1412.py`.

### 수정 (reader)
- `index`를 동시에 두 번 실행하면(다른 폴더) 색인(`index.json`) 항목이 사라지던 문제. 색인 읽기-고치기-쓰기를 `cache/index.lock`으로 잠급니다
  (O_CREAT|O_EXCL, 최대 10초 대기, 60초보다 오래된 잠금은 지움). `index`는 끝에 색인을 다시 읽어 빠진 문서를 다시 넣고(`recovered`) 확인한 뒤에만 `VERIFIED`,
  다시 넣지 못하면 `index_missing`과 `UNVERIFIED`. `index --all`은 색인에 없는 캐시 파일 중 원본이 있는 문서도 다시 넣습니다(`readded_to_index`).
- PDF 글자 없는 쪽: 스캔 쪽에 찍힌 쪽 번호(Bates)가 있으면 놓치고, 쪽마다 작은 로고가 있는 빈 쪽은 "OCR 필요"로 알리던 문제.
  이제 머리글·쪽 번호를 뺀 뒤 글이 없고 그림이 쪽 면적의 40% 이상을 덮는 쪽만 `pages_without_text`(OCR 글자 층이 있는 쪽은 제외, `PARSER_VERSION` 7).
- 캐시 파일이 없거나 손상됐는데 원본도 없는 문서: `index --all`을 안내해 같은 안내가 반복되던 문제. 원본 없음으로 세고(`missing_file_count`) `clear --missing --dry-run` 명령.
- `search --save x.xlsx|x.csv --full`이 `--full`을 무시하고 문맥을 3,000자에서 말없이 자르던 문제. `--full`이면 단락 전체(`.xlsx`는 셀 상한 확인), 아니면 잘린 행 수를 `context_clamped`와 directive로 알립니다.
  JSON `--full`의 `text`가 2,000자에서 잘려도 표시가 없던 문제(`text_truncated`, `text_chars`, `truncated: true`, 저장 명령에 `--full`).
- `search --top`, `--context`에 상한이 없어 출력이 끝없이 커지던 문제. `--top` 50, `--context` 500까지(`warnings`, 더 많은 결과는 `--save`).
- `clear --file/--folder`가 `~`를 풀지 않던 문제(`search`/`read`와 같게).

### 수정 (엑셀·워드·PPT·공통)
- 없는 시트: `--new-sheet`가 없는 명령(`excel replace/style/range/chart`)에도 "`--new-sheet`를 붙여 실행"을 안내해 인자 오류가 나던 문제. 그 명령들은 "찾을 수 없음"(`[후보]`/`[추천]`), `excel write`만 새 시트를 물음.
- 되돌리기: 쓰기 전에 멈춘 실행(입력 오류, `needs_user_decision`)도 스냅샷을 남겨, 다음 `office.py undo`가 바뀐 것 없는 그 지점으로 되돌리던 문제. 그 실행이 만든 지점을 지웁니다.
- 출력 크기: 엑셀 읽기의 첫 행, 찾은 칸 값, 슬라이드 글이 아무리 커도 그대로 들어가던 문제. 칸·찾은 값은 500자(`value_truncated`), 슬라이드 글은 약 8,000자(`text_truncated`)에서 자르고 `truncated: true`와 `--save` 명령.
  워드 단락·PPT 노트가 약 8,000자에서 잘려도 `truncated: false`였던 문제(`text_truncated`, `truncated_reason`). `_runtime.clamp_rows`는 (표, 행 잘림, 잘린 칸 수)를 돌려줍니다.
- 엑셀 칸 주소 `A0`, `0:3`, `A1048577`, `XFE1`을 받아 닫힌 파일에서 그대로 읽던 문제(STOP_RETRY).
- 도구가 쓴 CSV를 다시 읽을 때 구분자를 추측해 `' | '` 같은 칸이 있으면 행 수가 틀려 `UNVERIFIED`가 되던 문제(`check_saved_file(..., delimiter=",")`).
- 닫힌 xlsx에서 범위를 직접 준 읽기(`A1:C5`)가 데이터 뒤의 빈 행을 빼 `rows`가 열린 문서와 달랐던 문제(준 범위만큼 채움).

## [1.4.11] - 2026-11-02

2차 점검에서 찾은 오류 수정 (1.4.10 회귀 포함). 회귀 테스트: `tests/test_review_v1411.py`.

### 수정 (reader)
- PDF 글자 없는 쪽(1.4.10 회귀): 머리글·쪽 번호를 뺀 뒤에 세어 빈 쪽·간지까지 "OCR 필요"로 알리고, 그 문서의 모든 검색이 `UNVERIFIED`가 되던 문제.
  이제 빼기 전에도 글이 없고 그림이 있는 쪽만 `pages_without_text`(`PARSER_VERSION` 6). `search`는 `--pages`/`--section` 안의 쪽만 `not_searched`에 넣고, `read`는 읽으려는 범위 안의 쪽만 알립니다.
- 이전 형식 캐시: 검색 때 다시 읽는 상한(20개)을 넘어 시도하지 않은 문서도 "다시 읽지 못함"으로 알리던 문제(이유를 나눔), `skipped`와 `stale`에 두 번 세어 "옛 내용으로 검색"이라고 잘못 알리던 문제(`skipped`에만),
  원본이 없는 이전 형식 캐시에 `index --all`을 안내해 같은 안내가 반복되던 문제(`clear --missing --dry-run` 명령). 할 명령(`index --all`, `clear --missing`)을 `UNVERIFIED` directive 문구에도 넣습니다.
- `read --save x.xlsx`, `search --save x.xlsx --full`: 셀당 32,000자에서 잘라 저장하고도 `VERIFIED`였던 문제. 잘린 행이 있으면 `UNVERIFIED`("셀 글자 상한 32000자 초과: N행 잘림 (.md로 저장하면 전체)"). 긴 글이 잘린 `read`의 저장 명령은 `.md`로.
- `read`: `truncated_reason`이 저장할 때만 문자열이던 것을 늘 목록으로. 글자 없는 쪽 안내는 잘림이 아니므로 `notes`로. 엑셀 시트 행 상한에서만 잘렸으면 저장해도 그 뒤가 없으므로 `--save` 명령과 "전체 저장" 안내를 주지 않고 상한을 알립니다.

### 수정 (PPT·엑셀·워드)
- `ppt.py write --slide N --title`(1.4.10 회귀): 제목 개체틀이 없는 슬라이드의 본문 첫 줄을 기존 제목으로 보고 `needs_user_decision`을 내던 문제. 제목 개체틀의 글만 기존 제목으로 봅니다(개체틀이 없으면 확인 없음).
- `ppt.py write --title 제목 --file 내용.md`: `.md` 첫 절의 제목이 `--title`과 같으면 본문 첫 줄에 같은 제목이 또 들어가던 문제(공백·대소문자 무시 비교).
- `excel.py`/`word.py`/`ppt.py read --save`: 저장한 파일을 다시 열지 않고 `VERIFIED`였던 문제(R1). 엑셀 CSV는 빈 행을 뺀 행 수까지 비교, 확인하지 못하면 `UNVERIFIED`.
- 닫힌 xlsx 읽기: 작은 범위(`A1:B2`)에도 시트 전체를 훑어 데이터 범위를 세던 것을 범위 생략·열/행 전체(`B:B`, `5:7`)일 때만, `Sheet!A1`의 시트로 바꾼 뒤에 셉니다.

### 수정 (공통·outlook)
- 오류 로그 가리기: 짧은 이름(`-q` = `--query`)과 줄여 쓴 이름(`--que`, `--bod=`)의 값이 그대로 남던 문제, `--map`(찾아 바꿀 JSON) 추가. 모든 파서가 줄여 쓴 옵션을 받지 않습니다(`allow_abbrev=False`).
- `check_saved_file`: `.pptm`도 python-pptx로 다시 열어 확인(`.docm`은 python-docx가 열지 못해 그대로 `UNVERIFIED`).
- 쓰기 `--target`(1.4.10 W8): `%USERPROFILE%\Desktop\새.xlsx`처럼 환경 변수·`~`가 든 전체 경로를 파일 이름으로 보고 거부하던 문제(푼 뒤에 판단). 없는 전체 경로에 열린 문서 전부를 `[후보]`로 주던 문제(파일 이름만 줄 때만 `[후보]`).
- `outlook.py current --save`: 저장 확인 결과(`save_problem`)를 무시하고 늘 `VERIFIED`였던 문제(`read`와 같은 응답).
- `outlook.py --attach`: 폴더 이름에 `;`가 든 경로는 하나일 때만 첨부되던 문제. 나눈 조각이 있는 파일이 아니면 뒤 조각을 이어 붙여 찾습니다.

## [1.4.10] - 2026-11-01

전체 검토에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1410.py`.

### 수정 (엑셀)
- `excel replace --range '요약'!A1:D9`: 읽기는 `--range`의 시트, 쓰기는 지금 시트에 하고 `VERIFIED`였던 문제. 이제 범위의 시트에서 읽고 씁니다. `--all-sheets`와 시트 이름이 든 `--range`는 함께 쓸 수 없음(STOP_RETRY).
- `excel replace --find X`만 주면(바꿀 글·서식 없음) 아무것도 바뀌지 않는데 "N개 셀 처리" `VERIFIED`였던 문제. 이제 STOP_RETRY(`word replace`도 같게).
- `excel range --action merge`: 왼쪽 위가 비고 다른 칸 하나에만 값이 있으면 확인 없이 병합해 그 값이 지워지던 문제. 왼쪽 위 밖의 값이 하나라도 있으면 `needs_user_decision`.
- 덮어쓰기 확인(`write`, `range copy/move`, `merge`): 빈 글("")을 내는 수식 칸을 빈 칸으로 보고 확인 없이 덮어쓰던 문제. 수식도 확인합니다(`--append`의 마지막 행 찾기도).
- `excel write --append --formula/--value`: 첫 열만 보고 마지막 행을 정해 옆 열의 데이터를 덮어쓰던 문제. 범위의 열 수(또는 수식 개수)만큼 보고, `--append`에도 덮어쓰기 확인을 합니다.
- `excel range --action copy`: 상대 참조 수식은 값이 달라져 늘 `UNVERIFIED`였던 문제. 수식(R1C1)으로 비교합니다.
- `excel range --action autofit`: 확인하지 않고 `VERIFIED`(verified false)였던 것을 `UNVERIFIED`로.
- 닫힌 xlsx·CSV 읽기를 열린 문서와 같게: 거꾸로 준 범위(`C5:A1`)가 빈 결과였던 문제, 크기 정보(`<dimension>`)가 없는 파일에서 `B:B`가 `B:D`로·`5:7`이 끝 열 없이 읽히던 문제(실제 데이터 범위를 셈), 범위를 생략하면 A1부터 읽어 `rows`가 열린 문서와 달랐던 문제(데이터가 시작하는 칸부터).

### 수정 (워드·PPT)
- `word replace`: 문구당 5,000곳 상한에 닿으면 확인 없이 `VERIFIED`였던 것을 `UNVERIFIED`로. `--match-case` 없이 대소문자만 바꾸면(`office` -> `Office`) `UNVERIFIED`로 나오던 문제.
- `ppt.py write --slide new --title 제목 --file 내용.md`: `.md`가 있으면 `--title`이 버려지던 문제. 첫 슬라이드 제목으로 넣고 첫 절의 제목은 본문 첫 줄로.
- `ppt.py write --slide N --title`: 기존 슬라이드의 다른 제목을 확인 없이 바꾸던 문제. `needs_user_decision`(노트와 같은 방식), 사용자가 원하면 새 인자 `--overwrite`. 새 슬라이드는 해당 없음.
- `ppt.py read --slides 1,5,9`가 잘리면 이어 읽기 명령이 `2-9`였던 문제. 고른 슬라이드 중 남은 번호(`5,9`)로.

### 수정 (공통·office·outlook)
- `_runtime.ok()`가 `verified=False`를 받으면 내부 오류(ValueError). 확인하지 못한 결과는 `ok_unverified`로 무엇을 확인하지 못했는지 적습니다. `office.py transfer` 일부 실패도 `UNVERIFIED`로.
- `_runtime.ok_file`: 크기만 보던 것을 형식별로 다시 열어 확인(xlsx openpyxl, docx python-docx, pptx python-pptx, CSV 행 수, md/txt 읽기, PDF 머리). 열지 못하면 `UNVERIFIED`.
- 오류 로그(`errors.txt`)에 메일 본문·받는 사람·제목, `--text`, 검색어 등 내용 옵션 값이 그대로 남던 문제. `<N자 생략>`으로 남기고 경로와 그 밖의 인자는 그대로.
- 쓰기 `--target`에 파일 이름만 주었는데 열린 문서도 파일도 아니면 현재 폴더에 새 파일을 만들던 문제. "찾을 수 없음"(열린 문서가 있으면 `[후보]`), 새 파일은 전체 경로일 때만.
- `office.py undo --pick N`(N>1): 고른 지점만 지워 더 새 지점이 남아 다음 `undo`가 앞으로 가던 문제. 같은 문서의 더 새 지점도 지웁니다(`removed_newer_points`). 수정 시각 비교는 같은 문서의 가장 새 지점 기준(그 사이 도구 저장을 사용자 저장으로 보지 않음).
- `outlook.py read --save`: 제목을 그대로 파일 이름에 써 `RE:`·`/`·`?` 등에서 실패하던 문제(첨부 저장과 같은 정리). 저장한 파일을 다시 읽어 확인.
- `outlook.py --body`: `\n` 변환을 두 번 해 `--body-file`의 `C:\new\data`까지 깨던 문제. `--body`만 `unescape_newlines`(경로 보존), `--body-file`은 그대로.
- `outlook.py --attach`: 쉼표로만 나눠 이름에 쉼표가 든 파일을 첨부하지 못하던 문제. 값 전체가 있는 파일이면 하나로, 여러 개는 `;`(쉼표도 됨).

### 수정 (reader)
- `read`: 한 단위가 약 8,000자보다 길어 잘려도 `truncated: false`였던 문제. `truncated: true`, 단위의 `text_truncated`·`text_chars`, `truncated_reason`, 항상 `--save` 명령(잘린 뒤쪽은 저장 파일에 있음).
- 일부 쪽만 스캔된 PDF: 문서 전체에 글자가 없을 때만 알려 0건 검색이 `VERIFIED`였던 문제. 글자 없는 쪽을 `meta.pages_without_text`로 남기고 `search`의 `not_searched`와 `read` 결과에 알립니다(`PARSER_VERSION` 5, 기존 캐시는 다시 읽음).
- 엑셀 시트 행 상한(20,000행)에서 자른 문서를 `read`/`read --save`가 그대로 `VERIFIED`로 알리던 문제(`truncated`, 저장은 `UNVERIFIED`). `search --save`가 5,000건에서 멈춰도 `VERIFIED`였던 문제(`save_truncated`, `UNVERIFIED`).
- `read --save`, `search --save`: 저장한 파일을 다시 열어 확인(CSV는 행 수). 확인하지 못하면 `UNVERIFIED`.
- `index`: 캐시 저장에 실패해도 "캐시에 저장 확인"이던 문제. 캐시 파일을 확인해 `cache_failed`와 `UNVERIFIED`.
- 옛 캐시의 빈 키를 채우던 `normalize_doc` 삭제 (베타, 이전 형식 캐시는 `PARSER_VERSION`으로 다시 읽음). 다시 읽지 못한 이전 형식 캐시는 검색하지 않고 `skipped`로 알립니다.

### 정리
- pyflakes: `ppt.py`의 쓰지 않는 `nonlocal slide_no`, `word.py` `_compress`를 대입 전 참조 없이 다시 씀. 테스트의 BOM 글자를 `"\ufeff"`로.

## [1.4.9] - 2026-10-31

### 수정
- 닫힌 Word 파일 작업: Word 2007은 `ReadOnly`와 `PasswordDocument`를 함께 넘기면 암호 없는 파일도 "명령을 정상적으로 종료할 수 없습니다"로 열지 못했습니다. 파일 구조로 암호 여부를 먼저 판별해, 암호가 없다고 확인된 파일에는 넘기지 않습니다. 판별하지 못한 파일(DRM 등)은 전과 같이 넘깁니다.
- 암호 Word 문서: Word 2007의 오류 문구에는 '암호'가 없어 `Office 오류`로 나오던 것을 "열기 암호가 걸린 문서" 안내로 바꿈.

## [1.4.8] - 2026-10-30

### 변경 (R1 사후 확인 마무리)
- 쓰기 명령은 모두 결과를 다시 읽어 확인한 경우에만 `VERIFIED:` + `verified: true`, 아니면 `UNVERIFIED:` + `unverified`.
  - `excel replace`: 바꾼 칸의 값. `excel style`: 준 서식(배경·글자색·굵게·기울임·크기·숫자 서식), 테마 머리글 색, 조건부 강조.
  - `excel sheet`: 시트 목록 (추가·이름 변경·복사·삭제·선택). `excel range`: 복사·이동한 대상 값, 지운 칸, 병합 상태, 너비·높이, 숨김, 틀 고정 (자동 맞춤은 비교할 값이 없음).
  - `excel chart`: 만든 차트가 목록에 있고 제목이 같은지, 지운 차트가 목록에 없는지.
  - `word replace`: 문서를 다시 찾아 찾는 글이 남지 않았는지 (바꾼 글이 찾는 글을 품으면 바꾼 글이 그만큼 있는지). `ppt replace`: 바꾼 슬라이드를 다시 읽어 같은 확인.
  - `ppt slide`: 슬라이드 수, 옮긴 슬라이드 제목, 비운 슬라이드의 요소 수.
  - `transfer` 차트: 넣기 전후 그림 수 (붙여넣기 방식 relay는 `UNVERIFIED`).
  - `office.py undo`: 되돌린 원본 파일이 스냅샷과 같은지.
  - `outlook --send`: 실제 발송은 Outlook이 처리하므로 `UNVERIFIED`(보낸 편지함 확인 안내). 작성 창은 받는 사람·제목을 다시 읽어 `VERIFIED`.
- `_runtime.ok()`는 더 이상 `VERIFIED:`를 대신 붙이지 않음 (다른 문구가 오면 내부 오류). 읽기·조회 명령은 읽은 결과가 곧 확인이라 기본 문구 유지.

## [1.4.7] - 2026-10-30

### 변경 (R1 사후 확인)
- 내보내기(`excel export`, `word export`, `ppt export`): 저장한 파일이 있고 비어 있지 않은지 확인 (`_runtime.ok_file`).
- `office.py transfer`: 엑셀 범위를 워드·PPT 표로 넣으면 넣은 범위·슬라이드에서 머리글을, 워드·PPT 내용을 엑셀로 넣으면 넣은 칸을 다시 읽어 확인.

## [1.4.6] - 2026-10-30

### 변경 (R1 사후 확인)
- `excel write`: 쓴 칸을 다시 읽어 비교 (수식은 첫 칸의 수식, 값은 칸마다. 앞의 `'`는 값에 남지 않으므로 빼고 비교, Excel이 날짜 등으로 바꾼 글자는 비어 있지 않으면 인정).
  같으면 `VERIFIED:` + `verified: true`, 다르면 `UNVERIFIED:` + `unverified`(다른 칸).
- `word write`: 넣은 범위를 다시 읽어 비어 있지 않고 넣은 글의 첫 줄(마크다운 기호 제외)이 들어 있는지 확인.
- `ppt write`: 쓴 슬라이드를 다시 읽어 제목·첫 줄·표 머리글·카드 제목이 들어 있는지 확인.
- 위 명령의 "다시 읽어서 확인하지 마세요" 문구를 "다시 읽어 확인했습니다"로.

## [1.4.5] - 2026-10-30

전체 검토와 회사 PC 오류 로그에서 찾은 오류 수정.

### 수정
- 문서: 문자 규칙(유니코드 화살표, 긴 줄표) 정리, README 구성도(word outline/comment/revisions, ppt notes, reader clear), 결과 처리 표에 `UNVERIFIED:`, 깨진 SKILL.md 링크.
- `excel write`: 머리글을 먼저 쓰고 수식 검사에서 실패해 머리글만 남던 문제. 모든 검사를 마친 뒤 씀. 수식이 여러 개인데 범위가 없으면
  머리글 행부터의 범위를 `[추천]` (이전 예시를 따르면 한 행씩 밀려 `#VALUE!`인데 ok).
- `excel write`: 값이 있는 칸을 확인 없이 덮어쓰던 문제 (SKILL.md "덮어쓸 칸"과 불일치). 이제 `needs_user_decision`, 사용자가 원하면 `--overwrite`.
- `excel write --formula 73.4 49.9`처럼 숫자만 주면 `--values`를 `[추천]`.
- `office.py transfer --to excel`: 워드/PPT 글이 숫자·날짜·수식으로 바뀌던 문제 ("001", "1-2", "(100)", "=..."는 글자 그대로).
- `excel read`: 닫힌 xlsx·CSV에서 `B:B`, `5:7` 같은 열·행 범위를 읽지 못하던 문제.
- 워드·PPT `--text`의 `\n` 변환이 Windows 경로(`C:\new\data`)를 깨던 문제.
- `office.py undo --list`: 항목마다 그대로 실행할 `command`(같은 `--app` 포함). 필터에 따라 번호가 달라 다른 문서가 되돌려질 수 있었음.

### 기능
- 인자 오류에 `[추천]`: 없는 명령(`excel.py list` -> `info`, 오타 -> 가까운 명령), 다른 옵션 이름(`office.py undo --target` -> `--app`),
  따옴표 없이 띄어쓰기가 든 경로(큰따옴표 안내).
- 열린 문서가 없으면 `needs_user_decision`: 사용자에게 전체 경로를 물어 `--target`으로 1번 실행 (경로 추측 금지).

## [1.4.4] - 2026-10-30

### 변경
- SKILL.md 규칙 9: 도구 코드(.py)는 사용자가 코드를 열어 보라고 직접 요청할 때만 읽음 (ml-kit과 같게).

## [1.4.3] - 2026-10-29

### 구조
- 폴더 정리: `CLI/office-cli/`에는 실행에 필요한 파일과 `SKILL.md`만 남김 (`excel.py`, `word.py`, `ppt.py`, `outlook.py`, `reader.py`, `office.py`, `_runtime.py`, `_office.py`).
  `README.md`, `CHANGELOG.md`, `TEST_CHECKLIST.md`, `requirements.txt`, `tests/`, `tools/`는 `CLI-dev/office-cli/`로 옮김.
  공통 계약 `CLI/README.md`와 `run_all_tests.py/.bat`도 `CLI-dev/`로.
- 테스트와 실기 러너는 옆 폴더 `../CLI/office-cli`의 도구를 시험. 사내 PC에서는 실기 테스트 때만 `CLI-dev`를 `C:\CLI-dev`로 `C:\CLI`와 나란히 둠.
- CLI 인자, JSON 키, 도구 동작은 그대로.
- `pyproject.toml` 삭제: pytest `testpaths` 2줄뿐이었고 반입 검사에 자주 걸림.
- 실기 테스트 스킬(`office-cli-test`)의 러너 경로: `C:\CLI-dev\office-cli\tools\run_win_tests.py`.

### 수정
- `-h` 결과 JSON에 `skill_file`(SKILL.md 경로)과 "도구 코드(.py)는 읽지 말고, 모르면 사용자에게 물어보라"는 안내 추가.
- SKILL.md 규칙 9: 도구 코드(.py)를 읽지 않고, 방법을 못 찾으면 사용자에게 알리고 멈춤.

## [1.4.2] - 2026-10-29

### 기능
- 새 명령 `reader.py clear`: 캐시 지우기(색인 항목과 캐시 파일, 원본 문서는 그대로). 대상은 `--file`, `--name`, `--folder`, `--missing`(원본이 없어진 문서), `--all`(색인에 없는 남은 캐시 파일 포함).
  `--dry-run`은 지울 목록(`planned`)과 실제로 지우는 명령을 줍니다. 지운 뒤 색인과 캐시 파일을 다시 확인해 `verified`.
- `reader.py search`: 띄어쓰기·줄바꿈에 더해 기호(`# - . : / ( )` 등)도 무시하고 찾습니다. `SOI 1`로 `SOI #1`, `SOI-1`을 찾습니다(이전에는 `SOI #1`로만 찾힘).
  숫자 경계는 지켜 `SOI 1`이 `SOI #10`과, `1.2`가 `11.2`·`3.1.2`와 맞지 않습니다. 기호를 빼면 숫자만 남는 검색어(`#4`, `1.2`)는 기호까지 그대로 찾습니다.
  결과 강조(`**`)와 저장 파일의 문장도 같은 기준입니다.

### 문서
- SKILL.md(office-cli, office-cli-test) 보고 규칙: 파일 경로를 `file:///` 링크로 바꾸지 않고 결과 JSON 값을 그대로 1번만 쓰게 함.
  긴 대화에서 26B 모델이 한글 경로를 링크로 변환하다가 같은 조각(`l'al l'al ...`)을 끝없이 반복한 사례 대응.

## [1.4.1] - 2026-10-29

배경: 회사 PC 실기 (2026-10-29 로그). 실기 러너 71/71 통과(E3, K1~K5 포함). 오프라인 테스트 실패 4건·오류 10건과 OpenCode 대화 6개에서 나온 문제를 고쳤습니다.

### 수정
- `ppt.py notes --file`, `word.py comment --file` 등: Windows 메모장 형식(CRLF) 파일을 넣으면 줄마다 빈 줄이 하나 더 생기던 문제. `read_text_file`이 줄바꿈을 `\n`으로 맞춥니다.
- `excel.py read --target <CSV>`: CSV/TSV도 Excel 없이 읽습니다(인코딩 자동 판별, 숫자 변환, `--range`). 이전에는 확장자 오류가 나 Gemma가 pandas 코드를 직접 짰습니다.
- reader 제목 인식: `(3) Verification Review; and`처럼 `;` `,` `and/or/및/또는`으로 끝나는 나열 항목은 장절 제목이 아닙니다(검색 근거의 장절이 나열 항목으로 바뀌던 문제). `PARSER_VERSION` 4.

### 테스트
- 회사 PC처럼 python-docx 기본 템플릿에 numbering.xml이 없으면 번호 문단을 만드는 테스트 도우미가 `NotImplementedError`로 실패하던 문제(오류 10건, 실패 3건). 도우미가 번호 파트를 직접 만듭니다. 도구 동작과는 무관합니다(reader는 번호 파트가 없어도 읽음).

## [1.4.0] - 2026-10-28

배경: 동료가 만드는 office MCP와 기능을 비교해 필요한 것만 가져왔습니다. 규격서·보고서 검토(장절을 골라 읽고, 검토 결과를 메모로 남기고, 수정 기록을 확인)와 발표 준비(노트)에 씁니다.

### 기능 (추가, 기존 인자·키 유지)
- `word.py read`에 범위 고르기 `--pages 3-5`, `--paragraphs 120-140`(`141-`는 끝까지), `--section "제목 일부"`(+ `--occurrence N`). 셋 중 하나만 씁니다.
  - `--pages`: Word가 계산한 실제 쪽(`GoTo`, `Information(wdActiveEndPageNumber)`). 앞 쪽에서 이어지는 단락도 넣고 단락마다 `page`. 결과에 `page_total`. 닫힌 .docx는 reader 캐시와 같은 쪽 정보(저장된 쪽 경계)로 고르고, 쪽 정보가 없으면 `--paragraphs`/`--section`을 안내합니다.
  - `--section`: 제목 전체가 같은 것 -> 제목이 그 글로 시작하는 것(`1.1`은 `1.1. Scope`와 맞고 `1.1.1 Data`와는 안 맞음) -> 제목에 들어 있는 것 순으로 찾습니다. 여럿이면 `[후보]`, 없으면 비슷한 제목 `[추천]`.
  - 결과에 `selection`, `count`. 잘리면 나머지를 `--paragraphs a-b`(또는 `--pages`)로 이어 읽는 명령과 고른 범위 전체를 `--save`하는 명령. 0건은 `count: 0`과 안내.
- 새 명령 `word.py outline`: 제목마다 `paragraph`(read `--paragraphs`와 같은 번호), `level`, `text`, `page`. 200개씩 `--start`.
- 새 명령 `word.py comment`: 검토 메모 목록(`--action list`, 기본) / 메모 달기(`--action add --anchor 글 --text 메모` 또는 `--file`).
  메모는 기준 글이 든 단락을 `write --anchor`와 같은 규칙으로 고른 뒤 그 글 자체에 답니다. `--file 메모.csv`(열: 기준 글, 메모 내용[, 몇 번째 단락])는 여러 메모를 한 번에 달고, 기준 글 하나라도 정할 수 없으면 하나도 달지 않습니다(어느 줄이 왜 안 되는지 message). 단 뒤 메모 수와 내용을 다시 읽어 확인합니다(`verified`, 다르면 `UNVERIFIED:`). `office.py undo`로 되돌릴 수 있습니다.
- 새 명령 `word.py revisions`: 변경 내용 추적 목록(`type` 삽입/삭제/서식 변경 등, `author, date, page, text`, `track_changes`). 읽기만 합니다.
- 새 명령 `ppt.py notes`: 발표자 노트 읽기(`--slides`, 노트가 있는 슬라이드만, 0건 안내) / 쓰기(`--slide N --text` 또는 `--file`, `--append`). 쓴 뒤 다시 읽어 확인(`verified`, `before`, `after`).
- 공통 런타임: `ok_unverified()` (R1의 `verified: false` + `UNVERIFIED:` 응답). 이번에 추가한 쓰기 명령(`comment --action add`, `notes` 쓰기)부터 `verified` 키를 씁니다.

### 점검 반영: 공통 런타임 (모든 명령에 영향)
전체 코드 점검에서 경량 LLM을 잘못 이끌 수 있는 공통 동작을 고쳤습니다 (R3, R4, R5, R7, R8).
- `-h`/`--help`: 도움말은 stderr, stdout에는 `{"status":"ok","command":"help","verified":true,...}` 1줄 (이전에는 stdout에 사람용 글만 나와 JSON 파싱 실패). 서브명령도 같음.
- `[추천]` 기준: 표기만 다른(대소문자·공백·`_`·`-`) 이름이 하나일 때만 `[추천]`. 부분·유사 일치는 뜻이 다를 수 있어 `[후보]`(사용자가 고름, `error_type: ambiguous`).
  예: 시트 `2024` -> 이전 `[추천] 2024_1분기`, 이제 `[후보] 2024_1분기, 2024_2분기`. 첨부 `견적서.pdf`가 없을 때 `견적서_구버전.pdf`를 추천해 `--send`와 함께 다른 파일이 발송될 수 있던 문제.
  도구가 정한 옵션 값(색 이름, 테마, 형광펜 색, 앱 이름)의 오타는 가까운 값 하나를 계속 `[추천]`.
- 오류 분류: 일반 `ValueError`/`KeyError`는 코드 버그일 수 있어 `ERROR:`(내부 오류, 오류 로그에 traceback). 입력 문제는 `InputError`로 냅니다.
- 출력 파일: `--save`, `export` 등의 저장 경로에 같은 이름의 파일이 있으면 덮어쓰지 않고 `_YYYYMMDD_HHMMSS`를 붙인 새 이름으로 저장합니다. 결과에 `recovered`.
  잠긴 파일은 잠깐 기다려 보고(`retry_fs`) 그래도 잠겨 있으면 새 이름으로 저장하고 `recovered`. 이미 시각이 붙은 이름에 시각을 두 번 붙이지 않습니다.
- `note_recovered()`와 응답의 `recovered` 키(성공·오류 모두). `retry_fs`는 일시적 오류(접근 거부, 공유·잠금 위반)만 재시도하고, 없는 파일 같은 영구 오류는 바로 올립니다.

### 점검 반영: 1.4.0 새 기능
- `ppt.py notes` 쓰기: 기존 노트가 있으면 `--append` 또는 `--overwrite`가 있어야 바꿉니다. 둘 다 없으면 노트를 지우지 않고 `needs_user_decision`.
  노트 본문 개체틀이 없는 파일에서 다른 개체틀(머리글·쪽 번호)에 쓰던 문제. 읽기에서 `--slide 3`도 그 슬라이드만 읽습니다. 이어 읽기 명령이 고르지 않은 슬라이드를 넣던 문제.
- `word.py comment`: `--anchor`/`--text`/`--file`을 주고 `--action add`를 빠뜨려도 메모를 답니다 (이전에는 목록만 보여 주고 완료로 응답).
- `word.py read --paragraphs`에 떨어진 범위(예: `1-12,20`)를 주고 결과가 잘리면 오류가 나던 문제. 이어 읽기는 남은 단락 번호로 줍니다.
- `ppt.py read/info/notes --slides` 형식 오류 문구가 `--pages`를 안내하던 문제.

### 점검 반영 2: 데이터 손상 방지
- 엑셀 쓰기(`write`, `replace`): Excel이 COM으로 받은 글자를 직접 입력한 것처럼 다시 해석해 값이 바뀌던 문제. 앞자리 0 코드(`01234` -> 1234),
  15자리 초과 숫자, 월-일·분수처럼 보이는 글(`1-2`, `3/4` -> 날짜), 괄호 숫자(`(100)` -> -100), 지수(`1E5`)는 앞에 `'`를 붙여 글자 그대로 씁니다(`'`는 셀 값에 남지 않음).
  15자리를 넘는 숫자 글자는 숫자로 바꾸지 않습니다(정밀도 손실). 연도가 있는 날짜(`2026-10-28`)는 계속 날짜로 바뀝니다.
- `excel.py range --action clear`: `--range`를 빼면 시트 전체를 지우던 문제. `--range`가 필요합니다.
- `range --action merge`: 값이 있는 칸이 둘 이상이면 왼쪽 위 밖의 값이 지워지므로 병합하지 않고 `needs_user_decision`(지워질 칸과 값). `--overwrite`로만 병합.
- `range --action copy/move`: 대상 범위에 값이 있으면 덮어쓰지 않고 `needs_user_decision`(원본과 겹치는 칸은 제외). `--overwrite`로만 덮어씀.
- `range --action hide/unhide/autofit`: `--axis`를 생략하고 `5:7`처럼 행 주소를 주면 시트의 모든 열이 숨겨지던 문제. 행 주소면 행으로 판단합니다.
- `write --append`: 시작 열 하나만 보고 마지막 행을 찾아 다른 열의 데이터를 덮어쓰던 문제(쓸 열 전체에서 가장 아래 행). 머리글 비교를 1행이 아니라 시작 행과 하고, `--header`도 같은 머리글이면 생략. 시작 행 위쪽에 쓰던 문제.
- `word.py write --anchor`: 앵커가 표 칸 안이면 `after`는 옆 칸이 아니라 표 뒤에 넣고, `before`/`replace`는 거절합니다. 문서 마지막 단락이 앵커일 때 앵커 단락에 글이 붙고 스타일이 바뀌던 문제. 표로 시작하는 문서의 `--pos start` 거절.
- `office.py undo`:
  - 스냅샷을 만들지 못한 쓰기와 새 파일 쓰기는 "되돌릴 지점 없음" 표식을 남깁니다. 직전 작업이 그렇다면 더 오래된 지점으로 넘어가 그 사이 작업까지 지우지 않고 `needs_user_decision`(`--list`에 `undo_available: false`).
  - 되돌리기 지점 뒤에 파일이 다시 저장되었으면 `needs_user_decision`(`--overwrite`로만 되돌림).
  - 열린 문서를 닫은 뒤 복사가 실패하면 원본을 다시 열고 `undo_failed` 오류에 `pre_undo_backup` 경로. 복사는 `retry_fs`로 재시도.
- 쓰기 대상(`--target`, transfer 대상)이 열린 문서도 파일도 아니고 이름이 비슷한 열린 문서가 있으면 새 파일을 만들지 않고 `[후보]`. 새 파일은 전체 경로일 때만 바로 만듭니다.

### 점검 반영 3: reader 근거 정확도 (`PARSER_VERSION` 3, 이전 캐시는 다시 읽을 때 갱신)
- PDF: 여러 쪽의 위/아래 여백(8%)에 반복되는 줄(문서 제목 머리글, 쪽 번호)을 뺍니다(`meta.header_footer_removed`). 이전에는 머리글이 매 쪽 장절 경로를 끊고 검색 결과에 쪽마다 걸렸습니다.
- PDF 2단 편집: 가로로 걸친 블록(제목, 표) 사이마다 왼쪽 단을 먼저 읽습니다. 이전에는 y 좌표로만 정렬해 두 단의 줄이 섞여 장절이 틀어졌습니다.
- 제목 인식: 번호 뒤가 단위인 수치 줄(`10.5 kV 이하`, `3 mm`, `2 개월`)과 목차 줄(`1.2 설계 기준 ....... 12`, 탭 뒤 쪽 번호)은 제목이 아닙니다.
- DOCX: 변경 추적으로 삽입된 글(`w:ins`, `w:moveTo`)을 넣고 지운 글은 뺍니다. 콘텐츠 컨트롤(`w:sdt`) 안의 본문도 읽습니다. 여러 쪽에 걸친 표는 행마다 그 행의 쪽(이전에는 모두 표가 시작한 쪽).
  빈 번호 단락도 번호 카운터를 진행해 뒤 조항 번호가 밀리지 않습니다.
- HWP: 확장 한자 등 BMP 밖 글자(서로게이트 쌍)를 바르게 읽고, 짝 없는 글자는 대체 문자로 바꿔 캐시 저장이 실패하지 않습니다.
- `search`:
  - 검색 범위가 불완전하면(등록하지 못한 파일, 읽지 못한 파일, 다시 읽지 못한 캐시, 스캔 PDF, 20,000행에서 자른 엑셀, 들어가지 않은 하위 폴더) `not_searched`와 함께 `UNVERIFIED:`. 이전에는 일부를 못 찾고도 "VERIFIED: 0건, 다시 실행하지 마세요"였습니다.
  - 원본이 바뀐 캐시는 검색할 때 그 자리에서 다시 읽습니다(20개까지, `recovered`). 다시 읽지 못한 것은 결과 항목에 `stale: true`.
  - 검색어: 따옴표 구문, 소문자 `and`/`or`, 띄어쓰기·줄바꿈 무시(`설계기준` = `설계 기준`), NFC 정규화, `--regex`의 빈 일치 제외.
- 색인: 캐시는 최신인데 색인에 없으면(저장 전 중단, 동시 실행) 다시 등록할 때 색인을 고칩니다. Windows에서 경로 대소문자만 다른 항목은 하나로 합칩니다.
- `read --pages`: 쪽 정보가 있는 DOCX/HWPX에서도 실제 쪽으로 고릅니다(`search`의 근거 `p.12`를 그대로 읽을 수 있게). 결과에 `pages_mode`.
- `excel.py read`(닫힌 xlsx): 크기 정보(`<dimension>`)가 없는 파일을 A1 한 칸만 읽고 성공으로 응답하던 문제.

### 점검 반영 4: 실제 환경 견고성 (회사 PC 실기 확인 필요)
- 닫힌 파일을 보이지 않는 Office로 열 때 암호 창이나 열기 매크로 창이 뜨면 끝나지 않던 문제:
  Excel/Word는 틀린 암호를 넘겨 바로 "열기 암호가 걸린 문서" 오류(STOP_RETRY)가 나게 하고, 백그라운드 인스턴스의 매크로를 끕니다(`AutomationSecurity` 3).
  Excel은 읽기 전용 권장 창도 띄우지 않습니다. PowerPoint와 DRM 인증 창은 아직 막지 못합니다(남은 할 일).
- 쓰기 대상 파일을 다른 사람·프로그램이 열고 있어 읽기 전용으로 열리면, 고친 내용을 저장할 수 없으므로 쓰기 전에 멈춥니다(이전에는 저장 실패 또는 저장 안 된 채 성공 응답).
- Office가 재계산·자동 저장으로 잠깐 바빠 호출을 거부하면 COM 메시지 필터가 15초까지 기다렸다 다시 부릅니다(`recovered`). 넘으면 기존 `office_busy` 오류.
- Outlook `--folder`: `inbox\프로젝트`처럼 기본 폴더 별칭으로 시작하는 경로를 지원하고, 없는 하위 폴더의 후보를 전체 경로로 줍니다(일부만 주면 상위 폴더 전체를 검색함).
- `office.py focus`: 창 전환이 실패해도 Alt 키가 눌린 채 남지 않습니다.

### 테스트 러너 (`tools/run_win_tests.py`)
- K 그룹 추가(K1~K5): outline·장절, 쪽·단락 범위, 메모 달기·목록, 변경 추적 목록, PPT 노트. 가짜 COM 테스트로는 Word 쪽 계산(`GoTo`, `Information`)과 `Comments.Add`, 노트 개체틀이 실제와 같은지 확인할 수 없어 회사 PC 실기로 확인합니다.

## [1.3.0] - 2026-10-28

배경: 회사 PC에서 reader.py의 두 가지 쓰임새를 지원합니다.
1. 작성 중인 문서(주로 Word/Excel, 가끔 PPT)가 규정·규격서(주로 PDF, Word)에 맞는지 확인.
2. 여러 문서에서 키워드를 찾아 관련 정보를 근거와 함께 모으기. 근거는 쪽 번호, 장절 경로, 문항/조항 번호가 필수입니다.
이전 캐시는 형식마다 단위 모양이 달랐고(PDF는 쪽 하나가 단위, 표는 통째로), 장절은 Word 제목 스타일만 알아서 근거를 댈 수 없었습니다.

### 기능 (추가, 기존 인자·키 유지)
- 공통 캐시 형식(`PARSER_VERSION` 2): 모든 형식의 단위가 같은 키 `id, kind, loc, page, page_source, section, number, heading, text`를 가집니다.
  `heading`은 하위 호환용(`section`의 마지막 값). 다시 읽으면(`read`, `index`) 캐시가 새 형식으로 바뀝니다.
- 장절/번호 인식: 번호 토큰을 계열(로마 숫자, 점 번호 `1.1.1`/`1.0#`, 가나다, `1)`, `(1)`, `가)`, `(가)`, `제N장/절/조`, `Chapter N`, `Annex A`, `부록 1`, 문항 `Q3` 등)로 나누고 문서에서 처음 나온 순서로 수준을 정합니다.
  글머리 기호(`O`, `-`, 동그라미/네모 기호 등)는 번호도 제목도 아닙니다. 짧은 한 줄이고 문장 끝으로 끝나지 않을 때만 제목(오탐 방지).
  PDF는 글자 크기/굵게(본문 크기 = 가장 흔한 크기)를 제목 신호로, DOCX는 제목 스타일·개요 수준과 Word 자동 번호(numbering.xml의 `lvlText`, 시작 값, decimal/로마/가나다 형식)를 씁니다.
  자동 번호는 글 앞에 붙여(`3.2.1 세부 조건`) 검색과 근거에 나옵니다. 번호 해석이 실패하면 글의 번호 패턴으로 대신합니다. HWPX는 개요 수준과 개요 번호, HWP/TXT는 번호 패턴.
- 쪽 번호: PDF는 쪽마다 단락/줄 묶음 단위로 정확한 쪽. DOCX는 `w:lastRenderedPageBreak`(Word가 저장한 쪽 경계, `word_saved_layout`)가 있으면 그것으로,
  없으면 강제 쪽 나눔만 세어(`page_break`, 근거에 `(추정)`), 둘 다 없으면 null. PPTX는 슬라이드 번호(`slide`), HWPX는 문단 쪽 나눔이 있을 때만. XLSX/CSV/HWP/TXT는 null.
- 표는 행 단위(`kind: row`): 첫 행이 머리글이면 `항목: 내전압 | 기준: 1.5 kV 1분 | 시험방법: KS C 1234`. 병합/빈 머리글은 앞 칸을 이어 쓰고 여러 줄 칸은 공백으로 잇습니다.
  DOCX 표(`loc` `표 3 5행`), PDF 표(pymupdf `find_tables()`, `p.3 표 1 2행`, 표 영역의 글은 단락에서 뺌). 행 `number`는 첫 칸이 번호(`1`, `1.2`, `REQ-001`)일 때.
- 새 명령 `reader.py index --folder 폴더 [--depth N | --recursive] [--max-files 300]`, `index --file 파일`, `index --all`: 내용 출력 없이 캐시에 등록.
  결과 `count, indexed, updated, unchanged, failed, failures(20개), remaining, truncated, skipped_types, total_units`. `--max-files`는 새로 읽는 파일 수만 세므로 같은 명령을 다시 실행하면 이어서 등록합니다.
- `search`: 결과마다 `page, page_source, section, section_text, number, evidence`(`규격서.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)`). 새 인자 `--pages 3-7`, `--section "장절 일부"`, `--max-files`.
  `--save`에 `.xlsx` 추가(머리글 굵게, 줄바꿈, 열 너비; 기본은 계속 `.md`). 열: 파일, 쪽, 장절, 번호, 위치, 내용(문장/단락), 문맥. `.md`는 항목마다 `근거:` 줄.
  저장 파일에는 `--top`과 관계없이 전체 결과(최대 5,000건)를 담습니다(이전에는 표시한 건수만). R2 키 `count`, `shown`, `truncated` 추가, 잘리면 `--save "근거.xlsx"` 명령.
- `search --file 파일/폴더`: 지금처럼 그 대상을 검색하되, 캐시가 없거나 낡은(수정 시각/크기/파서 버전) 파일만 읽어 등록하고 최신이면 캐시를 씁니다. 결과에 `indexed_now`(새로 등록한 파일, 20개까지), `indexed_now_count`, `from_cache`.
  폴더는 `--max-files`(기본 300)까지 새로 읽고 남으면 `not_indexed_remaining`과 `index --folder` 명령.
- `search`(`--file` 없음)는 캐시만 검색합니다. 이전 형식 캐시도 빈 키를 채워(`p.N`/`슬라이드 N`에서 쪽, `heading`에서 장절) 검색하고, 원본이 바뀌었거나 이전 형식이면 `stale_cache_count`와 `index --all` 명령을 줍니다.
  등록된 문서가 없으면 `error_type: not_indexed`(STOP_RETRY: 사용자에게 폴더를 물어 `--file`로 실행).
- `read`: 단위마다 `kind, page, section, number`. `--save`에 `.xlsx` 추가, `.md`에 `근거:` 줄. `read --pages`는 PDF에서 쪽 번호로 고릅니다(PDF 단위가 쪽에서 단락으로 바뀜). 새 인자 `--units`(단위 번호, 쪽 중간에서 잘렸을 때 이어 읽기 명령이 씀).
- `list`: 캐시된 파일 조회. 문서마다 `type, units, pages, sections, cached_at, parser_version, cache(current|stale), stale, missing`. 새 인자 `--name 이름일부`, `--folder 폴더`. R2 키 `count, shown, truncated`, 100개를 넘으면 `full_result_file`(CSV).
- 공통 런타임: `InputError(directive=, error_type=, fields=)`로 오류 응답에 전용 안내와 키(예: `next_commands`)를 붙일 수 있습니다.
- PDF를 읽을 때 pymupdf의 안내/경고 출력(예: pymupdf_layout 권장 문구)이 stdout으로 나와 1줄 JSON 계약을 깨지 않게 stderr로 돌립니다.

### 제거
- EML 읽기 기능 제거 (Outlook 통합 도구에서 별도로 다룰 예정). `.eml`은 다른 미지원 확장자와 같은 "지원하지 않는 형식입니다" 오류이고, `index --folder`는 `skipped_types`에 셉니다.

### 동작 변경
- `search`(`--file` 없음)는 이제 캐시만 봅니다. 이전에는 색인된 문서를 검색할 때마다 원본이 바뀌었으면 다시 읽었습니다. 최신 내용이 필요하면 `index --all` 또는 `search --file`.
- 표가 표 하나 = 단위 하나에서 행 하나 = 단위 하나로 바뀌었습니다 (`loc` `표 1` -> `표 1 2행`). PDF 단위는 쪽에서 단락으로 바뀌었습니다 (`loc`은 그대로 `p.N`).
- `read`의 PDF `unit_name`이 `페이지`에서 `단락`으로 바뀌었습니다. `--pages`의 뜻(PDF 쪽, PPT 슬라이드, 그 밖은 항목 번호)은 그대로입니다.

### 스킬
- SKILL.md reader 행: 문서 등록(`index --folder`, `read --file`), 근거 검색(`--pages`, `--section`, `--save "근거.xlsx"`), 캐시된 파일 조회(`list --name`). 근거를 `evidence`(파일 p.N, 장절 (번호))로 알리고, page가 null이면 장절과 번호로 알리는 규칙.

### 테스트
- `tests/test_reader_evidence.py`: 번호 계열 분류, 글머리, 제목 오탐 방지, 세 가지 번호 체계 중첩(PDF, 글자 크기 포함), DOCX 제목 스타일·자동 번호·쪽 나눔·저장된 쪽 경계,
  PPTX/XLSX/HWPX 장절과 쪽, 표 행 형식(병합 머리글, 여러 줄 칸), PDF 선 표, `index` 집계와 이어서 등록, 캐시 기준 검색과 자동 등록, `--pages`/`--section`, xlsx/md/csv 저장, 이전 캐시 정규화와 stale, `list` 필터, `.eml` 미지원.
- `tests/test_reader.py`: EML 샘플과 테스트 제거, 표 행 단위와 폴더 검색의 `from_cache`에 맞춤.

## [1.2.2] - 2026-10-28

사내 경량 LLM(Gemma, OpenCode + PowerShell)으로 `docs/opencode/SCENARIOS.md`를 실행한 시나리오 테스트 결과 반영 (시나리오 10, 12).

### 스킬
- 도구 위치를 `C:\CLI\office-cli`로 고정: SKILL.md와 office-cli-test SKILL.md의 명령 경로를 실제 경로로 적음 (office-cli 옆 배치 조건 없앰).

### 수정
- (시나리오 10) Sheet1만 있는 빈 통합문서에서 "Sheet9에 '테스트'라고 써줘" -> Gemma가 사용자에게 묻지 않고 Sheet9를 만들어 썼습니다.
  원인: 없는 시트 오류(`_office.missing_sheet`)의 message에 "새 시트로 만들려면 같은 명령에 --new-sheet를 붙이세요"가 있었고, 비슷한 시트가 있으면 `[후보]`에 "(새 시트 만들기: --new-sheet)"를 섞어 넣어 LLM이 그것을 골랐습니다.
  - message와 `[후보]`에는 사실만 적습니다: 없는 시트, 있는 시트 목록, 아무 데도 쓰지 않았음, 사용자에게 물어볼 것. `--new-sheet`와 명령문을 넣지 않습니다.
  - 표기만 다른 시트(대소문자·공백·_-.)가 하나면 지금처럼 `[추천]`(error_type `input`). 그 밖에는 새 error_type `needs_user_decision`(첫 단어는 `STOP_RETRY:` 그대로). 비슷한 기존 시트가 있으면 그 이름만 `[후보]`와 `candidates`에 넣습니다.
  - `agent_directive`: `STOP_RETRY: 다시 실행하지 마세요. message를 사용자에게 그대로 보여 주고, 새 시트 '<이름>'을(를) 만들지 물어보세요. 사용자가 만들라고 답한 경우에만 같은 명령에 --new-sheet를 붙여 실행하세요.` (`[후보]`가 있으면 기존 시트를 고르는 경우도 안내)
  - 공통 런타임에 `UserDecisionError`(InputError 하위 클래스, 사용자 결정이 필요한 오류) 추가.
  - SKILL.md 3절 `[후보]` 행에서 `--new-sheet`를 고를 값 목록에서 빼고 "사용자가 새 시트를 원한다고 답한 경우에만"으로 바꿨습니다.
  - `office.py transfer --to excel --sheet X`에서 X가 이미 있을 때의 `[추천] X 2`는 유지합니다. transfer는 원래 항상 새 시트에 받고(`--sheet` 생략 시에도 `Word 내용 2`처럼 비어 있는 이름), 기존 시트를 덮어쓰는 길은 `--overwrite`뿐이라 추천대로 다시 실행해도 기존 데이터가 바뀌지 않습니다.
- (시나리오 12) 저장한 적 없는 새 워드 문서(`문서1`)에 엑셀 차트를 transfer한 뒤 "방금 한 거 되돌려줘" -> `office.py undo`가 빈 워드 창을 새로 열고, 차트가 들어간 문서는 그대로 남았습니다.
  원인: 저장한 적 없는 문서의 스냅샷은 항상 "새 제목 없는 문서"로 열었습니다(1.0 설계).
  - 스냅샷 때의 문서 이름으로 열린 문서 중 저장 안 된 문서를 찾습니다(`_office.open_untitled`, 스냅샷 meta에 `untitled` 추가).
  - 워드: 현재 내용을 `pre_undo_backup`(.xml)에 보관하고, 같은 창에서 내용만 스냅샷으로 바꿉니다(flat XML 스냅샷은 `InsertXML`, 그 밖은 `InsertFile`). `restored_as: "restored_in_place"`.
  - 엑셀/PPT: 현재 상태를 `pre_undo_backup`에 보관하고 스냅샷을 새 제목 없는 문서로 연 뒤, 바뀐 문서를 저장하지 않고 닫습니다(문서 수 그대로). `restored_as: "replaced_untitled_document"`.
  - 바뀐 문서가 닫혔거나 같은 이름의 저장 안 된 문서가 둘 이상이면 고르지 않고 예전처럼 새 문서로만 엽니다(`new_untitled_document`). 이유를 `preview`에 적습니다.
  - `preview`와 `agent_directive`에 무엇을 했는지 적습니다(예: `'문서1'을(를) 같은 창에서 되돌렸습니다`). 결과에 `restored_name`(되돌린 내용이 있는 문서 이름) 추가. 기존 키(`restored`, `app`, `name`, `restored_as`, `pre_undo_backup`) 유지.

### 실기 테스트 러너 (`tools/run_win_tests.py`)
- C3(저장 안 한 새 통합문서 undo): 기대값을 "통합문서 수 +1"에서 "수 그대로, 바뀐 통합문서는 닫힘, 되돌린 통합문서의 A1이 쓰기 전 값"으로 바꿨습니다.
- C8 추가: 새 워드 문서에 쓰고 undo -> 같은 창에서 문장이 사라지고 문서 수 그대로 (시나리오 12).

### 테스트
- 없는 시트: message·`[후보]`에 `--new-sheet`와 명령문이 없음, directive가 사용자에게 묻게 함, `[추천]`(표기만 다름) 유지, `[후보]`는 실제 시트만, 시나리오 10 재현.
- 저장 안 한 문서 undo: 워드 같은 창 복원(새 문서 없음, 백업 기록), 엑셀/PPT 바꿔 열고 바뀐 문서 닫기, 문서가 닫혀 있으면 예전 동작, 같은 이름 둘이면 고르지 않음.

## [1.2.1] - 2026-10-28

회사 PC 실기 테스트(Windows 11, Office 16.0 / Microsoft 365, WinPython 3.13.13, DRM) 결과 반영과 오류 로그(R6).

### 스킬
- SKILL.md를 PowerShell 기준으로 정리: 꼭 지킬 규칙 8개(인코딩 줄, 1회 실행, 표·여러 줄·큰따옴표 값은 파일로, `$`는 작은따옴표), PowerShell에서 따옴표가 사라지는 `--values` JSON 예시를 `--file "표.csv"`로 교체, `UNVERIFIED:` 행 추가, 유니코드 화살표 제거. 도구 동작은 바뀌지 않습니다.

### 기능 (추가만, 기존 인자·키 유지)
- 오류 로그: 실패한 실행(exit 1, 인자 오류 exit 2)을 `<data 폴더>/logs/errors.txt`에 자동으로 덧붙입니다 (`CLI-dev/README.md` R6, 기준 doc-classifier `write_error_log`).
  머리줄 `===== <시각> | office-cli <버전> | Python <버전> | <OS>`, `명령:`, `결과: <error_type> | <message>`, 한도를 넘긴 복구 시도는 `복구 시도:`, 내부 오류(`ERROR:`)만 `traceback:`.
  성공과 사용자 중단(130)은 기록하지 않고, 2MB를 넘으면 `errors.old.txt`로 넘깁니다(1세대). 회전 실패와 기록 실패는 따로 삼켜 stdout 1줄 JSON에 영향이 없습니다. 로그 경로는 JSON에 넣지 않습니다.
- `_runtime.VERSION`("1.2.1"), `record_recovery_attempt()`. `retry_fs`가 한도를 넘기면 `복구 시도:`에 남깁니다. SKILL.md `metadata.version`과 같은지 테스트합니다.

### 실기 테스트 러너 (`tools/run_win_tests.py`)
- A5(두 번째 Excel 정리)가 A.xlsx를 연 사용자 Excel까지 닫던 문제. 이후 B2, B3, C1, C2, D3, E, H1, H3, I1, I2가 `DISP_E_BADINDEX`(Workbooks("A.xlsx") 없음)로 실패했습니다.
  원인: 정리 스크립트가 "실행 시작 뒤 생긴 Excel/PowerPoint의 숨은 창 또는 제목 빈 창"에 WM_CLOSE를 보냈는데, A.xlsx를 연 Excel도 러너가 띄운 것이고 Office 16에는 숨은 XLMAIN 창이 있습니다 (Office 2007에서는 드러나지 않음).
  - 문서가 남아 있는 프로세스(COM 문서 수 > 0, 제목 있는 창, 창 제목에 A.xlsx 등)는 Quit도 WM_CLOSE도 하지 않습니다.
  - A5는 Excel을 띄우기 전후의 PID를 비교해 자기가 띄운 PID만 끝냅니다. 종료 정리는 실행 시작 뒤 생긴 빈 프로세스만, 실행 전부터 있던 Office는 문서가 없어도 끝내지 않습니다.
  - 같은 인스턴스가 두 번 잡히면 건너뛰고, Office 16의 빈 창 제목(`Excel`, `PowerPoint`)도 인식합니다.
  - 판단 규칙은 `tools/win_test_rules.py`로 분리해 오프라인 테스트(`tests/test_win_runner.py`, 가짜 Office/pywin32로 회사 PC 상황 재현)를 붙였습니다.
- JIT 복구: 그룹 시작 전과 문서 조회 때 A.xlsx/A.docx/A.pptx가 닫혀 있으면 준비 단계와 같은 방법(`os.startfile`)으로 다시 열고 `[WARN] A.xlsx가 닫혀 있어 다시 열었습니다`를 남깁니다.
  다시 열지 못하면 그 그룹의 남은 항목을 "테스트 문서 A.xlsx가 닫혀 다시 열지 못함" 하나의 이유로 바로 ERROR 처리합니다.
- `com_retry`가 '그 이름의 항목 없음'(`DISP_E_BADINDEX`)을 재시도하지 않습니다. 중첩된 재시도(최대 10x10x10회)로 멈춘 것처럼 보여 Ctrl+C로 끝내야 했습니다.
- 열린 문서 조회가 첫 번째 Excel 인스턴스만 보지 않고 모든 인스턴스를 찾습니다.
- `errors_*.txt`에 "러너 자동 복구와 창/프로세스 정리"(다시 연 문서, 닫은 문서, Quit/WM_CLOSE를 보낸 PID와 창, 보호한 PID)와 "도구 오류 로그"(이번 실행 중 `<테스트 폴더>/_data/logs/errors.txt`에 추가된 부분) 절을 넣습니다. 종료 정리 뒤에 한 번 더 써서 정리 기록까지 담습니다.

### 테스트
- (회사 PC 2차 실기 후 보완) 테스트용 docx를 폴더 템플릿으로 만들 때, 폴더에서 빠진 파일(회사 PC: `docProps/thumbnail.jpeg`)을 가리키는 관계는 빼고 엽니다 (`tests/sample_docs.py`). 도구 동작은 바뀌지 않습니다.
- DRM PC에서 python-pptx 기본 템플릿(`site-packages/pptx/templates/default.pptx`)이 zip이 아니라(암호화) `pptx.Presentation()`이 실패하던 문제(1 FAILED + 12 ERROR).
  템플릿이 필요한 테스트만 "python-pptx 기본 템플릿이 손상됨 (DRM 추정)"으로 건너뛰고, reader 샘플(docx/xlsx/pdf/eml/hwpx)은 형식별로 만들어 나머지 테스트는 그대로 돕니다.
  python-docx는 기본 .docx가 손상돼 있으면 폴더 템플릿(`docx/templates/default-docx-template`)으로 만듭니다 (`tests/sample_docs.py`). 도구 코드는 인자 없는 `Presentation()`/`Document()`를 쓰지 않습니다.

## [1.2.0] - 2026-10-28

저장소 지침(AGENTS.md 2026-10-28) 반영 1차: 사용자 확인 없이 덮어쓰거나 추측하던 동작을 멈춥니다.
근거 `docs/history/COMPLIANCE_REVIEW_2026-10-28.md` 2절, 구현 `CLI-dev/README.md` R8(`[추천]`/`[후보]`).

### 기능 (추가만, 기존 인자·키 유지)
- 후보가 둘 이상이라 도구가 고를 수 없을 때의 오류 형식: `error_type: "ambiguous"`, `candidates`(최대 10개), message 끝 `[후보] a, b, c`,
  directive `STOP_RETRY: 후보가 여러 개라 ... 재시도하지 말고 [후보]를 사용자에게 보여주고 고르게 한 뒤 ...`. SKILL.md 결과 처리 표에 추가.
- `excel.py write --new-sheet`: 없는 시트를 새로 만들어 쓸 때만 붙입니다. 이미 있는 이름에 붙이면 만들지 않고 오류.
- `word.py write --occurrence N`, `office.py transfer --occurrence N`: `--anchor`가 여러 단락에 있을 때 몇 번째 단락인지 (오류의 `[후보]` 번호).
- `office.py transfer --to excel --overwrite`: `--sheet`로 준 시트가 이미 있을 때 그 시트에 덮어쓰기.

### 수정 (동작 변경)
- `excel.py write --sheet`(또는 `--range '시트'!A1`)에 없는 시트 이름을 주면 그 이름으로 새 시트를 조용히 만들던 문제. 오타가 새 시트가 되었습니다.
  이제 만들지 않고 오류를 냅니다.
  - 대소문자·공백·`_-.`만 다른 시트가 하나면 `[추천]` (예: `Sheet 1` -> `Sheet1`).
  - 비슷하지만 다른 이름이면 `[후보]`로 사용자에게 고르게 합니다 (예: `Sheet2`를 `Sheet1`의 오타로 단정해 덮어쓰지 않음).
  - 비슷한 것이 없으면 기존 시트 목록과 `--new-sheet` 안내.
- `office.py transfer --to excel --sheet 이름`에 이미 있는 시트를 주면 A1부터 덮어쓰던 문제. 덮어쓰지 않고 비어 있는 이름을 `[추천]`합니다(`--overwrite`로 덮어쓰기). `--sheet`를 생략하면 지금처럼 새 시트(`Word 내용 2` 등).
- `word.py write --anchor`가 여러 단락에 맞으면 첫 번째를 고르고, `--pos replace`면 그 단락을 지우던 문제. 아무것도 바꾸지 않고 `[후보]`(번호와 단락 앞부분 30자)로 오류를 냅니다. 같은 단락에 여러 번 나오면 한 곳으로 셉니다.
  `office.py transfer --anchor`도 같습니다. 이 오류는 차트 전달 방식을 바꿔도 같으므로 다음 방식으로 넘어가지 않습니다.
- `--anchor`를 찾지 못하면 문서에서 가장 비슷한 문구(그대로 다시 찾을 수 있는 원문 일부)를 안내합니다.
  공백·대소문자만 다를 때, 또는 앵커 글자 사이에 2자 이하만 끼어들고(`2026 계획` -> `2026년 계획`) `--pos replace`가 아닐 때만 `[추천]`.
  글자가 바뀌거나 빠진 문구(`Q2` -> `Q1`)는 뜻이 다를 수 있어 `[추천]` 없이 "비슷한 문구: ..."로만 알립니다.

### 실기 테스트 러너
- I그룹이 미리 만들지 않은 `I_DATA`, `I_BIG` 시트에 쓰던 첫 write에 `--new-sheet` 추가 (1.2.0 동작 변경으로 실패하던 문제).

### 다음으로 미룸
- `--save`(CSV/MD/텍스트)와 COM 내보내기가 기존 출력 파일을 덮어쓰는 문제는 공통 런타임 반영(R4) 때 함께 처리 (남은 할 일).
- 열린 문서 이름(`--target 보고.xlsx`, `office.py focus --name`)이 여러 문서에 맞으면(다른 폴더의 같은 파일명, 다른 Excel 창의 `Book1`) 첫 번째를 고르던 문제. `[후보]`(전체 경로)로 오류를 냅니다. 전체 경로를 주면 그 문서로 정해집니다.

## [1.1.0] - 2026-10-27

### 구조 (이름 변경, 동작 불변)
- 도구 이름을 `atomic-cli`에서 `office-cli`로 바꿨습니다. 진입점 6개로 합치면서 "작은 도구 여러 개"라는 뜻이 더 이상 맞지 않기 때문입니다.
  - 폴더 `CLI/atomic-cli` -> `CLI/office-cli`, 스킬 이름(SKILL.md `name`) `office-cli`. OpenCode 스킬 폴더도 `office-cli`로 바꿔 배치하세요.
  - 환경변수 `ATOMIC_DATA_DIR` / `ATOMIC_CACHE_DIR` / `ATOMIC_VERBOSE` -> `OFFICE_CLI_DATA_DIR` / `OFFICE_CLI_CACHE_DIR` / `OFFICE_CLI_VERBOSE`. 예전 이름은 읽지 않습니다.
  - 실기 테스트 기본 폴더 `C:\atomic_test` -> `C:\office_test`.
- 내부 함수 `_runtime.atomic_write_text`(중단돼도 기존 파일이 깨지지 않는 저장)를 `safe_write_text`로 바꿨습니다. 코드에 `atomic` 이름이 남지 않습니다.
- 명령, 인자, JSON 키, `config.json`, `cache`/`data` 폴더 구조는 그대로입니다.

### 테스트 러너 (`tools/run_win_tests.py`)
- 분석용 오류 로그 `errors_<시각>.txt`를 `report_*.txt`와 함께 항상 만듭니다. 환경(Windows, Python, 라이브러리, Office 버전과 PDF 가능 여부, `config.json`, 러너 옵션)과, 실패/오류 항목마다 실행한 도구 명령 전체, 종료 코드, 걸린 시간, stdout JSON 전체, stderr 전체, traceback, 그 항목의 진행 로그를 담습니다. 기존 보고서는 JSON 400자, stderr 600자에서 잘려 원인을 알기 어려웠습니다.
- 준비 단계 실패(테스트 파일 만들기, 창 열기), 러너 자체 오류, Ctrl+C 중단에도 그 시점까지의 오류 로그를 남깁니다.
- `--agent`: 에이전트(OpenCode 스킬)용 실행. `--auto`와 같고, 진행 로그는 stderr로, stdout에는 결과 JSON 1줄(`summary`, `failed`, `report_file`, `error_log`, `agent_directive`)만 씁니다. 준비 실패는 exit 1, 중단은 130.
- 실기 테스트 스킬 `tools/office-cli-test/SKILL.md` 추가. 경량 모델이 사용자 확인 후 러너를 1회 실행하고 결과와 오류 로그 경로를 보고합니다.

## [1.0.1] - 2026-10-27

Windows 실기 테스트(집 PC: Windows 11 Pro 26200, Python 3.13.13, Office 2007 32비트) 결과 반영. 최종 65개 중 PASS 63, ENV-SKIP 2(PDF 추가 기능 없음).

### 기능 (추가만, 기존 인자·키 유지)
- `excel.py / word.py / ppt.py info --detail`: 결과가 제대로 보이는지 도구로 확인. PPT는 도형별 위치·크기·글자 크기·목록 종류와 경고(슬라이드 밖, 글자 넘침, 겹침, `--slides`로 범위 지정), 엑셀은 차트 종류·계열 수·위치와 데이터를 가리는 차트 경고, 워드는 그림 크기(쪽 너비 대비)와 제목 수준. 경고 수는 `layout_warnings`.
- `office.py windows`가 실행 중인 Office 버전과 PDF 내보내기 가능 여부(`office`)를 알려 줍니다. PDF 기능이 없는 Office(추가 기능 없는 2007 등)에서는 `export --save *.pdf`가 시도하지 않고 바로 STOP_RETRY로 안내합니다.
- 되돌리기 스냅샷에 문서 식별값(파일 속 만든 시각 `docProps/core.xml`)을 저장합니다. 같은 경로에 다른 문서를 새로 만들면 이전 문서의 스냅샷은 undo 목록에서 빠지고 정리됩니다(Office 저장으로는 바뀌지 않음).

### 수정
- `excel.py chart`를 `--at` 없이 만들면 데이터 범위 바로 오른쪽에 놓여 시트의 다른 데이터를 가리던 문제. 사용 범위 오른쪽에 놓습니다 (실기 J2에서 발견).
- `ppt.py export --save *.pdf`가 항상 내부 오류(`TypeError: The Python instance can not be converted to a COM object`)를 내던 문제. `ExportAsFixedFormat`의 개체 인자 `PrintRange`를 생략하면 pywin32가 인자 변환에서 실패하므로 `PrintRange=None`을 명시합니다.
- PowerPoint 2007(12.x)에 PDF 추가 기능이 없을 때 E_FAIL이 내부 오류로 나가던 문제. `.pptx`로 저장하라는 STOP_RETRY 입력 오류로 안내합니다. 다른 버전의 E_FAIL은 그대로 Office 오류로 보고합니다.
- Word 2007에 없는 `SaveAs2` 호출(새 파일 만들기, `word.py export --save *.docx`, undo의 Flat XML 복원)이 AttributeError로 실패하던 문제. 공통 `_office.word_save_as()`가 `SaveAs`로 대체합니다.
- `ppt.py write`의 `-` 글머리 목록이 "1. 2." 번호 목록으로 나오던 문제. 글머리 표시만 켜고 종류를 정하지 않아 PowerPoint 2007 글상자가 번호 목록으로 만들었습니다. 종류(글머리표/번호)를 명시하고, 마크다운과 `--text`의 `1.` 번호 목록도 구분합니다.
- `excel.py chart`의 분산형 종류가 틀리던 문제. `--type scatter`가 표식만이 아니라 선+표식(74)으로, `scatter_lines`(숫자 x축 `line` 자동 전환 포함)가 표식 없는 선(75) 대신 선+표식으로 그려졌습니다. `scatter` 상수를 -4169로 고치고, 종류를 데이터 지정 앞뒤로 모두 정합니다(앞: 첫 열을 x값으로, 뒤: Excel 2007이 바꾼 세부 종류 복원).
- 열린 워드 문서를 읽을 때 본문 그림 자리가 글자 `/`(Word 2007, 새 버전은 `\x01`)로 나오던 문제. `word.py read`와 `office.py transfer --from word`에서 그림 위치의 글자만 `[그림]`으로 바꿉니다(본문의 진짜 `/`는 유지).

### 테스트 러너 (`tools/run_win_tests.py`)
- 테스트 파일 준비가 Word 2007에서 `SaveAs2`로 멈추던 문제.
- 더블클릭으로 연 Excel은 포커스를 잃기 전까지 ROT에 등록되지 않아(`GetActiveObject` 실패, MS KB 238610) "A.xlsx가 열리지 않았습니다"로 중단되던 문제. 도구와 같은 창 핸들 바인딩(`_office.find_apps`)을 씁니다.
- 이전 실행의 스냅샷이 남아 같은 경로로 다시 만든 A 파일에 undo가 이어져 C7이 두 번째 실행부터 실패하던 문제. 도구 캐시/data를 테스트 폴더의 `_cache`(실행마다 비움), `_data`로 분리합니다.
- PDF 기능이 없는 PC의 D1, D4를 FAIL 대신 `ENV-SKIP`으로 기록합니다.
- I그룹 27개 추가 (모두 자동 판정): 엑셀 CSV/`--values`/`--append` 쓰기, `--search`·`--save`, 300행 잘림, 차트 종류·목록·삭제, 조건부 서식·숫자 서식·지우기, 시트 이름변경·복사·삭제, 범위 복사·숨김·틀 고정, 워드 그림·CSV 표·앵커 앞/바꾸기·제목 수준·쪽 나누기·목차·치환(굵게, `--map`, 바닥글), PPT 그림·카드·마크다운·긴 표·슬라이드 관리·표/그룹 치환·`.pptx` 사본, 엑셀 범위 -> 워드/PPT 표, 워드 표 -> 엑셀, redraw 강제, `undo --list/--pick`, 닫힌 파일 undo, 창 전환 시 최대화 유지. 집 PC(Office 2007): 60개 중 PASS 60(ENV-SKIP 2 제외).
- J그룹 3개 추가: `info --detail`로 PPT 전체 슬라이드, 엑셀 차트, 워드 그림의 배치 경고가 없는지 판정.
- 끝나면 A.pptx 모든 슬라이드와 A.xlsx 모든 차트를 `shots\`에 PNG로 저장하고, 테스트 폴더 문서와 실행 중 생긴 제목 없는 문서를 저장하지 않고 닫습니다(`--keep-open`으로 유지). 문서가 남지 않은 앱은 끝내고, `Quit`을 받고도 남는 빈 Excel/PowerPoint(이 실행 뒤에 생긴 것만)는 창 닫기로 정리합니다. 정리는 러너가 COM 참조를 놓은 뒤 별도 프로세스에서 합니다.

## [1.0.0] - 2026-10-27

첫 배포. 기존 마이크로 CLI 34개를 진입점 6개로 다시 설계했습니다. 이전 버전과의 CLI 호환은 없습니다.

### 구조
- 진입점 6개: `excel.py`, `word.py`, `ppt.py`, `outlook.py`, `reader.py`, `office.py`. 공통 모듈 2개: `_runtime.py`(응답 계약, 경로, 입출력), `_office.py`(COM 세션, 스냅샷, 색/테마).
- 저장소 ACI 계약 준수: stdout 1줄 JSON, exit 0/1/2/130, `VERIFIED:` / `STOP_RETRY:` / `ERROR:`, 인자 오류도 JSON.
- 닫힌 파일 읽기는 라이브러리(openpyxl, python-docx, python-pptx, pymupdf, olefile)로, Office를 띄우지 않습니다. 열린 문서와 쓰기는 COM입니다.
- 결과에 `target_resolved`(실제 작업한 문서), `next_commands`(다음 명령), 이름 오타 시 `[추천]`을 넣습니다.

### 기능
- 화면 규칙: 읽기는 사용자 화면(활성 시트, 스크롤, 슬라이드)을 움직이지 않습니다. 쓰기는 고친 시트/슬라이드로만 이동합니다(`--no-focus`로 끔).
- 쓰기 전 자동 스냅샷과 `office.py undo`(`--list`, `--pick`). 원본은 저장하지 않으며 복원 직전 상태를 백업합니다.
- 차트 전달 3방식(`export` / `redraw` / `relay`)을 자동으로 이어서 시도합니다(DRM으로 PNG가 암호화되는 환경 대응). `--probe`로 이 PC에서 되는 방식을 저장합니다.
- 엑셀: 범위 수식(열마다 수식), 머리글, 이어 쓰기, 테마, 실제 조건부 서식, 숫자 x축 꺾은선의 분산형 자동 전환.
- 워드: 마크다운(제목/목록/표/굵게)을 기본 스타일로 변환, 표는 ConvertToTable로 빠르게, 그림 비율 유지, 앵커 단락 앞/뒤/교체.
- PPT: 마크다운 제목마다 새 슬라이드, 긴 표는 머리글을 반복해 나눔, 그림은 남은 공간에 비율 유지, 카드.
- reader: 읽은 문서 자동 색인과 전체 검색(AND/OR/정규식), `--save`로 md/csv.
- outlook: 기본은 작성 창만 표시, `--send`는 명시 요청 시에만 수신자 확인 후 발송.

### 수정 (기존 코드 대비)
- 파일 대상 작업이 사용자 Excel/Word 창을 숨기던 문제 (전용 인스턴스 사용).
- 읽기 작업이 없는 시트/파일을 만들고 저장하던 문제, PPT 읽기만 해도 마지막 슬라이드로 이동하던 문제.
- Word 스냅샷이 원본을 저장하거나 경로를 바꾸던 문제, 활성 PPT는 스냅샷이 없던 문제, undo가 실패해도 성공으로 보고하던 문제.
- 한 행 범위 처리 오류, 범위 이동이 겹치면 데이터가 지워지던 문제, 치환이 수식을 덮어쓰던 문제.
- 워드 일반 텍스트 쓰기가 항상 실패하던 문제, 앵커 뒤 삽입이 단락을 쪼개던 문제, 빈 문자열 치환이 무시되던 문제.
- HWP 제어 문자 해석 오류, PPT 슬라이드 순서와 단어 쪼개짐, EML 한글 헤더와 전달 메일 본문 섞임.
- Outlook 검색 필터 구문 오류, 첨부 덮어쓰기, 서명 사라짐.

## 남은 할 일

### 1.0 제거 기능 (필요하면 다시 추가할 후보)
- 스타일 DNA 추출 (`office_theme --action extract`). 테마 프리셋은 유지.
- 검색 결과를 Office로 바로 넣기 (`search --to`). 현재는 `reader.py search --save` 후 `write --file`.
- 검색어가 파일명과 같으면 파일을 반환하던 처리.
- 샌드박스 가상 데스크톱 전환 (exebox, `SetThreadDesktop`).
- `clean fix-env`, JIT 자가복구의 "유령 프로세스 정리".
- 새 문서 만들기 / 파일 열기 (`office_window --new/--open`).
- EML 읽기·마스터 JSON 빌드·스레드는 Outlook 통합 도구에서 별도로 다룸 (reader.py에서는 1.3.0에서 제거).
- 워드 그림/도형 -> PPT 전달 (`office_bridge` word->ppt), PPT -> 워드 전달.

### 1.4.0 점검에서 남긴 것
- PowerPoint 닫힌 파일의 열기 암호: 사용자가 이미 띄운 PowerPoint로 열면 감시하지 않음(그 프로세스는 끝낼 수 없음). (DRM 인증 창 감시는 1.4.58)
- DRM 문서를 reader 캐시(평문)에 등록하기: 회사 보안 정책 결정이 필요해 하지 않음 (1.4.58 B4 2부).
- 엑셀: 열린 문서의 열 전체 범위(`A:A`)를 사용 범위로 자르기 (닫힌 파일·CSV는 1.4.5), 조건부 강조 `포함:`의 상대참조(활성 셀 기준), `--map` 연쇄 치환. (보호된 시트 사전 확인은 1.4.55) (잘못된 입력 검사 전에 새 시트가 만들어지는 문제는 1.4.13, 같은 이름 차트가 여러 시트에 있을 때 `[후보]`는 1.4.23)
- 워드: 열린 문서의 자동 번호 제목(`ListFormat.ListString`)이 `outline`/`--section`에 나오지 않음(닫힌 파일과 번호 체계 다름), 목차(TOC) 안의 앵커 후보 구별, 머리글 안 글상자 치환. (`replace --find`만 준 경우는 1.4.10)
- reader: HWP(바이너리) 개요 번호, PDF `find_tables` 성능(선이 있는 쪽만), 색인 저장 일괄화, BOM 없는 UTF-16 판별.
- transfer relay 실패 시 넣은 빈 단락 정리.
- "바뀐 것 없음"을 싸게 판정할 수 없어 늘 바꾼 것으로 보는 명령(되돌리기 지점이 남음): `excel style`(같은 서식 다시 적용, 서식 없는 범위 `--clear`),
  `range autofit/border/freeze`(같은 칸 다시 고정), 같은 값으로 `write --overwrite`. (1.4.15에서 replace·slide move·notes·range unhide/unmerge/unfreeze/clear 등은 처리)

### 23차 점검에서 남긴 것 (1.4.32)
- L2 성능: 범위를 생략한 닫힌 xlsx 읽기는 수식이 있는 시트 XML 전체를 한 번 더 해석함(30만 행 시트에서 수 초). 열린 값 읽기와 함께 한 번에 훑기.

### 1.4.37에서 남긴 것
- range sort: 사용자 지정 순서, 가로 정렬. 엑셀과 다른 글자 비교(기호 등)는 `UNVERIFIED:`로만 알림. (기준 4개 이상은 1.4.46)
- 실기(회사 PC)에서 outlook `--days`/`--since`의 Restrict 날짜 형식 확인.

### 1.4.67에서 남긴 것 (check 판단필요 과다)
- `Output voltage 5.0 V Typ., 4.75 V Min., 5.25 V Max.`(일치가 맞음), `정격 전압은 5 V이며 변동은 ±5 % 이내`(`변동은`이 한계 낱말로 안 잡힘), 기준 값 없는 `오프셋 전압 ±10 mV`
  (2차 말뭉치의 `허용 오차 ±0.1 mm`는 판단필요 기대라 함께 정해야 함), `리플 50 mVp-p`(괄호 없는 꼬리), `습도 45~85 %RH`, `기울기 ±0.5°`(각도 단위 없음), `approx.`,
  `각각` 나열(`입력 전압과 출력 전압은 각각 220 V, 110 V`)은 아직 판단필요.

### 확인 필요
- Windows 실기 검증: 회사 PC(Office 16)에서 1.4.1 실기 러너 71/71 통과. 남은 것: 1.4.5 쓰기 변경(write 덮어쓰기 확인, transfer 글자 유지)의 실기 확인, 눈 확인 항목.
- 실제 HWP 5.0 파일과 DRM 문서로 `reader.py` 검증. 실제 규격서(PDF/Word/HWPX)로 장절·쪽 인식 확인 (1.3.0).
- 회사 PC에서 `office.py transfer --probe` 결과 확인. (집 PC Office 2007: export, redraw 성공 / relay는 "워드에 붙여넣기 안 됨"으로 실패)
- 남은 Office 2007 실기 확인: `.xls/.doc/.ppt`(OOXML 아님)는 문서 식별값이 없어 같은 경로 새 파일 판별을 하지 않습니다.

### 저장소 지침 반영 (AGENTS.md, 2026-10-28)
상세 근거: `docs/history/COMPLIANCE_REVIEW_2026-10-28.md`. 공통 키·함수는 `CLI-dev/README.md` "공통 런타임 구현 규칙"(R1~R9)대로 구현. 우선순위 순.
- [ ] `recovered` 나머지: transfer 방식 대체, 오래된 스냅샷 정리 (잠금·기존 파일 새 이름 저장, 낡은 캐시는 1.4.0)
- [ ] `remove_tree` 추가, `os.remove`(_office.py:314, _runtime.py:311, word.py:593)에 적용. (1.4.0: undo 복원 복사는 `retry_fs`, retry_fs는 일시적 오류만 재시도 완료)
- [ ] COM 내보내기(excel.py, word.py, ppt.py export)의 잠금 대비 `save_with_lock_fallback`. (1.4.0: 기존 출력 파일은 새 이름 저장 + `recovered` 완료)
- [ ] `--save`로 기존 파일을 덮어쓰기를 원할 때의 `--overwrite` 인자 (기존 파일 새 이름 저장은 1.4.0)
- [ ] `--dry-run`: excel replace/range clear·move/sheet delete/chart delete --all, word replace/write --pos replace, ppt replace/slide delete·clear. (excel range delete·insert는 1.4.37) (1.4.0: 스냅샷이 없는 쓰기는 undo가 건너뛰지 않음, `undo --list`의 `undo_available: false`. 쓰기 결과에도 표시하는 것은 남음)
- [ ] 빈 결과 `count: 0` + "찾은 항목 없음": excel read 빈 범위, word/ppt read 0단위 (office windows·info --all·chart list·undo --list는 1.4.26, outlook list/search·ppt slide list는 1.4.27)
- [ ] 잘리는 결과에 `truncated` + 전체 결과 파일 경로: excel replace cells(30), undo --list(20), outlook list/search. `body_truncated` 옆에 `truncated` (excel read --search는 1.4.31)
- [ ] tools/run_win_tests.py `--agent` 모드: JsonArgumentParser와 최상위 예외 처리, 로그 태그 `[준비]` `[주의]` `[확인]` -> `[INFO]` `[WARN]`
- [ ] 인자 이름 정리 (베타라 별칭 없이 바꿈): `-r`/`-c`/`-n` 충돌, `--limit` 대 `--top`, `--pages` 대 `--slides`
