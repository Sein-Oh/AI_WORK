# office-cli 전체 검토 (2026-11-09, 1.4.91 기준)

`CLI/office-cli/` 전체를 영역별(엑셀 읽기·비교·대조, 엑셀 쓰기·COM·undo, 워드·PPT, reader, Outlook·공통 런타임)로 읽고 실행해 본 결과입니다.
검토 때는 저장소 파일을 고치지 않았습니다. 이 문서는 수정 작업을 이어받는 에이전트가 그대로 따라 할 수 있게 쓴 작업 목록입니다.

- 기준 커밋: main `e4ee10b` (office-cli 1.4.91)
- 테스트 기준선(클라우드, Linux): `python -m pytest -q CLI-dev/office-cli/tests CLI-dev/tests` -> 3395 passed, 3 failed.
  실패 3건은 `CLI-dev/tests`의 `test_run_all_tests.py` 2건·`test_setup.py` 1건이고, 원인을 확인한 것은 컨테이너에 scikit-learn·scipy·joblib이 없어서입니다
  (ml-kit 라이브러리를 설치하면 사라지는지 먼저 확인). office-cli 테스트는 모두 통과.
- 확인 표기:
  - **재현**: Linux에서 실제 명령으로 재현 (닫힌 파일 경로, reader, 공통 런타임)
  - **가짜 재현**: `CLI-dev/office-cli/tests`의 가짜 COM(`fake_excel.py` 등)으로 재현. 실제 Office 동작은 실기 확인 필요
  - **코드**: 코드 순서·경로를 읽어 확인
  - **추정 N%**: 실제 Office COM 동작에서 추론. 실기(집 PC·회사 PC) 확인 전에는 고치는 방향만 정하고, 확인 항목을 실기 러너에 넣는다
- 줄 번호는 1.4.91 기준입니다. 고치기 전에 그 근처를 다시 읽어 위치를 확인하십시오.
- 검토 때 쓴 임시 샘플·스크립트는 세션과 함께 사라졌습니다. 아래 재현 설명대로 테스트 안에서 다시 만듭니다.

## 0. 작업 방법 (이어받는 에이전트가 먼저 읽을 것)

1. `AGENTS.md`를 먼저 읽는다. 특히 "반드시 지킬 ACI 계약", "사용법 단순성", "문자 규칙", "작업 절차".
2. 항목마다: 회귀 테스트(실패 재현)를 먼저 쓰고 -> 고치고 -> `python -m pytest -q CLI-dev/office-cli/tests CLI-dev/tests`.
   동작이 바뀌면 같은 커밋에서 `SKILL.md`(6144바이트 한도, `test_skill_docs.py`가 검사)·`CLI-dev/office-cli/README.md`·`CHANGELOG.md`를 함께 고친다.
3. 버전은 묶음 단위로 올린다 (예: 1장 A 묶음 = 1.4.92, B 묶음 = 1.4.93). 스킬 5개의 `metadata.version`과 `_runtime.py`의 버전을 함께.
4. 버그 수정과 새 기능(4장)은 별도 커밋. 4장은 사용자와 범위를 정한 뒤에만 시작한다.
5. "추정" 항목은 실제 Office가 필요하다. 클라우드에서는 (a) 가짜 COM이 그 동작을 흉내 내게 고친 뒤 회귀 테스트를 쓰고,
   (b) `CLI-dev/office-cli/tools/run_win_tests.py`와 `TEST_CHECKLIST.md`에 실기 확인 항목을 추가한다. 확신이 낮은 것(70% 미만)은 실기 확인 뒤에 고친다.
6. 처리한 항목은 이 문서에서 지우지 말고 제목 끝에 `(1.4.NN에서 수정)`을 붙인다. `CHANGELOG.md` "남은 할 일"의 이 문서 줄은 전부 끝나면 지운다.
7. 테스트는 `CLI/office-cli/cache`, `CLI/office-cli/data`에 쓰지 않게 한다 (`OFFICE_CLI_DATA_DIR`와 `OFFICE_CLI_CACHE_DIR`는 따로다. 둘 다 임시 폴더로).

권장 순서: 1장 A1~A6, A10~A13 -> 2장 B1~B6 -> 1장 A7~A9(가짜 COM 보강 포함) -> 2장 나머지 -> 3장(실기 항목 추가) -> 5장 -> 4장(사용자 결정 뒤).

## 1. 틀린 결과를 VERIFIED로 내는 결함 (가장 급함)

경량 LLM은 `VERIFIED:`를 보면 그대로 사용자에게 전하고 다시 확인하지 않는다.

### A1. `read --where`의 크기 비교가 문자열 비교로 떨어짐 (재현, 높음) (1.4.92에서 수정)
- 위치: `excel.py:1000-1004` (조건 비교), `964` (`_where_num`)
- 재현: 머리글 `날짜, 비율, 코드`, 값 (2026-01-05 날짜 칸, 0.5 `0%` 서식, "010") / (2026-01-15, 0.75, "10") / (2026-01-25, 0.2, "0010")인 xlsx.
  - `read --target w.xlsx --where "비율 >= 50%"` -> `VERIFIED: 0건` (기대: 2행)
  - `--where "날짜 >= 2026-1-10"`, `"날짜 > 2026/01/10"` -> `VERIFIED: 0건` (기대: 2행). `2026-01-10`(0 채움)만 동작
  - `--where "금액 >= 1500원"` -> 1500 행이 빠짐
- 고칠 방향: `> < >= <=`에서 조건 값의 `%`(0.5와 50% 맞춤), 단위 꼬리(원, 개 등 숫자 뒤 글자), 느슨한 날짜(`2026-1-10`, `2026/01/10`, `2026.1.10`)를 해석한다.
  해석하지 못한 값으로 크기 비교를 하게 되면 문자열 비교로 답하지 말고 `STOP_RETRY`(무엇을 해석 못 했는지 + 가능하면 고친 전체 명령 `[추천]`).
- 테스트: 위 세 경우 + 해석 불가 값이 0건 VERIFIED가 아닌 것.
- 처리(1.4.92): 글 열끼리의 크기 비교(`이름 >= 다`, 기존 테스트)는 의도된 동작이라 남김. 그래서 '해석 못 한 값이면 STOP_RETRY' 대신 '그 열에 조건 값과 같은 종류로 읽히는 칸이 하나도 없으면 STOP_RETRY'로 함 (날짜 열에 `지난달`, 글 열에 숫자 조건). `[추천]`은 고칠 값을 도구가 알 수 없어 넣지 않음.

### A2. `--where "코드 = 010"`이 `010`, `10`, `0010`을 모두 같다고 봄 (재현, 중간) (1.4.92에서 수정)
- 위치: `excel.py:984-988`
- 재현: A1의 파일에서 `--where "코드 = 010"` -> 3행 (기대: 1행). office-data 스킬은 "품번의 앞자리 0은 그대로 다른 값"이라고 적고 있다.
- 고칠 방향: `=`/`!=`에서 조건 값이나 칸 값이 앞자리 0이 있는 글이면 글자 그대로 비교. 칸이 숫자형이면 지금처럼 숫자 비교.

### A3. 시트 XML의 `<dimension>`이 실제보다 작으면 앞부분만 읽음 (재현, 높음) (1.4.92에서 수정)
- 위치: `excel.py:301-302` (`_closed_bounds`). `read`, `export`, `check`가 씀
- 재현: 11행 xlsx의 `xl/worksheets/sheet1.xml`에서 `<dimension ref="A1:B11"/>`를 `A1:B5`(또는 `A1`)로 바꿔 다시 압축 (Excel이 아닌 프로그램이 만든 파일에 흔함).
  `read` -> 5행 VERIFIED, `--where "값 >= 50"`·`--search n9` -> `VERIFIED: 0건`, `export` -> 5행 "행 수 확인". `--range A1:B11`을 주면 정상.
- 고칠 방향: `_excel_compare.py:273`처럼 `ws.reset_dimensions()` 뒤 실제 범위를 구한다.

### A4. `search --pages`가 쪽 정보 없는 문서에서 `VERIFIED: 0건` (재현, 높음) (1.4.92에서 수정)
- 위치: `reader.py:3801-3806`, `4012-4020`
- 재현: `정격 전압은 5 V 이다.`가 든 TXT에 `reader.py search --query "정격 전압" --pages 1-3 --file a.txt` -> `count: 0`, `page_unknown_skipped: 2`, VERIFIED.
  TXT, CSV, XLSX, HWP, 쪽 나눔 없는 DOCX 모두 해당. `read --pages`는 같은 파일에서 항목 번호로 동작해 뜻도 다르다.
- 고칠 방향: `page_unknown_skipped > 0`이면 `not_searched`에 넣고 `UNVERIFIED:`. (쪽 없는 문서는 `--pages`를 무시하고 찾은 뒤 `warnings`로 알리는 쪽이
  "사용법 단순성 4번: 필요 없는 명령에서는 받고 무시"에 더 맞다. 둘 중 하나로 정하고 `read --pages`와 뜻을 맞춘다.)
- 처리(1.4.92): 무시하는 쪽은 `read --pages`(쪽 없는 문서는 항목 번호, 도구의 이어 읽기 명령도 그렇게 냄)와 뜻이 달라져, `search --pages`도 쪽 없는 문서는 항목 번호로 보고 `warnings`로 알림 (한 옵션 한 뜻). 쪽 문서 안의 쪽 모르는 단위만 `not_searched`.

### A5. 머리글 없는 표의 첫 행을 머리글로 써서 근거가 틀림 (재현, 높음) (1.4.92에서 수정)
- 위치: `reader.py:318-364` (`table_rows`, `_looks_like_header`)
- 재현 a: DOCX 2열 표 (모델명|PSU-100 / 입력 전압|220 V / 출력 전압|5.0 V)를 `read` -> 단위가 `모델명: 입력 전압 | PSU-100: 220 V`,
  `모델명: 출력 전압 | PSU-100: 5.0 V` 두 개뿐 (첫 행은 단위로 나오지 않음).
- 재현 b: PDF에서 다음 쪽으로 이어진 표 (2쪽 첫 행 `Ripple | 20 MHz BW | 50 mVp-p`) -> 다음 행들이 `Ripple: Efficiency | 20 MHz BW: 230 Vac | 50 mVp-p: 90 %`.
  `search --query "Ripple;50 mVp-p"`가 다른 행 2건을 VERIFIED(literal true)로 반환. `excel.py check`도 같은 행 글을 쓴다.
- 고칠 방향: (1) 머리글로 쓴 행도 단위로 낸다(검색에서 빠지지 않게). (2) 2열 표(항목|값 꼴)는 머리글 추정을 하지 않는다.
  (3) 앞 쪽 표와 열 수가 같은 이어진 표는 앞 표의 머리글을 다시 쓴다.
- 주의: 클라우드 기본 pymupdf가 사용자 환경(1.28.2)과 다르면 표 칸 글이 다르게 읽힌다. 표 테스트 전에 `python -m pip install pymupdf==1.28.2`.

### A6. PDF 본문의 줄바꿈된 줄이 장절 제목이 됨 (재현, 높음) (1.4.92에서 수정)
- 위치: `reader.py:585-590`, `classify_number` `177-187`, `_UNIT_AFTER` `117`
- 재현: "3. ELECTRICAL" > "3.2 Output" 아래 한 단락(같은 블록, 본문 글꼴) 안에 줄이 `1.5 times the rated current for`,
  `5.0 VDC nominal with decoupling of`, `0.1 uF capacitors placed near`, `1.5 N-m applied to each screw`로 시작하게 줄바꿈된 PDF를 `read`.
  그 줄들이 heading이 되어 단락이 쪼개지고, 뒤 단위의 근거가 `..., 5.0 VDC nominal with decoupling of > 1.5 N-m applied to each screw`.
  `5.0`이 깊이 1이라 "3. ELECTRICAL"이 닫혀 뒤의 "3.3 Protection"도 상위 장을 잃음. `search --query "limit of 1.5 times"`와 `--section 3.2` 검색은 `VERIFIED: 0건`.
- 고칠 방향: PDF에서 블록의 첫 줄이 아닌 줄은 (a) 본문보다 큰 글씨·굵게이거나 (b) 앞 줄이 문장 끝일 때만 제목 후보. 번호가 지금 장절 흐름과 이어지는지(3.2 다음에 5.0은 아님)도 본다.
  `_UNIT_AFTER`에 `VDC`, `uF`, `N-m`, `times` 같은 꼬리도 보탠다.

### A7. 통화 서식 칸(Decimal)을 숫자로 보지 않음 (가짜 재현, 높음. 실제 Excel이 Decimal을 준다는 전제는 추정 85%) (1.4.94에서 수정, 실기 확인 항목 있음)
- 위치: 뿌리 `_office.py:1169` (`range_to_matrix`). 증상 `excel.py:3319` (`_is_num`), `3349` (`_sort_class`), `3498`, `964`
- 재현(가짜 COM에 `decimal.Decimal` 값): B열 10, 100, 9에 `range --action sort --range A1:B4 --by B` -> "이미 정렬되어 있어 바꾸지 않았습니다" VERIFIED.
  머리글 자동 판별도 실패. 열린 문서 `read --where "금액 >= 10"` -> `VERIFIED: 0건`.
- 고칠 방향: `range_to_matrix`에서 `Decimal` -> `float`. 실기 러너에 통화 서식 열 정렬 항목 추가.

### A8. `+`/`-`로 시작하는 글이 숫자·수식으로 바뀌어 써지고 검증을 통과 (검증 구멍은 가짜 재현, 변환은 추정 80%) (1.4.94에서 수정, 실기 확인 항목 있음)
- 위치: `excel.py:174-181` (`xl_text`), `2605-2618` (`_same_cell`)
- 재현: `write --range A1 --value "+82-10-1234-5678"` -> Excel이 계산해 -6840. `-007` -> -7, `- 검토 필요` -> `#NAME?` 또는 COM 오류
  ("시트 보호·병합된 칸인지 확인"이라는 틀린 안내). `_same_cell`이 다시 읽은 값이 글이 아니면 통과시켜 VERIFIED.
  `2024-01`, `true`, 글로 저장된 숫자 칸의 `replace`도 같은 구멍.
- 고칠 방향: `coerce()`가 글로 남긴 값 중 `+ - @ =`로 시작하는 것은 앞에 `'`를 붙여 글로 쓴다(수식은 `--formula`·`--cells`의 `==` 경로만).
  글을 쓰려 했는데 다시 읽은 값이 글이 아니면 `UNVERIFIED:`(무엇이 바뀌었는지).

### A9. 수식은 글자만 다시 읽어 확인 (코드, 높음. 동적 배열은 추정 85%) (1.4.94에서 수정, 실기 확인 항목 있음)
- 위치: `excel.py:1410`, `2627-2631`
- 증상 a: `--formula "=합계(B2:B3)"`나 함수 이름 오타 -> 칸은 `#NAME?`인데 VERIFIED.
- 증상 b: Office 2021에서 `=UNIQUE(A2:A50)`, `FILTER`, `SORT`, `SEQUENCE`를 `Range.Formula`로 쓰면 암시적 교차(@)로 들어가 펼쳐지지 않는다.
- 고칠 방향: 쓴 뒤 값(`Value`)도 읽어 오류 값(`#NAME?`, `#REF!`, `#VALUE!` 등)이면 `UNVERIFIED:` + 어느 칸이 어떤 오류인지. `Formula2`가 있으면 그것으로 쓴다
  (없는 버전 대비 `try`). 가짜 COM에 오류 값 흉내를 넣는다.
- 처리(1.4.94): `Formula2`는 동적 배열 함수가 든 수식에만 씀 (모든 수식을 Formula2로 쓰면 `=A:A`처럼 예전에 한 칸이던 수식이 펼쳐져 뜻이 바뀜).
  A8의 `2024-01`, `true`는 Excel이 날짜·논리값으로 바꾸는 것이 사용자 의도일 때가 많아 글로 고정하지 않고, 다시 읽은 값이 그 종류일 때만 인정.

### A10. Outlook `list`/`search`가 항목 읽기 실패를 조용히 건너뛰고 "0건" (가짜 재현, 높음) (1.4.92에서 수정)
- 위치: `outlook.py:153-161` (`_collect`의 `except Exception: continue`)
- 재현: 모든 항목에서 `SenderEmailAddress`가 예외(Outlook 보안 창에서 "거부"를 누른 경우와 같음) -> `VERIFIED: 0건`, 다시 실행하지 말라고 함.
- 고칠 방향: `cmd_calendar`의 `bad`처럼 실패 수를 세어 일부 실패면 `UNVERIFIED:`(읽지 못한 메일 N개), 전부 실패면 오류.

### A11. 참조(CC) 수신자가 발송 전 확인 화면에 나오지 않음 (재현, 중간. 메일 오발송 위험) (1.4.92에서 수정)
- 위치: `outlook.py:870` (`_finalize`), `956` (`cmd_send`)
- 재현: 초안에 CC가 있는 `reply --all --body ok`, 또는 `new --cc c@x.com` -> 결과·`preview`에 `to`만. 사용자는 그것만 보고 "예".
- 고칠 방향: 결과와 `preview`, `send`의 확인 글에 `cc`(있으면 `bcc`)를 넣는다. office-mail 스킬 발송 순서 2번은 "받는 사람·참조·제목"으로.

### A12. `read --save` -> `write --file` 왕복이 내용을 망가뜨림 (재현, 높음) (1.4.92에서 수정)
- 위치: `word.py:312-325`, `ppt.py:209-221` (둘 다 `reader.units_to_markdown`, `reader.py:2001`), `_markdown.py:94` (단락 잇기)
- 재현 a: 5단락 docx를 `word.py read --target small.docx --save`, 결과의 `ppt.py write --file <md>` -> 슬라이드 6장, 제목이 "단락 N - <제목>". 30단락을 넘으면 "슬라이드 N장 분량"으로 멈춤.
- 재현 b: 같은 md를 `word.py write --file` -> 단락마다 "단락 N" 제목 2가 붙고, 실제 제목 "1. 개요"는 목록 항목 "개요"(번호 사라짐), "2. 결론"은 글자 그대로.
- 재현 c: `ppt.py read --save` -> `word.py write --file` -> 슬라이드의 줄들이 한 단락으로 합쳐지고 표 행 ` | v`가 따로 1x1 표.
- 이 경로는 office-doc 스킬 규칙 6번이 권하는 방법이다.
- 고칠 방향: `--save`가 만드는 md를 `_markdown.py`가 그대로 되읽는 형식으로 쓴다: 실제 제목은 `#` 수준으로, 본문 단락은 빈 줄로 나눈 한 줄씩, 표는 마크다운 표,
  "단락 N" 같은 위치 표지는 제목으로 넣지 않는다. 회귀 테스트는 "저장 -> 다시 쓰기 -> 단락 수·제목·표 모양이 같음"으로.

### A13. 검색 범위가 빠졌는데 VERIFIED (재현, 중간) (1.4.92에서 수정)
- 위치: `reader.py:4096` (`_search_gaps`), `3729-3731`, `477`, `1203`, `4987`
- 경우:
  1. 원본이 지워진 문서의 캐시 결과가 "VERIFIED: 검색 완료" (항목의 `file_missing: true`와 `missing_file_count`뿐). 폴더에 새로 넣은 파일은 `--file` 없이는 검색되지 않는다.
  2. 폴더의 `.doc/.xls/.ppt`는 `skipped_types`에만 있고 `not_searched`에 없어 `VERIFIED: 0건`.
  3. DOCX 각주·머리글·바닥글은 읽지 않는다. PDF는 반복 머리글·바닥글을 지워 문서번호·개정 표기가 검색되지 않는다.
  4. 쪽 경계에 걸친 문장(`humidity shall not exceed 95`가 p.1/p.2로 나뉨)은 `VERIFIED: 0건`.
  5. `diff`: 일부 쪽만 스캔인 PDF(`pages_without_text`)도 VERIFIED (`search`는 UNVERIFIED). `ocr_needed`만 본다.
- 고칠 방향: 1·2·5는 `not_searched`에 넣어 `UNVERIFIED:`. 3은 각주·머리글을 읽어 단위로 넣거나(권장) 읽지 않은 사실을 `not_searched`에. 4는 쪽 끝·다음 쪽 첫 단위를 이어 붙여 한 번 더 찾는다.
- 처리(1.4.92): 1의 '폴더에 새로 넣은 파일'은 `--file` 없는 검색이 색인만 보는 설계라 그대로 둠 (색인에 폴더 목록이 없어 새 파일을 알 수 없음, 스킬은 `--file 폴더`를 씀). 2는 문서 형식(.doc .xls .ppt .rtf .xlsb .odt .ods .odp)만 `not_searched` (그림·압축 파일까지 넣으면 모든 폴더 검색이 UNVERIFIED). 3은 단위로 넣음 (PDF 반복 머리글은 글마다 한 번).

## 2. 계약 위반·내부 오류·실행 못 하는 추천

### B1. 데이터 폴더를 만들 수 없으면 stdout에 JSON이 없음 (재현, 높음) (1.4.93에서 수정)
- 위치: `_runtime.py:825`, `832` (`continue_clear` -> `_continue_state:744` -> `get_data_dir` -> `os.makedirs:1139`)
- 재현: `OFFICE_CLI_DATA_DIR`를 파일 경로로 주고 `excel.py read --target x.xlsx --range A1:B2` -> traceback, stdout 0줄. 읽기 전용 `data` 폴더도 같은 경로.
- 고칠 방향: `_continue_state`/`continue_clear`가 예외를 삼키고, `run_cli`가 어떤 경우에도 JSON 1줄을 내게 가장 바깥에서 한 번 더 막는다.

### B2. `[추천]` 뒤가 실행할 수 없는 글 (재현, 중간) (1.4.93에서 수정)
- `outlook.py:1434`: `calendar --days 0` -> `[추천] 7`. 고칠 방향: `rt.retry_command({'--days': 7})`.
- `outlook.py:617`: `attachments --name a`(--save 없음) -> `[추천] 같은 명령 끝에 --save를 붙이세요.` 고칠 방향: 지금 인자에 `--save`를 붙인 전체 명령.
- `excel.py:2424-2425`: `export --encoding cp949` + CP949로 못 쓰는 글자 -> `[추천] utf-8-sig`. 고칠 방향: `rt.retry_command({'--encoding': 'utf-8-sig'})`.
- `reader.py:2397`: index 결과 `next_commands`의 `search --query '<검색어>'`가 `needs_input` 없이 나가 그대로 실행하면 `VERIFIED: 0건`. 고칠 방향: `needs_input: true`.
- 함께: 저장소 전체에서 `[추천]` 뒤가 `python `으로 시작하지 않는 곳을 찾는 테스트를 `test_contract.py`에 추가한다 (소스의 `"[추천] "` 뒤 글을 정적 검사).

### B3. `--rows`를 주면 값 없는 `--max-rows`를 추천 (재현, 중간) (1.4.93에서 수정)
- 위치: `_runtime.py:851` (`SAME_MEANING`의 `--rows` -> `--max-rows`), 원인 제공은 office-excel 스킬의 "(행 수 `rows`, 수식 보기 `--formulas`)" 표기
- 재현: `excel.py read --target sales.xlsx --rows` -> `[추천] ... read --target sales.xlsx --max-rows` -> 실행하면 "expected one argument".
  `read ... rows`(스킬 표기를 인자로 읽은 경우)는 `[추천]` 없이 인자 오류.
- 고칠 방향: 값 없는 `--rows`와 따로 떨어진 `rows`는 "행 수는 결과의 `rows`"라는 뜻이므로 그 인자를 뺀 전체 명령을 `[추천]`. 같은 뜻 표의 일반 규칙으로
  "값이 필요한 옵션으로 바꿀 때 값이 없으면 추천하지 않음"도 넣는다. 스킬 문구는 "행 수는 결과의 `rows` 값"처럼 옵션으로 읽히지 않게.

### B4. 수식 여러 개를 칸 하나에 주면 `[추천]`이 100행 범위를 지어냄 (재현, 중간) (1.4.93에서 수정)
- 위치: `excel.py:1844`
- 재현: 데이터 2행인 표에 `write --range D2 --formula "=B2*2" "=C2*3"` -> `[추천] ... --range D2:E101` -> 그대로 실행하면 수식 200개, VERIFIED.
- 고칠 방향: 옆 열의 사용 범위에서 마지막 행을 구해 그 범위로 추천. 못 구하면 `[후보]`(사용자에게 범위를 물음).
- 처리(1.4.93): 못 구하면 `[후보]` 대신 `[추천]` 없이 "사용자에게 마지막 행을 물어" (고를 값 목록이 없는 질문이라 `[후보]`·`confirm_commands`를 만들 수 없음).

### B5. `-`로 시작하는 값이 `[추천]` 없는 인자 오류 (재현, 중간) (1.4.93에서 수정)
- 위치: `_runtime.py` 인자 오류 처리, `reader.py:5058` 등
- 재현: `reader.py search --query "-40℃" --file a.txt`, `outlook.py reply --id 3 --body "-확인"` -> "expected one argument". `--query=-40℃`나 `"-40 ℃"`(띄어 씀)는 동작.
- 고칠 방향: 파싱 전에, 글 값을 받는 옵션(`--query`, `--body`, `--value`, `--text`, `--find`, `--replace`, `--subject` 등) 바로 뒤 토큰이 `-`로 시작하고
  그 도구의 옵션 이름이 아니면 `--opt=<값>`으로 합친다 (도구가 스스로 고침. 사용법은 그대로 한 가지).

### B6. 차트 시트가 있는 통합문서에서 `ERROR:` (재현, 중간) (1.4.93에서 수정)
- 위치: `excel.py:438-440` (`'Chartsheet' object has no attribute 'max_row'`), `_excel_compare.py:251-252` (`iter_rows`)
- 재현: openpyxl `wb.create_chartsheet()`가 든 xlsx. 차트 시트가 활성이거나 `--sheet`로 지정한 `read`, 시트 지정 없는 `compare`.
- 고칠 방향: `wb.worksheets`만 돈다. 차트 시트를 직접 지정하면 `STOP_RETRY`(워크시트 `[후보]`).

### B7. `XFD1048576`에 떠돌이 칸이 있으면 멈추거나 MemoryError (재현, 중간) (1.4.95에서 수정)
- 위치: `_excel_compare.py:317` (끝 열 자르기가 너비 x 행), `323` (모든 행을 전체 너비로 채움), `excel.py:463` (`CLOSED_MAX_CELLS` 자르기가 `--range`를 줄 때만)
- 재현: 데이터 10행 + `XFD1048576`에 `" "`. `compare`는 140초 넘게 JSON 없음, 범위 없는 `read`는 MemoryError.
- 고칠 방향: 행마다 마지막 채워진 열을 한 번에 구한다. 범위 없는 읽기에도 칸 수 상한을 적용. 빈 값뿐인 끝 행·열은 범위에서 뺀다.

### B8. 쓰기가 아닌 명령도 경로 오타면 빈 새 파일을 만들고 VERIFIED (가짜 재현, 중간) (1.4.95에서 수정)
- 위치: `excel.py:268-272` (`write_session`이 모든 쓰기 명령에 `create=True`), `_office.py:938`
- 재현: `style --target <없는 경로>.xlsx --range A1:C3 --bold` -> 빈 파일 생성, "서식 적용" VERIFIED (`replace`, `range sort/freeze`, `sheet add`, `chart delete`도).
- 고칠 방향: `create=True`는 `write`(와 엑셀로 보내는 `transfer`)만. 나머지는 "파일 없음" + 비슷한 이름 `[후보]`. word.py·ppt.py도 같은지 확인.
- 처리(1.4.95): word.py·ppt.py는 이미 `write`만 `create=True`라 고칠 것 없음 (`transfer`는 office.py가 따로 엶).

### B9. 출력에 반쪽 이모지(lone surrogate)가 있으면 JSON 없이 죽음 (재현, 중간) (1.4.93에서 수정)
- 위치: `_runtime.py:66-73` (`setup_console`이 이미 UTF-8이면 `reconfigure`를 건너뜀), `305-309`, `739` (`_continue_key`)
- 재현: `PYTHONIOENCODING=utf-8`(전역 지침의 인코딩 줄)에서 payload나 argv에 `\ud83d` -> `UnicodeEncodeError`, stdout 0줄. cp949에서는 `?`로 나가 정상.
- 고칠 방향: 늘 `reconfigure(errors="replace")`. `_continue_key`는 `encode("utf-8", "replace")`.

### B10. `--regex`에 시간 보호 없음 (재현, 낮음~중간) (1.4.95에서 수정)
- 위치: `reader.py:2509`
- 재현: 5쪽 PDF에 `search --query "(\w+\s?)+$X" --regex` -> 25초가 지나도 끝나지 않음. 120초 제한에 걸리면 JSON 없음.
- 고칠 방향: 중첩 수량자(`(...+)+`, `(...*)+`)를 사전에 거절(`STOP_RETRY`)하거나, 단위별 시한을 두고 넘으면 `UNVERIFIED:`.
- 처리(1.4.95): 파이썬 `re`는 한 번의 일치를 중간에 멈출 수 없어 단위별 시한은 효과가 없음. 사전 거절로 하되, 바깥 반복 안에 반드시 나오는 한 글자(구분 글자)가 있으면(`(\d{1,3},)+`) 되추적이 커지지 않아 받음.

### B11. 종류를 알 수 없는 PPTX 도형 하나로 파일 전체가 `ERROR:` (재현, 낮음~중간) (1.4.95에서 수정)
- 위치: `reader.py:1272` (`shape.shape_type`이 `NotImplementedError`)
- 재현: `prstGeom`도 `txBox`도 없는 `p:sp`가 든 pptx를 `read`.
- 고칠 방향: `shape_type` 대신 `hasattr(shape, "shapes")`, 또는 도형 단위 `try`. ppt.py의 닫힌 파일 읽기도 같은 코드를 쓰는지 확인.

### B12. 잘못된 주소가 `ERROR:`(내부 오류) (가짜 재현, 낮음~중간) (1.4.95에서 수정)
- 위치: `excel.py:2758`, `3781`, `2881`
- 재현: `range --action copy --dest 요약`, `chart --at 요약!E1`, `range --action freeze --range 3:3`(실패 전에 기존 고정을 이미 풂).
- 고칠 방향: 주소 검사를 변경 전에 하고 `STOP_RETRY`.

### B13. 그 밖 (모두 재현, 낮음) (1.4.95에서 수정)
- `read --save res.xlsx`가 경고 없이 `res.csv`로 저장 (`excel.py:832`). reader `search`/`read --save x.docx`도 경고 없이 .md. (`compare`·`diff`는 `warnings`로 알림. 같게 맞춘다)
- `read --search "1,000"`, `"25%"`가 표시 형식 숫자를 못 찾고 `VERIFIED: 0건` (`excel.py:513`, 원시 값만 찾음).
- `read --range B:B`의 `rows`가 서식만 있는 끝 행까지(1048576) 셈. 범위 없는 읽기는 자름. 스킬은 행 수를 `rows`로 알리라고 함.
- 기간 칸(`[h]:mm`)이 `"1 day, 6:00:00"`으로 나옴.
- xlsx가 아닌 zip, 깨진 시트 XML이 `ERROR:` (KeyError/ParseError). `STOP_RETRY`여야 함.
- `compare --target s.xlsx --target2 s.xlsx --sheet Sheet`가 자기 자신과 비교해 `VERIFIED: 0건. 바뀐 것 없음` (`_excel_compare.py:673-675`, `694`). 같은 경로·같은 시트면 거절.
- `check --target x.csv`가 CSV에 메모 다는 `next_commands`를 냄 (`_spec_check.py:280`. `compare`는 `not new["csv"]`로 막음).
- `write --range C5:A1 --value x`(뒤집힌 범위)가 "-3행 x -1열"로 보고하고 아무것도 확인하지 않음 (`excel.py:1791`, `2633`).
- `style --align`/`--border`만 주면 VERIFIED지만 둘 다 다시 읽지 않음 (`excel.py:2372`).
- `sheet --action select`를 닫힌 파일에 하면 "선택했습니다" VERIFIED인데 남는 것이 없음 (`excel.py:2540`).
- `--append`: 범위 없이 표가 A열에서 시작하지 않으면(B3:C5) A6에 머리글과 함께 씀 (`excel.py:1147`, `1789`).
- `read --pages 2,3,4,5,6,300`의 이어 읽기 `next_commands`가 `--units '10-600'`(요청하지 않은 쪽 포함) (`reader.py:2275-2291`).
- 첨부 차단 확장자 검사가 이름 끝의 공백·점으로 우회됨: `evil.exe `, `a.js.` (`outlook.py:422`와 `_unique_dest:566-568`). 정리한 뒤의 이름으로 검사.
- 오류 로그에 `--folder 'hong@corp.com\받은 편지함'`의 계정 주소가 그대로 남음 (`REDACT_OPTIONS:494`에 `--folder` 없음).
  - 처리(1.4.95): `--folder`를 통째로 가리면 기존 설계(기본 폴더 별칭 경로 `inbox\없음`은 남겨 원인을 보게 함, 시험 있음)가 깨져, 로그의 명령·결과에서 메일 주소 모양만 `<메일 계정>`으로 바꿈.
  옵션 없이 준 본문·검색어(`reply --id 3 "확인했습니다"`)도 `명령:`·`결과:`에 남고, 그때의 안내("띄어쓰기... 큰따옴표로")도 틀림(빠진 것은 `--body`).
- 다른 폴더의 같은 이름 파일은 evidence가 같아 구분되지 않음 (`spec.txt 단락 2`).
- 전부 스캔인 PDF를 index해도 경고 없이 "등록 완료, 검색 가능" (`total_units: 0`).

## 3. 실제 Office 동작에서 추정한 결함 (실기 확인 필요) (1.4.96에서 처리: 표 끝 열 참고)

클라우드에서는 가짜 COM을 실제 동작에 가깝게 고치고 테스트를 쓴 뒤, 실기 러너에 확인 항목을 넣는다.

| # | 증상 | 위치 | 확신 |
|---|---|---|---|
| C1 | 워드 `write --anchor "<표 바로 위 단락>" --pos after`(또는 `replace`)가 표의 첫 칸 안에 씀. `para.End`가 표 첫 칸의 시작이라서. 마크다운 표면 표 안에 표. `_verify_inserted`는 통과 | `word.py:1021`, `1006-1015` | 90% (1.4.96에서 수정, 실기 L4) |
| C2 | 두 번째 번호 목록이 앞 목록에 이어 4, 5로 매겨짐 (`WD_STYLE_LIST_NUMBER`만 씀). `ListFormat.ApplyListTemplate(..., ContinuePreviousList=False)` 필요 | `word.py:742-745` | 90% (1.4.96에서 수정, 실기 L5) |
| C3 | 기존 표 바로 뒤에 쓴 표·요약 상자가 앞 표와 합쳐지고, 머리글 서식이 사용자 표의 첫 행을 바꿈. 표 뒤면 빈 단락을 먼저 | `word.py:747-773`, `775-790` | 80% (1.4.96에서 수정, 실기 L6) |
| C4 | `image()`, `toc()`, `callout()`이 단락 스타일을 되돌리지 않아 그림 단락이 제목·글머리 기호가 됨 (`outline`·목차에 "[그림]") | `word.py:792-819`, `775` | 80% (1.4.96에서 수정, 실기 L7) |
| C5 | `write --toc --file report.md`가 본문을 넣기 전에 목차를 만들고 갱신해 빈 목차인데 VERIFIED. 이후 쓰기·`export`도 목차를 갱신하지 않음 (코드 순서로 확인) | `word.py:1056-1077`, `812-818` | 코드 (1.4.96에서 수정, 실기 L8) |
| C6 | 변경 추적이 켜진 문서에서 `replace`/`fill`: 지운 글이 Find에 계속 걸려 "N곳 남아 있음" UNVERIFIED, 다시 실행하면 중복 치환. `read`는 옛 글과 새 글이 붙어 나옴 | `word.py:1280`, `1733` | 60% (코드 그대로, 실기 L9만) |
| C7 | `--pos replace`가 내용을 넣기 전에 앵커 단락을 지움. 뒤에서 COM이 실패하면 단락만 사라짐 | `word.py:1012` | 코드 (1.4.96에서 수정) |
| C8 | PPT 새 글상자가 슬라이드 밖에 놓임: `next_top()`이 모든 도형의 가장 아래를 기준으로 삼는데 슬라이드 번호·바닥글·날짜 개체가 바닥에 있음. `body()`·`table()`은 공간 확인 없음 (가짜 재현: Top 500 개체 -> 540 슬라이드에 top 542) | `ppt.py:420-430`, `467`, `489` | 가짜 재현 (1.4.96에서 수정, 실기 L10) |
| C9 | `--target` 없는 쓰기가 다른 통합문서에 씀: 제한된 보기(ActiveWorkbook 없음)나 숨은 PERSONAL.XLSB만 있을 때 `Workbooks(1)`로 넘어감. 쓰기 세션은 넘어가지 말고 "편집 사용" 안내 | `_office.py:195-209` | 70% (1.4.96에서 수정, 실기 L11) |
| C10 | undo: 저장된 경로의 열린 통합문서를 되돌리면 스냅샷(사용자의 저장 안 한 편집 포함)을 디스크 파일에 덮어씀. "저장하지 않고 닫기"로 원래 파일을 되찾을 수 없고 `preview`에 그 말이 없음 | `office.py:1059-1073` | 코드 (1.4.96에서 수정, 실기 L12) |
| C11 | `--append`의 `End(xlUp)`이 필터·숨긴 행을 건너뛰어 숨은 데이터 행을 대상으로 잡음 | `excel.py:1147` | 75% (1.4.96에서 수정, 실기 L13) |
| C12 | 닫힌 파일 `range --action freeze`가 VERIFIED인데 저장되지 않을 수 있음 (dirty일 때만 저장) | `_office.py:1063` | 55% (코드 그대로, 실기 L14만) |
| C13 | OneDrive/SharePoint 통합문서(FullName이 URL)를 제목 없는 문서로 봐서 undo가 저장 없이 닫음 | `_office.py:408`, `office.py:979` | 낮음 (코드 그대로, 실기 L15만) |
| C14 | Outlook `received`가 현지 시각인데 `+00:00`으로 표기 (pywin32가 현지 벽시계에 UTC tzinfo를 붙임). `_when()`으로 표기 | `outlook.py:134` | 85% (1.4.96에서 수정, 실기 M1) |
| C15 | `--id` 없는 `current`/`reply`/`forward`/`attachments`가 뒤에 열어 둔 메일 창을 고름 (`ActiveInspector()`). `app.ActiveWindow()`가 Explorer인지 먼저 확인 | `outlook.py:668-681` | 85% (1.4.96에서 수정, 실기 M2) |
| C16 | `search --from <smtp 주소>`가 사내 Exchange 보낸 사람을 못 찾음 (X.500 DN). `PR_SMTP_ADDRESS`도 대조 | `outlook.py:353`, `337` | 70% (1.4.96에서 수정, 실기 M3) |
| C17 | `Display()` 실패(Outlook에 대화 상자가 떠 있음) 때 초안은 이미 저장됐는데 일반 오류. 다시 실행할 때마다 초안이 늘어남. 초안 정보와 함께 `UNVERIFIED:` | `outlook.py:875-876` | 코드 (1.4.96에서 수정, 실기 M4) |
| C18 | 서명 로고 같은 본문 그림이 첨부로 셈 (`PR_ATTACHMENT_HIDDEN`만 봄, Content-ID만 있는 경우) | `outlook.py` 첨부 목록 | 낮음 (코드 그대로, 실기 M5만) |

마크다운 처리의 작은 것 (재현, 낮음. 사용자와 범위를 정한 뒤):
- 표 칸의 `**굵게**`가 워드·PPT 표에 별표 그대로 (`word.py:749`, `ppt.py:502`). LLM은 합계 행을 굵게 쓰는 일이 잦다.
- `*기울임*`, 백틱, 링크, `>`는 글자 그대로. 중첩 목록은 한 수준으로. `1. a / - sub / 2. b`에서 `2. b`가 일반 글이 됨 (`_markdown.py`).
- `--text "## 결론\n수율 98%\n불량 2%"`는 두 줄이 한 단락으로 합쳐지고, 제목이 없으면 두 단락 (마크다운으로 보이는지에 따라 달라짐).

## 4. 부족한 기능 (새 기능. 사용자와 범위를 정한 뒤 별도 커밋)

우선순위는 경량 LLM이 스스로 못 하는 일 순.

1. **표 집계** (가장 먼저) (1.4.98: `excel.py stats`, 닫힌 파일 `info`. 열 선택·정렬해서 읽기는 남음): 합계, 평균, 건수, 그룹별 집계, 상위 N, 고유값. 지금은 `--where`가 거르기만 해서 LLM이 최대 200행을 직접 더해야 한다.
   office-data 스킬이 "표 집계도 앞으로 이 스킬에 둔다"고 적고 있다. 닫힌 파일·CSV에서 Office 없이 동작해야 하고, 결과는 `preview`에 숫자를 글로.
   함께: 닫힌 파일의 시트 목록(지금 `info`·`sheet --action list`는 COM 필요), 열 선택, 정렬해서 읽기.
2. **`--where` 보강**: 제목 행 아래의 머리글 찾기(지금은 제목을 열 이름 `[후보]`로 냄, `excel.py:1012-1013`. `compare`의 머리글 찾기를 재사용),
   범위(이상~이하), 목록(여러 값 중 하나), "포함 안 함"(`!~`가 지금은 열 이름 `상태 !`로 읽힘).
   `compare`도 품번 열이 없으면 늘 `--key`를 묻는다: 두 파일에서 값이 겹치지 않는 열이 하나면 그것을 쓴다. 제목 행 2칸을 머리글로 잡는 문제(`_excel_compare.py:208-211`).
3. **기존 내용 고치기**: PPT 기존 본문·도형의 글 바꾸기(지금은 쓸 때마다 새 글상자), 도형 하나 지우기·옮기기. 워드 표의 행·열 넣기·지우기, 단락 지우기,
   머리글·바닥글·쪽 번호(원본에), 변경 내용 수락·거부, 메모 답글·지우기, 기존 글의 정렬·글꼴·스타일.
4. **엑셀 일상 기능**: 열린 문서 저장·다른 이름으로 저장, 자동 필터 걸기·풀기, 중복 찾기·지우기, 시트 간 조회(VLOOKUP류)·합치기, 피벗, 데이터 유효성(목록),
   인쇄 설정, 시트 숨기기·순서, 기존 차트 고치기, 수식을 값으로 바꾸기. `info`에 숨긴 시트·필터·표·보호 여부.
5. **메일·일정**: 일정 만들기·고치기·회의 요청, 메일 이동·플래그·읽음 표시·분류, 전체 폴더 검색, `--until`, 첨부 있는 메일만, `--bcc`,
   `--body-file`의 .md를 서식으로 넣기(지금은 `**굵게**`가 글자 그대로인데 스킬은 .md를 쓰라고 함), 회의 요청·배달 보고가 목록에서 빠짐(`Class == 43`만).
6. **문서 읽기**: 스캔 PDF OCR, `.msg/.eml`, DOCX 각주·미주·머리글·메모, PPTX 차트·SmartArt 글, XLSX 표시 형식(5%가 0.05)·칸 메모, 장절 목록만 주는 명령
   (reader에 `outline`이 없어 큰 문서는 8,000자씩 이어 읽어야 함), 큰 PDF 한 파일 안의 시간 나누기(350쪽 19초, 파일 사이에서만 `CONTINUE`).
7. `diff`: 옮겨진 단락을 삭제+추가로만 내고 그 안의 숫자 변경을 놓침 (`reader.py:4931-4948`). 남은 삭제·추가를 유사도로 다시 짝지어 "이동"/"변경"으로.

## 5. 경량 모델 운용 관점

1. **Gemma 실측이 얇다** (1.4.97: O6~O8 추가): office-cli 자동 시나리오는 O1~O5 다섯 개(대부분 닫힌 파일)다. 1장을 고친 뒤 `docs/opencode/SCENARIOS.md`에 조건 검색(날짜·%),
   규격 검색(음수 온도), 문서 왕복(read --save -> write --file), 메일 참조 확인을 넣는다.
2. **0건의 "다시 실행하지 마세요"** (1.4.92~1.4.95에서 처리): 조건·검색어를 도구가 해석하지 못해 0건인 경우까지 LLM이 멈춘다. 해석 실패는 `STOP_RETRY` + `[추천]`, 진짜 0건만 VERIFIED.
3. **대상 문서 옵션 이름** (사용자 결정 대기): 엑셀·워드·PPT는 `--target`, reader는 `--file`, 엑셀·워드 쓰기의 `--file`은 내용 파일. office-spec 스킬이 둘을 섞어 쓴다.
   같은 뜻 표가 고쳐 주지만 왕복이 한 번 는다. 이름을 하나로 맞출지는 사용자 결정 (베타라 바꿀 수 있음. CHANGELOG "남은 할 일"의 인자 이름 정리와 함께).
4. **출력 크기** (1.4.97에서 수정): Outlook 목록 행마다 긴 `entry_id`(짧은 `id`가 있음), 계산 안 된 수식 목록이 `uncomputed_formulas`·`unverified`·`agent_directive`에 세 번,
   상한 없는 `folders`(`--depth`), reader `list` 100개(40개에 14.6KB), 엑셀 읽기 기본 200행(8열에 약 16KB).
5. **시간 제한 보호** (1.4.97에서 수정: 닫힌 xlsx 읽기. compare·reader index는 원래 CONTINUE): 20만 행 x 10열 xlsx에서 `read`·`--where`·`--search`·`export` 36~38초, `read --formulas` 58초(시트를 두 번 읽음), `compare` 78초.
   `CONTINUE`는 `compare`와 reader index에만 있다. 느린 PC에서 `read --formulas`가 120초에 JSON 없이 끊길 수 있으므로 시한 전에 멈추고 `UNVERIFIED:`/`CONTINUE:`.

## 6. 문제없다고 확인한 것 (다시 검토하지 않아도 됨)

- `quote_arg`: PowerShell·pwsh·bash 2만 건 무작위 왕복 통과 (cmd는 줄바꿈·`\%`에서 실패, 문서에 최선 지원으로 적혀 있음)
- 여섯 도구 모든 하위 명령의 `-h`: JSON 1줄, 도움말은 stderr. 옵션 줄임 거절. 중단 130. 내부 오류는 `ERROR:`, traceback은 로그에만
- 오류 로그 2MB 넘김, `--body`/`--to`/`--query` 가림. 발송 가드(도구가 만든 초안만, 임시 보관함만, 1번만, 못 푼 수신자는 보내지 않음)는 우회 못 함
- reader: `CONTINUE` 이어서 등록, 동시 실행 색인, 색인·캐시 손상 복구, 바뀐 원본 자동 재등록, 부호·단위·범위 숫자 검색, 동의어, `clear --dry-run`, 암호·손상·빈 파일
- 엑셀 닫힌 파일: CSV 인코딩·따옴표 줄바꿈, 범위 해석, 공유 수식, `check` 단위 변환, BOM 비교
- 워드·PPT: wdStyle/ppLayout/내보내기 상수, `^` 처리와 255자 한도, fill 자리 해석과 그 `[추천]`·`confirm_commands`, dry-run, 세션 범위
