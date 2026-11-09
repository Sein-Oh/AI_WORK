# 작업 인계 (2026-11-07, 클라우드 세션 -> 집 PC / 회사 PC)

PC에서 저장소를 받은 뒤 새 Claude Code(또는 OpenCode) 대화를 열고 맨 아래 "붙여넣을 첫 메시지"(집 PC용 / 회사 PC용)를 보내면 이어서 진행할 수 있습니다.
이 파일은 작업이 이어지면 최신 상태로 고치거나 지웁니다.

## 1. 지금 상태
- 도구 버전: **ml-kit 2.9.78**, **office-cli 1.4.83** (main은 2.9.77·1.4.83. 2.9.78은 작업 브랜치 `claude/gemma-loop-1108`, PR을 머지했으면 main)
- Office는 Office 16만 대상 (2026-11-07 Office 2007 지원 폐기, 집 PC도 Office 2019로 바꿈)
- 2026-11-07 클라우드 작업: 사용법 단순화 (AGENTS.md "사용법 단순성"). `[추천]`은 늘 전체 명령, `[후보] 1) a`와 같은 번호의 `confirm_commands`,
  ml-kit 스킬 4개(ml-train, ml-predict, ml-report, ml-map), 예측 `-m <데이터 파일>`, 개선·`CONTINUE` 횟수를 도구가 셈, transfer `--from-app`/`--to-app`,
  ppt slide `--dest`, `--cards` 구분 `;`, 별칭·줄임말 제거. 클라우드 테스트: ml-kit 894 passed·1 skipped, office-cli+공통 3296 passed.
  집 PC 실기 확인 끝 (2026-11-07, 4장): 고친 명령의 실제 PowerShell 5.1 따옴표, 스킬 설치(setup.py), N1~N4·B1~B5 모두 통과.
  남은 실기 확인: Gemma의 스킬 선택·번호 `[후보]` 처리 (회사 PC, 3장)
- 저장소 규칙은 `AGENTS.md` 한 곳 (Claude Code는 `CLAUDE.md`가 불러옴). 특히:
  - 답변은 한국어, 이모지·장식 기호 금지 (CP949 안전 문자만)
  - 베타 단계: 이전 버전 호환 코드를 넣지 않음. CLI 인자·JSON 키는 바꿔도 되고, 바꾸면 SKILL.md·README·CHANGELOG·테스트를 같은 커밋에서
  - 리팩터링과 기능 변경은 별도 커밋/PR, 버그 수정에는 회귀 테스트
  - `archive/`는 수정하지 않음
- 배포 파일에 사용자 이름이 든 전체 경로를 넣지 않음 (성능예측표 등은 `<데이터>_ml_data\vN`처럼 끝부분만)
- 실기 로그는 사용자가 GitHub 웹으로 `docs/test_logs/`에 올림 (`.gitignore` 처리). 사용자가 요청하면 에이전트가 `git add -f`로 올림.
  경로 안의 사용자 폴더 이름(`Users\<이름>`)은 `run_all_tests.py`·`run_simplify_check.py`가 결과 파일을 쓸 때 `<사용자>`로 가림
  (2026-11-07: 손으로 가리다 잘린 출력 하나를 놓쳐 러너가 하도록 바꿈). 저장소 위치는 필요하면 `<저장소>`로 손으로 가림. 프로젝트 마무리 때 저장소 이력에서 정리 예정
- 클라우드 점검: 17~25차 점검(검토 -> 수정 -> 재검사 반복) 완료 (2026-11-03). 마지막 차수의 회귀·높음 결함은 모두 수정, 경미·중간은 각 도구 CHANGELOG "남은 할 일"
- 클라우드 테스트: 전체 pytest 1549 passed, 1 skipped (ml-kit 2.9.36, office-cli 1.4.32 기준). 회사 PC 수는 `run_all_tests.py`의 pytest 결과로 확인

## 2. 최근 바뀐 것 (클라우드 전체 점검 반영, 실기에서 볼 것)
| 버전 | 내용 |
|---|---|
| office-cli 1.4.10 | `ok()`가 `verified=False`를 거부하고 저장 파일은 형식별로 다시 열어 확인. `excel replace`가 `--range`의 시트에 씀. 결과가 ""인 수식 칸도 덮어쓰기 확인. `ppt write --title`로 다른 제목을 바꾸려면 `--overwrite`. 파일 이름만 준 쓰기 대상은 새로 만들지 않음(전체 경로만). reader 잘림·글자 없는 PDF 쪽을 알림. 오류 로그에서 본문·받는 사람 등 내용 값은 `<N자 생략>` |
| office-cli 1.4.11 | 문서 색인 형식이 바뀜(PARSER_VERSION 6): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행 |
| office-cli 1.4.12 | 문서 색인 형식이 다시 바뀜(PARSER_VERSION 7, PDF 스캔 쪽 판정): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행. 동시에 실행한 `index`도 색인 항목을 잃지 않음 |
| office-cli 1.4.13 | 문서 색인 형식이 다시 바뀜(PARSER_VERSION 8, 도장·Bates 번호만 남은 스캔 쪽도 판정): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행. 입력 오류로 멈춘 쓰기는 문서(새 시트·서식 포함)를 바꾸지 않고 되돌리기 지점도 밀어내지 않음. 색인을 잠깐 못 읽어도 비우지 않음. 모든 쪽이 스캔인 PDF의 `read`는 `UNVERIFIED`(OCR 필요). 첨부 저장 실패는 `UNVERIFIED` |
| office-cli 1.4.14 | 바꾼 것 없이 성공한 쓰기(0곳 바꿈, 차트 0개 삭제, 이미 있는 시트 add)는 되돌리기 지점을 남기지 않음(다음 `undo`가 실제로 바꾼 작업을 되돌림). 엑셀에서 쓸 수 없는 시트 이름(`/ \ ? * [ ] :`, 31자 초과)은 시트를 만들기 전에 STOP_RETRY와 `[추천]` 이름. 스캔 PDF를 `read --save`하거나 스캔 쪽만 고르면 `UNVERIFIED`. 깨진 색인 파일도 다시 만듦 |
| office-cli 1.4.15 | 엑셀이 예약한 시트 이름 `History`(대소문자 무관)도 시트를 만들기 전에 STOP_RETRY와 `[추천] History_`. `transfer --sheet`의 `[추천]`이 받는 문서의 있는 시트와 겹치지 않음. 바꾼 것 없이 성공한 명령(같은 글로 바꾸는 `replace`, 같은 자리로 `slide move`, 숨긴 것 없는 `unhide`, 고정 없는 `unfreeze`, 빈 범위 `clear` 등)은 "0곳/0칸 바뀜" 또는 "바뀐 것 없음"으로 답하고 되돌리기 지점을 남기지 않음. 휘발성 수식(NOW 등)이 든 닫힌 파일도 바꾼 것이 없으면 다시 저장하지 않음. `ppt replace --find`만 주면 STOP_RETRY |
| office-cli 1.4.16 | `ppt replace`로 대소문자만 다른 글을 통일하는 치환(`iphone` -> `iPhone`)을 실제로 바꿈. 새 파일로 `transfer --sheet Sheet1`처럼 기본 시트 이름을 주면 그 시트에 받음. `range hide`/`unhide`와 `clear`(전체·서식)는 늘 바꾼 것으로 봄(되돌리기 지점을 잃지 않게) |
| office-cli 1.4.17 | `excel`/`word`/`ppt replace`의 찾을 글이 빈 문자열이면 STOP_RETRY(빈 칸 채우기는 `write`). 기본 시트가 여럿인 새 파일로 `transfer --sheet Sheet1`도 그 시트에 받음. 대소문자만 다른 치환(`iphone` -> `iPhone`)에서 이미 바꿀 글인 곳은 바꾼 곳으로 세지 않음(`unchanged_same_text`) |
| office-cli 1.4.18 | 이모지 같은 BMP 밖 글자 뒤의 글도 `ppt replace`가 모두 바꿈. Word 찾기에서 `^`를 글자 그대로 찾음(`감사합니다^^`, `10^6`) |
| office-cli 1.4.19 | Word 치환에서 바꿀 글이 255자를 넘어도 바꾼 뒤 확인이 오류로 끝나지 않음(Word 찾기 대신 직접 셈). 찾을 글·앵커 길이는 `^`를 2자로 셈 |
| office-cli 1.4.20 | `transfer --to word,ppt`는 무엇이든 넣기 전에 받는 문서를 모두 확인: 한 앱만 열려 있으면 아무것도 바꾸지 않고 STOP_RETRY와 `[추천]`(열린 앱만 `--to word` 등, 그대로 1회 실행). 넣은 뒤 한 앱만 실패하면 `UNVERIFIED`(`succeeded`/`failed`, 다시 실행하지 않음). 열린 문서가 없을 때 안내는 `--to-target`/`--from-target`(없는 `--target`을 안내하지 않음). `excel chart`의 `next_commands`는 열린 앱만 `--to`에 넣음 |
| office-cli 1.4.21 | 시트를 주지 않은 차트 찾기(`transfer --chart`, `chart --action delete --name`)는 이름이 같은 차트를 제목보다 우선하고, 그래도 여러 시트에 있으면 하나를 고르지 않고 `[후보] 시트!이름`(그 형식을 그대로 받음). `excel chart`의 `next_commands`에 `--sheet`. 이름이 같은 열린 문서가 여럿일 때 안내는 그 명령의 옵션(`--to-target`/`--from-target`/`--name`) |
| office-cli 1.4.22 | 차트 `[후보]`의 `시트!이름`을 `--sheet`와 함께 줘도, 차트 이름에 `!`가 있어도 찾음. SKILL의 문서 `[후보]` 안내가 message가 말한 옵션(`--to-target` 등)을 따르게 함 |
| office-cli 1.4.23 | outlook 검색어의 `%`·`_`를 글자 그대로 찾음(`견적서_v2`가 0건이던 문제). 기본 파일 이름의 `/ : ?` 등을 `_`로(차트 `1/4분기 매출` 전달 실패). 차트 `'시트'!이름` 형식, 글자 그대로 `!`가 든 차트 이름과의 `[후보]`, 한 시트의 같은 이름 차트는 `[후보] 시트!#번호`. 손상된 HWP/HWPX는 STOP_RETRY. `undo --pick 0` 거절, outlook `--limit` 상한 200 |
| office-cli 1.4.24 | 차트 `[후보]`는 그 시트의 모든 차트로 이름이 겹치는지 보고 `#번호`로(후보마다 차트 하나). FileHeader 없는 OLE `.hwp`는 STOP_RETRY "HWP 5.0 문서가 아닙니다". `excel style --highlight` 값의 `"`·`~ ? *`를 글자 그대로. `excel write --append --formula`는 상대 참조를 실제 쓴 행 기준으로. 마크다운 펜스 코드 블록(백틱 3개)을 줄마다 글자 그대로, 표 정렬 줄 `:-`, 칸 글자 `\|`. `ppt write --cards`는 공간이 모자라면 새 슬라이드에, 슬라이드 밖 도형은 UNVERIFIED. reader `--units 1-999999999`가 멈추지 않음 |
| office-cli 1.4.25 | `.md` 코드 블록 줄의 `__init__`, `x**2`를 굵게로 읽지 않고 그대로 넣음(Word·PowerPoint). 줄 첫머리 인라인 코드 ```` ```x``` ````는 코드 블록 시작이 아님. 코드 블록으로 시작하는 파일도 다시 읽기 확인 |
| office-cli 1.4.26 | `ppt write`가 코드 줄·제목의 `__`·`**`(`def __init__`, `## __init__ 설명`)로 거짓 UNVERIFIED가 되지 않음. `word write`는 `---`로 시작하는 `.md`, 첫 표 행의 `\|`도 다시 읽기 확인. `ppt write --text`의 펜스 코드 줄은 글자 그대로, `###` 제목의 `**`는 굵게 표시로 뺌. `excel write --file .md`는 첫 번째 표만(코드 블록·두 번째 표 제외). 0건 결과(`office windows`, `undo --list`, `info --all`, `chart --action list`)는 count 0과 "0건". `reader clear --dry-run`이 20개를 넘으면 `full_result_file`. 오류 로그에서 `--section`, excel `--highlight` 값을 가림 |
| office-cli 1.4.27 | 제목 굵게 표시 규칙을 word·ppt 모든 수준(슬라이드 제목, `--title` 포함)에서 하나로(`**x**`, `__x__`는 표시를 뺀 글). 단어 안의 `__`(`report__2024__final.xlsx`)는 그대로. `2026. 11. 3.`처럼 4자리 이상 숫자는 번호 목록이 아니고, 1이 아닌 번호로 시작하는 목록은 사용자 번호를 글로 둔 단락. 오류 로그 message에서 가린 인자 값과 문서에서 가져온 후보·문구를 뺌. `undo --list`(되돌릴 수 없는 작업만), `ppt slide --action list`(0장), `outlook list/search`(0통)는 `count`와 `VERIFIED: 0건`. `reader clear`의 `applied`·완료 CSV는 실제로 지운 것만(`removed` 열) |
| office-cli 1.4.28 | 오류 로그에 문서 글이 남지 않게 구조를 바꿈: 문서 글(칸 값, 제목, 노트, 시트·문서·파일·메일 폴더 이름 목록, 후보 단락)을 message에 넣는 오류는 그 글을 뺀 `log_message`를 주고, 표식 자르기(`사용 가능한 <공백 포함 종류>:`, `후보:`, `추천:` 추가)는 대비로 남김. 가린 값의 10자 이상 앞부분과 traceback도 가림, 1~2자 값은 따옴표 안만. `undo`는 되돌릴 수 없는 작업만 있으면 `VERIFIED: 0건.`, `undo --list`는 그런 작업이 20개 넘게 앞에 있어도 되돌릴 수 있는 지점을 보여 줌 |
| office-cli 1.4.29 | `transfer --chart`는 받는 문서를 찾지 못하면(오타 `[추천]`, `[후보]`, 확장자) 다음 방식(redraw, relay)으로 넘어가지 않고 원래 오류 그대로 멈춤(`[추천]`이 message 끝, 클립보드를 덮어쓰지 않음). outlook 폴더 `[후보]`는 끝 이름끼리 비교(같은 메일 계정 때문에 관계없는 폴더가 후보가 되지 않음), 오류 로그에 메일 계정 이름을 남기지 않음. 차트 다시 그리기는 `$`를 수식으로 읽지 않음. `reader search/index --file <파일 하나>`의 내부 오류는 `ERROR:`. `run_win_tests.py --agent -h`는 pywin32 없이도 JSON exit 0 |
| office-cli 1.4.30 | outlook `--folder` 오류의 `(경로: ...)`도 오류 로그에서는 계정 이름 자리를 `<메일 계정>`으로(`not_found`의 `log_extra`). `transfer --method relay --to ppt`는 받는 PPT의 `--slide`를 확인한 뒤에 클립보드·임시 워드를 씀(없는 슬라이드로 클립보드를 덮어쓰지 않음). relay `--to-target ""`는 활성 문서(`KeyError` 아님) |
| office-cli 1.4.31 | relay `--to ppt`는 열린·활성 PPT의 `--slide`를 찾기만으로 확인하고 쓰기 세션(되돌리기 지점)은 임시 워드 단계 뒤에 엶(워드 단계가 실패해도 바꾼 것 없는 지점이 남지 않음). 닫힌 xlsx의 계산된 값이 없는 수식 칸(openpyxl이 쓰고 엑셀에서 연 적 없는 파일)은 excel read/export, transfer 범위, reader read/search에서 UNVERIFIED와 칸 목록(`uncomputed_formulas`). docx 글상자 글을 따로 단위로 읽음(두 벌 저장은 한 번). 닫힌 docx 표 번호는 콘텐츠 컨트롤 안의 표도 셈(read·COM과 같음). word 굵게 위치는 UTF-16 단위. excel read --search에 `count`와 `VERIFIED: 0건.`. 마크다운 `---`는 word에서 구분선(아래 테두리 빈 단락), 맨 앞 YAML 머리말은 넣지 않고 `warnings` |
| office-cli 1.4.32 | 닫힌 xlsx를 범위 없이 읽을 때(`excel read`·`--search`·`--save`, `export`, transfer, `B:B`, 큰 범위 줄이기) 계산된 값이 없는 수식 칸의 행·열(합계 행 등)을 자르지 않고 남기며 UNVERIFIED. workbook.xml `fullCalcOnLoad`(xlsxwriter·openpyxl)면 모든 수식 칸을 계산되지 않은 칸으로 봄(`<v>0</v>`을 값으로 믿지 않음, 엑셀은 저장할 때 지움). 수식 칸 확인이 실패하면(MemoryError 포함) `uncomputed_check: "unknown"`과 UNVERIFIED "수식 계산 여부를 확인하지 못함". reader 파서 버전 10 |
| office-cli 1.4.33 | En-Genius 스킬로 실행하면(`ENGENIUS_OUT`이 있으면) `config.json`, data, cache를 `<사용자>\Documents\En-Genius\CLI-data\office-cli`에 둠 (트레이가 스킬 폴더를 실행마다 지우므로). `OFFICE_CLI_CONFIG`로 설정 파일 위치 지정 가능. OpenCode·`C:\CLI` 실행은 그대로 |
| office-cli 1.4.34 | 새 명령 `outlook.py attachments`: 첨부 목록(`index`, `name`, `size`, `hidden_inline`), `--save [전체 경로 폴더]`로 내려받기(기본 `%USERPROFILE%\Downloads\메일 첨부\<시각>_<제목>`, 자동 정리 안 함), `--name 일부`/`--index 1,3`으로 고르기(맞는 것이 없으면 `[후보]`), 같은 이름은 ` (2)`와 `recovered`. `read/current --save-attach` 없앰(대신 `attachment_names`와 `next_commands` "첨부 내려받기"). `new/reply/forward --attach`는 없는 전체 경로 파일이면 메일을 만들기 전에 같은 폴더의 비슷한 파일 `[후보]`, 붙인 뒤 `Attachments.Count`로 확인(못 하면 UNVERIFIED, `--send`면 보내지 않음) |
| office-cli 1.4.35 | En-Genius 스킬 zip은 지시문과 `run.py`만 담고 `C:\CLI\office-cli`를 실행: OpenCode와 같은 data, cache, config.json을 씀 (1.4.33의 고정 폴더는 도구가 트레이 임시 폴더에서 돌 때만) |
| office-cli 1.4.36 | `excel.py read --formulas`: 읽은 범위의 수식 칸 목록(`formulas`: `cell`, `formula`, `value`, 배열 수식은 `array`·`array_range`), `formula_count`, 200개 상한(`formulas_truncated`, 전체는 `full_result_file` CSV). 열린 문서는 Formula2(없으면 Formula)를 범위 한 번에, 닫힌 xlsx는 openpyxl. SKILL.md가 바뀌어 En-Genius 스킬 zip을 다시 가져와야 함 |
| office-cli 1.4.40~1.4.48 | 관리자 권한 실행이면 "열린 문서 없음" 대신 권한 안내(`elevated_process`), `range delete` 전 `#REF!`가 될 수식 칸 알림, reader 캐시의 확인 실패 다시 확인, 배열 수식 범위 전체를 수식 칸으로, 멀리 떨어진 계산 안 된 수식 칸 처리, `read --where`의 또는(`\|`)과 따옴표 값, 정렬 기준 4개 이상, 라이브러리 없음은 `missing_library`, next_commands 다듬기 |
| office-cli 1.4.49~1.4.51 | 규격서 대조 작업: `reader.py read --section`(캐시 문서 장절 읽기), `excel.py comment`(칸 메모 목록·달기·지우기, CSV로 여러 개), `reader.py search` 재현율(조사·단위 표기·전각 허용, 0건이면 부분/오타 일치로 넓힘 `match_mode`, 점수 순위, `data\synonyms.txt` 동의어) |
| office-cli 1.4.52~1.4.61 | `transfer --method`·`search -c` 삭제, 여러 칸 한 번에 `excel.py write --cells`, 시트 보호 `sheet_protected`와 거부된 수식 `[추천]`(되돌림), reader 등록·검색 `CONTINUE:`(`OFFICE_CLI_TIME_LIMIT`), 규격 대조 `excel.py check`, 열기 감시 30초 `drm_or_dialog`, `reader.py diff`(판 비교), 메일 짧은 `id`(긴 값은 `entry_id`)·`thread`·`calendar --days`, `word.py fill --file` |
| office-cli 1.4.62 | 통합 스킬 `office-cli`(16KB)를 작업별 스킬 4개로 나눔: `office-excel`, `office-doc`, `office-mail`, `office-spec` (`CLI/office-cli/skills/<이름>/SKILL.md`, 각 6KB 이하, 명령은 전체 명령). `apply_skills.py`(지금은 `CLI/setup.py`)가 4개를 설치하고 옛 `skills\office-cli` 폴더는 보관 후 지움. 도구 동작은 그대로 (`-h` 결과의 `skill_file`만 해당 작업 스킬) |
| office-cli 1.4.63 | 독립 검토 버그 수정: 2열 표 파일(칸값·메모·값 파일)의 셋째 열은 쓰기 전에 STOP_RETRY(큰따옴표 안내), `--cells` 따옴표·수식 나누기, `comment --append` 재실행 중복, 엑셀 쓰기 실패 후 되돌리기가 글자 상수(`007`, `1-2`)를 바꾸던 것, `check`가 같은 문장의 다른 항목 조건을 섞던 것과 칸 단위 없음은 `판단필요`, `word.py fill` 표 칸에 행·열 머리글 표시·문서 보호 확인, 열기 감시가 사용자 문서가 있는 Word·PowerPoint를 끝내지 않음 |
| office-cli 1.4.73 | 10차 검토 수정: 사용자 확인이 필요한 명령(메일 send, 미리보기 뒤 실제 변경)은 `next_commands`가 아니라 `confirm_commands`로(`needs_user_confirm` 없앰, 전역 지침 3장 6번 예외), send는 발송 원장(`outlook_drafts.json`, 잠금 아래 '보냄' 먼저 기록)에 있는 이 도구가 만든 초안이고 임시 보관함에 있으며 `Sent`·`Submitted`가 아닐 때만(동시·두 번 발송, 지운 초안, 사용자의 다른 초안 거부, id 오류는 `next_commands` 없는 거부), 없앤 `--send`는 exit 2 + `[추천]`. reader 층 A 경계를 허용 목록으로(숫자 앞·단위 뒤), 앞이 숫자뿐인 `/`의 분수·비율 제외, 스니펫 층 B 자리는 `[[ ]]` |
| office-cli 1.4.74 | 관리자 권한 안내는 UAC로 권한을 올린 창일 때만 (UAC가 꺼진 회사 PC는 모든 프로그램이 같은 권한이라 안내하지 않음) |
| office-cli 1.4.77 | 회사 PC 실기(2026-11-06) 대응: 닫힌 파일에 쓴 뒤 닫힌 시트를 읽어 실패로 보고하던 문제(`write --cells`, `chart`, `comment list`), `--cells` 쉼표 구분은 `;`로 바꾼 명령을 `[추천]`, `check --unit-col`(값 열 오른쪽 `단위` 열 자동), `word fill --slots "자리=값; 자리=값"`(값 파일 없이) |
| office-cli 1.4.78 | `word fill`: 괄호 안 띄어쓰기만 다른 자리(`{{ 업체명 }}`)는 문서의 자리 하나와 맞으면 채우고 `recovered`, `--slots` 값 안의 줄바꿈 유지. 스킬 크기 여유(각 64~118바이트) |
| office-cli 1.4.79 | 사용법 단순화: `[추천]`은 늘 전체 명령, transfer `--from-app`/`--to-app`, ppt slide `--dest`, `--cards` 구분 `;`, `CONTINUE` 횟수를 도구가 셈 |
| office-cli 1.4.80 | `excel compare`(표 두 판 비교, BOM 개정 비교: 머리글·열 역할·짝지을 키를 도구가 정함, 추가·삭제·수량·위치·속성 변경, 품번 변경 추정, 결과 xlsx와 새 판 칸 메모 미리보기)와 새 스킬 `office-data` |
| office-cli 1.4.81 | 남은 `[후보]`에 번호와 같은 번호의 `confirm_commands`, `reader --query` 구분자 `;`(그리고)·`\|`(또는)(옛 AND/OR는 `[추천]`), 빈 결과 `count 0`, 파일 지우기 재시도, `--dry-run` 확대(excel replace·range clear/move·sheet delete·chart delete, word replace·write, ppt replace·slide) |
| office-cli 1.4.82 | Office 2007 지원 폐기 (Office 16만): PDF 추가 기능 확인·PowerPoint 2007 E_FAIL 판정·Word `SaveAs` 대체·그림 글자 `/` 처리 삭제, 실기 러너 D1·D4의 PDF ENV-SKIP 삭제 |
| office-cli 1.4.76 | 도구가 내는 명령(`next_commands`, `confirm_commands`, `[추천]` 전체 명령)의 따옴표를 실행 셸에 맞춤: OpenCode처럼 `pwsh`가 있으면 PowerShell 7, 없으면 Windows PowerShell 5.1(`OFFICE_CLI_SHELL`로 강제). PowerShell은 작은따옴표(`$`·백틱·`{}`·`@`·`;` 글자 그대로), 5.1은 값 안의 `"`·끝 백슬래시·빈 값을 미리 바꿈. 테스트에 셸 시뮬레이터(`CLI-dev/tests/shell_sim.py`) |
| office-cli 1.4.75 | 11차 검토 수정: `confirm_commands`의 덮어쓰기·삭제·발송 항목에 `"destructive": true`("묻지 말고"라고 해도 늘 물음, 전역 지침 3장 6번), 문서를 바꾸는 이어 하기(`read --save` 뒤 write, write 뒤 chart, chart 뒤 transfer 등)는 `next_commands`에서 `confirm_commands`로, 파괴적 명령 오타는 `[후보]`, 발송 실패 초안은 다른 거부 문구, reader 층 A 허용 목록 완화(짝 따옴표·괄호, 전각 콜론, `±` 뒤 숫자, 단위 뒤 `(`) |
| office-cli 1.4.72 | 9차 검토 수정: outlook 발송은 저장한 초안을 그 항목 그대로 보내는 `send --id <draft_id>` 하나로(`new`·`reply`·`forward`의 `--send` 없앰, 초안 결과의 `next_commands[0]`에 `needs_user_confirm: true`, 이미 보냈거나 지운 초안은 거부해 중복 발송 없음), 답장·전달은 보내지 않은 초안을 원본으로 쓰지 않음(열린 초안 창에 답장하던 위험). reader 층 A: `/` 값 구분은 양쪽 숫자+단위이고 `/` 뒤가 1·10^n이 아닐 때만(숫자만 앞이면 Hz·V 목록), AND 스니펫은 검색어마다 창(` / `로 이음), 단위 없는 검색어의 DC/AC 예외 없앰, 붙은 기호(천 단위 따옴표, U+2044·U+FF0F·U+2215, 곱하기, U+3161, 폭 없는 공백, 지수 `m-1`) 차단 |
| office-cli 1.4.71 | 8차 검토 수정: 층 A(`VERIFIED`)의 뜻 확정 - 검색어 글자열이 그 자리에 글자 그대로, 숫자·단위 낱말 경계가 맞게 있음(띄어 쓴 부호 `- 40 ℃`·범위·이하/이상·한글 꼬리의 뜻은 판정하지 않고 스니펫으로 보임). 스니펫·근거 문장·check 조건은 층 A 자리(일치 앞 최소 15자), 숫자 검색 directive에 앞뒤 함께 전달 안내. 낱말 수준 오일치(띄어 쓴 천 단위, 곱 단위·띄어 쓴 단위 이어짐, 숫자 코드 뒤 `-숫자`)는 층 A 금지, 형식 차이(`/` 값 구분, `(DC)`·`(Max)` 괄호, 붙은 `DC`/`AC`, `IP 67`/`M 10`)는 층 A. check `--header-row 0`(머리글 없음)을 남은 행 명령에 |
| office-cli 1.4.70 | 7차 검토 수정: reader 숫자가 든 검색어는 두 층 판정. 글자 그대로 일치(층 A: 띄어쓰기·탭 여러 칸 = 한 칸, 숫자와 단위·± 앞뒤 띄어쓰기 0~1칸 차이만 허용, 엄격한 앞뒤 경계)만 `exact` VERIFIED, 정규화 매처만 찾은 곳(층 B)은 결과에 넣되 `literal: false` + `match_mode: exact_mixed` + `UNVERIFIED:`(`literal_count` 새 키). check는 층 B로 찾은 항목의 reason에 표시. 층 B에서 모두 대문자 단위(`10 MM`, `10UF`, `20MA`)는 대소문자 무관. 숫자 안전망 시간 상한은 단위마다 확인, check 시간 제한으로 대조하지 못한 행은 `대조하지 못함`(match_mode null, 메모 제외) + 남은 범위 check 명령(`--header-row` 새 인자), 규격 수치의 쉼표 뒤가 3자리가 아니면 수치 아님 |
| office-cli 1.4.69 | 6차 검토 수정: reader 숫자 정확 일치는 보수적으로(숫자 앞 부호가 분명히 같을 때만 exact, 글의 숫자 앞(띄어쓰기·`/` 건너)이 `-`·`±`·`+/-`·`+ / -`·`+`인데 검색어와 다르거나 범위인지 부호인지 애매하면 숫자 안전망 후보, 범위 대시는 양옆을 띄운 `0 - 20`만), 숫자 안전망 선형 탐색(지수 백트래킹 제거, 20초 상한), 숫자 아닌 낱말도 같은 단위에 있어야 후보, 단위 접두 대소문자(`mHz`/`MHz`, `mohm`=mΩ, μ=u), `Vp-p`·`Vrms`·`barg`·`dBA`·지수·`φ`는 exact 아님, `spec_conditions` 긴 숫자열 제곱 시간 제거, `check` 시간 제한 전 남은 행은 판단필요 |
| office-cli 1.4.68 | 5차 검토 수정: reader 숫자 검색 안전망(정확 일치 0건 + 숫자가 든 검색어면 숫자 값만 같은 곳을 `candidates`(`match_mode: number`, `number_words`)와 `UNVERIFIED: 정확히 일치한 곳은 0건, 숫자가 같은 곳 N건 - 표기를 확인하세요`로, 0건 VERIFIED는 이것도 0건일 때만), `±` 앞 글자 제한 제거(`100±0.5 mm`, `5V±5%`), 한글·`=`·`:` 뒤 빼기는 부호, 숫자 뒤 띄어쓴 대시는 범위, `+`·부호 없는 숫자는 `±`와 다름, 같은 단위 꼬리(`Vrms`, `sec`, `hrs`, `dBA`, `barg`), `kΩ`=`kohm`·`mm`=`㎜`·`M6x1.0`=`M6×1.0`·`Φ`=`Ø`·`08:30`=`8:30`·코드 경계, `spec_norm` 지수 2자리까지, `check` 안내에 메모 미리보기 행동(등록이 남으면 같은 명령 다시 대조만) |
| office-cli 1.4.67 | 4차 검토 수정: `excel.py check` 결과는 늘 `UNVERIFIED:`(자동 판정은 참고용, 행마다 읽은 조건·근거를 보여 주고 사용자 확인 뒤 메모. 메모에 `자동 판정(참고)`·읽은 조건·근거), 해석 버그(유니코드 빼기, `±N 이내`, ohm 대소문자, 지수 표기, `dB(A)` 꼬리, 최대·최소·약·`~ 시`·단서·부정문, 단위 없는 수치·치수·소수점 쉼표, (c)는 한계 낱말이 있을 때만, `A에서 B로`), 말뭉치 261건. reader 숫자 검색: 부호(`-20 ℃` != `20 ℃`)·단위 끝 경계(`10 m` != `10 mm`)·기호 소수 경계, `10%` = `10.0 %`, 날짜 표기, 천 단위 쉼표 3자리. outlook 번호 잠금은 운영체제 잠금(`msvcrt.locking`/`fcntl.flock`)으로 바꿔 번호 중복 경합 제거 |
| office-cli 1.4.66 | 3차 검토 수정: `excel.py check` 보수적 판정(합치는 규칙은 아래·위 한계 -> 범위, 기준 값 + `±N 단위`/`±N %`, 기준 값 + 품는 양쪽 범위 셋뿐, 그 밖에 조건이 둘 이상이면 `판단필요`, `사용 범위`·`동작 범위`는 별개 항목), 표준 판정 말뭉치 `tests/check_corpus.tsv`(132건), 숫자 검색어 `2.0`이 `2.5`에 맞던 회귀, 조사 `로`는 ㄹ 받침 뒤만, 전각 단위 글자(`㎜` 등), 메일 번호 잠금(쥔 동안 10초 상한, pid 확인), 숨은 Word 종료(문서 수 재시도, 자기 문서뿐이면 끔) |
| office-cli 1.4.65 | 독립 검토 2차 재검토 수정: 메일 짧은 `id`는 번호 파일을 못 읽거나 깨져도 1부터 다시 매기지 않음(그 결과는 긴 EntryID, 잠금 파일·`outlook_ids.next`), `excel.py check`의 허용·변동 범위 절을 앞 항목 한계로, 숨은 Word는 남은 문서가 없을 때만 끝냄, 사용자가 띄운 PowerPoint의 DisplayAlerts는 바꾸지 않음, 셋째 열 오류에 `[추천]` 대신 "고친 줄", 숫자 검색어 `2.1` 장절 경계, 넓혀 찾은 검색 안내 유지·조사 `로`, office-mail `--send` 확인 순서 |
| office-cli 1.4.64 | 독립 검토 2차 수정: 메일 짧은 `id`는 이어서 늘어나는 번호(옛 번호는 같은 메일이거나 "없는 번호"), reply/forward 결과에 `original`, reader 읽지 못한 파일 기록으로 `CONTINUE:`가 끝나게, `--section` 번호 경계, 넓혀 찾은 결과는 `count: 0`+`candidates`+`UNVERIFIED:`, 한글 fuzzy 엄격히, 숫자 검색어 정규화, calendar/thread 미확인 표시, 스킬 안전 규칙 보충, apply_skills 링크 처리, 채점기 명령 나눠 판정 |
| office-cli 1.4.39 | 실제 사용 오류 로그 반영: `office.py write`, `office.py transfer --text`, `reader.py read --section`에 맞는 명령(워드/PPT/엑셀 write, reader search --section) 안내, `excel.py write --cell`에 `[추천] --range` |
| office-cli 1.4.38 | `opencode run` 기록에서: 엑셀 `write`의 `--text`에 `[추천] --value`(전에는 `--target`), 표 파일 못 찾음 후보에 표 파일만, SKILL.md에 칸 하나 쓰기(`--value`) 줄 |
| office-cli 1.4.37 | SKILL.md에 빠졌던 기존 기능(메일 list/folders·search --folder/--where·reply --all·new --cc·--body-file, PPT slide add/duplicate/delete/move, 엑셀 sheet/range/export/style/replace --all-sheets·차트 종류, 워드 --toc/--callout·replace --highlight)을 넣고 SKILL.md의 명령·옵션이 argparse에 있는지 테스트로 대조. 새 기능: `outlook.py list|search --days N|--since YYYY-MM-DD`(period), `excel.py range --action insert|delete`(값이 있으면 needs_user_decision, `--dry-run`, `--overwrite`), `range --action sort --by "부서,금액" [--desc]`, `read --where "금액 >= 100; 부서 = 영업"`(`row_numbers`) |
| ml-kit 2.9.11 | 날짜 열을 학습·예측에서 같은 방식으로 해석(YYYYMMDD 포함). `-i` 값은 글자로 받아 범주 코드 `007` 유지. 머리글·단위 행 옵션이 예측 `next_commands`에도 들어감. `-i` 6건 이상은 결과 CSV 저장 |
| ml-kit 2.9.12 | 성능예측표: 작은 값 출력의 자리수, "실측 없는 영역" 기준을 표 해상도에 맞춤, **입력은 타겟마다 최대 8개**, 타겟별 `--inputs "T1=a,b;T2=c"`. `ml_export`·`ml_map`에 `--overwrite`. `ml_map`은 열마다 그 열에 값이 있는 행으로 계산. 예측값은 작은 값이면 유효숫자 4자리 |
| ml-kit 2.9.13 | 타겟 값이 모두 같으면 `STOP_RETRY`(학습 안 함). 완전히 같은 행은 교차검증에서 같은 fold로 묶고 `warnings`에 중복 개수. **`[추천]` 뒤가 `python`으로 시작하는 전체 명령이면 그대로 1회 실행**(레이아웃 옵션을 뺀 예측, 성능예측표 `--inputs` 전체. 전역 지침 `CLI/opencode/AGENTS.md`도 바뀜). `ml_map` 여러 구할 열은 `basis_by_column`. 작은 입력값(1e-6 단위)의 학습 범위·`--find` 결과 |
| ml-kit 2.9.14 | 숫자 입력에 쉼표(`-i "압력=500,000"`)가 있으면 `STOP_RETRY` + 쉼표를 뺀 전체 명령 `[추천]`. 예측 JSON의 `drift_warnings` 키 삭제(`extrapolation_warnings`는 20개까지, 전체 수 `extrapolation_warning_count`). 한 버전이 모든 타겟의 최고면 `ml_report -f`가 다시 학습하지 않고 그 버전을 씀(`made_best: false`). 머리글 행 안내는 열 이름 `[추천]`/`[후보]` 또는 `--header-row`. 입력+타겟이 다른 행이 10개 미만이면 `STOP_RETRY` |
| ml-kit 2.9.15 | 클래스가 많고 행이 적은 분류(10~15행)도 내부 오류 없이 학습. 숫자가 대부분인 단위 행은 `STOP_RETRY`, `[추천]` 머리글 행에 단위 행도 함께(전체 명령). 단위 행이 있는 새 파일을 옵션 없이 예측하면 `--unit-row`를 넣은 `[추천]` 명령. `--history`의 `best_version`(새 키 `best_version_dir`)과 보고서 `-d`는 모든 타겟의 최고가 있는 버전. 엑셀이 없는 PC의 보고서는 성능예측표를 확인하지 못해 `UNVERIFIED:` |
| ml-kit 2.9.16 | 범주·글자 타겟이 든 첫 데이터 행을 단위 행으로 보지 않음(`[추천]`에 틀린 `--unit-row`가 없고, 그런 `--unit-row`는 `STOP_RETRY`). 단위 없이 학습한 모델로 단위 행이 있는 새 파일을 예측하면 `--data-row`를 넣은 `[추천]` 명령. `ml_map --points`도 머리글 행 `[추천]` |
| ml-kit 2.9.17 | 단위 행 판단을 그 아래 값 최대 10개의 과반으로 함: 첫 값이 `측정불가`·`NG`·`12%`이거나 날짜 열(`yyyy-mm-dd`)이 있는 단위 행도 `--unit-row`로 받고, 옵션 없이 예측해도 단위 행을 건너뛰는 `[추천]`. 빈 행이나 머리글 위의 `--unit-row`는 `STOP_RETRY`(이전에는 진짜 단위 행을 데이터로 예측). 제목 행을 머리글로, 진짜 머리글을 단위 행으로 보던 `[추천]`(두 번 다시 실행) 수정 |
| ml-kit 2.9.18 | 숫자가 든 칸이 있는 행은 단위 행이 아님(첫 행이 `측정불가, 측정불가, A, 61.66`이면 데이터). 첫 데이터 행이 비었거나 `측정불가`여도 그 위의 단위 행을 건너뛰는 `[추천]`. 제목 줄이 좁은 CSV를 읽고 머리글 행 `[추천]`. 단위가 든 모델로 제목+머리글 파일을 예측하면 `[추천] --header-row`. 페이지마다 반복된 머리글(단위) 행은 빼고 `warnings`, 합계·평균 행은 빼지 않고 `warnings`('제외' 열 안내) |
| ml-kit 2.9.19 | 단위 행을 칸마다 아래 값과 비교해 판단: 타겟 열이 없는 예측 파일의 첫 행 `측정불가, 측정불가, A`는 데이터(모든 행 예측), 배율 숫자(`°C, 1000, , %`)와 이름 없는 머리글 아래 `비고`가 든 단위 행은 단위 행. `제외` 열에 표시한 합계 행과 집계 범주(`평균`/`최대`)는 합계 행 경고 없음. `;` CSV의 소수점 쉼표(`175,6`)는 소수점으로 읽고 `warnings`(천 단위와 섞이면 `STOP_RETRY`) |
| ml-kit 2.9.20 | `;` CSV의 `27,8448`(소수 4자리)을 278448로 읽던 문제 수정(정규식 이름이 겹쳐 있었음). `0,149`는 항상 소수, `1.346,8`(점 천 단위 + 소수점 쉼표)은 1346.8로 읽고 `warnings`(다른 형식과 섞이면 `STOP_RETRY`). 합계·소계 행은 열에서 차지하는 비율과 관계없이 경고(교대별 소계 25%, 빈 열의 합계 1개), 집계 범주로 보는 것은 `평균`류만. 단위 `ng`(나노그램)는 빈 값 표시가 아님(`NG`만). 큰 파일 읽기가 빨라짐(10만 행 x 32열 CSV 약 4.8초 -> 1초) |
| ml-kit 2.9.21 | 평균·합계 행을 값으로 판단: 바로 위 데이터 행들의 평균(합계)과 같은 `평균`(`합계`) 행은 비율과 관계없이 경고(교대별 `A조, B조, C조, 평균`). 통계 이름 열(`집계` = 합계/평균/최대/최소, `Sum/Avg`, 모든 행이 `평균`)은 값이 맞지 않으면 경고 없음. 소수점 쉼표 `;` CSV의 `1.300`(점으로 묶인 정수)을 1.3으로 읽던 문제: 같은 열과 파일이 일관되면 1300으로 읽고 `warnings`, 아니면 `STOP_RETRY` |
| ml-kit 2.9.22 | 합계 글자 행은 값과 관계없이 다시 항상 경고(하단 `합계`/`평균`/`최대`/`최소` 블록, `소계` 아래 총 `합계`, `평균`이 범주인 열의 `소계`), 통계 이름 열(합계·집계 글자가 값 칸의 절반 이상)만 경고 없음. 평균 행 값 비교가 천 단위 쉼표(`2,452`)를 숫자로 읽음. 통계 이름 열 10만 행 읽기 111초 -> 0.3초. 소수점 쉼표 CSV의 점 묶음 숫자(`1.300`, `2.125`)는 항상 `STOP_RETRY`(2.9.21의 1300 읽기 없앰), 모델에 소수점 쉼표 판단을 남겨 예측 파일·`ml_map --points` 조회 파일도 같게 읽음 |
| ml-kit 2.9.23 | 세로로 병합된 일자·라인 열 아래의 `합계`/`소계` 행도 다시 경고(통계 이름 열은 최대·최소·평균 같은 말이 함께 있을 때만). 엑셀 글자 칸의 `175,6`처럼 천 단위가 아닌 쉼표 숫자는 1756으로 바꾸지 않고 읽을 수 없는 값으로 경고 |
| ml-kit 2.9.24 | 일자 병합 일일 보고서의 일별 `합계`는 맨 아래 `최대`/`최소` 줄이 있어도 경고. `0,149`처럼 0으로 시작하는 쉼표 숫자는 천 단위가 아님(읽을 수 없는 값). 학습에서 숫자 열의 글자 칸을 중앙값으로 채우면 `warnings`에 알림 |
| ml-kit 2.9.25 | 읽지 못한 숫자 칸 경고는 학습하는 행(타겟 있음, `제외` 아님, 머리글 아래 단위 행 아님)과 모델이 숫자로 읽는 열(`--eval` 원본 열 포함)만: SKILL 예제의 단위 행 거짓 경고 5개 없앰, 열마다 1번, 5열을 넘으면 개수로. 합계 글자 넓힘(`일계`, `월계`, `누계`, `총 합계`, `전체합계`, `합계:`, `Total:`, `소계(A)`, `주간계`, `평균값`, `AVG.`; `설계`·`단계`·`온도계`는 아님). 선택한 피처 칸의 절반 이상이 소수점 쉼표 글자(`0,149`)면 `STOP_RETRY`, 그보다 적으면 경고 |
| ml-kit 2.9.26 | 2.9.25 회귀 수정: 범주 열의 `A계`/`B계`, `생산계`/`품질계`, `레벨계`/`pH계`를 합계 행으로 모든 행 경고하던 문제. 넓힌 합계 글자(`...합계`, 짧은 `...계`, 낱말 `Total`/`Sum`, `평균값`)는 열 안에서 합계 행처럼 보일 때만(채운 칸 절반 이하, 행 수 절반 미만, 글자 2개 이하) 경고, 정확한 낱말(`합계`, `소계`, `일계` 등)은 그대로. `SUM-01`, `Total_Loss`, `평균온도`는 합계·평균 글자 아님 |
| ml-kit 2.9.27 | 2.9.26 회귀 수정: 일자 병합 일일 보고서의 `일합계` + `월합계`/`총 합계`, 빈 열의 `총 합계` 1개를 경고하지 않던 문제. 넓힌 합계 글자는 값으로 판단: 행의 숫자 칸 절반 이상이 바로 위 데이터 행의 합, 위 합계 행의 합(`월합계` = `일합계`들), 위 모든 데이터 행의 합이면 경고, 합계 행이 적은 열에서는 숫자가 데이터보다 훨씬 크면 경고. 범주 값(`생산계`/`품질팀`, 센서 사이 `pH계`, `A계`/`B계`)은 경고 없음. `-i 'k=v'`의 같은 이름 두 번, `온도`와 `온도(°C)`를 함께 주면 `STOP_RETRY` |
| ml-kit 2.9.28 | 2.9.27 회귀 수정: 공정 열(`온도`/`압력`/`속도`)은 평균, `생산량`만 합인 `총 합계`/`월합계`/`일합계` 행을 경고하지 않던 문제 (블록 평균과 같은 칸은 판단에서 빼고, 합 1칸 이상 + 합·평균 칸이 절반 이상이면 경고). `-i`의 단위가 다른 이름(`압력(MPa)`, 모델은 `압력(bar)`)은 이름 바꾸기 `[추천]` 없이 모델 단위로 값을 환산하라는 `STOP_RETRY`. `온도`+`온도(°C)` 동시 입력은 모델 입력(ml_map은 데이터 열)과 맞출 때 판단: 모델에 둘 다 있으면 사용, `온도(°C)`만 있으면 `STOP_RETRY`. JSON `-i`의 같은 키 두 번, 같은 피처로 가는 별칭 쌍도 `STOP_RETRY` |
| ml-kit 2.9.29 | 단위가 다른 이름은 기본 이름이 달라도(`압력(MPa)`, 모델은 `반응압력(bar)`) `[추천]` 없이 값 환산 `STOP_RETRY`, `[후보]`의 단위 다른 후보에 "(단위 다름: 값 환산 필요)" (ml_map `-i`·`--find`·`--points`, `--find -i` 포함). `--find`의 목표 단위가 타겟과 다르면(`생산량(kg)`, 타겟 `생산량(t)`) `STOP_RETRY`. `--find`의 `-i`·목표에 여러 행 JSON 배열은 `STOP_RETRY`. 2.9.28 회귀 수정: 일정한 설정값 열이 평균 칸으로 세어져 범주 행(`A계`)을 합계 행으로 경고하던 문제 (블록 안에서 값이 바뀌는 열만 평균 칸) |
| ml-kit 2.9.30 | 타겟 이름(`--find` 목표, `-y`, `ml_report.py --target`)은 열 이름처럼 맞춤: `수율`이 `수율(1차)`·`수율(최종)`에 모두 맞으면 `[후보]` (이전에는 `--find`가 하나를 말없이 고름). 예측 `-i`의 모델이 쓰지 않는 이름(`압력(MPa)`, `습도`)은 `STOP_RETRY` (이전에는 버림). `ml_map.py --points`의 데이터 열과 비슷한 이름 열(`반응 압력(bar)`, `압력(MPa)`)은 `STOP_RETRY`. 합계 행: 블록 안에서 일정한 열의 값과 같은 칸은 판단에서 빼고 합 칸은 0이 아닌 수 2개 이상 (설정값을 적은 교대 보고서 `일합계` 경고) |
| ml-kit 2.9.31 | `[추천]`은 오타일 때만(`sure_typo`: 같은 이름, 한 이름이 다른 이름 안, 자모 하나 차이. 숫자가 다르거나 짧은 영문 이름의 길이가 다르면 아님), 아니면 `[후보]` (`습도`/`속도`, `T1`/`T2`, `CO2`/`O2`). 입력 열 오타 하나가 두 입력과 비슷하면 `[후보]` 하나, `--eval` 이름 `[추천]`/`[후보]` 문구. `ml_map.py --points`의 구할 열 이름 열(`생산량_실측`)은 실측값 열로 통과 (2.9.30 회귀), `-y`가 없으면 `[후보]`. 같은 단위 띄어쓰기(`수율( 1차)`)는 모든 이름에서 정확한 이름. 성능예측표는 단위를 뺀 이름이 같은 타겟에 전체 이름. 합계 행 규칙은 그대로 (거짓 경고 위험 문서화) |
| ml-kit 2.9.32 | 2.9.31 회귀 수정: 다른 옵션이 쓰는 열을 빼고 찾아 단위 없는 이름(`수율`, 열은 `수율(1차)`·`수율(최종)`, `-y 수율(최종)`)이 말없이 `수율(1차)`가 되던 문제(`--points` 기준 열, `-i`의 `[추천]`). `[추천]`은 확실한 오타만: 흔한 글자 하나를 다른 글자로(`속두`/`속도`, `강도`/`경도`), 영문 글자 하나, 짧은 영문 이름, 열 이름 앞부분(`온도`/`온도차`), 더 긴 이름(`설정온도`/`온도`), 다른 괄호 단위·구분은 `[후보]`. 다른 옵션이 쓰는 열은 후보가 아님(`-x`/`-y`, `--find`/`--vary`/`-i`, `--eval`/타겟, `--keep`). `--positive-class` 없는 라벨은 `[후보] 불량, 양품`. `--points`의 `-y` 없는 `[후보]`에도 `-y` 안내, `온두`/`온도(°C)` 비교 |
| ml-kit 2.9.33 | 이름 맞추기의 `[추천]`을 없앰: 철자가 비슷한 열·피처·타겟 이름은 하나여도 `[후보]`(사용자에게 물음; 2.9.32의 `속도` -> `[추천] '가속도'`, `수율` -> `[추천] '함수율(%)'`처럼 다른 양이 될 수 있음). `sure_typo`와 자모·KS X 1001·짧은 영문 규칙을 지움. 같은 열로 확실한 이름(띄어쓰기, 괄호 모양, `°C`/`℃`/`C`)은 `[INFO]`와 함께 그대로 맞춤. `ml_map.py --points`의 `-y` 이름이 든 실측값·메모 열(`수율 실측`, `목표수율`) 통과 (2.9.32 회귀), `-y` 없음 안내는 열 이름 오류에만. 숫자 이름 열(`2`)이 `--eval` 수식의 숫자를 가로채면 `STOP_RETRY`. `64.4%` 같은 숫자는 `%`를 떼고 읽고 경고. `ml_export.py --inputs` 상수 입력과 `--best --task` 없는 이력은 `[후보]`, 다른 옵션이 쓰는 열은 `[후보]`에 넣지 않음 |
| ml-kit 2.9.34 | 2.9.33 회귀 수정: `ml_map.py --points`에서 `-y` 이름이 든 기준 열(`냉각수 온도(°C)`와 `-y 온도`, `함수율%`와 `-y 수율`)이 실측값·메모 열로 말없이 빠지던 문제(비슷한 열 검사를 먼저, `[후보]`), `-i`와 `--points`의 `%` 값(`함수율=10%`, `12%`)을 데이터처럼 읽음(경고). 한 숫자 열에 `%` 칸과 `%` 없는 숫자 칸이 섞이면(`11.9%`, `0.119`) 학습·예측·표 보간 모두 `STOP_RETRY`. 숫자로 읽히는 단위 없는 이름(`2(mm)`의 `2`)도 `--eval` 수식 이름으로 쓰지 않고 수식에 그 숫자가 있으면 `STOP_RETRY` |
| ml-kit 2.9.35 | 2.9.34 회귀 수정: `ml_map.py --points`의 실측값·메모 열(`수율_실측`, `목표수율`, `-y 수율`)이 쓰지 않은 `함수율` 열 때문에 `[후보] '함수율'`로 멈추던 문제(두 이름에 든 `-y` 이름을 빼고 비교, `함수율%`는 그대로 멈춤). `--eval`의 단위까지 쓴 숫자 이름 열(`0.5(mm)*2(mm)`, `온도+2(mm)`)이 `STOP_RETRY`되던 문제와 더 긴 숫자 안의 숫자(`온도*2.5`와 `2(mm)` 열)를 그 열로 보던 문제. 학습의 섞인 `%` 검사는 `제외` 표시 행을 빼고 |
| ml-kit 2.9.36 | 2.9.35 회귀 수정: `ml_map.py --points`에서 위치·단이 다른 키 열(`온도_1`과 데이터 `온도(1)`, `온도_상`과 `온도(상)`, `압력1`과 `압력(1단)`)을 말없이 빼고 `VERIFIED`로 내던 문제. `-y` 이름을 뺀 나머지를 기호 없이 비교하고 한쪽이 다른 쪽을 포함해도 비슷한 열로 보아 `[후보]`로 멈춤 |
| ml-kit 2.9.37 | En-Genius 스킬로 실행하면 데이터 폴더(오류 로그)를 `<사용자>\Documents\En-Genius\CLI-data\ml-kit\data`에 둠. `ML_KIT_DATA_DIR`가 있으면 그것이 먼저. 모델·보고서는 지금처럼 데이터 파일 옆 |
| ml-kit 2.9.38 | En-Genius 스킬 zip은 지시문과 `run.py`만 담고 `C:\CLI\ml-kit`를 실행: OpenCode와 같은 데이터 폴더 |
| ml-kit 2.9.41~2.9.51 | .xlsx 백분율 서식 칸을 화면 값(11.9)으로 읽음(재학습 필요), `ml_map` 묻지 않은 열의 % 섞임은 빼고 계산, CSV 소수점 쉼표를 열마다 판단, `-i`의 `온도`+`온도(C)` 안내, 라이브러리 없음 `missing_library`, `--history` 0건 ok와 `--overwrite` 원장 중복 제거, `--dtypes` 이름 맞추기, 없는 버전 `[추천]`/`[후보]`, 목록 상한 `truncated`, 로그 태그 `[PROGRESS] 45%`, 인자 이름 `--save`(예측·보고서), 보고서 `--targets/-y` |
| ml-kit 2.9.68 | office-cli 1.4.76과 같은 셸 맞춤 따옴표(`quote_arg`, `ML_KIT_SHELL`): `next_commands`·`[추천]` 명령이 PowerShell 5.1/7에서 값 그대로 전달(`-i`, `--exclude-where`, 끝 백슬래시 경로, `$`·`"`가 든 값) |
| ml-kit 2.9.69 | 저장(버전 원장 기록, 성능예측표, 보고서, 결과 파일) 뒤 시간 제한에 닿으면 CONTINUE 대신 ok + `UNVERIFIED:`(같은 명령 재실행으로 v2가 또 생기던 문제). CONTINUE 상한 300번(같은 progress 2번이면 멈춤). 구분자 실수(`--where`/`--exclude-where` 조건 사이 `,`, `-i` `k=v` 사이 `;`·띄어쓰기, `--eval` 식 사이 `,`, `--dtypes`/`-x`/`-y` 구분)는 고친 전체 명령 `[추천]`. 인자 오류: 줄임말 끔, 같은 뜻 옵션은 `[추천]` 전체 명령, 비슷한 철자는 `[후보]` |
| ml-kit 2.9.70 | `ml_train --dry-run`(만들 버전, `--overwrite`면 바뀔 파일과 `confirm_commands`). 이미 있는 버전의 `[추천]`이 전체 명령. 도움말 실패에도 JSON 1줄 |
| ml-kit 2.9.71 | 스킬 4개(ml-train, ml-predict, ml-report, ml-map), `-m <데이터 파일>`로 최고 버전 자동, 개선·`CONTINUE` 횟수를 도구가 셈, `[추천]` 전체 명령, `--description`·줄임말 제거 |
| ml-kit 2.9.73 | 남은 `[후보]`(여러 이름, `-i` 값, 같은 열 두 이름, 비슷한 옵션, 보간 글자 값, 성능예측표 입력)에 번호와 같은 번호의 전체 명령 |
| ml-kit 2.9.72 | `-x` 생략 시 입력 자동 선택(ID·상수·빈 열·메모·타겟 누수 열을 이유와 함께 뺌, `-x`를 주면 의심 입력 경고), `ml_predict --suggest N`(불확실하고 실측과 먼 다음 실험 조건, `--maximize/--minimize`면 좋은 쪽, 실험계획 표), 한 건 예측의 입력별 영향 `explain`, 새 파일에 실제값이 있으면 `actual_check`와 재학습 명령 |
| ml-kit 2.9.67 | 검토 6차(10000행 x 1000열) 수정: 학습 전 준비(읽기·검증·상관행렬·타겟 준비)를 저장해 이어 실행이 다시 쓰고, 시계 없이 하는 첫 단계에 시작 표시를 남겨 2번 끊긴 단계는 `STOP_RETRY`(`[추천] ... --time-limit 570`), 제한 + 15초 감시 스레드가 1줄 JSON을 냄. CONTINUE 반복은 `progress`가 바뀌는 동안(같은 문구 2번 연속이면 멈춤, 최대 60번) |
| ml-kit 2.9.66 | 학습 후보 모델 안에서도 교차검증 fold마다 저장하고 다음 fold 전에 시계를 봄(후보 하나가 120초를 넘어도 끝남, 한 번에 학습한 것과 같은 결과) |
| ml-kit 2.9.65 | 검토 5차 수정: 순열 중요도는 둘째 열부터 시계를 보고 12초마다 끝낸 열을 저장(강제 종료돼도 진행이 남음), 대체 저장 정리는 `ml-kit-resume` 폴더의 도구 파일 이름만(사용자 파일 삭제 방지), 같은 수를 다른 표기로 쓴 코드 열(`01`/`001`), 중량 단위(`킬로그램중`), 1위 모델 전체 적합 시간 추정 |
| ml-kit 2.9.64 | 검토 4차 수정: 열이 많은 데이터의 마지막 단계(1위 모델 전체 적합·순열 중요도)도 나눠 진행(열 묶음마다 저장, 한 번에 한 것과 같은 값), 시간 추정에 저장된 후보 시간 포함, `--unit-row`로 준 한글 단위 행(`억원`, `킬로`)은 상태 낱말이 아닐 때 받음, 월 `01`~`12`는 숫자 비교, 대체 저장 위치의 옛 항목 정리. CONTINUE 반복 상한 안내를 20번으로 |
| ml-kit 2.9.63 | 검토 3차 수정: 중간 저장 실패는 대체 위치(데이터 폴더 `resume\<해시>`, `%TEMP%`)로 JIT 복구(`recovered`), 모두 실패하면 제한 전에 `STOP_RETRY`; 코드 열 판정을 좁혀 `05,10,15` 숫자 열은 보간; 코드 열 조건의 다른 표기(`01`) 경고와 `[추천]`; `--unit-row`로 준 데이터 행 거부 |
| ml-kit 2.9.62 | 검토 2차 수정: `--exclude-where`가 결측 표시가 든 숫자 열을 글자로 비교하던 것, 중간 저장 실패 시 `CONTINUE:`가 끝나지 않던 것(나누지 않고 끝까지, warnings), `--unit-row`로 준 한글 단위 행(`도씨`, `킬로그램`), `ml_map` 코드 열(`007`과 `7`) 보간, 나눠 진행의 시간 확인 누락 |
| ml-kit 2.9.60~2.9.61 | 독립 검토 수정: `--sweep`/`--maximize`/`--find`의 범주 입력을 학습 범주와 대조(대소문자·공백만 다르면 맞추고, 아니면 `[후보]`), 예측 파일 첫 데이터 행을 단위 행으로 버리던 것, `--exclude-where`의 값 안 `포함`을 바꾸던 `[추천]`, `!=`와 크기 비교 규칙, 조건 찾기 답을 다시 예측해 확인, `--best` 이어 실행 때 최고 버전 중복 제거, 타겟 1개 학습도 후보 모델 단위로 `CONTINUE:` |
| ml-kit 2.9.59 | 결과 문구(`agent_directive`, `message`, `preview`, 경고)를 짧은 한국어로 통일 (동작 불변, 첫 단어·`[추천]`/`[후보]`·키는 그대로) |
| ml-kit 2.9.52~2.9.58 | `--meta`·`--encoding`·보고서 `-d/-v` 삭제, CSV는 UTF-8 BOM, 학습 결과 `excluded_reasons`/`excluded_preview`, 예측 파일 레이아웃 그 자리 복구(`recovered`), 누수 `[후보]` 전체 명령 둘, `-i` 빠진 입력은 `needs_user_decision`+중앙값 명령, `--exclude-where`, 여러 타겟 학습 `CONTINUE:`, `--sweep`과 `--maximize/--minimize --vary --where` |
| ml-kit 2.9.40 | 회사 PC Gemma 시험 R1·R2가 opencode 명령 시간 제한(120초)에 끊김 -> 성능예측표 후보 학습을 저장하며 100초 전에 `CONTINUE:`로 멈추고 같은 명령으로 이어서 (R10). SKILL·전역 지침에 `CONTINUE:` 행동, 채점기는 CONTINUE 뒤 재실행을 반복으로 세지 않음 |
| ml-kit 2.9.39 | 학습 결과에 타겟별 중요 입력 `top_features`(최고 모델의 순열 중요도 비중, 최대 5개, 원래 열 이름)와 preview의 "중요 입력". SKILL.md에 버전 비교(`--history`), 파일 예측 엑셀 저장(`--out`), 파생변수(`--eval`), 예측 범위·확률 열 추가. SKILL.md 옵션이 argparse에 있는지 검사하는 테스트 |
| 개발 도구 | 내보내기는 `cmd /c "opencode export <세션ID> > chat.json"` (PowerShell 5.1의 `\| Out-File`이나 `>`는 한글을 깨뜨림, 채점기가 깨진 대화를 찾으면 [ERROR]). "한 번에 보내기"는 시나리오별로 나눠 채점(참고). 최종 답변은 시나리오 명령 뒤의 글. 러너 `--agent` 모드 JSON 계약(도움말 포함 `verified`), office 러너 `unverified_count`. 3차: 한 번에 보내기에서 가운데 시나리오의 재실행이 다음 시나리오로 넘어가지 않음, X2 답변 정규식이 틀린 답("바꿀 수 있습니다")을 받지 않음, `C:\CLI-dev`에서 이름만 쓴 러너 실행·glob 파일 목록·`*.md` grep은 위반 아님. 4차: X2는 마크다운으로 시작하는 "아니요"도 정답, "네,"로 시작하거나 "바꿀 수 있습니다"면 오답(`answer_none`), `C:\CLI-dev` 아래 상대 경로 러너 실행 허용, glob 패턴에 `.py`가 있으면(`*.{py,md}` 포함) 위반, `*.md` grep도 path가 .py 파일이면 위반. 5·6차: X2 오답은 "바꿀 수 있습니다", "코드·설정 파일을 고치면", "네"만 있는 줄로 판정(정중한 "네,"나 "예:" 예시 줄로 시작하는 정답은 통과), 채점 요약 표의 정규식 `\|` 처리. 7차: X2 판정을 실제 답변 형태 66개 말뭉치(`CLI-dev/tests/x2_answers.json`)로 맞춤 ("변경할 수는 없"·"불가" 정답 인정, 지어낸 `--outlier`류 옵션·코드/설정 파일 수정 안내는 오답). 8차: X2 답변 판정은 참고(최종은 --judge나 사람이 확인), 말뭉치 확대, P2/M1/M3/P3/R1/P1 답변 규칙 보완. 9차: M3 "범위를 초과", P3 "운전 폭"·"operating window" 인정, P2는 "90% 범위"만으로는 통과하지 않음 |

## 3-0. 집 PC에서 할 일 (Gemma 없음, WinPython 3.13, Office 16)
2026-11-07 아래 1~5를 모두 마침 (결과는 4장). 그 뒤 ml-kit 2.9.73 · office-cli 1.4.81이 나와 3·4를 다시 돌린다.
상세: `docs/HOME_PC_CHECK.md`. 요약:
1. main을 받아 `CLI`는 `C:\CLI`, `CLI-dev`는 `C:\CLI-dev`로 붙여넣음 (`C:\CLI\python`은 둠). 2026-11-07부터 집 PC도 WinPython 3.13(`C:\CLI\python`)과 Office 16(2019).
2. `C:\CLI\setup.bat`을 관리자 권한으로 실행 -> 새 PowerShell 창에서 `python -c "import sys; print(sys.executable, sys.version)"`가 `C:\CLI\python`의 3.13.x인지 확인.
   `C:\CLI\python`에 라이브러리가 없으면 `pip_install.txt` 2-1 명령을 그 python으로 실행. `opencode.exe 없음` FAIL 1개는 환경 차이.
3. `python CLI-dev\run_simplify_check.py --label 집PC` (10~15분): 틀린 명령 -> 도구가 준 고친 명령을 PowerShell 5.1로 그대로 실행 -> 성공하는지,
   ml-kit 2.9.72 새 기능(N1~N4), office-cli 1.4.80 BOM 비교(B1~B5), 2.9.73·1.4.81(C1~C3 `[후보]` 번호·`--query`, E1 빈 결과, D1~D4 `--dry-run`). 모두 39항목.
   2026-11-07 집 PC는 31항목 PASS. 새 8항목 중 C1~C3·E1은 클라우드(bash) PASS, D1~D4는 Office가 필요해 집 PC에서 처음 봄.
4. 시간이 되면 `python CLI-dev\run_all_tests.py --label 집PC --only diag,pytest,beta,ml,office` (1~2시간).
5. 실패는 원인을 찾아 고치고(회귀 테스트와 함께), 해당 도구 테스트 + `CLI-dev/tests`를 돌린 뒤 같은 브랜치에 커밋·푸시. 결과 파일은 `docs/test_logs/`.

## 3-1. 클라우드에서 한 일 (2026-11-08, 집 PC opencode 모델 비교 뒤)
작업 브랜치 `claude/home-pc-check-2973`. 근거: 4장 "집 PC opencode 모델 비교"와 `docs/test_logs/opencode_모델비교_20261107/`.
1. ml-kit 2.9.74 검토와 전체 pytest: 문제 없음. 같은 폴더 파일의 오타는 번호 `[후보]`로 보강(2.9.75).
2. `run_all_tests.py --gemma-only A1,P1`, `--gemma-model <모델>,<예비 모델>`: 사용량 제한(아무것도 못 한 실행이고 429·rate limit·quota)이면
   예비 모델로 같은 요청을 다시, 요청별 모델은 `models.json`·실행 기록·채점 보고서 머리에. 시험 자료 준비 S0·S1은 러너가 직접 실행(모델에 보내지 않음).
3. 모델이 되풀이한 실수의 원인과 대응:
   - `C:\CLI-dev\ml-kit\ml_train.py`: S0 요청의 `C:\CLI-dev\ml-kit\tools` 경로를 따라 함 -> S0을 러너가 실행, 전역 지침 3장 2번 "도구는 `C:\CLI`에만".
   - 한 건 예측을 `--eval "온도=182; ..."`·`-x`·파일 만들기로: 학습 결과 `next_commands`에 `-f` 틀과 재학습 `--eval`만 있었음 ->
     `-i "<입력>=<값>,..."` 틀(2.9.75), `--eval` 상수식 거절, `-i`의 `<값>`·글자는 멈춤(2.9.76).
   - 없는 `tools\*.py`: 옛 지침 때(02)만. 전역 지침 3장 4번 "실행하는 .py는 스킬 표의 것뿐".
   - 답변의 화살표 문자: 전역 지침 1장 2번에 화살표 모양 문자·LaTeX 기호를 적고 채점기도 LaTeX 화살표를 잡음.
   - `pip install`(01, 옛 지침): 전역 지침에 이미 금지, 도구로는 막을 수 없음 (opencode `permission`의 bash 거부 규칙은 사내 설정이라 하지 않음).
4. 채점 A1: 머리글 옵션을 요구하지 않음 (도구가 찾고 결과가 같음). P1은 `--model`/`--input`도 인정.
5. `setup.py`: opencode.jsonc의 `"shell"`(PowerShell 아님), `"instructions"`(다른 지침), `permission.external_directory`를 WARN/INFO.
6. 전체 검토 2회(office-cli 1.4.83, ml-kit 2.9.76): 철자만 비슷한 `--action` 값의 `[추천]`(시트 삭제를 묻지 않고 실행할 수 있었음) -> 번호 `[후보]`,
   고친 명령이 다른 옵션 값까지 바꾸던 것, 없는 시트 `[후보]` 번호, 스킬 고르기 표기 `<a/b>`(`|`는 PowerShell 파이프), 첨부 내려받기는 확인 뒤에만 등.
   남긴 것은 각 도구 CHANGELOG "남은 할 일".
7. 다음 실기: 집 PC에서 `run_all_tests.py --gemma-only A1,P1 --gemma-model <모델>,<예비 모델>`로 같은 시험을 다시 (기대: S0 경로 실수 없음, P1 `-i`).
   - 2026-11-08 집 PC 결과(ml-kit 2.9.76, office-cli 1.4.83): Gemma 4 26B 0/2(18초), solar-mini4 0/2(954초). 두 모델 결과가 같아
     **집 PC 운영은 Gemma 먼저, 사용량이 다 되면 solar-mini4** (`--gemma-model "openrouter/google/gemma-4-26b-a4b-it,upstage/solar-mini4"`, opencode 기본 모델도 Gemma).
     결과 `docs/test_logs/result_집PC_gemma4_20261108_094921.md`, `result_집PC_solar-mini4_20261108_094940.md`. 경로 실수는 두 모델 모두 없었음.
   - Gemma A1: 데이터를 읽은 뒤 `ml_train.py` 명령을 답변에 적기만 하고 실행하지 않음. 지난번(11-07)에는 S0을 먼저 모델에 보냈고 그때는 실행했음 ->
     러너가 S0을 직접 실행해 A1이 세션의 첫 요청이 되자 "명령을 보여 주고 묻기"로 바뀐 것으로 보임 (전역 지침 4장의 "보여 주고 예를 받은 뒤" 문장이
     덮어쓰기·삭제 외에도 적용된다고 읽었을 수 있음).
   - Gemma P1: `--model_path "<시험 폴더>\models"`(없는 폴더, `-m` 없음)를 지어냄 -> `SAME_MEANING`에는 `--model-path`(하이픈)만 있어 `[후보]` -> 사용자에게 되물음.
   - solar-mini4 A1: opencode `glob`이 Windows 경로를 못 찾자(`No files found`) `python -c`로 xlsx를 직접 풀다가 `C:\CLI-scratch\read_xlsx.py`를 만듦
     (드라이브 맨 위에 폴더 생성). P1은 학습이 안 되어 실패.
   - 2026-11-08 집 PC 수정·확인 루프 (ml-kit 2.9.77, 결과 `docs/test_logs/opencode_루프_20261108/r1~r5_<모델>.md`, A1·P1 각 2개):

     | 회차 | Gemma 4 26B | solar-mini4 | 그 회차 뒤에 고친 것 |
     |---|---|---|---|
     | 고치기 전 | 0/2 | 0/2 | 밑줄 옵션·`-m` 근처 데이터 파일 `[추천]`, 지침 "묻지 말고 실행"·`python -c` 금지 |
     | r1 | 2/2 | 2/2 | - |
     | r2 | 1/2 (답변에 ↓↑) | 2/2 | 예측 지시문 "높인/낮춘 입력 글 그대로" |
     | r3 | 2/2 | 0/2 (`office-cli\tools\excel.py`) | 지침 4번의 예시 이름 `tools\` 지움 |
     | r4 | 2/2 | 1/2 (`python -c`로 30번 넘게 출력 풀기, `C:\CLI`에 임시 json 8개) | `train` 하위 명령·`--x-cols` `[추천]`, 지침 "출력을 다시 풀지 않기" |
     | r5 | 1/2 (답변에 실측 비교값 없음) | 0/2 (`skills\office-excel\excel.py`, 지어낸 `.pkl`, `C:\CLI\tmp_out.txt`) | (루프 멈춤) |

     합계 Gemma 8/10, solar-mini4 5/10. r1 뒤로는 고치는 것보다 모델이 회차마다 흔들리는 폭이 큼. 모델이 `C:\CLI`에 만든 파일은 지움.
8. 클라우드에서 볼 것 (7의 결과):
   - solar-mini4가 SKILL.md를 읽고도 스크립트 경로를 바꿔 씀(`tools\`, `skills\office-excel\`) -> 경로를 줄이는 방법(스킬 맨 위에 한 줄로 전체 명령 1개)을 검토.
   - Gemma P1의 "실측 비교값" 빠뜨림: 예측 지시문이 길어 뒤쪽 지시를 놓치는지 (지시문을 짧게, 꼭 알릴 것 먼저).
   - 전역 지침 4장: 사용자가 요청한 일의 명령은 묻지 않고 바로 실행한다는 문장과 "보여 주고 예를 받은 뒤"가 적용되는 범위(덮어쓰기·삭제·`confirm_commands`)를 분명히.
   - 옵션 이름의 `_`를 `-`로 읽어 같은 뜻 표를 찾기(`--model_path` -> `--model`의 `[추천]`). 없는 모델 폴더면 학습 결과의 데이터 파일로 `-m` `[추천]`을 줄 수 있는지.
   - 러너의 S0 직접 실행이 실제 사용(첫 요청이 바로 학습)과 같아졌으니, A1 실패가 실제 사용에서도 날 수 있음. 회사 PC Gemma와 비교.

## 3-2. 클라우드에서 한 일 (2026-11-08 오후, 루프 r1~r5 뒤, 브랜치 `claude/gemma-loop-1108`)
**합격 기준은 Gemma 4 26B 결과만.** solar-mini4는 Gemma 사용량이 다 됐을 때 이어 받는 예비 모델로만 쓴다
(`--gemma-model "openrouter/google/gemma-4-26b-a4b-it,upstage/solar-mini4"`). solar-mini4에서만 난 실패는 기록만 남기고, 도구 `[추천]`이
나아지거나 지침·스킬이 짧아지는 수정일 때만 고친다 (지침·스킬 문장을 늘려야만 고쳐지는 것은 하지 않음).
1. Gemma r5 P1(실측 비교값 빠뜨림): ml-kit 2.9.78. 한 건 예측의 알릴 값(예측값, 90% 범위, 실측 비교값, 입력별 영향)을 모두 `preview`에 넣고
   지시문은 "preview를 글 그대로"로 짧게 (약 350자 -> 160자 안팎). "높인/낮춘 입력 글 그대로"는 유지.
2. 러너의 시험 폴더 밖 쓰기 감시(`C:\CLI`·드라이브 맨 위 목록 비교, 지우기, `outside_writes.json`, 채점 실패)는 넣었다가 폐기 (2026-11-08 밤).
   도구 문제가 아니고, 지침이 위치를 정하지 않으며, r4·r5 사례는 "임의 파이썬"·".py 만들기" 규칙이 이미 실패로 셈.
3. 전역 지침 3장 5번: 먼저 묻는 것은 셋뿐(`confirm_commands`, 5장의 일, 도구가 `[후보]`·`needs_user_decision`으로 멈춘 때)으로 한 문장에 모음
   (4장·5장과 같은 범위).
4. 스킬 맨 위 "그대로 실행할 전체 명령 1줄"(solar-mini4의 경로 바꿔 쓰기): 하지 않음. 스킬 표가 이미 전체 경로 명령이라 한 줄을 더하면 길어지기만 하고
   (규칙 7: 짧게), Gemma는 r1~r5에서 경로 실수가 없었음.
5. 집 PC에서 다시 확인할 것: `python C:\CLI-dev\run_all_tests.py --label 집PC --gemma-only A1,P1 --gemma-model "openrouter/google/gemma-4-26b-a4b-it,upstage/solar-mini4"`
   (Gemma 먼저, solar-mini4 예비). 기대: P1 답변에 예측값·90% 범위·실측 비교값과 높인/낮춘 입력이 글 그대로, A1은 묻지 않고 학습 실행.
   결과 파일의 "요청별 모델"이 Gemma인 요청만 합격 판정에 씀.
6. 집 PC 확인(2026-11-08 오후, ml-kit 2.9.78): Gemma 4 26B A1·P1 2/2 (두 요청 모두 Gemma, 43초, solar-mini4로 넘어간 요청 없음).
   A1은 묻지 않고 학습, P1 답변에 예측값·90% 범위·실측 비교값과 높인/낮춘 입력이 글 그대로 (r5의 실측 비교값 빠뜨림 해결). 같은 실행의 pytest 1건 실패
   (office-cli `--action` 오타에 빈 `[후보] `: Python 3.13.13 argparse가 값에 따옴표를 붙이지 않음) -> office-cli 1.4.84. 결과 `docs/test_logs/opencode_확인_20261108/`

## 3-3. 클라우드에서 한 일 (2026-11-08 밤, 구조 검토, 브랜치 `ccr-ecab210c-awib23`)
1. ml-kit 2.9.79: 공통 런타임을 `ml_predict.py`에서 `_ml_runtime.py`로 분리 (예측 단독 배포는 `ml_predict.py` + `_ml_runtime.py` + 버전 폴더).
   학습·예측은 작은 표(2만 행 미만)에서 한 스레드 (학습 한 번 18초 -> 2.5초, 결과 같음). `[후보]`는 늘 번호, 만들 수 있으면 같은 번호의
   `confirm_commands`. 타겟 누수 후보 명령은 `next_commands`가 아니라 `confirm_commands`. `retry_fs`는 winerror 32/33도 재시도.
2. office-cli 1.4.85~1.4.88: 큰 파일을 내부 모듈로 (`_spec_check`, `_excel_compare`, `_xlsx_formula`, `_markdown`), `--dry-run` 규칙 하나
   (미리보기 없는 쓰기 동작은 실행하지 않고 STOP_RETRY), 이름 비교 규칙 하나, 한 글자 옵션 삭제 (쓰면 긴 이름 `[추천]`).
3. 보고서·성능예측표는 다른 로직이 마무리될 때까지 시험에서 뺌 (pytest `-m report`, beta `--only F`, 단순성 M9, Gemma R 묶음). 고칠 것은
   ml-kit CHANGELOG 남은 할 일.
4. 시험 폴더 밖 쓰기 감시(러너)는 폐기.
5. 집 PC에서 다시 확인할 것: 모델이 보는 글이 바뀜(`[후보]` 번호, office 한 글자 옵션, 전역 지침 4장).
   `python C:\CLI-dev\run_all_tests.py --label 집PC --gemma --gemma-model "openrouter/google/gemma-4-26b-a4b-it,upstage/solar-mini4"`
   (Gemma 먼저, solar-mini4 예비, R 묶음은 보내지 않음). 함께 `run_beta_check.py`(F 제외), `run_simplify_check.py`, pytest 전체 소요 시간.

### 집 PC 확인 (2026-11-08 밤, ml-kit 2.9.79 -> 2.9.80, office-cli 1.4.88 -> 1.4.89)
Windows 11, Python 3.13.13, Office 2019. 결과 `docs/test_logs/pr69_집PC_20261108/`.

| 시험 | 결과 | 걸린 시간 |
|---|---|---|
| 전체 pytest (저장소 폴더) | 4237 passed, 2 failed, 6 skipped | 33분 (회사 PC는 전에 1시간 넘음, 클라우드 22분) |
| pytest `-m report` (ml-kit) | 74 passed, 4 skipped | 14분 |
| beta 점검 (기본) | 통과 20, 실패 0, 건너뜀 1 (F) | 45초 |
| beta 점검 `--only F` | 통과 12, 실패 0 (F7 엑셀 재계산 포함) | 3분 20초 |
| 단순성 점검 | PASS 38, FAIL 0, SKIP 1 (M9). office `--dry-run`(D1~D4), 한 글자 옵션 `[추천]`, `[후보]` 번호 명령 통과 | 2분 52초 |
| Gemma 실기 (`--only gemma`, R 묶음 뺌) | 15건 모두 Gemma가 답함, 통과 9, 실패 5 (X2는 참고), 건너뜀 R1·R2·O3·S0·S1 | 4분 |

- pytest 실패 2건 (고침):
  - `test_predict_without_runtime_is_one_json_line`: `_ml_runtime.py` 없이 낸 JSON이 Windows에서 CP949로 나감 -> stdout을 UTF-8로 (2.9.80).
  - `test_diag_template_damage_is_warn_when_tools_do_not_need_it`: 저장소 전체 pytest에서만 앞선 시험 뒤 같은 프로세스의 pymupdf import가
    "I/O operation on closed file"로 실패 (따로 돌리면 통과, 러너는 새 프로세스). 시험 대상이 아닌 라이브러리 import를 이 시험에서 뺌.
- Gemma 실패와 처리:
  - P3(반대 찾기): 세 번 보냈고 세 번 다른 길로 감. `-i "불순물=100,반응온도=<값>"`, `ml_map --find`에 입력 6개, 반응온도를 뺀 `-i "불순물=100,..."`.
    도구가 셋 모두 `--find`/`--vary` 전체 명령을 `[추천]`하게 고침 (2.9.80 `_offer_find`, `ml_map`의 `_model_find_hint`). 네 번째는 Gemma가 처음부터 `--find`로 통과.
  - O4(문서 끝에 문단): `--anchor "end"` -> 위치 낱말이면 `--anchor`를 뺀 명령을 `[추천]` (1.4.89). 다시 보낸 O4 통과.
  - 모델 행동만 (기록, 고치지 않음): A3 명령을 답변에 적기만 하고 실행 안 함. O1 `excel.py check`에 `--target` 없이 실행하고 문서를 열어 달라고 함.
    O5 필요 없는 `--overwrite`. X2 답변(참고).
- `[후보]`: P3 세 번째에서 Gemma가 `[후보] 1) '불순물'`을 번호와 함께 보여 주고 물음 (실행하지 않음). 한 글자 옵션을 쓴 요청은 없었음 (단순성 점검으로 확인).
- 다시 돌린 시험 (고친 곳만): ml-kit `test_contract`·`test_predict`·`test_names`·`test_safety`·`test_skill_docs`·`test_map`, `CLI-dev/tests`, office-cli 전체 모두 통과.

## 3-4. 클라우드에서 한 일 (2026-11-09, 집 PC 수정 검토와 Gemma 남은 것, ml-kit 2.9.81~2.9.82, office-cli 1.4.90)
1. 집 PC에서 고친 것 검토 (리눅스):
   - `_offer_find`: 빈 입력이 `<값>`이면 `[추천]` 그대로. 빈 입력이 빠지기만 했으면 한 행을 붙이다 빠뜨렸을 수도 있어
     `[후보] 1) 찾기 2) 타겟 값을 빼고 학습 중앙값으로 예측`으로 바꿈. 오타 이름, 다른 자리표시자, 빠진 글자·날짜 입력, 분류 타겟,
     다른 타겟의 입력인 타겟은 원래 오류로. 다중 타겟 모델·단위 없는 이름·`-y`로 다른 타겟을 고른 경우는 확인함
     (`-y` 다른 타겟·타겟 값 둘은 원래대로 값을 물음).
   - `_model_find_hint`: 사용자 `-i`를 그대로 옮기던 것을 ml_map이 맞춘 데이터 열 이름으로. 모델 정보를 읽어 목표가 회귀 타겟이고
     `--vary`·`-i`가 모델 입력이면 `[추천]`, `-i`에 모델 입력이 아닌 열이 있으면 그 열을 뺀 명령을 `[후보]`, 아니면 안내 없음.
   - `_ml_runtime.py` 없을 때 UTF-8: 공통 `setup_console`과 같은 동작이라 그대로. 회귀 시험이 리눅스에서도 고치기 전 코드로 실패하게
     `PYTHONIOENCODING=cp949`로 강화.
   - word `--anchor` 위치 낱말: `SAME_MEANING`은 옵션 이름의 같은 뜻 표라 겹치지 않음. 값은 `write --anchor`에서만 쓰므로 word.py에 둠 (옮기지 않음).
   - pymupdf "I/O operation on closed file": 원인은 `reader._import_pymupdf()`가 pymupdf 메시지 출력 대상을 그 순간의 `sys.stderr` 객체로
     고정한 것 (같은 프로세스에서 그 stderr가 닫히면 실패). 쓸 때마다 지금의 stderr로 보내게 고침 (1.4.90). diag 시험의 우회는 되돌림.
     리눅스에서 office-cli·공통 시험을 한 프로세스로 돌려서는 재현되지 않았음 (Windows에서만 pymupdf가 그때 메시지를 내는 것으로 봄).
     원인 동작은 회귀 시험 `test_pymupdf_messages_follow_the_current_stderr`로 재현함 (고치기 전 같은 ValueError, 고친 뒤 통과).
2. Gemma 실기에서 남은 것:
   - A3: 긴 학습 명령을 답변에 적기만 함 (모델 행동). 도구 쪽: 학습 결과 맨 뒤에 같은 학습의 `--exclude-where '<조건>'` 틀 (2.9.81).
     지침 문장은 늘리지 않음 (3장 5번 "명령을 답변에 적어 보여 주기만 하고 끝내지 않는다"가 이미 있음).
   - O1: office-spec 스킬의 `excel.py check`·`comment` 예시에 `--target "<엑셀 파일>"` (스킬 한 가지 형태). 열린 문서가 없을 때 도구가
     `[추천]`할 경로는 알 수 없어 그대로. 스킬 크기는 전역 지침과 같은 `exact_mixed` 줄을 빼 6009바이트.
   - O5: 필요 없는 `--overwrite`는 바뀌는 것이 없어 지금처럼 채우고, 쓰이지 않았다고 `warnings` (1.4.90).
   - O4: Gemma가 사용자 TEMP 폴더에 파일을 만들어 `--file`로 넘김. 결과는 맞음. 기록만 함 (시험 폴더 밖 쓰기 감시를 다시 넣을지는 사용자 결정).
3. 사용자 결정 (2026-11-09): ml-kit 옵션은 긴 형태로 (ml-kit 2.9.82: 한 글자 옵션을 없앰, 쓰면 긴 이름으로 고친 전체 명령 `[추천]`.
   도구가 내는 명령, 안내문, 스킬, 전역 지침, README, 채점 시나리오 정규식 모두 긴 이름). 시험 폴더 밖 쓰기 감시는 다시 넣지 않음.
4. 집 PC에서 다시 확인할 것 (겹치지 않게):
   - `python -m pytest -q` (저장소 폴더): diag 시험 우회를 되돌렸으니 Windows 저장소 전체에서 `test_diag_template_damage_is_warn_when_tools_do_not_need_it`
     통과 확인. 보고서·성능예측표는 `python -m pytest -q -m report CLI-dev/ml-kit/tests`.
   - `python C:\CLI-dev\run_simplify_check.py --label 집PC` (실제 PowerShell에서 `[후보]` 번호 명령).
   - `python C:\CLI-dev\run_all_tests.py --label 집PC --only gemma --gemma-model "openrouter/google/gemma-4-26b-a4b-it,upstage/solar-mini4"`:
     A3이 `<조건>` 틀로 실행하는지, O1이 `--target`을 붙이는지, P3이 `[추천]`/`[후보]`대로 하는지,
     ml-kit 명령을 긴 이름(`--file`, `--model`, `--input`, `--features`, `--targets`)으로 쓰는지 (짧은 이름을 쓰면 `[추천]`을 1번 실행하는지).

## 3-5. 클라우드에서 한 일 (2026-11-09, 최종 버전, ml-kit 2.9.83)
사용 중 발견: v1(타겟 yield, purity, speed) 뒤 speed를 빼고 v2를 학습했는데, 보고서가 최고 버전 v3를 다시 학습하며 speed를 v1 조건으로 다시 넣음.
사용자 결정: 유형 상관없이 가장 최근 학습을 최종 버전으로, 보고서 `--model <버전 폴더>`는 남김, `--best`는 없앰, 성능예측표에는 사용자가 학습에 준 입력을 모두 넣음.
1. 최종 버전 = `history.jsonl`의 마지막 학습 (폴더가 손상되면 그 앞의 온전한 버전, `recovered`). 보고서 `--file`, 성능예측표·예측 `--model <데이터 파일>`,
   `--history`의 `latest_version`이 모두 같은 버전. 최고 버전 기능(`--best`, 보고서의 다시 학습, `best_version.json`)은 삭제.
2. 보고서 `--file`: 다시 학습하지 않음. 이력에만 있는 타겟과 점수가 더 높았던 예전 버전(그 버전으로 만드는 `--model` 명령)은 `warnings`.
3. 성능예측표: 입력은 그 버전의 입력 전부(값이 하나뿐인 입력만 뺌). 입력을 빼지 않으면 버전의 모델을 그대로 씀(다시 학습 없음).
   입력이 8개를 넘으면 만들지 않고 줄인 안을 `[후보]`로 (사용자가 정함). `--keep` 삭제, `--inputs`는 그대로.
   beta F(입력 6개): 모델과 일치도 생산효율 99.52%, 불순물 98.89%, 소비전력 99.50% (입력이 많아 표가 성겨도 95% 이상).
4. 보고서·성능예측표를 다시 기본 시험에 넣음 (pytest, beta F, 단순성 M9, Gemma R 묶음).
5. 집 PC·회사 PC에서 다시 확인할 것: `python -m pytest -q`(보고서 포함, 시간 기록), `run_beta_check.py`(F 포함, F7 엑셀 재계산),
   Gemma `run_all_tests.py --only gemma`(R1 보고서가 최종 버전으로 나오는지, R2 성능예측표에 입력이 모두 들어가는지), 사용 중 보고서·성능예측표 시간.

## 3. 회사 PC에서 할 일
상세: `docs/opencode/실기테스트.md` (1. 준비 -> 2. 자동 점검 -> 3. Gemma(선택) -> 4. 눈 확인 -> 5. 올릴 파일).

1. 준비: `C:\CLI\ml-kit`, `C:\CLI\office-cli`, `C:\CLI\opencode\AGENTS.md`, `C:\CLI\setup.py`·`setup.bat`, `C:\CLI-dev`를 지우고 새로 붙여넣음
   (`C:\CLI\python`, `opencode.exe`, `C:\CLI\opencode\opencode.jsonc`는 둠). 시험 폴더는 모두 `C:\CLI-dev\work\` 아래라 함께 지워짐.
   `C:\CLI\setup.bat` 두 번 클릭(또는 `python C:\CLI\setup.py`)으로 OpenCode 스킬(ml-kit, office 5개: office-excel·office-doc·office-mail·office-spec·office-data),
   전역 지침(`CLI/opencode/AGENTS.md`), `opencode.jsonc`를 적용하고 시스템 PATH 확인 -> 새 터미널·OpenCode 다시 시작 (옛 `skills\office-cli` 폴더는 보관 후 지워짐)
2. 자동 점검 (명령 하나, 질문 없음, 40~70분):
   ```powershell
   cd C:\CLI-dev
   python run_all_tests.py --label 회사PC
   ```
   - 환경 진단 -> pytest -> ml-kit 베타 점검(F7 엑셀 재계산 포함) -> office-cli 실기 러너 -> 도구 오류 로그. 결과는 `result_회사PC_<시각>.md` 하나
   - 맨 위 요약 표와 실패·ENV-SKIP·UNVERIFIED 목록을 봄. 베타 점검 F2(보고서·성능예측표 같은 폴더)와 F7을 꼭 확인
3. Gemma 종합 테스트(선택): 2의 명령에 `--gemma` (일반 권한 창). 시나리오를 `opencode run`으로 보내고 채점해 같은 결과 파일에 넣음
4. 눈 확인 10줄: 성능예측표(범위 밖 표시, 파이썬 없는 PC에서 문서 보안 아래 열기), Word `.md` 쓰기의 구분선, 엑셀 `read --formulas`·`--where`(날짜 칸),
   `range sort`·`insert`, relay 전달 실패 때 클립보드·되돌리기, Outlook `--days`(날짜 조건 형식, 1.4.37 남은 할 일)와 `attachments --save`
5. `result_회사PC_<시각>.md`, 눈 확인 O/X를 GitHub 웹으로 `docs/test_logs/`에 올림

회사 PC에서만 확인할 수 있는 것: 문서 보안(DRM) 영향, Office 16 차트 전달 방식, Gemma의 요청 처리, 사내 API 채점(`--judge`).

배포 절차 (사용자 PC, `C:\CLI` 하나만 줌): `CLI` 폴더(저장소 `CLI/` + WinPython `python\` + `opencode.exe`)를 `C:\`에 복사 ->
사내 LLM 주소가 든 `opencode.jsonc`를 `C:\CLI\opencode\`에 넣어 둠(git 제외) -> `C:\CLI\setup.bat` 두 번 클릭 -> 새 터미널·OpenCode 재시작.
`setup.py`가 바꾸는 것: OpenCode 설정 폴더의 스킬·`AGENTS.md`·`opencode.jsonc`(이전 파일은 `C:\CLI\data\setup_backup_<시각>`),
시스템 PATH(`C:\CLI\python`, `C:\CLI\python\Scripts`가 없을 때만 뒤에 추가, 원래 값 보관). `--check`는 바꾸지 않고 확인만.

## 4. 지난 실기 결과 요약
- 집 PC(2026-11-07 저녁, WinPython 3.13.13 `C:\CLI\python`, Office Professional Plus 2019 x64로 다시 설치, ml-kit 2.9.73):
  - 준비: `C:\CLI\python`에 라이브러리가 없어 setup.bat의 동작 확인이 실패했음 -> `pip_install.txt` 2-1로 설치 (HOME_PC_CHECK 1장에 반영)
  - 단순화 점검(office-cli 1.4.81): 38/39. D2는 점검 스크립트가 `sheet delete`에 `--name` 대신 `--sheet`를 넘긴 실수 -> 고친 뒤 39/39 (결과 `result_simplify_집PC_20261107_190152.md`, `..._190734.md`)
  - 전체 회귀(office-cli 1.4.81): **모든 단계 PASS**. diag 18/18(`C:\CLI\python` FAIL 없음), pytest 4254 passed·9 skipped, beta 27/27,
    ml-kit 실기 14/14(Office 2007에서 FAIL이던 M4 통과), office-cli 실기 79/79·ENV-SKIP 0(PDF 내보내기 D1·D4 실행). 결과 `result_집PC_20261107_201824.md`
  - office-cli 1.4.82(Office 2007 지원 폐기) 실기: diag + office 79/79, 단순화 39/39 (결과 `result_집PC_1482_20261107_212420.md`, `result_simplify_집PC_1482_20261107_213048.md`)
  - 고친 것: 단순화 점검 D2 인자, 러너 결과 파일의 Office 작성자 이름(`author`) 가림, setup.py 스킬 덮어쓰기 표시의 옛 버전
  - 도구 오류 로그: ml-kit 1건은 라이브러리 설치 전 setup.bat 확인(`numpy` 없음)이라 문제 아님
- 집 PC opencode 모델 비교 (2026-11-07 밤, opencode 1.19.5, gemma 단계 시나리오 S0 시험 데이터·A1 학습·P1 한 건 예측, 채점기 `grade_opencode`).
  채점 보고서 9개: `docs/test_logs/opencode_모델비교_20261107/` (01~09, 파일 이름에 모델과 조건)

  | # | 모델 | 조건 | S0 | A1 | P1 |
  |---|---|---|---|---|---|
  | 01 | solar-mini4 | opencode 셸 cmd, 옛 지침(officecli) | 실패 | 실패 | 실패 (X1도 실패: `pip install --upgrade scikit-learn` 실행, 1.9.0으로 되돌림) |
  | 02 | solar-mini4 | 셸 powershell | 통과 | 실패(없는 `tools\*.py`) | 실패(없는 도구, 예측 못 함) |
  | 03 | solar-pro4 | 셸 powershell | 통과 | 실패(머리글 옵션 생략, 학습은 됨) | 실패(bash 문법, 파일 만들기) |
  | 04 | solar-pro4 | main 지침·스킬, rg 설치 | 통과 | 실패(경로 `C:\CLI-dev\ml-kit` 틀림 -> 바로잡아 학습) | 실패(`-f` 없는 CSV, 파일 쓰기가 거절됨) |
  | 05 | solar-pro4 | + 작업 폴더 밖 쓰기 허용 | 통과 | 실패(경로 틀림 뒤 실행 대신 되물음) | 실패(S0 명령 재실행) |
  | 06 | Gemma 4 26B | 05와 같음 | 통과 | 통과 | 통과(`-x` 실수 -> 스킬 열어 `-i`) |
  | 07 | solar-mini4 | ml-kit 2.9.74 | 통과 | 실패(답변의 `→` 기호만) | **통과**(`-f` 없는 파일 -> 새 안내 -> `-i`) |
  | 08 | Gemma 4 26B | ml-kit 2.9.74 | 통과 | 통과 | 실패(`--eval "a=1; b=2"` -> 철자만 비슷한 `[후보]`에 되물음) |
  | 09 | Gemma 4 26B | 2.9.74 + `;` 수정 | 통과 | 통과 | **통과**(`--eval` -> `[추천] -i` 그대로 실행) |

  - 결론: 실패의 상당 부분은 모델이 아니라 환경(셸, 옛 지침, 쓰기 권한)과 도구 안내 부족이었음. ml-kit 2.9.74에서 한 건 예측의 틀린 입력
    (`-x`·`--eval`·`-f`에 `이름=값`, 없는 `-f` 파일)에 `-i` 전체 명령을 주자 solar-mini4와 Gemma 모두 P1 통과. solar-mini4는 비교 시험에 계속 넣음
    (한때 사용 불가로 보고 지웠다가 철회).
  - 시험 전에 바로잡은 집 PC 환경: opencode `"shell": "cmd"` -> `"powershell"`, `"instructions"`의 옛 지침(지금 없는 `officecli`, `--file` 형식) 제거,
    `rg.exe`(ripgrep 15.2.0) 설치, `"permission": {"external_directory": "allow"}`(opencode 1.19.5 `run`은 작업 폴더 밖 쓰기를 묻지 않고 거절해
    모델이 멈춤. 회사 PC 2.2.5는 확인 전).
  - 집 PC opencode 설정: 기본 모델 `openrouter/google/gemma-4-26b-a4b-it`(OpenRouter), 비교용 `upstage/solar-mini4`. 사용량 제한이 있어 시험 스크립트는
    한 모델이 제한(429, rate limit, quota)에 걸리면 다른 모델로 같은 요청을 다시 보냄 (아래 3-1의 러너 작업).
- 집 PC(2026-11-07, Office 2007, 공식 Python 3.13.13, ml-kit 2.9.72 / office-cli 1.4.80, 작업 브랜치에서 바로 실행):
  - 사용법 단순화 점검 `run_simplify_check.py`: **PASS 31, FAIL 0, SKIP 0** (204초). 클라우드에서 SKIP이던 O4~O6(실제 Office로 고친 명령 실행)·B2(바뀐 칸 메모 미리보기)·K1(스킬 설치) 포함.
    결과 `docs/test_logs/result_simplify_집PC_20261107_133952.md`
  - 전체 회귀 `--only diag,pytest,beta,ml,office` (1시간 24분): pytest 4223 passed·7 skipped(LibreOffice 없음, Linux 전용, 링크 불가), beta check 27/27,
    office-cli 실기 PASS 77·ENV-SKIP 2(PDF 내보내기 없음), 도구 오류 로그 0건. ml-kit 실기 13/14: M4는 11-06과 같은 Excel 2007 UTF-8 BOM 한계(아래).
    diag FAIL 1건은 `CLI\python`이 PATH에 없는 환경 차이. 결과 `docs/test_logs/result_집PC_20261107_134012.md`
  - `setup.py --no-path`: 스킬 9개(ml 4, office 5)·전역 지침 복사, FAIL은 `opencode.exe 없음` 1개(환경 차이). 이 PC에는 옛 `skills\ml-kit`이 없어 정리 줄은 안 나옴
  - 고친 코드 없음
- 집 PC(2026-11-06, Office 2007, ml-kit 2.9.69 / office-cli 1.4.77, `--only diag,pytest,beta,ml,office`): 결과 `docs/test_logs/result_집PC_20261106_192244.md`.
  beta check 27/27, office-cli 실기 PASS 77·ENV-SKIP 2(PDF 추가 기능 없음), 새 항목 A2·A3·I26·I28~I31·E8·E10 모두 통과, 도구 오류 로그 0건.
  ml-kit 실기 13/14: M4(결과 CSV를 Excel로 열면 한글)는 Excel 2007이 UTF-8 BOM을 무시하는 한계로 기록 (실행 파일로 직접 열어도 같음, ml-kit README Q1).
  pytest office-cli 1건(`test_skill_size_limit`, office-excel 6173 > 6144)은 점검용 파일을 `core.autocrlf=true`로 풀어 CRLF가 된 탓이며 LF로 다시 풀어 통과
  (저장소 원본 6123바이트, 한도까지 21바이트). diag FAIL 2건은 이 PC에 `C:\CLI\python`·opencode가 없고 스킬을 복사하지 않은 환경 차이.
- 집 PC(2026-10-31, Office 2007, ml-kit 2.9.10 / office-cli 1.4.9): beta check 27/27(F7 포함), pytest 통과, ml-kit 실기 14/14, office-cli 실기 PASS 69·ENV-SKIP 2(PDF 추가 기능 없음), UNVERIFIED 0건, 눈 확인 V1~V7·V9~V10 통과
- 회사 PC(2026-10-29~30): beta check 26/27(F2), Gemma 12/14(X1·X2에서 도구 코드 읽기 -> 지침 보강, X2 시나리오 변경)
- 회사 PC(2026-11-04, `run_all_tests.py --only gemma`, opencode CLI 2.2.5, ml-kit 2.9.39): Gemma 12/14. R1(보고서)·R2(성능예측표)가 opencode 명령 시간 제한 120초에 끊겨 결과 JSON 없음(R1은 같은 명령 4번) -> 2.9.40 나눠서 진행(`CONTINUE:`). 나머지 12건 통과, X2 답변 판정 통과(참고)

## 5. 남은 일
- 회사 PC 실기 결과 분석과 수정
- 각 도구 `CLI-dev/<tool>/CHANGELOG.md` 하단 "남은 할 일"
- 프로젝트 마무리 때 `docs/test_logs/` 등 로그 이력 정리 (저장소를 다시 만드는 방식)
- 도구 개발을 마친 뒤 다시 판단하기로 한 검토 (2026-11-01, 지금은 하지 않음):
  - MCP 전환: CLI 유지. 다시 본다면 CLI를 하위 프로세스로 부르는 얇은 MCP 층부터
  - 전용 에이전트: 사내 에이전트 En-Genius에 스킬로 붙이는 패키지를 만들었으나 업데이트를 멈추고 `archive/engenius`에 보관 (2026-11-04).
    다시 쓸 때 원래 자리로 옮기고 zip을 새로 만듦 (archive/README.md). 도구 쪽 En-Genius 대응(임시 폴더 실행 시 고정 데이터 폴더)은 그대로 둠
  - ml-kit 웹 서버: 동료 수요(예측이면 성능예측표로 충분)와 IT 정책·문서 보안 확인 후
  - `opencode run` 자동화: 사용 중. opencode CLI v2.2.5, 일반 권한 창 (관리자 창에서는 열린 문서를 못 봄), 기본 모델은 `%USERPROFILE%\.config\opencode\opencode.jsonc`의 `model`.
    `run_all_tests.py --gemma`가 시나리오 14개를 한 세션에 보내고 `grade_opencode.py --runs`로 채점 (2026-11-04 회사 PC 12/14, R1·R2는 2.9.40에서 수정).
    이후 office-cli 묶음 O(S1 샘플, O1~O5)를 더해 자동 19개 + 수동 전용 O3(Outlook 메일). 회사 PC 오류 로그의 지어낸 office-cli 옵션은 공통 규칙 `no_fake_office`

## 6. 붙여넣을 첫 메시지
클라우드용 (Claude Code 클라우드 세션):
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md(1장, 3-1장, 4장의 "집 PC opencode 모델 비교")를 읽어.
작업 브랜치는 claude/home-pc-check-2973 (ml-kit 2.9.74, office-cli 1.4.82). 근거 자료는 docs/test_logs/opencode_모델비교_20261107/.
1. 3-1장 1번: ml-kit 2.9.74 변경을 검토하고 ml-kit 전체 pytest와 CLI-dev/tests를 돌려 줘.
2. 3-1장 2~5번을 검토해 고칠 것은 회귀 테스트와 함께 고치고(도구 변경은 SKILL.md·README·CHANGELOG 같이), 판단이 필요한 것은 나에게 물어 줘.
3. 같은 브랜치에 커밋·푸시하고, 끝나면 PR을 만들어 줘 (머지는 내가 정함).
답변은 한국어로, 이모지 없이.
```

집 PC용 (Claude Code, 저장소 폴더에서):
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md(특히 1장과 3-0장), docs/HOME_PC_CHECK.md를 읽어.
지금은 집 PC(Windows, WinPython 3.13, Office 16(2019), Gemma 없음)야. 작업 브랜치는 ccr-ecab210c-awib23.
1. HANDOFF 3-0장 1~2번대로 준비(setup.bat 관리자 실행, PATH의 python이 C:\CLI\python 3.13인지)를 확인하고 3번(run_simplify_check.py --label 집PC)을 실행해 줘.
   명령은 python만 쓰고, 실행 전에 무엇을 할지 짧게 알려 줘.
2. 결과 파일의 PASS/FAIL/SKIP을 표로 정리하고, FAIL은 원인과 고칠 곳(파일:줄)을 찾아 줘.
3. 고칠 것은 회귀 테스트와 함께 고치고, 바꾼 도구의 테스트와 CLI-dev/tests를 돌린 뒤 같은 브랜치에 커밋·푸시해 줘.
4. 시간이 되면 3-0장 4번 전체 회귀 점검도 돌려 줘.
답변은 한국어로, 이모지 없이.
```

회사 PC용:
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md를 읽고 지금 상태를 파악해.
지금은 회사 PC(Windows 11, WinPython 3.13.13, Office 16, 문서 보안)에서 실기 테스트를 하려고 해.
1. HANDOFF.md 3장대로 준비를 같이 하고, 자동 점검 `python run_all_tests.py --label 회사PC`(C:\CLI-dev에서) 한 번을 실행해 줘. 명령은 python만 쓰고, 실행 전에 무엇을 할지 짧게 알려 줘.
2. 끝나면 결과 파일 result_회사PC_<시각>.md 맨 위의 요약 표와 실패·ENV-SKIP·UNVERIFIED 목록을 표로 정리해 줘. 원인 추정과 고칠 곳(파일:줄)도.
3. 고칠 것은 바로 고치지 말고 계획을 먼저 보여 줘.
답변은 한국어로, 이모지 없이.
```
