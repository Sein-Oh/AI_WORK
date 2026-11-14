# CHANGELOG - pandas-kit

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`.

## [0.3.0] - 2026-11-15

pandas에서 자주 쓰는 메서드를 명령으로 (사용자 결정 2026-11-15 "가": CLI지만 인터프리터처럼, pandas 코드를 짜게 하지 않는 원칙은 그대로).
새 명령은 모두 `--from`(꼭 씀)의 표로 새 표(`--name`)를 만들고, 작업 기록에 단계로 남아 스냅샷이 없어도 기록으로 다시 만든다.

### 기능 (1차)
- `describe --from` (`--columns`): 열마다 한 행. 숫자 열은 개수·평균·표준편차·최소·25%·50%·75%·최대, 그 밖의 열은 개수·종류 수·가장 많은 값과 그 개수.
  글자 열인데 칸이 대부분 숫자로 읽히면 `warnings` (앞자리 0이 있는 코드 열은 빼고).
- `unique --from --columns`: 열마다 서로 다른 값과 개수, `엑셀 행 목록`. 빈 칸은 `(빈 칸)`.
- `group`에 `--mean`, `--min`, `--max`(숫자 열), `--count`(빈 칸이 아닌 칸 수), `--nunique`(서로 다른 값의 수). `--sum`과 같은 모양(열 목록 `,`),
  단위 섞임 경고도 `--sum`과 같음.
- `calc --from --expr "금액=Qty*[Unit Price]"`: 식으로 계산한 열을 더함. 열 이름, 숫자, `+ - * /`, 괄호만 도구가 직접 읽음 (`eval` 없음).
  공백·기호가 든 열 이름은 `[열 이름]`, 여러 식은 `;`. 0으로 나눈 행·숫자가 아닌 칸·빈 칸은 빈 값 + `warnings`. 있는 열 이름으로는 만들지 않음.
- `--where`에 `^=`(로 시작), `$=`(로 끝).
- `top --from --by <숫자 열> --n <개수>`: 큰 순 앞 몇 행 (`--asc`면 작은 순, `--n` 생략은 10).
- 시험 `tests/test_pandas_methods.py`: 명령마다 정상, 0건, 틀린 열 `[후보]`, 구분자 `[추천]`, 스냅샷을 지운 뒤 기록으로 다시 만들기.

### 스킬
- 명령 표에 `describe`, `unique`, `calc`, `top`과 `group`의 새 옵션, 조건 연산자 `^=` `$=`.

## [0.2.12] - 2026-11-14

### 수정
- 표 가운데 빈 행이 있으면 `[후보]`로 범위를 묻던 것: 묻지 않고 빈 행까지 한 표로 읽고 `warnings` 첫 줄로 알림
  (사용자 결정 2026-11-14, bom-kit 0.2.1과 같은 규칙: 표 가운데 빈 행은 표가 끝난 것이 아님). 일부만 읽으려면 `--range`.
  표 아래 작은 메모처럼 열이 다르거나 작은 덩어리는 전처럼 "빈 행 아래에 내용이 있는 행"으로 알리고 읽지 않음.
  시험 `test_a_blank_row_inside_the_table_is_read_as_one_table_without_asking`.

## [0.2.11] - 2026-11-14

### 수정
- 두 실행이 같은 분에 같은 원본을 처음 읽으면 같은 작업 폴더 이름을 골라 한쪽이 FileExistsError로 멈추던 것 (bom-kit 동시 실행 시험에서 드러남):
  폴더를 만들 때 이미 있으면 다음 번호(_2, _3)로.

## [0.2.10] - 2026-11-14

### 수정 (집 PC 실기 2026-11-14, `docs/test_logs/집PC_bom_20261114/결과.md`)
- `--file`·`--from` 없이 `--where`(또는 `--columns`)만 준 `select`: 마지막에 만든 표를 `[추천]`하던 것 (P7: 마지막 표가 0행 `df4`라
  그대로 실행하면 0건). 표가 여럿이면 짐작하지 않고 표마다 `[후보]`(`표 df1(110행)에서 고르기`, 최근 6개)와 열린 통합문서를 함께 묻는다.
  모두 합쳐 하나일 때만 `[추천]`. 단계 기록(`steps.json`)에 표의 행 수(`rows`)를 남김.
- 같은 작업 폴더에 두 명령을 동시에 실행하면 한쪽 단계가 기록에서 사라지고, 기록과 스냅샷이 뒤바뀐 표를 `VERIFIED`로 보여 주던 것
  (Fable 검토 2026-11-14 A1, 재현함): 기록을 고치는 `add`·`undo`를 작업 폴더의 잠금 파일(`.lock`, 20초 기다림, 2분 지난 잠금은 버림)로
  한 실행씩 하고, 잠근 뒤 기록을 다시 읽어 단계 번호를 정한다. 스냅샷에 단계 표시(번호·이름·계보·시각)를 넣어 기록과 다르면 버리고
  기록으로 다시 만든다(`recovered`). `VERIFIED`는 저장한 기록을 다시 읽어 그 단계가 있을 때만. 시험 `test_two_runs_at_once_keep_both_steps_and_the_right_snapshots`,
  `test_a_snapshot_that_does_not_match_the_record_is_rebuilt`.
- `--name`이 다른 원본의 표를 가리게 되면 늘 경고 (Fable 검토 2026-11-14 A3, 사용자 결정): 전에는 그 이름으로 만든 표가 있을 때만
  "옛 df1 기준" 경고가 나와, 새 대화에서 다른 파일을 `--name df1`로 읽으면 어제 df1이 조용히 바뀌었다. 시험 `test_a_name_that_now_points_to_another_file_is_warned`.
- 원본이 바뀌었다는 경고를 쓰는 표의 원본만, 값이 다를 때만 (Fable 검토 A4, 사용자 결정): 전에는 작업 안의 모든 원본을 매번 확인해
  diff 뒤 원본에 칠하면(저장됨) 그 뒤 명령마다 "다시 읽으려면: python ..." 경고가 붙었다. 이제 그 명령이 읽는 표의 조상 원본만 보고,
  바뀌었으면 다시 읽어 값이 같으면 기록만 조용히 갱신(stderr `[INFO]`), 다르면 "(값이 다름)" 경고와 다시 읽는 전체 명령.
  시험 `test_source_saved_with_only_formatting_is_quiet_and_only_used_tables_are_checked`.
- `--from`을 빼면 마지막 표(지금 표)를 쓰던 것을 없앰 (사용자 결정 2026-11-14: 짐작을 엄격하게 막고 늘 `--from` 또는 `--file`을 쓰게 함).
  `show`·`group`·`sort`·`chart`·`export`에서 `--from`이 빠지면 `STOP_RETRY:`로 멈추고, 표가 하나면 `[추천]`, 여럿이면 `[후보]`(최근 6개, 행 수).
  `undo` 결과에서 `current`(지금 표)를 빼고(남은 표는 원래 있던 `tables`), preview는 "남은 표: df1, df2". 스킬 문장도 "늘 `--from <표 이름>`을 쓴다"로. 시험 `test_from_is_always_needed_and_never_guessed`.
- `select --file`의 결과에 `hidden_rows`(읽은 숨긴 행의 엑셀 행, 있을 때만): bom-kit이 결과에 든 숨긴 행을 알릴 때 씀.
- 표 위 내용이 한 행이면 `엑셀 행 1~1` 대신 `엑셀 행 1`.
- 틀린 명령 이름의 message에서 argparse 영어 문장을 빼고 한국어만.
- 시험: `test_select_without_file_or_from_stops_and_gives_full_commands`, `test_hidden_rows_merged_cells_and_uncomputed_formulas_are_warned`,
  `test_title_block_above_is_not_taken_as_the_table`, `test_misspelt_command_name_gives_candidates` (모두 고침).

## [0.2.9] - 2026-11-14

### 수정 (저장소 점검에서 직접 써 보며)
- `--file`도 `--from`도 없는 `select`: 고를 조건(`--where`)·열(`--columns`)이 없으면 '지금 표에서 고르기'를 권하지 않는다.
  전에는 작업 폴더가 남아 있고 Excel이 닫혀 있으면 `[추천] select --from <지금 표>`가 나와, 경량 모델이 그것을 실행하면 표를 그대로
  복사하고 "열린 엑셀을 읽었다"고 답할 수 있었다 (집 PC 2026-11-13 4-2와 같은 뿌리). 이제 열린 통합문서가 없으면 파일 경로를 묻게 하고,
  하나면 그것을 읽는 `[추천]`. 조건·열이 있을 때는 전처럼 지금 표에서 고르는 명령을 준다.
- 숫자 열에 숫자가 아닌 값으로 크기를 비교(`Qty>abc`)하면 글자 비교가 되어 말없이 0건이던 것: `STOP_RETRY:`로 멈추고 숫자로 쓰라고 알림
  (날짜 열과 같은 처리). 같음·다름(`Qty=abc`)은 그대로 비교하고 `warnings`로 알림.
- 틀린 명령 이름(`selct`)에 `[후보]`와 고친 전체 명령(`confirm_commands`). 철자만 비슷한 것이라 `[추천]`이 아니라 사용자에게 물음.
- 시험: `test_select_without_file_or_from_stops_and_gives_full_commands`(고침), `test_number_column_with_a_value_that_is_not_a_number`,
  `test_misspelt_command_name_gives_candidates`.

## [0.2.8] - 2026-11-14

### 스킬
- `where-used` 기본 스킬 추가 (`AIKIT/.opencode/skills/where-used`, setup 기본 스킬 15개). BOM 한 판에서 품목이 쓰인 상위 조립체와
  제품 1대당 수량(상위 수량을 곱한 값)을 찾는다. 품번이 같거나 품명에 든 글자로 찾고, 품번 여럿이 걸리면 품번·단위별로 따로 센다.
  전에는 mbom-diff의 `mbom_path.py`가 두 판을 요구해 같은 표를 두 이름으로 export해야 했다 (개선점_MBOM.md K).
  시험 `test_where_used_skill.py` (쓰인 곳과 수량, 경로를 따라 곱하기, 품명으로 찾기, 0건·열 못 찾음, SKILL.md 크기).

## [0.2.7] - 2026-11-14

### 수정 (체결 부품을 상위까지 정리하며, 개선점_MBOM.md K)
- 숨긴 행(폐기 품목)이 집계에 들어가는데 어느 행인지 알 수 없어 뺄 방법이 없었음 -> 경고에 엑셀 행 번호와 그 행을 뺀 전체 명령.
- `--where`가 `엑셀 행`을 받지 않았음(`join --on`은 받음) -> 받는다. 숨긴 행·메모 줄을 행 번호로 뺄 수 있음.
- `group --sum`의 합계가 정수를 더해도 소수로 나왔음(`33.0`) -> 더한 값이 모두 정수면 정수로.
- 시험 `test_hidden_rows_can_be_left_out_and_group_totals_stay_whole`.

## [0.2.6] - 2026-11-14

### 수정 (사용자가 고친 MBOM 두 판을 비교하다가, 개선점_MBOM.md J)
- 표 가운데 빈 행(사용자가 구역을 나누려고 넣은 18·49행) 때문에 표가 셋으로 갈려 가장 큰 덩어리(69행)만 읽고 그 첫 줄을 머리글로 삼았음.
  BOM 111행 중 69행만 비교할 뻔했고 경고만 있었음 -> 열이 거의 같은 큰 덩어리가 더 있으면 `[후보]`로 묻는다
  (1. 빈 행까지 한 표로, 2~4. 덩어리마다 하나씩, 각각 `--range`를 넣은 전체 명령). 표 아래 메모처럼 작거나 열이 다른 덩어리는 지금처럼 `warnings`만.
  시험 `test_a_blank_row_inside_the_table_asks_which_range`.
- 0.2.5의 "기호 붙은 숫자" 경고가 `P100`, `MC-03` 같은 코드 열까지 잡았음 -> 앞의 기호는 통화(`₩ $ € ¥ £`)만.
  시험 `test_a_code_like_p100_is_not_called_a_number_with_a_symbol`.

## [0.2.5] - 2026-11-14

### 수정 (회사 엑셀에 흔한 모양으로 써 보며, 개선점_MBOM.md I)
- 머리글이 두 줄인 표(위는 병합한 묶음 이름 `기본`·`금액`, 아래가 `업체`·`품번`)에서 위 줄을 머리글로 잡아 열 이름이 `기본`, `열_B`가 되고
  진짜 머리글이 데이터 첫 행이 됐음 -> 아래 줄을 열 이름으로, `warnings`로 알림. 시험 `test_two_row_header_uses_the_lower_row_and_keeps_every_column`.
- 머리글 줄에서 값이 찬 칸의 처음~마지막까지만 열로 잡아, 묶음 이름이 없는 끝 열(`입고일`)이 통째로 빠졌음 (조용히 데이터가 사라짐)
  -> 머리글이 비어도 값이 여러 행에 있으면 그 열도 읽고 `warnings`. 한 행에만 있는 칸은 열로 보지 않음.
  시험 `test_a_column_whose_header_is_empty_is_read_but_a_stray_cell_is_not`.
- 천 단위 쉼표(`1,200`)·글자 날짜(`2026.01.05`)가 글자 열이어서 합계·날짜 조건을 쓸 수 없었고 경고도 없었음 -> 숫자·날짜로 읽고 `warnings`
  (뜻이 둘인 `05/01/2026`은 글자 그대로). 기호가 붙은 숫자(`₩1,500`)는 글자 열로 두되 `warnings`를 내고, 조건·`sort`는 기호를 뺀 숫자로 비교.
  시험 `test_thousands_separator_currency_and_text_dates`.

## [0.2.4] - 2026-11-14

### 수정 (BOM이 아닌 표로 써 보며, 개선점_MBOM.md H)
- 머리글 아래 단위 행(`°C`, `bar`, `wt%` ...)이 데이터로 읽혀 숫자 열이 모두 글자 열이 됐고, 조건으로 그 행을 빼도 글자 열 그대로여서
  `group --sum`·`chart --y`를 아예 쓸 수 없었음 (ml-kit 공정 샘플: 9열 전부) -> 단위 행으로 보고 빼며 `warnings`, 단위는 `columns`의 `unit`.
  숫자 열이 둘 이상 살아날 때만 단위 행으로 본다 (보통 글자 열은 데이터 그대로). 시험 `test_unit_row_under_the_header_is_read_as_column_units`,
  `test_a_text_row_that_is_not_units_is_kept_as_data`.
- `sort`: 글자 열을 글자 순서로 정렬해 유량 5.02가 10.02보다 크다고 했고 알리지도 않았음 (`--where`는 숫자로 비교하는데 `sort`만 달랐음)
  -> 칸이 대부분 숫자인 글자 열은 숫자로 정렬하고 `warnings`, 숫자가 아닌 칸은 맨 뒤. 시험 `test_sort_orders_a_text_number_column_as_numbers`.
- 열마다 나오던 같은 읽기 경고(공정 샘플 9줄)를 분류별 한 줄로 묶음. 시험 `test_same_read_warning_for_many_columns_is_one_line`.

## [0.2.3] - 2026-11-14

### 사용자 결정 반영 (개선점_MBOM.md F)
- `select`: `--file`·`--from` 둘 다 없으면 멈춤 (전에는 만든 표가 있으면 지금 표, 없으면 열린 Excel 통합문서를 말없이 골랐음).
  지금 표에서 고르는 명령과 열린(저장된) 통합문서를 읽는 명령을 `[후보]`로, 하나면 `[추천]`. 시험 `test_select_without_file_or_from_stops_and_gives_full_commands`.
- `chart --kind line`: 같은 `--x` 값이 여러 행이면 합계하지 않고 `warnings` + 막대로 그리는 전체 명령. 시험 `test_chart_sums_repeated_x_limits_bars_and_checks_columns`.
- `group`: 빈 값 묶음 이름이 `null` -> `(빈 칸)`. 시험 `test_group_names_the_empty_value_group`.
- `--work`: 지금 작업 폴더를 바꾸면(보기 명령 `show`도) preview 앞에 "지금 작업 폴더를 <이름>(으)로 바꿨습니다"를 붙임 (전에는 말없이 바뀌어
  다음 명령이 다른 폴더의 같은 이름 표를 썼음). 시험 `test_work_option_says_when_the_current_work_folder_changes`.
- `join --keep-left`: 왼쪽 표의 행을 모두 남김 (짝이 없으면 오른쪽 열은 빈 칸, 그 행 수 `warnings`). 짝이 없어 빠진 행이 있는 보통 `join`의
  경고에 `--keep-left`를 붙인 전체 명령. 시험 `test_mbom_join_keep_left_keeps_every_left_row`.
- 그대로 두기로 한 것: 글자 열의 `=`는 글자 그대로 비교(0.2.1의 경고 유지), `group --sum`은 비율 열도 더함.

### 스킬
- `mbom-diff`를 기본 스킬로 (`AIKIT/.opencode/skills/mbom-diff`, setup 기본 스킬 14개). 다른 상위로 옮겨진 품목은 `[이동]`이 아니라
  삭제·추가로, 대체 후보는 품번만 보고 같은 자리에서 (옮겨진 품번은 짝짓지 않음).

## [0.2.2] - 2026-11-14

### 수정 (두 판 비교 스킬로 v2 -> v2_mod를 찾아 보며, 개선점_MBOM.md G)
- `join --on "엑셀 행=엑셀 행번호"`: 결과가 엑셀 행을 `<표> 행`으로 바꿔 그 표로 `diff`하면 엑셀 행과 칠하기 명령이 말없이 빠졌음
  -> 왼쪽 키가 엑셀 행 하나면 결과는 원래 표의 엑셀 행·열 글자를 그대로 가짐. 시험 `test_mbom_join_on_excel_row_keeps_the_original_excel_rows`.
- `diff` 칠하기 명령의 설명이 삭제(추가)된 행과 바뀐 칸을 합해 "바뀐 칸 11곳"이라 적었음 -> "삭제된 3행, 바뀐 칸 8곳".

### 시험 자료
- `make_mbom_samples.py`: `MBOM_DR200_v2_mod.xlsx` (v2를 회사 BOM 개정처럼 고친 판, 정답 `MOD_CHANGES`, 함정 `MOD_TRAPS`).
- 조합 스킬 예 `tools/skills/mbom-diff/`(0.2.3에서 기본 스킬로) (SKILL.md, `scripts/mbom_path.py`, `scripts/mbom_summary.py`): 경로로 두 판을 맞춰
  이동·대체 후보·추가·삭제·바뀐 칸을 알림. 업무 판단이라 도구가 아니라 스킬. 시험 `test_mbom_diff_skill.py`(정답과 같은지, 함정이 안 나오는지).

## [0.2.1] - 2026-11-14

### 수정 (MBOM 두 판으로 기반 명령을 써 보며 찾은 '결과가 틀리는데 알리지 않음', 개선점_MBOM.md E)
- 글자 열에 `=`로 숫자를 주면(`Part Number=125`, `Op No=10`) 글자 그대로 비교해 0건이었는데 말이 없었음 -> 숫자로 같은 값이 다르게 적힌 칸 수와
  적힌 값(`000125`, `0010`)을 `warnings`로, 적힌 그대로 쓰는 예. 비교 방식은 그대로 (값을 바꾸지 않음).
- 날짜 열 조건 값이 연-월-일이 아니면(`2026-13`, `2026-12`, 없는 날 `2026-02-30`) 글자 비교로 그럴듯한 결과를 냈음 -> `STOP_RETRY` (연-월-일로).
- `join`(양쪽에 있는 행만): 짝이 없어 빠진 행 수를 알리지 않았음 (MBOM과 Routing: 110행 중 104행) -> 양쪽 빠진 행 수 `warnings`.
- `join`의 오른쪽 키 열 `[후보]`: 철자 순이라 맞는 열(`Item`)이 둘째였음 -> 왼쪽 키와 겹치는 값이 많은 열부터, 후보마다 겹치는 값 수.
- `only`: 키가 행마다 다르지 않으면(MBOM `Part Number` 34행 중복) 지운 행의 키가 다른 행에 남아 있을 때 결과에서 빠지는데 알리지 않았음 -> `warnings`.
- `diff`: 키가 같은 행이 있어 멈출 때 어느 키·행인지 몰랐음 -> 예 3개(`000125 (엑셀 행 16, 30, 35, 38, 58 등)`). BOM 용어(상위 품번) 안내는 뺌 (업무 판단은 스킬).
- 표 위의 다른 내용 경고가 한 칸짜리 제목 줄을 세지 않아 "3행(엑셀 행 2~4)"이었음 -> "4행(엑셀 행 1~4)".
- `export` xlsx: 날짜만 있는 열이 엑셀에서 `2026-01-01 00:00:00`으로 보였음 -> 날짜 형식 `yyyy-mm-dd` (시각이 있으면 시각까지).
- 시험: `test_mbom_title_rows_counted_and_number_written_differently_is_warned`, `test_mbom_date_condition_must_be_a_full_date`,
  `test_mbom_join_says_unmatched_rows_and_ranks_key_candidates_by_shared_values`, `test_mbom_only_and_diff_with_repeated_keys`,
  `test_mbom_export_keeps_dates_as_dates`.

## [0.2.0] - 2026-11-13

### 시험 자료
- `tools/make_mbom_samples.py`: 집 PC에서 쓴 MBOM 두 판을 클라우드에서도 같게 만듦 (xlsx는 반입 규칙상 저장소에 넣지 않음). 시험 `test_mbom_samples_are_made_and_read`.

### 구조
- 성격을 완결형 작업 도구에서 기반 도구로 바꿈 (사용자 결정 2026-11-13, office-cli와 같은 개념): 표를 고르고·보고·비교하고·잇고·저장하는
  단위 기능만 둔다. BOM 분석(Level로 상위 찾기, 누적 수량, 총소요량) 같은 업무 판단은 이 명령들을 조합한 스킬이 한다.
  MBOM으로 써 보며 넣었던 `tree`·`rollup`과 `diff`의 BOM 안내는 같은 이유로 넣지 않음 (MBOM 스킬로).

### 기능 (MBOM_DR200.xlsx로 직접 써 보며 찾은 기반 기능 부족)
- 표 찾기: 빈 행으로 나뉜 내용 중 가장 큰 것을 표로 (제목 블록 `공장·Rev·작성일`을 표로 잡던 것), 표 위의 다른 내용은 `warnings`.
- `group --by --sum`: 값별 행 수·합계·엑셀 행 목록, 단위 섞임 `warnings`. `sort --by --desc`: 정렬(같은 값은 표 순서, 빈 칸은 끝).
- `show --start`: 21행부터 보기, 더 있으면 다음 20행 `next_commands`.
- 작업 폴더 이름: `pandas_YYMMDDHHMM` -> `<파일 이름>_YYMMDDHHMM`, 같은 분이면 `_2` (사용자 결정, 어느 파일의 작업인지 이름으로 보임).
- `--on "왼쪽 열=오른쪽 열"`: 이름이 다른 키(`Part Number=Item`), `엑셀 행` 키로 스킬이 계산한 csv를 원래 표에 붙이기.
  머리글 `엑셀 행`처럼 도구의 행 번호 열 이름은 `엑셀 행번호`로 바꿔 읽음. 키의 정수는 종류와 상관없이 같게 봄.
- 시험: `test_title_block_above_is_not_taken_as_the_table`, `test_join_with_different_key_names_and_back_by_excel_row`,
  `test_group_counts_sums_and_warns_mixed_units`, `test_sort_keeps_excel_rows_and_puts_empty_cells_last`, show `--start` (기존 시험에 더함).

### 수정
- `diff`: 한 판에만 소수 수량이 있으면 열 종류가 달라(정수/소수) 2와 2.0을 바뀐 칸으로 셌음 (numpy 정수를 숫자로 보지 않던 것).
  시험 `test_diff_same_number_in_int_and_decimal_columns_is_not_a_change`.

## [0.1.2] - 2026-11-13

### 수정 (MBOM_DR200.xlsx로 직접 써 보며 찾은 '결과가 틀리는데 알리지 않음' 5가지)
- 숫자 사이 글자로 저장된 숫자 한 칸이 열 전체를 글자로 만들던 것 -> 숫자 열로 읽고 `warnings`(칸 수, 엑셀 행). 숫자 사이 숫자가 아닌 칸이 조금 섞인 글자 열도 `warnings`.
- 글자 열에 숫자 값으로 `>` `<` `>=` `<=`: 글자 순서 비교("18" > "3"이 거짓) -> 숫자로 읽히는 칸끼리 숫자 비교, 숫자가 아닌 칸 수 `warnings`.
- `!=`·`!~` 갈래만 `|`로 이으면 "어느 것도 아님" (`Item Type!=RAW|CONS`가 모든 행을 남기던 것).
- 값 없는 조건(`Qty>=`, `Desc~`)이 모든 행을 내던 것 -> 거절 (빈 칸은 `열=`·`열!=`만).
- `chart`: 단위 열(`UoM` 등)에 단위가 여러 가지면 `warnings` + 단위 하나로 고르는 전체 명령.
- 시험: `test_numbers_written_as_text_and_numeric_conditions_on_text_columns`, `test_not_equal_with_choices_means_none_of_them_and_empty_value_needs_equal`, `test_chart_warns_when_units_are_mixed`.

## [0.1.1] - 2026-11-13

### 수정
- 집 PC 실기(2026-11-13): `join`·`only`의 preview·계보에 받침과 맞지 않는 조사("df1와 df2를")가 나옴 -> "df1, df2을(를) <키> 기준으로 짝 맞춤", "(<키> 기준으로 df1과(와) 비교)". 시험 `test_join_only_stack_keep_both_row_numbers`.
- 시험: `test_diff_adds_deletes_changed_cells_with_rows_and_paint_files`가 명령을 `split()`으로 나눠 Windows 경로의 따옴표가 남아 실패 -> `shlex.split`.

## [0.1.0] - 2026-11-13

### 구조
- 새 도구 (완결형 작업 도구, 설계 `설계.md`, 결정 `docs/DECISIONS.md` 2026-11-13). `AIKIT/pandas-kit/pandas_kit.py`(진입점),
  `_pd_runtime.py`(ml-kit `_ml_runtime.py`에서 필요한 부분만 같은 이름·같은 동작으로 복사: `run_cli`, `JsonArgumentParser`, `quote_arg`,
  오류 로그, `retry_fs`, `save_with_lock_fallback`, `read_text_file`). 기본 스킬 `pandas-kit` (setup 기본 스킬 13개, `.gitignore`, 지침).

### 기능
- 첫 판 명령: `select`, `show`, `join`, `diff`, `only`, `stack`, `undo`, `new`, `export`, `chart` (README.md).
- 표 찾기(제목 한 줄 건너뜀, 빈 행에서 끝), 엑셀 행 번호·열 글자 유지, 값은 바꾸지 않음(앞자리 0 코드는 글자), 병합 칸·숨긴 행·저장 안 된 수식 `warnings`.
- 상태는 작업 폴더의 파일: 단계 기록(원본) + pickle 스냅샷(캐시, 깨지면 기록으로 다시 만듦), 원본 변경 감지, 옛 표 기준 경고.
- `diff`: 바뀐 곳마다 한 행 + 원본 파일별 칠할 칸 목록과 office-cli 칠하기 명령(`confirm_commands`).
- `chart`: 표를 차트 그림(.png)으로 (막대·선·산점·원·히스토그램·상자, 사용자 요청 2026-11-13). 같은 x는 합계, 막대 30개·원 12조각까지, 한글 글꼴 찾기, 다시 열어 확인.

## 남은 할 일
- **pandas 메서드 명령 늘리기** (사용자 결정 2026-11-15, "가": pandas에서 자주 쓰는 메서드를 명령 하나씩, CLI지만 인터프리터처럼.
  pandas 코드를 짜게 하지 않는 원칙은 그대로, 새 명령 이름은 pandas 이름 그대로, 있는 `group`·`stack` 등은 그대로).
  모든 명령은 `--from` 필수, 결과는 새 표(`--name`), 엑셀 행 유지, 한국어 preview, 다시 읽어 VERIFIED, 구분자 규칙(`,` 이름 목록, `;` 조건·식, `|` 또는).
  - 2차: `dropna --columns`, `fillna --columns --value`(위 값으로 `--down`, 병합 칸 정리), `drop-duplicates --columns`(`--keep first` 기본), `rename --columns "Qty=수량"`
  - 3차: `pivot --rows --cols --sum`(값 옵션은 group과 같음), `calc`의 `round`·`abs`
  - 검토(사용자와 정할 것): `~~`(대소문자 무시 포함), `Qty=1..10`(사이 값, 지금도 `Qty>=1;Qty<=10`), `cumsum`
  - 스킬이 8KB를 넘으면 작업별로 나눔 (나누는 방식은 사용자와 정함, 예: 읽기·고르기·비교 / 묶기·연산·정리)
- 다음 판 명령: `replay --file <새 판>`(같은 단계를 다른 파일에), 명시적 `fill`·`trim`. (`where-used`·`rollup` 같은 BOM 계산은 도구가 아니라 MBOM 스킬로)
- 긴 작업 나눠서 진행(`CONTINUE:`, R10): 첫 판은 한 번에 끝냄 (큰 파일 시간 측정 뒤 넣음).
- 범위 밖 주소(`C300`처럼 읽은 범위 밖)를 받으면 범위를 넓혀 다시 읽는 `[추천]`: 첫 판은 `--columns`의 열 글자가 표에 없으면 열 목록만 알림.
- 사용 가이드 HTML (`AIKIT/guide/`), 실기 러너·Gemma 시나리오.
- 자유 코드 `eval`(README에만, 대형 모델용).
