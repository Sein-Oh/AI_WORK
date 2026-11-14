# CHANGELOG - pandas-kit

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`.

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
- 다음 판 명령: `replay --file <새 판>`(같은 단계를 다른 파일에), 명시적 `fill`·`trim`. (`where-used`·`rollup` 같은 BOM 계산은 도구가 아니라 MBOM 스킬로)
- 긴 작업 나눠서 진행(`CONTINUE:`, R10): 첫 판은 한 번에 끝냄 (큰 파일 시간 측정 뒤 넣음).
- 범위 밖 주소(`C300`처럼 읽은 범위 밖)를 받으면 범위를 넓혀 다시 읽는 `[추천]`: 첫 판은 `--columns`의 열 글자가 표에 없으면 열 목록만 알림.
- 사용 가이드 HTML (`AIKIT/guide/`), 실기 러너·Gemma 시나리오.
- 자유 코드 `eval`(README에만, 대형 모델용).
