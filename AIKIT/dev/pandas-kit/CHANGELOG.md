# CHANGELOG - pandas-kit

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`.

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
- 다음 판 명령: `where-used`(이 부품이 쓰인 상위), `rollup`(하위 수량 곱해 올리기), `replay --file <새 판>`(같은 단계를 다른 파일에), 명시적 `fill`·`trim`.
- 긴 작업 나눠서 진행(`CONTINUE:`, R10): 첫 판은 한 번에 끝냄 (큰 파일 시간 측정 뒤 넣음).
- 범위 밖 주소(`C300`처럼 읽은 범위 밖)를 받으면 범위를 넓혀 다시 읽는 `[추천]`: 첫 판은 `--columns`의 열 글자가 표에 없으면 열 목록만 알림.
- 사용 가이드 HTML (`AIKIT/guide/`), 실기 러너·Gemma 시나리오.
- 자유 코드 `eval`(README에만, 대형 모델용).
