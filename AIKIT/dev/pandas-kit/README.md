# pandas-kit (0.1.0, 첫 판)

엑셀·CSV 표를 단계별로 고르고 비교하는 완결형 작업 도구입니다 (첫 대상: BOM 분석). 원본 파일은 쓰지 않고, 단계마다 이름 붙인 표(`df1`, `df2`)를
작업 폴더에 쌓습니다. 결과는 원본의 엑셀 행 번호와 열 글자를 함께 가져서, 원본의 어느 칸인지 잃지 않습니다.
읽기·칠하기·메모(원본 Excel에 쓰는 일)는 office-cli가 하고, pandas-kit은 그 명령을 전체 명령 문자열로만 줍니다 (도구끼리 import하지 않음).

- 에이전트용 사용법: `AIKIT/.opencode/skills/pandas-kit/SKILL.md` (8KB 이하)
- 공통 계약: [`../README.md`](../README.md) (stdout 1줄 JSON, `VERIFIED`/`UNVERIFIED`, `[추천]`/`[후보]`, 오류 로그)
- 남은 결정: [`설계.md`](설계.md)

## 파일
| 파일 | 하는 일 |
|---|---|
| `AIKIT/pandas-kit/pandas_kit.py` | 진입점. `python pandas_kit.py <명령> ...` |
| `AIKIT/pandas-kit/_pd_runtime.py` | 공통 런타임 (진입점 아님): `run_cli`, `JsonArgumentParser`, `quote_arg`, 오류 로그, 파일 복구 함수, 열린 Excel 경로 찾기. ml-kit `_ml_runtime.py`와 같은 이름·같은 동작 |
| `tests/` | `python -m pytest -q AIKIT/dev/pandas-kit/tests` |

필요한 라이브러리: pandas, openpyxl (xlsx), pywin32 (Windows에서 열린 Excel 경로 찾기만, 없어도 됨). `requirements.txt`.

## 명령 (첫 판)
| 명령 | 옵션 | 결과 |
|---|---|---|
| `select` | `--file`, `--sheet`, `--range`, `--from`, `--columns`, `--where`, `--name` | 파일(시트·범위)이나 만든 표에서 행·열을 골라 새 표 |
| `show` | `--from`, `--columns` | 앞 20행 (`엑셀 행` 포함). 넘으면 `truncated: true` + `full_result_file`(작업 폴더 `results/<표>_전체.csv`) |
| `join` | `--left`, `--right`, `--on`, `--name` | 키 열로 짝 맞춰 붙이기 (양쪽에 있는 행만). 오른쪽의 같은 이름 열은 `<열>_<오른쪽 표>` |
| `only` | `--left`, `--right`, `--on`, `--name` | 왼쪽 표에만 있는 행 |
| `diff` | `--old`, `--new`, `--on`, `--name` | 바뀐 곳마다 한 행: `구분`(추가/삭제/바뀜), 키 열, `열`, `옛 값`, `새 값`, `<옛 표> 행`, `<새 표> 행` |
| `stack` | `--tables "df1,df2"`, `--name` | 위아래로 잇기. 맨 앞 `표` 열에 어느 표의 행인지 |
| `undo` | | 마지막 단계 하나 되돌리기 (그 이름은 그 전 표를 다시 가리킴) |
| `new` | | 새 작업 시작 (다음 `select --file`이 새 작업 폴더를 만듦) |
| `export` | `--from`, `--save`, `--overwrite` | 새 파일(.xlsx, .csv). 생략하면 작업 폴더의 `<표>.xlsx`, 이름만 주면 작업 폴더에 |

모든 명령은 `--work "<작업 폴더>"`를 받습니다 (지난 작업으로 돌아가기, `new` 제외). `--name`을 빼면 `df1`, `df2`, ... 중 안 쓴 첫 이름,
`--from`을 빼면 지금 표(마지막 단계의 표)입니다.

## 대상 파일과 표 찾기
- `.xlsx`, `.xlsm`, `.csv`. `--file`이 없고 만든 표도 없으면 열린 Excel 통합문서의 경로를 COM으로 읽습니다 (경로만, 문서는 건드리지 않음).
  하나면 그 파일, 여럿이면 `[후보]`와 같은 번호의 `confirm_commands`, 없으면 경로를 물으라고 안내합니다.
- 저장 안 된 통합문서(`Saved=False`, 저장한 적 없는 새 문서 포함)는 읽지 않습니다: `STOP_RETRY` "Excel에서 저장한 뒤 같은 명령을 다시 실행".
- 시트: `--sheet`가 없으면 파일을 열었을 때 보이는 시트. 없는 이름은 `[후보]`(시트 목록).
- 표 찾기: 칸이 둘 이상 찬 첫 행이 머리글 (위의 한 칸짜리 제목 행은 건너뜀), 머리글의 처음~마지막 찬 칸이 열, 완전히 빈 행에서 끝.
  빈 행 아래에 내용이 더 있으면 `warnings`. `--range B5:H230`(또는 `B:H`, 왼쪽 위 칸 `B5`)를 주면 그 안의 첫 내용 행이 머리글이고,
  끝을 준 범위는 끝까지 읽으며 범위 안의 빈 행은 뺍니다(`warnings`). 찾은 범위는 preview에 (`B4:G8 (머리글 4행)`).
- 열 이름: 머리글 글자 그대로 (앞뒤 공백만 뺌). 빈 머리글은 `열_F`, 같은 이름은 `Qty_2`.
- csv 인코딩은 자동 판별 (UTF-8/BOM/UTF-16/CP949), 구분자는 `,` `;` 탭 중에서 찾습니다.

## 값은 바꾸지 않음
- 숫자 열: 열 전체가 숫자이거나, 전체가 숫자 모양 글자이고 앞자리 0이 없을 때만 (`Int64`/`float64`). 그 밖에는 글자 그대로 (`0102`는 `0102`).
  날짜 열은 전체가 날짜일 때만. 열마다 읽은 종류(`숫자`, `글자`, `날짜`, `빈 열`)를 `columns`에 줍니다.
- 공백·대소문자를 고치지 않고, 병합 칸의 빈자리도 채우지 않습니다. 병합 칸, 숨긴 행(자동 필터로 가려진 행 포함), 값이 저장되지 않은
  수식 칸(빈 값으로 읽음)은 개수를 `warnings`로 (30MB를 넘는 파일은 읽기 전용 모드라 이 셋은 확인하지 않고 그렇다고 알림).

## 열과 조건
- `--columns "Part Number,Qty"`: 머리글 이름(대소문자·띄어쓰기는 달라도 하나로 정해지면 그 열), 또는 원본 열 글자(`D`). `;`로 나누면 `[추천]`.
- `--where "Level=1;Qty>2"`: `;`는 그리고, `|`는 또는. `"Maker=A|B"`처럼 값만 쓴 갈래는 앞 갈래의 열·연산자를 씁니다.
  연산자 `=` `!=` `>` `<` `>=` `<=` `~`(포함) `!~`(포함 안 함), 빈 칸은 `"비고="`. 숫자 열은 숫자로, 날짜 열은 날짜로, 그 밖은 글자 그대로
  (대소문자 구분) 비교합니다. 조건을 `,`로 나누면 `;`로 고친 전체 명령 `[추천]`. 조건은 `--columns`보다 먼저 적용합니다 (남기지 않을 열로도 거름).
- 없는 열 이름은 철자가 비슷한 열을 `[후보]`(같은 번호의 `confirm_commands`)로 묻습니다.

## 작업 폴더 (상태는 파일로)
- 처음 `select --file`에서 그 파일 옆에 `pandas_YYMMDDHHMM\`(같은 분이면 `_2`)을 만들고, 그 뒤 명령은 "지금 작업 폴더"에 이어서 씁니다
  (다른 파일을 읽어도 같은 폴더). 지금 작업 폴더 표시는 `<도구 데이터 폴더>/current.json` (`PANDAS_KIT_DATA_DIR`, 기본 `C:\AIKIT\pandas-kit\data`).
- 작업 폴더: `.created_by_pandas-kit`(표시), `steps.json`(단계 기록, 원본), `tables/s<단계>.pkl`(표 스냅샷, 캐시), `results/`(전체 결과 csv, 칠할 칸 목록), 내보낸 파일.
- 스냅샷이 없거나 깨졌거나 다른 pandas 버전이 만들었으면 버리고 기록을 다시 실행해 만듭니다 (`recovered`). 표시 파일이 있는 자기 작업 폴더의 pickle만 읽습니다.
- 표마다 계보(`lineage`, 예: `df2 = df1에서 조건 Maker=A|B`)를 기록하고 preview에 한 줄로 줍니다.
- 같은 이름을 다시 만들면 새 단계이고 이름은 새 표를 가리킵니다. 옛 표에서 만든 표는 다시 계산하지 않고 `warnings`("df2는 옛 df1 기준").
- 원본 변경 감지: 읽을 때의 수정 시각·크기를 기록해 두고, 다음 명령에서 달라졌으면 `warnings`에 다시 읽는 전체 명령을 넣습니다.

## 결과 JSON (공통 키 외)
- `select`·`join`·`only`·`diff`·`stack`: `table`, `count`, `columns`(`name`, `letter`, `kind`, 30개까지), `location`, `lineage`, `work`,
  `next_commands`(그 표 `show`). 파일에서 읽으면 `anchor`(`file`, `sheet`, `header_row`, `range`)와 `source`.
  0행이면 `VERIFIED: 0건 ...`과 `next_commands` 없음.
- `show`: `rows`(`엑셀 행` 또는 `<표> 행` 포함), `shown`, `truncated`, `full_result_file`.
- `diff`: `summary`(추가·삭제·바뀐 칸 수), `paint_files`(원본 파일별 칠할 칸 목록 csv, 줄마다 `F7` 또는 행 범위 `B6:G6`),
  `confirm_commands`(그 목록으로 원본을 칠하는 office-cli `excel.py style --target ... --file ... --bg yellow`, 사용자가 원할 때만).
  `--on`이 같은 행이 있으면 짝을 정할 수 없어 `STOP_RETRY`.
- `export`: `output_file` (다시 읽어 행 수 확인 뒤 `VERIFIED`). 같은 이름이 있으면 시각을 붙인 이름(`recovered`), 덮어쓰기는 `--overwrite`일 때만.
  원본 파일 경로에는 쓰지 않습니다.

## 시험
`python -m pytest -q AIKIT/dev/pandas-kit/tests` (도구를 서브프로세스로 실행해 stdout 1줄 JSON까지 확인, 약 25초).
공통 시험 `AIKIT/dev/tests`는 setup의 기본 스킬 목록과 세 도구의 `quote_arg`가 같은지 봅니다.
