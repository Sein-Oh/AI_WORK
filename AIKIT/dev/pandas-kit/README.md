# pandas-kit (0.2.0)

엑셀·CSV 표를 단계별로 고르고 비교하는 기반 도구입니다 (office-cli처럼 조합해 쓰는 단위 기능, BOM 분석 같은 업무 판단은 스킬이 함). 원본 파일은 쓰지 않고, 단계마다 이름 붙인 표(`df1`, `df2`)를
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

필요한 라이브러리: pandas, openpyxl (xlsx), matplotlib (차트), pywin32 (Windows에서 열린 Excel 경로 찾기만, 없어도 됨). `requirements.txt`.

## 명령
| 명령 | 옵션 | 결과 |
|---|---|---|
| `select` | `--file`, `--sheet`, `--range`, `--from`, `--columns`, `--where`, `--name` | 파일(시트·범위)이나 만든 표에서 행·열을 골라 새 표 |
| `show` | `--from`, `--columns`, `--start` | 20행씩 (`엑셀 행` 포함, `--start`는 표 안의 몇 번째 행부터). 더 있으면 `truncated: true` + `full_result_file`(작업 폴더 `results/<표>_전체.csv`) + 다음 20행을 보는 `next_commands` |
| `join` | `--left`, `--right`, `--on`, `--name` | 키 열로 짝 맞춰 붙이기 (양쪽에 있는 행만). 오른쪽의 같은 이름 열은 `<열>_<오른쪽 표>` (아래 "키 열") |
| `only` | `--left`, `--right`, `--on`, `--name` | 왼쪽 표에만 있는 행 |
| `diff` | `--old`, `--new`, `--on`, `--name` | 바뀐 곳마다 한 행: `구분`(추가/삭제/바뀜), 키 열, `열`, `옛 값`, `새 값`, `<옛 표> 행`, `<새 표> 행` |
| `stack` | `--tables "df1,df2"`, `--name` | 위아래로 잇기. 맨 앞 `표` 열에 어느 표의 행인지 |
| `group` | `--from`, `--by`, `--sum`, `--name` | `--by` 값별 한 행: `행 수`, `<열> 합계`(`--sum`의 숫자 열), `엑셀 행 목록`(앞 10개). 단위 열(`UoM` 등)이 `--by`에 없고 한 묶음에 단위가 여럿이면 `warnings` + 단위를 넣은 전체 명령 |
| `sort` | `--from`, `--by`, `--desc`, `--name` | 정렬 (앞 열이 먼저, 기본 작은 것부터). 같은 값은 표 순서, 빈 칸은 끝. 엑셀 행 그대로 |
| `undo` | | 마지막 단계 하나 되돌리기 (그 이름은 그 전 표를 다시 가리킴) |
| `new` | | 새 작업 시작 (다음 `select --file`이 새 작업 폴더를 만듦) |
| `chart` | `--from`, `--kind`, `--x`, `--y`, `--title`, `--save`, `--overwrite` | 표를 차트 그림(.png)으로 (아래 "차트"). 표는 바꾸지 않고 단계로 남기지 않음 |
| `export` | `--from`, `--save`, `--overwrite` | 새 파일(.xlsx, .csv). 생략하면 작업 폴더의 `<표>.xlsx`, 이름만 주면 작업 폴더에 |

모든 명령은 `--work "<작업 폴더>"`를 받습니다 (지난 작업으로 돌아가기, `new` 제외). `--name`을 빼면 `df1`, `df2`, ... 중 안 쓴 첫 이름,
`--from`을 빼면 지금 표(마지막 단계의 표)입니다.

## 대상 파일과 표 찾기
- `.xlsx`, `.xlsm`, `.csv`. `--file`이 없고 만든 표도 없으면 열린 Excel 통합문서의 경로를 COM으로 읽습니다 (경로만, 문서는 건드리지 않음).
  하나면 그 파일, 여럿이면 `[후보]`와 같은 번호의 `confirm_commands`, 없으면 경로를 물으라고 안내합니다.
- 저장 안 된 통합문서(`Saved=False`, 저장한 적 없는 새 문서 포함)는 읽지 않습니다: `STOP_RETRY` "Excel에서 저장한 뒤 같은 명령을 다시 실행".
- 시트: `--sheet`가 없으면 파일을 열었을 때 보이는 시트. 없는 이름은 `[후보]`(시트 목록).
- 표 찾기: 완전히 빈 행으로 나뉜 내용 덩어리 중 가장 큰 것(머리글 칸 수 x 행 수, 같으면 앞의 것). 그 안에서 칸이 둘 이상 찬 첫 행이 머리글
  (위의 한 칸짜리 제목 행은 건너뜀), 머리글의 처음~마지막 찬 칸이 열. 표 위의 다른 내용(한 행에 두 칸 이상, 예: 공장·Rev 제목 블록)과
  빈 행 아래 내용은 `warnings`. `--range`를 주면 그 범위 전체가 표입니다.
  빈 행 아래에 내용이 더 있으면 `warnings`. `--range B5:H230`(또는 `B:H`, 왼쪽 위 칸 `B5`)를 주면 그 안의 첫 내용 행이 머리글이고,
  끝을 준 범위는 끝까지 읽으며 범위 안의 빈 행은 뺍니다(`warnings`). 찾은 범위는 preview에 (`B4:G8 (머리글 4행)`).
- 열 이름: 머리글 글자 그대로 (앞뒤 공백만 뺌). 빈 머리글은 `열_F`, 같은 이름은 `Qty_2`. `... 행`(예: `엑셀 행`)은 도구의 행 번호 열 이름이라
  `엑셀 행번호`처럼 `번호`를 붙이고 `warnings` (export한 표를 다시 읽을 때).
- csv 인코딩은 자동 판별 (UTF-8/BOM/UTF-16/CP949), 구분자는 `,` `;` 탭 중에서 찾습니다.

## 값은 바꾸지 않음
- 숫자 열: 칸이 모두 숫자이거나 앞자리 0이 없는 숫자 모양 글자일 때만 (`Int64`/`float64`). 그 밖에는 글자 그대로 (`0102`는 `0102`).
  숫자 사이에 글자로 저장된 숫자가 섞이면(Excel의 녹색 삼각형) 숫자 열로 읽고 `warnings`에 칸 수와 엑셀 행.
  숫자 사이에 숫자가 아닌 칸이 조금(10% 이하) 섞여 글자 열이 되면 그 칸 수와 엑셀 행을 `warnings`로 알립니다.
- `chart`: 단위 열(`UoM`, `Unit`, `단위` 등)에 단위가 여러 가지면 `warnings` (단위가 다른 수량을 더하거나 함께 그림).
  날짜 열은 전체가 날짜일 때만. 열마다 읽은 종류(`숫자`, `글자`, `날짜`, `빈 열`)를 `columns`에 줍니다.
- 공백·대소문자를 고치지 않고, 병합 칸의 빈자리도 채우지 않습니다. 병합 칸, 숨긴 행(자동 필터로 가려진 행 포함), 값이 저장되지 않은
  수식 칸(빈 값으로 읽음)은 개수를 `warnings`로 (30MB를 넘는 파일은 읽기 전용 모드라 이 셋은 확인하지 않고 그렇다고 알림).

## 열과 조건
- `--columns "Part Number,Qty"`: 머리글 이름(대소문자·띄어쓰기는 달라도 하나로 정해지면 그 열), 또는 원본 열 글자(`D`). `;`로 나누면 `[추천]`.
- `--where "Level=1;Qty>2"`: `;`는 그리고, `|`는 또는. `"Maker=A|B"`처럼 값만 쓴 갈래는 앞 갈래의 열·연산자를 씁니다.
  연산자 `=` `!=` `>` `<` `>=` `<=` `~`(포함) `!~`(포함 안 함), 빈 칸은 `"비고="`, 빈 칸 아님은 `"비고!="` (값 없는 `>`·`<`·`~` 등은 거절).
  갈래가 모두 `!=`·`!~`이면 "어느 것도 아님"입니다 (`"Type!=RAW|CONS"`는 RAW도 CONS도 아닌 행). 숫자 열은 숫자로, 날짜 열은 날짜로,
  글자 열에 숫자 값으로 `>` `<` `>=` `<=`를 쓰면 숫자로 읽히는 칸끼리 숫자로(숫자가 아닌 칸은 맞지 않음으로 보고 `warnings`), 그 밖은 글자 그대로
  (대소문자 구분) 비교합니다. 조건을 `,`로 나누면 `;`로 고친 전체 명령 `[추천]`. 조건은 `--columns`보다 먼저 적용합니다 (남기지 않을 열로도 거름).
- 없는 열 이름은 철자가 비슷한 열을 `[후보]`(같은 번호의 `confirm_commands`)로 묻습니다.

## 키 열 (`join`, `only`, `diff`의 `--on`)
- `"A,B"`: 두 표에서 같은 이름의 열. 이름이 다르면 `"Part Number=Item,Op No"`(왼쪽 열=오른쪽 열). 오른쪽에 없는 이름은
  오른쪽 표의 열을 `[후보]`로 묻고, 고른 번호의 `confirm_commands`는 `왼쪽=오른쪽`을 넣은 전체 명령입니다.
- `엑셀 행`은 그 표의 원본 엑셀 행 번호입니다. 스킬 스크립트가 export한 csv에 계산 열을 더해 돌려주면 `select --file`로 읽고
  `join --left df1 --right <읽은 표> --on "엑셀 행=엑셀 행번호"`로 원래 표에 붙입니다 (원래 표의 엑셀 위치를 그대로 가짐).
- 키 비교는 글자로 하되 정수는 종류와 상관없이 같게 봅니다 (16, 16.0, '16').

## 작업 폴더 (상태는 파일로)
- 처음 `select --file`에서 그 파일 옆에 `<파일 이름>_YYMMDDHHMMSS\`(예: `MBOM_DR200_261113153012\`, 같은 초면 `_2`)을 만들고, 그 뒤 명령은 "지금 작업 폴더"에 이어서 씁니다
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

## 차트 (`chart`)
- `--kind`: `bar`(막대, 기본), `line`(선), `scatter`(산점), `pie`(원), `hist`(히스토그램), `box`(상자). `--y`는 숫자 열(`,`로 여럿, 원은 하나).
- `--x`: 막대·선의 가로축(생략하면 `엑셀 행` 순서), 원의 조각 이름(필요), 산점의 가로축 숫자 열(필요). 히스토그램·상자는 `--x`를 쓰지 않음(`warnings`).
- 도구가 정하는 것: 막대·원에서 같은 `--x` 값이 여러 행이면 값마다 `--y`를 합계(`warnings`), 막대는 표 순서로 앞 30개까지(`truncated: true`),
  원은 큰 11개와 나머지 합계 `기타`, `--y`가 빈 행은 뺌(`warnings`), 제목은 생략하면 `<표>: <y 열> (<x>별)`.
- 숫자가 아닌 `--y`는 숫자 열을 `[후보]`로, 원에 `--y`가 여럿이면 첫 열로 고친 `[추천]`. 음수가 든 원은 `STOP_RETRY`.
- 저장: 작업 폴더 `results/<표>_<종류>.png` (이름만 준 `--save`도 작업 폴더 기준, `.png`만. 다른 확장자는 `[추천]` .png). 같은 이름이 있으면 시각을 붙인 이름(`recovered`),
  덮어쓰기는 `--overwrite`. 저장 뒤 그림을 다시 열어 확인하면 `VERIFIED`.
- 한글 글꼴은 맑은 고딕, 나눔고딕 순서로 찾고(ml-kit 차트와 같음), 없으면 한글이 네모로 보일 수 있다고 `warnings`. matplotlib는 창 없이(Agg) 그립니다.
- 그린 그림을 문서에 넣는 일은 office-cli (`excel.py write --image`, `word.py write --image`)에 `output_file`을 줍니다.

## 시험
`python -m pytest -q AIKIT/dev/pandas-kit/tests` (도구를 서브프로세스로 실행해 stdout 1줄 JSON까지 확인, 약 25초).
공통 시험 `AIKIT/dev/tests`는 setup의 기본 스킬 목록과 세 도구의 `quote_arg`가 같은지 봅니다.
