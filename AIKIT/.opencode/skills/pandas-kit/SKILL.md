---
name: pandas-kit
description: 엑셀·CSV 표를 단계별로 고르고 비교한다 (pandas_kit.py, 원본은 그대로). 파일·시트·범위에서 조건으로 행을 골라 이름 붙인 표(df1, df2) 만들기, 표 보기(엑셀 행 번호와 함께), 두 판 비교(추가·삭제·바뀐 칸), 짝 맞춰 붙이기, 한쪽에만 있는 행, 위아래로 잇기, 값별 행 수·합계, 정렬, 차트 그림(막대·선·산점·원·히스토그램·상자), 되돌리기, 새 파일로 저장. "BOM 두 판 비교", "레벨 1만 골라", "A에만 있는 품번", "업체별 수량 합계", "수량 큰 순", "업체별 수량 차트"에 사용. BOM 분석 같은 업무 판단은 이 명령들을 조합한 스킬이 한다. 원본에 칠하기·메모는 excel-kit.
license: Proprietary
compatibility: opencode
metadata:
  version: "0.2.5"
  platform: windows
---

# pandas-kit 스킬 (표를 단계별로 다루기)

## 규칙
1. 아래 명령을 그대로 쓰고 `<...>` 부분만 바꾼다. 옵션을 지어내지 않는다. 표 계산을 pandas 코드로 직접 짜지 않는다. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
2. 명령은 1번 실행하고 마지막 1줄 JSON만 읽는다. 결과 처리는 전역 지침 4장대로 한다.
3. 사용자에게는 `preview`와 `warnings`, `recovered`를 그대로 알린다. 이어서 할 일은 `next_commands`의 `command`를 글자 그대로 쓴다.
4. 표마다 이름이 있다 (`df1`, `df2`, ...). `--name`을 빼면 도구가 짓고 알린다. `--from`을 빼면 지금 표(마지막에 만든 표)를 쓴다.
5. 여러 값을 적을 때: 열 이름 목록은 `,`로, 조건 여러 개(그리고)는 `;`로, "또는"은 `|`로 나눈다 (예: `"Level=1;Qty>2"`, `"Maker=A|B"`).
6. 원본 파일은 고치지 않는다. 결과는 작업 폴더(원본 옆 `<파일 이름>_<날짜시각>`)에 쌓인다. 답할 때 경로는 결과 JSON의 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다.
7. 명령 표에 없는 일은 전역 지침 8장대로 한다.

## 명령
| 할 일 | 명령 |
|---|---|
| 파일에서 표 읽기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" select --file "<파일.xlsx 또는 .csv>"` (시트 `--sheet "<시트>"`, 범위 `--range B5:H230`, 열 `--columns "<열1>,<열2>"`, 조건 `--where "<조건>"`, 이름 `--name df1`) |
| 만든 표에서 고르기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" select --from df1 --where "<조건>" --name df2` (열만 `--columns "<열1>,<열2>"`) |
| 표 보기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" show --from df1` (열만 `--columns "<열1>,<열2>"`) |
| 두 판 비교 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" diff --old df1 --new df2 --on "<키 열1>,<키 열2>"` |
| 짝 맞춰 붙이기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" join --left df1 --right df2 --on "<키 열>"` (이름이 다르면 `--on "<왼쪽 열>=<오른쪽 열>"`, 왼쪽 행을 모두 남기려면 `--keep-left`) |
| 한쪽에만 있는 행 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" only --left df1 --right df2 --on "<키 열>"` |
| 값별 행 수·합계 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" group --from df1 --by "<열>" --sum "<숫자 열>"` (행 수만이면 `--sum` 없이) |
| 정렬 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" sort --from df1 --by "<열>"` (큰 것부터 `--desc`) |
| 위아래로 잇기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" stack --tables "df1,df2"` |
| 되돌리기 (한 단계) | `python "C:\AIKIT\pandas-kit\pandas_kit.py" undo` |
| 새 작업 시작 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" new` |
| 차트 그림 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" chart --from df1 --kind bar --x "<가로축 열>" --y "<숫자 열>"` (종류 `--kind bar`(막대) `line`(선) `scatter`(산점, `--x`도 숫자 열) `pie`(원, `--y` 하나) `hist`(히스토그램) `box`(상자), 제목 `--title "<제목>"`) |
| 새 파일로 저장 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" export --from df1` (다른 이름 `--save "<이름.xlsx 또는 .csv>"`, 같은 이름 파일 덮어쓰기는 사용자가 원할 때만 `--overwrite`) |

- `select`에는 `--file`이나 `--from`을 꼭 쓴다. 사용자가 "열린 엑셀"이라고만 하면 둘 다 빼고 실행한다: 도구가 열린 통합문서를 읽는 전체 명령을 `[추천]`이나 `[후보]`로 준다. 저장 안 된 문서는 읽지 않으니 사용자에게 저장해 달라고 한다.
- 표는 도구가 찾는다 (빈 행으로 나뉜 내용 중 가장 큰 것, 그 안에서 칸이 둘 이상 찬 첫 행이 머리글). 찾은 범위를 preview에 적는다. 틀렸다고 하면 `--range`를 준다.
- 열은 머리글 이름이나 엑셀 열 글자(`D`)로 준다. 품번 같은 값은 글자 그대로다 (`0102`).
- 조건 연산자: `=`, `!=`, `>`, `<`, `>=`, `<=`, `~`(포함), `!~`(포함 안 함). 날짜는 `2026-12-31`처럼 연-월-일로, 품번·코드는 적힌 그대로(`000125`) 쓴다. 빈 칸은 `"비고="`, 빈 칸 아님은 `"비고!="`. `"Type!=RAW|CONS"`는 RAW도 CONS도 아닌 행.
- `diff`의 `--on`은 행마다 달라야 한다 (예: `"Parent,Part Number"`). 같은 키가 여러 행이면 도구가 멈추고 알린다.
- 결과의 `엑셀 행`, `df1 행`은 원본 파일의 행 번호다. 사용자에게 알릴 때 그대로 쓴다.
- 지난 작업으로 돌아가려면 명령 끝에 `--work "<작업 폴더>"`를 붙인다 (결과의 `work` 값). 그 뒤 명령도 그 작업 폴더의 표를 쓴다 (바뀌었으면 preview에 나온다).

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `show`의 `rows`가 잘렸으면(`truncated: true`) 전체는 `full_result_file`에 있다고 알린다. 이어서 보려면 `next_commands`. 표를 직접 세거나 계산하지 않는다.
- `chart`는 그림 파일(`output_file`, .png)을 만든다. 문서에 넣어 달라고 하면 excel-kit·word-kit의 그림 넣기 명령에 그 경로를 쓴다. 같은 `--x` 값이 여러 행이면 도구가 합계로 그리고 `warnings`로 알린다.
- `diff`의 `confirm_commands`는 원본 파일의 바뀐 칸을 칠하는 office-cli 명령이다. 사용자가 칠해 달라고 했을 때만 1번 실행한다.
- `warnings`에 "옛 df1 기준"이 있으면, 그 표를 새 기준으로 다시 만들지 사용자에게 묻는다.
- `warnings`에 원본이 바뀌었다는 말이 있으면 그 안의 다시 읽는 명령을 사용자가 원할 때 실행한다.
