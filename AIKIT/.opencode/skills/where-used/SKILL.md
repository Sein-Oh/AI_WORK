---
name: where-used
description: BOM 엑셀에서 한 품목이 어느 상위 조립체에 쓰였는지와 제품 1대당 수량을 찾는다 (pandas-kit 명령 + bom-kit). 상위 품번 열 없이 Level과 행 순서로 구조를 적은 BOM에서, 최상위부터의 경로와 상위 수량을 곱한 소요량을 알린다. "이 부품 어디에 쓰여", "볼트 쓰인 곳 찾아줘", "제품 한 대에 몇 개 들어가", "역전개"에 사용. 원본 파일은 고치지 않는다.
license: Proprietary
compatibility: opencode
metadata:
  version: "0.2.0"
  platform: windows
---

# where-used 스킬 (품목이 쓰인 상위 찾기)

pandas-kit 명령(표 읽기·비교)과 bom-kit 명령(BOM 계산)을 이어 쓰는 스킬이다. 결과 처리는 전역 지침 4장, 승인은 5장을 따른다.
두 판을 비교하는 일은 이 스킬이 아니라 mbom-diff다.

## 찾는 기준 (업무 규칙)
1. 상위 품번 = 위쪽 행 중 가장 가까운, Level이 더 작은 행. 경로 = 최상위부터 품번을 `/`로 이은 것.
   같은 상위 아래 같은 품번이 둘 이상이면 `#1`, `#2`(위에서부터 순서)를 붙인다.
2. 제품 1대당 수량 = 그 행의 수량에 위쪽 상위들의 수량을 모두 곱한 값. 스크랩, 유효 기간, 대체 그룹(수량 0)은 보지 않는다.
3. `--find`는 품번이 그대로 같거나 품명에 그 글자가 든 행을 찾는다 (대소문자 가리지 않음). 품번 여럿이 걸리면 품번별로 따로 센다.
4. 숨긴 행(폐기 품목 등)도 BOM에 있으면 함께 센다. 빼려면 3단계 뒤 `select --where "엑셀 행!=<행>"`으로 고른 표를 export해서 쓴다.

## 순서
명령은 아래 순서로 한 번씩 실행한다. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다. `<...>`만 바꾼다.

| 단계 | 명령 |
|---|---|
| 1. 새 작업 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" new` |
| 2. BOM 읽기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" select --file "<BOM.xlsx>" --name bom` |
| 3. csv로 내보내기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" export --from bom --save bom.csv` |
| 4. 쓰인 곳 찾기 | `python "C:\AIKIT\bom-kit\bom.py" where-used --file "<3단계 output_file>" --find "<품번이나 품명>"` |

- BOM이 첫 시트에 없으면 2단계에 `--sheet "<시트>"`를 붙인다. 표 범위를 `[후보]`로 물으면 사용자에게 보여 주고 고른 번호의 명령을 실행한다.
- 2단계의 `warnings`는 모두 사용자에게 알린다 (숨긴 행, 빈 행, 두 줄 머리글 등).
- 4단계가 Level·품번·수량 열을 못 찾아 멈추면, message의 열 목록을 보여 주고 사용자가 고른 열 이름으로 명령 끝에
  `--level "<열>"`, `--part "<열>"`, `--qty "<열>"`을 붙여 1번 다시 실행한다.
- 같은 BOM에서 다른 품목을 또 찾으면 4단계만 `--find`를 바꿔 다시 실행한다 (1~3단계는 다시 하지 않는다).

## 결과 알리기
- 4단계 `preview`를 그대로 알린다 (첫 줄은 몇 곳·제품 1대당 수량, 그다음 줄마다 `[쓰임] <품번>: <상위> 아래 <수량> (경로, 엑셀 행)`).
  줄을 직접 세거나 고쳐 쓰지 않는다.
- 0건이면 "찾지 못했습니다"를 그대로 알리고 다시 실행하지 않는다. 품번 일부나 품명으로 다시 찾을지 사용자에게 묻는다.
- 잘렸으면(`truncated: true`) 전체는 `output_file`(쓰인곳 .csv)에 있다고 알린다.
- `warnings`가 있으면 그대로 알린다 (Level을 건너뛴 행, 수량이 숫자가 아닌 행 등).
- 결과를 엑셀로 보고 싶다고 하면 `output_file`을 excel-kit으로 연다. 더 다루려면 pandas-kit `select --file "<output_file>"`로 읽는다.
