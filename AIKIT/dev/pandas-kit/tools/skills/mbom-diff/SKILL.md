---
name: mbom-diff
description: MBOM(제조 BOM) 엑셀 두 판에서 무엇이 바뀌었는지 찾는다 (pandas-kit 명령 + 이 폴더 scripts). 상위 품번 열 없이 Level과 행 순서로 구조를 적은 BOM에서, 품목마다 최상위부터의 경로로 짝을 맞춰 추가·삭제·이동·대체 후보·바뀐 칸을 알린다. "BOM 두 판 뭐가 바뀌었어", "MBOM 개정 내용 찾아줘", "v2와 v2_mod 비교"에 사용. 원본 파일은 고치지 않는다.
license: Proprietary
compatibility: opencode
metadata:
  version: "0.1.0"
  platform: windows
---

# mbom-diff 스킬 (MBOM 두 판의 변경 찾기)

pandas-kit 기본 스킬의 명령을 이어 쓰는 조합 스킬이다. 결과 처리는 전역 지침 4장, 승인은 5장을 따른다.

## 비교 기준 (업무 규칙)
1. 상위 품번 = 위쪽 행 중 가장 가까운, Level이 더 작은 행. 품목마다 최상위부터 품번을 `/`로 이은 **경로**로 두 판의 짝을 맞춘다.
   같은 상위 아래 같은 품번이 둘 이상이면 `#1`, `#2`(위에서부터 순서)를 붙인다.
2. 품번만으로 짝을 맞추지 않는다 (공용 부품은 여러 상위 아래에 있어 짝이 틀림). 엑셀 행 번호로도 맞추지 않는다 (행이 하나 추가·삭제되면 그 아래가 모두 밀림).
3. 같은 상위 아래 순서만 바뀐 것, 글자로 적힌 숫자가 숫자로 바뀐 것(`'18'` -> `18`)은 변경이 아니다. 표 위 제목 블록(BOM Rev, 작성일)은 비교하지 않는다.
4. 같은 품번이 한 곳에서 빠지고 다른 상위 아래에 생기면 **이동**, 같은 자리에서 품번이 바뀌면 **대체 후보**다. 대체 후보는 짐작이니 사용자에게 맞는지 묻는다.

## 순서
명령은 아래 순서로 한 번씩 실행한다. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다. `<...>`만 바꾼다.

| 단계 | 명령 |
|---|---|
| 1. 새 작업 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" new` |
| 2. 옛 판 읽기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" select --file "<옛 판.xlsx>" --name old` |
| 3. 새 판 읽기 | `python "C:\AIKIT\pandas-kit\pandas_kit.py" select --file "<새 판.xlsx>" --name new` |
| 4. 옛 판 csv | `python "C:\AIKIT\pandas-kit\pandas_kit.py" export --from old --save old.csv` |
| 5. 새 판 csv | `python "C:\AIKIT\pandas-kit\pandas_kit.py" export --from new --save new.csv` |
| 6. 경로 계산 | `python "C:\AIKIT\.opencode\skills\mbom-diff\scripts\mbom_path.py" --old "<4단계 output_file>" --new "<5단계 output_file>"` |
| 7. 비교 | 6단계 결과의 `next_commands`를 처음부터 끝까지 순서대로 모두 실행한다 (경로 붙이기, 경로로 diff, 결과 export). 중간 명령이 주는 `next_commands`는 따르지 않는다. |
| 8. 요약 | `python "C:\AIKIT\.opencode\skills\mbom-diff\scripts\mbom_summary.py" --file "<7단계 마지막 명령의 output_file>"` |

- BOM이 첫 시트에 없으면 2·3단계에 `--sheet "<시트>"`를 붙인다. 표 범위가 틀렸다고 하면 `--range`를 붙인다.
- 4·5단계와 7단계의 `--save`는 파일 이름만 쓴다 (도구가 작업 폴더에 저장). 다음 단계에는 결과의 `output_file`을 그대로 넘긴다.
- 2·3단계의 `warnings`는 모두 사용자에게 알린다.
- 6단계가 Level이나 품번 열을 못 찾아 멈추면, message의 열 목록을 보여 주고 사용자가 고른 열 이름으로 6단계 명령 끝에 `--level "<열>"` 또는 `--part "<열>"`을 붙여 1번 다시 실행한다.

## 결과 알리기
- 8단계 `preview`를 그대로 알린다 (줄마다 `[이동]`, `[대체 후보]`, `[추가]`, `[삭제]`, `[바뀜]`과 옛·새 엑셀 행). 줄을 직접 세거나 고쳐 쓰지 않는다.
- 잘렸으면(`truncated: true`) 전체는 `output_file`(요약 .txt)에 있다고 알린다.
- `[대체 후보]`가 있으면 대체가 맞는지 사용자에게 묻는다.
- 원본 엑셀에 바뀐 칸을 칠해 달라고 하면 7단계 diff 결과의 `confirm_commands`를 1번 실행한다 (사용자가 원할 때만).
