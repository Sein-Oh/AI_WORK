# bom-kit (0.1.1)

BOM 업무 계산 도구입니다. pandas-kit(표를 고르고·비교하는 기반 도구)에 넣지 않는 BOM 업무 규칙을 한 곳에 둡니다
(사용자 결정 2026-11-14). 기본 스킬 `mbom-diff`(두 판 비교)와 `where-used`(쓰인 곳)가 씁니다.

**명령 하나로 끝납니다** (사용자 결정 2026-11-14): Gemma 같은 경량 모델은 여러 단계를 잇다가 중간 명령의 지시(`next_commands`)를 따라가
멈추기 쉬워서, 표 읽기·경로 계산·비교·요약을 도구가 차례로 하고 최종 결과만 돌려줍니다. 원본 엑셀은 쓰지 않습니다.

## 파일
| 파일 | 하는 일 |
|---|---|
| `AIKIT/bom-kit/bom.py` | 진입점. `python bom.py <명령> ...` |
| `AIKIT/bom-kit/_bom_runtime.py` | 공통 런타임 (진입점 아님): `run_cli`, `JsonArgumentParser`, `quote_arg`, 오류 로그, 파일 복구 함수. pandas-kit `_pd_runtime.py`와 같은 이름·같은 동작 (표준 라이브러리만) |
| `tests/` | `python -m pytest -q AIKIT/dev/bom-kit/tests` (샘플은 `AIKIT/dev/pandas-kit/tools/make_mbom_samples.py`) |

필요한 것: 표준 라이브러리, 그리고 같은 `C:\AIKIT` 아래의 pandas-kit (`C:\AIKIT\pandas-kit\pandas_kit.py`, 표 읽기·비교에 실행). `requirements.txt`.
데이터 폴더(오류 로그 `logs/errors.txt`): `AIKIT/bom-kit/data`, 환경변수 `BOM_KIT_DATA_DIR`로 바꿈.

## pandas-kit과의 관계
- import하지 않고 `python pandas_kit.py <명령>`으로 실행만 합니다 (도구 폴더는 단독 배포). 명령마다 1줄 JSON을 읽습니다.
- pandas-kit의 데이터 폴더를 `<bom-kit 데이터 폴더>/pandas-kit`으로 따로 써서, 사용자가 pandas-kit으로 하던 지금 작업(지금 표)을 바꾸지 않습니다.
  pandas-kit 작업 폴더는 원본 옆(`<파일 이름>_<시각>`)에 생기고 결과 파일(변경내용 csv, 요약 txt, 쓰인곳 csv)도 거기 둡니다 (`work`).
  폴더가 쌓이지 않게 멈춘 실행의 폴더와, 같은 명령을 다시 실행했을 때 앞 실행의 폴더는 지웁니다 (다른 `--find`·다른 파일의 결과는 그대로).
- pandas-kit 경고는 사실만 옮깁니다 (`옛 판:`, `새 판:`, `BOM:`을 붙임). pandas-kit 명령·옵션을 권하는 문장은 뺍니다 (경량 모델이 그 명령으로 여러 단계를 다시 잇지 않게).
- pandas-kit이 시트·표 범위를 `[후보]`·`[추천]`으로 물으면 같은 값을 넣은 **bom.py 전체 명령**으로 바꿔 줍니다
  (옛 판의 시트는 두 판 공통 `--sheet`, 새 판만이면 `--new-sheet`, 범위는 `--old-range`·`--new-range`, where-used는 `--sheet`·`--range`).
  그 밖의 실패(파일 없음, 저장 안 된 문서 등)는 메시지만 옮깁니다.

## BOM 업무 규칙 (이 도구 한 곳에만)
- 상위 = 위쪽 행 중 가장 가까운, Level이 더 작은 행 (상위 품번 열이 없는 BOM). Level이 두 단계 이상 건너뛰면 `warnings`.
- 경로 = 최상위부터 품번을 `/`로 이은 것. 같은 상위 아래 같은 품번이 둘 이상이면 모두 `#순번` (`300210#1`, `300210#2`).
- 제품 1대당 수량 = 그 행과 위쪽 상위들의 수량을 곱한 값 (스크랩·유효 기간·대체 그룹은 보지 않음).
- 두 판 비교: 경로(품번 기준)로 짝을 맞춤. 다른 상위로 옮겨진 품목은 삭제·추가로 둠. 같은 상위 아래 같은 자리(앞이나 뒤의 형제 품번이 같음)에서
  품번이 바뀐 것만 `[대체 후보]` (품명은 보지 않음, 삭제와 추가에 모두 있는 품번은 옮겨진 것이라 짝짓지 않음) (사용자 결정 2026-11-14).
- 열 찾기: Level(`Level`, `레벨` 등), 품번(`Part Number`, `품번`, `자재번호` 등), 수량(`Qty`, `수량` 등), 품명·단위(있으면).
  못 찾거나 준 이름이 없으면 열들을 `[후보]`로, 고른 번호의 `confirm_commands`는 `--level`·`--part`·`--qty`를 넣은 전체 명령.

## 명령
| 명령 | 옵션 | 결과 |
|---|---|---|
| `diff` | `--old`, `--new`, `--sheet` (README에만: `--old-sheet`, `--new-sheet`, `--old-range`, `--new-range`, `--level`, `--part`) | 두 판을 읽고(pandas-kit select·export) 경로 계산, 경로 붙이기(join `엑셀 행`), `diff --on 경로`, 결과 export 뒤 `[대체 후보]`·`[추가]`·`[삭제]`·`[바뀜]` 줄과 옛·새 엑셀 행, `counts`, 끝 줄에 읽은 표. `output_file`(요약 txt), `changes_file`(변경내용 csv), `confirm_commands`(원본에 바뀐 칸 칠하기, office-cli, 사용자가 원할 때만). 0건은 "두 판이 같습니다" |
| `where-used` | `--file`, `--find`, `--sheet` (README에만: `--range`, `--level`, `--part`, `--qty`) | BOM을 읽고 `--find`와 같은 품번(대소문자 무시) 또는 품명에 그 글자가 든 행마다 `[쓰임]` 줄(상위·제품 1대당 수량·경로·엑셀 행), `totals`(품번·단위별 곳 수와 합계), `output_file`(쓰인곳 csv). 0건은 `count: 0`과 "찾지 못했습니다" |

- 출력 상한: `lines` 30줄 (넘으면 `truncated: true`, 전체는 `output_file`).
- 시간: pandas-kit 명령 diff 12개·where-used 3개 (클라우드 diff 약 4초, where-used 약 1초). 명령 실행 도구 120초 안.
- 결과 파일은 다시 읽어 행 수를 확인하면 `VERIFIED`. 저장 대상이 잠겨 있으면 시각을 붙인 새 이름(`recovered`).
