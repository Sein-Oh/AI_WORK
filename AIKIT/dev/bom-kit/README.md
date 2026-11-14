# bom-kit (0.1.0)

BOM 업무 계산 도구입니다. pandas-kit(표를 고르고·비교하는 기반 도구)에 넣지 않는 BOM 업무 규칙을 한 곳에 둡니다
(사용자 결정 2026-11-14). 기본 스킬 `mbom-diff`(두 판 비교)와 `where-used`(쓰인 곳)가 pandas-kit 명령과 함께 씁니다.
원본 엑셀은 읽지도 쓰지도 않습니다: 입력은 pandas-kit `export`로 만든 csv(`엑셀 행` 열 포함)이고, 결과 csv·txt는 그 옆에 만듭니다.

## 파일
| 파일 | 하는 일 |
|---|---|
| `AIKIT/bom-kit/bom.py` | 진입점. `python bom.py <명령> ...` |
| `AIKIT/bom-kit/_bom_runtime.py` | 공통 런타임 (진입점 아님): `run_cli`, `JsonArgumentParser`, `quote_arg`, 오류 로그, 파일 복구 함수. pandas-kit `_pd_runtime.py`와 같은 이름·같은 동작 (표준 라이브러리만) |
| `tests/` | `python -m pytest -q AIKIT/dev/bom-kit/tests` (입력 csv는 pandas-kit으로 만들고, 샘플은 `AIKIT/dev/pandas-kit/tools/make_mbom_samples.py`) |

필요한 라이브러리: 없음 (표준 라이브러리). `requirements.txt`.
데이터 폴더(오류 로그 `logs/errors.txt`): `AIKIT/bom-kit/data`, 환경변수 `BOM_KIT_DATA_DIR`로 바꿈.

## BOM 업무 규칙 (이 도구 한 곳에만)
- 상위 = 위쪽 행 중 가장 가까운, Level이 더 작은 행 (상위 품번 열이 없는 BOM). Level이 두 단계 이상 건너뛰면 `warnings`.
- 경로 = 최상위부터 품번을 `/`로 이은 것. 같은 상위 아래 같은 품번이 둘 이상이면 모두 `#순번` (`300210#1`, `300210#2`).
- 제품 1대당 수량 = 그 행과 위쪽 상위들의 수량을 곱한 값 (스크랩·유효 기간·대체 그룹은 보지 않음).
- 두 판 비교: 경로(품번 기준)로 짝을 맞춤. 다른 상위로 옮겨진 품목은 삭제·추가로 둠. 같은 상위 아래 같은 자리(앞이나 뒤의 형제 품번이 같음)에서
  품번이 바뀐 것만 `[대체 후보]` (품명은 보지 않음, 삭제와 추가에 모두 있는 품번은 옮겨진 것이라 짝짓지 않음) (사용자 결정 2026-11-14).
- 열 찾기: Level(`Level`, `레벨` 등), 품번(`Part Number`, `품번`, `자재번호` 등), 수량(`Qty`, `수량` 등), 품명·단위(있으면). 못 찾으면 열 목록과
  `--level`·`--part`·`--qty`로 멈춤 (`STOP_RETRY`).

## 명령
| 명령 | 옵션 | 결과 |
|---|---|---|
| `path` | `--old`, `--new`, `--level`, `--part` | 두 판 csv 옆에 `<이름>_경로.csv`(`엑셀 행`, `상위 품번`, `경로`). `next_commands`: pandas-kit select·join(경로 붙이기, 두 판), diff `--on 경로 --name chg`, export `변경내용.csv` (6개, 차례로 모두). csv 파일 이름이 pandas-kit 표 이름 (`old.csv` -> 표 `old`, export가 붙인 시각 꼬리는 뗌) |
| `summary` | `--file` | `변경내용.csv`(diff를 경로 키로 export한 것) -> `[대체 후보]`·`[추가]`·`[삭제]`·`[바뀜]` 줄과 옛·새 엑셀 행, `counts`, 옆에 `<이름>_요약.txt`. 옆의 `<표>_경로.csv`로 형제 순서를 봄 (없으면 삭제 하나·추가 하나일 때만 대체 후보) |
| `where-used` | `--file`, `--find`, `--level`, `--part`, `--qty` | `--find`와 같은 품번(대소문자 무시) 또는 품명에 그 글자가 든 행마다 `[쓰임]` 줄(상위·제품 1대당 수량·경로·엑셀 행), `totals`(품번·단위별 곳 수와 합계), 옆에 `<이름>_쓰인곳.csv`. 0건은 `count: 0`과 "찾지 못했습니다" |

- 출력 상한: `lines` 30줄 (넘으면 `truncated: true`, 전체는 `output_file`).
- 결과 파일은 다시 읽어 행 수를 확인하면 `VERIFIED`. 저장 대상이 잠겨 있으면 시각을 붙인 새 이름(`recovered`).
- pandas-kit 명령은 `C:\AIKIT\pandas-kit\pandas_kit.py`(이 도구 폴더의 형제)를 부르는 전체 명령으로만 만든다. pandas-kit을 import하지 않는다.
