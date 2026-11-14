# CHANGELOG - bom-kit

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`.

## [0.1.0] - 2026-11-14

### 구조
- 새 도구 (사용자 결정 2026-11-14): mbom-diff·where-used 스킬 폴더의 스크립트 3개(`mbom_path.py`, `mbom_summary.py`, `where_used.py`)에
  따로 있던 같은 계산(Level로 상위·경로, 열 찾기, csv 읽기)을 `AIKIT/bom-kit/bom.py` 한 곳으로 (한쪽만 고치면 두 스킬이 같은 BOM을 다르게 읽음).
  공통 런타임 `_bom_runtime.py`는 pandas-kit `_pd_runtime.py`에서 복사 (표준 라이브러리만). 스킬은 그대로 둘 (필요한 순서만 읽음).
- 스킬 스크립트와 달리 공통 계약을 따름: 인자 오류·도움말 JSON, 같은 뜻 옵션 `[추천]`, 오류 로그, 잠긴 파일 새 이름, 인코딩 자동 판별,
  pandas-kit 명령의 셸 따옴표(`quote_arg`).
- 명령 하나로 (사용자 결정 2026-11-14, Gemma급 검토): 스킬이 pandas-kit 명령과 스크립트를 13단계(diff)·4단계(where-used)로 잇던 것을
  `bom.py diff`·`bom.py where-used` 하나로. bom-kit이 pandas-kit을 실행(import 아님, 데이터 폴더 따로)하고 최종 결과만 줌.
  중간 명령의 `next_commands`·지시를 경량 모델이 따라가 멈추는 위험, 긴 경로 옮겨 적기, 결과 JSON 13개가 쌓이는 문제를 없앰.
  스킬은 4KB -> 2.8KB.

### 기능
- `diff`, `where-used` (README). 결과는 스킬 스크립트와 같음 (v2 -> v2_mod 요약, where-used 4가지 찾기를 나란히 돌려 확인).
- pandas-kit이 시트·표 범위를 물으면 bom.py 전체 명령으로, 열을 못 찾으면 열 `[후보]`로. pandas-kit 경고는 사실만 (명령을 권하는 문장은 뺌).
- 시험: `test_mbom_diff.py`, `test_where_used.py`(pandas-kit 시험에서 옮김), `test_contract.py` (사용자의 pandas-kit 지금 작업을 바꾸지 않는지 포함).

## 남은 할 일
- 누적 수량·총소요량(제품 N대에 품목별 합계), BOM 한 판 점검(같은 상위 아래 같은 품번, 수량 0, 단위 섞임).
- Level 표기 `.1`, `..2`(점으로 깊이), 상위 품번 열이 이미 있는 BOM.
- 사용 가이드 HTML (`AIKIT/guide/`), 실기 러너·Gemma 시나리오 (지금은 `run_all_tests.py`가 버전과 오류 로그만 붙임).
