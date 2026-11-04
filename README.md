# AI_WORK

AI 에이전트가 CLI 도구로 호출하는 **Agent-Computer Interface(ACI)** 툴킷 저장소입니다.
대상 환경: **Windows 11 폐쇄망, WinPython (Python 3.13.13, PATH 등록)**, 26B~30B 경량 LLM 에이전트(스킬로 제공).
라이브러리 버전은 `AGENTS.md` "사용자 환경" 표 / [`docs/ENVIRONMENT.md`](docs/ENVIRONMENT.md).

## 트리

```text
AI_WORK/
├── README.md                  # 저장소 안내 (본 문서)
├── AGENTS.md                  # 에이전트 공통 규칙 + 사용자 환경 (단일 원본, OpenCode가 자동으로 읽음)
├── CLAUDE.md                  # @AGENTS.md 한 줄 (Claude Code용)
├── pytest.ini                 # python -m pytest -> CLI-dev/ 아래 모든 도구 테스트
├── docs/
│   ├── ENVIRONMENT.md         # 표준 사용자 환경 정의 + 도구별 호환 현황
│   ├── HANDOFF.md             # 작업 인계 (다른 PC·새 대화에서 이어 하기)
│   ├── WINDOWS_OFFLINE.md     # 폐쇄망 wheel 반입/설치 가이드 (모든 도구 공통)
│   ├── history/               # 지난 설계·검토 기록 (1.0 설계, 1차 지침 준수 검토)
│   ├── guide/                 # 사용자(사무직)용 사용 가이드 HTML: office-cli, ml-kit (파일 하나씩, 폐쇄망에서 브라우저로 열림)
│   ├── opencode/              # 사내 PC OpenCode 전역 AGENTS.md와 스킬 배치 안내, 실기테스트.md (회사 PC 실기 절차 한 장)
│   │   ├── 자동테스트/        # Gemma 종합 테스트 메시지 팩 (CLI-dev/agent_scenarios.yaml에서 생성)
│   │   └── 실기테스트_샘플/   # 정답을 아는 샘플 데이터 (CLI-dev/ml-kit/tools/make_beta_samples.py)
│   └── test_logs/             # 사내 PC 실기 결과 (GitHub 웹으로 업로드, 프로젝트 마무리 때 이력에서 정리)
├── archive/                   # 보관 도구 (수정·테스트 대상 아님, archive/README.md)
│   ├── engenius/              # 사내 에이전트 En-Genius 스킬 패키지 (업데이트 중지, 필요할 때 다시 꺼냄)
│   └── doc-classifier/        # 폴더 파일 목록화/분류 0.9.2 (다른 개발자 담당으로 보관)
├── CLI/                       # 사내 PC C:\CLI로 그대로 복사. 실행 파일과 SKILL.md만
│   ├── ml-kit/                # 회귀·분류 ML 툴킷
│   │   ├── SKILL.md           # 에이전트용 스킬 (짧은 명령 템플릿 + 결과 처리 규칙)
│   │   ├── ml_train.py        # 1단계: 학습
│   │   ├── ml_predict.py      # 2단계: 예측 (단독 배포 가능) + 공통 런타임
│   │   ├── ml_report.py       # 3단계: Word 보고서
│   │   ├── ml_map.py          # 표 보간·반대 찾기 (모델 불필요)
│   │   └── ml_export.py       # 성능예측표 (엑셀 계산기, 학습 PC용)
│   └── office-cli/            # Office(Excel/Word/PPT/Outlook) 자동화 + 문서 읽기/검색
│       ├── SKILL.md           # 에이전트용 스킬
│       ├── excel.py / word.py / ppt.py / outlook.py   # 앱별 진입점
│       ├── reader.py          # PDF/HWP/Office 파일 읽기, 캐시 등록, 근거(쪽/장절/번호) 검색 (Office 불필요)
│       ├── office.py          # 창 목록/전환, 앱 간 전달, 되돌리기
│       └── _runtime.py / _office.py   # 공통 모듈
└── CLI-dev/                   # 개발·실기 테스트용 (필요할 때만 C:\CLI-dev로, C:\CLI와 나란히)
    ├── README.md              # 도구 목록 + 새 도구 추가 규칙 / ACI 공통 계약
    ├── run_all_tests.py / .bat   # 실기 자동 점검 한 명령 (진단, pytest, 베타 점검, office 실기 -> result_<라벨>_<시각>.md 하나)
    ├── run_beta_check.py      # ml-kit 베타 기능 자동 점검 (정답을 아는 샘플)
    ├── grade_opencode.py      # OpenCode 대화 내보내기 채점 + 메시지 팩 생성
    ├── agent_scenarios.yaml   # Gemma 종합 테스트 요청과 채점 기준
    ├── tests/                 # 채점기·메시지 팩·점검 러너 테스트
    ├── ml-kit/
    │   ├── README.md          # 상세 설명서
    │   ├── CHANGELOG.md       # 변경 이력 + 남은 할 일
    │   ├── requirements.txt
    │   ├── tools/             # run_win_tests.py (Windows 실기 테스트 러너), make_beta_samples.py (샘플 생성)
    │   └── tests/             # CLI 계약 + 회귀 테스트
    └── office-cli/
        ├── README.md / CHANGELOG.md / TEST_CHECKLIST.md / requirements.txt
        ├── tools/             # run_win_tests.py, win_test_rules.py, office-cli-test/ (실기 테스트 스킬)
        └── tests/             # 오프라인 테스트 (가짜 COM + 실제 샘플 파일)
```

## 빠른 시작

```bash
python -m pip install -r CLI-dev/ml-kit/requirements.txt
python -m pip install -r CLI-dev/office-cli/requirements.txt
python -m pytest -q              # 전체 도구 테스트
```
폐쇄망 설치는 [`docs/WINDOWS_OFFLINE.md`](docs/WINDOWS_OFFLINE.md), 에이전트 연동은 각 도구의 `SKILL.md`.

## 관리 방법

1. **코드는 압축 파일 대신 평문 파일로 커밋** - diff/리뷰가 가능해야 합니다.
2. **CLI 인터페이스(인자 + stdout 1줄 JSON)는 공개 API** - 바꾸면 `SKILL.md`, `README.md`, `CHANGELOG.md`를 같이 갱신.
3. **작업 1건 = 브랜치 1개 = PR 1개**, 리팩터링과 동작 변경은 분리.
4. **모든 PR은 `python -m pytest` 통과가 조건.**
5. **배포 = 도구 폴더 복사** - `CLI/<tool>/` 폴더를 폐쇄망 PC의 `C:\CLI\<tool>`로 복사하고, OpenCode 스킬 폴더에는 `SKILL.md`만 복사 ([`docs/opencode/README.md`](docs/opencode/README.md)).
   ml-kit 예측만 배포할 때는 `ml_predict.py` + 학습된 `vX\` 폴더 2가지만 복사하면 됩니다.
