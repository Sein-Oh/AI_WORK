# AI_WORK

AI 에이전트가 CLI 도구로 호출하는 **Agent-Computer Interface(ACI)** 툴킷 저장소입니다.
대상 환경: **Windows 11 폐쇄망, WinPython (Python 3.13.13, PATH 등록)**, LLM 에이전트(스킬로 제공, 기준은 26B~30B 경량 모델 Gemma, 대형 모델 BEAM도 씀).
라이브러리 버전은 `AGENTS.md` "사용자 환경" 표 / [`docs/ENVIRONMENT.md`](docs/ENVIRONMENT.md).

## 트리

```text
AI_WORK/
├── README.md                  # 저장소 안내 (본 문서)
├── AGENTS.md                  # 에이전트 공통 규칙 + 사용자 환경 (단일 원본, OpenCode가 자동으로 읽음)
├── CLAUDE.md                  # @AGENTS.md 한 줄 (Claude Code용)
├── pytest.ini                 # python -m pytest -> AIKIT/dev/ 아래 모든 도구 테스트
├── pip_install.txt            # WinPython에 두 도구(ml-kit, office-cli)의 패키지를 표준 버전으로 설치하는 명령 (온라인, wheel 받기, 오프라인)
├── docs/
│   ├── ENVIRONMENT.md         # 표준 사용자 환경 정의 + 도구별 호환 현황
│   ├── DECISIONS.md           # 원칙급 결정 기록 (지금 유효한 원칙, 폐기·보류)
│   ├── HANDOFF.md             # 작업 인계: 지금 상태와 할 일만 (지난 경위는 history/)
│   ├── WINDOWS_OFFLINE.md     # 폐쇄망 wheel 반입/설치 가이드 (모든 도구 공통)
│   ├── history/               # 지난 설계·검토 기록 (1.0 설계, 1차 지침 준수 검토)
│   ├── opencode/              # 사내 PC OpenCode 스킬 배치 안내, 실기테스트.md (회사 PC 실기 절차 한 장)
│   │   └── 실기테스트_샘플/   # 정답을 아는 샘플 데이터 (AIKIT/dev/ml-kit/tools/make_beta_samples.py)
│   └── test_logs/             # 사내 PC 실기 결과 (GitHub 웹으로 업로드, 프로젝트 마무리 때 이력에서 정리)
├── archive/                   # 보관 도구 (수정·테스트 대상 아님, archive/README.md)
│   ├── engenius/              # 사내 에이전트 En-Genius 스킬 패키지 (업데이트 중지, 필요할 때 다시 꺼냄)
│   └── doc-classifier/        # 폴더 파일 목록화/분류 0.9.2 (다른 개발자 담당으로 보관)
├── AIKIT/                     # 사내 PC C:\AIKIT로 통째로 복사 (아래 AIKIT/dev 포함). OpenCode 프로젝트 폴더 (이 폴더를 프로젝트로 열 때만 우리 도구를 씀)
│   ├── AGENTS.md              # 전역 지침 (C:\AIKIT 프로젝트의 지침, 설정 파일 없이 폴더 모양으로 읽힘)
│   ├── .opencode/skills/<이름>/SKILL.md # 프로젝트 스킬: 폴더 하나 = 스킬 하나. 기본 12개(ml-train, ml-dataset, ml-predict, ml-report, ml-map,
│   │                          #   excel-kit, excel-format-kit, word-kit, ppt-kit, outlook-kit, search-kit, stats-kit), 그 밖은 사용자 스킬(git 제외)
│   ├── guide/                 # 사용자(사무직)용 사용 가이드 HTML: office-cli, ml-kit (파일 하나씩, 폐쇄망에서 브라우저로 열림)
│   ├── opencode.json          # 셸·권한·폐쇄망 잠금 (setup이 사용자 전역 설정에도 합침). 사내 LLM 설정 opencode.jsonc는 사용자 전역 설정에
│   ├── setup.py / setup.bat   # 배포 설치: 전역 설정에 우리 설정 합치기, 시스템 PATH 등록, 동작 확인
│   ├── web.bat                # OpenCode Web UI를 바로 띄움 (opencode 2.x: opencode pair의 일회용 로그인 링크를 브라우저로 엶)
│   ├── ml-kit/                # 회귀·분류 ML 툴킷
│   │   ├── ml_train.py        # 학습, 성능 개선, 학습 이력
│   │   ├── ml_dataset.py      # 학습용 데이터셋 만들기 (조건·행 번호로 뽑기, 학습에서 빼기)
│   │   ├── ml_predict.py      # 예측, 목표값 찾기, 조건별 예측, 최적 조건 (예측만 배포: 이 파일 + _ml_runtime.py + 버전 폴더)
│   │   ├── ml_report.py       # Word 보고서 (성능예측표도 함께)
│   │   ├── ml_export.py       # 성능예측표 (엑셀 계산기)
│   │   ├── ml_map.py          # 표 보간·반대 찾기 (모델 불필요)
│   │   └── _ml_runtime.py     # 공통 런타임 (진입점 아님)
│   └── office-cli/            # Office(Excel/Word/PPT/Outlook) 자동화 + 문서 읽기/검색
│       ├── excel.py / word.py / ppt.py / outlook.py   # 앱별 진입점
│       ├── reader.py          # PDF/HWP/Office/eml 파일 읽기, 캐시 등록, 근거(쪽/장절/번호) 검색 (Office 불필요)
│       ├── office.py          # 창 목록/전환, 앱 간 전달, 되돌리기
│       └── _runtime.py / _office.py / _markdown.py / _xlsx_formula.py   # 공통·내부 모듈 (진입점 아님)
└── AIKIT/dev/                 # AIKIT 안의 개발·실기 테스트 폴더 (C:\AIKIT\dev, 도구와 함께 배포, 도구는 읽지 않음)
    ├── README.md              # 도구 목록 + 새 도구 추가 규칙 / ACI 공통 계약
    ├── run_all_tests.py / .bat   # 실기 자동 점검 한 명령 (진단, pytest, 베타 점검, office 실기, --gemma -> result_<라벨>_<시각>.md 하나)
    ├── run_beta_check.py      # ml-kit 베타 기능 자동 점검 (정답을 아는 샘플)
    ├── run_simplify_check.py  # 사용법 단순성 점검 (run_all_tests.py의 simplify 단계로 실행)
    ├── grade_opencode.py      # Gemma 시험 채점과 실행 기록 묶음 (run_all_tests.py의 gemma 단계)
    ├── agent_scenarios.py     # Gemma 종합 테스트 요청과 채점 기준 (데이터 모듈)
    ├── tests/                 # 채점기·점검 러너·setup.py 테스트
    ├── work/                  # (git 제외) 모든 시험 폴더: 베타 점검, 실기 러너, Gemma, 눈 확인
    ├── ml-kit/
    │   ├── README.md          # 상세 설명서
    │   ├── CHANGELOG.md       # 변경 이력 + 남은 할 일
    │   ├── requirements.txt
    │   ├── tools/             # run_win_tests.py (Windows 실기 테스트 러너), make_beta_samples.py (샘플 생성)
    │   └── tests/             # CLI 계약 + 회귀 테스트
    └── office-cli/
        ├── README.md / CHANGELOG.md / TEST_CHECKLIST.md / requirements.txt
        ├── tools/             # run_win_tests.py, win_test_rules.py, make_office_samples.py (Gemma 테스트 O 샘플)
        └── tests/             # 오프라인 테스트 (가짜 COM + 실제 샘플 파일)
```

## 빠른 시작

```bash
python -m pip install -r AIKIT/dev/ml-kit/requirements.txt
python -m pip install -r AIKIT/dev/office-cli/requirements.txt
python -m pytest -q              # 전체 도구 테스트
```
WinPython에서는 `pip_install.txt`의 명령을 WinPython Command Prompt에 붙여넣어 두 도구의 패키지를 표준 환경 버전으로 설치합니다
(온라인 설치, 인터넷 PC에서 wheel 받기, wheel 폴더로 오프라인 설치).
폐쇄망 설치는 [`docs/WINDOWS_OFFLINE.md`](docs/WINDOWS_OFFLINE.md), 에이전트 연동은 `AIKIT/.opencode/skills/<이름>/SKILL.md`와 전역 지침 `AIKIT/AGENTS.md`.

## 배포 방법

배포는 `C:\AIKIT` 폴더 하나입니다 (저장소 `AIKIT/` + WinPython `python\` + `opencode.exe`).
1. `AIKIT` 폴더를 사용자 PC의 `C:\`에 복사합니다 (`C:\AIKIT\setup.bat`, `C:\AIKIT\python\python.exe`가 되도록, `dev` 포함).
2. 사내 LLM 주소가 든 `opencode.jsonc`를 사용자 전역 설정 폴더 `%USERPROFILE%\.config\opencode\`에 넣어 둡니다 (모든 프로젝트에서 사내 LLM 사용).
   `C:\AIKIT\opencode.jsonc`에 두어도 setup이 전역 폴더로 옮깁니다 (원래 파일은 보관).
3. `C:\AIKIT\setup.bat`을 두 번 클릭합니다 (또는 `python C:\AIKIT\setup.py`). `C:\AIKIT\opencode.json`의 셸·권한·폐쇄망 잠금을 전역 설정
   `opencode.json`에 합쳐 모든 프로젝트에 적용하고(원래 파일 보관), 예전 설치가 남긴 `OPENCODE_CONFIG_DIR=C:\AIKIT`를 지웁니다.
   시스템 PATH에 `C:\AIKIT\python`, `C:\AIKIT\python\Scripts`를 없을 때만 추가한 뒤, 도구와 OpenCode 설정을 확인합니다. `--check`는 확인만.
   우리 지침·스킬은 `C:\AIKIT`을 프로젝트로 열 때만 읽힙니다 (Web UI는 `web.bat`, TUI는 `C:\AIKIT`에서 `opencode`).
4. 새 터미널을 열고 OpenCode(TUI, Desktop, Web UI)를 다시 시작합니다. setup이 OpenCode 백그라운드 서비스를 꺼 두므로 새 PATH·스킬·지침으로 다시 뜹니다
   (TUI와 web.bat은 이미 떠 있는 서비스를 그대로 쓰므로, 설치 전에 뜬 서비스가 남으면 python을 못 찾고 옛 스킬을 씀. 그럴 때는 `C:\AIKIT\web.bat stop`).
5. (선택) 브라우저 화면으로 쓰려면 `C:\AIKIT\web.bat`을 두 번 클릭합니다. `opencode pair`가 OpenCode 백그라운드 서비스를 띄우고
   (이 PC에서만 접속, 화면은 opencode.exe 안에 들어 있어 폐쇄망에서도 됨) 일회용 로그인 링크를 브라우저로 엽니다. 링크는 한 번만 쓰이고
   몇 분 뒤 만료되니 다시 열 때는 web.bat을 다시 실행합니다. 서비스는 창을 닫아도 계속 돌며 `web.bat stop`으로 끕니다(`web.bat status`로 상태).
   브라우저는 Edge(없으면 Chrome)를 `C:\AIKIT\data\web_profile` 전용 프로필의 앱 창으로 엽니다. 평소 브라우저 프로필에서
   "Internal error opening backing store for indexedDB.open"이 나던 회사 PC 대응 (2026-11-10). 둘 다 없으면 기본 브라우저.
   Web UI는 세션을 프로젝트(작업 폴더) 안에 만듭니다. 처음 쓰는 PC는 폴더가 없어 새 세션이 만들어지지 않으니, 처음 한 번 화면에서 프로젝트 열기로
   `C:\AIKIT`(web.bat이 서비스를 띄우는 기본 프로젝트 폴더)을 고른 뒤 새 세션을 만듭니다. 프로젝트 목록은 브라우저 프로필(`C:\AIKIT\data\web_profile`)에 저장됩니다.

## 관리 방법

1. **코드는 압축 파일 대신 평문 파일로 커밋** - diff/리뷰가 가능해야 합니다.
2. **CLI 인터페이스(인자 + stdout 1줄 JSON)는 공개 API** - 바꾸면 `SKILL.md`, `README.md`, `CHANGELOG.md`를 같이 갱신.
3. **작업 1건 = 브랜치 1개 = PR 1개**, 리팩터링과 동작 변경은 분리.
4. **모든 PR은 `python -m pytest` 통과가 조건.**
5. **배포 = `CLI` 폴더 하나** - 아래 "배포 방법". OpenCode 스킬 폴더에는 `SKILL.md`만 들어갑니다 ([`docs/opencode/README.md`](docs/opencode/README.md)).
   ml-kit 예측만 배포할 때는 `ml_predict.py` + `_ml_runtime.py` + 학습된 `vX\` 폴더를 복사하면 됩니다.
