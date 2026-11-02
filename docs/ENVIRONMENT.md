# 표준 사용자 환경

`CLI/` 아래 모든 도구는 이 환경에서 **추가 설치 없이** 동작하는 것을 기본 목표로 합니다.
버전 기준은 루트 `AGENTS.md`의 "사용자 환경" 표와 아래 표입니다 (둘을 함께 갱신).

| 항목 | 값 |
|---|---|
| OS | Windows 11 (폐쇄망) |
| Python | **WinPython** 배포판, Python 3.13.13 |
| 실행 방식 | WinPython의 `python.exe`를 시스템 환경변수 `PATH`에 추가하여 `python` 명령으로 호출 |
| 에이전트 LLM | 26B~30B 경량 모델 (스킬로 도구 호출) |

## WinPython 사용 시 유의점

- 명령은 항상 **`python`** 으로 호출합니다. WinPython은 `py` 런처를 등록하지 않으므로 `py`, `python3`는 쓰지 않습니다.
- 패키지 설치도 `python -m pip install ...` 형식으로 해야 PATH에 잡힌 WinPython에 설치됩니다 (`pip` 단독 호출 금지).
- 설치 위치 확인: `python -c "import sys; print(sys.executable)"` 가 WinPython 폴더를 가리켜야 합니다.
- WinPython은 포터블 배포판이라 다른 PC로 폴더째 복사해도 같은 환경이 됩니다. 학습 PC와 예측 PC의 라이브러리 버전을 맞추기 쉬우므로, 가능하면 같은 WinPython 폴더를 그대로 복사해 쓰는 것을 권장합니다.

## 설치된 주요 라이브러리

| 분류 | 패키지 (버전) |
|---|---|
| 데이터/ML | pandas 3.0.5, scikit-learn 1.9.0, scipy 1.18.1 (numpy, joblib은 의존성으로 포함) |
| 시각화 | matplotlib 3.11.1, seaborn 0.13.2 |
| 문서 | openpyxl 3.1.5 (xlsx), python-docx 1.2.0 (docx), python-pptx 1.0.2 (pptx), pymupdf 1.28.2 / pymupdf4llm 1.28.2 (PDF), olefile 0.47 (HWP 5.0 읽기, office-cli용으로 추가) |
| 유틸 | PyYAML 6.0.3, python-dateutil 2.9.0.post0 |
| 영상 | opencv-python 5.0.0.93 |
| 에이전트/서버 | fastapi 0.141.1, fastmcp 3.4.7, mcp 1.28.1, langgraph 1.2.11, openai 2.54.0 |
| Windows | pywin32 312 |
| 테스트 | pytest 9.1.1 |

위 표는 주요 라이브러리만 적은 것입니다. 그 외 일반적인 패키지도 대부분 설치되어 있고,
없는 패키지는 Python 3.13용 wheel을 반입해 설치할 수 있습니다.

## 도구별 호환 현황

| 도구 | 상태 | 비고 |
|---|---|---|
| `ml-kit` 학습/예측/보고서 | 추가 설치 없이 동작 | 동일 버전(pandas 3.0.5, scikit-learn 1.9.0, openpyxl 3.1.5, python-docx 1.2.0)으로 전체 테스트 통과 |
| `office-cli` Office 자동화/문서 읽기 | 1.4.13, 회사 PC(Office 16) 실기 1.4.1에서 71/71 통과, 오프라인 테스트는 1.4.3까지 회사 PC 통과 | 열린 문서 작업은 pywin32 312 + Office 16.0. 닫힌 파일 읽기는 openpyxl / python-docx / python-pptx / pymupdf, HWP는 olefile 0.47, 차트 다시 그리기는 matplotlib. 오프라인 테스트는 이 버전으로 통과 |

## 새 도구를 만들 때 규칙

1. 위 목록에 있는 라이브러리를 우선 사용합니다.
2. 목록에 없는 패키지가 필요하면 Python 3.13 / win_amd64 wheel이 있는 것만 쓰고, 이 문서와 `AGENTS.md` 표에 버전을 기록합니다.
   import 실패 시에도 traceback 대신 `STOP_RETRY` JSON으로 안내되게 합니다.
3. 테스트는 위 표의 버전으로 실행해 통과해야 합니다.
