# Windows 11 폐쇄망 설치 가이드

`AIKIT/` 아래 모든 도구 공통 (각 도구의 `requirements.txt`는 `AIKIT/dev/<tool>/`에 있음). 폐쇄망 PC에서는 `pip install`이 인터넷에 접근할 수 없으므로
**인터넷 PC에서 wheel 파일을 미리 받아 반입**합니다.

## 1. 인터넷 PC에서 준비

표준 환경은 WinPython (Python 3.13) 입니다 ([`ENVIRONMENT.md`](ENVIRONMENT.md)). 대부분의 패키지는 이미 설치되어 있으므로
**표준 환경에 없는 패키지만** 반입하면 됩니다. 현재 ml-kit은 추가 반입이 필요 없습니다.

```bat
cd AIKIT\dev\ml-kit
python -m pip download -r requirements.txt -d wheels ^
  --only-binary=:all: --platform win_amd64 --python-version 3.13 --implementation cp
```
- 인터넷 PC가 폐쇄망과 **같은 Windows + 같은 Python 버전**이면 `--platform` 이하 옵션은 생략 가능합니다.
- 표준 환경과 같은 버전을 받으려면 `requirements.txt` 대신 필요한 패키지만 `openpyxl==3.1.5`처럼 버전을 붙여 적습니다 (버전: `docs/ENVIRONMENT.md`).

## 2. 반입 후 폐쇄망 PC에서 설치

PATH에 잡힌 `python`이 WinPython인지 먼저 확인합니다: `python -c "import sys; print(sys.executable)"`

```bat
cd AIKIT\dev\ml-kit
python -m pip install --no-index --find-links wheels -r requirements.txt
```
`wheels\` 폴더만 사용합니다(`--no-index`).

## 3. 버전 일치 (중요)

`.joblib` 모델은 scikit-learn 버전에 민감합니다. **학습 PC와 예측 PC에 같은 wheel 묶음**을 설치하십시오.
버전이 다르면 `ml_predict.py`가 `[WARN] scikit-learn 버전 불일치`를 출력합니다.
설치 후 `python -m pip freeze > wheels\lock.txt`로 고정 버전을 기록해 두면 다음 반입 때 같은 버전을 받을 수 있습니다
(`python -m pip download -r wheels\lock.txt ...`).

## 4. Windows 사용 시 주의

| 상황 | 동작 |
|---|---|
| 결과 CSV/보고서를 Excel·Word로 열어둔 채 재실행 | 잠금 감지 후 `_YYYYMMDD_HHMMSS` 붙은 새 파일로 저장 |
| 백신/탐색기가 버전 폴더를 잡고 있음 | 이름 변경·삭제를 최대 6회 재시도 |
| 경로가 240자 초과 | `[WARN]` 출력. 데이터를 `C:\work\` 같은 짧은 경로로 이동 |
| 메모장/PowerShell로 만든 txt·json (UTF-8 BOM, UTF-16, CP949) | 자동 인식 |
| CSV 출력 | CP949 기본 (Excel 더블클릭 시 한글 정상), 표현 불가 문자 시 utf-8-sig |
| 에이전트가 stdout 읽기 | 항상 UTF-8 1줄 JSON. PowerShell에서는 한글이 깨지지 않게 SKILL.md의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 명령 앞에 붙임 |
