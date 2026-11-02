# office-cli 1.0 작업계획서

- 최종 갱신: 2026-10-27
- 전제: 이전 이력과 상관없이 **첫 배포(1.0.0)**로 보고 구조를 다시 설계합니다. 기존 34개 파일은 새 진입점이 테스트를 통과하면 한꺼번에 삭제합니다.
- 부록 A에 기존 코드 분석(버그 목록 포함)을 남겨 둡니다. 기능을 새 구조로 옮길 때 이 목록의 버그를 함께 고칩니다.

---

## 1. 확정된 결정

| # | 결정 | 근거 |
|---|---|---|
| D1 | **진입점 6개 + 공통 모듈 2개**로 통합 | 26B 에이전트의 도구 선택 부담과 SKILL.md 길이를 줄임. 소스 파일은 진입점당 1,500줄 안팎으로 유지(넘으면 `_<앱>_*.py` 내부 모듈로 분리) |
| D2 | **닫힌 파일 읽기는 라이브러리**(openpyxl, python-docx, python-pptx, pymupdf, olefile), **열린 문서와 쓰기는 COM** | Office를 띄우지 않아 빠르고 창 간섭이 없음. 회사 PC에서 라이브러리 직접 읽기 확인됨 |
| D3 | olefile 0.47을 표준 환경에 추가 | HWP 5.0 읽기. constraints / ENVIRONMENT.md / AGENTS.md 반영 완료 |
| D4 | 제거 기능은 과감히 삭제하고, 나중에 필요하면 다시 추가 | 6절 "재추가 후보" 목록으로 관리 |
| D5 | merged_code.md와 merge 도구 삭제 | 반입은 개별 파일로 함 |
| D6 | 메일은 기본으로 작성 창만 표시하고, `--send`를 명시했을 때만 발송 (수신자 확인 후) | 사용자 결정 |
| D7 | **회사 DRM 대응**: 앱 간 차트 전달은 여러 방식을 두고 회사 PC에서 동작하는 방식을 자동 선택 | Excel이 쓴 PNG를 Word는 읽지만 PPT는 못 읽음. 클립보드도 회사에서 불안정 |
| D8 | **화면 이동 규칙**: 읽기·조회는 사용자 화면(슬라이드, 시트, 스크롤)을 절대 움직이지 않음. 쓰기는 실제로 고친 위치로만 이동(`--no-focus`로 끔) | PPT를 읽기만 해도 마지막 슬라이드로 이동하던 문제 |
| D9 | 저장소 ACI 계약 준수 (stdout 1줄 JSON, exit 0/1/2/130, `STOP_RETRY:`/`ERROR:`, 로그 태그) | AGENTS.md, CLI/README.md |

## 2. 구조

```text
CLI/office-cli/
├── SKILL.md         # OpenCode 스킬 (짧게)
├── README.md        # 사람용 상세 설명서
├── CHANGELOG.md     # 1.0.0부터 관리 + 하단 "남은 할 일"
├── requirements.txt # pywin32, openpyxl, python-docx, python-pptx, pymupdf, matplotlib, olefile
├── excel.py         # 엑셀
├── word.py          # 워드
├── ppt.py           # 파워포인트
├── outlook.py       # 메일
├── reader.py        # PDF/HWP/EML 등 파일 읽기 + 읽은 문서 전체 검색 (Office 불필요)
├── office.py        # 창 목록/전환, 앱 간 전달, 되돌리기
├── _runtime.py      # 공통: 응답 계약(run_cli, JsonArgumentParser), 경로, 텍스트/CSV 입출력, 작업 폴더, 설정
├── _office.py       # 공통: COM 세션(active/파일), 스냅샷, 화면 이동 규칙
├── tools/run_win_tests.py  # Windows 실기 테스트 러너
└── tests/           # 계약 테스트 + 라이브러리 기반 실제 파일 테스트 + 가짜 COM 테스트
```
- `reader.py`와 `_runtime.py`는 pywin32 없이도 import·실행되어야 합니다.

## 3. 명령 체계

파일 이름은 프로그램이고, 서브커맨드는 하고 싶은 일(동사)입니다. 같은 일은 앱이 달라도 같은 동사를 씁니다.

| 명령 | 뜻 | 기존 파일 |
|---|---|---|
| `excel.py / word.py / ppt.py read` | 내용 읽기 (`--save`로 CSV/md 저장) | excel_read, word_read, ppt_read |
| `... info` | 문서 상태 요약 (열린 문서 목록 포함) | *_inspect |
| `... write` | 값·표·수식·텍스트·이미지 쓰기와 추가 (`--file`로 CSV/md 입력) | excel_fill, word_write, ppt_write |
| `... replace` | 찾아 바꾸기 (+서식) | *_edit의 치환 기능 |
| `... export` | PDF나 사본으로 저장 | *_export |
| `excel.py sheet` | 시트 추가·이름변경·복사·삭제·선택 | excel_sheet |
| `excel.py range` | 너비·높이·병합·틀고정·숨김·이동·복사·지우기 | excel_range |
| `excel.py style` | 테마·조건 강조·배경·글꼴·숫자 서식 | excel_edit의 서식 기능 |
| `excel.py chart` | 차트 만들기·지우기 | excel_chart |
| `ppt.py slide` | 슬라이드 추가·복제·이동·삭제·비우기 | ppt_slide |
| `outlook.py list / search / read / current / folders` | 메일 조회 | outlook_read |
| `outlook.py new / reply / forward` | 메일 작성 (D6) | outlook_write |
| `reader.py read / search / list` | 파일 읽기 / 읽은 문서 전체 검색 / 색인 목록 | pdf_read, hwp_read, eml_read, search |
| `office.py windows / focus` | 열린 Office 창 목록과 화면 전환 | office_window |
| `office.py transfer` | 앱 간 전달 (차트, 표, 텍스트) | office_bridge |
| `office.py undo` | 직전 작업 되돌리기 (`--list`로 목록 보고 선택) | clean |

공통 옵션: `--target`(active, 열린 문서 이름, 경로), `--file`(입력), `--save`(결과 파일), `--range`, `--sheet`, `--slide`, `--no-focus`

## 4. 체이닝 규칙 (데이터는 파일이나 Office 객체로, LLM은 짧은 참조만)

1. **교환 형식**: 표는 CSV(utf-8-sig), 텍스트는 마크다운(.md), 이미지는 PNG(Python이 쓴 파일)
2. **결과 JSON의 참조값**: `target_resolved`(실제 작업한 문서), `output_file`, `chart_name`, `sheet`, `slide`, `range_written`
3. **`next_commands`**: 자주 이어지는 다음 명령을 완성된 형태로 제공 (chart 다음 transfer, search 다음 write 등)
4. **작업 폴더**: `--save`에 경로를 주지 않으면 `data_dir\_work\`에 자동으로 이름을 붙여 저장. 오래된 파일은 자동 정리
5. **결과 길이 제한**: 읽기 결과가 상한(기본 200행 또는 8KB)을 넘으면 `truncated: true`로 자르고, 다음 구간을 읽는 명령을 `next_commands`로 제공

대표 시나리오:
```bat
python excel.py write --range A1:B101 --header "x,sin(x)" --formula "=(ROW()-2)*4*PI()/99" "=SIN(A2)"
python excel.py chart --range A1:B101 --type line --title "사인 2주기"
python office.py transfer --from excel --chart "사인 2주기" --to word,ppt
python office.py transfer --from word --what text --to excel
python reader.py search --query "마일스톤" --save
python word.py write --file "<output_file>"
```

## 5. 앱 간 차트 전달 방식 (D7)

| 방식 | 동작 | 비고 |
|---|---|---|
| `export` | Excel `Chart.Export` PNG -> 삽입 | PNG 첫 바이트를 검사해 암호화됐으면 즉시 다음 방식으로. **PPT 대상에는 쓰지 않음** (회사 DRM에서 PPT가 못 읽음) |
| `redraw` | COM으로 계열 데이터(숫자)만 읽고 matplotlib으로 PNG를 그려 삽입 | 클립보드·Office 파일 쓰기 없음. 원본 서식은 근사치 |
| `relay` | Excel 차트 -> Word에 그림으로 넣은 뒤 -> PPT로 넘김 | 회사에서 Word는 받아들인 이력 반영 |

- `office.py transfer --probe`: 세 방식을 실제로 시험해 대상 앱별로 동작하는 방식을 `config.json`에 저장합니다. 이후 자동 사용합니다.
- 결과 JSON의 `method`에 실제 사용한 방식을 적습니다.

## 6. 1.0에서 제거하는 기능 (재추가 후보)

사용자가 추가 기능을 물으면 이 목록을 안내합니다.
- 스타일 DNA 추출 (`office_theme --action extract`). 테마 프리셋(`--theme navy` 등)은 유지
- 검색 결과를 Office로 바로 넣기 (`search --to`). `--save` 후 `write --file`로 대체
- 검색어가 파일명과 같으면 파일을 반환하는 처리
- 샌드박스 가상 데스크톱 전환 (exebox, `SetThreadDesktop`)
- `clean fix-env`, JIT 자가복구의 "유령 프로세스 정리"
- 캐시의 "파일당 최근 2개 버전" 정책. 파일당 텍스트 캐시 1개와 색인으로 단순화
- recipe 관련 코드, `run.py` 호환 흔적, `excel_style.py` 등 별칭 진입점
- 새 문서 만들기, 파일 열기 (`office_window --new/--open`)
- EML 마스터 JSON 빌드·병합·스레드 (`eml_read build/merge/thread`). EML 단일 파일 읽기는 `reader.py read`로 유지

## 7. 진행 단계

| 단계 | 내용 | 상태 |
|---|---|---|
| 0 | `.py` 복원, 정적 검사, 가짜 COM 테스트, CLI/office-cli 이동 | 완료 |
| 1 | (구 구조) 데이터 손상·창 간섭 P0 수정 15건 | 완료. 새 구조로 그대로 옮김 |
| 2 | `_runtime.py` + 계약 테스트 | 완료 (`tests/test_runtime.py`) |
| 3 | `reader.py` (라이브러리 기반, Linux에서 실제 파일로 테스트) | 완료 (`tests/test_reader.py`). HWP 5.0은 실제 파일 검증 필요 |
| 4 | `_office.py` (COM 세션, 스냅샷, 화면 이동 규칙) | 완료 (`tests/test_office_core.py`, 가짜 COM) |
| 5 | `excel.py` -> `word.py` -> `ppt.py` (부록 A의 버그 함께 수정) | 완료 (`tests/test_excel.py`, `test_word.py`, `test_ppt.py`) |
| 6 | `outlook.py`, `office.py` (transfer 3방식 + probe, undo) | 완료 (`tests/test_outlook.py`, `test_office_cmd.py`) |
| 7 | SKILL.md, README, CHANGELOG 1.0.0, requirements.txt, run_win_tests 갱신, 기존 파일 삭제 | 완료. Windows 실기 검증 대기 |

---

# 부록 A. 기존 코드 분석과 1차 수정 기록 (2026-10-27)

- 작성일: 2026-10-27
- 분석 대상: `merged_code.md`에 담긴 루트 소스 파일 34개(약 11,000줄)와 `README.md`(구버전)
- 분석 방법:
  - merged_code.md를 파일별로 나눠 `py_compile`(문법 검사)와 `pyflakes`(정적 분석)를 돌렸습니다.
  - 34개 파일을 모두 정독했습니다.
  - Windows/COM 환경이 없어 실제 실행은 하지 않았습니다. 아래 항목은 모두 코드를 읽고 판단한 결과입니다. 확실하지 않은 항목은 `(?)`로 표시했습니다.

---

## 0. 현황 요약

### 0.1 코드 상태
| 항목 | 결과 |
|---|---|
| 문법 오류 | 없음 (34개 파일 모두 컴파일 통과) |
| 정의되지 않은 이름 | `word_export.py:105` `fixed_err_msg`: 실제로 실행되는 경로에서 NameError 발생 |
| 사용되지 않는 import | 42건 |
| 공통 모듈 | `_com_guard`(1020줄), `_office_utils`, `_config`, `_cache`, `_outlook_guard` |
| 진입점 시그니처 | `execute(wb, ws, params)`, `execute(app, doc, params)`, `execute(params)` 세 가지가 섞여 있음 |

### 0.2 README와 실제 코드의 차이
| README에 있음 | 코드 현황 | 조치 |
|---|---|---|
| `atomic_agent.py`, `atomic_agent_tui.py` | merged_code.md에 없음 | 하위 폴더에 있는지, 삭제했는지 **확인 필요** |
| `ml_pandas`, `ml_eda`, `ml_feature`, `ml_train`, `ml_report` | 없음 | 위와 같음 |
| `excel_style.py` (호환 진입점) | 없음 | README에서 삭제하거나 별칭(alias)으로 다시 추가 |
| `config.json`의 `recipe_dir` 관련 함수 | `_config.py`에 남아 있음 (Zero-JSON 원칙과 충돌) | 사용하지 않으면 제거 |
| JIT 자가복구 "유령 프로세스 정리" | 실제로는 `CoUninitialize` 후 `CoInitialize`만 함 (`killed: []`) | 문서를 실제 동작에 맞추거나 기능을 구현 |
| `excel_inspect` 다중 프로세스 전수 탐색 | 첫 번째 인스턴스를 찾으면 탐색을 멈춤 | 버그로 분류해 수정 (P1) |
| (README에 없음) `eml_read.py` | 코드에 있음 | README에 추가 |

---

## 1. 단계별 로드맵

| 단계 | 목표 | 산출물 |
|---|---|---|
| **Phase 0** | 작업 기반 마련 | 실제 파일 구조 복원, 정적 검사, 오프라인 테스트 장치 |
| **Phase 1 (P0)** | 사용자 데이터 손상과 화면 간섭 제거 | 치명 버그 약 15건 수정 |
| **Phase 2 (P1)** | 기능 오류 수정 | 모듈별 기능 버그 약 50건 |
| **Phase 3** | 공통화와 성능 최적화 | 세션, 응답, 경로 처리 통합, COM 호출 수 축소 |
| **Phase 4** | 고도화 | 안전장치, 도구 스키마, 테스트 확대, README 재작성 |

Phase 1과 2는 버그 수정만 하고 구조 변경은 최소로 둡니다. 구조를 바꾸는 큰 작업은 Phase 3에서 몰아서 합니다.

---

## 2. Phase 0: 기반 정비

1. **소스 복원**: merged_code.md를 `office-cli/*.py`로 다시 나눠 저장소에서 diff로 추적할 수 있게 합니다. merged_code.md는 반출용 산출물로만 남깁니다.
2. **정적 검사**: `ruff`(pyflakes와 import 정리 규칙)와 `python -m compileall`을 기본 검사로 둡니다.
3. **오프라인 테스트 장치** (COM이 없어도 돌아가는 부분부터):
   - 순수 로직 단위 테스트:
     - `range_to_matrix`, `parse_color_to_bgr`, `get_target_range_address`
     - `parse_markdown_table`, `normalize_and_expand_rules`
     - HWP 레코드 파서, EML 디코딩, search 쿼리 파서
   - 가짜 COM 객체(FakeRange / FakeWorkbook)로 `execute()`의 분기를 검증합니다.
   - Windows 실기 검증용 스모크 스크립트 `smoke_windows.ps1`: 모듈마다 대표 명령 1~2개를 실행하고 JSON 파싱 결과를 확인합니다.
4. ~~반출 자동화 (`tools/merge_code.py`)~~: 사용자 결정으로 삭제했습니다. 반입은 개별 파일로 합니다.

---

## 3. Phase 1 (P0): 데이터 손상, 화면 간섭, 보안

| # | 위치 | 문제 | 수정 방향 |
|---|---|---|---|
| 1 | `_com_guard.py:482, 722` | 파일 모드에서 `Dispatch("Excel/Word.Application")`가 **사용자가 이미 띄운 인스턴스에 붙은 뒤** `Visible=False`, `DisplayAlerts=False`를 실행합니다. 결과적으로 사용자 창이 숨겨지고, 마지막에 `Quit`이 호출될 수도 있습니다 | 백그라운드 작업은 `DispatchEx`로 전용 인스턴스를 만들고, 끝나면 그 인스턴스만 `Quit`합니다 |
| 2 | `_cache.py:242-251` | Word 스냅샷이 (1) 저장된 문서는 `doc.Save()`로 **사용자의 미저장 변경을 원본에 강제로 저장**하고 (2) 저장 안 된 문서는 `SaveAs(캐시 경로)`로 **문서 경로 자체를 캐시 폴더로 바꿉니다**. 이 파일은 나중에 회전 정책으로 삭제됩니다 | `doc.Range().WordOpenXML`(또는 `doc.WordOpenXML`)을 Flat OPC `.xml`로 저장합니다. Word의 SaveCopy 대체 수단으로, 원본 경로를 바꾸지 않습니다 |
| 3 | `_com_guard.py:953` | 조건에 `snapshot and ... and file_exists`가 있어 **활성 PPT는 스냅샷을 찍지 않습니다**. undo가 동작하지 않습니다 | 조건에서 `file_exists`를 빼고 활성/파일 모드 모두 `SaveCopyAs`를 씁니다 |
| 4 | `_com_guard.py:456-464, 507-513` | 시트명을 잘못 쓰면 **읽기 도구에서도 새 시트를 만들고**, 파일 모드에서는 그대로 저장합니다 | `create_sheet_if_missing` 인자를 추가합니다(쓰기 도구만 True). 대소문자를 구분하지 않고 비교합니다 |
| 5 | `_com_guard.py:972-985` | PPT 파일 모드가 변경이 없어도 항상 `Save()`합니다. mtime이 바뀌어 캐시가 무효화됩니다. 새 파일은 `Save()`만 하고 지정 경로로 `SaveAs`하지 않습니다 | `Saved` 플래그를 확인하고, 새 파일은 `SaveAs(abs_path)`로 저장합니다 |
| 6 | `word_session`, `ppt_session`과 읽기 도구 | 존재하지 않는 경로를 읽으면 **빈 문서를 그 경로에 만들고 저장합니다** (`word_read` 등) | 세션에 `mode="read"/"write"`를 둡니다. 읽기 모드에서 파일이 없으면 오류를 반환합니다 |
| 7 | `word_export.py:48-55` | `--page-number`가 **원본 문서에 바닥글을 영구적으로 넣습니다**. 실행할 때마다 중복됩니다 | 사본을 만든 뒤 사본에 적용해 내보냅니다. 또는 내보낸 뒤 `Undo`로 되돌립니다 |
| 8 | `word_export.py:113-122` | docx 내보내기가 원본에 `Save()`를 호출합니다. 실패 시 대체 경로의 `SaveAs`가 열린 문서의 경로를 바꿉니다 | `SaveAs2`를 사본 문서에 적용합니다. 원본은 건드리지 않습니다 |
| 9 | `excel_range.py:218-220` | `move`를 "복사 후 원본 `Clear()`"로 구현해서 **원본과 대상 범위가 겹치면 데이터가 사라집니다** | `Range.Cut(Destination=...)`를 씁니다 |
| 10 | `excel_edit.py:377, 407` | 계산 결과값으로 매칭한 뒤 텍스트를 다시 써서 **수식이 상수로 바뀝니다** | `HasFormula`인 셀은 건너뛰거나 `.Formula`에서 치환합니다 |
| 11 | `excel_export.py:179-201` | XLSX 내보내기에 `try/finally`가 없어 오류 시 고아 통합문서가 남고 `DisplayAlerts=False`가 유지됩니다 | finally에서 정리하고 원래 설정을 복원합니다 |
| 12 | `clean.py:50` | `undo`가 확인도, 현재 버전 백업도 없이 원본을 덮어씁니다. 파일이 열려 있으면 복사에 실패하는데도 "restored"(ok)로 보고합니다 | 현재본을 `.bak`으로 백업하고, 열린 문서는 COM으로 닫은 뒤 다시 엽니다. 실패하면 error를 반환합니다 |
| 13 | `outlook_write.py:62, 112, 162` | `send=True`면 확인 없이 즉시 발송하고, 아무 경로의 파일이나 첨부할 수 있습니다 | 발송 기본값은 `Display()`를 유지합니다. `--send`는 CLI에서 사람이 직접 지정할 때만 허용하고 에이전트 도구 스키마에서는 제거합니다. 첨부는 `data_dir`로 제한하는 옵션을 둡니다 |
| 14 | `ppt_slide.py:120-126` | `clear`에 `--slide`가 없으면 **모든 슬라이드의 모든 도형을 삭제합니다** | `--slide`를 필수로 하거나 `--all`을 명시해야 동작하게 합니다 |
| 15 | `excel_range` merge | 범위를 지정하지 않으면 UsedRange 전체를 병합합니다. 활성 모드에서는 확인 대화상자가 떠서 COM 호출이 멈춥니다 | 범위를 필수로 하고, 병합할 때만 `DisplayAlerts=False`를 잠시 적용한 뒤 복원합니다 |

---

## 4. Phase 2 (P1): 기능 오류

### 4.1 공통 및 핵심 모듈
- **창 탐색 콜백** (`_com_guard.py:160-279, 357-362`)
  - `EnumWindows`/`EnumChildWindows` 콜백이 False를 반환하면 pywin32에서 예외가 납니다(?). 이 때문에 포그라운드 창 바인딩 경로가 매번 실패하는 것으로 보입니다.
  - `find_all_*`가 인스턴스 1개를 찾으면 멈춰서 다중 프로세스 탐색이 동작하지 않습니다.
  - 수정: 콜백은 항상 True를 반환하고 결과만 모읍니다. 찾은 Application은 `Hwnd` 기준으로 중복을 제거합니다.
- **열린 문서 매칭** (`get_all_running_*`): 이름(`Name`)만으로 중복을 제거해서 이름이 같고 폴더가 다른 문서가 누락됩니다. `FullName` 기준으로 바꿉니다.
- **JIT 복구**: `jit_repair_office_environment`가 COM 객체가 살아 있는 상태에서 `CoUninitialize`를 호출해 위험합니다. 재초기화는 빼고 "지연 후 재바인딩"만 남깁니다.
- **응답 생성**: `format_response`의 `extra`가 `status`/`action`/`preview`를 덮어씁니다(clean.py에서 실제로 발생). 예약 키는 보호합니다.
- **색상 변환**: `parse_color_to_bgr(0)`이 검정을 None으로 처리합니다. `office_theme.py:328-343`도 같은 원인으로 검정 글꼴을 놓칩니다. `is None`으로 판정합니다.
- **오류 메시지 변환**: `friendly_error_message`의 `"1004" in msg`, `"RPC" in msg`가 너무 넓게 걸립니다. HRESULT 정수로 비교합니다. 이 함수는 어느 모듈도 쓰지 않으므로 공통 오류 응답에 연결합니다.
- **종료 코드**: `os._exit(0)`(ppt_edit, ppt_export, ppt_slide, ppt_write, office_bridge) 때문에 실패해도 종료 코드가 0입니다. 정상 종료로 바꾸고 status에 따라 0/1을 반환합니다.
- **예외 시 JSON 미출력**: `excel_fill.py:16`, `excel_sheet.py:15`, `excel_chart.py:91-95`, `excel_edit.py:282`, `word_edit` / `ppt_edit` / `office_theme` / `office_window` / `outlook_*` 의 `execute`, `search` 폴더 모드, 세션 열기 실패 등에서 예외가 나면 JSON 없이 traceback만 출력됩니다. 공통 래퍼를 둬서 어떤 예외든 JSON 1줄로 내보내게 합니다.

### 4.2 Excel
- **한 행 범위 처리** (`excel_read.py:25`, `excel_export.py:71`, `excel_inspect.py:34`, `excel_edit.py:381-383`): 한 행에 여러 칸인 범위의 `.Value`(`((a,b,c),)`)를 평탄한 튜플로 잘못 처리합니다. 결과가 깨지거나, 오류를 삼킨 채 no_match가 됩니다. 공통 `range_to_matrix()` 함수로 통일합니다.
- **excel_read**
  - 캐시 키에 sheet/search/color가 빠져 있어 다른 시트의 데이터나 필터링 안 된 데이터를 반환합니다.
  - 파일이 Excel에 열려 있고 미저장 변경이 있으면 디스크 기준의 오래된 캐시를 반환합니다.
  - 2500셀을 넘으면 요청 범위와 상관없이 `A1:Z50`만 탐색합니다.
  - `--color`는 red와 hex만 인식합니다.
- **excel_edit**: 규칙 `=100:green`이 부분 문자열로 매칭돼 1000, 2100 같은 값에도 색이 칠해집니다. `underline`은 받기만 하고 적용하지 않습니다.
- **excel_export**
  - 길이 2 이하 토큰을 열로 판단해서 `A1,B2` 같은 셀 목록을 열로 오인합니다.
  - CSV 직통 경로는 `--format`과 `--columns`를 무시합니다. 그래서 `.xlsx` 이름으로 CSV 텍스트가 저장될 수 있습니다.
  - 값이 화면 표시값(`.Text`)이 아니라 `str(.Value)`로 저장돼 정수가 `1.0`이 됩니다.
- **excel_fill**: append 모드에서 헤더만 있는 CSV를 넣으면 `max()` 오류가 납니다. append와 `--theme`를 같이 쓰면 첫 데이터 행에 헤더 서식이 칠해집니다.
- **excel_chart**
  - 제목으로 삭제할 때 일치하는 차트가 없어도 1번 차트를 지우고 ok를 반환합니다.
  - CLI는 `--range`를 필수로 요구하는데 `execute`는 범위를 자동 감지합니다. 둘이 어긋납니다.
  - 꺾은선형 차트 테마에 `Fill`을 써서 색이 반영되지 않습니다. `Format.Line`을 써야 합니다.
- **excel_range**
  - `'My Sheet'!A1`처럼 따옴표가 붙은 시트명을 처리하지 못합니다.
  - 다른 시트의 셀로 freeze를 하면 COM 오류가 납니다.
  - 범위 없이 freeze하면 화면 중간 위치에 고정됩니다.
- **excel_sheet**
  - 시트명 대소문자 처리 때문에 `ensure`와 `copy` 실패 시 "Sheet4"나 "X (2)" 같은 고아 시트가 남습니다.
  - `DisplayAlerts`를 finally에서 복원하지 않습니다.
  - 기본 action이 `execute`와 CLI에서 서로 다릅니다.
- **excel_inspect**: 시트명을 대소문자 구분해서 비교하고, 활성 문서 판정 방식이 모드마다 다릅니다.

### 4.3 Word
- **word_write.py:311 `matrix` UnboundLocalError (중요)**: 일반 텍스트, heading, bullet 쓰기가 **항상 error를 반환합니다**. 210행 앞에 `matrix = None`을 추가합니다.
- **word_export.py:105 `fixed_err_msg`**: NameError가 진짜 COM 오류를 가립니다. 또 `-2147467259`(E_FAIL, 일반 실패 코드)를 "PDF 추가기능 없음"으로 오진합니다. `ppt_export`도 같습니다.
- **word_inspect.py:89**: `Information(14)`는 행 번호 상수입니다. 열 번호는 `16`(wdStartOfRangeColumnNumber)입니다.
- **word_edit**
  - `--replace ""`(텍스트 삭제)가 무시됩니다. ppt_edit도 같습니다.
  - Find 옵션(MatchWildcards, MatchCase, MatchWholeWord)을 명시하지 않아 사용자의 마지막 Find 설정이 그대로 남습니다.
  - `"blue": 9`는 wdDarkBlue입니다. wdBlue는 2입니다.
  - 5000건 상한에 걸려도 VERIFIED를 반환합니다.
- **word_write**
  - `pos=after`가 문단 중간에 단락 기호를 넣어 문단을 쪼갭니다.
  - markdown 렌더러가 anchor를 무시합니다.
  - 제목에 Heading 스타일 대신 직접 서식을 써서 목차(TOC)와 검색이 제목을 인식하지 못합니다.
  - data_file 경로를 CWD 기준으로만 찾습니다.
- **word_read**
  - COM 경로에서 search와 format을 무시합니다.
  - `.doc`는 BadZipFile 오류가 나고 대체 경로가 없습니다.
  - 중첩 표의 셀이 상위 표에 섞입니다.
  - `w:sdt`(콘텐츠 컨트롤) 안의 문단이 누락됩니다.
- **word_export**: `--format pdf -o x.docx`처럼 형식과 확장자가 다르면 PDF가 `.docx` 이름으로 저장됩니다.

### 4.4 PowerPoint
- **ppt_read**
  - 슬라이드 순서를 `slideN.xml` 파일 번호로 정해서 순서를 바꾼 뒤에는 번호가 틀립니다. `presentation.xml`의 `sldIdLst`를 써야 합니다.
  - `<a:t>` 런 단위로 `\n\n`을 넣어 서식이 바뀌는 지점에서 단어가 쪼개집니다. `<a:p>` 단위로 합쳐야 합니다.
  - `total_slides`가 텍스트가 있는 슬라이드 수만 셉니다.
- **ppt_write**
  - 기본값 판정용 숫자(`320`)가 실제 기본 높이와 달라 이미지 가로세로 비율이 깨지고, 명시한 `--width 620`도 기본값으로 취급됩니다.
  - 범위 밖 `--slide`를 조용히 마지막 슬라이드로 바꿉니다.
  - 75행 초과 표를 잘라내고도 ok를 반환합니다.
  - 요소가 겹치게 배치되고, 제목은 620pt 고정폭입니다.
- **ppt_slide**: `CustomLayouts(7)`이 기업 템플릿에서는 빈 레이아웃이 아닐 수 있습니다.
- **ppt_session 종료 시 `GotoSlide(마지막)`**: 읽기나 진단만 해도 사용자 화면이 마지막 슬라이드로 이동합니다. 쓰기 작업이 실제로 수정한 슬라이드로만 이동하게 합니다.
- **ppt_inspect**: 슬라이드 제목을 `Shapes.Title`이 아니라 처음 발견한 텍스트로 잡습니다.

### 4.5 office_bridge
- excel->excel이 `to_target`을 무시하고, 원본 범위(`range`)를 붙여넣을 위치로도 써서 **원본 위에 붙여넣습니다**.
- CLI의 `--copy-as` 기본값 "table"이 PPT/차트용 자동 기본값("picture")을 항상 덮어씁니다. 기본값을 None으로 바꿉니다.
- ppt->word에서 bookmark 위치를 무시하고 항상 문서 끝에 붙여넣습니다.
- 차트 이미지를 기본 500x300으로 넣어 원본 비율이 깨집니다.
- Word/PPT 표를 셀 단위로 추출하는데, 병합 셀은 빈칸으로 처리됩니다.

### 4.6 Outlook, EML
- **outlook_read 검색**
  - `@SQL=` 접두어를 절마다 붙이고 괄호 없이 AND로 이어 **query와 from을 같이 쓰면 필터가 깨집니다**.
  - 실패 시 대체 경로가 정확히 일치하는 제목만 찾고 sender 조건을 버립니다. from만 준 경우에는 필터 없는 최신 메일을 "일치"로 반환합니다.
  - `Restrict` 다음에 Sort를 해야 정렬이 보장됩니다.
- **outlook_read 첨부 저장**: 같은 이름의 파일을 덮어쓰고, 서명 이미지까지 저장하며, 첨부 하나에서 예외가 나면 전체가 중단됩니다.
- **outlook_write**
  - HTML 접두부를 `<body>` 안이 아니라 `<html>` 앞에 넣습니다.
  - `Display` 없이 `Body`를 지정하면 사용자 서명이 사라집니다.
- **_outlook_guard**: 연결 함수가 `Dispatch`를 써서 Outlook이 꺼져 있으면 실행시킵니다(문서 설명과 다름).
- **eml_read**
  - `thread` 액션이 JSON 앞에 여러 줄 텍스트를 출력해 **1줄 JSON 규약을 깹니다**.
  - `extract`와 `thread`의 기본 입력 `eml_master.json`이 `build`의 출력 이름(`eml_master_<name>.json`)과 달라 기본값으로는 찾지 못합니다.
  - charset 대체(fallback)가 실제로는 동작하지 않습니다. `errors='replace'`로 먼저 디코딩해 버리기 때문입니다.
  - 전달된 첨부 메일(`message/rfc822`)의 본문이 메인 본문에 섞입니다.
  - 날짜를 문자열로 정렬하고 시간대를 무시합니다.

### 4.7 PDF, HWP, search, 기타
- **hwp_read 제어문자 파싱 (중요)**: HWP 5.0 규격 기준으로 세 가지가 틀렸습니다.
  - 확장/인라인 컨트롤(8 WCHAR)을 건너뛸 때 14바이트가 아니라 16바이트를 건너뛰어 **컨트롤 뒤 한 글자가 사라집니다**.
  - 코드 1~3(확장 컨트롤)은 코드 2바이트만 건너뛰고 나머지 14바이트를 남겨 `dces`, `dloc` 같은 쓰레기 문자가 섞입니다.
  - 코드 24~31(1 WCHAR 문자 컨트롤: 하이픈, 묶음빈칸 등)을 확장 컨트롤로 처리해 실제 글자 8개를 지웁니다.
- **hwp_read 기타**
  - `Section10`이 `Section2`보다 앞에 오도록 문자열로 정렬합니다. 숫자로 정렬해야 합니다.
  - HWPX를 `<hp:t>` 단위로 쪼개 문단이 조각나고 공백이 사라집니다.
  - 제목 판정 정규식이 "이 사업은…" 같은 일반 문장까지 제목으로 잡습니다.
  - 배포용, 암호화 문서를 빈 결과의 ok로 반환합니다.
- **pdf_read**
  - pymupdf가 실패하거나 빈 결과를 내면 pypdf 결과가 같은 리스트에 이어 붙어 페이지가 중복됩니다.
  - 스캔 PDF는 시간 제한 없이 Word COM 대체 경로로 넘어갑니다.
- **search**
  - 최상위가 리스트인 JSON 캐시 파일 하나 때문에 AttributeError로 종료됩니다.
  - 질의어가 우연히 기존 파일명과 같으면 검색 대신 가짜 매칭을 반환합니다.
  - `--to` 전송이 실패해도 ok를 반환합니다.
- **clean**: `fix-env`가 실제로는 아무것도 하지 않으면서 "VERIFIED"를 반환합니다.
- **office_theme**
  - 읽기 전용이어야 할 `extract`가 전역 `current_theme.json`을 **다른 형식(스키마)으로 덮어씁니다**. 이후 `resolve_theme("auto")`를 쓰는 도구가 깨집니다.
  - 헤더에 채우기가 없으면 `#1B365D`를 지어내서 기록합니다.
- **office_window**
  - `--focus` 플래그가 실제로 전달되지 않습니다.
  - excel/word가 아닌 이름은 전부 PPT로 매핑합니다.
  - `ShowWindow(9)`가 최대화된 창을 원래 크기로 되돌립니다.
  - `FindWindow`가 해당 문서가 아닌 첫 번째 창을 잡습니다.

---

## 5. Phase 3: 공통화와 성능 최적화

### 5.1 구조 공통화 (중복 제거)
| 대상 | 현재 | 목표 |
|---|---|---|
| Excel/Word/PPT 세션 | 약 90% 같은 코드 3벌(약 600줄) | `OfficeSession` 기반 클래스와 앱별 어댑터(문서 컬렉션 이름, 창 클래스, 저장 방식) |
| 앱 탐색 | `find_all_native_*` 3벌, `_get_*_app_from_hwnd` 3벌 | `find_office_apps(app_kind)` 하나 |
| 경로 해석 | `_config.resolve_data_path`와 별도 resolver 5개, `C:\cli_data\data` 하드코딩 3곳 | `resolve_input_path` / `resolve_output_path` 하나. 검색 순서를 일관되게 하고 하드코딩은 config로 옮김 |
| 응답과 오류 | 모듈마다 error 응답 복사, `main`/`execute` 중복 | `@tool_entry` 데코레이터: 인자 파싱 -> execute -> 예외를 JSON으로 -> 종료 코드 |
| Range를 행렬로 변환 | 4곳 (모두 버그) | `range_to_matrix()` |
| 테마 표 서식 | `excel_edit`와 `excel_fill`에 2벌 (로직 다름) | `excel_edit.apply_theme_to_table` 하나 |
| 마크다운 표 파서, 출력 경로, 대상 파일 해석 | word/ppt에 똑같은 함수 3쌍 | `_office_utils`로 이동 |
| 도형 순회(그룹, 표, 텍스트) | ppt_edit, ppt_inspect, ppt_read 3벌 | `iter_text_frames(slide)` |
| 키워드 스니펫 | pdf와 hwp에 같은 코드 | `_text_search.snippets()` |
| CSV 쓰기와 인코딩 대체 | excel_export에만 4벌 | `write_csv(rows, path, encoding)` |

### 5.2 응답 규약(JSON contract) 정리
- `status` 값을 `ok | no_match | error | partial` 네 가지로 고정합니다. 잘라냈거나 상한에 걸린 결과는 `partial`로 하고 VERIFIED를 붙이지 않습니다.
- `agent_directive`를 오류 성격에 따라 나눕니다.
  - 일시적 오류(RPC 중단, Excel 사용 중 0x800AC472): `RETRY_ONCE`
  - 입력 오류: `FIX_ARGS`
  - 그 외: `STOP_RETRY`
- `action`은 항상 CLI에서 받은 action을 쓰고, 세부 동작은 `sub_action` 필드로 분리합니다.
- 읽기 결과에서 `items`와 `data`가 같은 내용을 중복으로 담는 문제를 없앱니다. 기본은 `items`만 주고 `data`는 `--format text`일 때만 줍니다. LLM에 들어가는 토큰이 절반 가까이 줄어듭니다.
- 미리보기(preview) 언어를 하나로 통일합니다(권장: 한국어).

### 5.3 COM 성능
| 위치 | 현재 | 개선 |
|---|---|---|
| 쓰기 세션 전반 | `ScreenUpdating=True`를 강제하고 Calculation, Events를 그대로 둠 | 쓰기 블록 동안 `ScreenUpdating=False`, 수동 계산(`xlCalculationManual`), `EnableEvents=False`로 두고 finally에서 복원 |
| excel_read 검색, excel_edit 규칙, excel_inspect | 셀마다 COM 호출 3~4회 | `.Value`를 한 번에 읽어 Python에서 처리하고, 해당 셀만 쓰기. 규칙은 `FormatConditions.Add`(진짜 조건부 서식) 옵션 제공 |
| excel_fill 수식 | 셀 단위로 두 번째 패스 | `Range.Formula = 2D tuple` 한 번 |
| 테마 행 서식 | 행마다 COM 3회 | 줄무늬는 `Union`이나 조건부 서식(`MOD(ROW(),2)`)으로 일괄 처리 |
| word_edit 단순 치환 | 매칭마다 Execute와 속성 설정 | 서식이 없는 치환은 `Execute(Replace=wdReplaceAll)` 한 번 |
| word_read COM 경로 | `doc.Paragraphs(i)` 인덱스 접근(O(n²))과 문단마다 `Information(3)` | `for p in doc.Paragraphs` 순회, 페이지 정보는 필요할 때만 조회 |
| Word/PPT 표 생성 | 셀마다 COM 3~6회 | 텍스트를 탭으로 구분해 `ConvertToTable`로 변환하고 행/열 단위로 서식 지정 |
| word_inspect | `Characters.Count`, `Words.Count` | `ComputeStatistics` |
| ppt_edit | 규칙 × 도형마다 `tr.Text`를 다시 읽음 | 프레임마다 텍스트를 한 번만 읽고 Python에서 규칙을 미리 걸러냄 |
| office_theme | 열 수 × 8회 COM | `UsedRange`를 실제 데이터 범위로 줄이고 필요한 속성만 조회 |

### 5.4 캐시와 I/O
- **캐시 파일명에 메타데이터 넣기**: `cache_{pathhash}_{mtime}_{size}_{subkeyhash}.json` 형식으로 바꾸면 glob만으로 적중 여부를 판정할 수 있습니다. 지금은 조회할 때마다 후보 JSON을 전부 열어 파싱합니다.
- **보관 정책**: "파일당 2개"가 sub_key(시트, 범위, 페이지 등)가 달라도 서로 밀어내 캐시가 제 역할을 못 합니다. 파싱 원본은 1개만 두고 질의별 결과는 캐시하지 않는 방식으로 바꿉니다. 원본만 있으면 page/search/format은 Python에서 바로 계산할 수 있습니다.
- `indent=2`를 제거하고 원자적으로 씁니다(임시 파일에 쓴 뒤 `os.replace`).
- **search 전역 검색**: 매 질의마다 모든 캐시 JSON을 읽습니다. 경량 색인 `index.json`(파일 경로, 종류, 캐시 파일명)을 따로 두고 중복을 먼저 제거합니다.
- **search 폴더 모드**: 깊이와 파일 수 상한, 숨김/시스템 폴더 제외를 두고, COM이 필요한 형식(.doc/.xls/.ppt)은 옵션을 켰을 때만 처리합니다.
- **_config**: `load_config()`가 호출될 때마다 파일을 다시 읽습니다. 프로세스 단위로 메모이즈하고, 사용하지 않는 recipe 함수는 제거합니다.

---

## 6. Phase 4: 고도화

1. **안전 정책 계층** (`_safety.py`)
   - 쓰기 경로 화이트리스트(`data_dir`, 사용자가 지정한 출력 폴더)를 둡니다.
   - 메일 발송과 파괴적 작업(시트/슬라이드 삭제, clear, undo)은 `--confirm`이 있어야 실행합니다. 에이전트 도구 스키마에서는 발송 기능을 아예 노출하지 않습니다.
   - Outlook 첨부 저장은 별도 폴더에 저장하고, 파일명이 겹치면 번호를 붙이고, 실행 파일 확장자는 차단합니다.
2. **Undo 고도화**
   - 스냅샷을 작업 단위(transaction id)로 묶고, `clean --action undo --list`로 목록을 조회하고 선택해서 복원할 수 있게 합니다.
   - 열린 문서는 COM으로 다시 불러와 복원합니다. 디스크 덮어쓰기는 문서가 닫혀 있을 때만 합니다.
3. **도구 스키마 단일화**
   - 각 모듈이 `TOOL_SPEC`(이름, 설명, JSON Schema)을 직접 내보내게 합니다.
   - 에이전트(`atomic_agent.py`)는 이 스펙을 모아 Function Calling 정의를 자동 생성합니다. README의 파라미터 표도 같은 스펙에서 생성해 문서와 코드가 어긋나지 않게 합니다.
4. **파서 품질**
   - HWP: FileHeader의 압축/암호화/배포용 플래그를 해석하고, 표와 셀 텍스트를 추출합니다. HWPX는 `<hp:p>` 단위로 읽습니다.
   - DOCX/PPTX: 머리글/바닥글, 각주, 노트, 콘텐츠 컨트롤을 추출하고 실제 슬라이드 순서를 반영합니다.
   - PDF: 페이지 수 상한을 두고, 스캔본은 OCR이 필요하다고 알리기만 하고 Word 대체 경로는 옵션으로 돌립니다.
   - EML: charset 감지(엄격 디코딩 -> cp949 -> euc-kr -> utf-8 순)를 고칩니다. 스레드는 `In-Reply-To`/`References` 헤더 기준으로 묶습니다.
5. **가시 위치 보정**: ppt_write의 Smart Stacking에 슬라이드 경계 검사, 16:9 너비 반영, 제목 placeholder 사용을 추가합니다.
6. **관측성**: `OFFICE_CLI_DEBUG=1`일 때만 stderr에 디버그를 출력합니다. 지금은 항상 출력합니다. 도구별 실행 시간과 COM 호출 수를 JSON의 `meta` 필드에 넣습니다.
7. **README 전면 재작성**
   - 실제 존재하는 모듈과 인자 기준으로 다시 씁니다.
   - "0.001초", "100% 보장" 같은 측정되지 않은 수치는 실제 측정값이나 정성적 설명으로 바꿉니다.
   - `config.json` 예시에서 개인 경로와 API 키 자리는 일반 예시로 바꿉니다.

---

## 7. 진행 방식과 검증

- **작업 단위**: Phase 1은 항목마다 커밋 1개로 나눕니다. Phase 3은 공통 모듈을 먼저 도입한 뒤 모듈을 하나씩 옮깁니다(한 번에 1~3개 파일).
- **검증 방법**
  1. ruff와 compileall을 통과해야 합니다.
  2. 오프라인 단위 테스트(순수 로직과 가짜 COM 객체)를 통과해야 합니다.
  3. Windows 실기 스모크 스크립트를 사용자가 실행하고 결과 JSON을 공유합니다.
- **반입**: 개별 파일로 합니다. merged_code.md와 merge 도구는 사용하지 않습니다.

## 8. 확정 사항 (2026-10-27 사용자 회신)

1. **범위**: merged_code.md에 없는 파일(`atomic_agent*.py`, `ml_*.py`, `excel_style.py`)은 작업 대상에서 제외합니다. README 재작성(Phase 4) 때 해당 절도 정리합니다.
2. **작업 방식**: 저장소의 개별 `.py`로 작업하고, 반입도 개별 파일로 합니다. (merged_code.md와 merge 도구는 삭제)
3. **대상 환경**: pywin32 312, pywin32-ctypes 0.2.3, Office 16.0 (Word/Excel/PowerPoint/Outlook, 2016/2019/2021/365 계열)
   - `SaveAs2`, `WordOpenXML`, `ExportAsFixedFormat`, `DispatchEx`를 모두 쓸 수 있습니다.
   - pywin32 312에서 Enum 콜백이 False를 반환하면 예외가 날 수 있으므로, 콜백은 항상 True를 반환하도록 바꿨습니다.
4. **메일 발송 정책**: 기본은 작성 창만 띄우고 사용자가 [보내기]를 누릅니다. 사용자가 특별히 요청한 경우에만 `--send`(또는 `send=true`)로 발송합니다.
   - 발송 전에 수신자 주소를 확인합니다. 확인되지 않은 수신자가 있으면 발송하지 않고 작성 창으로 전환합니다(`status=partial`).
   - `"false"` 같은 문자열은 발송 요청으로 인정하지 않습니다.

## 9. 진행 현황

| 단계 | 상태 | 비고 |
|---|---|---|
| Phase 0 | 완료 | `.py` 복원, `pyproject.toml`(ruff, pytest), `tests/`(가짜 COM 기반 29개). merge 도구는 이후 삭제 |
| Phase 1 | 완료 (Windows 실기 검증 대기) | 아래 표 참고 |
| Phase 2 | 일부 선반영 | `word_write` matrix, `word_export` fixed_err_msg, `excel_edit`의 한 행 범위 처리, 창 탐색 콜백, 응답 예약 키 보호 |

Phase 1 처리 내역:

| # | 처리 |
|---|---|
| 1 | Excel/Word 파일 모드를 `DispatchEx` 전용 인스턴스로 바꾸고, 종료 시 전용 인스턴스만 Quit합니다. PPT는 사용자 인스턴스의 Visible/WindowState를 바꾸지 않고, DisplayAlerts는 복원합니다 |
| 2 | Word 스냅샷: 저장된 문서는 파일을 복사하고, 미저장/새 문서는 `WordOpenXML`(.xml)로 남깁니다. 원본 Save/SaveAs를 제거했습니다 |
| 3 | 활성 PPT도 스냅샷을 생성합니다 |
| 4, 6 | 세션에 `readonly` 인자를 추가했습니다(기본: `snapshot=False`이면 읽기 전용). 읽기 전용은 시트/파일을 만들지 않고, 저장하지 않으며, 활성 시트·스크롤·슬라이드 위치를 바꾸지 않습니다. 시트명은 대소문자를 구분하지 않습니다 |
| 5 | PPT는 변경이 있을 때만 저장합니다. 새 파일은 지정 경로로 `SaveAs`합니다 |
| 7, 8 | `word_export`는 사본 문서에서 페이지 번호와 DOCX 저장을 처리해 원본을 바꾸지 않습니다. 형식과 확장자 불일치를 보정하고, E_FAIL 오진을 제거했습니다(`ppt_export`도 동일) |
| 9 | `excel_range move`를 `Cut(Destination)`으로 바꾸고, 따옴표 시트명을 해석합니다 |
| 10 | `excel_edit`는 수식 셀의 내용을 보존하고 서식만 적용합니다 |
| 11 | `excel_export` XLSX는 try/finally로 임시 통합문서를 정리하고 경고 설정을 복원합니다 |
| 12 | `clean undo`: 복원 전 상태를 `cache/undo_backup_*`에 백업합니다. 열린 문서는 닫은 뒤 교체하고 다시 엽니다. 실패하면 error를 반환하고, 성공한 스냅샷은 소비합니다. `fix-env`는 실제 동작(COM 응답 점검)대로 보고합니다 |
| 13 | `outlook_write` 발송 정책 적용(위 8-4). 서명을 유지하고, HTML은 `<body>` 안에 삽입합니다 |
| 14 | `ppt_slide clear`는 `--slide` 또는 `--all-slides`가 있어야 동작합니다 |
| 15 | `excel_range merge`는 범위가 필수이고, 경고 대화상자를 잠시 끈 뒤 복원합니다 |
| 공통 | 읽기 도구 `main()`은 `run_session_tool`로 세션 오류도 JSON 1줄로 내보내고, 종료 코드는 0/1입니다 |

Windows 실기 확인을 부탁드리는 항목은 다음과 같습니다.
- Excel/Word 창을 띄운 상태에서 `excel_read.py --target <다른 파일>`과 `word_read.py --target <다른 파일>`을 실행해도 사용자 창이 숨겨지지 않는지
- `clean.py --action undo`가 열린 Excel/Word/PPT 문서를 복원하는지
- `outlook_write.py --action new`가 작성 창만 띄우는지(서명 포함)
