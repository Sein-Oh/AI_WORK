# office-cli Windows 실기 테스트

## 실행
`CLI-dev/office-cli` 폴더(사내 PC `C:\CLI-dev\office-cli`)에서 실행합니다. 시험 대상 도구는 옆 폴더 `CLI/office-cli`입니다.
```bat
python tools/run_win_tests.py
python tools/run_win_tests.py --self-email 본인주소     :: G3 실제 발송 테스트 포함
python tools/run_win_tests.py --only H                  :: 일부 그룹만
python tools/run_win_tests.py --only L,M,N              :: 2026-11-09 검토 실기 확인과 1.6.0 새 기능 (M·N7·N8은 Outlook, 질문 없음)
python tools/run_win_tests.py --skip-outlook --auto     :: Outlook 제외, 질문 없이 자동 판정 항목만
python tools/run_win_tests.py --agent --skip-outlook    :: 에이전트용: --auto + 로그는 stderr, stdout은 결과 JSON 1줄
```
- 테스트 파일(A/B .xlsx .docx .pptx)은 `C:\CLI-dev\work\office`에 자동으로 만듭니다(`--dir`로 변경 가능).
- 실행 전에 다른 Office 문서는 저장하고 닫아 주세요. 실행 중에는 Office 창을 직접 만지지 마세요.
- 결과는 테스트 폴더에 두 파일로 저장됩니다.
  - `report_*.txt`: 항목별 판정 요약.
  - `errors_*.txt`: 분석용 상세 로그. 환경(Windows, Python, 라이브러리, Office 버전, `config.json`), 실패/오류 항목마다 실행한 명령 전체와 종료 코드, 걸린 시간, stdout JSON 전체, stderr 전체, traceback, 그 항목의 진행 로그를 담습니다. 준비 단계에서 멈추거나 Ctrl+C로 중단해도 그 시점까지 기록합니다. **문제가 있으면 이 파일을 편집하지 말고 그대로 Claude에게 전달해 주세요.**
- OpenCode 에이전트에게 맡기기: `C:\CLI-dev\office-cli\tools\office-cli-test` 폴더를 스킬 폴더의 `office-cli` 옆에 복사한 뒤 "office-cli 실기 테스트 돌려줘"라고 요청합니다. 에이전트는 Office 문서를 닫았는지 먼저 묻고, `--agent --skip-outlook`으로 실행해 결과와 `errors_*.txt` 경로를 알려 줍니다. `--agent`에서는 눈 확인 항목이 `MANUAL-SKIP`이 됩니다. 결과 JSON의 `unverified_count`는 도구가 결과를 다시 읽어 확인하지 못한 호출 수이며, 1 이상이면 `agent_directive`가 `UNVERIFIED:`로 시작합니다(통과 아님). `-h`와 인자 오류(없는 그룹 포함, 인자 오류는 exit 2)도 JSON 1줄이고, Ctrl+C는 exit 130입니다.
- 도구의 캐시(스냅샷, 색인)와 data 폴더는 테스트 폴더의 `_cache`, `_data`를 쓰고, 캐시는 실행마다 비웁니다. 그래서 여러 번 실행해도 undo 결과가 이전 실행에 섞이지 않고, 도구 폴더의 `cache`/`data`도 건드리지 않습니다. (`config.json`의 `--probe` 결과는 도구 폴더에 그대로 저장됩니다.)
- 끝나면 확인용 그림(A.pptx 슬라이드, A.xlsx 차트)을 테스트 폴더의 `shots\`에 저장하고, 테스트 문서는 저장하지 않고 닫습니다. 열어 두려면 `--keep-open`.
- 판정: `PASS` / `FAIL` / `MANUAL-SKIP`(`--auto`라 눈 확인 생략) / `ENV-SKIP`(이 PC에 기능이 없음).

## 그룹
| 그룹 | 확인 내용 |
|---|---|
| A | 닫힌 파일 읽기는 Office를 띄우지 않음. 닫힌 파일에 쓸 때 사용자 창이 닫히거나 숨겨지지 않음. 다중 Excel 프로세스 목록 |
| B | 읽기 전용: 없는 시트/파일을 만들지 않음, 활성 시트 유지, 시트명 대소문자 무시, PPT 읽기 후 화면 슬라이드 유지 |
| C | 되돌리기: 엑셀/워드/PPT 쓰기 후 `office.py undo`, 연속 undo, 저장 안 한 새 통합문서, 원본 자동 저장 안 됨 |
| D | 내보내기: 워드 PDF 쪽 번호(원본 미변경), DOCX 사본, 엑셀 XLSX, PPT PDF |
| E | 엑셀 편집: 범위 이동, 병합, 따옴표 시트명, 수식 보존 치환, 테마 |
| F | PPT: clear 보호, 특정 슬라이드 비우기, 새 슬라이드 쓰기와 화면 이동 |
| G | Outlook: 기본은 초안 저장 + 작성 창만, 잘못된 수신자 초안은 send해도 발송 안 함, 초안 send 발송(두 번째 send는 거부), 목록 |
| I | 기능 확대(자동 판정 27개): 엑셀 쓰기·읽기·차트·서식·시트·범위, 워드 그림·표·앵커·제목·목차·치환, PPT 그림·카드·마크다운·긴 표·슬라이드 관리·치환·사본, 범위/표 전달, redraw, undo 선택·닫힌 파일, 최대화 유지. 결과는 COM과 파일로 다시 읽어 판정합니다 |
| J | 배치 점검(`info --detail`): PPT 슬라이드 밖/넘침/겹침, 데이터를 가리는 엑셀 차트, 쪽보다 넓은 워드 그림이 없음 |
| K | 검토 기능: 워드 `outline`(수준·쪽)과 `read --section`(하위 절 포함), `read --pages`(Word 쪽 번호와 같음)와 `--paragraphs`, 기준 글에 메모 달기와 메모 목록, 변경 추적 목록(삽입), PPT 노트 쓰기·덧붙이기·읽기 |
| L | 2026-11-09 검토 실기 확인(엑셀·워드·PPT, 1.4.94~1.4.96): L1 통화 서식 열 정렬·`--where`(A7), L2 `+82-...`·`- 검토` 글자 그대로(A8), L3 `=합계()`는 #NAME?로 UNVERIFIED·`=SORT()` 펼침(A9), L4 표 바로 위 단락 뒤에 쓴 글이 표 첫 칸에 안 들어감(C1), L5 새 번호 목록은 1부터(C2), L6 표 뒤 표가 합쳐지지 않음(C3), L7 제목 뒤 그림 단락은 본문 수준(C4), L8 `--toc --file` 목차에 제목(C5), L9 변경 추적 문서 replace(C6, 확인만), L10 바닥글 개체틀 슬라이드의 글상자 위치(C8), L11 제한된 보기 쓰기(C9, 러너가 인터넷 영역 표시를 붙인 파일을 제한된 보기로 열어 자동), L12 열린 통합문서 undo의 디스크 백업(C10), L13 필터 숨긴 행 `--append`(C11), L14 닫힌 파일 freeze 저장(C12, 확인만). C13 OneDrive는 쓰지 않아 뺌 |
| M | 2026-11-09 검토 실기 확인(Outlook, `--skip-outlook`이면 빠짐, 모두 러너가 판정): M1 받은 시각(C14), M2 `current`는 맨 앞 창(C15), M3 `search --from <SMTP>`(C16, `--self-email` 필요), M4 대화 상자가 떠 있을 때 작성(C17, 러너가 옵션 창을 띄우고 닫음), M5 본문 그림(C18) |
| N | 1.6.0 새 기능(질문 없음): N1 수식을 값으로, N2 열린 통합문서 저장·다른 이름으로 저장, N3 중복 찾기, N4 워드 단락 지우기, N5 워드 표 행 넣기·지우기, N6 PPT 도형 글 바꾸기, N7 일정 만들기·지우기(1.6.3: 도구로 지움), N8 메일 `--until`·`--with-attachments` (N7·N8은 Outlook, `--skip-outlook`이면 빠짐). 1.6.1: N9 PPT 도형 옮기기·크기, N10 워드 메모 답글·지우기, N11 워드 단락 서식. 1.6.3: N12 BOM 비교(새 판에 비교결과 시트를 묻지 않고 넣음, 메모는 `confirm_commands[0]`) |
| H | 체이닝: 사인 2주기 그래프 -> `transfer --probe` -> 차트를 워드/PPT에 그림으로 -> 워드/PPT 내용을 엑셀로 -> 문서 검색 결과를 워드에 |

H2(`--probe`)는 이 PC에서 되는 차트 전달 방식(export / redraw / relay)을 저장합니다. 회사 PC에서 꼭 한 번 실행해 주세요.

## 수동 확인용 명령 (러너 없이 시험할 때)
```bat
python excel.py write --range A1:B101 --header "x,sin(x)" --formula "=(ROW()-2)*4*PI()/99" "=SIN(A2)"
python excel.py chart --range A1:B101 --type line --title "사인 2주기"
python office.py transfer --from excel --chart "사인 2주기" --to word,ppt
python office.py transfer --from word --what text --to excel
python reader.py read --file "실제문서.hwp" --format summary
python reader.py search --query "검색어" --file "문서폴더" --save
python office.py undo --list
```
