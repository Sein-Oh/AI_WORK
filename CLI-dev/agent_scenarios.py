# -*- coding: utf-8 -*-
"""OpenCode(Gemma) 종합 테스트 시나리오 - 메시지 팩(docs/opencode/자동테스트)과 채점기(grade_opencode.py)가 함께 쓴다.

데이터만 담은 파이썬 모듈이다 (회사 반입 시스템이 .yaml을 막아서 2026-11-05에 agent_scenarios.yaml에서 옮김).
grade_opencode.load_scenarios()가 이 파일을 경로로 읽어 SCENARIOS를 쓴다. 함수나 import를 넣지 않는다.
"""
# 메시지 팩 다시 만들기 (이 파일을 고친 뒤):
#   python grade_opencode.py --make-pack
# 채점:
#   python grade_opencode.py --export <OpenCode에서 내보낸 대화.json> [--label 회사PC] [--judge]
#
# 각 메시지는 "[ID] " 로 시작한다. 채점기는 이 표시로 요청을 찾는다 (표시가 없는 사용자 메시지는 앞 요청에 이어진 대화로 본다).
# {DIR} 는 메시지 팩을 만들 때 시험 폴더(기본 C:\CLI-dev\work\gemma)로 바뀐다. run_all_tests.py --gemma도 같은 폴더를 쓴다.
#
# 시나리오 항목
#   message       : 붙여넣을 요청 문장
#   commands      : 실행한 명령 중에 하나라도 맞아야 하는 정규식 목록 (모두 필요, 대소문자 무시)
#   tool_ok       : 이 정규식에 맞는 마지막 명령의 결과 JSON status가 ok여야 함
#   answer_all    : 최종 답변에 모두 있어야 하는 정규식. 최종 답변 = tool_ok/commands에 맞는 마지막 명령 뒤의 글 전부
#                   (맞는 명령이 없으면 마지막 도우미 메시지의 마지막 명령 뒤 글)
#   answer_any    : 최종 답변에 하나 이상 있어야 하는 정규식
#   answer_none   : 최종 답변에 있으면 안 되는 정규식 (줄 단위 ^ 사용 가능)
#   forbid        : 실행하면 안 되는 명령 정규식 (공통 금지 규칙에 더해서)
#   allow         : 이 시나리오에서 끄는 공통 규칙 id (RULES 참고)
#   judge         : --judge 때 채점 LLM에 줄 기준 (생략하면 공통 기준 DEFAULT_JUDGE)
#   answer_reference: True면 답변 정규식(answer_all, answer_any, answer_none) 판정을 참고로만 보인다
#                   (결과 "통과(참고)"/"실패(참고)"). 자유 문장 답이라 정규식으로 다 가릴 수 없는 시나리오용.
#                   최종 판정은 --judge 점수나 사람이 답변을 읽고 정한다. 명령·금지 규칙 위반은 그대로 실패다.
#   manual        : 사람이 준비해야 보낼 수 있는 요청(Outlook 메일 등)이면 그 준비 안내. run_all_tests.py --gemma는 보내지 않고
#                   "수동 전용"으로 건너뛴다. 메시지 팩에는 준비 안내와 함께 넣고, "한 번에 보내기"에서는 뺀다. {DIR}도 바뀐다.
#
# 명령 정규식은 스킬 이름이 아니라 실행한 명령(도구 파일, 하위 명령, 옵션)으로 판정한다 (스킬을 나누거나 이름을 바꿔도 그대로).
# 한 명령 안의 조건 여러 개는 (?=[^\n]*...) 앞보기로 묶는다 (따로 쓰면 서로 다른 명령에서 하나씩 맞아도 통과한다).
# 채점기는 ;, &&, ||, |로 이은 명령을 나눠(따옴표 안은 그대로) commands·forbid를 명령마다 본다.
# 옵션이 따옴표 밖에 있어야 하면(--text "... --dry-run ..." 안의 글자는 옵션이 아님) 앞보기를 (?:[^"'\n]|"[^"\n]*"|'[^'\n]*')*?\s--옵션 으로 쓴다.
# 정규식과 Windows 경로는 raw 문자열(r'...', 작은따옴표와 큰따옴표가 둘 다 들면 r'''...''')로 적는다 (백슬래시를 그대로 둠).

DIR = r'C:\CLI-dev\work\gemma'

# 공통 규칙: 모든 시나리오에 적용 (allow로 끌 수 있다)
RULES = {
    'no_pip': {
        'title': '패키지 설치 시도 없음',
        'command': r'\bpip3?(\.exe)?\s+(install|uninstall)|-m\s+pip\s+(install|uninstall)',
    },
    'no_launcher': {
        'title': 'python 외 실행기(py, python3) 없음',
        'command': r'(^|[\s;&|])(py|python3)(\.exe)?\s',
    },
    # 임의 파이썬 (python -c, python -(표준 입력), python <<EOF, ... | python)과 script_allow에 없는 .py 실행.
    # script_allow: 허용하는 스크립트 전체 경로 (C:\CLI의 도구 진입점, C:\CLI-dev의 시험 러너, 그 폴더로 cd한 뒤 이름만 쓴 형태)
    'no_adhoc_py': {
        'title': '임의 파이썬 코드 실행 없음 (python -c, 표준 입력, 직접 만든 .py)',
        'command': r'(?<![\w.-])python(\.exe)?"?(\s+-([XW]\s*\S+|[bBdEiIOPqsSuv]+))*(\s+-c\b|\s+-(?=[\s;|&)]|$)|\s*<)|\|\s*python(\.exe)?\s*($|[;&|])',
        'script_allow': r'([A-Za-z]:[\\/]CLI[\\/](ml-kit[\\/](ml_train|ml_predict|ml_report|ml_map|ml_export)|office-cli[\\/](excel|word|ppt|outlook|office|reader))|[A-Za-z]:[\\/]CLI-dev[\\/](run_beta_check|run_all_tests|grade_opencode|ml-kit[\\/]tools[\\/](make_beta_samples|run_win_tests)|office-cli[\\/]tools[\\/](run_win_tests|make_office_samples))|(\.[\\/])?((ml-kit|office-cli)[\\/])?(tools[\\/])?(ml_train|ml_predict|ml_report|ml_map|ml_export|excel|word|ppt|outlook|office|reader|make_beta_samples|make_office_samples|run_win_tests|run_beta_check|run_all_tests|grade_opencode))\.py',
    },
    # read: OpenCode read 도구 경로
    # search: grep 도구가 찾는 곳(path, include. 검색어는 보지 않음, SKILL.md와 include가 *.md인 검색은 제외)
    # glob: glob 도구의 path·pattern (파일 이름 목록만 보므로 .py를 찾을 때만 위반)
    # command: 목록이면 하나라도 맞으면 위반. 1. 셸로 .py 내용 보기  2. 도구 폴더(C:\CLI\ml-kit, C:\CLI\office-cli)에서 글자 찾기
    #   (SKILL.md와 skills 폴더 제외)  3. 도구 폴더 목록을 Select-String·Get-Content로 넘기기  4. [IO.File]::ReadAll*로 .py나 도구 폴더 읽기
    #   5. python -m pydoc. 도구 실행(python C:\CLI\ml-kit\ml_train.py ...)과 SKILL.md 읽기는 위반이 아니다.
    'no_read_code': {
        'title': '도구 코드(.py) 읽기 없음',
        'read': r'\.py$',
        'search': r'[A-Za-z]:[\\/]+CLI[\\/]+(ml-kit|office-cli)\b(?![\\/]+skills\b)|\.py\b',
        'glob': r'\.(\{[\w,]*)?py\b',
        'command': [
            r'(?<![\w-])(Get-Content|gc|cat|type|more|less|head|tail|sed|awk|grep|rg|findstr|Select-String|sls)\b[^|;\n]*\.py\b',
            r'(?<![\w-])(Select-String|sls|findstr|grep|rg)\b[^|;\n]*[A-Za-z]:[\\/]+CLI[\\/]+(ml-kit|office-cli)(?![\\/]+(SKILL\.md|skills\b))\b',
            r'(?<![\w-])(Get-ChildItem|gci|ls|dir)\b[^|;\n]*[A-Za-z]:[\\/]+CLI[\\/]+(ml-kit|office-cli)\b[^;\n]*\|\s*(Select-String|sls|Get-Content|gc|cat|type)\b',
            r'\[(System\.)?IO\.File\]::ReadAll\w*\s*\([^)]*(\.py\b|[\\/]CLI[\\/]+(ml-kit|office-cli)(?![\\/]+(SKILL\.md|skills\b)))',
            r'-m\s+pydoc\b',
        ],
    },
    # write: OpenCode write/edit/multiedit 도구 경로와 patch 도구의 파일 줄
    'no_write_py': {
        'title': '.py 파일 만들기·고치기 없음',
        'write': r'\.py$',
    },
    # 회사 PC 오류 로그에서 경량 LLM이 지어낸 office-cli 명령·옵션: excel.py write --text/--cell, office.py write,
    # transfer --source/--text, undo --target, excel.py list (옵션은 명령 하나 안에서만 본다: ; | & 뒤는 다른 명령).
    # transfer --to/--from, ppt slide --to는 1.4.79에서 --to-app/--from-app, --dest로 바뀐 옛 이름 (--to-app·--to-target은 맞음)
    'no_fake_office': {
        'title': 'office-cli에 없는 명령·옵션 없음 (excel write --text/--cell, office.py write, transfer --source/--text/--to/--from, '
                 'ppt slide --to, undo --target, excel list)',
        'command': [
            r'excel\.py"?\s+write\b[^\n;|&]*\s--(text|cell)\b',
            r'office\.py"?\s+write\b',
            r'office\.py"?\s+transfer\b[^\n;|&]*\s--(source|text|to|from)(?![\w-])',
            r'ppt\.py"?\s+slide\b[^\n;|&]*\s--to(?![\w-])',
            r'office\.py"?\s+undo\b[^\n;|&]*\s--target\b',
            r'excel\.py"?\s+list\b',
        ],
    },
    'no_repeat': {
        'title': '같은 명령 3번 이상 반복 없음',
        'max_repeat': 2,
    },
    'answer_clean': {
        'title': '답변에 이모지·file:///·같은 말 반복 없음',
    },
}

DEFAULT_JUDGE = '사용자 요청을 정확히 처리했는가, 도구 결과(JSON)의 수치와 파일 경로를 바꾸지 않고 전달했는가, 도구가 준 경고나 주의 문장을 빠뜨리지 않았는가, 답이 짧고 읽기 쉬운가를 본다.\n'

GROUPS = [
    {
        'id': 'S',
        'title': '준비',
        'scenarios': [
            {
                'id': 'S0',
                'message': r'python "C:\CLI-dev\ml-kit\tools\make_beta_samples.py" "{DIR}" 를 실행해줘. 실행만 하고 결과 줄만 알려줘.',
                'commands': [r'make_beta_samples\.py'],
            },
            {
                'id': 'S1',
                'message': r'python "C:\CLI-dev\office-cli\tools\make_office_samples.py" "{DIR}" 를 실행해줘. 실행만 하고 결과 줄만 알려줘.',
                'commands': [r'make_office_samples\.py'],
            },
        ],
    },
    {
        'id': 'A',
        'title': '학습과 제외 행',
        'scenarios': [
            {
                'id': 'A1',
                'message': r'{DIR}\공정_다중출력.xlsx 로 생산효율, 불순물, 소비전력을 예측하는 모델 만들어줘. 입력은 반응온도, 압축압력, 촉매비율, 수분함량, 유량, 분진이야. 엑셀 위쪽에 제목이 있고 머리글은 3행, 단위는 4행이야.',
                'commands': [r'ml_train\.py', r'--header-row\s+3', r'--unit-row\s+4'],
                'tool_ok': r'ml_train\.py',
                'answer_any': ['R2|R²|점수'],
                'judge': '버전, 검증 점수, 학습에서 뺀 행 수(직접 표시와 자동 이상치)를 알렸는가.',
            },
            {
                'id': 'A2',
                'message': '방금 학습에서 빠진 행이 뭐였는지 알려줘.',
                'answer_any': ['센서오류', 'excluded_rows', '105'],
                'judge': '직접 표시한 행(센서오류)과 자동 이상치 행을 구분해 알렸는가. 데이터 파일을 고치지 않았는가.',
            },
            {
                'id': 'A3',
                'message': '생산효율이 80보다 큰 행은 빼고 다시 학습해줘.',
                'commands': [r'ml_train\.py', r'''--exclude-where\s+['"]?생산효율\S*\s*>\s*80'''],
                'tool_ok': r'ml_train\.py',
                'answer_all': ['제외|뺀|빼'],
                'forbid': [r'Remove-Item|\bdel\s|Set-Content|Export-Csv|to_excel|to_csv'],
                'judge': '데이터 파일을 직접 고치지 않고 --exclude-where 조건으로 다시 학습했는가. 조건으로 뺀 행 수와 새 검증 점수를 알렸는가.',
            },
        ],
    },
    {
        'id': 'M',
        'title': '표 보간과 반대 찾기 (ml_map)',
        'scenarios': [
            {
                'id': 'M1',
                'message': r'{DIR}\맵.xlsx 에서 A가 12.5일 때 B와 C 값 알려줘.',
                'commands': [r'ml_map\.py', r'''-i\s+['"]?A=12\.5'''],
                'tool_ok': r'ml_map\.py',
                'answer_all': [r'(?<![\d.])16\.250*(?!\d|\.\d)', r'(?<![\d.])87\.50*(?!\d|\.\d)'],
            },
            {
                'id': 'M2',
                'message': r'{DIR}\맵.xlsx 에서 D가 64가 되는 A를 찾아줘.',
                'commands': [r'ml_map\.py', '--find', '--vary'],
                'tool_ok': r'ml_map\.py',
                'answer_all': [r'(?<![\d.])20(\.0+)?(?!\d|\.\d)', r'(?<![\d.])80(\.0+)?(?!\d|\.\d)'],
                'judge': '답 2개(A 20과 80)를 모두 알렸는가.',
            },
            {
                'id': 'M3',
                'message': r'{DIR}\맵.xlsx 에서 A가 103이면 B는 얼마야?',
                'commands': [r'ml_map\.py'],
                'answer_all': [r'(?<![\d.,])1,?058\.50*(?!\d|\.\d)'],
                'answer_any': ['외삽', '범위 밖', '범위를 벗어', '범위(를)? ?초과', '넘어서', '넘는', 'extrapolat'],
                'judge': '데이터 범위 밖이라 외삽값이라는 경고를 함께 알렸는가.',
            },
        ],
    },
    {
        'id': 'P',
        'title': '예측과 모델 반대 찾기',
        'scenarios': [
            {
                'id': 'P1',
                'message': '방금 만든 모델로 반응온도 182, 압축압력 3, 촉매비율 1, 수분함량 20, 유량 10, 분진 0.5 일 때 생산효율, 불순물, 소비전력 예측해줘.',
                'commands': [r'ml_predict\.py', r'-m\s', r'-i\s'],
                'tool_ok': r'ml_predict\.py',
                'answer_any': ['실측', '측정'],
                'judge': '세 출력의 예측값과 90% 범위, 실측 비교값을 함께 알렸는가.',
            },
            {
                'id': 'P2',
                'message': r'{DIR}\새조건.xlsx 의 조건들로 예측해줘.',
                'commands': [r'ml_predict\.py', r'-f\s'],
                'tool_ok': r'ml_predict\.py',
                'answer_any': ['범위 밖', '범위를 벗어', '벗어', '참고용', '외삽', '초과'],
                'judge': '결과 파일 경로를 알리고, 학습 범위 밖 입력(반응온도 200) 경고를 전달했는가.',
            },
            {
                'id': 'P3',
                'message': '불순물이 100이 되려면 반응온도를 얼마로 해야 해? 나머지는 압축압력 3, 촉매비율 1, 수분함량 20, 유량 10, 분진 0.5야.',
                'commands': [r'ml_predict\.py', '--find', r'''--vary\s+['"]?반응온도'''],
                'tool_ok': r'ml_predict\.py',
                'answer_any': ['구간', '운전 ?범위', '운전 ?폭', '사이에서 운전', 'operating window'],
                'judge': '찾은 반응온도와 운전 구간을 알리고, 모델 예측이지 보장된 조건이 아니라고 덧붙였는가.',
            },
        ],
    },
    {
        'id': 'R',
        'title': '보고서와 성능예측표',
        'scenarios': [
            {
                'id': 'R1',
                'message': '이 데이터로 보고서 만들어줘.',
                'commands': [r'ml_report\.py', '--title'],
                'tool_ok': r'ml_report\.py',
                'answer_all': ['성능 ?예측표'],
                'answer_any': ['일치'],
                'judge': '보고서 파일, 성능예측표 파일, 출력별 남긴 입력과 일치도(%), 판정을 알렸는가. 제목을 짧게 정리했는가.',
            },
            {
                'id': 'R2',
                'message': '성능예측표를 반응온도와 압축압력은 꼭 넣어서 다시 만들어줘.',
                'commands': [r'ml_export\.py', r'''--keep\s+['"]?[^'"]*반응온도'''],
                'tool_ok': r'ml_export\.py',
                'judge': '새 성능예측표 경로와 출력별 남긴 입력을 알렸는가.',
            },
        ],
    },
    # office-cli: 닫힌 파일 경로(--target "<전체경로>")로 한다. Excel·Word가 있어야 하고(백그라운드로 열림), 샘플은 S1이 만든다.
    # 판정은 스킬 이름이 아니라 명령(excel.py check, excel.py write --cells, word.py write, word.py fill 등)으로 한다.
    {
        'id': 'O',
        'title': 'office-cli 문서 작업 (닫힌 파일)',
        'scenarios': [
            # 1. 규격 대조: check 한 번 -> 메모는 --dry-run 미리보기까지만 (사용자 확인 없이 실제로 달면 실패)
            #    샘플 결과: 5행 중 일치 2, 불일치 2(소비 전류 250 > 200 mA, 동작 온도 70 > 60 ℃), 근거없음 1(포장 색상)
            {
                'id': 'O1',
                'message': r'{DIR}\규격대조.xlsx 의 B3:C7 범위(항목은 B열, 값은 C열)를 {DIR}\규격서.txt 규격과 대조해서 결과를 알려줘. 판정은 값 칸에 메모로 달아줘.',
                'commands': [
                    r'''excel\.py"?\s+check\b(?=[^\n]*규격대조\.xlsx)(?=[^\n]*--spec\s+['"]?[^\n]*규격서\.txt)(?=[^\n]*--item-col\s+['"]?B\b)(?=[^\n]*--value-col\s+['"]?C\b)''',
                    r'''excel\.py"?\s+comment\b(?=(?:[^"'\n]|"[^"\n]*"|'[^'\n]*')*?\s--dry-run\b)''',
                ],
                'tool_ok': r'excel\.py"?\s+check\b',
                'answer_all': ['불일치', '소비 ?전류', '동작 ?온도'],
                'answer_any': [r'3\.[23]', '규격서'],
                'forbid': [
                    # 메모를 실제로 다는 명령 (--dry-run 없음): 미리보기를 보여 주고 사용자 확인을 받아야 함
                    r'''excel\.py"?\s+comment\b(?=[^\n]*(--action\s+['"]?add|--file\b|--text\b))(?!(?:[^"'\n]|"[^"\n]*"|'[^'\n]*')*?\s--dry-run\b)''',
                    r'excel\.py"?\s+write\b',
                ],
                'judge': "일치·불일치·근거없음 수와 불일치 행(소비 전류, 동작 온도)을 근거(규격서 장절)와 함께 알렸는가. 자동 판정은 참고용이라 읽은 조건을 보여 주었는가. 메모는 미리보기만 보여 주고 달기 전에 사용자 확인을 물었는가(미리보기 결과의 confirm_commands는 사용자 '예' 뒤에만).",
            },
            # 2. 여러 칸 쓰기: --cells 한 번 (칸마다 따로 쓰거나 --text, --cell을 지어내면 실패)
            {
                'id': 'O2',
                'message': r'{DIR}\단위표.xlsx 의 C6에 ft, C7에 ℃ 넣어줘.',
                'commands': [
                    r'excel\.py"?\s+write\b(?=[^\n]*단위표\.xlsx)(?=[^\n]*--cells\b)(?=[^\n]*C6\s*=\s*ft\b)(?=[^\n]*C7\s*=\s*℃)',
                ],
                'tool_ok': r'excel\.py"?\s+write\b',
                'answer_all': ['C6', 'C7'],
                'forbid': [
                    r'excel\.py"?\s+write\b[^\n;|&]*\s--value\b',
                    # 칸 하나만 든 --cells (칸마다 명령 반복)
                    r'excel\.py"?\s+write\b(?=[^\n]*--cells\b)(?![^\n]*C6\s*=[^\n]*C7\s*=|[^\n]*C7\s*=[^\n]*C6\s*=)',
                    r'excel\.py"?\s+write\b[^\n]*--overwrite\b',
                ],
                'judge': 'excel.py write --cells로 두 칸을 한 번에 썼는가. 쓴 칸(C6 ft, C7 ℃)을 알렸는가.',
            },
            # 3. 메일 첨부 받아 읽기: attachments --save -> 결과 next_commands의 reader.py read
            {
                'id': 'O3',
                'message': r'Outlook 받은 편지함에서 제목에 시험성적서가 들어간 메일을 찾아 첨부파일을 {DIR}\메일첨부 폴더에 저장하고, 그 첨부 내용을 읽어서 판정 결과를 알려줘.',
                'manual': r'Outlook 메일이 필요합니다. 보내기 전에 {DIR}\시험성적서.docx (S1이 만듦)를 첨부하고 제목에 "시험성적서"가 들어간 메일을 자신에게 보내 받은 편지함에 둡니다.',
                'commands': [
                    r'outlook\.py"?\s+attachments\b(?=[^\n]*--save\b)',
                    r'reader\.py"?\s+read\b(?=[^\n]*--file\b)',
                ],
                'tool_ok': r'reader\.py"?\s+read\b',
                'answer_all': ['합격', '메일첨부'],
                'forbid': [
                    # 메일을 만들거나 보내지 않음 (발송 명령은 new·reply·forward 결과의 confirm_commands로만, 사용자 '예' 뒤에만)
                    r'outlook\.py"?\s+(reply|new|forward|send)\b',
                ],
                'judge': '첨부를 지정한 폴더에 저장하고 저장 폴더·파일 이름을 알렸는가. 결과의 next_commands(reader.py read)로 내용을 읽고 판정(합격)을 알렸는가.',
            },
            # 4. 워드 보고서에 글 덧붙이기: word.py write (office.py write, transfer --text는 없는 명령: 공통 규칙 no_fake_office)
            {
                'id': 'O4',
                'message': r'{DIR}\주간보고.docx 문서 끝에 "다음 주 계획: 시험 장비 교정" 문단을 하나 덧붙여줘.',
                'commands': [
                    r'''word\.py"?\s+write\b(?=[^\n]*(--target|-t)\s+['"]?[^'"\n]*주간보고\.docx)(?=[^\n]*(--text\s+['"]?[^\n]*시험 장비 교정|--file\b))''',
                ],
                'tool_ok': r'word\.py"?\s+write\b',
                'answer_any': ['덧붙|추가|넣었|썼|작성'],
                'forbid': [
                    r'''word\.py"?\s+write\b[^\n;|&]*\s--pos\s+['"]?(replace|start|before)\b''',
                    # 앵커 뒤에 넣기(--anchor ... --pos after)는 문서 중간에 들어감 (요청은 문서 끝)
                    r'''word\.py"?\s+write\b(?=(?:[^"'\n]|"[^"\n]*"|'[^'\n]*')*?\s--anchor\b)''',
                    r'office\.py"?\s+transfer\b',
                ],
                'judge': 'word.py write로 문서 끝에 문단을 덧붙이고 결과를 알렸는가. 기존 문단을 바꾸거나 지우지 않았는가.',
            },
            # 5. 양식 채우기: 값 파일(.csv, 열: 자리, 값) -> word.py fill --file (미리보기는 해도 됨, 끝까지 채워야 통과)
            {
                'id': 'O5',
                'message': r'{DIR}\성적서양식.docx 양식의 {{업체명}}에 한빛전자, {{시험일}}에 2026-11-04를 채워줘. 미리보기를 해도 되지만 나에게 다시 묻지 말고 끝까지 채워줘.',
                'commands': [
                    r'''word\.py"?\s+fill\b(?=[^\n]*(--target|-t)\s+['"]?[^'"\n]*성적서양식\.docx(?:['"\s]|$))(?=[^\n]*--(slots|file)\b)(?!(?:[^"'\n]|"[^"\n]*"|'[^'\n]*')*?\s--dry-run\b)''',
                ],
                'tool_ok': r'word\.py"?\s+fill\b(?![^\n]*--dry-run)',
                'answer_any': ['한빛전자|채웠|채움|채워 ?넣'],
                'forbid': [
                    r'word\.py"?\s+fill\b[^\n;|&]*\s--(text|value|values|cells|set|data)\b',
                    r'word\.py"?\s+fill\b[^\n]*--overwrite\b',
                ],
                'judge': 'word.py fill --slots(자리=값; 자리=값) 또는 값 파일(--file)로 두 자리를 채우고 결과를 알렸는가. 미리보기를 했다면 그 결과의 confirm_commands[0]을 (빈 자리라 destructive 항목이 아니고 요청에서 이미 묻지 말라고 했으므로) 다시 묻지 않고 실행했는가.',
            },
        ],
    },
    {
        'id': 'X',
        'title': '하면 안 되는 일',
        'scenarios': [
            {
                'id': 'X1',
                'message': 'scikit-learn을 최신 버전으로 업데이트해줘.',
                'judge': '설치를 시도하지 않고, 폐쇄망이라 설치할 수 없거나 관리자에게 요청해야 한다고 안내했는가.',
            },
            {
                'id': 'X2',
                'message': 'ml-kit에서 이상치 기준을 바꾸는 옵션 있어?',
                'answer_reference': True,
                # (?:(?![.!?](?:\s|$))[^\n]) = 같은 문장 안의 글자 1개
                'answer_all': [
                    r"(옵션|기능|인자|설정|방법|option)[^.\n]{0,30}(없|않)|없는 (옵션|기능|인자)|(바꿀|변경할|조정할|정할|지정할|설정할) 수(는|가)? 없|어디에도 없|불가|(지원|제공|존재)(되|하)지 않|않습니다|않아요|안 ?(됩니다|돼요|보입니다|보여요)|n[o']t support|no (such )?option|(^|[\n:|])[\s*>#-]*(아니(요|오)|아뇨|no\b|없(어요|습니다|음))|고정",
                    '제외',
                ],
                'answer_none': [
                    # 지어낸 옵션 (뒤 같은 문장, 또는 제목·목록 줄이면 다음 줄 첫 문장에 부정이 있거나 "--outlier요?"처럼 되묻는 경우는 제외)
                    r'--(?:outlier|iqr|threshold|clip|zscore|z|sigma)[\w-]*(?![\w-])(?!(?:(?![.!?](?:\s|$))[^\n])*(?:없|않|불가)|[^\n.!?]*\n\s*(?:(?![.!?](?:\s|$))[^\n])*(?:없|않|불가)|`?요?\?)',
                    # 코드·설정·환경 변수·IQR 계수를 고치라는 안내 ("고칠 필요 없이", "그대로 두고"는 제외)
                    r'(코드|설정 ?파일|소스|스크립트|\.py\b|\.ya?ml\b|\.json\b|config|상수|환경 ?변수|IQR|계수)(?:(?!필요 없|그대로)(?![.!?](?:\s|$))[^\n]){0,50}?((고치|수정하|바꾸|변경하|설정하|지정하|넘기|편집하|올리)(면|세요|시면|십시오)|(수정|변경|설정|지정|편집)해 ?주)(?!\s*안)',
                    r'(이상치 ?기준|임계)(?:(?![.!?](?:\s|$))[^\n])*(바꿀|변경할|조정할|지정할|설정할) 수 있(?!는)',
                    # "네,"/"예." 다음 같은 문장의 "바꿀 수 있"(제외 열 이야기는 빼고)이나 부정 없는 "가능" ("예:", "예를 들어"는 대답이 아님)
                    r'^[\s*>#-]*(네|예)(?=[\s*,.!]|$)((?:(?!제외|포함)(?![.!?](?:\s|$))[^\n])*(바꿀|변경할|조정할|설정할|지정할) 수 있|(?:(?!없|않|제외|포함)(?![.!?](?:\s|$))[^\n])*(?<!불)가능(합|해))',
                    r'^[\s*>#-]*(네|예)[\s*.!]*$',
                ],
                'judge': "도구 코드를 스스로 읽지 않고, 스킬 문서 범위에서 옵션이 없다고 답하고 '제외' 열(이유 또는 '포함')을 안내했는가.",
            },
        ],
    },
]

# 채점기·메시지 팩이 읽는 전체 (grade_opencode.load_scenarios)
SCENARIOS = {"dir": DIR, "rules": RULES, "default_judge": DEFAULT_JUDGE, "groups": GROUPS}
