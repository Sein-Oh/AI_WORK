# -*- coding: utf-8 -*-
"""셸 시뮬레이터: 도구가 낸 명령 문자열(next_commands 등)을 실제 셸이 실행할 때 파이썬이 받는 인자 목록을 만든다.

OpenCode 셸 도구는 Windows에서 pwsh(PowerShell 7) -> powershell(Windows PowerShell 5.1) -> Git Bash -> cmd 순서로
처음 찾은 셸을 쓴다. 단계는 둘이다.
1. 셸 문법: PowerShell 토크나이저(작은·큰따옴표, 백틱, $ 확장 감지) / cmd(%NAME% 확장, ^, &|<>) / bash.
2. 네이티브 전달: PowerShell은 인자 값들로 명령줄 문자열을 다시 만든다.
   - Windows PowerShell 5.1 (레거시): 값을 그대로 이어 붙이고, 공백이 있으면 바깥만 "..."로 감싼다(안의 "는 그대로),
     빈 값은 버린다.
   - PowerShell 7.3+ (Standard): .NET ArgumentList 규칙으로 올바르게 이스케이프한다.
   그 문자열을 파이썬(UCRT, CommandLineToArgvW와 같은 계열)이 argv로 나눈다.

도구의 quote_arg가 셸 단계에서 위험한 모양을 내면(따옴표 밖의 $, ;, @ 등) ShellError로 알린다: 실제 셸이라면
값이 바뀌거나 다른 명령이 실행되는 경우다. 테스트 도우미 split_command(office_helpers, mlkit_helpers)가 이 모듈을 쓴다.
"""
import os
import re
import shutil

SHELLS = ("pwsh", "powershell", "cmd", "bash")
PS_SINGLE = "'\u2018\u2019\u201a\u201b"
PS_DOUBLE = '"\u201c\u201d\u201e'
PS_DASHES = "-\u2013\u2014\u2015"
_PLAIN_INT = re.compile(r"-?(?:0|[1-9]\d{0,8})")
_PS_OPTION = re.compile(r"--?[A-Za-z][A-Za-z0-9-]*")
# 따옴표 밖에서 PowerShell 문법이 되는 글자 (값이 바뀌거나 명령이 나뉨). 쉼표는 네이티브 명령에는 글자지만
# 맨 앞·맨 뒤·연속 쉼표는 지원되지 않음 (about_Parsing)
_PS_BARE_SPECIAL = set(";|&(){}<>$`[]\"'") | set(PS_SINGLE) | set(PS_DOUBLE)


class ShellError(ValueError):
    """셸이 값을 글자 그대로 넘기지 않는 모양 (확장, 명령 구분, 주석, 숫자 변환 등)."""


def current_shell(env_var):
    """도구의 command_shell()과 같은 규칙: 환경변수 > pwsh가 PATH에 있으면 pwsh > powershell."""
    forced = os.environ.get(env_var, "").strip().lower()
    if forced in SHELLS:
        return forced
    return "pwsh" if shutil.which("pwsh") else "powershell"


# ---------------------------------------------------------------------------
# 1. PowerShell 토크나이저 (네이티브 명령 호출 한 줄, 인자 모드)
# ---------------------------------------------------------------------------

def _net_space(ch):
    import unicodedata
    return ch in "\t\n\v\f\r\x85\xa0" or unicodedata.category(ch) in ("Zs", "Zl", "Zp")


def _check_bare(word):
    """따옴표 없는 낱말이 PowerShell에서 글자 그대로인지 (숫자 변환·매개변수 쪼개기·배열·주석이 아닌지)."""
    if word.startswith("#"):
        raise ShellError(f"맨 앞 #는 주석: {word!r}")
    if word.startswith("@"):
        raise ShellError(f"맨 앞 @는 배열·해시·스플래팅: {word!r}")
    bad = [ch for ch in word if ch in _PS_BARE_SPECIAL]
    if bad:
        raise ShellError(f"따옴표 밖의 PowerShell 특수문자 {bad[0]!r}: {word!r}")
    if word[0] == "," or word[-1] == "," or ",," in word:
        raise ShellError(f"맨 앞·맨 뒤·연속 쉼표는 배열: {word!r}")
    if word[0] in PS_DASHES and not (_PS_OPTION.fullmatch(word) or _PLAIN_INT.fullmatch(word)):
        raise ShellError(f"-로 시작하는 낱말은 매개변수로 읽혀 나뉠 수 있음: {word!r}")
    if word == "--" or word.startswith("--%"):
        raise ShellError(f"PowerShell이 따로 다루는 낱말: {word!r}")
    if (word[0].isdigit() or word[0] in "+." and word[1:2].isdigit()) and not _PLAIN_INT.fullmatch(word):
        raise ShellError(f"숫자로 읽혀 바뀔 수 있음: {word!r}")


def _ps_double(line, i):
    """큰따옴표 문자열 (i는 여는 따옴표 다음). (값, 다음 위치). $ 확장은 ShellError."""
    out = []
    n = len(line)
    while i < n:
        ch = line[i]
        if ch == "`":
            if i + 1 >= n:
                raise ShellError("끝나지 않은 백틱")
            nxt = line[i + 1]
            out.append({"0": "\0", "a": "\a", "b": "\b", "f": "\f", "n": "\n", "r": "\r", "t": "\t",
                        "v": "\v"}.get(nxt, nxt))
            i += 2
            continue
        if ch == "$" and i + 1 < n and (line[i + 1].isalnum() or line[i + 1] in "_{(?^$:"):
            raise ShellError(f"큰따옴표 안의 $ 확장: {line[i:i + 10]!r}")
        if ch in PS_DOUBLE:
            if i + 1 < n and line[i + 1] in PS_DOUBLE:
                out.append(line[i + 1])
                i += 2
                continue
            return "".join(out), i + 1
        out.append(ch)
        i += 1
    raise ShellError("닫히지 않은 큰따옴표")


def _ps_single(line, i):
    """작은따옴표 문자열 (i는 여는 따옴표 다음). 안의 작은따옴표 두 개 -> 뒤의 것 하나 (PowerShell 토크나이저)."""
    out = []
    n = len(line)
    while i < n:
        ch = line[i]
        if ch in PS_SINGLE:
            if i + 1 < n and line[i + 1] in PS_SINGLE:
                out.append(line[i + 1])
                i += 2
                continue
            return "".join(out), i + 1
        out.append(ch)
        i += 1
    raise ShellError("닫히지 않은 작은따옴표")


def ps_tokenize(line):
    """PowerShell 한 줄(네이티브 명령 호출) -> [명령, 인자 값...]. 인자 하나는 따옴표 문자열 하나이거나 따옴표 없는
    낱말 하나여야 한다(붙여 쓰기는 PowerShell 버전마다 달라 ShellError)."""
    args = []
    i, n = 0, len(line)
    while i < n:
        ch = line[i]
        if ch in " \t":
            i += 1
            continue
        if ch in "\r\n":
            raise ShellError("줄바꿈은 명령을 나눔")
        if ch in PS_SINGLE or ch in PS_DOUBLE:
            value, i = (_ps_single if ch in PS_SINGLE else _ps_double)(line, i + 1)
            if i < n and line[i] not in " \t":
                raise ShellError(f"따옴표 뒤에 붙은 글자: {line[i:i + 10]!r}")
            args.append(value)
            continue
        j = i
        while j < n and line[j] not in " \t\r\n":
            j += 1
        word = line[i:j]
        _check_bare(word)
        args.append(word)
        i = j
    return args


# ---------------------------------------------------------------------------
# 2. 네이티브 명령줄 조립 (PowerShell -> 파이썬)
# ---------------------------------------------------------------------------

def _need_quotes(arg):
    """Windows PowerShell 5.1 NativeCommandParameterBinder.NeedQuotes: 짝이 맞지 않은 " 밖에 공백이 있으면 감쌈.
    \\ 바로 뒤의 "는 세지 않음."""
    quotes, after_slash = 0, False
    for ch in arg:
        if ch == '"' and not after_slash:
            quotes += 1
        elif _net_space(ch) and quotes % 2 == 0:
            return True
        after_slash = ch == "\\"
    return False


def legacy_command_line(args):
    """Windows PowerShell 5.1 (레거시): 빈 값은 버리고, 공백이 있으면 바깥만 "..."로 감싸 그대로 이어 붙임."""
    parts = [args[0]]
    for arg in args[1:]:
        if arg == "":
            continue
        parts.append('"' + arg + '"' if _need_quotes(arg) else arg)
    return " ".join(parts)


def _paste_argument(arg):
    """.NET PasteArguments.AppendArgument (PowerShell 7.3+ Standard)."""
    if arg and not any(ch.isspace() or ch == '"' for ch in arg):
        return arg
    out, slashes = ['"'], 0
    for ch in arg:
        if ch == "\\":
            slashes += 1
            continue
        if ch == '"':
            out.append("\\" * (2 * slashes + 1) + '"')
        else:
            out.append("\\" * slashes + ch)
        slashes = 0
    out.append("\\" * (2 * slashes) + '"')
    return "".join(out)


def standard_command_line(args):
    """PowerShell 7.3+ ($PSNativeCommandArgumentPassing = Windows/Standard, python.exe는 Standard)."""
    return " ".join([args[0]] + [_paste_argument(a) for a in args[1:]])


# ---------------------------------------------------------------------------
# 3. Windows 명령줄 -> argv (UCRT parse_command_line, 파이썬 sys.argv; CommandLineToArgvW와 같은 계열)
# ---------------------------------------------------------------------------

def windows_argv(line):
    """백슬래시는 " 앞에서만 특별 (2n개 + " -> n개 + 따옴표 열고 닫기, 2n+1개 + " -> n개 + 글자 "),
    따옴표 안의 ""는 글자 " 하나 (따옴표 안에 머묾). shlex(posix)는 C:\\a 의 백슬래시를 지우므로 쓰지 않는다."""
    args, cur, have, quoted, i, n = [], [], False, False, 0, len(line)
    while i < n:
        ch = line[i]
        if ch in " \t" and not quoted:
            if have:
                args.append("".join(cur))
                cur, have = [], False
            i += 1
            continue
        if ch == "\\":
            j = i
            while j < n and line[j] == "\\":
                j += 1
            count = j - i
            if j < n and line[j] == '"':
                cur.append("\\" * (count // 2))
                have = True
                if count % 2:
                    cur.append('"')
                    i = j + 1
                    continue
                i = j                     # 짝수: 뒤의 "는 아래 따옴표 규칙으로
                continue
            cur.append("\\" * count)
            i, have = j, True
            continue
        if ch == '"':
            have = True
            if quoted and i + 1 < n and line[i + 1] == '"':
                cur.append('"')
                i += 2
                continue
            quoted = not quoted
            i += 1
            continue
        cur.append(ch)
        have = True
        i += 1
    if have:
        args.append("".join(cur))
    return args


# ---------------------------------------------------------------------------
# 4. cmd (명령줄 모드: %NAME% 확장 -> ^ 처리와 &|<> 검사 -> 파이썬 argv), bash
# ---------------------------------------------------------------------------

CMD_ENV = {"PATH": r"C:\Windows\system32", "USERNAME": "user", "X": "확장됨", "TEMP": r"C:\Temp"}


def cmd_line(line, env=None):
    """cmd.exe가 python에 넘기는 명령줄. 정의된 %NAME%는 바뀌고(바뀌면 ShellError 대신 값이 달라짐),
    따옴표 밖의 ^는 다음 글자를 글자로, 따옴표 밖의 & | < >는 ShellError."""
    env = {k.upper(): v for k, v in (CMD_ENV if env is None else env).items()}
    out, i, n = [], 0, len(line)
    while i < n:
        if line[i] == "%":
            j = line.find("%", i + 1)
            if j > i + 1 and line[i + 1:j].upper() in env:
                out.append(env[line[i + 1:j].upper()])
                i = j + 1
                continue
        out.append(line[i])
        i += 1
    expanded = "".join(out)
    out, quoted, i, n = [], False, 0, len(expanded)
    while i < n:
        ch = expanded[i]
        if ch == '"':
            quoted = not quoted
        elif not quoted and ch == "^":
            if i + 1 < n:
                out.append(expanded[i + 1])
            i += 2
            continue
        elif not quoted and ch in "&|<>":
            raise ShellError(f"따옴표 밖의 cmd 특수문자 {ch!r}")
        elif ch in "\r\n":
            raise ShellError("줄바꿈은 명령을 나눔")
        out.append(ch)
        i += 1
    return "".join(out)


def bash_argv(line):
    """bash 한 줄 -> argv. 작은따옴표는 글자 그대로, 큰따옴표 안의 $·백틱과 따옴표 밖의 특수문자는 ShellError."""
    args, cur, have, i, n = [], [], False, 0, len(line)
    while i < n:
        ch = line[i]
        if ch in " \t":
            if have:
                args.append("".join(cur))
                cur, have = [], False
            i += 1
            continue
        if ch == "'":
            j = line.find("'", i + 1)
            if j < 0:
                raise ShellError("닫히지 않은 작은따옴표")
            cur.append(line[i + 1:j])
            i, have = j + 1, True
            continue
        if ch == '"':
            j = i + 1
            while j < n and line[j] != '"':
                if line[j] in "$`":
                    raise ShellError("큰따옴표 안의 확장")
                if line[j] == "\\" and j + 1 < n and line[j + 1] in '"\\$`':
                    cur.append(line[j + 1])
                    j += 2
                    continue
                cur.append(line[j])
                j += 1
            if j >= n:
                raise ShellError("닫히지 않은 큰따옴표")
            i, have = j + 1, True
            continue
        if ch == "\\":
            if i + 1 < n:
                cur.append(line[i + 1])
            i, have = i + 2, True
            continue
        if ch in ";&|<>()$`*?[]{}~#\n":          # ! 는 대화형 bash에서만 (셸 도구는 bash -c)
            raise ShellError(f"따옴표 밖의 bash 특수문자 {ch!r}")
        cur.append(ch)
        i, have = i + 1, True
    if have:
        args.append("".join(cur))
    return args


# ---------------------------------------------------------------------------
# 셸 하나로 끝까지
# ---------------------------------------------------------------------------

def run_shell(cmd, shell):
    """명령 문자열을 그 셸로 실행할 때 파이썬이 받는 argv ([python, 스크립트, 인자...])."""
    if shell == "powershell":
        return windows_argv(legacy_command_line(ps_tokenize(cmd)))
    if shell == "pwsh":
        return windows_argv(standard_command_line(ps_tokenize(cmd)))
    if shell == "cmd":
        return windows_argv(cmd_line(cmd))
    if shell == "bash":
        return bash_argv(cmd)
    raise ValueError(shell)


def split_command(cmd, env_var):
    """도구가 낸 명령 문자열을 도구와 같은 셸 규칙(env_var 강제 > pwsh 있음 > powershell)으로 인자 목록으로 나눈다."""
    return run_shell(cmd, current_shell(env_var))
