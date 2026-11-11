"""run_all_tests.py pytest 단계용 pytest 플러그인 (`-p pytest_live_report`, 사람이 쓰는 점검 도구).

pytest -q는 실패 이름(-rfE 요약)을 맨 끝에만 출력하므로, 시간 상한에 걸려 중간에 끊기면 무엇이 실패했는지 남지 않는다
(회사 PC 2026-11-05: ml-kit 3600초 초과, F 11개 이름 없음). 이 플러그인은 실패·오류가 생기는 즉시 한 줄을 출력한다:
    [LIVE] FAILED <테스트 id> - <짧은 원인>
    [LIVE] ERROR <테스트 id> (setup) - <짧은 원인>
run_all_tests.parse_live()가 이 줄을 모은다.
"""
LIVE_TAG = "[LIVE]"
_reporter = []


def pytest_configure(config):
    tr = config.pluginmanager.get_plugin("terminalreporter")
    if tr is not None:
        _reporter.append(tr)


def short_reason(report, limit=200):
    """실패 원인 한 줄: 예외 메시지의 첫 줄 (없으면 마지막 줄)."""
    crash = getattr(report.longrepr, "reprcrash", None)
    text = getattr(crash, "message", None) or str(report.longrepr or "")
    lines = [x.strip() for x in str(text).splitlines() if x.strip()]
    return (lines[0] if lines else "")[:limit]


def pytest_runtest_logreport(report):
    if not report.failed:
        return
    kind = "FAILED" if report.when == "call" else "ERROR"
    where = "" if report.when == "call" else f" ({report.when})"
    line = f"{LIVE_TAG} {kind} {report.nodeid}{where} - {short_reason(report)}"
    if _reporter:                       # 바로 내보냄: 시간 초과로 프로세스가 끝나도 이 줄은 남음
        _reporter[0].write_line(line)
        try:
            _reporter[0]._tw.flush()
        except Exception:
            pass
    else:
        print(line, flush=True)
