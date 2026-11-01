import subprocess, sys, os, json
PY = r"C:\WPy64-313130\python\python.exe"


def emit(line, code):
    # UTF-8 bytes: print() would fail on a cp949 pipe for characters such as U+2022, dashes or NBSP
    sys.stdout.buffer.write((line + "\n").encode("utf-8", "replace"))
    sys.stdout.buffer.flush()
    sys.exit(code)


def fail(message, stderr="", code=1):
    emit(json.dumps({"status": "error", "error_type": "runner", "message": message, "stderr": stderr,
                     "agent_directive": "ERROR: 도구를 실행하지 못했거나 결과 JSON이 없습니다. 재시도하지 말고 "
                                        "message와 stderr를 사용자에게 그대로 알리세요."}, ensure_ascii=False), code)


env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
try:
    r = subprocess.run([PY] + sys.argv[1:], capture_output=True, env=env)
except KeyboardInterrupt:
    fail("사용자가 중단함", code=130)
except OSError as e:
    fail(f"python 실행 실패 ({PY}): {e}. run.py 2행의 WinPython 경로를 확인하세요.")
out = [x for x in r.stdout.decode("utf-8", "replace").strip().splitlines() if x.strip()]
err = r.stderr.decode("utf-8", "replace")[-800:]
if not out:
    fail(f"출력 없음 (exit {r.returncode})", err, r.returncode or 1)
try:
    is_json = isinstance(json.loads(out[-1]), dict)
except ValueError:
    is_json = False
if not is_json:
    fail(f"마지막 출력 줄이 JSON이 아님 (exit {r.returncode}): {out[-1][:300]}", err, r.returncode or 1)
emit(out[-1], r.returncode)
