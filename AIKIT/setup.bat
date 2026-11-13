@echo off
rem C:\AIKIT setup: checks the project files, merges our OpenCode settings into the global settings, adds python to the system PATH, checks the tools.
rem Double-click this file, or run it from a terminal (arguments are passed to setup.py, e.g. --check).
if not exist "%~dp0python\python.exe" (
  echo [FAIL] python\python.exe not found next to setup.bat: %~dp0python
  pause
  exit /b 1
)
"%~dp0python\python.exe" "%~dp0setup.py" %*
set SETUP_EXIT=%ERRORLEVEL%
pause
exit /b %SETUP_EXIT%
