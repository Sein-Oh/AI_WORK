@echo off
rem OpenCode Web UI launcher (double-click), for opencode 2.x.
rem 'opencode pair' starts the OpenCode background service if needed and prints one-time login links.
rem This opens the local (127.0.0.1/localhost) link in the default browser.
rem The web UI is embedded in opencode.exe, so it works without internet.
rem Usage: web.bat         open the Web UI
rem        web.bat stop    stop the background service
rem        web.bat status  show the background service status
setlocal
set "OC=%~dp0opencode.exe"
if not exist "%OC%" (
  echo [FAIL] opencode.exe not found: %OC%
  pause
  exit /b 1
)
if /i "%~1"=="stop" goto service
if /i "%~1"=="status" goto service

rem Run 'pair' once and keep its whole output in a file (no pipe inside for /f).
if not exist "%~dp0data" mkdir "%~dp0data"
set "OUT=%~dp0data\web_pair.txt"
"%OC%" pair > "%OUT%" 2>&1
set "RC=%ERRORLEVEL%"
set "LINK="
for /f "usebackq tokens=*" %%L in (`findstr /R /C:"http://127\.0\.0\.1[:/]" /C:"http://localhost[:/]" "%OUT%"`) do (
  if not defined LINK set "LINK=%%L"
)
if not defined LINK (
  echo [FAIL] no local login link from 'opencode pair' ^(exit code %RC%^). Output:
  type "%OUT%"
  pause
  exit /b 1
)
del "%OUT%" >nul 2>&1
echo [INFO] Opening OpenCode Web UI: %LINK%
echo [INFO] The link works once and expires in a few minutes. Run web.bat again for a new one.
echo [INFO] The background service keeps running after this window closes. Stop it with: web.bat stop
start "" "%LINK%"
endlocal
exit /b 0

:service
"%OC%" service %~1
set "RC=%ERRORLEVEL%"
pause
exit /b %RC%
