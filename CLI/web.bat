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
if /i "%~1"=="stop" (
  "%OC%" service stop
  pause
  exit /b %ERRORLEVEL%
)
if /i "%~1"=="status" (
  "%OC%" service status
  pause
  exit /b %ERRORLEVEL%
)
set "LINK="
for /f "tokens=*" %%L in ('""%OC%" pair 2^>nul ^| findstr /R /C:"http://127\.0\.0\.1[:/]" /C:"http://localhost[:/]""') do (
  if not defined LINK set "LINK=%%L"
)
if not defined LINK (
  echo [FAIL] no local login link from 'opencode pair'. Output:
  "%OC%" pair
  pause
  exit /b 1
)
echo [INFO] Opening OpenCode Web UI: %LINK%
echo [INFO] The link works once and expires in a few minutes. Run web.bat again for a new one.
echo [INFO] The background service keeps running after this window closes. Stop it with: web.bat stop
start "" "%LINK%"
endlocal
exit /b 0
