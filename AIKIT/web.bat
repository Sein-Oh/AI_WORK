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
rem Our tools are used only when this folder is the OpenCode project (AGENTS.md, .opencode\skills\, opencode.json),
rem user decision 2026-11-13. The internal LLM settings (opencode.jsonc) are in the user's global OpenCode settings.
set "AIKIT=%~dp0"
set "AIKIT=%AIKIT:~0,-1%"
if /i "%~1"=="stop" goto service
if /i "%~1"=="status" goto service

rem Project folder = this folder (C:\AIKIT, user decision 2026-11-12). The Web UI makes every session inside a project folder,
rem and a new PC has none, so "new session" did nothing (team PC 2026-11-11). The service starts here; the user opens it once
rem in the Web UI (Open project).
set "WORK=%AIKIT%"
cd /d "%WORK%"

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
echo [INFO] First time on this PC: in the Web UI choose Open project and pick %WORK% , then start a new session there.

rem Open in Edge (or Chrome) with its own profile folder under data\web_profile, as an app window.
rem The normal browser profile can fail with "Internal error opening backing store for indexedDB.open"
rem (broken or cleaned site data), while a fresh profile works. Falls back to the default browser.
set "PROFILE=%~dp0data\web_profile"
set "BROWSER="
set "PF86=%ProgramFiles(x86)%"
if exist "%PF86%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%PF86%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "BROWSER=%LocalAppData%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER goto default_browser
if not exist "%PROFILE%" mkdir "%PROFILE%"
echo [INFO] Browser: %BROWSER% (own profile: %PROFILE%)
start "" "%BROWSER%" --user-data-dir="%PROFILE%" --no-first-run --no-default-browser-check --app="%LINK%"
rem Stop the background service when the Web UI window is closed (user decision 2026-11-13). A hidden PowerShell
rem watches for browser processes that use this profile folder (the user's normal browser has another profile and
rem is not counted) and runs 'opencode service stop' when none is left. Nothing stays on screen or in the tray.
echo [INFO] The background service stops when the Web UI window is closed. Stop it any time with: web.bat stop
start "" /min powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "$p='%PROFILE%'; $oc='%OC%'; Start-Sleep 15; while (Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'msedge.exe' -or $_.Name -eq 'chrome.exe') -and $_.CommandLine -like ('*' + $p + '*') }) { Start-Sleep 5 }; & $oc service stop"
endlocal
exit /b 0

:default_browser
echo [WARN] Edge/Chrome not found. Opening the default browser.
echo [INFO] The background service keeps running after the browser closes. Stop it with: web.bat stop
start "" "%LINK%"
endlocal
exit /b 0

:service
"%OC%" service %~1
set "RC=%ERRORLEVEL%"
pause
exit /b %RC%
