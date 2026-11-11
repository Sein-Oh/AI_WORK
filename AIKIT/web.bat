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
rem OpenCode settings folder = this folder (AGENTS.md, skills\, opencode.json, opencode.jsonc). setup.bat also sets it
rem as a user environment variable; setting it here makes web.bat right even in a window opened before setup.
set "OPENCODE_CONFIG_DIR=%~dp0"
set "OPENCODE_CONFIG_DIR=%OPENCODE_CONFIG_DIR:~0,-1%"
if /i "%~1"=="stop" goto service
if /i "%~1"=="status" goto service

rem Work folder: the Web UI makes every session inside a project folder, and a new PC has none, so "new session" did nothing
rem (team PC 2026-11-11). The service starts in this folder; the user opens it once in the Web UI (Open project).
set "WORK=%~dp0work"
if not exist "%WORK%" mkdir "%WORK%"
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
echo [INFO] The background service keeps running after this window closes. Stop it with: web.bat stop
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
endlocal
exit /b 0

:default_browser
echo [WARN] Edge/Chrome not found. Opening the default browser.
start "" "%LINK%"
endlocal
exit /b 0

:service
"%OC%" service %~1
set "RC=%ERRORLEVEL%"
pause
exit /b %RC%
