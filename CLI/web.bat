@echo off
rem OpenCode Web UI launcher (double-click). Starts 'opencode web' on this PC only (127.0.0.1)
rem and opens the browser. Usage: web.bat [work folder] [port]
rem   work folder: folder OpenCode works in (default: %USERPROFILE%\Documents)
rem   port: default 4096. Close this window to stop the server.
setlocal
set "OC=%~dp0opencode.exe"
if not exist "%OC%" (
  echo [FAIL] opencode.exe not found: %OC%
  pause
  exit /b 1
)
set "WORK=%~1"
if "%WORK%"=="" set "WORK=%USERPROFILE%\Documents"
if not exist "%WORK%\" (
  echo [FAIL] work folder not found: %WORK%
  pause
  exit /b 1
)
set "PORT=%~2"
if "%PORT%"=="" set "PORT=4096"
netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul
if not errorlevel 1 (
  echo [INFO] port %PORT% is already in use. Opening http://127.0.0.1:%PORT%/
  echo [INFO] If that is not OpenCode, run: web.bat "%WORK%" 4097
  start "" "http://127.0.0.1:%PORT%/"
  exit /b 0
)
cd /d "%WORK%"
echo [INFO] OpenCode Web UI: http://127.0.0.1:%PORT%/  (work folder: %WORK%)
echo [INFO] Keep this window open while using the Web UI. Close it to stop.
"%OC%" web --port %PORT% --hostname 127.0.0.1
echo [INFO] OpenCode server stopped (exit %ERRORLEVEL%).
pause
endlocal
