@echo off
rem Install the Python packages for the two CLI tools (ml-kit, office-cli) into WinPython.
rem Versions follow the standard environment table (AGENTS.md "user environment", docs/ENVIRONMENT.md).
rem Use the same versions on every PC: ml-kit models (joblib) are version-sensitive.
rem
rem Usage (from any folder):
rem   install_requirements.bat                  install from the pip index (internet PC or company mirror)
rem   install_requirements.bat D:\wheels        install offline from a wheel folder (no index)
rem   install_requirements.bat --download D:\wheels
rem                                             on an internet PC: download win_amd64 / Python 3.13 wheels into the folder
rem
rem Runs 'python -m pip' of the python on PATH (C:\CLI\python after setup.bat).
setlocal
set PKGS=numpy joblib pandas==3.0.5 scikit-learn==1.9.0 scipy==1.18.1 openpyxl==3.1.5 python-docx==1.2.0 python-pptx==1.0.2 matplotlib==3.11.1 pymupdf==1.28.2 olefile==0.47 pywin32==312 pytest==9.1.1

where.exe python >nul 2>&1
if errorlevel 1 (
  echo [FAIL] python not found on PATH. Run C:\CLI\setup.bat first, then open a new window.
  pause
  exit /b 1
)
for /f "tokens=*" %%P in ('where.exe python') do (
  echo [INFO] python: %%P
  goto found
)
:found

if /i "%~1"=="--download" goto download
if "%~1"=="" (
  python -m pip install %PKGS%
) else (
  if not exist "%~1" (
    echo [FAIL] wheel folder not found: %~1
    pause
    exit /b 1
  )
  python -m pip install --no-index --find-links "%~1" %PKGS%
)
set "RC=%ERRORLEVEL%"
goto done

:download
if "%~2"=="" (
  echo [FAIL] give the folder to save wheels: install_requirements.bat --download D:\wheels
  pause
  exit /b 2
)
python -m pip download --only-binary=:all: --platform win_amd64 --python-version 3.13 --implementation cp --dest "%~2" %PKGS%
set "RC=%ERRORLEVEL%"

:done
if "%RC%"=="0" (
  echo [OK] done
) else (
  echo [FAIL] pip exit code %RC%
)
pause
exit /b %RC%
