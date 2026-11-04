@echo off
rem CLI tools: automatic field checks, one result file result_LABEL_TIME.md (see docs\opencode)
chcp 65001 >nul
cd /d "%~dp0"
python run_all_tests.py %*
pause
