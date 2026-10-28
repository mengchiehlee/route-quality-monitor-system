@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    py -3 "%~dp0update_dsp_monitor.py"
) else (
    python "%~dp0update_dsp_monitor.py"
)
set "run_result=%errorlevel%"
if not "%run_result%"=="0" echo Update failed. See error above.
pause
exit /b %run_result%
