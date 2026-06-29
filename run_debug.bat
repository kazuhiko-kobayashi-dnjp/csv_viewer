@echo off
REM run_debug.bat - run CSV Viewer with a visible console to see errors.
REM Double-click this. If the app crashes, the traceback stays on screen.

setlocal
cd /d "%~dp0"

REM Clear env vars that break the Windows Python stdlib (often leaked from WSL)
set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONSTARTUP="

REM Prefer the 'py' launcher (the bare 'python' on PATH may be a broken install)
where py >nul 2>&1
if %errorlevel%==0 (
    py -3 csv_viewer.py %*
) else (
    python csv_viewer.py %*
)

echo.
echo ---- exit code: %errorlevel% ----
pause
