@echo off
REM diagnose.bat - identify which Python is on PATH and whether it works.
setlocal
echo ===== PATH lookups =====
echo [where python]
where python
echo.
echo [where pythonw]
where pythonw
echo.
echo [where py]
where py
echo.
echo ===== Env vars =====
echo PYTHONHOME=%PYTHONHOME%
echo PYTHONPATH=%PYTHONPATH%
echo.
echo ===== Installs known to the py launcher =====
py -0p
echo.
echo ===== Try py -3 version =====
py -3 -c "import sys; print(sys.executable); print(sys.version)"
echo.
echo ===== Try python version =====
python -c "import sys; print(sys.executable); print(sys.version)"
echo.
echo ---- done ----
pause
