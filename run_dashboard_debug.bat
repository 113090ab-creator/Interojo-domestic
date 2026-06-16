@echo off
cd /d "%~dp0"

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
  set "PYEXE=%VENV_PY%"
) else (
  set "PYEXE=python"
)

echo [INFO] Working directory: %cd%
echo [INFO] Python: %PYEXE%
echo [INFO] Starting Streamlit on http://127.0.0.1:8512
echo.

"%PYEXE%" -m streamlit run dashboard.py --server.port 8512 --server.address 0.0.0.0 --server.headless true

echo.
echo [INFO] Streamlit process exited. Check the error lines above.
pause
