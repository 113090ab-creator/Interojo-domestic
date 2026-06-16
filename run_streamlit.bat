@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "PORT=8514"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "PYEXE="
set "OUT_LOG=%~dp0streamlit_8514_out.log"
set "ERR_LOG=%~dp0streamlit_8514_err.log"

if exist "%VENV_PY%" (
  set "PYEXE=%VENV_PY%"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo [ERROR] Python executable not found.
    echo [HINT] Install Python or create .venv in this folder.
    pause
    exit /b 1
  )
  for /f "delims=" %%I in ('where python') do (
    set "PYEXE=%%I"
    goto :python_ready
  )
)

:python_ready
echo [INFO] Running Streamlit on http://127.0.0.1:%PORT%
echo [INFO] STDOUT log: %OUT_LOG%
echo [INFO] STDERR log: %ERR_LOG%
echo [INFO] Press Ctrl+C in this window to stop.

:run_loop
echo ==================================================>>"%OUT_LOG%"
echo [%date% %time%] START dashboard.py on port %PORT%>>"%OUT_LOG%"
echo ==================================================>>"%ERR_LOG%"
echo [%date% %time%] START dashboard.py on port %PORT%>>"%ERR_LOG%"

"%PYEXE%" -m streamlit run dashboard.py --server.port %PORT% --server.address 127.0.0.1 --server.headless true 1>>"%OUT_LOG%" 2>>"%ERR_LOG%"
set "EXIT_CODE=%ERRORLEVEL%"
echo [%date% %time%] EXIT CODE !EXIT_CODE!>>"%ERR_LOG%"

if "!EXIT_CODE!"=="0" (
  echo [WARN] Streamlit stopped normally. Restarting in 3 seconds...
) else (
  echo [WARN] Streamlit crashed with code !EXIT_CODE!. Restarting in 3 seconds...
)

ping -n 4 127.0.0.1 >nul
goto :run_loop
