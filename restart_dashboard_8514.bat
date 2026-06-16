@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PORT=8514"

echo [1/3] Stop listeners on port %PORT% ...
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
  taskkill /PID %%P /F >nul 2>nul
)

echo [2/3] Start Streamlit dashboard ...
start "Streamlit Dashboard" cmd /k "%~dp0run_streamlit.bat"

echo [3/3] Wait for server ...
for /l %%S in (1,1,30) do (
  for /f "tokens=5" %%P in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
    echo [DONE] Dashboard is ready on port %PORT% ^(PID %%P^).
    start "" "http://127.0.0.1:%PORT%"
    exit /b 0
  )
  ping -n 2 127.0.0.1 >nul
)

echo [ERROR] Dashboard did not start on port %PORT%.
echo [HINT] Check logs:
echo   %~dp0streamlit_8514_out.log
echo   %~dp0streamlit_8514_err.log
exit /b 1
