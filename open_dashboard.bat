@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PORT=8514"
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
  set "EXIST_PID=%%P"
  goto :already_running
)

start "Streamlit Dashboard" cmd /k "%~dp0run_streamlit.bat"

for /l %%S in (1,1,30) do (
  for /f "tokens=5" %%P in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
    set "CAND_PID=%%P"
    ping -n 3 127.0.0.1 >nul
    for /f "tokens=5" %%Q in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
      set "NEW_PID=%%Q"
      goto :ready
    )
  )
  ping -n 2 127.0.0.1 >nul
)

echo [ERROR] Streamlit server did not start on port %PORT%.
echo [HINT] Check log files:
echo   %~dp0streamlit_8514_out.log
echo   %~dp0streamlit_8514_err.log
exit /b 1

:already_running
echo [INFO] Dashboard already running on PID %EXIST_PID%.
start "" "http://127.0.0.1:%PORT%"
exit /b 0

:ready
echo [DONE] Dashboard ready on PID %NEW_PID%.
start "" "http://127.0.0.1:%PORT%"
exit /b 0
