@echo off
setlocal

cd /d "%~dp0packaging-dashboard-react"

if not exist node_modules (
  echo [1/2] npm install ...
  call npm install
  if errorlevel 1 (
    echo npm install failed.
    pause
    exit /b 1
  )
)

echo [2/2] npm run dev ...
start "" cmd /c "cd /d \"%~dp0packaging-dashboard-react\" && npm run dev"
timeout /t 3 >nul
start "" http://127.0.0.1:8512

echo Dashboard opened: http://127.0.0.1:8512
endlocal
