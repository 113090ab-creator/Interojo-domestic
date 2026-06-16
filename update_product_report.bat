@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PYEXE=.venv\Scripts\python.exe"
) else (
  set "PYEXE=python"
)

echo [INFO] Using Python: %PYEXE%
echo [INFO] Building product_total_report.html ...
"%PYEXE%" build_product_report.py

if errorlevel 1 (
  echo [ERROR] Failed to build report.
  pause
  exit /b 1
)

if not exist "product_total_report.html" (
  echo [ERROR] product_total_report.html not found.
  pause
  exit /b 1
)

echo [DONE] Report updated.
start "" "%cd%\product_total_report.html"
endlocal

