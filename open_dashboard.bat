@echo off
setlocal

set "DASHBOARD_URL=https://interojo-domestic.streamlit.app/"

start "" "%DASHBOARD_URL%"
echo Dashboard opened: %DASHBOARD_URL%

endlocal
