@echo off
REM Double-click to refresh Steve's World Cup tracker with live ESPN scores.
python -m pip install --quiet --disable-pip-version-check openpyxl
python "%~dp0update_wc.py"
set "DESKTOP=%USERPROFILE%\Desktop"
if not exist "%DESKTOP%" if exist "%OneDrive%\Desktop" set "DESKTOP=%OneDrive%\Desktop"
if exist "%DESKTOP%\World Cup 2026 Tracker.xlsx" (
    start "" "%DESKTOP%\World Cup 2026 Tracker.xlsx"
) else (
    start "" "%~dp0World Cup 2026 Tracker.xlsx"
)
pause
