@echo off
rem ppt-skills installer for Windows (double-click)
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (py -3 install.py %*) || (python install.py %*)
echo.
pause
