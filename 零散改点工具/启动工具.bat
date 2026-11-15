@echo off
setlocal EnableExtensions
set "NoDefaultCurrentDirectoryInExePath=1"
chcp 65001 >nul
cd /d "%~dp0"

"%SystemRoot%\System32\where.exe" pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%~dp0app_gui.py"
    exit /b 0
)

if exist "%~dp0零散改点整理工具.exe" (
    start "" "%~dp0零散改点整理工具.exe"
    exit /b 0
)

python "%~dp0app_gui.py"
