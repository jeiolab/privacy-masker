@echo off
chcp 949 >nul
cd /d "%~dp0"
if not exist "logs" mkdir "logs"
set "PYW=pythonw.exe"
where pythonw.exe >nul 2>&1 || set "PYW=pyw.exe"
start "" /b %PYW% "app.py" 2>>"logs\실행오류.txt"
exit /b
