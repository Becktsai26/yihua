@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

REM 找到可用的 Python 指令
set PYTHON_CMD=
where python >nul 2>&1 && set PYTHON_CMD=python
if not defined PYTHON_CMD where python3 >nul 2>&1 && set PYTHON_CMD=python3
if not defined PYTHON_CMD where py >nul 2>&1 && set PYTHON_CMD=py

if not defined PYTHON_CMD (
    echo [錯誤] 找不到 Python，請先執行 install.bat
    pause
    exit /b 1
)

REM 先殺掉佔用 5678 port 的舊進程
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5678 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
timeout /t 1 /nobreak >nul

%PYTHON_CMD% reconcile_app.py
