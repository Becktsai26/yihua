@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8

REM 檢查 Python
set PYTHON_CMD=
where python >nul 2>&1 && set PYTHON_CMD=python
if not defined PYTHON_CMD where python3 >nul 2>&1 && set PYTHON_CMD=python3
if not defined PYTHON_CMD where py >nul 2>&1 && set PYTHON_CMD=py

if not defined PYTHON_CMD (
    echo [錯誤] 找不到 Python，請先執行 install.bat
    pause
    exit /b 1
)

echo === 安裝依賴 ===
%PYTHON_CMD% -m pip install pandas openpyxl flask pyinstaller
if %errorlevel% neq 0 (
    echo [錯誤] 安裝依賴失敗
    pause
    exit /b 1
)

echo === 開始打包 ===
%PYTHON_CMD% -m PyInstaller --onefile --name "發票對帳工具" reconcile_app.py
if %errorlevel% neq 0 (
    echo [錯誤] 打包失敗
    pause
    exit /b 1
)

echo === 完成 ===
echo 執行檔位於 dist\發票對帳工具.exe
pause
