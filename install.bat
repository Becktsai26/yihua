@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
echo ====================================
echo   對帳系統 - 首次安裝
echo ====================================
echo.

REM 檢查 Python 是否已安裝（依序嘗試 python / python3 / py）
set PYTHON_CMD=
where python >nul 2>&1 && set PYTHON_CMD=python
if not defined PYTHON_CMD where python3 >nul 2>&1 && set PYTHON_CMD=python3
if not defined PYTHON_CMD where py >nul 2>&1 && set PYTHON_CMD=py

if not defined PYTHON_CMD (
    echo [錯誤] 找不到 Python，請先安裝：
    echo.
    echo   1. 前往 https://www.python.org/downloads/
    echo   2. 下載最新版本
    echo   3. 安裝時務必勾選「Add Python to PATH」
    echo.
    pause
    exit /b 1
)

echo [OK] Python 已安裝：
%PYTHON_CMD% --version
echo.

REM 升級 pip 避免舊版問題
echo 正在更新 pip...
%PYTHON_CMD% -m pip install --upgrade pip >nul 2>&1
echo.

REM 安裝依賴套件（用 python -m pip 確保裝到正確的環境）
echo 正在安裝依賴套件...
%PYTHON_CMD% -m pip install -r "%~dp0requirements.txt"
if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 安裝失敗，請確認網路連線正常
    pause
    exit /b 1
)

REM 驗證所有套件都能正常載入
echo.
echo 正在驗證安裝...
%PYTHON_CMD% -c "import flask; import pandas; import openpyxl; print('[OK] 所有套件驗證通過')"
if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 套件驗證失敗，請嘗試重新執行 install.bat
    pause
    exit /b 1
)

echo.
echo ====================================
echo   安裝完成！
echo   之後每次使用請雙擊 start.bat
echo ====================================
echo.
pause
