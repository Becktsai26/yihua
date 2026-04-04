@echo off
chcp 65001 >nul 2>&1
echo ====================================
echo   對帳系統 - 首次安裝
echo ====================================
echo.

REM 檢查 Python 是否已安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [錯誤] 找不到 Python，請先安裝：
    echo   https://www.python.org/downloads/
    echo   安裝時務必勾選「Add Python to PATH」
    echo.
    pause
    exit /b 1
)

echo [OK] Python 已安裝：
python --version
echo.

REM 安裝依賴套件
echo 正在安裝依賴套件...
pip install -r "%~dp0requirements.txt"
if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 安裝失敗，請確認網路連線正常
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
