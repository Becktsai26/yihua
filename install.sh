#!/bin/bash
echo "===================================="
echo "  對帳系統 - 首次安裝"
echo "===================================="
echo

# 檢查 Python
if ! command -v python3 &> /dev/null; then
    echo "[錯誤] 找不到 Python3，請先安裝："
    echo "  https://www.python.org/downloads/"
    exit 1
fi

echo "[OK] Python 已安裝：$(python3 --version)"
echo

# 安裝依賴
echo "正在安裝依賴套件..."
pip3 install -r "$(dirname "$0")/requirements.txt"

if [ $? -ne 0 ]; then
    echo
    echo "[錯誤] 安裝失敗，請確認網路連線正常"
    exit 1
fi

echo
echo "===================================="
echo "  安裝完成！"
echo "  之後每次使用請執行 bash start.sh"
echo "===================================="
