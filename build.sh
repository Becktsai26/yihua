#!/bin/bash
echo "=== 安裝依賴 ==="
pip3 install pandas openpyxl flask pyinstaller

echo "=== 開始打包 ==="
python3 -m PyInstaller --onefile --name "發票對帳工具" reconcile_app.py

echo "=== 完成 ==="
echo "執行檔位於 dist/發票對帳工具"
