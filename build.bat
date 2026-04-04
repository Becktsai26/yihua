@echo off
echo === 安裝依賴 ===
pip install pandas openpyxl pyinstaller

echo === 開始打包 ===
pyinstaller --onefile --windowed --name "發票對帳工具" reconcile_app.py

echo === 完成 ===
echo 執行檔位於 dist\發票對帳工具.exe
pause
