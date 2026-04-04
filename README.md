# 每日發票對帳工具

將「發票系統匯出檔」與「8591 交易平台匯出檔」進行自動比對，產出對帳結果與 Excel 報表。

## 支援格式

兩個上傳欄位都支援 `.xlsx` 和 `.csv`，系統會自動辨識。

---

## 快速開始（直接用 Python 跑）

### 1. 安裝 Python

| 平台 | 方式 |
|------|------|
| **Mac** | 通常已內建。終端機輸入 `python3 --version` 確認。若沒有，到 [python.org](https://www.python.org/downloads/) 下載安裝 |
| **Windows** | 到 [python.org](https://www.python.org/downloads/) 下載安裝。**安裝時務必勾選「Add Python to PATH」** |

### 2. 下載專案

```bash
git clone https://github.com/Becktsai26/yihua.git
cd yihua
```

或直接從 GitHub 頁面點「Code > Download ZIP」解壓縮。

### 3. 安裝依賴

Mac：
```bash
pip3 install -r requirements.txt
```

Windows：
```bash
pip install -r requirements.txt
```

### 4. 啟動工具

Mac：
```bash
python3 reconcile_app.py
```

Windows：
```bash
python reconcile_app.py
```

啟動後會自動打開瀏覽器，顯示對帳介面（網址 `http://localhost:5678`）。

### 5. 操作步驟

1. 上傳「發票檔案」（從發票系統匯出的 XLSX 或 CSV）
2. 上傳「交易檔案」（從 8591 匯出的 XLSX 或 CSV）
3. 點「開始對帳」
4. 查看結果，需要時點「匯出 Excel 報表」下載

### 6. 停止工具

在終端機按 `Ctrl + C` 即可停止。

---

## 打包成執行檔（不需要安裝 Python）

如果要給其他人使用，可以打包成獨立執行檔。

**注意：Mac 打包的只能在 Mac 上跑，Windows 打包的只能在 Windows 上跑。需要在各自的系統上分別打包。**

### Mac

```bash
chmod +x build.sh
./build.sh
```

產出檔案：`dist/發票對帳工具`

使用方式：在終端機執行 `./dist/發票對帳工具`，會自動開啟瀏覽器。

### Windows

雙擊執行 `build.bat`，或在命令提示字元中執行：

```bash
build.bat
```

產出檔案：`dist\發票對帳工具.exe`

使用方式：雙擊 `發票對帳工具.exe`，會自動開啟瀏覽器。

---

## 對帳邏輯說明

- **比對依據**：發票檔的「明細備註」欄 = 交易檔的「賣場編號」欄
- **8591 紀錄**：賣場編號以 `S` 開頭的為 8591 交易
- **銀行紀錄**：賣場編號非 `S` 開頭的自動排除（不列入差異）
- **金額比對**：發票「總計」 vs 交易「金額」
- **CSV 備註行**：若 CSV 第一行不是有效交易資料，會自動跳過

## 匯出報表內容

| Sheet | 內容 |
|-------|------|
| 對帳摘要 | 筆數統計、總額比對、差額 |
| 吻合明細 | 所有對得上的逐筆紀錄 |
| 差異明細 | 金額不符 / 只在發票 / 只在交易 |
| 排除紀錄 | 被排除的銀行公司戶紀錄 |

---

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `reconcile_core.py` | 對帳核心邏輯 |
| `reconcile_app.py` | 網頁介面（Flask） |
| `requirements.txt` | Python 依賴套件 |
| `build.sh` | Mac 打包腳本 |
| `build.bat` | Windows 打包腳本 |

## 常見問題

**Q: 啟動後瀏覽器沒有自動打開？**
手動開啟瀏覽器，輸入 `http://localhost:5678`

**Q: 出現 port 5678 already in use？**
有其他程式佔用了 5678 port。先關掉之前的對帳工具，或在終端機執行：
- Mac: `lsof -i :5678` 找到 PID，再 `kill <PID>`
- Windows: `netstat -ano | findstr :5678` 找到 PID，再 `taskkill /PID <PID> /F`

**Q: Windows 上 `pip` 找不到？**
安裝 Python 時沒有勾選「Add Python to PATH」。重新安裝 Python 並勾選該選項。
