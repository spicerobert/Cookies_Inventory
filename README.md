# 餅乾禮盒生產排程系統 (Cookies Inventory Management System)

## 專案簡介

這是一個自動化生產排程系統，專為餅乾禮盒生產工廠設計。系統能夠：
- 根據組裝計畫自動建議餅乾生產排程
- 動態調整排程以因應庫存和需求變化
- 進行齊料缺口分析
- 優化生產批次以減少換線損失

## 系統架構

本系統使用 **Google Sheets** 作為資料來源和結果輸出的介面，透過 Python 進行排程計算和資料處理。

## 快速開始

### 1. 建立虛擬環境 (Python 3.12)

本專案目前使用 Python 3.12。使用 `uv` 建立虛擬環境：

```bash
uv venv --python 3.12 --seed --link-mode=symlink --clear .venv
```

### 2. 啟動虛擬環境

**Windows (Git Bash):**
```bash
source .venv/Scripts/activate
```

**Linux/macOS:**
```bash
source .venv/bin/activate
```

### 3. 安裝依賴套件

```bash
uv pip install -e .
```

如果需要安裝 ERP 資料庫連接套件（根據您的資料庫類型選擇）：

```bash
# SQL Server (預設使用)
uv pip install -e ".[sqlserver]"

# MySQL
uv pip install -e ".[mysql]"

# PostgreSQL
uv pip install -e ".[postgresql]"

# Oracle
uv pip install -e ".[oracle]"
```

### 4. 設定 Google Sheets

確保 `config.ini` 中的 Google Sheet URL 正確，並且 `service_account.json` 中的服務帳戶有適當的權限。

### 5. 測試連接

```bash
python test_google_sheets.py
```

### 6. 初始化工作表結構

```bash
python setup_sheets.py
```

這會建立系統所需的所有工作表（BOM、組裝計劃、生產排程建議等）。

### 7. 填入基礎資料

請參考 `工作表說明.md` 文件，在各工作表中填入：
- BOM（禮盒組成表）
- 生產參數
- 庫存狀態
- 組裝計劃
- 出貨預測
- 等基礎資料

### 8. 執行排程計算

（待開發）執行排程計算腳本，系統會自動產生生產排程建議。

### 9. 同步 ERP 庫存資料

從 ERP 系統同步餅乾庫存到 Google Sheets：

```bash
python sync_inventory_from_erp.py
```

**功能說明：**
- 只同步 Index 工作表中存在的餅乾代號
- 合併期初庫存（INVLC）和每日進出數量（INVLA）計算即時庫存
- 支援多庫別（SP40, SP50, SP60）
- 自動更新或新增庫存資料到「庫存狀態」工作表

**注意事項：**
- 請確保 `config.ini` 中的 ERP 資料庫連接設定正確
- 禮盒成品庫存不從 ERP 同步，請手動更新「成品庫存」工作表

## 檔案說明

- `config.ini` - 系統設定檔（包含 Google Sheet URL）
- `service_account.json` - Google Service Account 憑證（請妥善保管）
- `google_sheets_helper.py` - Google Sheets 操作工具模組
- `test_google_sheets.py` - 連接測試腳本
- `setup_sheets.py` - 工作表結構初始化腳本
- `系統設計規劃.md` - 系統設計規劃文件
- `工作表說明.md` - 各工作表的詳細說明

## 工作表列表

### 輸入資料工作表（由使用者填入）
1. **BOM** - 禮盒組成表
2. **組裝計劃** - 後段組裝作業排程
3. **出貨預測** - 預估出貨計畫
4. **庫存狀態** - 目前餅乾庫存
5. **成品庫存** - 目前禮盒成品庫存
6. **生產參數** - 餅乾生產規格
7. **組裝產能** - 後段組裝線產能
8. **產線產能** - 前段生產線產能
9. **訂單與預測** - 需求數據

### 系統輸出工作表（由系統計算並寫入）
10. **生產排程建議** - 餅乾生產排程建議
11. **組裝調整建議** - 組裝計畫調整建議
12. **齊料缺口分析** - 齊料缺口分析報告

## 下一步

1. 填入基礎資料到各工作表
2. 開發排程計算核心邏輯
3. 實作每週自動調整機制

## 注意事項

- 請妥善保管 `service_account.json`，不要將其提交到公開的版本控制系統
- 建議定期備份 Google Sheets 資料
- 系統設計細節請參考 `系統設計規劃.md`