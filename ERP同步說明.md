# ERP 庫存同步功能說明

## 功能概述

此功能可以從 ERP 系統的 MS-SQL Server 資料庫直接查詢即時餅乾庫存數量，並自動更新到 Google Sheets 的「庫存狀態」工作表。

**重要說明：**
- **餅乾庫存**：只同步 Index 工作表中存在的餅乾代號
- **禮盒成品庫存**：不從 ERP 同步（因為禮盒組裝沒有進 ERP 系統），請手動更新「成品庫存」工作表

## 支援的資料庫類型

- **MS-SQL Server** (使用 pyodbc)

## 安裝步驟

### 1. 安裝資料庫驅動套件

安裝 MS-SQL Server 驅動套件：

```bash
pip install pyodbc
```

### 2. 設定 config.ini

在 `config.ini` 中已設定 `[ERP_DATABASE]` 區段，請確認資料庫連接資訊正確：

```ini
[ERP_DATABASE]
server = 192.168.98.10
database = AS_online
username = DS
password = dsc@23225889
```

### 3. 設定 SQL 查詢（可選）

如果您想使用自訂的 SQL 查詢，可以在 `config.ini` 的 `[ERP_QUERIES]` 區段中設定：

```ini
[ERP_QUERIES]
# 餅乾庫存查詢（請根據實際資料表結構調整）
cookie_inventory_query = 
    SELECT 
        cookie_code,
        current_stock as qty,
        '片' as unit
    FROM inventory_table
    WHERE item_type = 'COOKIE'
```

**重要：** 查詢結果必須包含以下欄位：
- 餅乾庫存：`cookie_code` (或 `餅乾代號`), `qty` (或 `目前庫存數量`), `unit` (或 `單位`)

**注意：** 禮盒庫存不從 ERP 同步，因此不需要設定 `box_inventory_query`

## 使用方法

### 同步餅乾庫存

```bash
python sync_inventory_from_erp.py
```

**同步邏輯：**
1. 讀取 Google Sheets 的 Index 工作表，取得所有餅乾代號
2. 從 ERP 查詢這些餅乾代號的庫存資料
3. 只同步 Index 工作表中存在的餅乾代號
4. 如果代號在「庫存狀態」工作表中已存在則更新，不存在則新增

**注意：** 禮盒成品庫存不從 ERP 同步，請手動更新「成品庫存」工作表

## 工作流程

1. **連接 Google Sheets**：讀取 Index 工作表，取得需要同步的餅乾代號列表
2. **連接 ERP 資料庫**：使用 `erp_db_helper.py` 建立資料庫連接
3. **查詢庫存資料**：執行 SQL 查詢取得即時餅乾庫存
4. **過濾資料**：只保留 Index 工作表中存在的餅乾代號
5. **更新工作表**：
   - 如果代號在「庫存狀態」工作表中已存在，更新該行的庫存數量和更新時間
   - 如果代號不存在，新增一筆新資料
6. **記錄更新時間**：自動記錄最後更新日期時間

## 資料對應

### 庫存狀態工作表（自動同步）
- **餅乾代號** ← ERP 的餅乾代號欄位（僅同步 Index 工作表中存在的代號）
- **目前庫存數量** ← ERP 的庫存數量欄位
- **單位** ← ERP 的單位欄位（預設為「片」）
- **最後更新日期** ← 自動填入當前時間

### 成品庫存工作表（手動更新）
- **禮盒代號** ← 手動輸入
- **目前庫存數量** ← 手動輸入
- **單位** ← 手動輸入（通常為「盒」）
- **最後更新日期** ← 手動輸入或自動填入

**注意：** 禮盒成品庫存不從 ERP 同步，因為禮盒組裝沒有進 ERP 系統，請手動更新「成品庫存」工作表

## 自動化排程

您可以設定定時任務（例如 Windows 工作排程器或 Linux cron）來定期執行同步：

**Windows 工作排程器範例：**
- 每小時執行一次：`python E:\Git\Cookies_Inventory\sync_inventory_from_erp.py`

**Linux cron 範例：**
```bash
# 每小時執行一次
0 * * * * cd /path/to/Cookies_Inventory && python sync_inventory_from_erp.py
```

## 關於 MCP (Model Context Protocol)

**MCP** 是 Cursor 編輯器的一個功能，允許 AI 助手連接到外部工具和服務。不過，對於這個任務：

1. **直接使用 Python 更可靠**：Python 的資料庫連接庫（如 pyodbc, pymysql）已經非常成熟和穩定
2. **更好的錯誤處理**：可以直接處理連接錯誤、查詢錯誤等
3. **更容易除錯**：可以直接看到 SQL 查詢和結果
4. **更靈活**：可以根據實際需求調整查詢邏輯

如果您想使用 MCP，需要：
1. 安裝 MCP 伺服器（例如 MCP SQL Server 伺服器）
2. 在 Cursor 中設定 MCP 連接
3. 透過 MCP 協議執行 SQL 查詢

但對於這個專案，建議直接使用 Python 腳本，因為：
- 更簡單直接
- 更容易維護
- 可以輕鬆整合到現有的 Google Sheets 更新流程中

## 疑難排解

### 問題：無法連接資料庫

**解決方案：**
1. 檢查 `config.ini` 中的連接資訊是否正確
2. 確認資料庫伺服器是否可訪問
3. 檢查防火牆設定
4. 確認已安裝對應的資料庫驅動套件

### 問題：SQL 查詢失敗

**解決方案：**
1. 檢查資料表名稱和欄位名稱是否正確
2. 確認 SQL 語法是否符合您的資料庫類型
3. 使用資料庫管理工具（如 SSMS, MySQL Workbench）先測試 SQL 查詢
4. 查看錯誤訊息中的詳細資訊

### 問題：資料未更新到 Google Sheets

**解決方案：**
1. 檢查 Google Sheets 的服務帳戶權限
2. 確認工作表名稱是否正確（「庫存狀態」）
3. 確認 Index 工作表中是否有餅乾代號
4. 查看執行日誌中的錯誤訊息

### 問題：某些餅乾代號沒有同步

**解決方案：**
1. 確認該餅乾代號是否在 Index 工作表中
2. 只有 Index 工作表中存在的餅乾代號才會被同步
3. 如果需要在 Index 中新增餅乾代號，請手動在 Index 工作表中新增

## 安全注意事項

1. **保護資料庫密碼**：`config.ini` 包含敏感資訊，請不要提交到公開的版本控制系統
2. **使用最小權限原則**：建議建立一個只有讀取權限的資料庫帳號
3. **網路安全**：如果資料庫在遠端，建議使用 VPN 或加密連接

## 下一步

1. 根據您的 ERP 資料庫類型，安裝對應的驅動套件
2. 在 `config.ini` 中設定資料庫連接資訊
3. 測試 SQL 查詢是否正確
4. 執行同步腳本測試
5. 設定自動化排程（可選）

