# 專案簡介 — Cookies Inventory

## 核心目標
自動化餅乾庫存算料系統，專為餅乾禮盒生產工廠設計：
- 從 ERP 系統同步餅乾庫存資料到 Google Sheets
- 根據期初庫存、生產排程、組裝排程，計算未來 14 天每日每種餅乾庫存
- 檢測庫存不足（負庫存）並產生警示
- 計算並更新生產排程的生產片數

## 現況
- **狀態：開發完成（結案）** — git: `開發完成，結案`
- 目前以 Google Sheets 為資料來源/輸出介面

## 未來計畫
**整合至 Production_Scheduler**：
- 資料來源：Google Sheets → **MS-SQL**
- UI：獨立腳本 → **Gradio Web UI**（與生產排程共用）
- 邏輯與流程不變，僅底層資料來源替換
- 整合文件：`E:/Git/Production_Scheduler/gradio_scheduler/docs/COOKIE_INVENTORY_INTEGRATION_ARCHITECTURE.md`

## 專案路徑
- Git repo：`E:/Git/Cookies_Inventory`
- 主程式：`LINEBOT_Cookie_inventory.py`（Line Bot 通知）
- GUI：`cookie_inventory_gui.py`
- 核心邏輯：`cookies_inventory/`（Python 模組）
- 設定：`config.ini`（從 `config.ini.template` 複製）

## 系統核心邏輯
- **拉式生產**：後段組裝拉動前段餅乾生產
- **集中生產**：同款餅乾需求合併，最小批量 1000 片
- **3條產線**：每週 6 天、每天 10 小時
- **保存期限**：120 天（消費者需求 ≥45 天）
