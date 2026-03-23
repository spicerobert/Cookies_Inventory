# 當前工作重點 — Cookies Inventory

## 專案狀態
**已結案（開發完成）** — 本專案功能已完整實作。

## 最近 git 提交
- `448d7c5` 開發完成，結案
- `8c8a640` 合併遠端更改：保留庫存不足檢測功能說明
- `f3143aa` 將「負庫存」改為「庫存不足」，修正判斷邏輯

## 目前未提交的變更（新增文件）
- `.cursor/commands/check-inventory-calculation.md`
- `.cursor/commands/code-review-checklist.md`
- `.cursor/commands/security-audit.md`
- `.cursor/commands/test-and-fix.md`
- `.vscode/settings.json`
- `README.md` — 已修改

## 下一步
此專案本身無新功能開發計畫。主要工作為：

**整合至 Production_Scheduler**（見該專案進度）：
- 實作 `schema_cookie_inventory.sql`（MS-SQL 餅乾庫存表）
- 將 `cookies_inventory/` 核心邏輯移植/適配至 Production_Scheduler 後端
- 在 Gradio Web UI 新增餅乾庫存算料頁面

## 注意事項
- 本專案保留作為邏輯參考，整合時「邏輯不變，僅資料來源替換」
