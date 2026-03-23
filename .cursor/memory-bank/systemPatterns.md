# 系統架構與模式 — Cookies Inventory

## 庫存計算邏輯
```
期初庫存
  + 生產排程（未來14天每日生產片數）
  - 組裝排程（每日消耗片數 = 禮盒數量 × BOM）
= 每日每種餅乾庫存數量
```

### 庫存不足判斷
- 任一天任一種餅乾庫存 ≤ 0 → 觸發警示
- 警示透過 Line Bot 發送

## 生產計算模式
- **集中生產**：將同款餅乾多個訂單合併計算
- **最小批量**：1000 片/批次
- **換線時間**：標準 0.5 小時
- **產能計算**：片/小時 × 可用工時

## 資料流
```
Google Sheets（ERP 同步）
    ↓ 讀取期初庫存
cookies_inventory 模組
    ├─ 讀取生產排程
    ├─ 讀取組裝排程
    └─ 計算14天庫存預估
         ├─ 輸出 Google Sheets（庫存明細）
         └─ 庫存不足 → Line Bot 警示
```

## 整合至 Production_Scheduler 的模式變更
| 項目 | 目前 | 整合後 |
|------|------|--------|
| 資料來源 | Google Sheets | MS-SQL（Procedure_List、新餅乾庫存表）|
| UI | Tkinter GUI / Line Bot | Gradio Web UI |
| 計算邏輯 | 不變 | 不變 |
| 輸出 | Google Sheets | MS-SQL + Web UI 顯示 |

## 關聯文件
- `系統設計規劃.md` — 完整系統設計（拉式生產、齊料策略等）
- `工作表說明.md` — Google Sheets 工作表結構說明
- `打包說明.md` — 打包/部署說明
- `GUI使用說明.md` — Tkinter GUI 使用方式
