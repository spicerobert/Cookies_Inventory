# 技術背景 — Cookies Inventory

## 技術堆疊
| 層次 | 技術 |
|------|------|
| 資料來源/輸出 | Google Sheets API |
| 通知 | Line Bot（Messaging API）|
| GUI | Tkinter（`cookie_inventory_gui.py`）|
| 語言 | Python 3.12 |
| 套件管理 | **uv**（嚴格禁用 pip/poetry）|
| 認證 | Google Service Account（`service_account.json`）|

## 重要路徑
```
Cookies_Inventory/
├── cookie_inventory_gui.py      # Tkinter GUI 主程式
├── LINEBOT_Cookie_inventory.py  # Line Bot 通知主程式
├── setup_sheets.py              # Google Sheets 初始化設定
├── cookies_inventory/           # 核心邏輯模組
│   └── __pycache__/
├── Cookies/                     # 餅乾相關資料（設定/快取）
├── config.ini                   # 實際設定（不提交 git）
├── config.ini.template          # 設定範本
├── service_account.json.template # Google SA 範本
├── Line_Access_token.json       # Line Bot token（不提交 git）
├── pyproject.toml               # uv 套件設定
└── uv.lock                      # 鎖定版本
```

## 設定檔管理
- `config.ini`：從 `config.ini.template` 複製並填入實際值
- `service_account.json`：Google Service Account 金鑰（不提交 git）
- `Line_Access_token.json`：Line Bot Access Token

## 虛擬環境
```bash
uv venv --python 3.12 --seed --link-mode=symlink --clear .venv
source .venv/Scripts/activate  # Windows Git Bash
uv sync
```

## 特殊注意
- 專案目錄名 `Cookies/` vs 模組導入 `cookies_inventory`：需建立符號連結
- Google Sheets 為唯一資料介面（未來整合後改為 MS-SQL）
