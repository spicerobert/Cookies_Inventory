"""
測試 Google Sheets API 連接
"""
import configparser
import json
import gspread
from google.oauth2.service_account import Credentials

def load_config():
    """讀取設定檔"""
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config['GOOGLE_SHEETS']['data_sheet_url']

def get_google_sheets_client():
    """建立 Google Sheets 客戶端"""
    # 讀取服務帳戶憑證
    with open('service_account.json', 'r', encoding='utf-8') as f:
        creds_info = json.load(f)
    
    # 建立憑證物件
    credentials = Credentials.from_service_account_info(
        creds_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    
    # 建立 gspread 客戶端
    client = gspread.authorize(credentials)
    return client

def extract_sheet_id(url):
    """從 Google Sheets URL 提取 Sheet ID"""
    # URL 格式: https://docs.google.com/spreadsheets/d/{SHEET_ID}/...
    parts = url.split('/')
    sheet_id_index = parts.index('d') + 1
    if sheet_id_index < len(parts):
        return parts[sheet_id_index].split('?')[0]
    raise ValueError("無法從 URL 中提取 Sheet ID")

def test_connection():
    """測試 Google Sheets 連接"""
    try:
        print("正在讀取設定檔...")
        sheet_url = load_config()
        print(f"Sheet URL: {sheet_url}")
        
        print("\n正在建立 Google Sheets 客戶端...")
        client = get_google_sheets_client()
        
        print("\n正在開啟試算表...")
        sheet_id = extract_sheet_id(sheet_url)
        spreadsheet = client.open_by_key(sheet_id)
        print(f"[成功] 開啟試算表: {spreadsheet.title}")
        
        # 列出現有工作表
        print("\n現有工作表:")
        worksheets = spreadsheet.worksheets()
        for i, worksheet in enumerate(worksheets, 1):
            print(f"  {i}. {worksheet.title} (ID: {worksheet.id})")
        
        # 測試讀取（如果有工作表）
        if worksheets:
            print(f"\n測試讀取工作表 '{worksheets[0].title}' 的資料...")
            values = worksheets[0].get_all_values()
            if values:
                print(f"[成功] 讀取資料，共 {len(values)} 行")
                if len(values) > 0:
                    print(f"  第一行: {values[0]}")
            else:
                print("  工作表是空的")
        
        # 測試寫入（建立一個測試工作表）
        print("\n測試寫入功能...")
        try:
            test_sheet = spreadsheet.worksheet('_測試連接')
            print("  測試工作表已存在，正在刪除...")
            spreadsheet.del_worksheet(test_sheet)
        except gspread.exceptions.WorksheetNotFound:
            pass
        
        test_sheet = spreadsheet.add_worksheet(title='_測試連接', rows=10, cols=5)
        test_sheet.update(range_name='A1', values=[['測試連接', '成功！']])
        print(f"[成功] 建立並寫入測試工作表 '{test_sheet.title}'")
        
        # 清理測試工作表
        print("\n正在清理測試工作表...")
        spreadsheet.del_worksheet(test_sheet)
        print("[成功] 測試完成，所有功能正常！")
        
        return True
        
    except Exception as e:
        print(f"\n[錯誤] {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    print("=" * 50)
    print("Google Sheets API 連接測試")
    print("=" * 50)
    success = test_connection()
    if success:
        print("\n" + "=" * 50)
        print("[成功] 連接測試成功！可以開始建立工作表。")
        print("=" * 50)
    else:
        print("\n" + "=" * 50)
        print("[失敗] 連接測試失敗，請檢查設定檔和權限。")
        print("=" * 50)

