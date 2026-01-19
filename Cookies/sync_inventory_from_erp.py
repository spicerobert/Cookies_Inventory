"""從 ERP 系統同步餅乾庫存資料到 Google Sheets。功能說明：
- 只同步 Index 工作表中存在的餅乾代號
- 合併期初庫存（INVLC）和每日進出數量（INVLA）計算即時庫存
- 支援多庫別（SP40, SP50, SP60, SP80）
- 自動將庫存數量統一轉換為「片」為單位（根據餅乾代號最後一碼換算）
- 每次同步會先刪除所有舊資料，再寫入最新的清單"""
import sys
from typing import List, Dict, Any
from .google_sheets_helper import GoogleSheetsHelper
from .erp_db_helper import ERPDBHelper
from .sync_utils import (
    logger,
    convert_qty_to_float,
    get_cookie_codes_from_index,
    filter_data_by_index,
    write_worksheet_data,
    get_update_date
)

# 工作表欄位定義
INVENTORY_HEADERS = ['餅乾代號', '餅乾品名', '目前庫存數量', '庫別代號', '單位', '最後更新日期']


def get_unit_conversion_factor(cookie_code: str) -> float:
    """根據餅乾代號最後一碼取得單位換算倍數
    換算規則：- C: 乘1, D: 乘2, H: 乘8, F: 乘4, A: 乘1, 其他: 乘1（預設） 
    Args: cookie_code: 餅乾代號
    Returns: 換算倍數"""
    if not cookie_code:
        return 1.0
    last_char = cookie_code[-1].upper()
    conversion_map = {'C': 1.0,'D': 2.0,'H': 8.0,'F': 4.0,'A': 1.0}
    return conversion_map.get(last_char, 1.0)

def convert_to_pieces(qty: float, cookie_code: str) -> float:
    """將庫存數量轉換為「片」為單位    
    Args: qty: 原始庫存數量, cookie_code: 餅乾代號        
    Returns: 轉換為「片」的庫存數量"""
    factor = get_unit_conversion_factor(cookie_code)
    return qty * factor

def normalize_cookie_code(cookie_code: str) -> str:
    """將餅乾代號的最後一碼統一改為 C（對應「片」的單位）    
    轉換規則：- D → C（包轉片，×2）, H → C（×8）, F → C（×4）, A → C（×1）, C → C（保持不變）, 其他: 保持不變    
    Args: cookie_code: 原始餅乾代號        
    Returns: 標準化後的餅乾代號（最後一碼為 C）
    """
    if not cookie_code or len(cookie_code) < 2:
        return cookie_code    
    last_char = cookie_code[-1].upper()    
    # 如果最後一碼是 D, H, F, A，則改為 C
    if last_char in ['D', 'H', 'F', 'A']:
        return cookie_code[:-1] + 'C'    
    # 如果已經是 C 或其他，保持不變
    return cookie_code


def sync_cookie_inventory() -> bool:
    """同步餅乾庫存到 Google Sheets 的「帳上庫存」工作表
    
    Returns:
        同步是否成功
    """
    # 連接 Google Sheets
    sheets_helper = GoogleSheetsHelper()
    logger.info("已連接到 Google Sheets")
    
    # 取得需要同步的餅乾代號
    cookie_codes = get_cookie_codes_from_index(sheets_helper)
    if not cookie_codes:
        logger.warning("無法進行同步：Index 工作表中沒有餅乾代號")
        return False
    
    # 連接 ERP 資料庫並查詢庫存
    with ERPDBHelper() as erp_db:
        logger.info("已連接到 ERP 資料庫")
        logger.info("查詢 ERP SP40, SP50, SP60, SP80 庫存資料...")
        inventory_data = erp_db.get_cookie_inventory()
        logger.info(f"從 ERP 查詢到 {len(inventory_data)} 筆庫存資料(包含餅乾和包材)")
        
        if not inventory_data:
            logger.warning("ERP 中未查詢到任何庫存資料")
            return False
        
        # 過濾：只保留 Index 中存在的餅乾代號
        filtered_inventory = filter_data_by_index(inventory_data, cookie_codes, data_type='庫存')
        if not filtered_inventory:
            logger.warning("沒有需要同步的餅乾庫存資料（所有代號都不在 Index 中）")
            return False
        
        # 準備更新資料
        update_date = get_update_date()
        worksheet = sheets_helper.get_worksheet('帳上庫存', create_if_not_exists=True)
        if worksheet is None:
            logger.error("無法取得或建立「帳上庫存」工作表")
            return False
        
        # 準備所有要同步的資料（在記憶體中處理）
        processed_rows = []
        
        for item in filtered_inventory:
            original_cookie_code = str(item.get('cookie_code', '')).strip()
            if not original_cookie_code:
                continue
            
            # 取得原始庫存數量
            original_qty = convert_qty_to_float(item.get('qty', 0))
            # 轉換為「片」為單位
            qty_in_pieces = convert_to_pieces(original_qty, original_cookie_code)
            # 標準化餅乾代號：將最後一碼統一改為 C
            cookie_code = normalize_cookie_code(original_cookie_code)
            # 從 ERP 查詢結果取得餅乾品名
            cookie_name = str(item.get('cookie_name', '')).strip()
            warehouse_code = str(item.get('warehouse_code', '')).strip()
            # 統一使用「片」作為單位
            unit = '片'
            row_data = [cookie_code, cookie_name, qty_in_pieces, warehouse_code, unit, update_date]
            processed_rows.append(row_data)
        
        # 將列表進行排序（第一優先：餅乾代號，第二優先：庫別代號）
        logger.info("對資料進行排序（第一優先：餅乾代號，第二優先：庫別代號）...")
        sorted_rows = sorted(processed_rows,
            key=lambda row: (
                str(row[0]).strip() if len(row) > 0 and row[0] else '',  # 餅乾代號
                str(row[3]).strip() if len(row) > 3 and row[3] else ''   # 庫別代號
            )
        )
        
        logger.info(f"排序完成：共 {len(sorted_rows)} 筆資料")
        
        # 批次寫入所有新資料
        write_worksheet_data(worksheet, INVENTORY_HEADERS, sorted_rows, clear_first=True)
        logger.info(f"同步完成: 已刪除舊資料，寫入 {len(sorted_rows)} 筆新資料，已排序")
        return True

if __name__ == '__main__':
    """
    從 ERP 系統同步餅乾庫存到 Google Sheets 的「帳上庫存」工作表
    
    注意：
    - 只同步 Index 工作表中存在的餅乾代號
    - 同步的資料會寫入「帳上庫存」工作表（這是從 ERP 系統查詢到的帳上庫存資料）
    - 「實盤庫存」工作表為手動更新，不會被此程式覆蓋
    - 禮盒成品庫存不從 ERP 同步，請手動更新「成品庫存」工作表
    """
    from .sync_utils import sync_with_error_handling
    success = sync_with_error_handling("同步餅乾庫存", sync_cookie_inventory)
    sys.exit(0 if success else 1)
