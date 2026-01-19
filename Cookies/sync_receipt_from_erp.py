"""從 ERP 系統同步完工入庫資料到 Google Sheets。功能說明：
- 查詢入庫單表頭（MOCTF）和單身（MOCTG）的合併資料
- 查詢條件：TF003（入庫日期）在從今天到（今天-5天）這段期間，TF011='P104'
- 每次同步會先刪除所有舊資料，再寫入最新的清單"""
import sys
from typing import List, Dict, Any
from .google_sheets_helper import GoogleSheetsHelper
from .erp_db_helper import ERPDBHelper
from .sync_utils import (
    logger,
    convert_qty_to_float,
    format_date_yyyymmdd_to_slash,
    write_worksheet_data,
    get_update_date
)

# 工作表欄位定義
RECEIPT_HEADERS = ['入庫日期', '餅乾代號', '品名', '驗收數量', '單位', '規格', '最後更新日期']

def sync_receipt_data(days_back: int = 5) -> bool:
    """同步完工入庫資料到 Google Sheets 的「完工入庫」工作表
    
    Args:
        days_back: 往前查詢的天數（預設為5天）
        
    Returns:
        同步是否成功
    """
    # 連接 Google Sheets
    sheets_helper = GoogleSheetsHelper()
    logger.info("已連接到 Google Sheets")
    
    # 連接 ERP 資料庫並查詢入庫資料
    with ERPDBHelper() as erp_db:
        logger.info("已連接到 ERP 資料庫")
        logger.info(f"查詢 ERP 完工入庫資料（最近 {days_back} 天，TF011='P104'，TF001 IN ('5801', '5802')）...")
        receipt_data = erp_db.get_receipt_data(days_back=days_back)
        logger.info(f"從 ERP 查詢到 {len(receipt_data)} 筆完工入庫資料")
        
        if not receipt_data:
            logger.warning("ERP 中未查詢到任何完工入庫資料")
            return False
        
        # 準備更新資料
        update_date = get_update_date()
        worksheet = sheets_helper.get_worksheet('完工入庫', create_if_not_exists=True)
        
        if worksheet is None:
            logger.error("無法取得或建立「完工入庫」工作表")
            return False
        
        # 準備所有要同步的資料（在記憶體中處理）
        # 注意：同一天同一餅乾可能有多筆入庫，需要合併驗收數量
        processed_data = {}  # key: "入庫日期|餅乾代號", value: row_data
        
        for item in receipt_data:
            receipt_date = format_date_yyyymmdd_to_slash(str(item.get('receipt_date', '')).strip())
            cookie_code = str(item.get('cookie_code', '')).strip()
            cookie_name = str(item.get('cookie_name', '')).strip()
            spec = str(item.get('spec', '')).strip()
            unit = str(item.get('unit', '')).strip()
            receipt_qty = convert_qty_to_float(item.get('receipt_qty', 0))
            
            if not receipt_date or not cookie_code:
                continue
            
            key = f"{receipt_date}|{cookie_code}"
            
            # 如果已存在相同日期和餅乾代號的記錄，合併驗收數量
            if key in processed_data:
                # 累加驗收數量
                existing_qty = processed_data[key][3]  # 驗收數量在第4欄（索引3）
                processed_data[key][3] = existing_qty + receipt_qty
            else:
                row_data = [
                    receipt_date,
                    cookie_code,
                    cookie_name,
                    receipt_qty,
                    unit,
                    spec,
                    update_date
                ]
                processed_data[key] = row_data
        
        # 將字典轉換為列表並進行排序
        # 排序優先順序：入庫日期（降序）→ 餅乾代號（升序）
        logger.info("對資料進行排序（第一優先：入庫日期降序，第二優先：餅乾代號升序）...")
        
        def sort_key(row):
            """自定義排序鍵：入庫日期降序，餅乾代號升序"""
            # 入庫日期（轉換為可比較的格式，用於降序）
            date_str = str(row[0]).strip() if len(row) > 0 and row[0] else ''
            # 將日期轉換為 YYYYMMDD 格式用於排序（取負值實現降序）
            try:
                if date_str and len(date_str) == 10:  # YYYY/MM/DD
                    date_parts = date_str.split('/')
                    if len(date_parts) == 3:
                        date_int = int(date_parts[0] + date_parts[1].zfill(2) + date_parts[2].zfill(2))
                        # 使用負值實現降序
                        date_key = -date_int
                    else:
                        date_key = 0
                else:
                    date_key = 0
            except (ValueError, IndexError):
                date_key = 0
            
            return (
                date_key,  # 入庫日期（降序）
                str(row[1]).strip() if len(row) > 1 and row[1] else ''  # 餅乾代號（升序）
            )
        
        sorted_rows = sorted(processed_data.values(), key=sort_key)
        logger.info(f"排序完成：共 {len(sorted_rows)} 筆資料")
        
        # 批次寫入所有新資料
        write_worksheet_data(worksheet, RECEIPT_HEADERS, sorted_rows, clear_first=True)
        logger.info(f"同步完成: 已刪除舊資料，寫入 {len(sorted_rows)} 筆新資料，已排序")
        return True

if __name__ == '__main__':
    """
    從 ERP 系統同步完工入庫資料到 Google Sheets
    
    功能說明：
    - 查詢入庫單表頭（MOCTF）和單身（MOCTG）的合併資料
    - 查詢條件：
      * TF003（入庫日期）在從今天到（今天-5天）這段期間
      * TF011='P104'
      * TF001 IN ('5801', '5802')
    - 取出欄位：
      * 入庫日期（TF003）
      * 餅乾代號（TG004）
      * 品名（TG005）
      * 規格（TG006）
      * 單位（TG007）
      * 驗收數量（TG013）
    - 注意：同一天同一餅乾的多筆入庫會自動合併驗收數量
    """
    from .sync_utils import sync_with_error_handling
    success = sync_with_error_handling("同步完工入庫資料", sync_receipt_data)
    sys.exit(0 if success else 1)
