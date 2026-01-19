"""從 ERP 系統同步在製品庫存資料到 Google Sheets。功能說明：
- 只同步 Index 工作表中存在的餅乾代號
- 從製令單頭（MOCTA）查詢生產中的在製品數量
- 計算邏輯：已領數量（TA016）- 已生產數量（TA017）
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
WIP_HEADERS = ['餅乾代號', '餅乾品名', '製令單別', '製令單號', '在製品數量', '單位', '最後更新日期']

def sync_wip_inventory() -> bool:
    """同步在製品庫存到 Google Sheets 的「在製品庫存」工作表
    
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
    
    # 連接 ERP 資料庫並查詢在製品庫存
    with ERPDBHelper() as erp_db:
        logger.info("已連接到 ERP 資料庫")
        logger.info("查詢 ERP 在製品庫存資料...")
        wip_data = erp_db.get_wip_inventory()
        logger.info(f"從 ERP 查詢到 {len(wip_data)} 筆在製品庫存資料")
        
        if not wip_data:
            logger.warning("ERP 中未查詢到任何在製品庫存資料")
            return False
        
        # 過濾：只保留 Index 中存在的餅乾代號
        filtered_wip = filter_data_by_index(wip_data, cookie_codes, data_type='在製品庫存')
        if not filtered_wip:
            logger.warning("沒有需要同步的在製品庫存資料（所有代號都不在 Index 中）")
            return False
        
        # 準備更新資料
        update_date = get_update_date()
        worksheet = sheets_helper.get_worksheet('在製品庫存', create_if_not_exists=True)
        if worksheet is None:
            logger.error("無法取得或建立「在製品庫存」工作表")
            return False
        
        # 準備要同步的新資料
        processed_rows = []
        
        for item in filtered_wip:
            cookie_code = str(item.get('cookie_code', '')).strip()
            mo_type = str(item.get('mo_number_type', '')).strip()
            mo_number = str(item.get('mo_number', '')).strip()
            if not cookie_code:
                continue
            
            # 從 ERP 查詢結果取得餅乾品名
            cookie_name = str(item.get('cookie_name', '')).strip()
            # 取得在製品數量
            wip_qty = convert_qty_to_float(item.get('wip_qty', 0))
            unit = str(item.get('unit', '片')).strip()
            row_data = [cookie_code, cookie_name, mo_type, mo_number, wip_qty, unit, update_date]
            processed_rows.append(row_data)
        
        # 對資料進行排序（第一優先：餅乾代號，第二優先：製令單別，第三優先：製令單號）
        logger.info("對資料進行排序（第一優先：餅乾代號，第二優先：製令單別，第三優先：製令單號）...")
        sorted_rows = sorted(processed_rows,
            key=lambda row: (
                str(row[0]).strip() if len(row) > 0 and row[0] else '',  # 餅乾代號
                str(row[2]).strip() if len(row) > 2 and row[2] else '',  # 製令單別
                str(row[3]).strip() if len(row) > 3 and row[3] else ''   # 製令單號
            )
        )
        
        logger.info(f"排序完成：共 {len(sorted_rows)} 筆資料")
        
        # 批次寫入所有新資料
        write_worksheet_data(worksheet, WIP_HEADERS, sorted_rows, clear_first=True)
        logger.info(f"同步完成: 已刪除舊資料，寫入 {len(sorted_rows)} 筆新資料，已排序")
        return True

if __name__ == '__main__':
    """
    從 ERP 系統同步在製品庫存到 Google Sheets
    
    注意：
    - 只同步 Index 工作表中存在的餅乾代號
    - 查詢條件：開單日期 >= 20251101 且狀態碼為 '3'（生產中）
    - 計算邏輯：已領數量（TA016）- 已生產數量（TA017）
    """
    from .sync_utils import sync_with_error_handling
    success = sync_with_error_handling("同步在製品庫存", sync_wip_inventory)
    sys.exit(0 if success else 1)
