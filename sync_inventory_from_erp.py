"""
從 ERP 系統同步餅乾庫存資料到 Google Sheets

功能說明：
- 只同步 Index 工作表中存在的餅乾代號
- 合併期初庫存（INVLC）和每日進出數量（INVLA）計算即時庫存
- 支援多庫別（SP40, SP50, SP60）
- 自動更新或新增庫存資料
"""
import sys
from datetime import datetime
from typing import List, Dict, Set, Any
from google_sheets_helper import GoogleSheetsHelper
from erp_db_helper import ERPDBHelper
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 工作表欄位定義
INVENTORY_HEADERS = ['餅乾代號', '目前庫存數量', '庫別代號', '單位', '最後更新日期']


def get_cookie_codes_from_index(sheets_helper: GoogleSheetsHelper) -> Set[str]:
    """
    從 Index 工作表取得需要同步的餅乾代號列表
    
    Args:
        sheets_helper: Google Sheets 輔助物件
        
    Returns:
        餅乾代號集合
    """
    logger.info("讀取 Index 工作表...")
    index_dict = sheets_helper.get_index_dict()
    cookie_codes = set(index_dict.get('餅乾', {}).keys())
    
    if not cookie_codes:
        logger.warning("Index 工作表中沒有餅乾代號")
    else:
        logger.info(f"Index 工作表中找到 {len(cookie_codes)} 個餅乾代號")
    
    return cookie_codes


def filter_inventory_by_index(
    inventory_data: List[Dict[str, Any]], 
    cookie_codes: Set[str]
) -> List[Dict[str, Any]]:
    """
    過濾庫存資料，只保留 Index 中存在的餅乾代號
    
    Args:
        inventory_data: 從 ERP 查詢到的庫存資料
        cookie_codes: Index 工作表中的餅乾代號集合
        
    Returns:
        過濾後的庫存資料列表
    """
    filtered = []
    skipped = 0
    
    for item in inventory_data:
        cookie_code = str(item.get('cookie_code', '')).strip()
        if cookie_code in cookie_codes:
            filtered.append(item)
        else:
            skipped += 1
    
    logger.info(f"過濾後：需同步 {len(filtered)} 筆，跳過 {skipped} 筆（不在 Index 中）")
    return filtered


def build_row_mapping(existing_data: List[List[Any]]) -> Dict[str, int]:
    """
    建立「餅乾代號+庫別代號」到行號的對應
    
    Args:
        existing_data: 現有工作表資料
        
    Returns:
        字典：key 為 "餅乾代號|庫別代號"，value 為行號
    """
    mapping = {}
    if len(existing_data) > 1:
        for idx, row in enumerate(existing_data[1:], start=2):
            if len(row) >= 3 and row[0] and row[2]:
                cookie_code = str(row[0]).strip()
                warehouse_code = str(row[2]).strip()
                key = f"{cookie_code}|{warehouse_code}"
                mapping[key] = idx
    return mapping


def convert_qty_to_float(qty: Any) -> float:
    """
    將庫存數量轉換為 float（Google Sheets API 需要可序列化的類型）
    
    Args:
        qty: 庫存數量（可能是 Decimal 或其他類型）
        
    Returns:
        float 類型的庫存數量
    """
    if hasattr(qty, '__float__'):
        return float(qty)
    return float(qty) if qty else 0.0


def ensure_headers(worksheet, sheets_helper: GoogleSheetsHelper):
    """確保工作表標題行正確"""
    existing_data = sheets_helper.read_worksheet('庫存狀態')
    if len(existing_data) == 0 or existing_data[0] != INVENTORY_HEADERS:
        worksheet.update(range_name='1:1', values=[INVENTORY_HEADERS])
        logger.info("已更新工作表標題行")


def sync_cookie_inventory() -> bool:
    """
    同步餅乾庫存到 Google Sheets 的「庫存狀態」工作表
    
    Returns:
        同步是否成功
    """
    logger.info("=" * 60)
    logger.info("開始同步餅乾庫存")
    logger.info("=" * 60)
    
    try:
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
            logger.info("查詢 ERP 餅乾庫存資料...")
            
            inventory_data = erp_db.get_cookie_inventory()
            logger.info(f"從 ERP 查詢到 {len(inventory_data)} 筆餅乾庫存資料")
            
            if not inventory_data:
                logger.warning("ERP 中未查詢到任何餅乾庫存資料")
                return False
            
            # 過濾：只保留 Index 中存在的餅乾代號
            filtered_inventory = filter_inventory_by_index(inventory_data, cookie_codes)
            
            if not filtered_inventory:
                logger.warning("沒有需要同步的餅乾庫存資料（所有代號都不在 Index 中）")
                return False
            
            # 準備更新資料
            update_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            worksheet = sheets_helper.get_worksheet('庫存狀態', create_if_not_exists=True)
            
            # 確保標題行正確
            ensure_headers(worksheet, sheets_helper)
            
            # 讀取現有資料並建立對應關係
            existing_data = sheets_helper.read_worksheet('庫存狀態')
            code_warehouse_to_row = build_row_mapping(existing_data)
            
            # 更新或新增資料
            updated_count = 0
            new_count = 0
            
            for item in filtered_inventory:
                cookie_code = str(item.get('cookie_code', '')).strip()
                if not cookie_code:
                    continue
                
                qty = convert_qty_to_float(item.get('qty', 0))
                warehouse_code = str(item.get('warehouse_code', '')).strip()
                unit = str(item.get('unit', '片')).strip()
                
                row_data = [cookie_code, qty, warehouse_code, unit, update_date]
                key = f"{cookie_code}|{warehouse_code}"
                
                if key in code_warehouse_to_row:
                    # 更新現有資料
                    row_num = code_warehouse_to_row[key]
                    cell_range = f'A{row_num}:E{row_num}'
                    worksheet.update(range_name=cell_range, values=[row_data])
                    updated_count += 1
                else:
                    # 新增資料
                    worksheet.append_rows([row_data])
                    new_count += 1
            
            logger.info(f"同步完成: 更新 {updated_count} 筆，新增 {new_count} 筆")
            logger.info("=" * 60)
            return True
            
    except Exception as e:
        logger.error(f"同步餅乾庫存失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    """
    從 ERP 系統同步餅乾庫存到 Google Sheets
    
    注意：
    - 只同步 Index 工作表中存在的餅乾代號
    - 禮盒成品庫存不從 ERP 同步，請手動更新「成品庫存」工作表
    """
    success = sync_cookie_inventory()
    sys.exit(0 if success else 1)
