"""從 ERP 系統同步在製品庫存資料到 Google Sheets。功能說明：
- 只同步 Index 工作表中存在的餅乾代號
- 從製令單頭（MOCTA）查詢生產中的在製品數量
- 計算邏輯：已領數量（TA016）- 已生產數量（TA017）
- 自動更新或新增在製品庫存資料"""
import sys
from datetime import datetime
from typing import List, Dict, Set, Any
from google_sheets_helper import GoogleSheetsHelper
from erp_db_helper import ERPDBHelper
import logging

logging.basicConfig(level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
# 工作表欄位定義
WIP_HEADERS = ['餅乾代號', '製令單別', '製令單號', '在製品數量', '單位', '最後更新日期']

def get_cookie_codes_from_index(sheets_helper: GoogleSheetsHelper) -> Set[str]:
    """從 Index 工作表取得需要同步的餅乾代號列表    
    Args:sheets_helper: Google Sheets 輔助物件        
    Returns:餅乾代號集合"""
    logger.info("讀取 Index 工作表...")
    index_dict = sheets_helper.get_index_dict()
    cookie_codes = set(index_dict.get('餅乾', {}).keys())    
    if not cookie_codes:
        logger.warning("Index 工作表中沒有餅乾代號")
    else:
        logger.info(f"Index 工作表中找到 {len(cookie_codes)} 個餅乾代號")    
    return cookie_codes

def filter_wip_by_index(wip_data: List[Dict[str, Any]],cookie_codes: Set[str]) -> List[Dict[str, Any]]:
    """過濾在製品資料，只保留 Index 中存在的餅乾代號    
    Args: wip_data: 從 ERP 查詢到的在製品資料, cookie_codes: Index 工作表中的餅乾代號集合        
    Returns: 過濾後的在製品資料列表"""
    filtered = []
    skipped = 0    
    for item in wip_data:
        cookie_code = str(item.get('cookie_code', '')).strip()
        if cookie_code in cookie_codes:
            filtered.append(item)
        else:
            skipped += 1    
    logger.info(f"過濾後：需同步 {len(filtered)} 筆，跳過 {skipped} 筆（不在 Index 中）")
    return filtered

def convert_qty_to_float(qty: Any) -> float:
    """將在製品數量轉換為 float（Google Sheets API 需要可序列化的類型）    
    Args: qty: 在製品數量（可能是 Decimal 或其他類型）        
    Returns: float 類型的在製品數量"""
    if hasattr(qty, '__float__'):
        return float(qty)
    return float(qty) if qty else 0.0

def sync_wip_inventory() -> bool:
    """同步在製品庫存到 Google Sheets 的「在製品庫存」工作表
    Returns:同步是否成功"""
    logger.info("=" * 60)
    logger.info("開始同步在製品庫存")
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
            filtered_wip = filter_wip_by_index(wip_data, cookie_codes)
            if not filtered_wip:
                logger.warning("沒有需要同步的在製品庫存資料（所有代號都不在 Index 中）")
                return False
            # 準備更新資料
            update_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            worksheet = sheets_helper.get_worksheet('在製品庫存', create_if_not_exists=True)
            # 只讀取一次現有資料
            logger.info("讀取現有在製品庫存資料...")
            existing_data = sheets_helper.read_worksheet('在製品庫存')
            # 確保標題行正確
            headers = WIP_HEADERS
            if len(existing_data) == 0 or existing_data[0] != WIP_HEADERS:
                headers = WIP_HEADERS
            else:
                headers = existing_data[0]            
            # 建立「餅乾代號|製令單別|製令單號」到資料行索引的對應
            data_rows = existing_data[1:] if len(existing_data) > 1 else []
            key_to_index = {}
            for idx, row in enumerate(data_rows):
                if len(row) >= 3 and row[0] and row[1] and row[2]:
                    cookie_code = str(row[0]).strip()
                    mo_type = str(row[1]).strip()
                    mo_number = str(row[2]).strip()
                    key = f"{cookie_code}|{mo_type}|{mo_number}"
                    key_to_index[key] = idx            
            # 準備所有要同步的資料（在記憶體中處理）
            processed_data = {}  # key: "餅乾代號|製令單別|製令單號", value: row_data
            updated_count = 0
            new_count = 0            
            for item in filtered_wip:
                cookie_code = str(item.get('cookie_code', '')).strip()
                mo_type = str(item.get('mo_number_type', '')).strip()
                mo_number = str(item.get('mo_number', '')).strip()
                if not cookie_code:
                    continue                
                # 取得在製品數量
                wip_qty = convert_qty_to_float(item.get('wip_qty', 0))
                unit = str(item.get('unit', '片')).strip()
                row_data = [cookie_code, mo_type, mo_number, wip_qty, unit, update_date]
                key = f"{cookie_code}|{mo_type}|{mo_number}"
                # 記錄處理後的資料
                if key in key_to_index:
                    updated_count += 1
                else:
                    new_count += 1                
                processed_data[key] = row_data
            
            # 合併現有資料和處理後的資料
            # 建立最終資料字典：key 為 "餅乾代號|製令單別|製令單號"，value 為 row_data
            final_data_dict = {}            
            # 先將現有資料加入字典（保留未被更新的現有資料）
            for idx, row in enumerate(data_rows):
                if len(row) >= 3 and row[0] and row[1] and row[2]:
                    cookie_code = str(row[0]).strip()
                    mo_type = str(row[1]).strip()
                    mo_number = str(row[2]).strip()
                    key = f"{cookie_code}|{mo_type}|{mo_number}"
                    # 如果這個 key 不在處理後的資料中，保留原資料
                    if key not in processed_data:
                        final_data_dict[key] = row            
            # 將處理後的資料加入字典（會覆蓋現有資料）
            for key, row_data in processed_data.items():
                final_data_dict[key] = row_data            
            # 將字典轉換為列表並進行排序（第一優先：餅乾代號，第二優先：製令單別，第三優先：製令單號）
            logger.info("對資料進行排序（第一優先：餅乾代號，第二優先：製令單別，第三優先：製令單號）...")
            sorted_rows = sorted(final_data_dict.values(),
                key=lambda row: (
                str(row[0]).strip() if len(row) > 0 and row[0] else '',  # 餅乾代號
                str(row[1]).strip() if len(row) > 1 and row[1] else '',  # 製令單別
                str(row[2]).strip() if len(row) > 2 and row[2] else ''   # 製令單號
                )
            )
            
            logger.info(f"排序完成：共 {len(sorted_rows)} 筆資料")            
            # 組合標題行和排序後的資料行
            final_data = [headers] + sorted_rows            
            # 一次性批次寫入所有資料（只使用一次 API 請求）
            logger.info("批次寫入所有資料到 Google Sheets...")
            if len(final_data) > 0:
                worksheet.clear()
                num_cols = len(headers)
                end_col = chr(ord('A') + num_cols - 1)
                range_name = f'A1:{end_col}{len(final_data)}'
                worksheet.update(range_name=range_name, values=final_data)            
            logger.info(f"同步完成: 更新 {updated_count} 筆，新增 {new_count} 筆，已排序")
            logger.info("=" * 60)
            return True
            
    except Exception as e:
        logger.error(f"同步在製品庫存失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    """
    從 ERP 系統同步在製品庫存到 Google Sheets
    
    注意：
    - 只同步 Index 工作表中存在的餅乾代號
    - 查詢條件：開單日期 >= 20251101 且狀態碼為 '3'（生產中）
    - 計算邏輯：已領數量（TA016）- 已生產數量（TA017）
    """
    success = sync_wip_inventory()
    sys.exit(0 if success else 1)
