"""從 ERP 系統同步品名、生重、熟重資料到 Google Sheets Index 工作表。功能說明：
- 讀取 Index 工作表的所有代號
- 從 ERP 資料庫 INVMB 表查詢品名（MB002）、生重（MB104）、熟重（MB105）
- 更新 Index 工作表的對應欄位
- 如果 Index 工作表沒有「生重」和「熟重」欄位，會自動新增"""
import sys
from typing import List, Dict, Set, Any
from google_sheets_helper import GoogleSheetsHelper
from erp_db_helper import ERPDBHelper
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Index 工作表標準欄位定義
INDEX_HEADERS = ['類型', '代號', '名稱', '生重', '熟重', '備註']

def get_all_codes_from_index(sheets_helper: GoogleSheetsHelper) -> List[str]:
    """從 Index 工作表取得所有代號列表
    Args: sheets_helper: Google Sheets 輔助物件
    Returns: 代號列表"""
    logger.info("讀取 Index 工作表...")
    data = sheets_helper.read_worksheet('Index')
    if not data or len(data) < 2:
        logger.warning("Index 工作表沒有資料")
        return []    
    codes = []
    for row in data[1:]:  # 跳過標題行
        if len(row) >= 2 and row[1]:  # 至少要有「類型」和「代號」欄位
            code = str(row[1]).strip()
            if code:
                codes.append(code)    
    logger.info(f"從 Index 工作表讀取到 {len(codes)} 個代號")
    return codes

def ensure_index_headers(worksheet, sheets_helper: GoogleSheetsHelper):
    """確保 Index 工作表標題行包含所有必要欄位（按照正確順序：類型、代號、名稱、生重、熟重、備註）"""
    existing_data = sheets_helper.read_worksheet('Index')
    
    if len(existing_data) == 0:
        # 如果工作表是空的，直接寫入標題
        worksheet.update(range_name='1:1', values=[INDEX_HEADERS])
        logger.info("已建立 Index 工作表標題行")
        return INDEX_HEADERS
    
    existing_headers = existing_data[0] if existing_data else []
    
    # 標準順序：類型、代號、名稱、生重、熟重、備註
    required_headers = ['類型', '代號', '名稱', '生重', '熟重', '備註']
    
    # 檢查現有標題是否與標準順序一致
    if existing_headers == required_headers:
        return existing_headers
    
    # 如果不一致，使用標準順序（確保所有欄位都存在）
    new_headers = []
    for header in required_headers:
        if header in existing_headers:
            new_headers.append(header)
        else:
            new_headers.append(header)
    
    # 更新標題行
    worksheet.update(range_name='1:1', values=[new_headers])
    if existing_headers != new_headers:
        logger.info(f"已更新 Index 工作表標題行順序為: {', '.join(new_headers)}")
    
    return new_headers

def sync_index_from_erp() -> bool:
    """同步品名、生重、熟重到 Google Sheets 的 Index 工作表
    
    Returns:
        同步是否成功
    """
    logger.info("=" * 60)
    logger.info("開始同步 Index 工作表資料")
    logger.info("=" * 60)
    
    try:
        # 連接 Google Sheets
        sheets_helper = GoogleSheetsHelper()
        logger.info("已連接到 Google Sheets")
        
        # 取得 Index 工作表
        worksheet = sheets_helper.get_worksheet('Index', create_if_not_exists=True)
        
        # 確保標題行正確
        headers = ensure_index_headers(worksheet, sheets_helper)
        if not headers:
            headers = INDEX_HEADERS
        
        # 讀取現有資料
        existing_data = sheets_helper.read_worksheet('Index')
        if len(existing_data) < 2:
            logger.warning("Index 工作表沒有資料行")
            return False
        
        # 取得現有標題（用於映射舊資料）
        old_headers = existing_data[0] if existing_data else []
        old_header_to_idx = {header: idx for idx, header in enumerate(old_headers)}
        
        # 取得所有代號（從代號欄位讀取）
        codes = get_all_codes_from_index(sheets_helper)
        if not codes:
            logger.warning("Index 工作表中沒有代號")
            return False
        
        # 連接 ERP 資料庫並查詢資料
        with ERPDBHelper() as erp_db:
            logger.info("已連接到 ERP 資料庫")
            logger.info("查詢 ERP 品名、生重、熟重資料...")
            item_info = erp_db.get_item_info_by_codes(codes)
            logger.info(f"從 ERP 查詢到 {len(item_info)} 個代號的資訊")
            
            if not item_info:
                logger.warning("ERP 中未查詢到任何資料")
                return False
            
            # 新順序的欄位索引對應：類型、代號、名稱、生重、熟重、備註
            header_to_idx = {header: idx for idx, header in enumerate(headers)}
            type_idx = header_to_idx.get('類型', 0)
            code_idx = header_to_idx.get('代號', 1)
            name_idx = header_to_idx.get('名稱', 2)
            raw_weight_idx = header_to_idx.get('生重', 3)
            cooked_weight_idx = header_to_idx.get('熟重', 4)
            note_idx = header_to_idx.get('備註', 5)
            
            # 準備更新資料
            updated_count = 0
            not_found_count = 0
            updated_rows = []
            
            # 處理每一行資料（按照新順序重新構建）
            for old_row in existing_data[1:]:  # 跳過標題行
                # 從舊資料行中取得各欄位的值
                old_type = old_row[old_header_to_idx.get('類型', 0)] if old_header_to_idx.get('類型', 0) < len(old_row) else ''
                old_code = old_row[old_header_to_idx.get('代號', 1)] if old_header_to_idx.get('代號', 1) < len(old_row) else ''
                old_name = old_row[old_header_to_idx.get('名稱', 2)] if old_header_to_idx.get('名稱', 2) < len(old_row) else ''
                old_raw_weight = old_row[old_header_to_idx.get('生重', 3)] if old_header_to_idx.get('生重', 3) < len(old_row) else ''
                old_cooked_weight = old_row[old_header_to_idx.get('熟重', 4)] if old_header_to_idx.get('熟重', 4) < len(old_row) else ''
                old_note = old_row[old_header_to_idx.get('備註', 5)] if old_header_to_idx.get('備註', 5) < len(old_row) else ''
                
                # 按照新順序構建新行：類型、代號、名稱、生重、熟重、備註
                new_row = [''] * len(headers)
                new_row[type_idx] = old_type
                new_row[code_idx] = old_code
                new_row[name_idx] = old_name
                new_row[raw_weight_idx] = old_raw_weight
                new_row[cooked_weight_idx] = old_cooked_weight
                new_row[note_idx] = old_note
                
                # 如果找到代號，從 ERP 查詢結果更新名稱、生重、熟重
                code = str(old_code).strip() if old_code else ''
                if code and code in item_info:
                    info = item_info[code]
                    # 更新品名（名稱欄位）
                    new_row[name_idx] = info['cookie_name']
                    # 更新生重
                    new_row[raw_weight_idx] = info['raw_weight'] if info['raw_weight'] > 0 else ''
                    # 更新熟重
                    new_row[cooked_weight_idx] = info['cooked_weight'] if info['cooked_weight'] > 0 else ''
                    
                    updated_rows.append(new_row)
                    updated_count += 1
                else:
                    # 即使沒有找到資料，也要保留原行（按照新順序）
                    updated_rows.append(new_row)
                    if code:
                        not_found_count += 1
        
        # 批次更新所有資料
        if updated_rows:
            # 組合標題行和更新後的資料行
            final_data = [headers] + updated_rows
            
            # 計算更新範圍
            num_cols = len(headers)
            end_col = chr(ord('A') + num_cols - 1)
            range_name = f'A1:{end_col}{len(final_data)}'
            
            # 批次寫入
            worksheet.update(range_name=range_name, values=final_data)
        
        logger.info(f"同步完成: 更新 {updated_count} 筆，未找到 {not_found_count} 筆")
        logger.info("=" * 60)
        return True
            
    except Exception as e:
        logger.error(f"同步 Index 工作表失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    """
    從 ERP 系統同步品名、生重、熟重到 Google Sheets Index 工作表
    
    注意：
    - 只更新 Index 工作表中已存在的代號
    - 如果 Index 工作表沒有「生重」和「熟重」欄位，會自動新增
    - 品名會更新到「名稱」欄位
    """
    success = sync_index_from_erp()
    sys.exit(0 if success else 1)
