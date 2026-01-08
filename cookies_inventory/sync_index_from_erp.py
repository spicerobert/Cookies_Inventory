"""從 ERP 系統同步品名、生重、熟重資料到 Google Sheets Index 工作表。功能說明：
- 讀取 Index 工作表的所有代號
- 從 ERP 資料庫 INVMB 表查詢品名（MB002）、生重（MB104）、熟重（MB105）
- 更新 Index 工作表的對應欄位
- 如果 Index 工作表沒有「生重」和「熟重」欄位，會自動新增"""
import sys
from typing import List, Dict, Set, Any, Tuple
from .google_sheets_helper import GoogleSheetsHelper
from .erp_db_helper import ERPDBHelper
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Index 工作表標準欄位定義
INDEX_HEADERS = ['類型', '代號', '名稱', '生重', '熟重', '備註']

def get_header_index(headers: List[str], header_name: str, default: int = 0) -> int:
    """取得標題欄位索引"""
    return headers.index(header_name) if header_name in headers else default

def ensure_index_headers(worksheet, sheets_helper: GoogleSheetsHelper) -> List[str]:
    """確保 Index 工作表標題行包含所有必要欄位
    
    Args:
        worksheet: Google Sheets 工作表物件
        sheets_helper: Google Sheets 輔助物件
    
    Returns:
        標題行列表
    """
    existing_data = sheets_helper.read_worksheet('Index')
    
    if not existing_data:
        worksheet.update(range_name='1:1', values=[INDEX_HEADERS])
        logger.info("已建立 Index 工作表標題行")
        return INDEX_HEADERS
    
    existing_headers = existing_data[0] if existing_data else []
    
    if existing_headers == INDEX_HEADERS:
        return existing_headers
    
    worksheet.update(range_name='1:1', values=[INDEX_HEADERS])
    logger.info(f"已更新 Index 工作表標題行順序為: {', '.join(INDEX_HEADERS)}")
    return INDEX_HEADERS

def get_all_codes_from_index(sheets_helper: GoogleSheetsHelper) -> List[str]:
    """從 Index 工作表取得所有代號列表
    
    Args:
        sheets_helper: Google Sheets 輔助物件
    
    Returns:
        代號列表
    """
    logger.info("讀取 Index 工作表...")
    data = sheets_helper.read_worksheet('Index')
    if not data or len(data) < 2:
        logger.warning("Index 工作表沒有資料")
        return []
    
    headers = data[0]
    code_idx = get_header_index(headers, '代號', 1)
    
    codes = [
        str(row[code_idx]).strip()
        for row in data[1:]
        if len(row) > code_idx and row[code_idx]
    ]
    
    logger.info(f"從 Index 工作表讀取到 {len(codes)} 個代號")
    return codes

def safe_get_value(row: List[Any], header_to_idx: Dict[str, int], header_name: str, default_idx: int = 0) -> str:
    """安全地從資料行中取得欄位值
    
    Args:
        row: 資料行
        header_to_idx: 標題到索引的對應字典
        header_name: 欄位名稱
        default_idx: 預設索引
    
    Returns:
        欄位值（字串）
    """
    idx = header_to_idx.get(header_name, default_idx)
    return str(row[idx]).strip() if idx < len(row) and row[idx] else ''

def build_updated_row(
    old_row: List[Any],
    old_header_to_idx: Dict[str, int],
    new_header_to_idx: Dict[str, int],
    item_info: Dict[str, Dict[str, Any]]
) -> Tuple[List[Any], bool]:
    """構建更新後的資料行
    
    Args:
        old_row: 舊資料行
        old_header_to_idx: 舊標題到索引的對應
        new_header_to_idx: 新標題到索引的對應
        item_info: ERP 查詢結果
    
    Returns:
        (更新後的資料行, 是否已更新)
    """
    # 從舊資料行中取得各欄位的值
    old_values = {
        header: safe_get_value(old_row, old_header_to_idx, header, idx)
        for idx, header in enumerate(INDEX_HEADERS)
    }
    
    # 按照新順序構建新行
    new_row = [''] * len(INDEX_HEADERS)
    for header, idx in new_header_to_idx.items():
        new_row[idx] = old_values.get(header, '')
    
    # 如果找到代號，從 ERP 查詢結果更新名稱、生重、熟重
    code = old_values.get('代號', '').strip()
    if code and code in item_info:
        info = item_info[code]
        new_row[new_header_to_idx['名稱']] = info.get('cookie_name', '')
        new_row[new_header_to_idx['生重']] = info['raw_weight'] if info.get('raw_weight', 0) > 0 else ''
        new_row[new_header_to_idx['熟重']] = info['cooked_weight'] if info.get('cooked_weight', 0) > 0 else ''
        return new_row, True
    
    return new_row, False

def write_worksheet_data(worksheet, headers: List[str], rows: List[List[Any]]) -> None:
    """將資料寫入工作表
    
    Args:
        worksheet: Google Sheets 工作表物件
        headers: 標題行
        rows: 資料行列表
    """
    if not rows or worksheet is None:
        return
    
    final_data = [headers] + rows
    num_cols = len(headers)
    end_col = chr(ord('A') + num_cols - 1)
    range_name = f'A1:{end_col}{len(final_data)}'
    worksheet.update(range_name=range_name, values=final_data)

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
        
        # 讀取現有資料
        existing_data = sheets_helper.read_worksheet('Index')
        if len(existing_data) < 2:
            logger.warning("Index 工作表沒有資料行")
            return False
        
        # 取得現有標題（用於映射舊資料）
        old_headers = existing_data[0]
        old_header_to_idx = {header: idx for idx, header in enumerate(old_headers)}
        new_header_to_idx = {header: idx for idx, header in enumerate(headers)}
        
        # 取得所有代號
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
            
            # 處理每一行資料
            updated_rows = []
            updated_count = 0
            not_found_count = 0
            
            for old_row in existing_data[1:]:
                new_row, is_updated = build_updated_row(
                    old_row, old_header_to_idx, new_header_to_idx, item_info
                )
                updated_rows.append(new_row)
                if is_updated:
                    updated_count += 1
                elif new_row[new_header_to_idx['代號']].strip():
                    not_found_count += 1
        
        # 批次更新所有資料
        write_worksheet_data(worksheet, headers, updated_rows)
        
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
