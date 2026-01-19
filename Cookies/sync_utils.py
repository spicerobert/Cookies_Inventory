"""同步模組共用工具函數"""
import logging
from datetime import datetime
from typing import List, Dict, Set, Any, Optional, Callable
from .google_sheets_helper import GoogleSheetsHelper
import gspread

# 統一的 logging 設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def convert_qty_to_float(qty: Any) -> float:
    """將數量轉換為 float（Google Sheets API 需要可序列化的類型）
    
    Args:
        qty: 數量（可能是 Decimal 或其他類型）
        
    Returns:
        float 類型的數量
    """
    if hasattr(qty, '__float__'):
        return float(qty)
    return float(qty) if qty else 0.0


def get_cookie_codes_from_index(sheets_helper: GoogleSheetsHelper) -> Set[str]:
    """從 Index 工作表取得需要同步的餅乾代號列表
    
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


def write_worksheet_data(
    worksheet: gspread.Worksheet,
    headers: List[str],
    rows: List[List[Any]],
    clear_first: bool = True
) -> None:
    """批次寫入資料到 Google Sheets 工作表
    
    Args:
        worksheet: Google Sheets 工作表物件
        headers: 標題行
        rows: 資料行列表
        clear_first: 是否先清空工作表
    """
    if not rows or worksheet is None:
        return
    
    if clear_first:
        logger.info("清空工作表，刪除所有舊資料...")
        worksheet.clear()
    
    final_data = [headers] + rows
    num_cols = len(headers)
    end_col = chr(ord('A') + num_cols - 1)
    range_name = f'A1:{end_col}{len(final_data)}'
    
    logger.info("批次寫入資料到 Google Sheets...")
    worksheet.update(range_name=range_name, values=final_data)
    logger.info(f"已成功寫入 {len(rows)} 筆資料")


def get_update_date() -> str:
    """取得當前時間的格式化字串（用於最後更新日期欄位）
    
    Returns:
        格式化的日期時間字串（格式：YYYY-MM-DD HH:MM:SS）
    """
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def sync_with_error_handling(
    sync_name: str,
    sync_func: Callable[[], bool]
) -> bool:
    """執行同步函數並統一處理錯誤
    
    Args:
        sync_name: 同步操作名稱（用於日誌）
        sync_func: 同步函數
        
    Returns:
        同步是否成功
    """
    logger.info("=" * 60)
    logger.info(f"開始{sync_name}")
    logger.info("=" * 60)
    
    try:
        success = sync_func()
        if success:
            logger.info("=" * 60)
        return success
    except Exception as e:
        logger.error(f"{sync_name}失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        logger.info("=" * 60)
        return False


def filter_data_by_index(
    data: List[Dict[str, Any]],
    cookie_codes: Set[str],
    cookie_code_key: str = 'cookie_code',
    data_type: str = '資料'
) -> List[Dict[str, Any]]:
    """過濾資料，只保留 Index 中存在的餅乾代號
    
    Args:
        data: 從 ERP 查詢到的資料
        cookie_codes: Index 工作表中的餅乾代號集合
        cookie_code_key: 資料字典中餅乾代號的鍵名
        data_type: 資料類型名稱（用於日誌）
        
    Returns:
        過濾後的資料列表
    """
    filtered = []
    skipped = 0
    
    for item in data:
        cookie_code = str(item.get(cookie_code_key, '')).strip()
        if cookie_code in cookie_codes:
            filtered.append(item)
        else:
            skipped += 1
    
    logger.info(f"過濾後：需同步 {len(filtered)} 筆，跳過 {skipped} 筆（不在 Index 中）")
    return filtered


def format_date_yyyymmdd_to_slash(date_str: str) -> str:
    """格式化日期為 YYYY/MM/DD 格式
    
    Args:
        date_str: 日期字串（格式：YYYYMMDD）
        
    Returns:
        格式化後的日期字串（格式：YYYY/MM/DD）
    """
    if not date_str or len(date_str) != 8:
        return date_str
    
    try:
        year = date_str[:4]
        month = date_str[4:6]
        day = date_str[6:8]
        return f"{year}/{month}/{day}"
    except Exception:
        return date_str
