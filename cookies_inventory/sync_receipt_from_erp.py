"""從 ERP 系統同步完工入庫資料到 Google Sheets。功能說明：
- 查詢入庫單表頭（MOCTF）和單身（MOCTG）的合併資料
- 查詢條件：TF003（入庫日期）在從今天到（今天-5天）這段期間，TF011='P104'
- 自動更新或新增完工入庫資料"""
import sys
from datetime import datetime
from typing import List, Dict, Set, Any
from .google_sheets_helper import GoogleSheetsHelper
from .erp_db_helper import ERPDBHelper
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 工作表欄位定義
RECEIPT_HEADERS = ['入庫日期', '餅乾代號', '品名', '驗收數量', '單位', '規格', '最後更新日期']

def format_receipt_date(date_str: str) -> str:
    """格式化入庫日期為 YYYY/MM/DD 格式
    
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

def convert_qty_to_float(qty: Any) -> float:
    """將驗收數量轉換為 float（Google Sheets API 需要可序列化的類型）
    
    Args:
        qty: 驗收數量（可能是 Decimal 或其他類型）
        
    Returns:
        float 類型的驗收數量
    """
    if hasattr(qty, '__float__'):
        return float(qty)
    return float(qty) if qty else 0.0

def sync_receipt_data(days_back: int = 5) -> bool:
    """同步完工入庫資料到 Google Sheets 的「完工入庫」工作表
    
    Args:
        days_back: 往前查詢的天數（預設為5天）
        
    Returns:
        同步是否成功
    """
    logger.info("=" * 60)
    logger.info("開始同步完工入庫資料")
    logger.info("=" * 60)
    
    try:
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
            update_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            worksheet = sheets_helper.get_worksheet('完工入庫', create_if_not_exists=True)
            
            if worksheet is None:
                logger.error("無法取得或建立「完工入庫」工作表")
                return False
            
            # 讀取現有資料
            logger.info("讀取現有完工入庫資料...")
            existing_data = sheets_helper.read_worksheet('完工入庫')
            
            # 確保標題行正確
            headers = RECEIPT_HEADERS
            if len(existing_data) == 0 or existing_data[0] != RECEIPT_HEADERS:
                headers = RECEIPT_HEADERS
            else:
                headers = existing_data[0]
            
            # 建立「入庫日期|餅乾代號」到資料行索引的對應（用於去重）
            data_rows = existing_data[1:] if len(existing_data) > 1 else []
            key_to_index = {}
            for idx, row in enumerate(data_rows):
                # 新格式（7欄）：入庫日期、餅乾代號、品名、驗收數量、單位、規格、最後更新日期
                # 舊格式（9欄）：入庫日期、單別、單號、餅乾代號、品名、規格、單位、驗收數量、最後更新日期
                if len(row) >= 2 and row[0]:
                    receipt_date = str(row[0]).strip()
                    # 根據欄位數量判斷格式
                    if len(row) >= 9:
                        # 舊格式：餅乾代號在第4欄（索引3）
                        cookie_code = str(row[3]).strip() if len(row) > 3 and row[3] else ''
                    else:
                        # 新格式：餅乾代號在第2欄（索引1）
                        cookie_code = str(row[1]).strip() if len(row) > 1 and row[1] else ''
                    
                    if receipt_date and cookie_code:
                        key = f"{receipt_date}|{cookie_code}"
                        key_to_index[key] = idx
            
            # 準備所有要同步的資料（在記憶體中處理）
            # 注意：同一天同一餅乾可能有多筆入庫，需要合併驗收數量
            processed_data = {}  # key: "入庫日期|餅乾代號", value: row_data
            updated_count = 0
            new_count = 0
            
            for item in receipt_data:
                receipt_date = format_receipt_date(str(item.get('receipt_date', '')).strip())
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
                    updated_count += 1
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
                    if key in key_to_index:
                        updated_count += 1
                    else:
                        new_count += 1
            
            # 合併現有資料和處理後的資料
            # 建立最終資料字典：key 為 "入庫日期|餅乾代號"，value 為 row_data
            final_data_dict = {}
            
            # 先將現有資料加入字典（保留未被更新的現有資料）
            for idx, row in enumerate(data_rows):
                if len(row) >= 2 and row[0]:
                    receipt_date = str(row[0]).strip()
                    # 根據欄位數量判斷格式
                    if len(row) >= 9:
                        # 舊格式：餅乾代號在第4欄（索引3）
                        cookie_code = str(row[3]).strip() if len(row) > 3 and row[3] else ''
                    else:
                        # 新格式：餅乾代號在第2欄（索引1）
                        cookie_code = str(row[1]).strip() if len(row) > 1 and row[1] else ''
                    
                    if receipt_date and cookie_code:
                        key = f"{receipt_date}|{cookie_code}"
                        # 如果這個 key 不在處理後的資料中，保留原資料（但需要轉換為新格式）
                        if key not in processed_data:
                            # 如果是舊格式（9欄），轉換為新格式
                            if len(row) >= 9:  # 舊格式有9欄
                                # 舊格式：入庫日期、單別、單號、餅乾代號、品名、規格、單位、驗收數量、最後更新日期
                                # 新格式：入庫日期、餅乾代號、品名、驗收數量、單位、規格、最後更新日期
                                final_data_dict[key] = [
                                    row[0],  # 入庫日期
                                    row[3],  # 餅乾代號
                                    row[4] if len(row) > 4 else '',  # 品名
                                    row[7] if len(row) > 7 else 0,   # 驗收數量
                                    row[6] if len(row) > 6 else '',  # 單位
                                    row[5] if len(row) > 5 else '',  # 規格
                                    row[8] if len(row) > 8 else ''    # 最後更新日期
                                ]
                            elif len(row) == 7:
                                # 可能是舊的新格式（7欄但順序不同）
                                # 檢查是否為舊順序：入庫日期、餅乾代號、品名、規格、單位、驗收數量、最後更新日期
                                # 新順序：入庫日期、餅乾代號、品名、驗收數量、單位、規格、最後更新日期
                                # 如果第4欄（索引3）是規格（通常是文字），第5欄（索引4）是單位，第6欄（索引5）是數字，則是舊順序
                                try:
                                    # 嘗試判斷：如果索引5是數字，可能是舊順序
                                    float(str(row[5]))
                                    # 舊順序：需要重新排列
                                    final_data_dict[key] = [
                                        row[0],  # 入庫日期
                                        row[1],  # 餅乾代號
                                        row[2],  # 品名
                                        row[5],  # 驗收數量（從索引5移到索引3）
                                        row[4],  # 單位（從索引4移到索引4）
                                        row[3],  # 規格（從索引3移到索引5）
                                        row[6] if len(row) > 6 else ''  # 最後更新日期
                                    ]
                                except (ValueError, TypeError, IndexError):
                                    # 已經是正確的新順序
                                    final_data_dict[key] = row
                            else:
                                # 已經是新格式
                                final_data_dict[key] = row
            
            # 將處理後的資料加入字典（會覆蓋現有資料）
            for key, row_data in processed_data.items():
                final_data_dict[key] = row_data
            
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
            
            sorted_rows = sorted(final_data_dict.values(), key=sort_key)
            
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
        logger.error(f"同步完工入庫資料失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

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
    success = sync_receipt_data(days_back=5)
    sys.exit(0 if success else 1)
