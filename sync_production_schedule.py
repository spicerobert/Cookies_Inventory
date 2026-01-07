"""從 Google Sheets 讀取生產排程工作表，計算生產片數並更新回 Google Sheets。功能說明：
- 讀取生產排程工作表的所有資料
- 從 Index 工作表讀取生重資料
- 從 ERP 查詢餅乾品名（INVMB.MB002）
- 對每一行計算：生產片數 = 生產顆數 * 160000 / 生重
- 更新名稱欄位和生產片數欄位
- 按照標準順序重新排列欄位並更新回 Google Sheets"""
import sys
from typing import Dict, List, Any
from datetime import datetime, timedelta
from google_sheets_helper import GoogleSheetsHelper
from erp_db_helper import ERPDBHelper
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def _parse_date(date_value: Any) -> datetime | None:
    """將工作表日期值轉為 datetime（不含時間），失敗則回傳 None。"""
    if isinstance(date_value, datetime):
        return datetime(date_value.year, date_value.month, date_value.day)
    if not date_value:
        return None
    try:
        parts = str(date_value).split('/')
        if len(parts) == 3:
            y, m, d = map(int, parts)
            return datetime(y, m, d)
    except Exception:
        return None
    return None

def read_raw_weight_from_index(sheets_helper: GoogleSheetsHelper) -> Dict[str, float]:
    """從 Index 工作表讀取餅乾的生重資料
    Args:sheets_helper: Google Sheets 輔助物件
    Returns:字典：{餅乾代號: 生重}"""
    logger.info("讀取 Index 工作表的生重資料...")
    raw_weights = {}
    try:
        index_data = sheets_helper.read_worksheet('Index')
        if len(index_data) > 1:
            headers = index_data[0]
            # Index 工作表欄位順序：類型、代號、名稱、生重、熟重、備註
            type_idx = headers.index('類型') if '類型' in headers else 0
            code_idx = headers.index('代號') if '代號' in headers else 1
            raw_weight_idx = headers.index('生重') if '生重' in headers else 3
            
            for row in index_data[1:]:
                if len(row) > max(type_idx, code_idx, raw_weight_idx):
                    item_type = str(row[type_idx]).strip() if row[type_idx] else ''
                    cookie_code = str(row[code_idx]).strip() if row[code_idx] else ''
                    # 只處理餅乾類型的資料
                    if item_type == '餅乾' and cookie_code:
                        try:
                            raw_weight = float(row[raw_weight_idx]) if row[raw_weight_idx] else 0.0
                            if raw_weight > 0:
                                raw_weights[cookie_code] = raw_weight
                        except (ValueError, TypeError):
                            continue
            logger.info(f"從 Index 工作表讀取到 {len(raw_weights)} 種餅乾的生重資料")
    except Exception as e:
        logger.warning(f"讀取 Index 工作表生重資料失敗: {e}")
    return raw_weights

def sync_production_schedule() -> bool:
    """重新計算生產排程的生產片數並更新回 Google Sheets    
    新的欄位結構：日期、產線代號、餅乾代號、名稱、生產顆數、生產片數、預計完成日期、狀態、備註
    流程：
    1. 讀取生產排程工作表的所有資料
    2. 從 Index 工作表讀取生重資料
    3. 從 ERP 查詢餅乾品名（INVMB.MB002）
    4. 對每一行計算：生產片數 = 生產顆數 * 160000 / 生重 * 0.95（考慮5%損耗）
    5. 更新名稱欄位和生產片數欄位
    6. 按照標準順序重新排列欄位並更新回 Google Sheets    
    Returns:是否成功更新"""
    logger.info("=" * 60)
    logger.info("開始重新計算並更新生產排程的生產片數和名稱")
    logger.info("=" * 60)
    
    try:
        # 連接 Google Sheets
        sheets_helper = GoogleSheetsHelper()
        logger.info("已連接到 Google Sheets")        
        # 定義標準欄位順序
        standard_headers = ['日期', '產線代號', '餅乾代號', '名稱', '生產顆數', '生產片數', '預計完成日期', '狀態', '備註']        
        # 1. 讀取生重資料
        logger.info("讀取 Index 工作表的生重資料...")
        raw_weights = read_raw_weight_from_index(sheets_helper)
        if not raw_weights:
            logger.error("無法讀取生重資料，無法繼續")
            return False        
        # 2. 讀取生產排程工作表
        logger.info("讀取生產排程工作表...")
        schedule_data = sheets_helper.read_worksheet('生產排程')
        if len(schedule_data) < 2:
            logger.warning("生產排程工作表沒有資料")
            return False        
        headers = schedule_data[0]
        logger.info(f"現有欄位：{', '.join(headers)}")
        
        # 檢查必要欄位
        required_fields = ['日期', '產線代號', '餅乾代號', '生產顆數', '生產片數']
        missing_fields = [field for field in required_fields if field not in headers]
        if missing_fields:
            logger.error(f"缺少必要欄位：{', '.join(missing_fields)}")
            return False
        
        # 取得欄位索引（用於讀取資料）
        header_to_idx = {header: idx for idx, header in enumerate(headers)}
        date_idx = header_to_idx.get('日期', 0)
        line_code_idx = header_to_idx.get('產線代號', 1)
        cookie_code_idx = header_to_idx.get('餅乾代號', 2)
        pieces_idx = header_to_idx.get('生產顆數', 3)
        pieces_qty_idx = header_to_idx.get('生產片數', 4)
        completion_date_idx = header_to_idx.get('預計完成日期', 5)
        status_idx = header_to_idx.get('狀態', 6)
        note_idx = header_to_idx.get('備註', 7)
        name_idx = header_to_idx.get('名稱', -1)
        
        # 3. 收集所有餅乾代號，從 ERP 查詢品名
        logger.info("收集餅乾代號並從 ERP 查詢品名...")
        cookie_codes = set()
        for data_row in schedule_data[1:]:
            if cookie_code_idx < len(data_row) and data_row[cookie_code_idx]:
                cookie_code = str(data_row[cookie_code_idx]).strip()
                if cookie_code:
                    cookie_codes.add(cookie_code)
        
        cookie_names = {}
        if cookie_codes:
            try:
                with ERPDBHelper() as erp_db:
                    logger.info(f"從 ERP 查詢 {len(cookie_codes)} 個餅乾代號的品名...")
                    item_info = erp_db.get_item_info_by_codes(list(cookie_codes))
                    for code, info in item_info.items():
                        cookie_names[code] = info.get('cookie_name', '')
                    logger.info(f"成功查詢到 {len(cookie_names)} 個餅乾的品名")
            except Exception as e:
                logger.warning(f"從 ERP 查詢品名失敗: {e}，將使用 Index 工作表的資料")
                # 如果 ERP 查詢失敗，嘗試從 Index 工作表取得
                index_dict = sheets_helper.get_index_dict()
                cookie_names = index_dict.get('餅乾', {})
        
        # 4. 建立新欄位索引對應
        new_header_to_idx = {header: idx for idx, header in enumerate(standard_headers)}
        
        # 5. 重新計算每一行的生產片數並更新名稱
        logger.info("開始重新計算生產片數並更新名稱...")
        updated_rows = [standard_headers]  # 使用標準標題行
        updated_count = 0
        name_updated_count = 0
        skipped_count = 0
        error_count = 0
        
        for row_idx, row in enumerate(schedule_data[1:], start=2):  # 從第2行開始（第1行是標題）
            # 建立新行（按照標準順序）
            new_row: List[Any] = [''] * len(standard_headers)            
            # 從資料行讀取資料並填入新行
            production_date_raw = row[date_idx] if date_idx < len(row) else ''
            production_date = _parse_date(production_date_raw)
            if production_date_raw:
                new_row[new_header_to_idx['日期']] = production_date_raw
            if line_code_idx < len(row):
                new_row[new_header_to_idx['產線代號']] = row[line_code_idx]
            if cookie_code_idx < len(row):
                new_row[new_header_to_idx['餅乾代號']] = row[cookie_code_idx]
            if pieces_idx < len(row):
                new_row[new_header_to_idx['生產顆數']] = row[pieces_idx]
            if pieces_qty_idx < len(row):
                new_row[new_header_to_idx['生產片數']] = row[pieces_qty_idx]
            # 預計完成日期：若欄位有值則採用，否則使用預設「投料日期 + 2 天」
            completion_date_raw = row[completion_date_idx] if completion_date_idx < len(row) else ''
            if completion_date_raw:
                new_row[new_header_to_idx['預計完成日期']] = completion_date_raw
            elif production_date:
                default_completion = (production_date + timedelta(days=2)).strftime('%Y/%m/%d')
                new_row[new_header_to_idx['預計完成日期']] = default_completion
            if status_idx < len(row):
                new_row[new_header_to_idx['狀態']] = row[status_idx]
            if note_idx < len(row):
                new_row[new_header_to_idx['備註']] = row[note_idx]
            
            # 處理名稱欄位（優先使用 ERP 查詢結果，其次使用現有資料）
            cookie_code = str(row[cookie_code_idx]).strip() if cookie_code_idx < len(row) and row[cookie_code_idx] else ''
            if cookie_code:
                if cookie_code in cookie_names:
                    new_row[new_header_to_idx['名稱']] = cookie_names[cookie_code]
                    name_updated_count += 1
                elif name_idx >= 0 and name_idx < len(row) and row[name_idx]:
                    # 如果 ERP 查不到，保留現有的名稱
                    new_row[new_header_to_idx['名稱']] = row[name_idx]
            
            # 處理生產片數計算
            pieces_str = row[pieces_idx] if pieces_idx < len(row) else ''
            if cookie_code and pieces_str:
                try:
                    pieces = float(pieces_str) if pieces_str else 0.0
                    if pieces > 0:
                        # 計算生產片數 = 生產顆數 * 160000 / 生重 * 0.95（考慮5%損耗）
                        if cookie_code in raw_weights and raw_weights[cookie_code] > 0:
                            raw_weight = raw_weights[cookie_code]
                            qty_pieces = pieces * 160000 / raw_weight * 0.95
                            # 更新生產片數欄位（保留小數點後2位）
                            new_row[new_header_to_idx['生產片數']] = round(qty_pieces, 2)
                            updated_count += 1
                        else:
                            # 找不到生重，保留原值或設為空
                            if not new_row[new_header_to_idx['生產片數']]:
                                new_row[new_header_to_idx['生產片數']] = ''
                            skipped_count += 1
                            logger.debug(f"第 {row_idx} 行：餅乾 {cookie_code} 找不到生重，跳過")
                except (ValueError, TypeError) as e:
                    error_count += 1
                    logger.warning(f"第 {row_idx} 行：解析生產顆數失敗: {e}")            
            updated_rows.append(new_row)        
        logger.info(f"計算完成：更新生產片數 {updated_count} 筆，更新名稱 {name_updated_count} 筆，跳過 {skipped_count} 筆，錯誤 {error_count} 筆")
        
        # 6. 更新回 Google Sheets
        logger.info("開始更新 Google Sheets...")
        worksheet = sheets_helper.get_worksheet('生產排程')
        if not worksheet:
            logger.error("無法取得生產排程工作表")
            return False
        
        # 清空工作表並寫入新資料
        worksheet.clear()
        if len(updated_rows) > 0:
            num_cols = len(standard_headers)
            end_col = chr(ord('A') + num_cols - 1)
            range_name = f'A1:{end_col}{len(updated_rows)}'
            worksheet.update(range_name=range_name, values=updated_rows)
        
        logger.info(f"已成功更新 {len(updated_rows) - 1} 筆資料到生產排程工作表")
        logger.info(f"欄位順序已更新為：{', '.join(standard_headers)}")
        logger.info("=" * 60)
        logger.info("更新完成！")
        logger.info("=" * 60)
        return True
        
    except Exception as e:
        logger.error(f"重新計算並更新生產片數失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    """
    從 Google Sheets 讀取生產排程工作表，計算生產片數並更新回 Google Sheets
    
    功能：
    - 讀取生產排程工作表的所有資料
    - 從 Index 工作表讀取生重資料
    - 從 ERP 查詢餅乾品名
    - 計算生產片數 = 生產顆數 * 160000 / 生重 * 0.95（考慮5%損耗）
    - 更新名稱欄位和生產片數欄位
    - 按照標準順序重新排列欄位並更新回 Google Sheets
    
    注意：
    - 需要確保 Index 工作表有生重資料
    - 需要確保生產排程工作表有必要的欄位：日期、產線代號、餅乾代號、生產顆數、生產片數
    """
    success = sync_production_schedule()
    sys.exit(0 if success else 1)
