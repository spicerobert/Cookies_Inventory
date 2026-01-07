"""餅乾庫存算料系統
功能說明：
- 從Google Sheets讀取今天的期初庫存（從「庫存狀態」工作表）
  （注意：此工作表的資料應該已經過手動調整）
- 讀取BOM表、生產排程、組裝排程
- 計算未來14天每一天每種餅乾的庫存數量
- 檢測負庫存（餅乾不足）的情況（包含在「庫存預估明細」工作表的「是否負庫存」和「缺口數量」欄位）
- 輸出結果到「庫存預估明細」工作表
執行流程：
1. 先執行 sync_inventory_from_erp.py 從ERP同步資料(或省略)
2. 手動更新 Google Sheets 中的「庫存狀態」工作表
3. 執行此程式計算未來14天的庫存預估
"""
import sys
from datetime import datetime, timedelta
from typing import List, Dict, Set, Any, Tuple, Union, Optional
from collections import defaultdict
from google_sheets_helper import GoogleSheetsHelper
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 前置天數（投料到完工入庫）
LEAD_TIME_DAYS = 2
# 計算天數
FORECAST_DAYS = 21

# 輸出工作表標題
# 欄位順序：期初庫存 → 當天組裝需求 → 預估入庫數量 → 期末庫存
INVENTORY_DETAIL_HEADERS = ['日期', '餅乾代號', '餅乾品名', '期初庫存', '當天組裝需求', '預估入庫數量', '期末庫存', '是否負庫存', '缺口數量']
def parse_date(date_str: Any) -> Optional[datetime]:
    """解析日期字串（Google Sheets 格式：YYYY/M/D 或 YYYY/MM/DD）
    支援格式：YYYY/M/D（單數月份和日期，例如：2025/1/5）、YYYY/MM/DD（雙數月份和日期，例如：2025/01/05）
    Args:date_str: 日期字串
    Returns: datetime 物件（時間設為00:00:00，只保留日期部分），無法解析則返回 None"""
    if isinstance(date_str, datetime):
        return normalize_date(date_str)
    if not date_str:
        return None
    date_str = str(date_str).strip()
    try:
        parts = date_str.split('/')
        if len(parts) == 3:
            year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            return normalize_date(datetime(year, month, day))
    except (ValueError, IndexError):
        pass    
    logger.warning(f"無法解析日期（期望格式：YYYY/M/D 或 YYYY/MM/DD）: {date_str}")
    return None

def normalize_date(date: datetime) -> datetime:
    """標準化日期（只取日期部分，時間設為00:00:00）"""
    return datetime(date.year, date.month, date.day)

def get_today_date() -> datetime:
    """取得今天的日期（只取日期部分，時間設為00:00:00）"""
    return normalize_date(datetime.now())

def get_header_index(headers: List[str], header_name: str, default: int = 0) -> int:
    """取得標題欄位索引
    Args: headers: 標題行列表, header_name: 欄位名稱, default: 預設索引（如果找不到欄位）
    Returns: 欄位索引"""
    return headers.index(header_name) if header_name in headers else default

def format_date(date: Optional[datetime]) -> str:
    """格式化日期為 YYYY/MM/DD 字串
    Args: date: datetime 物件或 None
    Returns: 日期字串（格式：YYYY/MM/DD），如果為 None 則返回空字串"""
    return date.strftime('%Y/%m/%d') if date else ''

def _parse_numeric_string(value_str: str) -> str:
    """移除千分位逗號並清理字串"""
    return value_str.replace(',', '').strip()

def parse_number(value: Any) -> int:
    """將文字轉換為整數，處理千分位逗號格式    
    例如："1,000" -> 1000, "-500" -> -500    
    Args:value: 要轉換的值（可能是文字、數字或 None）    
    Returns:轉換後的整數，如果無法轉換則返回 0 """
    if value is None:
        return 0
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)    
    value_str = str(value).strip()
    if not value_str:
        return 0    
    try:
        return int(float(_parse_numeric_string(value_str)))
    except (ValueError, TypeError):
        logger.warning(f"無法將 '{value}' 轉換為整數，使用 0")
        return 0

def parse_float(value: Any) -> float:
    """將文字轉換為浮點數，處理千分位逗號格式    
    例如："1,000.5" -> 1000.5, "-500.25" -> -500.25, "1,234.56" -> 1234.56    
    Args:value: 要轉換的值（可能是文字、數字或 None）    
    Returns:轉換後的浮點數，如果無法轉換則返回 0.0"""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)    
    value_str = str(value).strip()
    if not value_str:
        return 0.0    
    try:
        return float(_parse_numeric_string(value_str))
    except (ValueError, TypeError):
        logger.warning(f"無法將 '{value}' 轉換為浮點數，使用 0.0")
        return 0.0

def read_initial_inventory(sheets_helper: GoogleSheetsHelper) -> Dict[str, int]:
    """讀取今天的期初庫存（從Google Sheets讀取「庫存狀態」工作表）
    說明：
    - 此函數從Google Sheets讀取「庫存狀態」工作表
    - 此工作表的資料就是今天的期初庫存數量（例如：1/5的期初庫存）
    - 此工作表的資料應該已經過手動調整（可能先從ERP同步，再手動修改）
    - 多庫別（SP40, SP50, SP60, SP80）會按餅乾代號合併加總
    - 會處理千分位逗號格式的數字（例如："1,000"）
    - 包括負庫存數量也會正確處理和加總
    Args: sheets_helper: Google Sheets 輔助物件    
    Returns: 字典：{餅乾代號: 庫存數量（整數）} - 今天的期初庫存"""
    logger.info("從Google Sheets讀取今天的期初庫存（從「庫存狀態」工作表）...")
    inventory = defaultdict(int)    
    # 讀取庫存狀態工作表（多庫別需要按餅乾代號合併加總）
    try:
        inventory_data = sheets_helper.read_worksheet('庫存狀態')
        if len(inventory_data) > 1:
            headers = inventory_data[0]
            # 找到欄位索引（新格式：餅乾代號、餅乾品名、目前庫存數量、庫別代號、單位、最後更新日期）
            code_idx = get_header_index(headers, '餅乾代號', 0)
            qty_idx = get_header_index(headers, '目前庫存數量', 2)
            warehouse_idx = get_header_index(headers, '庫別代號', 3)
            
            processed_count = 0
            skipped_count = 0
            for row in inventory_data[1:]:
                if len(row) > max(code_idx, qty_idx):
                    cookie_code = str(row[code_idx]).strip() if row[code_idx] else ''
                    if not cookie_code:
                        skipped_count += 1
                        continue
                    qty = parse_number(row[qty_idx])
                    warehouse_code = str(row[warehouse_idx]).strip() if len(row) > warehouse_idx and row[warehouse_idx] else ''
                    inventory[cookie_code] += qty
                    processed_count += 1
                    
                    # 記錄詳細資訊（僅在 debug 模式下）
                    if qty != 0:
                        logger.debug(f"  餅乾代號: {cookie_code}, 庫別代號: {warehouse_code}, 庫存數量: {qty}")
            
            logger.info(f"處理了 {processed_count} 筆庫存記錄，跳過 {skipped_count} 筆（無餅乾代號）")
            logger.info(f"從「庫存狀態」工作表讀取到 {len(inventory)} 種餅乾的庫存（已按餅乾代號合併加總）")
    except Exception as e:
        logger.warning(f"讀取「庫存狀態」工作表失敗: {e}")
        import traceback
        traceback.print_exc()
    
    # 保留所有餅乾代號（包括數量為0的，因為可能後續有生產或需求）
    result = dict(inventory)
    if result:
        total_qty = sum(result.values())
        positive_count = sum(1 for v in result.values() if v > 0)
        negative_count = sum(1 for v in result.values() if v < 0)
        zero_count = sum(1 for v in result.values() if v == 0)
        logger.info(f"今天的期初庫存總計：{positive_count} 種餅乾有正庫存，{negative_count} 種餅乾有負庫存，{zero_count} 種餅乾為零庫存，總數量：{total_qty} 片")
    return result

def read_bom(sheets_helper: GoogleSheetsHelper) -> Dict[str, Dict[str, float]]:
    """讀取BOM表（禮盒組成表）
    Args: sheets_helper: Google Sheets 輔助物件
    Returns: 字典：{禮盒代號: {餅乾代號: 每盒片數, ...}}"""
    logger.info("讀取BOM表...")
    bom = defaultdict(lambda: defaultdict(float))
    try:
        bom_data = sheets_helper.read_worksheet('BOM')
        if len(bom_data) > 1:
            headers = bom_data[0]
            box_code_idx = get_header_index(headers, '禮盒代號', 0)
            cookie_code_idx = get_header_index(headers, '餅乾代號', 1)
            qty_idx = get_header_index(headers, '每盒片數', 2)            
            for row in bom_data[1:]:
                if len(row) > max(box_code_idx, cookie_code_idx, qty_idx):
                    box_code = str(row[box_code_idx]).strip()
                    cookie_code = str(row[cookie_code_idx]).strip()
                    try:
                        qty = parse_float(row[qty_idx])
                        if box_code and cookie_code and qty > 0:
                            bom[box_code][cookie_code] = qty
                    except (ValueError, TypeError):
                        continue            
            logger.info(f"讀取到 {len(bom)} 種禮盒的BOM資料")
    except Exception as e:
        logger.error(f"讀取BOM表失敗: {e}")
        raise    
    return dict(bom)

def read_production_schedule(sheets_helper: GoogleSheetsHelper, today: datetime) -> Dict[datetime, Dict[str, float]]:
    """讀取生產排程（包含今天及前3天的投料）
    
    說明：
    - 直接讀取「生產片數」欄位（不需要重新計算，因為應該已經由 sync_production_schedule.py 計算完成）
    - 讀取「日期」、「餅乾代號」、「生產片數」、「預計完成日期」欄位
    
    計算邏輯：
    - 如果「預計完成日期」欄位有指定日期，則使用該日期作為完工入庫日期
    - 如果「預計完成日期」欄位為空白，則使用預設值：投料日期 + 2天（前置天數）
    - 讀取所有投料日期 >= (今天 - 3天) 的記錄
    - 例如：今天是1/5，則要計算1/2投料、1/3投料、1/4投料和1/5投料分別加入到對應的完工入庫日期
    
    Args:
        sheets_helper: Google Sheets 輔助物件
        today: 今天的日期
    
    Returns:
        字典：{完工入庫日期: {餅乾代號: 生產數量（片）, ...}}
    """
    logger.info("讀取生產排程（包含今天及前3天的投料）...")
    production = defaultdict(lambda: defaultdict(float))
    min_production_date = today.date() - timedelta(days=3)
    
    try:
        schedule_data = sheets_helper.read_worksheet('生產排程')
        if len(schedule_data) > 1:
            headers = schedule_data[0]
            # 讀取必要的欄位：日期、餅乾代號、生產片數、預計完成日期
            date_idx = get_header_index(headers, '日期', 0)
            cookie_code_idx = get_header_index(headers, '餅乾代號', 2)
            pieces_qty_idx = get_header_index(headers, '生產片數', 5)
            completion_date_idx = get_header_index(headers, '預計完成日期', -1)
            
            skipped_before_min_date = 0
            skipped_no_pieces = 0
            used_custom_date_count = 0
            used_default_date_count = 0
            
            for row in schedule_data[1:]:
                if len(row) > max(date_idx, cookie_code_idx, pieces_qty_idx):
                    # 解析投料日期
                    production_date = parse_date(row[date_idx])
                    if not production_date:
                        continue
                    # 只讀取投料日期 >= (今天 - 3天) 的記錄
                    if production_date.date() < min_production_date:
                        skipped_before_min_date += 1
                        continue
                    
                    cookie_code = str(row[cookie_code_idx]).strip() if row[cookie_code_idx] else ''
                    if not cookie_code:
                        continue
                    
                    try:
                        qty_pieces = parse_float(row[pieces_qty_idx])
                        if qty_pieces > 0:
                            # 判斷是否使用指定的預計完成日期
                            if completion_date_idx >= 0 and completion_date_idx < len(row) and row[completion_date_idx]:
                                # 如果有指定預計完成日期，使用該日期
                                custom_completion_date = parse_date(row[completion_date_idx])
                                if custom_completion_date:
                                    completion_date = normalize_date(custom_completion_date)
                                    used_custom_date_count += 1
                                else:
                                    # 如果解析失敗，使用預設值
                                    completion_date = normalize_date(production_date + timedelta(days=LEAD_TIME_DAYS))
                                    used_default_date_count += 1
                            else:
                                # 如果預計完成日期為空白，使用預設值：投料日期 + 2天
                                completion_date = normalize_date(production_date + timedelta(days=LEAD_TIME_DAYS))
                                used_default_date_count += 1
                            
                            production[completion_date][cookie_code] += qty_pieces
                        else:
                            skipped_no_pieces += 1
                    except (ValueError, TypeError) as e:
                        logger.warning(f"解析生產排程資料失敗（餅乾代號：{cookie_code}）: {e}")
                        continue
            
            if skipped_before_min_date > 0:
                logger.info(f"已排除 {min_production_date} 之前的投料記錄 {skipped_before_min_date} 筆")
            if skipped_no_pieces > 0:
                logger.debug(f"跳過 {skipped_no_pieces} 筆生產片數為0或空的記錄")
            if used_custom_date_count > 0:
                logger.info(f"使用指定預計完成日期：{used_custom_date_count} 筆")
            if used_default_date_count > 0:
                logger.info(f"使用預設完工日期（投料日期+2天）：{used_default_date_count} 筆")
            logger.info(f"讀取到 {len(production)} 天的生產排程（已轉換為完工入庫日期）")
    except Exception as e:
        logger.error(f"讀取生產排程失敗: {e}")
        raise
    return dict(production)


def read_assembly_schedule(sheets_helper: GoogleSheetsHelper, bom: Dict[str, Dict[str, float]]) -> Dict[datetime, Dict[str, float]]:
    """讀取組裝排程並展開為餅乾需求量    
    Args:sheets_helper: Google Sheets 輔助物件,bom: BOM表字典
    Returns:字典：{組裝日期: {餅乾代號: 需求量, ...}}"""
    logger.info("讀取組裝排程...")
    assembly = defaultdict(lambda: defaultdict(float))    
    try:
        assembly_data = sheets_helper.read_worksheet('組裝計劃')
        if len(assembly_data) > 1:
            headers = assembly_data[0]
            date_idx = get_header_index(headers, '日期', 0)
            box_code_idx = get_header_index(headers, '禮盒代號', 1)
            qty_idx = get_header_index(headers, '計畫組裝數量', 2)
            for row in assembly_data[1:]:
                if len(row) > max(date_idx, box_code_idx, qty_idx):
                    # 解析組裝日期
                    assembly_date = parse_date(row[date_idx])
                    if not assembly_date:
                        continue                    
                    box_code = str(row[box_code_idx]).strip()
                    try:
                        box_qty = parse_float(row[qty_idx])
                        if box_code and box_qty > 0:
                            assembly_date_key = normalize_date(assembly_date)
                            # 使用BOM表展開為餅乾需求量
                            if box_code in bom:
                                for cookie_code, pieces_per_box in bom[box_code].items():
                                    cookie_qty = box_qty * pieces_per_box
                                    assembly[assembly_date_key][cookie_code] += cookie_qty
                            else:
                                logger.warning(f"禮盒 {box_code} 在BOM表中找不到")
                    except (ValueError, TypeError):
                        continue            
            logger.info(f"讀取到 {len(assembly)} 天的組裝排程（已展開為餅乾需求）")
    except Exception as e:
        logger.error(f"讀取組裝排程失敗: {e}")
        raise    
    return dict(assembly)

def get_all_cookie_codes(
    initial_inventory: Dict[str, float],
    production_schedule: Dict[datetime, Dict[str, float]],
    assembly_schedule: Dict[datetime, Dict[str, float]]
) -> Set[str]:
    """取得所有需要計算的餅乾代號集合    
    Args: initial_inventory: 期初庫存, production_schedule: 生產排程（完工入庫日期）, assembly_schedule: 組裝排程（餅乾需求量）
    Returns: 所有餅乾代號的集合"""
    all_cookies = set(initial_inventory.keys())
    for schedule in production_schedule.values():
        all_cookies.update(schedule.keys())
    for schedule in assembly_schedule.values():
        all_cookies.update(schedule.keys())
    return all_cookies

def calculate_daily_inventory(
    date: datetime,
    cookie_code: str,
    current_inventory: Dict[str, float],
    production_schedule: Dict[datetime, Dict[str, float]],
    assembly_schedule: Dict[datetime, Dict[str, float]]
) -> Tuple[float, float, float, float]:
    """計算單一餅乾在特定日期的庫存變化
    
    計算邏輯：
    - 期初庫存 = 前一天的期末庫存（第一天使用從「庫存狀態」工作表讀取的期初庫存）
    - 當天組裝需求量 = 從組裝排程取得的當天組裝計劃所需的餅乾數量
    - 當天完工入庫數量 = 從生產排程取得的當天預計要完工入庫的餅乾數量（生產排程日期 + 2天 = 完工入庫日期）
    - 期末庫存 = 期初庫存 - 當天組裝計劃所需的餅乾 + 當天預計要完工入庫的餅乾
    - 當天的期末庫存會轉為明天的期初庫存
    
    Args:
        date: 計算日期
        cookie_code: 餅乾代號
        current_inventory: 當前庫存狀態（會更新，作為下一天的期初庫存）
        production_schedule: 生產排程（完工入庫日期: {餅乾代號: 生產數量}）
        assembly_schedule: 組裝排程（組裝日期: {餅乾代號: 需求量}）
    
    Returns:
        (期初庫存, 預估入庫數量, 當天組裝需求, 期末庫存)
    """
    date_key = normalize_date(date)
    beginning_qty = current_inventory.get(cookie_code, 0.0)
    demand_qty = assembly_schedule.get(date_key, {}).get(cookie_code, 0.0)
    completion_qty = production_schedule.get(date_key, {}).get(cookie_code, 0.0)
    ending_qty = beginning_qty - demand_qty + completion_qty
    current_inventory[cookie_code] = ending_qty
    return beginning_qty, completion_qty, demand_qty, ending_qty

def create_detail_row(
    date: datetime,
    cookie_code: str,
    cookie_name: str,
    beginning_qty: float,
    completion_qty: float,
    demand_qty: float,
    ending_qty: float
) -> List[Any]:
    """建立庫存明細記錄
    
    Args:
        date: 日期
        cookie_code: 餅乾代號
        cookie_name: 餅乾品名
        beginning_qty: 期初庫存
        completion_qty: 預估入庫數量
        demand_qty: 當天組裝需求
        ending_qty: 期末庫存
    
    Returns:
        明細記錄列表
    """
    shortage_qty = abs(ending_qty) if ending_qty < 0 else 0.0
    return [
        format_date(date),
        cookie_code,
        cookie_name,
        beginning_qty,
        demand_qty,
        completion_qty,
        ending_qty,
        '是' if ending_qty < 0 else '否',
        shortage_qty
    ]

def calculate_inventory_forecast(
    initial_inventory: Dict[str, Union[int, float]],
    production_schedule: Dict[datetime, Dict[str, float]],
    assembly_schedule: Dict[datetime, Dict[str, float]],
    today: datetime,
    cookie_names: Dict[str, str]
) -> List[List[Any]]:
    """計算未來14天的庫存預估    
    計算邏輯：
    - 從「庫存狀態」工作表讀取今天的期初庫存
    - 逐日計算每種餅乾的庫存變化：
      * 期初庫存 = 前一天的期末庫存（第一天使用從「庫存狀態」工作表讀取的期初庫存）
      * 當天組裝需求量 = 從組裝排程取得的當天組裝計劃所需的餅乾數量
      * 當天完工入庫數量 = 從生產排程取得的當天預計要完工入庫的餅乾數量（生產排程日期 + 2天 = 完工入庫日期）
      * 期末庫存 = 期初庫存 - 當天組裝計劃所需的餅乾 + 當天預計要完工入庫的餅乾
      * 當天的期末庫存會轉為明天的期初庫存（迭代計算）
    - 在明細記錄中包含「是否負庫存」和「缺口數量」欄位
    Args: initial_inventory: 期初庫存（從「庫存狀態」工作表讀取的今天的期初庫存）, production_schedule: 生產排程（完工入庫日期: {餅乾代號: 生產數量}，生產排程日期 + 2天 = 完工入庫日期）, assembly_schedule: 組裝排程（組裝日期: {餅乾代號: 需求量}）, today: 今天的日期, cookie_names: 餅乾名稱對應表
    Returns: 庫存明細列表"""
    logger.info(f"開始計算未來 {FORECAST_DAYS} 天的庫存預估（前置天數：{LEAD_TIME_DAYS} 天）...")
    
    detail_rows = []
    
    # 取得所有需要計算的餅乾代號
    all_cookies = get_all_cookie_codes(initial_inventory, production_schedule, assembly_schedule)
    logger.info(f"需要計算的餅乾種類：{len(all_cookies)} 種")
    
    current_inventory = {k: float(v) for k, v in initial_inventory.items()}
    
    for day_offset in range(FORECAST_DAYS):
        date = today + timedelta(days=day_offset)
        for cookie_code in sorted(all_cookies):
            beginning_qty, completion_qty, demand_qty, ending_qty = calculate_daily_inventory(
                date, cookie_code, current_inventory, production_schedule, assembly_schedule
            )
            
            cookie_name = cookie_names.get(cookie_code, '')
            detail_rows.append(create_detail_row(
                date, cookie_code, cookie_name,
                beginning_qty, completion_qty, demand_qty, ending_qty
            ))
    
    logger.info(f"計算完成：共 {len(detail_rows)} 筆明細記錄")
    return detail_rows


def write_worksheet_data(
    sheets_helper: GoogleSheetsHelper,
    worksheet_name: str,
    headers: List[str],
    rows: List[List[Any]]
) -> None:
    """將資料寫入指定的工作表
    
    Args:
        sheets_helper: Google Sheets 輔助物件
        worksheet_name: 工作表名稱
        headers: 標題行
        rows: 資料行列表
    """
    worksheet = sheets_helper.get_worksheet(worksheet_name, create_if_not_exists=True)
    if worksheet is None:
        raise ValueError(f"無法取得或建立「{worksheet_name}」工作表")
    
    all_data = [headers] + rows
    worksheet.clear()
    if len(all_data) > 0:
        num_cols = len(headers)
        end_col = chr(ord('A') + num_cols - 1)
        range_name = f'A1:{end_col}{len(all_data)}'
        worksheet.update(range_name=range_name, values=all_data)

def write_results(
    sheets_helper: GoogleSheetsHelper,
    detail_rows: List[List[Any]]
):
    """將計算結果寫入Google Sheets
    Args: sheets_helper: Google Sheets 輔助物件, detail_rows: 庫存明細列表"""
    logger.info("寫入計算結果到Google Sheets...")
    
    try:
        write_worksheet_data(sheets_helper, '庫存預估明細', INVENTORY_DETAIL_HEADERS, detail_rows)
        logger.info(f"已寫入 {len(detail_rows)} 筆資料到「庫存預估明細」工作表")
    except Exception as e:
        logger.error(f"寫入「庫存預估明細」工作表失敗: {e}")
        raise


def calculate_cookie_inventory():
    """主函數：執行餅乾庫存算料計算"""
    logger.info("=" * 60)
    logger.info("開始執行餅乾庫存算料計算")
    logger.info("=" * 60)
    
    try:
        # 連接 Google Sheets
        sheets_helper = GoogleSheetsHelper()
        logger.info("已連接到 Google Sheets")      
        # 取得今天的日期
        today = get_today_date()
        end_date = today + timedelta(days=FORECAST_DAYS-1)
        logger.info(f"計算基準日期：{format_date(today)}（今天）")
        logger.info(f"計算範圍：未來 {FORECAST_DAYS} 天（從 {format_date(today)} 到 {format_date(end_date)}）")
        
        # 1. 讀取今天的期初庫存（從Google Sheets讀取「庫存狀態」工作表）
        # 注意：此工作表的資料應該已經過手動調整（可能先從ERP同步，再手動修改）
        initial_inventory = read_initial_inventory(sheets_helper)
        
        # 取得餅乾名稱對應表（用於輸出餅乾品名）
        index_dict = sheets_helper.get_index_dict()
        cookie_names = index_dict.get('餅乾', {})
        
        # 2. 讀取BOM表
        bom = read_bom(sheets_helper)
        
        # 3. 讀取生產排程（從今天開始之後的投料，包含今天）
        production_schedule = read_production_schedule(sheets_helper, today)
        
        # 4. 讀取組裝排程並展開為餅乾需求
        assembly_schedule = read_assembly_schedule(sheets_helper, bom)
        
        # 5. 計算未來14天的庫存預估
        detail_rows = calculate_inventory_forecast(
            {k: float(v) for k, v in initial_inventory.items()},
            production_schedule,
            assembly_schedule,
            today,
            cookie_names
        )
        
        # 6. 輸出結果
        write_results(sheets_helper, detail_rows)
        
        # 統計負庫存數量（用於日誌顯示）
        shortage_count = sum(1 for row in detail_rows if row[7] == '是')  # 第8欄（索引7）為「是否負庫存」
        
        logger.info("=" * 60)
        logger.info("計算完成！")
        logger.info(f"負庫存警示：{shortage_count} 筆（已包含在「庫存預估明細」工作表中）")
        logger.info("=" * 60)
        return True
        
    except Exception as e:
        logger.error(f"計算失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    """
    餅乾庫存算料系統    
    功能：
    - 從Google Sheets讀取今天的期初庫存（從「庫存狀態」工作表）
    - 計算未來14天每一天每種餅乾的庫存數量
    - 檢測負庫存（餅乾不足）的情況（包含在「庫存預估明細」工作表的「是否負庫存」和「缺口數量」欄位）
    - 輸出結果到「庫存預估明細」工作表    
    執行前準備：
    1. 執行 sync_production_schedule.py 計算並更新生產排程的「生產片數」
    2. 執行 sync_inventory_from_erp.py 從ERP同步「庫存狀態」（可選）
    3. 手動調整 Google Sheets 中的「庫存狀態」工作表的資料
    4. 執行此程式進行計算    
    計算邏輯：
    - 期初庫存 = 從Google Sheets讀取的「庫存狀態」（已手動調整）
    - 生產排程：直接讀取「生產片數」欄位（不需要重新計算）
    - 只讀取從今天開始之後的投料記錄（包含今天）
    - 前置天數固定為2天（投料日期+2天=完工入庫日期）
    """
    import sys
    # 正常執行模式
    success = calculate_cookie_inventory()
    sys.exit(0 if success else 1)
