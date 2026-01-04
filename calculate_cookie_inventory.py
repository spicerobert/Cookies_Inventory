"""餅乾庫存算料系統
功能說明：
- 從Google Sheets讀取今天的期初庫存（合併「庫存狀態」和「在製品庫存」工作表）
  （注意：這兩個工作表的資料應該已經過手動調整）
- 讀取BOM表、生產排程、組裝排程
- 計算未來14天每一天每種餅乾的庫存數量
- 檢測負庫存（餅乾不足）的情況
- 輸出結果到Google Sheets

執行流程：
1. 先執行 sync_inventory_from_erp.py 和 sync_wip_from_erp.py 從ERP同步資料
2. 手動調整 Google Sheets 中的「在製品庫存」和「庫存狀態」工作表
3. 執行此程式計算未來14天的庫存預估
"""
import sys
from datetime import datetime, timedelta
from typing import List, Dict, Set, Any, Tuple
from collections import defaultdict
from google_sheets_helper import GoogleSheetsHelper
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 前置天數（投料到完工）
LEAD_TIME_DAYS = 5

# 計算天數
FORECAST_DAYS = 14

# 輸出工作表標題
# 欄位順序：期初庫存 → 當天組裝需求 → 當天入庫數量 → 期末庫存
INVENTORY_DETAIL_HEADERS = ['日期', '餅乾代號', '餅乾品名', '期初庫存', '當天組裝需求', '當天入庫數量', '期末庫存', '是否負庫存', '缺口數量']
SHORTAGE_ALERT_HEADERS = ['日期', '餅乾代號', '餅乾品名', '缺口數量', '當天組裝需求', '期初庫存', '當天入庫數量']

def parse_date(date_str: Any) -> datetime:
    """解析日期字串（Google Sheets標準格式：YYYY/MM/DD）    
    Args: date_str: 日期字串（格式：YYYY/MM/DD）    
    Returns: datetime 物件（時間設為00:00:00，只保留日期部分）"""
    if isinstance(date_str, datetime):
        # 如果已經是datetime，只保留日期部分
        return datetime(date_str.year, date_str.month, date_str.day)    
    if not date_str:
        return None
    
    date_str = str(date_str).strip()    
    # 只處理 Google Sheets 標準格式：YYYY/MM/DD
    try:
        dt = datetime.strptime(date_str, '%Y/%m/%d')
        return datetime(dt.year, dt.month, dt.day)  # 確保時間為00:00:00
    except ValueError:
        logger.warning(f"無法解析日期（期望格式：YYYY/MM/DD）: {date_str}")
        return None

def get_today_date() -> datetime:
    """取得今天的日期（只取日期部分，時間設為00:00:00）"""
    today = datetime.now()
    return datetime(today.year, today.month, today.day)

def format_date(date: datetime) -> str:
    """格式化日期為 YYYY/MM/DD 字串
    Args: date: datetime 物件    
    Returns: 日期字串（格式：YYYY/MM/DD）"""
    return date.strftime('%Y/%m/%d')

def read_initial_inventory(sheets_helper: GoogleSheetsHelper) -> Dict[str, float]:
    """讀取今天的期初庫存（從Google Sheets讀取並合併「庫存狀態」和「在製品庫存」工作表）
    說明：
    - 此函數從Google Sheets讀取「庫存狀態」和「在製品庫存」工作表
    - 這兩個工作表的資料應該已經過手動調整（可能先從ERP同步，再手動修改）
    - 合併後的庫存數量 = 昨天結束後的期末庫存 = 今天早上的期初庫存
    - 多庫別（SP40, SP50, SP60）會按餅乾代號合併加總
    - 在製品庫存也會按餅乾代號加總後合併到庫存中
    Args: sheets_helper: Google Sheets 輔助物件    
    Returns: 字典：{餅乾代號: 庫存數量} - 今天的期初庫存"""
    logger.info("從Google Sheets讀取今天的期初庫存（合併「庫存狀態」和「在製品庫存」工作表）...")
    inventory = defaultdict(float)    
    # 讀取庫存狀態工作表（多庫別需要按餅乾代號合併加總）
    try:
        inventory_data = sheets_helper.read_worksheet('庫存狀態')
        if len(inventory_data) > 1:
            headers = inventory_data[0]
            # 找到欄位索引（新格式：餅乾代號、餅乾品名、目前庫存數量、庫別代號、單位、最後更新日期）
            code_idx = headers.index('餅乾代號') if '餅乾代號' in headers else 0
            qty_idx = headers.index('目前庫存數量') if '目前庫存數量' in headers else 2  # 調整索引：現在是第3欄（索引2）            
            for row in inventory_data[1:]:
                if len(row) > max(code_idx, qty_idx):
                    cookie_code = str(row[code_idx]).strip()
                    try:
                        qty = float(row[qty_idx]) if row[qty_idx] else 0.0
                        if cookie_code:
                            inventory[cookie_code] += qty
                    except (ValueError, TypeError):
                        continue            
            logger.info(f"從「庫存狀態」工作表讀取到 {len([k for k, v in inventory.items() if v > 0])} 種餅乾的庫存")
    except Exception as e:
        logger.warning(f"讀取「庫存狀態」工作表失敗: {e}")
    
    # 讀取在製品庫存工作表（需要按餅乾代號加總後合併）
    try:
        wip_data = sheets_helper.read_worksheet('在製品庫存')
        if len(wip_data) > 1:
            headers = wip_data[0]
            # 找到欄位索引（新格式：餅乾代號、餅乾品名、製令單別、製令單號、在製品數量、單位、最後更新日期）
            code_idx = headers.index('餅乾代號') if '餅乾代號' in headers else 0
            qty_idx = headers.index('在製品數量') if '在製品數量' in headers else 4  # 調整索引：現在是第5欄（索引4）            
            for row in wip_data[1:]:
                if len(row) > max(code_idx, qty_idx):
                    cookie_code = str(row[code_idx]).strip()
                    try:
                        qty = float(row[qty_idx]) if row[qty_idx] else 0.0
                        if cookie_code:
                            inventory[cookie_code] += qty
                    except (ValueError, TypeError):
                        continue            
            logger.info(f"從「在製品庫存」工作表讀取資料（已合併到庫存中）")
    except Exception as e:
        logger.warning(f"讀取「在製品庫存」工作表失敗: {e}")
    
    # 保留所有餅乾代號（包括數量為0的，因為可能後續有生產或需求）
    result = dict(inventory)
    logger.info(f"今天的期初庫存總計：{len([k for k, v in result.items() if v > 0])} 種餅乾有庫存，總數量：{sum(result.values()):.0f} 片")
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
            box_code_idx = headers.index('禮盒代號') if '禮盒代號' in headers else 0
            cookie_code_idx = headers.index('餅乾代號') if '餅乾代號' in headers else 1
            qty_idx = headers.index('每盒片數') if '每盒片數' in headers else 2            
            for row in bom_data[1:]:
                if len(row) > max(box_code_idx, cookie_code_idx, qty_idx):
                    box_code = str(row[box_code_idx]).strip()
                    cookie_code = str(row[cookie_code_idx]).strip()
                    try:
                        qty = float(row[qty_idx]) if row[qty_idx] else 0.0
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
    """讀取生產排程（排除今天的投料）    
    Args:sheets_helper: Google Sheets 輔助物件,today: 今天的日期
    Returns:字典：{完工入庫日期: {餅乾代號: 生產數量, ...}}"""
    logger.info("讀取生產排程...")
    production = defaultdict(lambda: defaultdict(float))    
    try:
        schedule_data = sheets_helper.read_worksheet('生產排程建議')
        if len(schedule_data) > 1:
            headers = schedule_data[0]
            date_idx = headers.index('日期') if '日期' in headers else 0
            cookie_code_idx = headers.index('餅乾代號') if '餅乾代號' in headers else 2
            qty_idx = headers.index('建議生產數量_片') if '建議生產數量_片' in headers else 3
            skipped_today = 0
            for row in schedule_data[1:]:
                if len(row) > max(date_idx, cookie_code_idx, qty_idx):
                    # 解析投料日期
                    production_date = parse_date(row[date_idx])
                    if not production_date:
                        continue
                    # 排除今天的投料（因為已經包含在在製品庫存中）
                    if production_date.date() == today.date():
                        skipped_today += 1
                        continue                    
                    cookie_code = str(row[cookie_code_idx]).strip()
                    try:
                        qty = float(row[qty_idx]) if row[qty_idx] else 0.0
                        if cookie_code and qty > 0:
                            # 計算完工入庫日期 = 投料日期 + 5天
                            completion_date = production_date + timedelta(days=LEAD_TIME_DAYS)
                            # 標準化日期（只取日期部分，時間設為00:00:00）
                            completion_date_key = datetime(completion_date.year, completion_date.month, completion_date.day)
                            production[completion_date_key][cookie_code] += qty
                    except (ValueError, TypeError):
                        continue
            if skipped_today > 0:
                logger.info(f"已排除今天的投料記錄 {skipped_today} 筆")
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
            date_idx = headers.index('日期') if '日期' in headers else 0
            box_code_idx = headers.index('禮盒代號') if '禮盒代號' in headers else 1
            qty_idx = headers.index('計畫組裝數量') if '計畫組裝數量' in headers else 2
            for row in assembly_data[1:]:
                if len(row) > max(date_idx, box_code_idx, qty_idx):
                    # 解析組裝日期
                    assembly_date = parse_date(row[date_idx])
                    if not assembly_date:
                        continue                    
                    box_code = str(row[box_code_idx]).strip()
                    try:
                        box_qty = float(row[qty_idx]) if row[qty_idx] else 0.0
                        if box_code and box_qty > 0:
                            # 標準化日期（只取日期部分，時間設為00:00:00）
                            assembly_date_key = datetime(assembly_date.year, assembly_date.month, assembly_date.day)
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

def calculate_inventory_forecast(
    initial_inventory: Dict[str, float],
    production_schedule: Dict[datetime, Dict[str, float]],
    assembly_schedule: Dict[datetime, Dict[str, float]],
    today: datetime,
    cookie_names: Dict[str, str]
) -> Tuple[List[List[Any]], List[List[Any]]]:
    """計算未來14天的庫存預估    
    Args:initial_inventory: 期初庫存,production_schedule: 生產排程（完工入庫日期）,assembly_schedule: 組裝排程（餅乾需求量）,today: 今天的日期,cookie_names: 餅乾名稱對應表
    Returns:(庫存明細列表, 負庫存警示列表)"""
    logger.info("開始計算未來14天的庫存預估...")    
    detail_rows = []
    shortage_rows = []    
    # 取得所有需要計算的餅乾代號
    all_cookies = set(initial_inventory.keys())
    for schedule in production_schedule.values():
        all_cookies.update(schedule.keys())
    for schedule in assembly_schedule.values():
        all_cookies.update(schedule.keys())    
    # 每天的庫存狀態：{日期: {餅乾代號: 期末庫存}}
    daily_inventory = {}
    current_inventory = initial_inventory.copy()    
    # 逐日計算
    for day_offset in range(FORECAST_DAYS):
        date = today + timedelta(days=day_offset)
        date_key = datetime(date.year, date.month, date.day)        
        daily_inventory[date_key] = {}        
        # 對每一種餅乾計算
        for cookie_code in sorted(all_cookies):
            # 期初庫存 = 前一天的期末庫存
            beginning_qty = current_inventory.get(cookie_code, 0.0)            
            # 當天預計完工入庫數量
            completion_qty = production_schedule.get(date_key, {}).get(cookie_code, 0.0)            
            # 當天組裝需求量
            demand_qty = assembly_schedule.get(date_key, {}).get(cookie_code, 0.0)            
            # 期末庫存 = 期初 + 完工入庫 - 組裝需求
            ending_qty = beginning_qty + completion_qty - demand_qty            
            # 更新當前庫存（作為下一天的期初庫存）
            current_inventory[cookie_code] = ending_qty
            daily_inventory[date_key][cookie_code] = ending_qty            
            # 判斷是否負庫存
            is_shortage = ending_qty < 0
            shortage_qty = abs(ending_qty) if is_shortage else 0.0            
            # 取得餅乾品名
            cookie_name = cookie_names.get(cookie_code, '')
            
            # 建立明細記錄
            # 欄位順序：日期、餅乾代號、餅乾品名、期初庫存、當天組裝需求、當天入庫數量、期末庫存、是否負庫存、缺口數量
            detail_rows.append([
                format_date(date),      # 日期（YYYY/MM/DD）
                cookie_code,
                cookie_name,            # 餅乾品名
                beginning_qty,          # 期初庫存
                demand_qty,             # 當天組裝需求
                completion_qty,         # 當天入庫數量
                ending_qty,             # 期末庫存
                '是' if is_shortage else '否',
                shortage_qty
            ])
            
            # 如果是負庫存，加入警示列表
            if is_shortage:
                shortage_rows.append([
                    format_date(date),  # 日期（YYYY/MM/DD）
                    cookie_code,
                    cookie_name,        # 餅乾品名
                    shortage_qty,
                    demand_qty,         # 當天組裝需求
                    beginning_qty,      # 期初庫存
                    completion_qty      # 當天入庫數量
                ])
    
    logger.info(f"計算完成：共 {len(detail_rows)} 筆明細記錄，{len(shortage_rows)} 筆負庫存警示")
    return detail_rows, shortage_rows


def write_results(
    sheets_helper: GoogleSheetsHelper,
    detail_rows: List[List[Any]],
    shortage_rows: List[List[Any]]
):
    """將計算結果寫入Google Sheets
    
    Args:
        sheets_helper: Google Sheets 輔助物件
        detail_rows: 庫存明細列表
        shortage_rows: 負庫存警示列表
    """
    logger.info("寫入計算結果到Google Sheets...")
    
    # 寫入庫存預估明細
    try:
        worksheet = sheets_helper.get_worksheet('庫存預估明細', create_if_not_exists=True)
        all_data = [INVENTORY_DETAIL_HEADERS] + detail_rows
        
        worksheet.clear()
        if len(all_data) > 0:
            num_cols = len(INVENTORY_DETAIL_HEADERS)
            end_col = chr(ord('A') + num_cols - 1)
            range_name = f'A1:{end_col}{len(all_data)}'
            worksheet.update(range_name=range_name, values=all_data)
        
        logger.info(f"已寫入 {len(detail_rows)} 筆資料到「庫存預估明細」工作表")
    except Exception as e:
        logger.error(f"寫入「庫存預估明細」工作表失敗: {e}")
        raise
    
    # 寫入負庫存警示
    try:
        worksheet = sheets_helper.get_worksheet('負庫存警示', create_if_not_exists=True)
        all_data = [SHORTAGE_ALERT_HEADERS] + shortage_rows
        
        worksheet.clear()
        if len(all_data) > 0:
            num_cols = len(SHORTAGE_ALERT_HEADERS)
            end_col = chr(ord('A') + num_cols - 1)
            range_name = f'A1:{end_col}{len(all_data)}'
            worksheet.update(range_name=range_name, values=all_data)
        
        logger.info(f"已寫入 {len(shortage_rows)} 筆資料到「負庫存警示」工作表")
    except Exception as e:
        logger.error(f"寫入「負庫存警示」工作表失敗: {e}")
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
        
        # 1. 讀取今天的期初庫存（從Google Sheets讀取「庫存狀態」和「在製品庫存」工作表並合併）
        # 注意：這兩個工作表的資料應該已經過手動調整（可能先從ERP同步，再手動修改）
        initial_inventory = read_initial_inventory(sheets_helper)
        
        # 取得餅乾名稱對應表（用於輸出餅乾品名）
        index_dict = sheets_helper.get_index_dict()
        cookie_names = index_dict.get('餅乾', {})
        
        # 2. 讀取BOM表
        bom = read_bom(sheets_helper)
        
        # 3. 讀取生產排程（排除今天的投料）
        production_schedule = read_production_schedule(sheets_helper, today)
        
        # 4. 讀取組裝排程並展開為餅乾需求
        assembly_schedule = read_assembly_schedule(sheets_helper, bom)
        
        # 5. 計算未來14天的庫存預估
        detail_rows, shortage_rows = calculate_inventory_forecast(
            initial_inventory,
            production_schedule,
            assembly_schedule,
            today,
            cookie_names
        )
        
        # 6. 輸出結果
        write_results(sheets_helper, detail_rows, shortage_rows)
        
        logger.info("=" * 60)
        logger.info("計算完成！")
        logger.info(f"負庫存警示：{len(shortage_rows)} 筆")
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
    - 從Google Sheets讀取今天的期初庫存（合併「庫存狀態」和「在製品庫存」）
    - 計算未來14天每一天每種餅乾的庫存數量
    - 檢測負庫存（餅乾不足）的情況
    - 輸出結果到「庫存預估明細」和「負庫存警示」工作表
    
    執行前準備：
    1. 執行 sync_inventory_from_erp.py 從ERP同步「庫存狀態」
    2. 執行 sync_wip_from_erp.py 從ERP同步「在製品庫存」
    3. 手動調整 Google Sheets 中的上述兩個工作表的資料
    4. 執行此程式進行計算
    
    計算邏輯：
    - 期初庫存 = 從Google Sheets讀取的「庫存狀態」+「在製品庫存」（已手動調整）
    - 排除今天的投料記錄（因為已包含在在製品庫存中）
    - 前置天數固定為5天（投料日期+5天=完工入庫日期）
    """
    success = calculate_cookie_inventory()
    sys.exit(0 if success else 1)
