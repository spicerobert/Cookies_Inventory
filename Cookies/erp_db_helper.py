"""
ERP 資料庫連接工具模組
專用於 MS-SQL Server 資料庫連接，用於查詢餅乾庫存
"""
import configparser
from typing import List, Dict, Any, Optional
import logging

try:
    import pyodbc
except ImportError:
    raise ImportError("請安裝 pyodbc: pip install pyodbc")

logger = logging.getLogger(__name__)

# 預設查詢（移出 config，避免在設定檔暴露 SQL）
DEFAULT_COOKIE_INVENTORY_QUERY = """
SELECT 
    LC.LC001 as cookie_code,
    LC.LC003 as warehouse_code,
    LC.LC004 + COALESCE(SUM(LA.LA011 * LA.LA005), 0) as qty,
    COALESCE(MB.MB004, '片') as unit,
    COALESCE(MB.MB002, '') as cookie_name
FROM [AS_online].[dbo].[INVLC] LC
LEFT JOIN [AS_online].[dbo].[INVLA] LA 
    ON LA.LA001 = LC.LC001 
    AND LA.LA009 = LC.LC003
    AND LA.LA004 >= '20251201'
LEFT JOIN [AS_online].[dbo].[INVMB] MB
    ON MB.MB001 = LC.LC001
WHERE LC.LC001 IS NOT NULL 
    AND LC.LC002 = '202512' 
    AND LC.LC003 IN ('SP40', 'SP50', 'SP60', 'SP80')
GROUP BY LC.LC001, LC.LC003, LC.LC004, MB.MB004, MB.MB002
"""

DEFAULT_WIP_INVENTORY_QUERY = """
SELECT 
    MO.TA001 as mo_number_type,
    MO.TA002 as mo_number,
    MO.TA006 as cookie_code,
    (MO.TA016 - MO.TA017 - ISNULL(MO.TA018, 0)) as wip_qty,
    COALESCE(MB.MB004, '片') as unit,
    COALESCE(MB.MB002, '') as cookie_name
FROM [AS_online].[dbo].[MOCTA] MO
LEFT JOIN [AS_online].[dbo].[INVMB] MB
    ON MB.MB001 = MO.TA006
WHERE MO.TA003 >= '20251101'
    AND MO.TA011 = '3'
    AND MO.TA006 IS NOT NULL
    AND (MO.TA016 - MO.TA017 - ISNULL(MO.TA018, 0)) > 0
ORDER BY MO.TA006, MO.TA001, MO.TA002
"""


class ERPDBHelper:
    """ERP 資料庫連接輔助類別（MS-SQL Server）"""
    
    def __init__(self, config_file='config.ini'):
        """
        初始化 ERP 資料庫連接
        
        Args:
            config_file: 設定檔路徑
        """
        self.config_file = config_file
        self.connection = None
        self._connect()
    
    def _connect(self):
        """建立 MS-SQL Server 資料庫連接"""
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        
        if 'ERP_DATABASE' not in config:
            raise ValueError("config.ini 中缺少 [ERP_DATABASE] 區段")
        
        db_config = config['ERP_DATABASE']
        server = db_config.get('server')
        database = db_config.get('database')
        username = db_config.get('username', '')
        password = db_config.get('password', '')
        
        if not server or not database:
            raise ValueError("config.ini 中缺少 server 或 database 設定")
        
        # 選擇可用的 ODBC 驅動程式
        driver = self._find_available_driver()
        
        # 建立連接字串
        if username and password:
            conn_str = (
                f"DRIVER={{{driver}}};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"UID={username};"
                f"PWD={password}"
            )
        else:
            conn_str = (
                f"DRIVER={{{driver}}};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"Trusted_Connection=yes;"
            )
        
        try:
            self.connection = pyodbc.connect(conn_str)
            logger.info(f"已連接到 SQL Server: {server}/{database}")
        except Exception as e:
            logger.error(f"連接 SQL Server 失敗: {str(e)}")
            raise
    
    def _find_available_driver(self) -> str:
        """尋找可用的 ODBC 驅動程式"""
        available_drivers = pyodbc.drivers()
        preferred_drivers = [
            "ODBC Driver 17 for SQL Server",
            "ODBC Driver 18 for SQL Server",
            "ODBC Driver 13 for SQL Server",
            "SQL Server Native Client 11.0",
            "SQL Server"
        ]
        
        for preferred in preferred_drivers:
            if preferred in available_drivers:
                logger.info(f"使用 ODBC 驅動程式: {preferred}")
                return preferred
        
        raise ValueError(f"找不到可用的 SQL Server ODBC 驅動程式。可用的驅動程式: {', '.join(available_drivers)}")
    
    def execute_query(self, sql: str, params: Optional[tuple] = None) -> List[Dict[str, Any]]:
        """
        執行 SQL 查詢
        
        Args:
            sql: SQL 查詢語句
            params: 查詢參數（用於防止 SQL 注入）
            
        Returns:
            查詢結果列表，每個元素是一個字典（欄位名: 值）
        """
        if self.connection is None:
            raise ConnectionError("資料庫連接未建立")
        
        cursor = self.connection.cursor()
        try:
            if params:
                cursor.execute(sql, params)
            else:
                cursor.execute(sql)
            
            # 取得欄位名稱
            columns = [desc[0] for desc in cursor.description] if cursor.description else []
            
            # 轉換為字典列表
            results = []
            for row in cursor.fetchall():
                row_dict = {col: row[i] for i, col in enumerate(columns)}
                results.append(row_dict)
            
            return results
        finally:
            cursor.close()
    
    def get_cookie_inventory(self) -> List[Dict[str, Any]]:
        """查詢餅乾庫存
        優先使用 config.ini 的自訂 SQL，若無則使用內建預設查詢。
        Returns:庫存資料列表，格式: [
                {'cookie_code': 'COOKIE001',
                    'qty': 1000.0,
                    'warehouse_code': 'SP40',
                    'unit': '片',
                    'cookie_name': '餅乾品名'},...
            ]"""
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        
        cookie_sql = ''
        if 'ERP_QUERIES' in config:
            cookie_sql = config['ERP_QUERIES'].get('cookie_inventory_query', '').strip()
        
        if cookie_sql:
            logger.info("執行餅乾庫存查詢（使用 config.ini 自訂 SQL）")
        else:
            logger.info("執行餅乾庫存查詢（使用內建預設 SQL）")
            cookie_sql = DEFAULT_COOKIE_INVENTORY_QUERY
        
        results = self.execute_query(cookie_sql)
        
        # 標準化欄位名稱
        standardized = []
        for row in results:
            standardized.append({
                'cookie_code': str(row.get('cookie_code', '')).strip(),
                'qty': row.get('qty', 0),
                'warehouse_code': str(row.get('warehouse_code', '')).strip(),
                'unit': str(row.get('unit', '片')).strip(),
                'cookie_name': str(row.get('cookie_name', '')).strip()
            })
        
        return standardized
    
    def get_wip_inventory(self) -> List[Dict[str, Any]]:
        """
        查詢在製品庫存
        優先使用 config.ini 的自訂 SQL，若無則使用內建預設查詢。
        
        Returns:
            在製品資料列表，格式: [
                {
                    'cookie_code': 'COOKIE001',
                    'wip_qty': 500.0,
                    'unit': '片',
                    'cookie_name': '餅乾品名'
                },
                ...
            ]
        """
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        
        wip_sql = ''
        if 'ERP_QUERIES' in config:
            wip_sql = config['ERP_QUERIES'].get('wip_inventory_query', '').strip()
        
        if wip_sql:
            logger.info("執行在製品庫存查詢（使用 config.ini 自訂 SQL）")
        else:
            logger.info("執行在製品庫存查詢（使用內建預設 SQL）")
            wip_sql = DEFAULT_WIP_INVENTORY_QUERY
        
        results = self.execute_query(wip_sql)
        
        # 標準化欄位名稱
        standardized = []
        for row in results:
            standardized.append({
                'mo_number_type': str(row.get('mo_number_type', '')).strip(),
                'mo_number': str(row.get('mo_number', '')).strip(),
                'cookie_code': str(row.get('cookie_code', '')).strip(),
                'wip_qty': row.get('wip_qty', 0),
                'unit': str(row.get('unit', '片')).strip(),
                'cookie_name': str(row.get('cookie_name', '')).strip()
            })
        
        return standardized
    
    def get_item_info_by_codes(self, codes: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        根據代號列表查詢 INVMB 表的品名、生重、熟重
        
        Args:
            codes: 代號列表
            
        Returns:
            字典：{代號: {'cookie_name': '品名', 'raw_weight': 生重, 'cooked_weight': 熟重}}
        """
        if not codes:
            return {}
        
        # 建立 SQL IN 子句（使用安全的字串轉義）
        # 注意：這裡的代號都是內部系統代號，相對安全
        safe_codes = [code.replace("'", "''") for code in codes]  # SQL 注入防護
        codes_str = "','".join(safe_codes)
        sql = f"""
            SELECT 
                MB001 as code,
                COALESCE(MB002, '') as cookie_name,
                COALESCE(MB104, 0) as raw_weight,
                COALESCE(MB105, 0) as cooked_weight
            FROM [AS_online].[dbo].[INVMB]
            WHERE MB001 IN ('{codes_str}')
        """
        
        logger.info(f"查詢 {len(codes)} 個代號的品名、生重、熟重資訊...")
        results = self.execute_query(sql)
        
        # 轉換為字典格式
        info_dict = {}
        for row in results:
            code = str(row.get('code', '')).strip()
            if code:
                info_dict[code] = {
                    'cookie_name': str(row.get('cookie_name', '')).strip(),
                    'raw_weight': float(row.get('raw_weight', 0)) if row.get('raw_weight') else 0.0,
                    'cooked_weight': float(row.get('cooked_weight', 0)) if row.get('cooked_weight') else 0.0
                }
        
        logger.info(f"成功查詢到 {len(info_dict)} 個代號的資訊")
        return info_dict
    
    def get_receipt_data(self, days_back: int = 5) -> List[Dict[str, Any]]:
        """
        查詢入庫單表頭和單身資料（合併查詢）
        
        查詢條件：
        - TF003（入庫日期）在從今天到（今天-{days_back}天）這段期間
        - TF011='P104'
        - TF001 IN ('5801', '5802')
        
        JOIN 條件：
        - MOCTF.TF001 = MOCTG.TG001（單別相同）
        - MOCTF.TF002 = MOCTG.TG002（單號相同）
        
        Args:
            days_back: 往前查詢的天數（預設為5天）
            
        Returns:
            入庫單資料列表，格式: [
                {
                    'cookie_code': 'TG004',
                    'cookie_name': 'TG005',
                    'spec': 'TG006',
                    'unit': 'TG007',
                    'receipt_qty': 'TG013',
                    'receipt_date': 'TF003',
                    'receipt_type': 'TF001',
                    'receipt_number': 'TF002'
                },
                ...
            ]
        """
        from datetime import datetime, timedelta
        
        # 計算日期範圍
        today = datetime.now()
        start_date = today - timedelta(days=days_back)
        
        # 格式化日期為 YYYYMMDD
        today_str = today.strftime('%Y%m%d')
        start_date_str = start_date.strftime('%Y%m%d')
        
        sql = f"""
        SELECT 
            MOCTG.TG004 as cookie_code,
            COALESCE(MOCTG.TG005, '') as cookie_name,
            COALESCE(MOCTG.TG006, '') as spec,
            COALESCE(MOCTG.TG007, '') as unit,
            MOCTG.TG013 as receipt_qty,
            MOCTF.TF003 as receipt_date,
            MOCTF.TF001 as receipt_type,
            MOCTF.TF002 as receipt_number
        FROM [AS_online].[dbo].[MOCTF] MOCTF
        INNER JOIN [AS_online].[dbo].[MOCTG] MOCTG
            ON MOCTF.TF001 = MOCTG.TG001
            AND MOCTF.TF002 = MOCTG.TG002
        WHERE MOCTF.TF003 >= '{start_date_str}'
            AND MOCTF.TF003 <= '{today_str}'
            AND MOCTF.TF011 = 'P104'
            AND MOCTF.TF001 IN ('5801', '5802')
        ORDER BY MOCTF.TF003 DESC, MOCTF.TF001, MOCTF.TF002, MOCTG.TG004
        """
        
        logger.info(f"查詢入庫單資料（日期範圍：{start_date_str} 到 {today_str}，TF011='P104'，TF001 IN ('5801', '5802')）...")
        results = self.execute_query(sql)
        
        # 標準化欄位名稱
        standardized = []
        for row in results:
            standardized.append({
                'cookie_code': str(row.get('cookie_code', '')).strip(),
                'cookie_name': str(row.get('cookie_name', '')).strip(),
                'spec': str(row.get('spec', '')).strip(),
                'unit': str(row.get('unit', '')).strip(),
                'receipt_qty': float(row.get('receipt_qty', 0)) if row.get('receipt_qty') else 0.0,
                'receipt_date': str(row.get('receipt_date', '')).strip(),
                'receipt_type': str(row.get('receipt_type', '')).strip(),
                'receipt_number': str(row.get('receipt_number', '')).strip()
            })
        
        logger.info(f"成功查詢到 {len(standardized)} 筆入庫單資料")
        return standardized
    
    def close(self):
        """關閉資料庫連接"""
        if self.connection:
            self.connection.close()
            logger.info("已關閉資料庫連接")
    
    def __enter__(self):
        """支援 with 語句"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """支援 with 語句"""
        self.close()
