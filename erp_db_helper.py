"""
ERP 資料庫連接工具模組
專用於 MS-SQL Server 資料庫連接，用於查詢餅乾庫存
"""
import configparser
from typing import List, Dict, Any
import logging

try:
    import pyodbc
except ImportError:
    raise ImportError("請安裝 pyodbc: pip install pyodbc")

logger = logging.getLogger(__name__)


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
    
    def execute_query(self, sql: str, params: tuple = None) -> List[Dict[str, Any]]:
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
        """
        查詢餅乾庫存（從 config.ini 讀取 SQL 查詢）
        
        Returns:
            庫存資料列表，格式: [
                {
                    'cookie_code': 'COOKIE001',
                    'qty': 1000.0,
                    'warehouse_code': 'SP40',
                    'unit': '片'
                },
                ...
            ]
        """
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        
        if 'ERP_QUERIES' not in config:
            raise ValueError("config.ini 中缺少 [ERP_QUERIES] 區段")
        
        query_config = config['ERP_QUERIES']
        cookie_sql = query_config.get('cookie_inventory_query', '')
        
        if not cookie_sql:
            raise ValueError("config.ini 中缺少 cookie_inventory_query 設定")
        
        logger.info("執行餅乾庫存查詢")
        results = self.execute_query(cookie_sql)
        
        # 標準化欄位名稱
        standardized = []
        for row in results:
            standardized.append({
                'cookie_code': str(row.get('cookie_code', '')).strip(),
                'qty': row.get('qty', 0),
                'warehouse_code': str(row.get('warehouse_code', '')).strip(),
                'unit': str(row.get('unit', '片')).strip()
            })
        
        return standardized
    
    def get_wip_inventory(self) -> List[Dict[str, Any]]:
        """
        查詢在製品庫存（從 config.ini 讀取 SQL 查詢）
        
        Returns:
            在製品資料列表，格式: [
                {
                    'cookie_code': 'COOKIE001',
                    'wip_qty': 500.0,
                    'unit': '片'
                },
                ...
            ]
        """
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        
        if 'ERP_QUERIES' not in config:
            raise ValueError("config.ini 中缺少 [ERP_QUERIES] 區段")
        
        query_config = config['ERP_QUERIES']
        wip_sql = query_config.get('wip_inventory_query', '')
        
        if not wip_sql:
            raise ValueError("config.ini 中缺少 wip_inventory_query 設定")
        
        logger.info("執行在製品庫存查詢")
        results = self.execute_query(wip_sql)
        
        # 標準化欄位名稱
        standardized = []
        for row in results:
            standardized.append({
                'mo_number_type': str(row.get('mo_number_type', '')).strip(),
                'mo_number': str(row.get('mo_number', '')).strip(),
                'cookie_code': str(row.get('cookie_code', '')).strip(),
                'wip_qty': row.get('wip_qty', 0),
                'unit': str(row.get('unit', '片')).strip()
            })
        
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
