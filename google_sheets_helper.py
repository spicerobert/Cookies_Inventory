"""
Google Sheets 操作工具模組
用於讀寫 Google Sheets 資料
"""
import configparser
import json
import gspread
from google.oauth2.service_account import Credentials
from typing import List, Dict, Optional, Any


class GoogleSheetsHelper:
    """Google Sheets 操作輔助類別"""
    
    def __init__(self, config_file='config.ini', credentials_file='service_account.json'):
        """
        初始化 Google Sheets 連接
        
        Args:
            config_file: 設定檔路徑
            credentials_file: 服務帳戶憑證檔案路徑
        """
        self.config_file = config_file
        self.credentials_file = credentials_file
        self.client = None
        self.spreadsheet = None
        self._connect()
    
    def _connect(self):
        """建立 Google Sheets 連接"""
        # 讀取設定檔
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        sheet_url = config['GOOGLE_SHEETS']['data_sheet_url']
        
        # 讀取服務帳戶憑證
        with open(self.credentials_file, 'r', encoding='utf-8') as f:
            creds_info = json.load(f)
        
        # 建立憑證物件
        credentials = Credentials.from_service_account_info(
            creds_info,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        
        # 建立 gspread 客戶端
        self.client = gspread.authorize(credentials)
        
        # 開啟試算表
        sheet_id = self._extract_sheet_id(sheet_url)
        self.spreadsheet = self.client.open_by_key(sheet_id)
    
    def _extract_sheet_id(self, url: str) -> str:
        """從 Google Sheets URL 提取 Sheet ID"""
        parts = url.split('/')
        sheet_id_index = parts.index('d') + 1
        if sheet_id_index < len(parts):
            return parts[sheet_id_index].split('?')[0]
        raise ValueError("無法從 URL 中提取 Sheet ID")
    
    def get_worksheet(self, worksheet_name: str, create_if_not_exists: bool = False, rows: int = 1000, cols: int = 26) -> Optional[gspread.Worksheet]:
        """
        取得工作表
        
        Args:
            worksheet_name: 工作表名稱
            create_if_not_exists: 如果不存在是否建立
            rows: 建立時的列數
            cols: 建立時的欄數
            
        Returns:
            Worksheet 物件，如果不存在且不建立則返回 None
        """
        try:
            return self.spreadsheet.worksheet(worksheet_name)
        except gspread.exceptions.WorksheetNotFound:
            if create_if_not_exists:
                return self.spreadsheet.add_worksheet(title=worksheet_name, rows=rows, cols=cols)
            return None
    
    def create_worksheet(self, worksheet_name: str, rows: int = 1000, cols: int = 26) -> gspread.Worksheet:
        """
        建立新工作表
        
        Args:
            worksheet_name: 工作表名稱
            rows: 列數
            cols: 欄數
            
        Returns:
            Worksheet 物件
        """
        try:
            # 如果已存在，先刪除
            existing = self.spreadsheet.worksheet(worksheet_name)
            self.spreadsheet.del_worksheet(existing)
        except gspread.exceptions.WorksheetNotFound:
            pass
        
        return self.spreadsheet.add_worksheet(title=worksheet_name, rows=rows, cols=cols)
    
    def read_worksheet(self, worksheet_name: str) -> List[List[Any]]:
        """
        讀取整個工作表的資料
        
        Args:
            worksheet_name: 工作表名稱
            
        Returns:
            二維列表，每一行是一個列表
        """
        worksheet = self.get_worksheet(worksheet_name)
        if worksheet is None:
            return []
        return worksheet.get_all_values()
    
    def write_worksheet(self, worksheet_name: str, data: List[List[Any]], start_cell: str = 'A1'):
        """
        寫入資料到工作表
        
        Args:
            worksheet_name: 工作表名稱
            data: 二維列表資料
            start_cell: 起始儲存格位置（例如 'A1'）
        """
        worksheet = self.get_worksheet(worksheet_name, create_if_not_exists=True)
        worksheet.update(range_name=start_cell, values=data)
    
    def append_rows(self, worksheet_name: str, rows: List[List[Any]]):
        """
        在工作表末尾新增資料列
        
        Args:
            worksheet_name: 工作表名稱
            rows: 要新增的資料列（二維列表）
        """
        worksheet = self.get_worksheet(worksheet_name, create_if_not_exists=True)
        worksheet.append_rows(rows)
    
    def clear_worksheet(self, worksheet_name: str):
        """
        清空工作表內容
        
        Args:
            worksheet_name: 工作表名稱
        """
        worksheet = self.get_worksheet(worksheet_name)
        if worksheet:
            worksheet.clear()
    
    def update_cell(self, worksheet_name: str, cell: str, value: Any):
        """
        更新單一儲存格
        
        Args:
            worksheet_name: 工作表名稱
            cell: 儲存格位置（例如 'A1'）
            value: 要寫入的值
        """
        worksheet = self.get_worksheet(worksheet_name, create_if_not_exists=True)
        worksheet.update(range_name=cell, values=[[value]])
    
    def list_worksheets(self) -> List[str]:
        """
        列出所有工作表名稱
        
        Returns:
            工作表名稱列表
        """
        return [ws.title for ws in self.spreadsheet.worksheets()]
    
    def get_index_dict(self) -> Dict[str, Dict[str, str]]:
        """
        讀取 Index 工作表，建立代號與名稱的對應字典
        
        Returns:
            字典結構: {
                '餅乾': {'COOKIE001': '奶油餅乾', ...},
                '禮盒': {'BOX001': '經典禮盒', ...},
                '產線': {'LINE_A': '產線A', ...}
            }
        """
        try:
            data = self.read_worksheet('Index')
            if not data or len(data) < 2:
                return {'餅乾': {}, '禮盒': {}, '產線': {}}
            
            # 第一行是標題: ['類型', '代號', '名稱', '備註']
            result = {'餅乾': {}, '禮盒': {}, '產線': {}}
            
            for row in data[1:]:  # 跳過標題行
                if len(row) >= 3 and row[0] and row[1] and row[2]:
                    item_type = row[0].strip()
                    code = row[1].strip()
                    name = row[2].strip()
                    
                    # 將類型映射到標準鍵值
                    type_mapping = {
                        '餅乾': '餅乾',
                        '禮盒': '禮盒',
                        '產線': '產線',
                        'Cookie': '餅乾',
                        'Box': '禮盒',
                        'Line': '產線',
                        'LINE': '產線'
                    }
                    
                    mapped_type = type_mapping.get(item_type, item_type)
                    if mapped_type in result:
                        result[mapped_type][code] = name
            
            return result
        except Exception:
            return {'餅乾': {}, '禮盒': {}, '產線': {}}
    
    def get_name_by_code(self, code: str, item_type: str = '餅乾') -> Optional[str]:
        """
        根據代號查詢名稱
        
        Args:
            code: 代號
            item_type: 類型（'餅乾'、'禮盒'、'產線'）
            
        Returns:
            名稱，如果找不到則返回 None
        """
        index_dict = self.get_index_dict()
        type_mapping = {
            '餅乾': '餅乾',
            '禮盒': '禮盒',
            '產線': '產線',
            'Cookie': '餅乾',
            'Box': '禮盒',
            'Line': '產線',
            'LINE': '產線'
        }
        mapped_type = type_mapping.get(item_type, item_type)
        return index_dict.get(mapped_type, {}).get(code.strip())


def initialize_sheets_structure(helper: GoogleSheetsHelper):
    """
    初始化系統所需的工作表結構
    
    Args:
        helper: GoogleSheetsHelper 實例
    """
    # 定義所有需要的工作表及其欄位標題
    # 注意：名稱欄位已移除，統一使用 Index 工作表作為代號對應表
    sheets_structure = {
        'Index': [
            ['類型', '代號', '名稱', '備註']
        ],
        'BOM': [
            ['禮盒代號', '餅乾代號', '每盒片數', '備註']
        ],
        '組裝計劃': [
            ['日期', '禮盒代號', '計畫組裝數量', '已完成數量', '狀態', '備註']
        ],
        '出貨預測': [
            ['出貨日期', '禮盒代號', '預估出貨數量', '客戶類別', '備註']
        ],
        '庫存狀態': [
            ['餅乾代號', '目前庫存數量', '庫別代號', '單位', '最後更新日期']
        ],
        '成品庫存': [
            ['禮盒代號', '目前庫存數量', '單位', '最後更新日期']
        ],
        '生產參數': [
            ['餅乾代號', '產線歸屬', '每小時產量_片', '前置天數_天', '最小批量_片', '換線分類', '備註']
        ],
        '組裝產能': [
            ['禮盒代號', '組裝速度_盒每小時每人', '每日最大產能_盒', '備註']
        ],
        '產線產能': [
            ['產線代號', '每日最大產能_小時', '備註']
        ],
        '訂單與預測': [
            ['禮盒代號', '總目標銷售數量', '已確認訂單數量', '預測數量', '單位', '備註']
        ],
        '生產排程建議': [
            ['日期', '產線代號', '餅乾代號', '建議生產數量_片', '預計開始時間', '預計完成時間', '狀態', '備註']
        ],
        '組裝調整建議': [
            ['原始日期', '調整後日期', '禮盒代號', '原始數量', '調整後數量', '調整原因', '狀態']
        ],
        '齊料缺口分析': [
            ['日期', '禮盒代號', '缺料餅乾代號', '缺口數量_片', '優先級', '狀態']
        ]
    }
    
    print("正在建立工作表結構...")
    created_sheets = []
    updated_sheets = []
    existing_sheets = helper.list_worksheets()
    
    for sheet_name, headers in sheets_structure.items():
        if sheet_name not in existing_sheets:
            worksheet = helper.create_worksheet(sheet_name, rows=1000, cols=26)
            helper.write_worksheet(sheet_name, headers)
            created_sheets.append(sheet_name)
            print(f"  [建立] {sheet_name}")
        else:
            # 檢查現有工作表的標題行是否與新結構一致
            worksheet = helper.get_worksheet(sheet_name)
            if worksheet:
                existing_headers = worksheet.row_values(1)
                if existing_headers != headers[0]:
                    # 更新標題行（只更新第一行，不影響資料）
                    worksheet.update(range_name='1:1', values=[headers[0]])
                    updated_sheets.append(sheet_name)
                    print(f"  [更新標題] {sheet_name}")
                else:
                    print(f"  [已存在] {sheet_name}")
            else:
                print(f"  [已存在但無法讀取] {sheet_name}")
    
    if created_sheets:
        print(f"\n[完成] 已建立 {len(created_sheets)} 個新工作表")
    if updated_sheets:
        print(f"[完成] 已更新 {len(updated_sheets)} 個工作表的標題行")
    if not created_sheets and not updated_sheets:
        print("\n[完成] 所有工作表已存在且結構正確，無需更新")


if __name__ == '__main__':
    # 測試功能
    helper = GoogleSheetsHelper()
    print(f"已連接到試算表: {helper.spreadsheet.title}")
    print(f"現有工作表: {', '.join(helper.list_worksheets())}")

