"""LINE Bot 餅乾庫存查詢系統
功能：
1. 接收 LINE Bot 的用戶訊息
2. 解析訊息中的餅乾代號
3. 查詢指定庫別（SP50）的庫存
4. 回覆庫存訊息給用戶
獨立執行，可常駐運行
"""
import os
import sys
# 設置標準輸出編碼為 UTF-8，避免 Windows 編碼問題
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from typing import Optional, Dict, Any, List
from flask import Flask, request, abort
from linebot.v3.messaging import (Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage)
import json
import hmac
import hashlib
import base64
import traceback
import re
import logging
from datetime import datetime

# 導入 ERP 資料庫輔助模組
from Cookies.erp_db_helper import ERPDBHelper

# ==================== 配置區域 ====================
# LINE Bot 配置
LINE_TOKEN_FILE = 'Line_Access_token.json'
if not os.path.exists(LINE_TOKEN_FILE):
    raise FileNotFoundError(f"找不到 LINE Bot 憑證檔案: {LINE_TOKEN_FILE}")

with open(LINE_TOKEN_FILE, 'r', encoding='utf-8') as f:
    line_config = json.load(f)

CHANNEL_ACCESS_TOKEN = line_config.get("CHANNEL_ACCESS_TOKEN")
CHANNEL_SECRET = line_config.get("CHANNEL_SECRET")

if not CHANNEL_ACCESS_TOKEN or not CHANNEL_SECRET:
    raise ValueError("LINE Bot 憑證檔案中缺少 CHANNEL_ACCESS_TOKEN 或 CHANNEL_SECRET")

configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)

# 固定庫別代號
DEFAULT_WAREHOUSE_CODE = 'SP50'

# 設定日誌
logging.basicConfig(level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('linebot_inventory.log', encoding='utf-8'),logging.StreamHandler()])
logger = logging.getLogger(__name__)

# ==================== SQL 查詢定義 ====================
# 查詢指定餅乾代號在指定庫別的庫存
# 基於 config.ini 中的查詢邏輯，但加入 WHERE 條件過濾特定代號和庫別
# 從 INVMB 產品主檔取得品名和庫存單位
COOKIE_INVENTORY_BY_CODE_SQL = """
    SELECT 
        LC.LC001 as cookie_code,
        MB.MB002 as product_name,
        LC.LC003 as warehouse_code,
        LC.LC004 + COALESCE(SUM(LA.LA011 * LA.LA005), 0) as qty,
        MB.MB004 as unit
    FROM [AS_online].[dbo].[INVLC] LC
    LEFT JOIN [AS_online].[dbo].[INVLA] LA 
        ON LA.LA001 = LC.LC001 
        AND LA.LA009 = LC.LC003
        AND LA.LA004 >= '20251201'
    LEFT JOIN [AS_online].[dbo].[INVMB] MB
        ON MB.MB001 = LC.LC001
    WHERE LC.LC001 = ?
        AND LC.LC001 IS NOT NULL 
        AND LC.LC002 = '202512' 
        AND LC.LC003 = ?
    GROUP BY LC.LC001, LC.LC003, LC.LC004, MB.MB002, MB.MB004
"""

# 關鍵字查詢 SQL：查詢 SP50 庫別中品名包含關鍵字且有庫存的品項
# 支援單個字母關鍵字查詢（A、B、E、F、G、H、J、K、Y 等）
# 使用 UPPER 函數確保大小寫不敏感
COOKIE_INVENTORY_BY_KEYWORD_SQL = """
    SELECT 
        LC.LC001 as cookie_code,
        MB.MB002 as product_name,
        LC.LC003 as warehouse_code,
        LC.LC004 + COALESCE(SUM(LA.LA011 * LA.LA005), 0) as qty,
        MB.MB004 as unit
    FROM [AS_online].[dbo].[INVLC] LC
    LEFT JOIN [AS_online].[dbo].[INVLA] LA 
        ON LA.LA001 = LC.LC001 
        AND LA.LA009 = LC.LC003
        AND LA.LA004 >= '20251201'
    LEFT JOIN [AS_online].[dbo].[INVMB] MB
        ON MB.MB001 = LC.LC001
    WHERE LC.LC001 IS NOT NULL 
        AND LC.LC002 = '202512' 
        AND LC.LC003 = ?
        AND UPPER(MB.MB002) LIKE UPPER(?)
    GROUP BY LC.LC001, LC.LC003, LC.LC004, MB.MB002, MB.MB004
    HAVING (LC.LC004 + COALESCE(SUM(LA.LA011 * LA.LA005), 0)) > 0
    ORDER BY MB.MB002, LC.LC001
"""

# ==================== 訊息解析函數 ====================
def parse_user_input(message: str) -> tuple[str, str]:
    """
    解析使用者輸入，判斷是品號還是關鍵字
    
    品號格式：
    - 前5碼：必須是數字
    - 第6碼：數字或英文字母
    - 第7碼：英文字母（可選，沒有第7碼也可以）
    - 總長度：6碼或7碼，中間沒有空白
    
    如果不是品號格式，則視為關鍵字
    
    Args:
        message: 使用者輸入的訊息
        
    Returns:
        tuple: (類型, 值)
        - 類型: 'code' 表示品號，'keyword' 表示關鍵字
        - 值: 品號（大寫）或關鍵字（去除空白）
    """
    if not message:
        return ('keyword', '')
    
    message = message.strip()
    
    # 正則表達式：匹配品號格式
    # 格式：^[0-9]{5}[A-Za-z0-9][A-Za-z]?$
    pattern = r'^[0-9]{5}[A-Za-z0-9][A-Za-z]?$'
    
    # 檢查整個訊息是否符合品號格式
    if re.match(pattern, message):
        cookie_code = message.upper()
        logger.info(f"從訊息 '{message}' 中識別為品號: {cookie_code}")
        return ('code', cookie_code)
    
    # 不符合品號格式，視為關鍵字
    keyword = message.strip()
    logger.info(f"從訊息 '{message}' 中識別為關鍵字: {keyword}")
    return ('keyword', keyword)


# ==================== 庫存查詢函數 ====================
def query_cookie_inventory(cookie_code: str, warehouse_code: str = DEFAULT_WAREHOUSE_CODE) -> Optional[Dict[str, Any]]:
    """
    查詢指定餅乾代號在指定庫別的庫存
    
    Args:
        cookie_code: 餅乾代號
        warehouse_code: 庫別代號（預設為 SP50）
        
    Returns:
        庫存資料字典，格式: {
            'cookie_code': 'COOKIE001',
            'product_name': '品名',
            'warehouse_code': 'SP50',
            'qty': 1000.0,
            'unit': '片' 或 '包'（從資料庫取得）
        }
        如果查無資料或發生錯誤則返回 None
    """
    try:
        logger.info(f"查詢庫存: 餅乾代號={cookie_code}, 庫別={warehouse_code}")
        
        with ERPDBHelper() as erp_db:
            # 使用參數化查詢防止 SQL 注入
            results = erp_db.execute_query(
                COOKIE_INVENTORY_BY_CODE_SQL,
                params=(cookie_code, warehouse_code)
            )
            
            if results and len(results) > 0:
                row = results[0]
                # 取得各欄位資料
                cookie_code = str(row.get('cookie_code', '')).strip()
                product_name = str(row.get('product_name', '')).strip() if row.get('product_name') else ''
                warehouse_code = str(row.get('warehouse_code', '')).strip()
                qty = float(row.get('qty', 0)) if row.get('qty') is not None else 0.0
                unit = str(row.get('unit', '')).strip() if row.get('unit') else ''
                
                inventory_data = {
                    'cookie_code': cookie_code,
                    'product_name': product_name,
                    'warehouse_code': warehouse_code,
                    'qty': qty,
                    'unit': unit
                }
                logger.info(f"查詢成功: {inventory_data}")
                return inventory_data
            else:
                logger.warning(f"查無資料: 餅乾代號={cookie_code}, 庫別={warehouse_code}")
                return None
                
    except Exception as e:
        logger.error(f"查詢庫存時發生錯誤: {str(e)}")
        logger.error(traceback.format_exc())
        return None


def query_cookie_inventory_by_keyword(keyword: str, warehouse_code: str = DEFAULT_WAREHOUSE_CODE) -> List[Dict[str, Any]]:
    """
    使用關鍵字查詢 SP50 庫別中品名包含關鍵字且有庫存的品項
    
    Args:
        keyword: 關鍵字（會用於 LIKE 查詢，自動加上 % 前後綴）
        warehouse_code: 庫別代號（預設為 SP50）
        
    Returns:
        庫存資料列表，格式: [
            {
                'cookie_code': 'COOKIE001',
                'product_name': '品名',
                'warehouse_code': 'SP50',
                'qty': 1000.0,
                'unit': '片'
            },
            ...
        ]
        如果查無資料則返回空列表
    """
    try:
        logger.info(f"關鍵字查詢庫存: 關鍵字={keyword}, 庫別={warehouse_code}")
        
        # 關鍵字前後加上 % 用於 LIKE 查詢
        keyword_pattern = f'%{keyword}%'
        
        with ERPDBHelper() as erp_db:
            # 使用參數化查詢防止 SQL 注入
            results = erp_db.execute_query(
                COOKIE_INVENTORY_BY_KEYWORD_SQL,
                params=(warehouse_code, keyword_pattern)
            )
            
            inventory_list = []
            if results and len(results) > 0:
                for row in results:
                    cookie_code = str(row.get('cookie_code', '')).strip()
                    product_name = str(row.get('product_name', '')).strip() if row.get('product_name') else ''
                    wh_code = str(row.get('warehouse_code', '')).strip()
                    qty = float(row.get('qty', 0)) if row.get('qty') is not None else 0.0
                    unit = str(row.get('unit', '')).strip() if row.get('unit') else ''
                    
                    inventory_data = {
                        'cookie_code': cookie_code,
                        'product_name': product_name,
                        'warehouse_code': wh_code,
                        'qty': qty,
                        'unit': unit
                    }
                    inventory_list.append(inventory_data)
                
                logger.info(f"關鍵字查詢成功: 找到 {len(inventory_list)} 筆資料")
            else:
                logger.warning(f"關鍵字查詢無資料: 關鍵字={keyword}, 庫別={warehouse_code}")
            
            return inventory_list
                
    except Exception as e:
        logger.error(f"關鍵字查詢庫存時發生錯誤: {str(e)}")
        logger.error(traceback.format_exc())
        return []


# ==================== 回覆格式化函數 ====================
def format_inventory_reply(inventory_data: Dict[str, Any]) -> str:
    """
    格式化單筆庫存資料為 LINE Bot 回覆訊息
    
    Args:
        inventory_data: 庫存資料字典
        
    Returns:
        格式化的訊息字串
    """
    cookie_code = inventory_data.get('cookie_code', '')
    product_name = inventory_data.get('product_name', '').strip()
    warehouse_code = inventory_data.get('warehouse_code', '')
    qty = inventory_data.get('qty', 0)
    unit = inventory_data.get('unit', '').strip()
    update_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 格式化數量（加入千分位）
    qty_str = f"{qty:,.0f}" if qty == int(qty) else f"{qty:,.2f}"
    
    # 如果有單位則顯示，沒有則不顯示
    qty_display = f"{qty_str} {unit}" if unit else qty_str
    
    # 建立回覆訊息
    reply_lines = [
        "📦 庫存查詢結果",
        "",
        f"品號：{cookie_code}"
    ]
    
    # 如果有品名則顯示
    if product_name:
        reply_lines.append(f"品名：{product_name}")
    
    reply_lines.extend([
        f"庫別代號：{warehouse_code}",
        f"目前庫存：{qty_display}",
        f"查詢時間：{update_time}"
    ])
    
    reply = "\n".join(reply_lines)
    return reply


def format_keyword_reply(inventory_list: List[Dict[str, Any]], keyword: str) -> str:
    """
    格式化關鍵字查詢的多筆結果為 LINE Bot 回覆訊息
    
    Args:
        inventory_list: 庫存資料列表
        keyword: 查詢的關鍵字
        
    Returns:
        格式化的訊息字串
    """
    if not inventory_list:
        return f"❌ 查無符合條件的庫存資料（關鍵字：「{keyword}」）"
    
    # 建立回覆訊息，只顯示品名、品號、庫存數量、庫存單位
    reply_lines = []
    
    # 格式化每一筆資料
    for item in inventory_list:
        cookie_code = item.get('cookie_code', '')
        product_name = item.get('product_name', '').strip()
        qty = item.get('qty', 0)
        unit = item.get('unit', '').strip()
        
        # 格式化數量
        qty_str = f"{qty:,.0f}" if qty == int(qty) else f"{qty:,.2f}"
        qty_display = f"{qty_str} {unit}" if unit else qty_str
        
        # 顯示格式：品名 品號 庫存數量 庫存單位
        display_name = product_name if product_name else cookie_code
        reply_lines.append(f"{display_name} {cookie_code} {qty_display}")
    
    reply = "\n".join(reply_lines)
    return reply


def format_error_reply(error_type: str, cookie_code: str = None) -> str:
    """
    格式化錯誤回覆訊息
    
    Args:
        error_type: 錯誤類型（'no_code', 'not_found', 'system_error'）
        cookie_code: 餅乾代號（可選）
        
    Returns:
        錯誤訊息字串
    """
    if error_type == 'no_code':
        return """❌ 無法識別輸入

請輸入：
1️⃣ 品號（6-7碼格式）：
   • 前5碼：必須是數字
   • 第6碼：數字或英文字母
   • 第7碼：英文字母（可選）
   
   範例：12345A、123456、12345AB

2️⃣ 關鍵字（品名搜尋）：
   輸入品名中的關鍵字，例如：牛奶、草莓
   系統會搜尋 SP50 庫別中包含該關鍵字的所有品項"""
    
    elif error_type == 'not_found':
        code_msg = f"（代號：{cookie_code}）" if cookie_code else ""
        return f"""❌ 查無庫存資料{code_msg}

可能原因：
• 該餅乾代號不存在
• 該代號在 SP50 庫別中沒有庫存
• 代號輸入錯誤

請確認代號是否正確，或聯繫管理員。"""
    
    elif error_type == 'system_error':
        return """⚠️ 系統暫時無法查詢

請稍後再試，或聯繫系統管理員。

錯誤已記錄，我們會盡快處理。"""
    
    else:
        return "❌ 發生未知錯誤，請稍後再試。"


# ==================== LINE Bot 處理函數 ====================
def truncate_message(text: str, max_length: int = 5000) -> str:
    """
    截斷訊息以符合 LINE 文字訊息長度限制（最多 5000 字元）
    
    Args:
        text: 原始訊息
        max_length: 最大長度（預設 5000）
        
    Returns:
        截斷後的訊息
    """
    if len(text) <= max_length:
        return text
    truncated = text[:max_length - 50]
    return truncated + "\n\n...（訊息過長，已截斷）"


def process_user_message(user_text: str) -> str:
    """
    處理使用者訊息並返回回覆
    
    根據輸入內容判斷是品號查詢還是關鍵字查詢：
    - 符合品號格式（6-7碼，前5碼數字+第6碼數字或英文+第7碼英文或無）：品號查詢
    - 不符合品號格式：關鍵字查詢（在品名中搜尋）
    
    Args:
        user_text: 使用者輸入的文字
        
    Returns:
        回覆訊息
    """
    try:
        # 解析使用者輸入，判斷是品號還是關鍵字
        input_type, input_value = parse_user_input(user_text)
        
        if not input_value:
            return format_error_reply('no_code')
        
        if input_type == 'code':
            # 品號查詢
            inventory_data = query_cookie_inventory(input_value, DEFAULT_WAREHOUSE_CODE)
            
            if inventory_data is None:
                return format_error_reply('not_found', input_value)
            
            # 格式化單筆回覆
            reply = format_inventory_reply(inventory_data)
            return reply
        
        else:
            # 關鍵字查詢
            inventory_list = query_cookie_inventory_by_keyword(input_value, DEFAULT_WAREHOUSE_CODE)
            
            # 格式化多筆回覆
            reply = format_keyword_reply(inventory_list, input_value)
            return reply
        
    except Exception as e:
        logger.error(f"處理使用者訊息時發生錯誤: {str(e)}")
        logger.error(traceback.format_exc())
        return format_error_reply('system_error')


# ==================== Flask 應用程式 ====================
app = Flask(__name__)


@app.route("/", methods=["POST"])
def webhook():
    """處理 LINE Bot webhook"""
    try:
        # 驗證簽章
        body = request.get_data(as_text=True)
        signature = request.headers.get("X-Line-Signature", "")
        hash = hmac.new(CHANNEL_SECRET.encode(), body.encode(), hashlib.sha256).digest()
        
        if signature != base64.b64encode(hash).decode():
            logger.warning("❌ 簽章驗證失敗")
            abort(400)
        
        # 解析 JSON 資料
        data = request.get_json()
        
        # 確保有事件，且事件類型是訊息，且訊息類型是文字
        if 'events' in data and data['events'] and \
           data['events'][0].get('type') == 'message' and \
           data['events'][0].get('message', {}).get('type') == 'text':
            
            # 提取文字訊息和 replyToken
            user_text = data['events'][0]['message']['text']
            reply_token = data['events'][0]['replyToken']
            
            logger.info(f"👤 收到使用者訊息: {user_text}")
            
            # 處理訊息並產生回覆
            reply_text = process_user_message(user_text)
            
            # 處理訊息長度限制
            reply_text = truncate_message(reply_text)
            
            # 回覆訊息給用戶
            with ApiClient(configuration) as api_client:
                messaging_api = MessagingApi(api_client)
                messaging_api.reply_message_with_http_info(
                    ReplyMessageRequest(
                        reply_token=reply_token,
                        messages=[TextMessage(text=reply_text)]
                    )
                )
            
            logger.info(f"✅ 已回覆使用者")
        
        return "OK", 200
    
    except Exception as e:
        logger.error(f"❌ Webhook 處理發生錯誤: {str(e)}")
        logger.error(traceback.format_exc())
        return "Error", 500


@app.route("/health", methods=["GET"])
def health_check():
    """健康檢查端點"""
    return {"status": "ok", "service": "LINEBOT_Cookie_Inventory"}, 200


# ==================== 主程式 ====================
if __name__ == "__main__":
    print("=" * 60)
    print("🚀 LINE Bot 餅乾庫存查詢系統啟動中...")
    print("=" * 60)
    print(f"📋 服務名稱: LINE Bot Cookie Inventory Query")
    print(f"📁 憑證檔案: {LINE_TOKEN_FILE}")
    print(f"🏢 預設庫別: {DEFAULT_WAREHOUSE_CODE}")
    print(f"🔗 Webhook URL: http://localhost:3001/")
    print(f"📝 日誌檔案: linebot_inventory.log")
    print("=" * 60)
    print("💡 使用說明:")
    print("   使用者可輸入品號查詢庫存")
    print("   格式：前5碼數字 + 第6碼數字或英文 + 第7碼英文（可選）")
    print("   範例：401500D、501500、40382JD")
    print("=" * 60)
    
    try:
        app.run(host='0.0.0.0', port=3001, debug=False)
    except KeyboardInterrupt:
        print("\n👋 系統已停止")
    except Exception as e:
        logger.error(f"系統啟動失敗: {str(e)}")
        logger.error(traceback.format_exc())
        sys.exit(1)
