"""
RAG + LLM + LINE Bot 整合系統
整合功能：
1. 接收 LINE Bot 的用戶問題
2. 透過 RAG 系統檢索相關文檔
3. 使用 LLM 生成回答
4. 將結果回覆給 LINE 用戶

整合來源：
- RAG_3_LLM.py: RAG+LLM 查詢系統
- LINEBOT(第三版HMAC驗證和回覆).py: LINE Bot 回覆系統
"""
import os
import sys
# 設置標準輸出編碼為 UTF-8，避免 Windows 編碼問題
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from typing import Any
from flask import Flask, request, abort
from linebot.v3.messaging import (Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage)
import json
import hmac
import hashlib
import base64
import traceback

# RAG 系統相關導入
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnableLambda
from langchain_ollama import ChatOllama
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma
import chromadb

# ==================== 配置區域 ====================
# LINE Bot 配置
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path, 'r', encoding='utf-8') as f:
    config = json.load(f)
CHANNEL_ACCESS_TOKEN = config.get("CHANNEL_ACCESS_TOKEN")
CHANNEL_SECRET = config.get("CHANNEL_SECRET")
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)

# RAG 系統配置
current_dir = os.path.dirname(os.path.abspath(__file__))
persistent_dir = os.path.join(current_dir, "db", "chroma_db")
client = chromadb.PersistentClient(path=persistent_dir)

embeddings = HuggingFaceEmbeddings(model_name="BAAI/bge-m3")
db = Chroma(
    embedding_function=embeddings,
    client=client,
    collection_name="Leadership"
)

# LLM 配置
llm = ChatOllama(
    model="gpt-oss:120b-cloud",
    model_kwargs={"keep_alive": -1}, 
    base_url="http://localhost:11434"
)

# ==================== RAG 系統函數 ====================
def retriever_docs(question):
    """檢索相關文檔"""
    retriever = db.as_retriever(
        search_type="similarity",
        search_kwargs={"k": 10}
    )
    relevant_docs = retriever.invoke(question)
    
    content = "\n\n".join([doc.page_content for doc in relevant_docs])
    
    # 檢查 content 是否有內容，只在為空時顯示警告
    if not content or not content.strip():
        print("⚠ 警告: Content 為空或沒有內容！")
    
    return {
        "context": content,
        "question": question,
    }

# Prompt 模板
template = """你是一個公司的財務主管。請根據以下參考資料回答使用者的問題。
參考資料：
{context}
使用者問題：{question}
請用繁體中文回答，並且：
1. 主要根據參考資料回答，但可以加上你的專業知識和經驗
2. 如果參考資料中沒有答案，請誠實說「我在資料中找不到相關資訊」
3. 回答要清楚、具體、有條理
回答："""

prompt = ChatPromptTemplate.from_template(template)

# 建立 RAG Chain
rag_chain = (
    RunnableLambda[Any, dict[str, Any]](retriever_docs)
    | prompt
    | llm
    | StrOutputParser()
)

# ==================== LINE Bot 處理函數 ====================
def truncate_message(text, max_length=5000):
    """
    截斷訊息以符合 LINE 文字訊息長度限制（最多 5000 字元）
    如果超過限制，會在結尾添加提示
    """
    if len(text) <= max_length:
        return text
    # 截斷並添加提示
    truncated = text[:max_length - 50]
    return truncated + "\n\n...（回答過長，已截斷）"

def query_rag_system(question):
    """
    查詢 RAG 系統並返回結果
    包含錯誤處理
    """
    try:
        print(f"📥 收到問題: {question}")
        result = rag_chain.invoke(question)
        print(f"✅ RAG 查詢成功")
        return result
    except Exception as e:
        error_msg = f"❌ RAG 查詢發生錯誤: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return f"抱歉，查詢時發生錯誤：{str(e)}"

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
            print("❌ 簽章驗證失敗")
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
            
            print(f"👤 用戶訊息: {user_text}")
            
            # 查詢 RAG 系統
            answer = query_rag_system(user_text)
            
            # 處理訊息長度限制
            answer = truncate_message(answer)
            
            # 回覆訊息給用戶
            with ApiClient(configuration) as api_client:
                messaging_api = MessagingApi(api_client)
                messaging_api.reply_message_with_http_info(
                    ReplyMessageRequest(
                        reply_token=reply_token,
                        messages=[TextMessage(text=answer)]
                    )
                )
            
            print(f"✅ 已回覆用戶")
        
        return "OK", 200
    
    except Exception as e:
        print(f"❌ Webhook 處理發生錯誤: {str(e)}")
        print(traceback.format_exc())
        return "Error", 500

@app.route("/health", methods=["GET"])
def health_check():
    """健康檢查端點"""
    return {"status": "ok", "service": "RAG_LINEBOT_LLM"}, 200

# ==================== 主程式 ====================
if __name__ == "__main__":
    print("=" * 50)
    print("🚀 RAG + LLM + LINE Bot 整合系統啟動中...")
    print("=" * 50)
    print(f"📁 資料庫路徑: {persistent_dir}")
    print(f"🤖 LLM 模型: gemma3:27b")
    print(f"🔗 LINE Bot Webhook: http://localhost:3001/")
    print("=" * 50)
    app.run(port=3001, debug=True)

