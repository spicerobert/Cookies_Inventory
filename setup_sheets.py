"""
設定 Google Sheets 工作表結構
執行此腳本將建立系統所需的所有工作表
"""
from google_sheets_helper import GoogleSheetsHelper, initialize_sheets_structure


def main():
    """主程式：建立工作表結構"""
    print("=" * 60)
    print("餅乾禮盒生產排程系統 - Google Sheets 工作表初始化")
    print("=" * 60)
    
    try:
        # 建立 Google Sheets 連接
        print("\n[步驟 1/2] 連接 Google Sheets...")
        helper = GoogleSheetsHelper()
        print(f"[成功] 已連接到試算表: {helper.spreadsheet.title}")
        
        # 列出現有工作表
        existing_sheets = helper.list_worksheets()
        if existing_sheets:
            print(f"\n現有工作表 ({len(existing_sheets)} 個):")
            for i, sheet_name in enumerate(existing_sheets, 1):
                print(f"  {i}. {sheet_name}")
        
        # 建立工作表結構
        print("\n[步驟 2/2] 建立工作表結構...")
        initialize_sheets_structure(helper)
        
        # 顯示最終工作表列表
        print("\n" + "=" * 60)
        print("所有工作表列表:")
        final_sheets = helper.list_worksheets()
        for i, sheet_name in enumerate(final_sheets, 1):
            print(f"  {i}. {sheet_name}")
        
        print("\n" + "=" * 60)
        print("[完成] 工作表結構設定完成！")
        print("=" * 60)
        print("\n接下來您可以開始在各工作表中填入基礎資料：")
        print("  1. Index - 代號對應表（請先填入餅乾、禮盒、產線的代號與名稱對應）")
        print("  2. BOM - 禮盒組成表")
        print("  3. 組裝計劃 - 後段組裝作業排程")
        print("  4. 出貨預測 - 預估出貨計畫")
        print("  5. 庫存狀態 - 目前餅乾庫存")
        print("  6. 生產參數 - 餅乾生產規格")
        print("  7. 組裝產能 - 後段組裝線產能")
        print("  8. 產線產能 - 前段生產線產能")
        print("  9. 訂單與預測 - 需求數據")
        print("\n系統將會在以下工作表輸出計算結果：")
        print("  - 生產排程建議")
        print("  - 組裝調整建議")
        print("  - 齊料缺口分析")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n[錯誤] 執行失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    return True


if __name__ == '__main__':
    success = main()
    if not success:
        exit(1)

