"""
測試 Index 工作表功能
"""
from google_sheets_helper import GoogleSheetsHelper


def test_index_reading():
    """測試 Index 工作表的讀取功能"""
    print("=" * 60)
    print("測試 Index 工作表功能")
    print("=" * 60)
    
    try:
        # 建立連接
        helper = GoogleSheetsHelper()
        print(f"\n[成功] 已連接到試算表: {helper.spreadsheet.title}")
        
        # 讀取 Index 資料
        print("\n[測試] 讀取 Index 工作表...")
        index_dict = helper.get_index_dict()
        
        print("\nIndex 資料內容:")
        print(f"  餅乾對應表: {len(index_dict['餅乾'])} 筆")
        if index_dict['餅乾']:
            print("  範例:")
            for code, name in list(index_dict['餅乾'].items())[:3]:
                print(f"    {code} -> {name}")
        
        print(f"\n  禮盒對應表: {len(index_dict['禮盒'])} 筆")
        if index_dict['禮盒']:
            print("  範例:")
            for code, name in list(index_dict['禮盒'].items())[:3]:
                print(f"    {code} -> {name}")
        
        print(f"\n  產線對應表: {len(index_dict['產線'])} 筆")
        if index_dict['產線']:
            print("  範例:")
            for code, name in list(index_dict['產線'].items())[:3]:
                print(f"    {code} -> {name}")
        
        # 測試單一查詢
        if index_dict['餅乾']:
            test_code = list(index_dict['餅乾'].keys())[0]
            print(f"\n[測試] 查詢代號 '{test_code}' 的名稱...")
            name = helper.get_name_by_code(test_code, '餅乾')
            print(f"  結果: {name}")
        
        print("\n" + "=" * 60)
        print("[完成] Index 工作表功能測試完成！")
        print("=" * 60)
        print("\n提示: 如果 Index 工作表是空的，請先在 Google Sheets 中填入資料")
        
        return True
        
    except Exception as e:
        print(f"\n[錯誤] 測試失敗: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    success = test_index_reading()
    if not success:
        exit(1)


