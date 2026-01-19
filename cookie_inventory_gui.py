"""
餅乾庫存算料系統 - GUI 啟動腳本
在根目錄提供一個簡單的入口點來啟動 GUI
"""
# 開發環境直接使用 Cookies 目錄名
# 打包時 PyInstaller 會根據 build_exe.spec 中的 hiddenimports 處理模組導入
from Cookies.cookie_inventory_gui import main

if __name__ == '__main__':
    main()
