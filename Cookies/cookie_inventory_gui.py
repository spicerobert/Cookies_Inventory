"""餅乾庫存算料系統 - 圖形化工具
使用 tkinter 和 ttk 建立簡單易用的圖形介面"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
import sys
from datetime import datetime

# 導入功能模組（從套件 __init__.py 導入）
from . import sync_index_from_erp,sync_cookie_inventory,sync_wip_inventory,sync_production_schedule,sync_receipt_data,calculate_cookie_inventory

class TextHandler(logging.Handler):
    """自訂日誌處理器，將日誌輸出到 Text widget"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget        
    def emit(self, record):
        try:
            msg = self.format(record)
            # 使用 after 確保線程安全
            self.text_widget.after(0, self._append_text, msg)
        except Exception:
            pass    
    def _append_text(self, msg):
        """在 Text widget 中追加文字"""
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)

class CookieInventoryGUI:
    """餅乾庫存算料系統主視窗"""    
    def __init__(self, root):
        self.root = root
        self.root.title("餅乾庫存算料系統")
        self.root.geometry("800x600")
        self.root.resizable(True, True)        
        # 執行狀態
        self.is_running = False        
        # 建立介面
        self._create_widgets()        
        # 設定日誌
        self._setup_logging()        
    def _create_widgets(self):
        """建立 GUI 元件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")        
        # 設定 grid 權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)        
        # 標題
        title_label = ttk.Label(main_frame, text="餅乾庫存算料系統",font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))        
        # 功能按鈕框架
        button_frame = ttk.LabelFrame(main_frame, text="功能選單", padding="10")
        button_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))        
        # 按鈕樣式
        button_width = 20        
        # 同步 Index 按鈕
        self.btn_sync_index = ttk.Button(
            button_frame,
            text="1. 同步 Index 資料",
            command=self._sync_index,
            width=button_width
        )
        self.btn_sync_index.grid(row=0, column=0, pady=5, sticky="ew")
        
        # 同步帳上庫存按鈕
        self.btn_sync_inventory = ttk.Button(
            button_frame,
            text="2. 同步帳上庫存",
            command=self._sync_inventory,
            width=button_width
        )
        self.btn_sync_inventory.grid(row=1, column=0, pady=5, sticky="ew")
        
        # 同步在製品按鈕
        self.btn_sync_wip = ttk.Button(
            button_frame,
            text="3. 同步在製品庫存",
            command=self._sync_wip,
            width=button_width
        )
        self.btn_sync_wip.grid(row=2, column=0, pady=5, sticky="ew")
        
        # 同步生產排程按鈕
        self.btn_sync_production = ttk.Button(
            button_frame,
            text="4. 同步生產排程",
            command=self._sync_production,
            width=button_width
        )
        self.btn_sync_production.grid(row=3, column=0, pady=5, sticky="ew")
        
        # 同步完工入庫按鈕
        self.btn_sync_receipt = ttk.Button(
            button_frame,
            text="5. 同步完工入庫",
            command=self._sync_receipt,
            width=button_width
        )
        self.btn_sync_receipt.grid(row=4, column=0, pady=5, sticky="ew")
        
        # 計算庫存按鈕
        self.btn_calculate = ttk.Button(
            button_frame,
            text="6. 計算庫存預估",
            command=self._calculate_inventory,
            width=button_width
        )
        self.btn_calculate.grid(row=5, column=0, pady=5, sticky="ew")
        
        # 清除日誌按鈕
        self.btn_clear = ttk.Button(
            button_frame,
            text="清除日誌",
            command=self._clear_log,
            width=button_width
        )
        self.btn_clear.grid(row=6, column=0, pady=5, sticky="ew")
        
        # 日誌顯示框架
        log_frame = ttk.LabelFrame(main_frame, text="執行日誌", padding="10")
        log_frame.grid(row=1, column=1, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 日誌文字區域（帶滾動條）
        log_text_frame = ttk.Frame(log_frame)
        log_text_frame.grid(row=0, column=0, sticky="nsew")
        log_text_frame.columnconfigure(0, weight=1)
        log_text_frame.rowconfigure(0, weight=1)
        
        # 文字區域
        self.log_text = tk.Text(
            log_text_frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#f5f5f5",
            fg="#333333"
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        # 滾動條
        scrollbar = ttk.Scrollbar(log_text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 狀態列
        self.status_label = ttk.Label(
            main_frame,
            text="就緒",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        
    def _setup_logging(self):
        """設定日誌系統"""
        # 清除現有的 handlers
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)        
        # 設定日誌格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')        
        # 建立 Text widget handler
        text_handler = TextHandler(self.log_text)
        text_handler.setLevel(logging.INFO)
        text_handler.setFormatter(formatter)        
        # 設定根日誌記錄器
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(text_handler)        
        # 記錄啟動訊息
        logger = logging.getLogger(__name__)
        logger.info("=" * 60)
        logger.info("餅乾庫存算料系統已啟動")
        logger.info(f"啟動時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 60)
        
    def _update_status(self, message):
        """更新狀態列"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
        
    def _set_buttons_state(self, enabled):
        """設定按鈕狀態"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.btn_sync_index.config(state=state)
        self.btn_sync_inventory.config(state=state)
        self.btn_sync_wip.config(state=state)
        self.btn_sync_production.config(state=state)
        self.btn_sync_receipt.config(state=state)
        self.btn_calculate.config(state=state)
        self.is_running = not enabled
        
    def _sync_index(self):
        """同步 Index 資料"""
        if self.is_running:
            messagebox.showwarning("警告", "已有任務正在執行中，請稍候...")
            return            
        self._set_buttons_state(False)
        self._update_status("正在同步 Index 資料...")        
        def run():
            try:
                logger = logging.getLogger(__name__)
                logger.info("開始執行：同步 Index 資料")
                success = sync_index_from_erp()
                if success:
                    logger.info("✓ 同步 Index 資料完成")
                    self.root.after(0, lambda: messagebox.showinfo("完成", "同步 Index 資料完成！"))
                else:
                    logger.error("✗ 同步 Index 資料失敗")
                    self.root.after(0, lambda: messagebox.showerror("錯誤", "同步 Index 資料失敗，請查看日誌"))
            except Exception as e:
                logger = logging.getLogger(__name__)
                logger.error(f"執行錯誤: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"執行錯誤: {str(e)}"))
            finally:
                self.root.after(0, self._set_buttons_state, True)
                self.root.after(0, lambda: self._update_status("就緒"))        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        
    def _sync_inventory(self):
        """同步帳上庫存資料"""
        if self.is_running:
            messagebox.showwarning("警告", "已有任務正在執行中，請稍候...")
            return            
        self._set_buttons_state(False)
        self._update_status("正在同步帳上庫存資料...")        
        def run():
            try:
                logger = logging.getLogger(__name__)
                logger.info("開始執行：同步帳上庫存資料")
                success = sync_cookie_inventory()
                if success:
                    logger.info("✓ 同步帳上庫存資料完成")
                    self.root.after(0, lambda: messagebox.showinfo("完成", "同步帳上庫存資料完成！"))
                else:
                    logger.error("✗ 同步帳上庫存資料失敗")
                    self.root.after(0, lambda: messagebox.showerror("錯誤", "同步帳上庫存資料失敗，請查看日誌"))
            except Exception as e:
                logger = logging.getLogger(__name__)
                logger.error(f"執行錯誤: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"執行錯誤: {str(e)}"))
            finally:
                self.root.after(0, self._set_buttons_state, True)
                self.root.after(0, lambda: self._update_status("就緒"))        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        
    def _sync_wip(self):
        """同步在製品庫存"""
        if self.is_running:
            messagebox.showwarning("警告", "已有任務正在執行中，請稍候...")
            return            
        self._set_buttons_state(False)
        self._update_status("正在同步在製品庫存...")        
        def run():
            try:
                logger = logging.getLogger(__name__)
                logger.info("開始執行：同步在製品庫存")
                success = sync_wip_inventory()
                if success:
                    logger.info("✓ 同步在製品庫存完成")
                    self.root.after(0, lambda: messagebox.showinfo("完成", "同步在製品庫存完成！"))
                else:
                    logger.error("✗ 同步在製品庫存失敗")
                    self.root.after(0, lambda: messagebox.showerror("錯誤", "同步在製品庫存失敗，請查看日誌"))
            except Exception as e:
                logger = logging.getLogger(__name__)
                logger.error(f"執行錯誤: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"執行錯誤: {str(e)}"))
            finally:
                self.root.after(0, self._set_buttons_state, True)
                self.root.after(0, lambda: self._update_status("就緒"))        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        
    def _sync_production(self):
        """同步生產排程"""
        if self.is_running:
            messagebox.showwarning("警告", "已有任務正在執行中，請稍候...")
            return            
        self._set_buttons_state(False)
        self._update_status("正在同步生產排程...")        
        def run():
            try:
                logger = logging.getLogger(__name__)
                logger.info("開始執行：同步生產排程")
                success = sync_production_schedule()
                if success:
                    logger.info("✓ 同步生產排程完成")
                    self.root.after(0, lambda: messagebox.showinfo("完成", "同步生產排程完成！"))
                else:
                    logger.error("✗ 同步生產排程失敗")
                    self.root.after(0, lambda: messagebox.showerror("錯誤", "同步生產排程失敗，請查看日誌"))
            except Exception as e:
                logger = logging.getLogger(__name__)
                logger.error(f"執行錯誤: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"執行錯誤: {str(e)}"))
            finally:
                self.root.after(0, self._set_buttons_state, True)
                self.root.after(0, lambda: self._update_status("就緒"))        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        
    def _sync_receipt(self):
        """同步完工入庫資料"""
        if self.is_running:
            messagebox.showwarning("警告", "已有任務正在執行中，請稍候...")
            return            
        self._set_buttons_state(False)
        self._update_status("正在同步完工入庫資料...")        
        def run():
            try:
                logger = logging.getLogger(__name__)
                logger.info("開始執行：同步完工入庫資料")
                success = sync_receipt_data(days_back=5)
                if success:
                    logger.info("✓ 同步完工入庫資料完成")
                    self.root.after(0, lambda: messagebox.showinfo("完成", "同步完工入庫資料完成！"))
                else:
                    logger.error("✗ 同步完工入庫資料失敗")
                    self.root.after(0, lambda: messagebox.showerror("錯誤", "同步完工入庫資料失敗，請查看日誌"))
            except Exception as e:
                logger = logging.getLogger(__name__)
                logger.error(f"執行錯誤: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"執行錯誤: {str(e)}"))
            finally:
                self.root.after(0, self._set_buttons_state, True)
                self.root.after(0, lambda: self._update_status("就緒"))        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        
    def _calculate_inventory(self):
        """計算庫存預估"""
        if self.is_running:
            messagebox.showwarning("警告", "已有任務正在執行中，請稍候...")
            return            
        self._set_buttons_state(False)
        self._update_status("正在計算庫存預估...")        
        def run():
            try:
                logger = logging.getLogger(__name__)
                logger.info("開始執行：計算庫存預估")
                success = calculate_cookie_inventory()
                if success:
                    logger.info("✓ 計算庫存預估完成")
                    self.root.after(0, lambda: messagebox.showinfo("完成", "計算庫存預估完成！"))
                else:
                    logger.error("✗ 計算庫存預估失敗")
                    self.root.after(0, lambda: messagebox.showerror("錯誤", "計算庫存預估失敗，請查看日誌"))
            except Exception as e:
                logger = logging.getLogger(__name__)
                logger.error(f"執行錯誤: {str(e)}")
                import traceback
                logger.error(traceback.format_exc())
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"執行錯誤: {str(e)}"))
            finally:
                self.root.after(0, self._set_buttons_state, True)
                self.root.after(0, lambda: self._update_status("就緒"))        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()        
    def _clear_log(self):
        """清除日誌"""
        self.log_text.delete(1.0, tk.END)
        logger = logging.getLogger(__name__)
        logger.info("日誌已清除")

def main():
    """主函數"""
    root = tk.Tk()
    app = CookieInventoryGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()
