import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from ttkthemes import ThemedTk
import threading
import queue
import time
from datetime import datetime
import pandas as pd

# 导入自定义模块
from pdf_processor import PDFProcessor
from bank_parsers import get_bank_parser
from utils import setup_logging, get_config, save_config
from ui_components import ToolTipButton, ProgressWindow, BankMappingDialog

# 设置日志
logger = setup_logging()

class BankStatementApp(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc")
        
        # 应用程序配置
        self.title("银行账单批量处理工具")
        self.geometry("900x700")
        self.minsize(800, 600)
        self.config = get_config()
        
        # 创建队列用于线程间通信
        self.queue = queue.Queue()
        self.processing_thread = None
        self.is_processing = False
        
        # 创建UI组件
        self.create_menu()
        self.create_main_frame()
        
        # 初始化数据
        self.pdf_files = []
        self.output_file = ""
        self.bank_mapping = {}
        
        # 加载上次的配置
        self.load_last_session()
        
        # 定期检查队列
        self.after(100, self.check_queue)
    
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="选择PDF文件", command=self.select_pdf_files)
        file_menu.add_command(label="选择输出Excel文件", command=self.select_output_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="银行识别映射", command=self.open_bank_mapping)
        settings_menu.add_command(label="高级设置", command=self.open_advanced_settings)
        menubar.add_cascade(label="设置", menu=settings_menu)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.config(menu=menubar)
    
    def create_main_frame(self):
        """创建主界面"""
        # 创建主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 添加按钮
        self.select_pdf_btn = ToolTipButton(btn_frame, text="选择PDF文件", 
                                         command=self.select_pdf_files,
                                         tooltip="选择要处理的银行账单PDF文件")
        self.select_pdf_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.select_output_btn = ToolTipButton(btn_frame, text="选择输出文件", 
                                            command=self.select_output_file,
                                            tooltip="选择输出的Excel文件路径")
        self.select_output_btn.pack(side=tk.LEFT, padx=5)
        
        self.process_btn = ToolTipButton(btn_frame, text="开始处理", 
                                       command=self.start_processing,
                                       tooltip="开始处理所选PDF文件")
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ToolTipButton(btn_frame, text="停止处理", 
                                    command=self.stop_processing,
                                    tooltip="停止当前处理任务",
                                    state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表区域
        files_frame = ttk.LabelFrame(main_frame, text="PDF文件列表")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建文件列表
        self.file_tree = ttk.Treeview(files_frame, columns=("文件名", "银行", "状态"), show="headings")
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("银行", text="银行")
        self.file_tree.heading("状态", text="状态")
        
        self.file_tree.column("文件名", width=400)
        self.file_tree.column("银行", width=150)
        self.file_tree.column("状态", width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 右键菜单
        self.context_menu = tk.Menu(self.file_tree, tearoff=0)
        self.context_menu.add_command(label="移除", command=self.remove_selected_file)
        self.context_menu.add_command(label="查看内容", command=self.preview_pdf)
        self.file_tree.bind("<Button-3>", self.show_context_menu)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(status_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT)
        
        self.output_label = ttk.Label(status_frame, text="输出文件: 未选择")
        self.output_label.pack(side=tk.RIGHT)
    
    def select_pdf_files(self):
        """选择PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择银行账单PDF文件",
            filetypes=[("PDF文件", "*.pdf")],
            initialdir=self.config.get("last_pdf_dir", os.path.expanduser("~"))
        )
        
        if files:
            # 保存最后使用的目录
            self.config["last_pdf_dir"] = os.path.dirname(files[0])
            save_config(self.config)
            
            # 添加文件到列表
            for file_path in files:
                if file_path not in [f["path"] for f in self.pdf_files]:
                    # 尝试自动识别银行
                    bank_name = self.auto_detect_bank(file_path)
                    
                    file_info = {
                        "path": file_path,
                        "name": os.path.basename(file_path),
                        "bank": bank_name,
                        "status": "待处理"
                    }
                    
                    self.pdf_files.append(file_info)
                    self.file_tree.insert("", tk.END, values=(file_info["name"], file_info["bank"], file_info["status"]))
            
            self.log(f"已添加 {len(files)} 个PDF文件")
    
    def auto_detect_bank(self, file_path):
        """尝试自动检测银行类型"""
        try:
            processor = PDFProcessor()
            return processor.detect_bank_type(file_path)
        except Exception as e:
            logger.error(f"自动检测银行失败: {str(e)}")
            return "未知"
    
    def select_output_file(self):
        """选择输出Excel文件"""
        file_path = filedialog.asksaveasfilename(
            title="选择输出Excel文件",
            filetypes=[("Excel文件", "*.xlsx")],
            defaultextension=".xlsx",
            initialdir=self.config.get("last_output_dir", os.path.expanduser("~")),
            initialfile=f"银行账单汇总_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if file_path:
            self.output_file = file_path
            self.config["last_output_dir"] = os.path.dirname(file_path)
            save_config(self.config)
            
            self.output_label.config(text=f"输出文件: {os.path.basename(file_path)}")
            self.log(f"已选择输出文件: {file_path}")
    
    def start_processing(self):
        """开始处理PDF文件"""
        if not self.pdf_files:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
        
        if not self.output_file:
            messagebox.showwarning("警告", "请先选择输出Excel文件")
            return
        
        # 检查是否有未知银行
        unknown_banks = [f for f in self.pdf_files if f["bank"] == "未知"]
        if unknown_banks:
            if not messagebox.askyesno("确认", f"有 {len(unknown_banks)} 个文件的银行类型未识别，是否继续处理？"):
                return
        
        # 禁用按钮
        self.process_btn.config(state=tk.DISABLED)
        self.select_pdf_btn.config(state=tk.DISABLED)
        self.select_output_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 更新状态
        self.is_processing = True
        self.status_label.config(text="处理中...")
        
        # 创建进度窗口
        self.progress_window = ProgressWindow(self, len(self.pdf_files))
        
        # 启动处理线程
        self.processing_thread = threading.Thread(target=self.process_files)
        self.processing_thread.daemon = True
        self.processing_thread.start()
    
    def process_files(self):
        """在单独的线程中处理文件"""
        try:
            # 初始化结果DataFrame
            all_results = []
            
            for i, file_info in enumerate(self.pdf_files):
                if not self.is_processing:
                    # 如果处理被停止
                    self.queue.put(("log", "处理已停止"))
                    break
                
                try:
                    # 更新状态
                    self.queue.put(("update_status", (i, file_info["name"], "处理中")))
                    
                    # 获取对应的银行解析器
                    bank_parser = get_bank_parser(file_info["bank"])
                    
                    if bank_parser is None:
                        self.queue.put(("log", f"错误: 不支持的银行类型 '{file_info['bank']}'，跳过文件 {file_info['name']}"))
                        self.queue.put(("update_status", (i, file_info["name"], "失败")))
                        continue
                    
                    # 处理PDF文件
                    processor = PDFProcessor()
                    transactions = processor.process_pdf(file_info["path"], bank_parser)
                    
                    if transactions:
                        # 添加银行名称列
                        for trans in transactions:
                            trans["银行"] = file_info["bank"]
                            trans["文件名"] = file_info["name"]
                        
                        all_results.extend(transactions)
                        self.queue.put(("log", f"成功处理 {file_info['name']}，提取了 {len(transactions)} 条交易记录"))
                        self.queue.put(("update_status", (i, file_info["name"], "成功")))
                    else:
                        self.queue.put(("log", f"警告: 未能从 {file_info['name']} 提取任何交易记录"))
                        self.queue.put(("update_status", (i, file_info["name"], "无数据")))
                    
                except Exception as e:
                    logger.exception(f"处理文件 {file_info['name']} 时出错")
                    self.queue.put(("log", f"错误: 处理 {file_info['name']} 失败 - {str(e)}"))
                    self.queue.put(("update_status", (i, file_info["name"], "失败")))
                
                # 更新进度
                self.queue.put(("update_progress", i + 1))
            
            # 保存结果到Excel
            if all_results and self.is_processing:
                df = pd.DataFrame(all_results)
                
                # 标准化列名和顺序
                standard_columns = [
                    "银行", "文件名", "交易日期", "交易时间", "交易类型", "交易金额", 
                    "收入金额", "支出金额", "账户余额", "交易描述", "对方账号", "对方户名", 
                    "交易渠道", "交易地点", "备注"
                ]
                
                # 确保所有标准列都存在
                for col in standard_columns:
                    if col not in df.columns:
                        df[col] = ""
                
                # 重新排序列
                existing_cols = [col for col in standard_columns if col in df.columns]
                other_cols = [col for col in df.columns if col not in standard_columns]
                df = df[existing_cols + other_cols]
                
                # 保存到Excel
                df.to_excel(self.output_file, index=False, sheet_name="交易记录")
                
                # 创建汇总表
                self.create_summary_sheet(self.output_file, df)
                
                self.queue.put(("log", f"已成功将 {len(all_results)} 条交易记录保存到 {self.output_file}"))
            elif not all_results:
                self.queue.put(("log", "未提取到任何交易记录，未生成Excel文件"))
        
        except Exception as e:
            logger.exception("处理过程中发生错误")
            self.queue.put(("log", f"处理过程中发生错误: {str(e)}"))
        
        finally:
            # 处理完成
            self.queue.put(("processing_done", None))
    
    def create_summary_sheet(self, excel_file, df):
        """创建汇总表"""
        try:
            # 读取已保存的Excel文件
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
                # 按银行汇总
                bank_summary = df.groupby('银行').agg({
                    '收入金额': 'sum',
                    '支出金额': 'sum',
                    '交易日期': ['min', 'max'],
                    '文件名': 'nunique'
                }).reset_index()
                
                bank_summary.columns = ['银行', '总收入', '总支出', '起始日期', '结束日期', '文件数']
                bank_summary['净收入'] = bank_summary['总收入'] - bank_summary['总支出']
                
                # 保存汇总表
                bank_summary.to_excel(writer, sheet_name='银行汇总', index=False)
                
                # 按月汇总
                if '交易日期' in df.columns:
                    # 确保交易日期是日期类型
                    df['交易日期'] = pd.to_datetime(df['交易日期'], errors='coerce')
                    df['月份'] = df['交易日期'].dt.strftime('%Y-%m')
                    
                    month_summary = df.groupby('月份').agg({
                        '收入金额': 'sum',
                        '支出金额': 'sum',
                        '交易日期': 'count'
                    }).reset_index()
                    
                    month_summary.columns = ['月份', '总收入', '总支出', '交易笔数']
                    month_summary['净收入'] = month_summary['总收入'] - month_summary['总支出']
                    
                    # 按月份排序
                    month_summary = month_summary.sort_values('月份')
                    
                    # 保存月度汇总
                    month_summary.to_excel(writer, sheet_name='月度汇总', index=False)
            
            self.queue.put(("log", "已创建汇总报表"))
        
        except Exception as e:
            logger.exception("创建汇总表时出错")
            self.queue.put(("log", f"创建汇总表时出错: {str(e)}"))
    
    def stop_processing(self):
        """停止处理"""
        if self.is_processing:
            self.is_processing = False
            self.log("正在停止处理...")
    
    def check_queue(self):
        """检查队列中的消息"""
        try:
            while True:
                message_type, message = self.queue.get_nowait()
                
                if message_type == "log":
                    self.log(message)
                elif message_type == "update_status":
                    idx, filename, status = message
                    self.update_file_status(idx, status)
                elif message_type == "update_progress":
                    if hasattr(self, 'progress_window'):
                        self.progress_window.update_progress(message)
                elif message_type == "processing_done":
                    self.processing_done()
                
                self.queue.task_done()
        
        except queue.Empty:
            pass
        
        # 继续检查队列
        self.after(100, self.check_queue)
    
    def update_file_status(self, idx, status):
        """更新文件状态"""
        if 0 <= idx < len(self.pdf_files):
            self.pdf_files[idx]["status"] = status
            
            # 更新树形视图
            item_id = self.file_tree.get_children()[idx]
            values = self.file_tree.item(item_id, "values")
            self.file_tree.item(item_id, values=(values[0], values[1], status))
    
    def processing_done(self):
        """处理完成后的操作"""
        # 启用按钮
        self.process_btn.config(state=tk.NORMAL)
        self.select_pdf_btn.config(state=tk.NORMAL)
        self.select_output_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        # 更新状态
        self.is_processing = False
        self.status_label.config(text="就绪")
        
        # 关闭进度窗口
        if hasattr(self, 'progress_window'):
            self.progress_window.destroy()
        
        # 显示完成消息
        if os.path.exists(self.output_file):
            if messagebox.askyesno("处理完成", "所有文件处理完成，是否打开生成的Excel文件？"):
                self.open_excel_file()
        else:
            messagebox.showinfo("处理完成", "处理已完成，但未生成Excel文件")
    
    def open_excel_file(self):
        """打开生成的Excel文件"""
        if os.path.exists(self.output_file):
            os.startfile(self.output_file)
    
    def log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # 同时写入日志文件
        logger.info(message)
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        # 获取点击的项目
        item = self.file_tree.identify_row(event.y)
        if item:
            # 选中该项目
            self.file_tree.selection_set(item)
            # 显示菜单
            self.context_menu.post(event.x_root, event.y_root)
    
    def remove_selected_file(self):
        """移除选中的文件"""
        selected = self.file_tree.selection()
        if selected:
            for item in selected:
                idx = self.file_tree.index(item)
                if 0 <= idx < len(self.pdf_files):
                    del self.pdf_files[idx]
                    self.file_tree.delete(item)
            
            self.log(f"已移除 {len(selected)} 个文件")
    
    def preview_pdf(self):
        """预览PDF内容"""
        selected = self.file_tree.selection()
        if selected:
            item = selected[0]
            idx = self.file_tree.index(item)
            if 0 <= idx < len(self.pdf_files):
                file_path = self.pdf_files[idx]["path"]
                try:
                    # 使用系统默认PDF查看器打开文件
                    os.startfile(file_path)
                except Exception as e:
                    messagebox.showerror("错误", f"无法打开PDF文件: {str(e)}")
    
    def open_bank_mapping(self):
        """打开银行映射设置"""
        dialog = BankMappingDialog(self, self.bank_mapping)
        self.wait_window(dialog)
        if dialog.result:
            self.bank_mapping = dialog.result
            # 保存到配置
            self.config["bank_mapping"] = self.bank_mapping
            save_config(self.config)
    
    def open_advanced_settings(self):
        """打开高级设置"""
        # TODO: 实现高级设置对话框
        messagebox.showinfo("提示", "高级设置功能尚未实现")
    
    def show_help(self):
        """显示帮助信息"""
        help_text = """
使用说明：

1. 选择PDF文件：点击"选择PDF文件"按钮，选择需要处理的银行账单PDF文件。
2. 选择输出文件：点击"选择输出文件"按钮，选择保存结果的Excel文件路径。
3. 开始处理：点击"开始处理"按钮，程序将自动识别银行类型并提取交易记录。
4. 查看结果：处理完成后，可以选择打开生成的Excel文件查看结果。

支持的银行类型：
- 工商银行
- 建设银行
- 农业银行
- 中国银行
- 交通银行
- 招商银行
- 浦发银行
- 民生银行
- 中信银行
- 光大银行
- 华夏银行
- 广发银行
- 平安银行
- 邮储银行

如需添加其他银行支持，请在"设置"菜单中配置银行识别映射。
        """
        
        help_window = tk.Toplevel(self)
        help_window.title("使用说明")
        help_window.geometry("600x400")
        help_window.transient(self)
        help_window.grab_set()
        
        text = ScrolledText(help_window, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.insert(tk.END, help_text)
        text.config(state=tk.DISABLED)
    
    def show_about(self):
        """显示关于信息"""
        about_text = """
银行账单批量处理工具

版本: 1.0.0

功能：
- 支持多家银行账单PDF文件批量处理
- 自动识别银行类型
- 提取交易记录并导出到Excel
- 生成汇总报表

作者: Trea AI
        """
        
        messagebox.showinfo("关于", about_text)
    
    def load_last_session(self):
        """加载上次会话的配置"""
        # 加载银行映射
        self.bank_mapping = self.config.get("bank_mapping", {})
        
        # 加载上次的输出文件
        last_output = self.config.get("last_output_file", "")
        if last_output and os.path.dirname(last_output):
            self.output_file = last_output
            self.output_label.config(text=f"输出文件: {os.path.basename(last_output)}")
    
    def on_closing(self):
        """关闭应用程序时的操作"""
        # 保存当前配置
        if self.output_file:
            self.config["last_output_file"] = self.output_file
        
        save_config(self.config)
        
        # 关闭应用程序
        self.destroy()

if __name__ == "__main__":
    # 检查是否是PyInstaller打包的可执行文件
    if getattr(sys, 'frozen', False):
        # 如果是打包的可执行文件，设置工作目录为可执行文件所在目录
        os.chdir(os.path.dirname(sys.executable))
    
    app = BankStatementApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()