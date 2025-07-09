import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
import queue
import time
from datetime import datetime
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import pdfplumber


# 导入自定义模块
from pdf_processor import PDFProcessor
from bank_parsers import get_bank_parser
from utils import setup_logging, get_config, save_config
from ui_components import ToolTipButton, ProgressWindow, BankMappingDialog, AdvancedSettingsDialog
from exceptions import PDFProcessingError, BankDetectionError

# 设置日志
logger = setup_logging()

class SplashScreen(tk.Toplevel):
    """启动画面"""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("")
        self.geometry("400x250")
        self.overrideredirect(True)  # 无边框窗口
        self.configure(background="#ffffff")
        
        # 居中显示
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # 添加内容
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(frame, text="银行账单批量处理工具", 
                font=("Microsoft YaHei UI", 16, "bold"),
                foreground="#4a86e8").pack(pady=(20, 10))
        
        # 副标题
        ttk.Label(frame, text="正在加载应用...", 
                font=("Microsoft YaHei UI", 10),
                foreground="#666666").pack(pady=(0, 20))
        
        # 进度条
        self.progress = ttk.Progressbar(frame, orient="horizontal", 
                                      length=350, mode="indeterminate")
        self.progress.pack(pady=10)
        self.progress.start(15)
        
        # 版本信息
        ttk.Label(frame, text="版本 1.0", 
                font=("Microsoft YaHei UI", 8),
                foreground="#999999").pack(side=tk.BOTTOM, pady=10)

class BankStatementApp(ttk.Window):
    def __init__(self):
        # 使用ttkbootstrap的主题
        super().__init__(themename="litera")  # 可选主题：cosmo, flatly, litera, minty, lumen, sandstone, yeti等
        
        # 设置自定义颜色方案 - 可以使用ttkbootstrap的颜色常量
        self.primary_color = PRIMARY  # 使用ttkbootstrap常量
        self.secondary_color = SECONDARY
        self.accent_color = INFO
        self.success_color = SUCCESS
        self.warning_color = WARNING
        self.error_color = DANGER
        
        # 应用程序配置
        self.title("银行账单批量处理工具")
        self.geometry("950x700")
        self.minsize(850, 650)
        self.app_config = get_config()
        
        # 设置样式
        self.setup_styles()
        
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
    
    def setup_styles(self):
        """设置自定义样式 - 使用ttkbootstrap风格"""
        self.custom_style = ttk.Style()
        
        # 设置整体背景色 - ttkbootstrap已经处理了大部分样式
        # 设置标题标签样式
        self.custom_style.configure("Title.TLabel", 
                            font=("Microsoft YaHei UI", 12, "bold"), 
                            foreground=self.primary_color)
        
        # 设置Treeview样式
        self.custom_style.configure("Treeview", 
                            font=("Microsoft YaHei UI", 9))
        self.custom_style.configure("Treeview.Heading", 
                            font=("Microsoft YaHei UI", 9, "bold"))
    
    def batch_detect_bank_type(self):
        """批量重新检测银行类型"""
        if not self.pdf_files:
            messagebox.showwarning("警告", "没有文件可以检测")
            return
        
        processor = PDFProcessor()
        
        # 更新所有文件的银行类型
        for i, file_info in enumerate(self.pdf_files):
            try:
                bank_name = processor.detect_bank_type(file_info["path"], self.bank_mapping)
                file_info["bank"] = bank_name
                
                # 更新UI
                items = self.file_tree.get_children()
                if i < len(items):
                    item = items[i]
                    values = list(self.file_tree.item(item, "values"))
                    values[1] = bank_name
                    self.file_tree.item(item, values=values)
                
                self.log(f"重新检测文件: {file_info['name']}，银行类型: {bank_name}")
            except Exception as e:
                self.log(f"检测文件 {file_info['name']} 银行类型时出错: {str(e)}")

    def redetect_bank_type(self):
        """重新检测选中文件的银行类型"""
        selected = self.file_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要重新检测的文件")
            return
        
        processor = PDFProcessor()
        
        # 获取选中项的索引
        items = self.file_tree.get_children()
        for item in selected:
            index = items.index(item)
            
            if index < len(self.pdf_files):
                file_info = self.pdf_files[index]
                try:
                    bank_name = processor.detect_bank_type(file_info["path"], self.bank_mapping)
                    file_info["bank"] = bank_name
                    
                    # 更新UI
                    values = list(self.file_tree.item(item, "values"))
                    values[1] = bank_name
                    self.file_tree.item(item, values=values)
                    
                    self.log(f"重新检测文件: {file_info['name']}，银行类型: {bank_name}")
                except Exception as e:
                    self.log(f"检测文件 {file_info['name']} 银行类型时出错: {str(e)}")
    
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
        """创建主界面 - 使用ttkbootstrap样式"""
        # 创建主框架
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(title_frame, text="银行账单批量处理", 
                 font=("Microsoft YaHei UI", 12, "bold"),
                 bootstyle="primary").pack(side=tk.LEFT)  # 使用bootstyle
        
        # 顶部按钮区域 - 使用更好的分组和间距
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 文件操作按钮组
        file_btn_frame = ttk.Labelframe(btn_frame, text="文件操作", bootstyle="primary")  # 使用bootstyle
        file_btn_frame.pack(side=tk.LEFT, padx=(0, 10), fill=tk.Y)
        
        # 添加按钮 - 使用ttkbootstrap样式
        self.select_pdf_btn = ToolTipButton(file_btn_frame, text="选择PDF文件", 
                                         command=self.select_pdf_files,
                                         tooltip="选择要处理的银行账单PDF文件",
                                         bootstyle="primary")  # 使用bootstyle替代style
        self.select_pdf_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.select_folder_btn = ToolTipButton(file_btn_frame, text="选择文件夹", 
                                           command=self.select_pdf_folder,
                                           tooltip="选择包含银行账单PDF文件的文件夹",
                                           bootstyle="primary")
        self.select_folder_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.select_output_btn = ToolTipButton(file_btn_frame, text="选择输出文件", 
                                            command=self.select_output_file,
                                            tooltip="选择输出的Excel文件路径",
                                            bootstyle="primary")
        self.select_output_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 处理操作按钮组
        process_btn_frame = ttk.Labelframe(btn_frame, text="处理操作", bootstyle="primary")
        process_btn_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)
        
        self.process_btn = ToolTipButton(process_btn_frame, text="开始处理", 
                                       command=self.start_processing,
                                       tooltip="开始处理所选PDF文件",
                                       bootstyle="success")  # 使用success样式
        self.process_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.stop_btn = ToolTipButton(process_btn_frame, text="停止处理", 
                                    command=self.stop_processing,
                                    tooltip="停止当前处理任务",
                                    state=tk.DISABLED,
                                    bootstyle="danger")  # 使用danger样式
        self.stop_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 批量操作按钮组
        batch_btn_frame = ttk.Labelframe(btn_frame, text="批量操作", bootstyle="primary")
        batch_btn_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)
        
        self.batch_detect_btn = ToolTipButton(batch_btn_frame, text="重新检测银行", 
                                           command=self.batch_detect_bank_type,
                                           tooltip="重新检测所有文件的银行类型",
                                           bootstyle="secondary")
        self.batch_detect_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 文件列表区域
        files_frame = ttk.Labelframe(main_frame, text="PDF文件列表", bootstyle="primary")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 创建文件列表 - 添加条纹效果
        self.file_tree = ttk.Treeview(files_frame, columns=("文件名", "银行", "状态", "文件大小"), show="headings", style="Treeview")
        
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("银行", text="银行")
        self.file_tree.heading("状态", text="状态")
        self.file_tree.heading("文件大小", text="文件大小")
        
        self.file_tree.column("文件名", width=350)
        self.file_tree.column("银行", width=150)
        self.file_tree.column("状态", width=100)
        self.file_tree.column("文件大小", width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 右键菜单
        self.context_menu = tk.Menu(self.file_tree, tearoff=0)
        self.context_menu.add_command(label="移除", command=self.remove_selected_file)
        self.context_menu.add_command(label="查看内容", command=self.preview_pdf)
        self.context_menu.add_command(label="重新检测银行", command=self.redetect_bank_type)
        self.file_tree.bind("<Button-3>", self.show_context_menu)
        
        # 日志区域 - 改进样式
        log_frame = ttk.LabelFrame(main_frame, text="处理日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = ScrolledText(log_frame, height=10, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED, background="#fafafa", foreground="#333333")
        
        # 设置日志文本标签
        # 标签配置移至setup_styles方法中
        
        # 状态栏 - 添加边框和背景色
        status_frame = ttk.Frame(main_frame, relief="sunken", borderwidth=1)
        status_frame.pack(fill=tk.X, pady=(15, 0))
        
        self.status_label = ttk.Label(status_frame, text="就绪", padding=(5, 2))
        self.status_label.pack(side=tk.LEFT)
        
        self.output_label = ttk.Label(status_frame, text="输出文件: 未选择", padding=(5, 2))
        self.output_label.pack(side=tk.RIGHT)
    
    def select_pdf_files(self):
        """选择PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择银行账单PDF文件",
            filetypes=[("PDF文件", "*.pdf")],
            initialdir=self.app_config.get("last_pdf_dir", os.path.expanduser("~"))
        )
        
        if files:
            # 保存最后的目录
            self.app_config["last_pdf_dir"] = os.path.dirname(files[0])
            save_config(self.app_config)
            
            # 清空现有文件列表
            for item in self.file_tree.get_children():
                self.file_tree.delete(item)
            
            self.pdf_files = []
            
            # 添加新文件
            for file_path in files:
                file_name = os.path.basename(file_path)
                file_size = os.path.getsize(file_path)
                file_size_str = self.format_file_size(file_size)
                
                # 检测银行类型
                try:
                    processor = PDFProcessor()
                    bank_name = processor.detect_bank_type(file_path, self.bank_mapping)
                    
                    # 添加到列表
                    item_id = self.file_tree.insert("", tk.END, values=(file_name, bank_name, "待处理", file_size_str))
                    
                    # 设置行样式和状态标签
                    row_index = len(self.pdf_files)
                    row_tags = ["even_row" if row_index % 2 == 0 else "odd_row", "pending"]
                    self.file_tree.item(item_id, tags=row_tags)
                    
                    # 添加到文件列表
                    self.pdf_files.append({
                        "name": file_name,
                        "path": file_path,
                        "bank": bank_name,
                        "size": file_size
                    })
                    
                    self.log(f"添加文件: {file_name}，检测到银行类型: {bank_name}，大小: {file_size_str}")
                    
                except BankDetectionError as e:
                    # 银行类型检测失败
                    item_id = self.file_tree.insert("", tk.END, values=(file_name, "未知", "待处理", file_size_str))
                    
                    # 设置行样式和状态标签
                    row_index = len(self.pdf_files)
                    row_tags = ["even_row" if row_index % 2 == 0 else "odd_row", "unknown_bank"]
                    self.file_tree.item(item_id, tags=row_tags)
                    
                    # 添加到文件列表
                    self.pdf_files.append({
                        "name": file_name,
                        "path": file_path,
                        "bank": "未知",
                        "size": file_size
                    })
                    
                    self.log(f"添加文件: {file_name}，无法检测银行类型: {str(e)}，大小: {file_size_str}", "WARNING")
                
                except Exception as e:
                    # 其他错误
                    self.log(f"添加文件 {file_name} 时出错: {str(e)}", "ERROR")
            
            # 更新状态
            self.status_label.config(text=f"已添加 {len(self.pdf_files)} 个文件")
    
    def select_pdf_folder(self):
        """选择包含PDF文件的文件夹"""
        folder_path = filedialog.askdirectory(
            title="选择包含银行账单PDF文件的文件夹",
            initialdir=self.app_config.get("last_pdf_dir", os.path.expanduser("~"))
        )
        
        if not folder_path:
            return
            
        # 保存最后的目录
        self.app_config["last_pdf_dir"] = folder_path
        save_config(self.app_config)
        
        # 询问是否清空现有文件列表
        clear_existing = True
        if self.pdf_files:
            clear_existing = messagebox.askyesno("确认", "是否清空当前文件列表？")
        
        # 清空现有文件列表
        if clear_existing:
            for item in self.file_tree.get_children():
                self.file_tree.delete(item)
            self.pdf_files = []
        
        # 递归查找所有PDF文件
        self.log(f"正在扫描文件夹: {folder_path}")
        pdf_files = self.find_pdf_files(folder_path)
        
        if not pdf_files:
            messagebox.showinfo("提示", "所选文件夹中未找到PDF文件")
            return
        
        # 添加找到的PDF文件
        processor = PDFProcessor()
        added_count = 0
        
        for file_path in pdf_files:
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            file_size_str = self.format_file_size(file_size)
            
            # 检测银行类型
            try:
                bank_name = processor.detect_bank_type(file_path, self.bank_mapping)
                
                # 添加到列表
                self.file_tree.insert("", tk.END, values=(file_name, bank_name, "待处理", file_size_str))
                
                # 添加到文件列表
                self.pdf_files.append({
                    "name": file_name,
                    "path": file_path,
                    "bank": bank_name,
                    "size": file_size
                })
                
                added_count += 1
                self.log(f"添加文件: {file_name}，检测到银行类型: {bank_name}，大小: {file_size_str}")
                
            except Exception as e:
                # 添加到列表，但标记为未知银行
                self.file_tree.insert("", tk.END, values=(file_name, "未知", "待处理", file_size_str))
                
                # 添加到文件列表
                self.pdf_files.append({
                    "name": file_name,
                    "path": file_path,
                    "bank": "未知",
                    "size": file_size
                })
                
                added_count += 1
                self.log(f"添加文件: {file_name}，无法检测银行类型: {str(e)}，大小: {file_size_str}")
        
        # 更新状态
        self.status_label.config(text=f"已添加 {len(self.pdf_files)} 个文件")
        messagebox.showinfo("完成", f"已从文件夹添加 {added_count} 个PDF文件")
    
    def find_pdf_files(self, folder_path):
        """递归查找文件夹中的所有PDF文件"""
        pdf_files = []
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(".pdf"):
                    pdf_files.append(os.path.join(root, file))
        
        return pdf_files
        
    def format_file_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes/1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes/(1024*1024):.1f} MB"
        else:
            return f"{size_bytes/(1024*1024*1024):.1f} GB"
    
    def select_output_file(self):
        """选择输出文件"""
        # 添加文件类型选择
        file_types = [
            ("Excel文件", "*.xlsx"),
            ("CSV文件", "*.csv")
        ]
        
        file_path = filedialog.asksaveasfilename(
            title="选择输出文件",
            filetypes=file_types,
            defaultextension=".xlsx",
            initialdir=self.app_config.get("last_output_dir", os.path.expanduser("~")),
            initialfile=self.app_config.get("last_output_file", "银行账单汇总.xlsx")
        )
        
        if file_path:
            self.output_file = file_path
            self.output_label.config(text=f"输出文件: {os.path.basename(file_path)}")
            
            # 保存最后的目录和文件名
            self.app_config["last_output_dir"] = os.path.dirname(file_path)
            self.app_config["last_output_file"] = file_path
            save_config(self.app_config)
            
            self.log(f"设置输出文件: {file_path}")
    
    def start_processing(self):
        """开始处理PDF文件"""
        # 检查是否有未知银行
        unknown_banks = [f for f in self.pdf_files if f["bank"] == "未知"]
        if unknown_banks:
            if not messagebox.askyesno("警告", f"有 {len(unknown_banks)} 个文件的银行类型未知，是否继续处理？"):
                return
        
        # 检查是否有文件和输出路径
        if not self.pdf_files:
            messagebox.showwarning("警告", "请先选择要处理的PDF文件")
            return
        
        if not self.output_file:
            messagebox.showwarning("警告", "请先选择输出Excel文件路径")
            return
        
        # 禁用按钮
        self.select_pdf_btn.config(state=tk.DISABLED)
        self.select_output_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 更新状态
        self.status_label.config(text="处理中...")
        self.is_processing = True
        
        # 创建进度窗口
        self.progress_window = ProgressWindow(self, len(self.pdf_files))
        
        # 启动处理线程
        self.processing_thread = threading.Thread(target=self.process_files_thread)
        self.processing_thread.daemon = True
        self.processing_thread.start()
        
        self.log("开始处理文件...")
    
    def process_files_thread(self):
        """在后台线程中处理文件"""
        try:
            def process_single_file(file_info):
                """处理单个文件的内部函数"""
                try:
                    if not self.is_processing:
                        return None
                    
                    # 获取对应的银行解析器
                    bank_parser = get_bank_parser(file_info["bank"])
                    
                    if bank_parser is None:
                        self.queue.put(("log", f"错误: 不支持的银行类型 '{file_info['bank']}'，跳过文件 {file_info['name']}"))
                        self.queue.put(("update_status", (file_info["index"], file_info["name"], "失败")))
                        return None
                    
                    # 处理PDF文件 - 直接使用银行解析器
                    start_time = time.time()
                    with pdfplumber.open(file_info["path"]) as pdf:
                        transactions = bank_parser.parse(pdf)
                    processing_time = time.time() - start_time
                    
                    if transactions:
                        # 添加银行名称和文件名信息
                        for trans in transactions:
                            trans["银行"] = file_info["bank"]
                            trans["文件名"] = file_info["name"]
                        
                        # 生成输出文件路径
                        base_name = os.path.splitext(self.output_file)[0]
                        file_ext = os.path.splitext(self.output_file)[1]
                        output_path = f"{base_name}_{file_info['bank']}_{os.path.splitext(file_info['name'])[0]}{file_ext}"
                        
                        # 让解析器直接保存文件
                        bank_parser.save_to_excel(transactions, output_path, {
                            'bank_name': file_info['bank'],
                            'file_name': file_info['name']
                        })
                        
                        self.queue.put(("log", f"成功处理 {file_info['name']}，提取了 {len(transactions)} 条交易记录，耗时 {processing_time:.2f} 秒"))
                        self.queue.put(("log", f"文件已保存到: {output_path}"))
                        self.queue.put(("update_status", (file_info["index"], file_info["name"], "成功")))
                        return {
                            'file': file_info['name'],
                            'bank': file_info['bank'],
                            'transactions': len(transactions),
                            'output_path': output_path
                        }
                    else:
                        self.queue.put(("log", f"警告: 未能从 {file_info['name']} 提取任何交易记录，耗时 {processing_time:.2f} 秒"))
                        self.queue.put(("update_status", (file_info["index"], file_info["name"], "无数据")))
                        return None
                        
                except Exception as e:
                    logger.exception(f"处理文件 {file_info['name']} 时出错")
                    self.queue.put(("log", f"错误: 处理 {file_info['name']} 失败 - {str(e)}"))
                    self.queue.put(("update_status", (file_info["index"], file_info["name"], "失败")))
                    return None
            
            # 为每个文件添加索引
            for i, file_info in enumerate(self.pdf_files):
                file_info["index"] = i
            
            # 计算合适的线程数
            max_workers = min(os.cpu_count() or 4, len(self.pdf_files))
            self.queue.put(("log", f"使用 {max_workers} 个线程并行处理文件"))
            
            # 使用线程池并行处理文件
            processed_files = []
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(process_single_file, file_info): file_info for file_info in self.pdf_files}
                
                # 收集结果
                for future in futures:
                    result = future.result()
                    if result:
                        processed_files.append(result)
        
            # 显示处理结果汇总
            if processed_files and self.is_processing:
                total_transactions = sum(f['transactions'] for f in processed_files)
                self.queue.put(("log", f"处理完成！共处理 {len(processed_files)} 个文件，提取了 {total_transactions} 条交易记录"))
                self.queue.put(("log", "生成的文件列表:"))
                for f in processed_files:
                    self.queue.put(("log", f"  - {f['bank']}: {f['file']} -> {f['output_path']}"))
            elif not processed_files and self.is_processing:
                self.queue.put(("log", "警告: 未能从任何文件中提取交易记录"))
        
        except Exception as e:
            logger.exception("处理过程中发生错误")
            self.queue.put(("log", f"处理过程中发生错误: {str(e)}"))
            self.queue.put(("processing_done", None))
        
        finally:
            # 处理完成
            self.queue.put(("processing_done", None))
    
    def stop_processing(self):
        """停止处理"""
        if self.is_processing:
            self.is_processing = False
            self.log("正在停止处理...")
            self.status_label.config(text="正在停止...")
    
    def check_queue(self):
        """检查队列消息"""
        try:
            while True:
                message = self.queue.get_nowait()
                
                if message[0] == "log":
                    self.log(message[1])
                elif message[0] == "update_status":
                    self.update_file_status(*message[1])
                elif message[0] == "update_progress":
                    if hasattr(self, "progress_window") and self.progress_window:
                        self.progress_window.update_progress(message[1])
                elif message[0] == "processing_done":
                    self.processing_done()
                
                self.queue.task_done()
        except queue.Empty:
            pass
        
        # 继续检查队列
        self.after(100, self.check_queue)
    
    def update_file_status(self, index, file_name, status):
        """更新文件状态"""
        items = self.file_tree.get_children()
        if index < len(items):
            item = items[index]
            values = list(self.file_tree.item(item, "values"))
            values[2] = status
            self.file_tree.item(item, values=values)
    
    def processing_done(self):
        """处理完成后的操作"""
        # 启用按钮
        self.select_pdf_btn.config(state=tk.NORMAL)
        self.select_output_btn.config(state=tk.NORMAL)
        self.process_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        # 更新状态
        self.status_label.config(text="处理完成")
        
        # 关闭进度窗口
        if hasattr(self, "progress_window") and self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None
        
        # 显示完成消息
        if self.is_processing:  # 如果不是被用户中断的
            if messagebox.askyesno("处理完成", "所有文件处理完成，是否打开输出文件夹？"):
                self.open_output_folder()
        
        self.is_processing = False
    
    def open_output_folder(self):
        """打开输出文件夹"""
        if self.output_file:
            output_dir = os.path.dirname(self.output_file)
            if os.path.exists(output_dir):
                os.startfile(output_dir)
            else:
                messagebox.showerror("错误", "输出文件夹不存在")
        else:
            messagebox.showerror("错误", "未设置输出路径")
    
    def open_excel_file(self):
        """打开生成的Excel文件"""
        if self.output_file and os.path.exists(self.output_file):
            os.startfile(self.output_file)
        else:
            messagebox.showerror("错误", "无法打开Excel文件，文件不存在")
    
    def log(self, message, level="INFO"):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] [{level}] {message}"
        
        # 根据日志级别设置颜色
        tag = None
        if level == "ERROR":
            tag = "error"
        elif level == "WARNING":
            tag = "warning"
        elif level == "SUCCESS":
            tag = "success"
        
        # 启用编辑
        self.log_text.config(state=tk.NORMAL)
        
        # 添加日志消息
        self.log_text.insert(tk.END, log_message + "\n")
        
        # 应用标签
        if tag:
            line_count = int(self.log_text.index(tk.END).split(".")[0])
            start_index = f"{line_count-1}.0"
            end_index = f"{line_count-1}.end"
            self.log_text.tag_add(tag, start_index, end_index)
        
        # 滚动到底部
        self.log_text.see(tk.END)
        
        # 禁用编辑
        self.log_text.config(state=tk.DISABLED)
        
        # 同时输出到控制台
        print(log_message)
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        # 获取点击的项
        item = self.file_tree.identify_row(event.y)
        if item:
            # 选中该项
            self.file_tree.selection_set(item)
            # 显示菜单
            self.context_menu.post(event.x_root, event.y_root)
    
    def remove_selected_file(self):
        """移除选中的文件"""
        selected = self.file_tree.selection()
        if not selected:
            return
        
        # 获取选中项的索引
        items = self.file_tree.get_children()
        index = items.index(selected[0])
        
        # 从列表中移除
        if index < len(self.pdf_files):
            file_info = self.pdf_files.pop(index)
            self.log(f"移除文件: {file_info['name']}")
        
        # 从树中移除
        self.file_tree.delete(selected[0])
        
        # 更新状态
        self.status_label.config(text=f"已添加 {len(self.pdf_files)} 个文件")
    
    def preview_pdf(self):
        """预览PDF内容"""
        selected = self.file_tree.selection()
        if not selected:
            return
        
        # 获取选中项的索引
        items = self.file_tree.get_children()
        index = items.index(selected[0])
        
        # 获取文件路径
        if index < len(self.pdf_files):
            file_path = self.pdf_files[index]["path"]
            
            # 使用系统默认程序打开PDF
            try:
                os.startfile(file_path)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开PDF文件: {str(e)}")
    
    def open_bank_mapping(self):
        """打开银行映射设置"""
        dialog = BankMappingDialog(self, self.bank_mapping)
        self.wait_window(dialog)
        
        if dialog.result:
            self.bank_mapping = dialog.result
            self.app_config["bank_mapping"] = self.bank_mapping
            save_config(self.app_config)
            self.log("已更新银行识别映射设置")
    
    def open_advanced_settings(self):
        """高级设置"""
        dialog = AdvancedSettingsDialog(self, self.app_config)
        self.wait_window(dialog)
        
        if dialog.result:
            # 更新配置
            for key, value in dialog.result.items():
                self.app_config[key] = value
            
            save_config(self.app_config)
            self.log("已更新高级设置")
    
    def show_help(self):
        """显示帮助信息"""
        help_text = """
使用说明：

1. 选择PDF文件：点击"选择PDF文件"按钮，选择需要处理的银行账单PDF文件。
2. 选择输出文件：点击"选择输出文件"按钮，选择保存结果的Excel文件路径。
3. 开始处理：点击"开始处理"按钮，程序将自动识别银行类型并提取交易记录。
4. 查看结果：处理完成后，可以选择打开生成的Excel文件查看结果。

支持的银行类型：
- 玉山银行
- 渣打银行
- 汇丰银行
- 南洋银行
- 恒生银行
- 中银香港
- 东亚银行
- 大新银行


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
                        """

        messagebox.showinfo("关于", about_text)
    
    def load_last_session(self):
        """加载上次会话的配置"""
        # 加载银行映射
        self.bank_mapping = self.app_config.get("bank_mapping", {})
        
        # 加载上次的输出文件
        last_output = self.app_config.get("last_output_file", "")
        if last_output and os.path.dirname(last_output):
            self.output_file = last_output
            self.output_label.config(text=f"输出文件: {os.path.basename(last_output)}")
    
    def on_closing(self):
        """关闭应用程序时的操作"""
        # 保存当前配置
        if self.output_file:
            self.app_config["last_output_file"] = self.output_file
        
        save_config(self.app_config)
        
        # 关闭应用程序
        self.destroy()

# 简化版本 - 移除启动画面
def main():
    app = BankStatementApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()

# 主程序入口
if __name__ == "__main__":
    # 检查是否是PyInstaller打包的可执行文件
    if getattr(sys, 'frozen', False):
        # 如果是打包的可执行文件，设置工作目录为可执行文件所在目录
        os.chdir(os.path.dirname(sys.executable))
    
    main()