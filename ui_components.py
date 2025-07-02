import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import os

class ToolTipButton(ttk.Button):
    """带有工具提示的按钮"""
    
    def __init__(self, master=None, tooltip="", **kwargs):
        super().__init__(master, **kwargs)
        self.tooltip = tooltip
        self.tooltip_window = None
        
        # 绑定鼠标事件
        self.bind("<Enter>", self._show_tooltip)
        self.bind("<Leave>", self._hide_tooltip)
    
    def _show_tooltip(self, event=None):
        """显示工具提示"""
        if self.tooltip:
            x, y, _, _ = self.bbox("insert")
            x += self.winfo_rootx() + 25
            y += self.winfo_rooty() + 25
            
            # 创建工具提示窗口
            self.tooltip_window = tk.Toplevel(self)
            self.tooltip_window.wm_overrideredirect(True)  # 无边框窗口
            self.tooltip_window.wm_geometry(f"+{x}+{y}")
            
            # 添加标签
            label = ttk.Label(self.tooltip_window, text=self.tooltip, 
                             background="#ffffe0", relief="solid", borderwidth=1,
                             wraplength=180, justify="left", padding=(5, 2))
            label.pack()
    
    def _hide_tooltip(self, event=None):
        """隐藏工具提示"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class ProgressWindow(tk.Toplevel):
    """进度窗口"""
    
    def __init__(self, parent, total_files):
        super().__init__(parent)
        self.title("处理进度")
        self.geometry("400x150")
        self.transient(parent)  # 设置为父窗口的临时窗口
        self.resizable(False, False)
        
        # 设置窗口位置居中
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # 创建进度条
        self.progress_frame = ttk.Frame(self, padding="10")
        self.progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_label = ttk.Label(self.progress_frame, text="正在处理文件...")
        self.progress_label.pack(pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, 
                                          length=380, mode="determinate")
        self.progress_bar.pack(pady=(0, 10))
        
        self.file_label = ttk.Label(self.progress_frame, text="")
        self.file_label.pack()
        
        # 设置进度条最大值
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        
        # 禁止关闭按钮
        self.protocol("WM_DELETE_WINDOW", lambda: None)
    
    def update_progress(self, value, file_name=None):
        """更新进度"""
        self.progress_bar["value"] = value
        
        # 更新百分比
        percent = int((value / self.progress_bar["maximum"]) * 100)
        self.progress_label.config(text=f"正在处理文件... {percent}%")
        
        # 更新文件名
        if file_name:
            self.file_label.config(text=f"当前文件: {file_name}")
        
        # 刷新窗口
        self.update_idletasks()


class BankMappingDialog(tk.Toplevel):
    """银行映射设置对话框"""
    
    def __init__(self, parent, bank_mapping=None):
        super().__init__(parent)
        self.title("银行识别映射设置")
        self.geometry("600x400")
        self.transient(parent)  # 设置为父窗口的临时窗口
        self.grab_set()  # 模态对话框
        
        # 初始化结果
        self.result = None
        self.bank_mapping = bank_mapping or {}
        
        # 创建界面
        self.create_widgets()
        
        # 设置窗口位置居中
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 说明标签
        ttk.Label(main_frame, text="设置银行识别关键词映射，每行一个关键词，用于自动识别银行类型").pack(anchor=tk.W, pady=(0, 10))
        
        # 创建表格
        columns = ("银行名称", "关键词")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        # 设置列标题
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # 设置列宽
        self.tree.column("银行名称", width=150)
        self.tree.column("关键词", width=400)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 放置表格和滚动条
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 加载现有映射
        self.load_mapping()
        
        # 按钮框架
        btn_frame = ttk.Frame(self, padding="10")
        btn_frame.pack(fill=tk.X)
        
        # 添加按钮
        ttk.Button(btn_frame, text="添加", command=self.add_mapping).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="编辑", command=self.edit_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除", command=self.delete_mapping).pack(side=tk.LEFT, padx=5)
        
        # 确定取消按钮
        ttk.Button(btn_frame, text="确定", command=self.on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=self.on_cancel).pack(side=tk.RIGHT, padx=5)
    
    def load_mapping(self):
        """加载现有映射"""
        # 清空表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加映射
        for bank_name, keywords in self.bank_mapping.items():
            self.tree.insert("", tk.END, values=(bank_name, ", ".join(keywords)))
    
    def add_mapping(self):
        """添加映射"""
        dialog = BankMappingEntryDialog(self)
        self.wait_window(dialog)
        
        if dialog.result:
            bank_name, keywords = dialog.result
            
            # 检查是否已存在
            for item in self.tree.get_children():
                if self.tree.item(item, "values")[0] == bank_name:
                    messagebox.showwarning("警告", f"银行 '{bank_name}' 已存在，请编辑现有条目")
                    return
            
            # 添加到表格
            self.tree.insert("", tk.END, values=(bank_name, ", ".join(keywords)))
    
    def edit_mapping(self):
        """编辑映射"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要编辑的条目")
            return
        
        # 获取选中项的值
        item = selected[0]
        values = self.tree.item(item, "values")
        bank_name = values[0]
        keywords = [k.strip() for k in values[1].split(",")]
        
        # 打开编辑对话框
        dialog = BankMappingEntryDialog(self, bank_name, keywords)
        self.wait_window(dialog)
        
        if dialog.result:
            new_bank_name, new_keywords = dialog.result
            
            # 更新表格
            self.tree.item(item, values=(new_bank_name, ", ".join(new_keywords)))
    
    def delete_mapping(self):
        """删除映射"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的条目")
            return
        
        # 确认删除
        if messagebox.askyesno("确认", "确定要删除选中的条目吗？"):
            for item in selected:
                self.tree.delete(item)
    
    def on_ok(self):
        """确定按钮事件"""
        # 收集表格中的数据
        mapping = {}
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            bank_name = values[0]
            keywords = [k.strip() for k in values[1].split(",")]
            mapping[bank_name] = keywords
        
        self.result = mapping
        self.destroy()
    
    def on_cancel(self):
        """取消按钮事件"""
        self.destroy()


class BankMappingEntryDialog(tk.Toplevel):
    """银行映射条目编辑对话框"""
    
    def __init__(self, parent, bank_name=None, keywords=None):
        super().__init__(parent)
        self.title("编辑银行映射")
        self.geometry("400x300")
        self.transient(parent)  # 设置为父窗口的临时窗口
        self.grab_set()  # 模态对话框
        
        # 初始化结果
        self.result = None
        
        # 创建界面
        self.create_widgets(bank_name, keywords)
        
        # 设置窗口位置居中
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self, bank_name=None, keywords=None):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 银行名称
        ttk.Label(main_frame, text="银行名称:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.bank_name_entry = ttk.Entry(main_frame, width=30)
        self.bank_name_entry.grid(row=0, column=1, sticky=tk.W, pady=(0, 5))
        
        # 关键词
        ttk.Label(main_frame, text="关键词列表:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Label(main_frame, text="(每行一个关键词)").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        # 关键词文本框
        self.keywords_text = ScrolledText(main_frame, width=30, height=10)
        self.keywords_text.grid(row=1, column=1, rowspan=2, sticky=tk.W, pady=(0, 5))
        
        # 设置初始值
        if bank_name:
            self.bank_name_entry.insert(0, bank_name)
        
        if keywords:
            self.keywords_text.insert(tk.END, "\n".join(keywords))
        
        # 按钮框架
        btn_frame = ttk.Frame(self, padding="10")
        btn_frame.pack(fill=tk.X)
        
        # 确定取消按钮
        ttk.Button(btn_frame, text="确定", command=self.on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=self.on_cancel).pack(side=tk.RIGHT, padx=5)
    
    def on_ok(self):
        """确定按钮事件"""
        # 获取输入值
        bank_name = self.bank_name_entry.get().strip()
        keywords_text = self.keywords_text.get(1.0, tk.END).strip()
        
        # 验证输入
        if not bank_name:
            messagebox.showwarning("警告", "请输入银行名称")
            return
        
        if not keywords_text:
            messagebox.showwarning("警告", "请输入至少一个关键词")
            return
        
        # 处理关键词
        keywords = [k.strip() for k in keywords_text.split("\n") if k.strip()]
        
        self.result = (bank_name, keywords)
        self.destroy()
    
    def on_cancel(self):
        """取消按钮事件"""
        self.destroy()


class AdvancedSettingsDialog(tk.Toplevel):
    """高级设置对话框"""
    
    def __init__(self, parent, config=None):
        super().__init__(parent)
        self.title("高级设置")
        self.geometry("500x400")
        self.transient(parent)  # 设置为父窗口的临时窗口
        self.grab_set()  # 模态对话框
        
        # 初始化结果
        self.result = None
        self.config = config or {}
        
        # 创建界面
        self.create_widgets()
        
        # 设置窗口位置居中
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建选项卡
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 常规设置选项卡
        general_frame = ttk.Frame(notebook, padding="10")
        notebook.add(general_frame, text="常规设置")
        
        # 线程设置
        ttk.Label(general_frame, text="最大线程数:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.max_threads_var = tk.StringVar(value=str(self.config.get("max_threads", os.cpu_count() or 4)))
        tk.Spinbox(general_frame, from_=1, to=32, textvariable=self.max_threads_var, width=5).grid(row=0, column=1, sticky=tk.W, pady=(0, 5))
        
        # 自动打开结果文件
        self.auto_open_var = tk.BooleanVar(value=self.config.get("auto_open_result", True))
        ttk.Checkbutton(general_frame, text="处理完成后自动打开结果文件", variable=self.auto_open_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # 保存日志到文件
        self.save_log_var = tk.BooleanVar(value=self.config.get("save_log", True))
        ttk.Checkbutton(general_frame, text="保存日志到文件", variable=self.save_log_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # 显示处理时间
        self.show_time_var = tk.BooleanVar(value=self.config.get("show_processing_time", True))
        ttk.Checkbutton(general_frame, text="显示处理时间", variable=self.show_time_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # 导出设置选项卡
        export_frame = ttk.Frame(notebook, padding="10")
        notebook.add(export_frame, text="导出设置")
        
        # 默认导出格式
        ttk.Label(export_frame, text="默认导出格式:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.export_format_var = tk.StringVar(value=self.config.get("default_export_format", "xlsx"))
        ttk.Combobox(export_frame, textvariable=self.export_format_var, values=["xlsx", "csv"], width=10).grid(row=0, column=1, sticky=tk.W, pady=(0, 5))
        
        # 包含汇总表
        self.include_summary_var = tk.BooleanVar(value=self.config.get("include_summary", True))
        ttk.Checkbutton(export_frame, text="包含汇总表", variable=self.include_summary_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # 包含银行汇总
        self.include_bank_summary_var = tk.BooleanVar(value=self.config.get("include_bank_summary", True))
        ttk.Checkbutton(export_frame, text="包含银行汇总", variable=self.include_bank_summary_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # 包含月度汇总
        self.include_month_summary_var = tk.BooleanVar(value=self.config.get("include_month_summary", True))
        ttk.Checkbutton(export_frame, text="包含月度汇总", variable=self.include_month_summary_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        # 按钮框架
        btn_frame = ttk.Frame(self, padding="10")
        btn_frame.pack(fill=tk.X)
        
        # 确定取消按钮
        ttk.Button(btn_frame, text="确定", command=self.on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=self.on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="恢复默认", command=self.reset_defaults).pack(side=tk.LEFT, padx=5)
    
    def on_ok(self):
        """确定按钮事件"""
        # 收集设置
        settings = {
            "max_threads": int(self.max_threads_var.get()),
            "auto_open_result": self.auto_open_var.get(),
            "save_log": self.save_log_var.get(),
            "show_processing_time": self.show_time_var.get(),
            "default_export_format": self.export_format_var.get(),
            "include_summary": self.include_summary_var.get(),
            "include_bank_summary": self.include_bank_summary_var.get(),
            "include_month_summary": self.include_month_summary_var.get()
        }
        
        self.result = settings
        self.destroy()
    
    def on_cancel(self):
        """取消按钮事件"""
        self.destroy()
    
    def reset_defaults(self):
        """恢复默认设置"""
        # 设置默认值
        self.max_threads_var.set(str(os.cpu_count() or 4))
        self.auto_open_var.set(True)
        self.save_log_var.set(True)
        self.show_time_var.set(True)
        self.export_format_var.set("xlsx")
        self.include_summary_var.set(True)
        self.include_bank_summary_var.set(True)
        self.include_month_summary_var.set(True)