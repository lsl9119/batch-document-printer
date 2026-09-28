"""
打印设置对话框
提供打印机选择和参数配置界面
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional

from src.core.settings_manager import PrinterSettingsManager
from src.core.models import (PrintSettings, ColorMode, Orientation, DuplexMode,
                             OfficeBackendPreference)

BACKEND_CHOICES = [
    ("自动（优先 WPS Office）", OfficeBackendPreference.AUTO),
    ("仅 WPS Office", OfficeBackendPreference.WPS),
    ("仅 Microsoft Office", OfficeBackendPreference.MS_OFFICE),
]


class PrintSettingsDialog:
    """打印设置对话框类"""
    
    def __init__(self, parent, printer_manager: PrinterSettingsManager, current_settings: PrintSettings):
        """
        初始化打印设置对话框
        
        Args:
            parent: 父窗口
            printer_manager: 打印机管理器
            current_settings: 当前打印设置
        """
        self.parent = parent
        self.printer_manager = printer_manager
        self.current_settings = current_settings
        self.result: Optional[PrintSettings] = None
        
        # 创建对话框窗口
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("打印设置")
        self.dialog.geometry("600x600")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self._center_dialog()
        
        # 创建界面
        self._create_widgets()
        self._load_current_settings()
        
        # 等待对话框关闭
        self.dialog.wait_window()
    
    def _center_dialog(self):
        """将对话框居中显示"""
        self.dialog.update_idletasks()
        
        # 获取父窗口位置和大小
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # 计算对话框位置
        dialog_width = self.dialog.winfo_reqwidth()
        dialog_height = self.dialog.winfo_reqheight()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 打印机选择区域
        self._create_printer_section(main_frame)
        
        # 纸张设置区域
        self._create_paper_section(main_frame)
        
        # 打印选项区域
        self._create_options_section(main_frame)
        
        # 按钮区域
        self._create_button_section(main_frame)
    
    def _create_printer_section(self, parent):
        """创建打印机选择区域"""
        # 打印机框架
        printer_frame = ttk.LabelFrame(parent, text="打印机选择", padding="10")
        printer_frame.pack(fill="x", pady=(0, 10))
        
        # 打印机下拉框
        ttk.Label(printer_frame, text="打印机:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(
            printer_frame, 
            textvariable=self.printer_var,
            state="readonly",
            width=40
        )
        self.printer_combo.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        # 刷新按钮
        self.btn_refresh = ttk.Button(
            printer_frame, 
            text="刷新",
            command=self._refresh_printers
        )
        self.btn_refresh.grid(row=0, column=2)
        
        # 测试按钮
        self.btn_test = ttk.Button(
            printer_frame,
            text="测试连接",
            command=self._test_printer
        )
        self.btn_test.grid(row=0, column=3, padx=(5, 0))
        
        # 打印机信息显示
        self.lbl_printer_info = ttk.Label(printer_frame, text="", foreground="gray")
        self.lbl_printer_info.grid(row=1, column=0, columnspan=4, sticky="w", pady=(5, 0))
        
        printer_frame.grid_columnconfigure(1, weight=1)
        
        # 绑定打印机选择变化事件
        self.printer_combo.bind('<<ComboboxSelected>>', self._on_printer_changed)
    
    def _create_paper_section(self, parent):
        """创建纸张设置区域"""
        paper_frame = ttk.LabelFrame(parent, text="纸张设置", padding="10")
        paper_frame.pack(fill="x", pady=(0, 10))
        
        # 纸张尺寸行
        paper_row_frame = ttk.Frame(paper_frame)
        paper_row_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        
        ttk.Label(paper_row_frame, text="纸张尺寸:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.paper_var = tk.StringVar()
        self.paper_combo = ttk.Combobox(
            paper_row_frame,
            textvariable=self.paper_var,
            state="readonly",
            width=25
        )
        self.paper_combo.grid(row=0, column=1, sticky="w", padx=(0, 10))
        
        # 同步纸张按钮
        self.btn_sync_paper = ttk.Button(
            paper_row_frame,
            text="同步电脑纸张",
            command=self._sync_paper_sizes
        )
        self.btn_sync_paper.grid(row=0, column=2, padx=(5, 0))
        
        # 页面方向
        ttk.Label(paper_frame, text="页面方向:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(10, 0))
        
        self.orientation_var = tk.StringVar()
        orientation_frame = ttk.Frame(paper_frame)
        orientation_frame.grid(row=1, column=1, sticky="w", pady=(10, 0))
        
        for text, value in (("自动", "auto"), ("纵向", "portrait"), ("横向", "landscape")):
            ttk.Radiobutton(orientation_frame, text=text, variable=self.orientation_var,
                            value=value).pack(side="left", padx=(0, 10))

        # Office 文档是否强制套用纸张/方向
        self.force_page_setup_var = tk.BooleanVar()
        ttk.Checkbutton(
            paper_frame,
            text="Office 文档（文字/表格）也强制使用以上纸张和方向（默认遵循文档自身的页面设置）",
            variable=self.force_page_setup_var
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        
        # 配置网格权重
        paper_frame.grid_columnconfigure(1, weight=1)
        paper_row_frame.grid_columnconfigure(1, weight=1)
    
    def _create_options_section(self, parent):
        """创建打印选项区域"""
        options_frame = ttk.LabelFrame(parent, text="打印选项", padding="10")
        options_frame.pack(fill="x", pady=(0, 10))
        
        # 打印数量
        ttk.Label(options_frame, text="打印数量:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.copies_var = tk.IntVar()
        self.copies_spin = ttk.Spinbox(
            options_frame,
            from_=1,
            to=999,
            textvariable=self.copies_var,
            width=10
        )
        self.copies_spin.grid(row=0, column=1, sticky="w")
        
        # 颜色模式
        ttk.Label(options_frame, text="颜色模式:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(10, 0))
        
        self.color_var = tk.StringVar()
        color_frame = ttk.Frame(options_frame)
        color_frame.grid(row=1, column=1, sticky="w", pady=(10, 0))
        
        self.rb_color = ttk.Radiobutton(
            color_frame,
            text="彩色",
            variable=self.color_var,
            value="color"
        )
        self.rb_color.pack(side="left", padx=(0, 10))
        
        self.rb_grayscale = ttk.Radiobutton(
            color_frame,
            text="黑白",
            variable=self.color_var,
            value="grayscale"
        )
        self.rb_grayscale.pack(side="left")
        
        # 单双面
        ttk.Label(options_frame, text="单双面:").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(10, 0))
        self.duplex_var = tk.StringVar()
        duplex_frame = ttk.Frame(options_frame)
        duplex_frame.grid(row=2, column=1, sticky="w", pady=(10, 0))
        for text, value in (("单面", DuplexMode.SIMPLEX.value), ("双面-长边翻转", DuplexMode.DUPLEX_LONG.value),
                            ("双面-短边翻转", DuplexMode.DUPLEX_SHORT.value)):
            ttk.Radiobutton(duplex_frame, text=text, variable=self.duplex_var,
                            value=value).pack(side="left", padx=(0, 10))

        # Office 引擎
        ttk.Label(options_frame, text="Office 引擎:").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(10, 0))
        self.backend_var = tk.StringVar()
        self.backend_combo = ttk.Combobox(options_frame, textvariable=self.backend_var, state="readonly",
                                          width=28, values=[label for label, _ in BACKEND_CHOICES])
        self.backend_combo.grid(row=3, column=1, sticky="w", pady=(10, 0))

        ttk.Label(options_frame, foreground="gray",
                  text="提示：打印机驱动不支持的设置（如双面、彩色）会自动使用驱动默认值，并记录在日志中。"
                  ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
    
    def _create_button_section(self, parent):
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill="x", pady=(20, 10))
        
        # 左侧按钮框架
        left_frame = ttk.Frame(button_frame)
        left_frame.pack(side="left")
        
        # 重置按钮
        self.btn_reset = ttk.Button(
            left_frame,
            text="重置",
            command=self._on_reset,
            width=10
        )
        self.btn_reset.pack(side="left")
        
        # 右侧按钮框架
        right_frame = ttk.Frame(button_frame)
        right_frame.pack(side="right")
        
        # 取消按钮
        self.btn_cancel = ttk.Button(
            right_frame,
            text="取消",
            command=self._on_cancel,
            width=10
        )
        self.btn_cancel.pack(side="left", padx=(0, 5))
        
        # 确定按钮
        self.btn_ok = ttk.Button(
            right_frame,
            text="确定",
            command=self._on_ok,
            width=10
        )
        self.btn_ok.pack(side="left")
    
    def _refresh_printers(self):
        """刷新打印机列表"""
        try:
            self.printer_manager.refresh_printer_list()
            printers = self.printer_manager.available_printers
            
            self.printer_combo['values'] = printers
            
            if printers:
                # 如果有默认打印机，选择它
                default_printer = self.printer_manager.default_printer
                if default_printer and default_printer in printers:
                    self.printer_var.set(default_printer)
                elif not self.printer_var.get():
                    self.printer_var.set(printers[0])
                
                self._update_printer_info()
            else:
                self.lbl_printer_info.config(text="未找到可用的打印机")
                
        except Exception as e:
            messagebox.showerror("错误", f"刷新打印机列表失败: {e}")
    
    def _test_printer(self):
        """测试打印机连接"""
        printer_name = self.printer_var.get()
        if not printer_name:
            messagebox.showwarning("提示", "请先选择打印机")
            return
        
        try:
            if self.printer_manager.test_printer_connection(printer_name):
                # 去掉连接成功提示窗口
                print(f"打印机 '{printer_name}' 连接正常")
            else:
                messagebox.showerror("错误", f"无法连接到打印机 '{printer_name}'")
        except Exception as e:
            messagebox.showerror("错误", f"测试打印机连接失败: {e}")
    
    def _update_printer_info(self):
        """更新打印机信息显示"""
        printer_name = self.printer_var.get()
        if printer_name:
            try:
                info = self.printer_manager.get_printer_info(printer_name)
                if info:
                    info_text = f"驱动: {info.get('driver', 'N/A')}, 端口: {info.get('port', 'N/A')}"
                    if info.get('location'):
                        info_text += f", 位置: {info['location']}"
                    self.lbl_printer_info.config(text=info_text)
                else:
                    self.lbl_printer_info.config(text="无法获取打印机信息")
            except Exception as e:
                self.lbl_printer_info.config(text=f"获取信息失败: {e}")
        else:
            self.lbl_printer_info.config(text="")
    
    def _load_current_settings(self):
        """加载当前设置到界面"""
        # 首先自动刷新打印机列表，确保数据是最新的
        print("自动刷新打印机列表...")
        self._refresh_printers()
        
        # 设置打印机
        if self.current_settings.printer_name:
            # 检查打印机是否还存在
            available_printers = self.printer_manager.available_printers
            if self.current_settings.printer_name in available_printers:
                self.printer_var.set(self.current_settings.printer_name)
            else:
                # 如果原打印机不存在，使用默认打印机
                default_printer = self.printer_manager.default_printer
                if default_printer:
                    self.printer_var.set(default_printer)
                    print(f"原打印机不可用，已切换到默认打印机: {default_printer}")
            
            self._update_printer_info()
        
        # 设置默认纸张列表（标准纸张，不自动同步系统纸张）
        self._load_default_paper_sizes()
        
        # 设置纸张尺寸（优先使用A4）
        if self.current_settings.paper_size and self.current_settings.paper_size in list(self.paper_combo['values']):
            self.paper_var.set(self.current_settings.paper_size)
        else:
            # 默认使用A4
            self.paper_var.set('A4')
        
        # 设置页面方向
        self.orientation_var.set(self.current_settings.orientation.value)
        
        # 设置打印数量
        self.copies_var.set(self.current_settings.copies)
        
        # 设置颜色模式
        self.color_var.set(self.current_settings.color_mode.value)
        
        # 设置单双面
        self.duplex_var.set(self.current_settings.effective_duplex.value)

        # Office 选项
        self.force_page_setup_var.set(self.current_settings.office_force_page_setup)
        backend_label = next((label for label, value in BACKEND_CHOICES
                              if value == self.current_settings.office_backend), BACKEND_CHOICES[0][0])
        self.backend_var.set(backend_label)
        
        print("打印设置加载完成（纸张使用默认列表，如需同步系统纸张请点击'同步纸张'按钮）")

    def _load_default_paper_sizes(self):
        """加载默认纸张尺寸列表"""
        # 使用标准纸张尺寸列表
        standard_papers = list(self.printer_manager.STANDARD_PAPER_SIZES.keys())
        self.paper_combo['values'] = standard_papers
        print(f"已加载默认纸张列表，共 {len(standard_papers)} 种标准格式")
    
    def _validate_settings(self) -> bool:
        """验证设置是否有效"""
        # 检查打印机
        if not self.printer_var.get():
            messagebox.showerror("错误", "请选择打印机")
            return False
        
        # 检查纸张尺寸
        if not self.paper_var.get():
            messagebox.showerror("错误", "请选择纸张尺寸")
            return False
        
        # 检查打印数量
        try:
            copies = self.copies_var.get()
            if copies < 1 or copies > 999:
                messagebox.showerror("错误", "打印数量必须在1-999之间")
                return False
        except (tk.TclError, ValueError):
            messagebox.showerror("错误", "打印数量必须是有效数字")
            return False
        
        return True
    
    def _create_settings_from_form(self) -> PrintSettings:
        """从表单创建打印设置对象"""
        duplex_mode = DuplexMode(self.duplex_var.get() or DuplexMode.SIMPLEX.value)
        backend = next((value for label, value in BACKEND_CHOICES if label == self.backend_var.get()),
                       OfficeBackendPreference.AUTO)
        return PrintSettings(
            printer_name=self.printer_var.get(),
            paper_size=self.paper_var.get(),
            copies=self.copies_var.get(),
            duplex=duplex_mode != DuplexMode.SIMPLEX,
            duplex_mode=duplex_mode if duplex_mode != DuplexMode.SIMPLEX else DuplexMode.DUPLEX_LONG,
            color_mode=ColorMode(self.color_var.get()),
            orientation=Orientation(self.orientation_var.get() or "auto"),
            scaling=self.current_settings.scaling,
            office_force_page_setup=self.force_page_setup_var.get(),
            office_backend=backend,
        )
    
    def _on_ok(self):
        """确定按钮处理"""
        if self._validate_settings():
            try:
                self.result = self._create_settings_from_form()
                
                # 验证设置
                is_valid, errors = self.printer_manager.validate_settings(self.result)
                if not is_valid:
                    error_msg = "\n".join(errors)
                    messagebox.showerror("设置错误", f"打印设置验证失败:\n{error_msg}")
                    return
                
                print(f"打印设置已确认: {self.result.printer_name}")
                self.dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("错误", f"保存设置失败: {e}")
    
    def _on_cancel(self):
        """取消按钮处理"""
        self.result = None
        self.dialog.destroy()
    
    def _on_reset(self):
        """重置按钮处理"""
        if messagebox.askyesno("确认", "确定要重置所有设置吗？"):
            # 重置为默认设置
            default_settings = self.printer_manager.create_default_settings()
            self.current_settings = default_settings
            self._load_current_settings()
    
    def _on_printer_changed(self, event):
        """处理打印机选择变化事件"""
        self._update_printer_info()
        # 不自动同步纸张，用户需要手动点击"同步纸张"按钮
    
    def _sync_paper_sizes(self):
        """同步纸张尺寸列表"""
        printer_name = self.printer_var.get()
        if not printer_name:
            return
        
        try:
            # 获取选中打印机支持的纸张尺寸
            paper_sizes = self.printer_manager.get_printer_paper_sizes(printer_name)
            self.paper_combo['values'] = paper_sizes
            
            # 保存当前选择的纸张尺寸
            current_paper = self.paper_var.get()
            
            # 如果当前纸张在新列表中，保持选择；否则选择默认A4或第一个
            if current_paper and current_paper in paper_sizes:
                self.paper_var.set(current_paper)
            elif 'A4' in paper_sizes:
                self.paper_var.set('A4')
            elif paper_sizes:
                self.paper_var.set(paper_sizes[0])
            else:
                self.paper_var.set('')
            
            print(f"已更新打印机 '{printer_name}' 的纸张列表，共 {len(paper_sizes)} 种格式")
            
        except Exception as e:
            print(f"同步纸张尺寸失败: {e}")
            # 如果失败，使用标准纸张尺寸作为备用
            standard_papers = list(self.printer_manager.STANDARD_PAPER_SIZES.keys())
            self.paper_combo['values'] = standard_papers
            if 'A4' in standard_papers:
                self.paper_var.set('A4')
    
 