"""
主窗口界面
办公文档批量打印应用的主界面 (重构版本)
"""
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import List, Optional
import threading

# 拖拽支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    print("警告: tkinterdnd2 未安装，拖拽功能将不可用")
    DRAG_DROP_AVAILABLE = False

from src.core.document_manager import DocumentManager
from src.core.settings_manager import PrinterSettingsManager
from src.core.print_controller import PrintController
from src.core.models import Document, PrintStatus, OFFICE_FILE_TYPES
from src.core.environment import EnvironmentReport, collect_environment
from src.backends import process_utils
from src.backends.office_backend import KIND_BY_FILE_TYPE
from src.utils.config_utils import ConfigManager
from src.utils.path_utils import get_app_icon_path, get_logs_dir
from src.gui.print_settings_dialog import PrintSettingsDialog
from src.gui.page_count_dialog import show_page_count_dialog
from src.gui.environment_dialog import show_environment_dialog, show_about_dialog
from src.version import APP_NAME, VERSION

UI_POLL_MS = 100

# 导入功能处理器
from src.gui.components import FileImportHandler, ListOperationHandler, WindowManager, create_button_tooltip


class MainWindow:
    """主窗口类 (重构版本)"""
    
    def __init__(self, initial_paths: Optional[List[str]] = None):
        """
        初始化主窗口

        Args:
            initial_paths: 启动时要添加的文件/文件夹（例如拖到 exe 图标上的文件）
        """
        # 创建支持拖拽的主窗口
        if DRAG_DROP_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        # 初始化核心管理器
        self.document_manager = DocumentManager()
        self.printer_manager = PrinterSettingsManager()
        self.print_controller = PrintController()
        self.config_manager = ConfigManager()
        
        # 加载配置
        self.app_config = self.config_manager.load_app_config()
        self.current_print_settings = self.config_manager.load_print_settings()
        
        # 初始化功能处理器
        self._setup_handlers()
        
        # 后台线程 → 界面线程的消息队列（Tkinter 控件只能在主线程中操作）
        self._ui_queue: "queue.Queue" = queue.Queue()
        self.env_report: Optional[EnvironmentReport] = None

        # 设置打印控制器（回调在工作线程中执行，只向队列投递消息）
        self.print_controller.set_print_settings(self.current_print_settings)
        self.print_controller.set_progress_callback(
            lambda cur, total, msg: self._ui_queue.put(("progress", (cur, total, msg))))
        self.print_controller.set_document_callback(
            lambda doc: self._ui_queue.put(("document", doc)))
        self.print_controller.set_finished_callback(
            lambda summary: self._ui_queue.put(("finished", summary)))
        
        # 创建界面
        self._create_widgets()
        self._setup_layout()
        self._bind_events()
        
        # 设置窗口属性
        self._setup_window()
        
        # 启动界面消息轮询与后台环境检测
        self.root.after(UI_POLL_MS, self._process_ui_queue)
        if initial_paths:
            self.root.after(200, lambda: self._add_initial_paths(initial_paths))
        self._start_environment_check()
        print(f"{APP_NAME} v{VERSION} 已启动")
    
    def _setup_handlers(self):
        """初始化功能处理器"""
        # 窗口管理器
        self.window_manager = WindowManager(self.root, self.config_manager)
        
        # 文件导入处理器 (需要在创建界面后初始化拖拽功能)
        self.file_import_handler = FileImportHandler(
            self.document_manager, 
            self._get_enabled_file_types,
            self._on_files_imported
        )
        
        # 列表操作处理器 (需要在创建树形控件后初始化)
        self.list_operation_handler = None  # 稍后初始化
    
    def _setup_window(self):
        """设置窗口属性"""
        # 设置窗口标题
        self.window_manager.set_window_title(APP_NAME, f"v{VERSION}")
        icon = get_app_icon_path()
        if icon.exists():
            self.window_manager.set_window_icon(str(icon))
        
        # 设置窗口最小尺寸
        self.window_manager.set_window_minimum_size(900, 560)
        
        # 恢复窗口几何属性
        self.window_manager.restore_window_geometry(self.app_config)
        
        # 设置窗口关闭处理
        self.window_manager.setup_window_close_handler(self._on_window_closing)
    
    def _create_widgets(self):
        """创建界面组件"""
        # 工具栏
        self._create_toolbar()

        # 环境状态栏（Office Backend / WPS 组件 / SumatraPDF）
        self._create_environment_bar()
        
        # 主要内容区域
        self.main_frame = ttk.Frame(self.root)
        
        # 文档列表区域
        self._create_document_list()
        
        # 状态区域
        self._create_status_area()
        
        # 进度区域
        self._create_progress_area()
        
        # 创建列表操作处理器 (现在树形控件已创建)
        self.list_operation_handler = ListOperationHandler(
            self.document_manager, 
            self.doc_tree,
            self._on_list_operation_completed
        )
        
        # 设置界面提示功能
        self._setup_tooltips()
    
    def _create_toolbar(self):
        """创建工具栏"""
        self.toolbar = ttk.Frame(self.root)
        
        # 文件操作按钮
        self.btn_add_files = ttk.Button(
            self.toolbar, text="添加文件", 
            command=self._add_files
        )
        
        self.btn_add_folder = ttk.Button(
            self.toolbar, text="添加文件夹", 
            command=self._add_folder
        )
        
        self.btn_remove_selected = ttk.Button(
            self.toolbar, text="移除选中", 
            command=self._remove_selected_documents
        )
         
        self.btn_clear = ttk.Button(
            self.toolbar, text="清空列表", 
            command=self._clear_documents
        )

        self.btn_filter = ttk.Button(
            self.toolbar, text="文件过滤", 
            command=self._filter_documents
        )
        
        # 功能按钮
        self.btn_print_settings = ttk.Button(
            self.toolbar, text="打印设置", 
            command=self._show_print_settings
        )
        
        self.btn_help = ttk.Button(
            self.toolbar, text="使用说明", 
            command=self._show_help
        )

        self.btn_env = ttk.Button(
            self.toolbar, text="环境诊断",
            command=self._show_environment
        )

        self.btn_about = ttk.Button(
            self.toolbar, text="关于",
            command=lambda: show_about_dialog(self.root)
        )
        
        self.btn_calculate_pages = ttk.Button(
            self.toolbar, text="计算页数", 
            command=self._calculate_pages
        )
        
        self.btn_start_print = ttk.Button(
            self.toolbar, text="开始打印", 
            command=self._start_printing,
            style="Accent.TButton"
        )
    
    def _create_environment_bar(self):
        """创建环境状态栏"""
        self.env_bar = ttk.Frame(self.root)
        self.lbl_backend = ttk.Label(self.env_bar, text="Office Backend: 检测中...")
        self.lbl_wps_components = ttk.Label(self.env_bar, text="")
        self.lbl_sumatra = ttk.Label(self.env_bar, text="")
        self.lbl_env_warning = ttk.Label(self.env_bar, text="", foreground="#b35900", cursor="hand2")
        self.lbl_backend.pack(side="left", padx=(3, 12))
        self.lbl_wps_components.pack(side="left", padx=(0, 12))
        self.lbl_sumatra.pack(side="left", padx=(0, 12))
        self.lbl_env_warning.pack(side="left")
        self.lbl_env_warning.bind("<Button-1>", lambda e: self._show_environment())

    def _create_document_list(self):
        """创建文档列表组件"""
        # 创建标题框架
        title_frame = ttk.Frame(self.main_frame)
        
        # 文档列表标题
        title_label = ttk.Label(title_frame, text="文档列表")
        
        # 文件类型过滤器
        self._create_file_type_filters(title_frame)
        
        # 文档列表框架
        list_frame = ttk.LabelFrame(self.main_frame, labelwidget=title_frame, padding="5")
        self.list_frame = list_frame
        
        # 创建树形视图
        self._create_tree_view(list_frame)
    
    def _create_file_type_filters(self, parent):
        """创建文件类型过滤器"""
        # 文件类型勾选框变量
        self.var_word = tk.BooleanVar(value=self.app_config.enabled_file_types.get('word', True))
        self.var_ppt = tk.BooleanVar(value=self.app_config.enabled_file_types.get('ppt', True))
        self.var_excel = tk.BooleanVar(value=self.app_config.enabled_file_types.get('excel', True))
        self.var_pdf = tk.BooleanVar(value=self.app_config.enabled_file_types.get('pdf', True))
        self.var_image = tk.BooleanVar(value=self.app_config.enabled_file_types.get('image', True))
        self.var_text = tk.BooleanVar(value=self.app_config.enabled_file_types.get('text', True))
        
        # 标题
        title_label = ttk.Label(parent, text="文档列表")
        title_label.pack(side="left")
        
        # 文件类型勾选框
        self.chk_word = ttk.Checkbutton(
            parent, text="文字", variable=self.var_word,
            command=self._on_filter_changed
        )
        self.chk_ppt = ttk.Checkbutton(
            parent, text="演示", variable=self.var_ppt,
            command=self._on_filter_changed
        )
        self.chk_excel = ttk.Checkbutton(
            parent, text="表格", variable=self.var_excel,
            command=self._on_filter_changed
        )
        self.chk_pdf = ttk.Checkbutton(
            parent, text="PDF", variable=self.var_pdf,
            command=self._on_filter_changed
        )
        self.chk_image = ttk.Checkbutton(
            parent, text="图片", variable=self.var_image,
            command=self._on_filter_changed
        )
        self.chk_text = ttk.Checkbutton(
            parent, text="文本", variable=self.var_text,
            command=self._on_filter_changed
        )
        
        # 布局
        self.chk_word.pack(side="left", padx=(10, 2))
        self.chk_ppt.pack(side="left", padx=2)
        self.chk_excel.pack(side="left", padx=2)
        self.chk_pdf.pack(side="left", padx=2)
        self.chk_image.pack(side="left", padx=2)
        self.chk_text.pack(side="left", padx=2)
    
    def _create_tree_view(self, parent):
        """创建树形视图"""
        # 创建树形视图容器
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)
        
        # 创建Treeview
        columns = ("文件名", "类型", "大小", "状态", "路径", "备注")
        self.doc_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题和宽度
        self.doc_tree.heading("文件名", text="文件名")
        self.doc_tree.heading("类型", text="类型")
        self.doc_tree.heading("大小", text="大小(MB)")
        self.doc_tree.heading("状态", text="状态")
        self.doc_tree.heading("路径", text="文件路径")
        self.doc_tree.heading("备注", text="备注")
        
        self.doc_tree.column("文件名", width=220)
        self.doc_tree.column("类型", width=70)
        self.doc_tree.column("大小", width=65)
        self.doc_tree.column("状态", width=65)
        self.doc_tree.column("路径", width=280)
        self.doc_tree.column("备注", width=220)
        self.doc_tree.tag_configure("error", foreground="#c00000")
        self.doc_tree.tag_configure("ok", foreground="#1a7f37")
        self.doc_tree.tag_configure("printing", foreground="#0550ae")
        
        # 滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.doc_tree.yview)
        self.doc_tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.doc_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _create_status_area(self):
        """创建状态显示区域"""
        status_frame = ttk.LabelFrame(self.main_frame, text="状态信息", padding="5")
        self.status_frame = status_frame
        
        # 状态标签
        self.lbl_doc_count = ttk.Label(status_frame, text="文档总数: 0")
        self.lbl_total_size = ttk.Label(status_frame, text="总大小: 0 MB")
        self.lbl_printer = ttk.Label(status_frame, text="打印机: 未设置")
        self.lbl_print_status = ttk.Label(status_frame, text="就绪")
        
        # 布局
        self.lbl_doc_count.grid(row=0, column=0, sticky="w", padx=5)
        self.lbl_total_size.grid(row=0, column=1, sticky="w", padx=5)
        self.lbl_printer.grid(row=1, column=0, sticky="w", padx=5)
        self.lbl_print_status.grid(row=1, column=1, sticky="w", padx=5)
    
    def _create_progress_area(self):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(self.main_frame, text="打印进度", padding="5")
        self.progress_frame = progress_frame
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_frame, mode="determinate", length=400
        )
        
        # 进度标签
        self.lbl_progress = ttk.Label(progress_frame, text="等待开始...")
        
        # 布局
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=5)
        self.lbl_progress.grid(row=1, column=0, sticky="w", padx=5)
        
        progress_frame.grid_columnconfigure(0, weight=1)
    
    def _setup_layout(self):
        """设置布局"""
        # 工具栏布局 - 统一风格和间距
        self.btn_add_files.pack(side="left", padx=3)
        self.btn_add_folder.pack(side="left", padx=3)
        self.btn_remove_selected.pack(side="left", padx=3)
        self.btn_filter.pack(side="left", padx=3)
        self.btn_clear.pack(side="left", padx=3)
        self.btn_print_settings.pack(side="left", padx=3)
        self.btn_help.pack(side="left", padx=3)
        self.btn_env.pack(side="left", padx=3)
        self.btn_about.pack(side="left", padx=3)
        
        # 右侧按钮 - 统一间距
        self.btn_start_print.pack(side="right", padx=3)
        self.btn_calculate_pages.pack(side="right", padx=3)
        
        # 主框架布局
        self.toolbar.pack(fill="x", padx=10, pady=5)
        self.env_bar.pack(fill="x", padx=10)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 主内容布局
        self.list_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 5))
        self.status_frame.grid(row=1, column=0, sticky="ew", padx=(0, 5))
        self.progress_frame.grid(row=1, column=1, sticky="ew")
        
        # 设置权重
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
    
    def _bind_events(self):
        """绑定事件"""
        # 确保列表操作处理器已初始化
        if self.list_operation_handler:
            # 设置列排序功能
            self.list_operation_handler.setup_column_sorting()
        
        # 设置拖拽功能
        self.file_import_handler.setup_drag_drop(self.doc_tree)
        
        # 创建右键菜单
        self._create_context_menu()
        
        # 文档列表事件
        self.doc_tree.bind("<Button-3>", self._show_context_menu)
        self.doc_tree.bind("<Double-1>", self._on_double_click)
        self.doc_tree.bind("<Delete>", self._on_delete_key)
    
    def _setup_tooltips(self):
        """设置界面提示功能"""
        # 只为文件类型过滤器添加tooltip
        create_button_tooltip(self.chk_word, 
            "文字文档过滤器（WPS 文字 / Word）\n支持格式: .doc, .docx, .wps")
        
        create_button_tooltip(self.chk_ppt, 
            "演示文稿过滤器（WPS 演示 / PowerPoint）\n支持格式: .ppt, .pptx, .dps")
        
        create_button_tooltip(self.chk_excel, 
            "表格过滤器（WPS 表格 / Excel）\n支持格式: .xls, .xlsx, .et\n打印整个工作簿的所有可见工作表，遵循各表的打印区域和页面设置")
        
        create_button_tooltip(self.chk_pdf, 
            "PDF文档过滤器\n支持标准PDF文件")
        
        create_button_tooltip(self.chk_image, 
            "图片文件过滤器\n支持格式: .jpg, .jpeg, .png, .bmp, .tiff, .tif, .webp\n注意: TIFF可能包含多页，其他图片按1页计算")
        
        create_button_tooltip(self.chk_text, 
            "文本文件过滤器\n支持格式: .txt\n注意:此格式只能模糊统计页数")
        
        # 为工具栏按钮添加tooltip
        create_button_tooltip(self.btn_filter, 
            "文件过滤\n根据上方勾选的文件类型过滤列表\n只保留勾选类型的文档，移除未勾选类型的文档")
        

    

    
    # === 文件操作相关方法 ===
    def _add_files(self):
        """添加文件"""
        added_count = self.file_import_handler.add_files_dialog()
        if added_count > 0:
            self._refresh_document_list(force_rebuild=True)  # 添加文件需要重建列表
            self._update_status()
    
    def _add_folder(self):
        """添加文件夹"""
        added_count = self.file_import_handler.add_folder_dialog()
        if added_count > 0:
            self._refresh_document_list(force_rebuild=True)  # 添加文件需要重建列表
            self._update_status()
    
    def _remove_selected_documents(self):
        """删除选中的文档"""
        if not self.list_operation_handler:
            return
        removed_count = self.list_operation_handler.remove_selected_documents()
        if removed_count > 0:
            self._refresh_document_list(force_rebuild=True)  # 删除文件需要重建列表
            self._update_status()
    
    def _clear_documents(self):
        """清空文档列表"""
        if not self.list_operation_handler:
            return
        if self.list_operation_handler.clear_all_documents():
            self._refresh_document_list(force_rebuild=True)  # 清空列表需要重建
            self._update_status()
    
    def _filter_documents(self):
        """根据勾选的文件类型过滤文档列表"""
        if not self.list_operation_handler:
            return
        
        # 获取当前启用的文件类型
        enabled_types = self._get_enabled_file_types()
        
        # 执行过滤操作
        filtered_count = self.list_operation_handler.filter_documents_by_enabled_types(enabled_types)
        
        if filtered_count > 0:
            self._refresh_document_list(force_rebuild=True)  # 过滤文件需要重建列表
            self._update_status()
    
    # === 文件类型过滤器相关 ===
    def _get_enabled_file_types(self) -> dict:
        """获取当前启用的文件类型"""
        return {
            'word': self.var_word.get(),
            'ppt': self.var_ppt.get(),
            'excel': self.var_excel.get(),
            'pdf': self.var_pdf.get(),
            'image': self.var_image.get(),
            'text': self.var_text.get()
        }
    
    def _on_filter_changed(self):
        """文件类型过滤器变更事件"""
        enabled_types = self._get_enabled_file_types()
        self.window_manager.save_user_preferences(self.app_config, enabled_types)
        print(f"文件类型过滤器已更新: {enabled_types}")
    
    def _on_files_imported(self):
        """文件导入完成后的回调函数"""
        self._refresh_document_list(force_rebuild=True)  # 导入文件需要重建列表
        self._update_status()
    
    def _on_list_operation_completed(self):
        """列表操作完成后的回调函数"""
        self._refresh_document_list(force_rebuild=True)  # 列表操作完成需要重建列表
        self._update_status()
    
    # === 打印相关方法 ===
    def _show_print_settings(self):
        """显示打印设置对话框"""
        dialog = PrintSettingsDialog(
            self.root, 
            self.printer_manager, 
            self.current_print_settings
        )
        
        if dialog.result:
            self.current_print_settings = dialog.result
            self.print_controller.set_print_settings(self.current_print_settings)
            self.config_manager.save_print_settings(self.current_print_settings)
            self._update_status()
            print("打印设置已更新")
    
    def _start_printing(self):
        """开始批量打印 / 打印中点击则停止"""
        if self.print_controller.is_printing:
            self._stop_printing()
            return
        if self.document_manager.document_count == 0:
            messagebox.showwarning("提示", "请先添加要打印的文档")
            return
        self._launch_print(self.document_manager.documents)

    def _stop_printing(self):
        if messagebox.askyesno("停止打印", "当前文件打印完成后停止，剩余文件标记为“已取消”。\n确定停止吗？"):
            self.print_controller.cancel_current_print()
            self.btn_start_print.config(text="正在停止...", state="disabled")
            self.lbl_print_status.config(text="正在停止...")

    def _ensure_printer_selected(self) -> bool:
        """检查是否已设置打印机，没有则引导用户设置"""
        if self.current_print_settings.printer_name:
            return True
        if messagebox.askyesno("需要设置打印机", "检测到尚未设置打印机，是否现在进行设置？"):
            self._show_print_settings()
            if self.current_print_settings.printer_name:
                return True
            messagebox.showinfo("取消打印", "未完成打印机设置，打印操作已取消")
        else:
            messagebox.showinfo("取消打印", "需要设置打印机才能开始打印")
        return False

    def _office_precheck(self, documents: List[Document]) -> bool:
        """Office 组件与正在运行的 WPS 检查，返回是否继续"""
        office_docs = [d for d in documents if d.file_type in OFFICE_FILE_TYPES]
        if not office_docs:
            return True
        report = self.env_report
        if report is not None and report.selection is not None:
            missing = []
            for kind in {KIND_BY_FILE_TYPE[d.file_type] for d in office_docs}:
                backend, message = report.selection.backend_for(kind)
                if backend is None:
                    missing.append(message)
            if missing:
                if not messagebox.askyesno(
                        "Office 组件不可用",
                        "\n\n".join(missing) + "\n\n这些文件将打印失败（其它文件不受影响）。是否继续？"):
                    return False
        running = process_utils.list_running({"wps.exe", "et.exe", "wpp.exe"})
        if running:
            return messagebox.askyesno(
                "WPS 正在运行",
                "检测到 WPS 正在运行：\n" + "\n".join(running) +
                "\n\n为避免批量打印影响您正在编辑的文档，并确保卡死时可以自动恢复，"
                "建议先保存并关闭 WPS。\n\n是否仍然继续打印？")
        return True

    def _launch_print(self, documents: List[Document]):
        if not self._ensure_printer_selected():
            return
        if not self._office_precheck(documents):
            return
        count = len(documents)
        s = self.current_print_settings
        detail = (f"打印机：{s.printer_name}\n纸张：{s.paper_size}  方向：{s.orientation_display}\n"
                  f"{s.duplex_display}  {s.color_display}  份数：{s.copies}")
        if not messagebox.askyesno("确认", f"确定要打印 {count} 个文档吗？\n\n{detail}"):
            return
        try:
            self.print_controller.clear_queue()
            self.print_controller.add_documents_to_queue(documents)
            self.print_controller.start_batch_print()
        except Exception as e:  # noqa: BLE001 - 显示给用户
            messagebox.showerror("错误", f"启动打印失败: {e}")
            return
        self.progress_bar['value'] = 0
        self.btn_start_print.config(text="停止打印")
        self.btn_calculate_pages.config(state="disabled")
        self.lbl_print_status.config(text=f"打印中 (0/{count})")
        print(f"批量打印任务已启动，共 {count} 个文档")

    def _add_initial_paths(self, paths: List[str]):
        added = self.file_import_handler.process_dropped_paths(paths)
        if added:
            self._refresh_document_list(force_rebuild=True)
            self._update_status()

    # === 后台线程消息处理（主线程） ===
    def _process_ui_queue(self):
        try:
            while True:
                kind, payload = self._ui_queue.get_nowait()
                try:
                    if kind == "progress":
                        self._on_print_progress(*payload)
                    elif kind == "document":
                        self._update_document_row(payload)
                    elif kind == "finished":
                        self._on_print_finished(payload)
                    elif kind == "environment":
                        self._apply_environment_report(payload)
                except Exception as e:  # noqa: BLE001 - 单条消息异常不影响后续
                    print(f"界面更新失败: {e}")
        except queue.Empty:
            pass
        try:
            self.root.after(UI_POLL_MS, self._process_ui_queue)
        except tk.TclError:
            pass  # 窗口已销毁

    def _on_print_progress(self, current: int, total: int, message: str):
        """打印进度（主线程）"""
        progress = (current / total) * 100 if total else 100
        self.progress_bar['value'] = progress
        self.lbl_progress.config(text=f"{current}/{total} - {message}")
        if self.print_controller.is_printing:
            self.lbl_print_status.config(text=f"打印中 ({current}/{total})")

    def _on_print_finished(self, summary):
        """批次结束（主线程）"""
        self.btn_start_print.config(text="开始打印", state="normal")
        self._update_status()
        self._refresh_document_list()
        self.progress_bar['value'] = 100
        self.lbl_print_status.config(text=f"打印完成：{summary.text()}")
        self.lbl_progress.config(text=f"{summary.total}/{summary.total} - 批量打印完成！{summary.text()}")
        lines = [f"批量打印完成：{summary.text()}（共 {summary.total} 个）"]
        if summary.fatal_error:
            lines.append(f"\n错误：{summary.fatal_error}")
        failed = [r for r in summary.records if not r.success]
        if failed:
            lines.append("\n失败的文件：")
            for record in failed[:10]:
                lines.append(f"• {Path(record.file).name}：{record.error}")
            if len(failed) > 10:
                lines.append(f"……另有 {len(failed) - 10} 个")
        if summary.warnings:
            lines.append("\n打印机设置提示：")
            lines.extend(f"• {w}" for w in summary.warnings)
        lines.append(f"\n打印日志：{summary.log_file}")
        if summary.failed or summary.fatal_error:
            messagebox.showwarning("打印完成（有失败）", "\n".join(lines))
        else:
            messagebox.showinfo("打印完成", "\n".join(lines))

    # === 环境检测 ===
    def _start_environment_check(self):
        preference = self.current_print_settings.office_backend.value

        def worker():
            try:
                report = collect_environment(preference)
                print(report.text())
                self._ui_queue.put(("environment", report))
            except Exception as e:  # noqa: BLE001
                print(f"环境检测失败: {e}")

        threading.Thread(target=worker, name="EnvCheck", daemon=True).start()

    def _apply_environment_report(self, report: EnvironmentReport):
        self.env_report = report
        self.lbl_backend.config(text=f"Office Backend: {report.backend_label()}")
        if report.wps:
            self.lbl_wps_components.config(text="  ".join(s.summary() for s in report.wps.values()))
        self.lbl_sumatra.config(text=f"SumatraPDF {'✓' if report.sumatra_ok else '✗'}")
        problems = report.problems()
        if problems:
            first = problems[0] if len(problems[0]) <= 36 else problems[0][:35] + "…"
            more = f" 等 {len(problems)} 项" if len(problems) > 1 else ""
            self.lbl_env_warning.config(text=f"⚠ {first}{more}（点击查看诊断）")
        else:
            self.lbl_env_warning.config(text="")

    def _show_environment(self):
        if self.env_report is None:
            messagebox.showinfo("环境诊断", "正在检测运行环境，请稍候再试")
            return
        show_environment_dialog(self.root, self.env_report,
                                self.current_print_settings.office_backend.value,
                                on_updated=self._apply_environment_report)
    
    # === 页数统计相关 ===
    def _calculate_pages(self):
        """计算页数"""
        if self.document_manager.document_count == 0:
            messagebox.showwarning("提示", "请先添加要统计的文档")
            return
        if self.print_controller.is_printing:
            messagebox.showwarning("提示", "正在打印，请等待打印完成后再统计页数")
            return
        
        # 显示页数统计对话框
        show_page_count_dialog(self.root, self.document_manager.documents)
    
    def _calculate_selected_pages(self):
        """计算选中文档的页数"""
        if not self.list_operation_handler:
            return
        selected_documents = self.list_operation_handler.get_selected_document_objects()
        if not selected_documents:
            messagebox.showwarning("提示", "请先选择要计算页数的文档")
            return
        if self.print_controller.is_printing:
            messagebox.showwarning("提示", "正在打印，请等待打印完成后再统计页数")
            return
        
        # 调用页数统计功能
        try:
            show_page_count_dialog(self.root, selected_documents)
        except Exception as e:
            messagebox.showerror("错误", f"页数统计功能出错: {e}")
    
    # === 界面更新相关方法 ===
    def _refresh_document_list(self, force_rebuild: bool = False):
        """
        刷新文档列表显示
        
        Args:
            force_rebuild: 是否强制重建整个列表（默认为智能更新）
        """
        if force_rebuild or len(self.doc_tree.get_children()) != len(self.document_manager.documents):
            # 强制重建或列表长度不匹配时，完全重建列表
            self._rebuild_document_list()
        else:
            # 智能更新：只更新状态发生变化的项目
            self._update_document_list_status()
    
    @staticmethod
    def _row_values(doc: Document):
        return (doc.file_name, doc.type_display, doc.size_mb, doc.status_display,
                str(doc.file_path), doc.last_error)

    @staticmethod
    def _row_tags(doc: Document):
        return {PrintStatus.ERROR: ("error",), PrintStatus.COMPLETED: ("ok",),
                PrintStatus.PRINTING: ("printing",)}.get(doc.print_status, ())

    def _rebuild_document_list(self):
        """完全重建文档列表（行 iid 使用文档 id，便于按文档更新状态）"""
        selected_ids = set(self.doc_tree.selection())
        for item in self.doc_tree.get_children():
            self.doc_tree.delete(item)
        for doc in self.document_manager.documents:
            self.doc_tree.insert("", "end", iid=doc.id, values=self._row_values(doc), tags=self._row_tags(doc))
            if doc.id in selected_ids:
                self.doc_tree.selection_add(doc.id)
        if self.list_operation_handler:
            self.list_operation_handler.maintain_sort_indicators()

    def _update_document_row(self, doc: Document):
        """更新单个文档行（主线程）"""
        if self.doc_tree.exists(doc.id):
            self.doc_tree.item(doc.id, values=self._row_values(doc), tags=self._row_tags(doc))
            if doc.print_status == PrintStatus.PRINTING:
                self.doc_tree.see(doc.id)

    def _update_document_list_status(self):
        """智能更新文档列表状态"""
        documents = self.document_manager.documents
        if len(self.doc_tree.get_children()) != len(documents):
            self._rebuild_document_list()
            return
        for doc in documents:
            self._update_document_row(doc)
    
    def _update_status(self):
        """更新状态显示"""
        summary = self.document_manager.get_summary()
        
        # 更新文档统计
        self.lbl_doc_count.config(text=f"文档总数: {summary['total']}")
        self.lbl_total_size.config(text=f"总大小: {summary['total_size_mb']} MB")
        
        # 更新打印机信息
        printer_name = self.current_print_settings.printer_name or "未设置"
        self.lbl_printer.config(text=f"打印机: {printer_name}")
        
        # 更新按钮状态
        if self.document_manager.document_count == 0:
            self.btn_calculate_pages.config(state="disabled")
        else:
            self.btn_calculate_pages.config(state="normal")
    
    # === 右键菜单相关 ===
    def _create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
    
    def _show_context_menu(self, event):
        """显示右键菜单"""
        selection = self.doc_tree.selection()
        if not selection:
            return
        
        # 清空现有菜单项
        self.context_menu.delete(0, "end")
        
        # 根据选中项目数量构建菜单
        selected_count = len(selection)
        
        if selected_count == 1:
            # 单文件菜单
            self.context_menu.add_command(
                label="📁  打开所在文件夹",
                command=lambda: self.list_operation_handler.open_file_location() if self.list_operation_handler else None
            )
            self.context_menu.add_command(
                label="📄  仅打印选中文档",
                command=self._print_selected_documents
            )
            self.context_menu.add_command(
                label="❌  从列表中移除",
                command=self._remove_selected_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label="🔄  重置排序",
                command=self._reset_sort
            )
            self.context_menu.add_command(
                label="🔍  文件过滤",
                command=self._filter_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label="📊  计算选中文档页数",
                command=self._calculate_selected_pages
            )
        else:
            # 多文件菜单
            self.context_menu.add_command(
                label=f"📄  仅打印选中文档 ({selected_count}个文件)",
                command=self._print_selected_documents
            )
            self.context_menu.add_command(
                label=f"❌  从列表中移除 ({selected_count}个文件)",
                command=self._remove_selected_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label="🔄  重置排序",
                command=self._reset_sort
            )
            self.context_menu.add_command(
                label="🔍  文件过滤",
                command=self._filter_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label=f"📊  计算选中文档页数 ({selected_count}个文件)",
                command=self._calculate_selected_pages
            )
            self.context_menu.add_command(
                label=f"💾  导出选中文档列表 ({selected_count}个文件)",
                command=lambda: self.list_operation_handler.export_document_list("selected") if self.list_operation_handler else None
            )
        
        # 显示菜单
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def _print_selected_documents(self):
        """仅打印选中的文档"""
        if not self.list_operation_handler:
            return
        selected_documents = self.list_operation_handler.get_selected_document_objects()
        if not selected_documents:
            messagebox.showwarning("提示", "请先选择要打印的文档")
            return
        if self.print_controller.is_printing:
            messagebox.showwarning("提示", "打印任务正在进行中")
            return
        self._launch_print(selected_documents)
    
    def _reset_sort(self):
        """重置排序"""
        if not self.list_operation_handler:
            return
        self.list_operation_handler.reset_sort()
        self._refresh_document_list(force_rebuild=True)  # 重置排序需要重建列表
    
    # === 事件处理 ===
    def _on_double_click(self, event):
        """双击文档列表项处理"""
        # 检查是否点击在列标题区域
        region = self.doc_tree.identify_region(event.x, event.y)
        if region == "heading":
            # 点击在列标题区域，不处理双击事件
            return
        
        # 检查是否有选中的项目
        item = self.doc_tree.identify_row(event.y)
        if item and self.list_operation_handler:
            self.list_operation_handler.open_file_with_default_app()
    
    def _on_delete_key(self, event):
        """删除键处理"""
        self._remove_selected_documents()
    
    def _on_window_closing(self):
        """窗口关闭事件处理"""
        if self.print_controller.is_printing:
            if not messagebox.askyesno("确认", "打印任务正在进行中，退出将中止当前任务。\n确定要退出吗？"):
                return False
            # 立即结束本程序启动的 WPS 进程，工作线程随后会恢复默认打印机与打印参数
            self.print_controller.cancel_current_print(force=True)
        self.window_manager.save_window_geometry(self.app_config)
        return True
    
    # === 使用说明 ===
    def _show_help(self):
        """显示使用说明"""
        help_text = f"""
📖 {APP_NAME} v{VERSION} 使用说明

═══════════════════════════════════════

🎯 软件定位
Windows + WPS Office 批量打印工具。只安装 WPS Office、未安装 Microsoft Office
也能批量打印 Office 文档；PDF 与图片无需 WPS/Office。

═══════════════════════════════════════

📂 支持的文件格式

• 文字文档：.doc .docx .wps        → WPS 文字（KWPS.Application）
• 表格：    .xls .xlsx .et          → WPS 表格（KET.Application）
• 演示文稿：.ppt .pptx .dps        → WPS 演示（KWPP.Application）
• PDF：     .pdf                    → 内置 SumatraPDF
• 图片：    .jpg .jpeg .png .bmp .tiff .tif .webp → SumatraPDF（备用 GDI）
• 文本：    .txt                    → Windows GDI 直接打印

未安装 WPS 时，可在“打印设置 → Office 引擎”中改用 Microsoft Office。

═══════════════════════════════════════

📋 使用步骤

1. 拖入文件或文件夹（或点击“添加文件/添加文件夹”），自动识别类型
2. 点击“打印设置”：选择打印机、份数、纸张、单双面、方向、彩色/黑白
3. 点击“开始打印”，列表中实时显示 等待 / 正在打印 / 成功 / 失败
4. 打印完成后显示“成功 N，失败 M”，失败原因显示在“备注”列
5. 打印中可点击“停止打印”，当前文件完成后停止

═══════════════════════════════════════

⚙️ 打印参数说明

• 纸张/方向：PDF、图片、文本直接生效；文字/表格文档默认遵循文档自身页面设置，
  勾选“Office 文档也强制使用以上纸张和方向”后才会覆盖（不会保存到原文件）
• 单双面/彩色：通过打印机的“每用户默认打印参数”设置，结束后自动恢复
• 打印机驱动不支持的设置会使用驱动默认值，并在日志和完成提示中说明
• 打印 Office 文档时，若目标打印机不是系统默认打印机，程序会临时将其设为默认，
  打印结束（包括出错、取消）后自动恢复；即使程序意外退出，下次启动也会恢复

═══════════════════════════════════════

🛡️ 稳定性

• 单个文件失败不影响后续文件
• WPS 卡死超时（文字 3 分钟、表格/演示 5 分钟）会被强制结束并继续下一个文件
• 只结束本程序启动的 WPS 进程，不会影响您自己打开的 WPS
• 建议批量打印前保存并关闭正在编辑的 WPS 文档

═══════════════════════════════════════

🔍 故障排查

• 点击“环境诊断”查看 WPS 文字/表格/演示 COM 是否可用，可做“深度检测”
• 日志目录：{get_logs_dir()}
  - app.log：运行日志
  - print_YYYYMMDD.log：每个打印任务的记录（文件、打印机、参数、结果、耗时）

═══════════════════════════════════════

基于 batch-document-printer（作者 喵言喵语，MIT 许可）二次开发。
        """
        
        # 创建帮助窗口
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("650x700")
        help_window.resizable(True, True)
        help_window.transient(self.root)
        help_window.grab_set()
        
        # 居中显示
        self.window_manager.center_window(help_window)
        
        # 创建滚动文本框
        main_frame = ttk.Frame(help_window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 文本显示区域
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True)
        
        # 创建文本控件和滚动条
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Microsoft YaHei", 10),
            bg="white",
            fg="black",
            relief="flat",
            borderwidth=0,
            state="normal"
        )
        
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 插入帮助文本
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")  # 设为只读
        
        # 关闭按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        close_btn = ttk.Button(
            button_frame,
            text="关闭",
            command=help_window.destroy,
            width=10
        )
        close_btn.pack(side="right")
    
    def run(self):
        """运行主循环"""
        # 初始更新
        self._update_status()
        
        # 启动主循环
        self.root.mainloop()

        # 窗口关闭后等待打印线程结束：确保 WPS 进程已退出、默认打印机与打印参数已恢复
        self.print_controller.shutdown(wait=True)


def main():
    """主函数"""
    try:
        app = MainWindow()
        app.run()
    except Exception as e:
        print(f"应用程序启动失败: {e}")
        messagebox.showerror("错误", f"应用程序启动失败: {e}")


if __name__ == "__main__":
    main() 