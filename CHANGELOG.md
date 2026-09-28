# 更新日志

## v1.0.0 – Batch Document Printer – WPS Edition（2026-09-28）

基于 batch-document-printer v5.2（作者 喵言喵语）二次开发的首个 WPS 版本。

### 新增
- Office Backend 层（`src/backends/`）：WPS Office（`KWPS.Application` / `KET.Application` /
  `KWPP.Application`）优先，Microsoft Office 作为可选回退；统一的 `detect_backend()`、
  `open_writer()` / `open_spreadsheet()` / `open_presentation()`
- 表格改为直接通过 `KET.Application` 打印，移除 xlwings；支持多工作表、隐藏表、空表、打印区域
- WPS 演示：`PrintOptions.ActivePrinter` 不可写时临时切换 Windows 默认打印机并在 `finally` 中恢复；
  状态文件 + 下次启动自动恢复，防止默认打印机被永久修改
- 打印参数：单面 / 双面长边 / 双面短边、自动 / 纵向 / 横向、彩色 / 黑白、纸张、份数；
  通过每用户默认打印参数（`PRINTER_INFO_9`，无需管理员）生效，结束后恢复；驱动不支持时记录并使用驱动默认值
- 启动环境检测：Windows 版本、WPS 三个组件的 COM 注册（ProgID → CLSID → LocalServer32）、Microsoft Office、
  SumatraPDF、打印机列表；“环境诊断”对话框与“深度检测”（实际启动 COM）
- `BDP-WPS-Diagnose.exe` / `scripts/test_wps_com.py`：WPS COM 冒烟测试
- 日志：`logs/app.log`（滚动）与 `logs/print_YYYYMMDD.log`（每个任务一行）
- 批量稳定性：WPS 实例复用与回收、单文件失败不中断、看门狗超时强制结束本程序启动的 WPS 进程、
  WPS 崩溃后自动重启组件、批次结束兜底清理进程
- 占用中 / 网络共享 / 超长路径的文件复制到临时目录后打印
- 打印中可停止；完成后显示“成功 N，失败 M”及失败原因；列表新增“备注”列
- 文件/文件夹拖到 exe 图标上即可加入列表
- TXT 改为 GDI 直接打印（可指定打印机和份数，自动识别 UTF-8/GBK/UTF-16）
- 图片备用打印方案改为 GDI（不再“打开默认程序”冒充打印成功）
- PyInstaller onedir 打包、GitHub Actions（CI / Release / WPS E2E）、pytest 测试（全部 mock）

### 修复（v5.2 中存在的问题）
- PDF 打印未传递份数、纸张、彩色参数给 SumatraPDF
- 双面打印始终为长边翻转，无法选择短边
- 打印机全局设置使用 `SetPrinter` level 2，普通用户无权限时静默失败
- 后台线程直接操作 Tkinter 控件（打印进度、页数统计进度）
- `subprocess.run(capture_output=True)` 在外部程序子进程继承管道时可能无限等待
- 配置保存在源码目录，打包成单文件 exe 后每次启动都会丢失
- PDF 处理器使用相对路径查找 SumatraPDF，从其他目录启动时找不到
- 同名文件（不同文件夹）在选中 / 删除时会互相混淆
- 文件名含“副本”的正常文档在导入文件夹时被跳过
- 打开打印机句柄异常时未关闭
- 同一文件夹导入顺序不确定（现按路径排序）

### 移除
- 依赖：xlwings、comtypes
- 旧的 v5.x 打包配置（改为 `BatchDocumentPrinter-WPS.spec` + `build_exe.py`）
