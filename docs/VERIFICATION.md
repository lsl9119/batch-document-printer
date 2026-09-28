# 验证记录 – v1.0.0

开发/构建环境：Linux（Ubuntu 24.04）容器，**没有 Windows 桌面、没有 WPS Office、没有 Microsoft Office**。
因此本版本的验证分为三类，请勿混淆：

- **A. 自动化测试（mock）**：pytest，替换 COM、win32print、SumatraPDF 子进程、进程表
- **B. Wine 环境真实代码验证**：Windows 版 CPython 3.11 + 真实 pywin32 / psutil / PyInstaller 产物运行在 Wine 9.0 中，
  打印机为 CUPS-PDF 虚拟打印机。这验证的是本程序代码与 Windows API 的交互，**不是** Windows 10/11 真机，也**不是** WPS
- **C. 真实 Windows + WPS 验证**：**尚未进行**

## A. 自动化测试

`python -m pytest -q tests` → 111 passed（Linux CPython 3.11；Windows CPython 3.11 on Wine 同样 111 passed）

覆盖：文件类型识别（16 种扩展名）、PrintSettings 兼容旧配置、Backend 检测与回退（WPS 缺表格组件时的提示含
`ProgID: KET.Application`）、COM 错误分类与降级调用、会话池复用/回收/失效重启、看门狗、只结束自有进程、
WPS 文字/表格/演示的打开-打印-关闭（命名参数不支持时的降级、加密文件、空工作簿、打印区域、隐藏表、
ActivePrinter 只读）、打印队列（单个失败不中断：成功 4 失败 1、超时、取消、异常隔离、日志字段）、
默认打印机切换与恢复（含异常路径、崩溃恢复、多实例互斥）、每用户打印参数写入与恢复、驱动不支持时的降级提示、
临时副本（长路径/只读/被占用）、SumatraPDF 参数、外部进程不因继承管道而卡死、文本排版、页数统计。

mock 端到端测试（`tests/test_e2e_mocked.py`）：12 个文件（docx/xlsx/pptx/pdf/png/txt/wps/et/dps/jpg/bmp/tiff）
→ docx/wps 走 `KWPS.Application`、xlsx/et 走 `KET.Application`、pptx/dps 走 `KWPP.Application`，
每个组件只启动一次，全部成功；批次后 WPS 进程全部退出、默认打印机恢复、状态文件删除。

## B. Wine 环境验证（真实代码，非 WPS、非 Windows 真机）

| 项目 | 结果 |
|------|------|
| PyInstaller onedir 打包（Windows CPython 3.11 + PyInstaller 6.11.1） | 成功；打包内容校验通过（SumatraPDF、pywin32 DLL、win32ui、tkdnd） |
| `BatchDocumentPrinter-WPS.exe` 启动 | 界面正常、环境检测正常（正确报告 KWPS/KET/KWPP 未注册）、日志写入程序目录、命令行传入文件夹可加入列表 |
| `BDP-WPS-Diagnose.exe` | 正常运行并报告 3 个组件 FAIL（无 WPS，符合预期）；修复了两个仅在打包/真实 API 下出现的问题（见下） |
| 真实 COM 生命周期（用 Wine 自带的进程外 COM 服务器 `InternetExplorer.Application` 代替 WPS） | 注册表 ProgID→CLSID→LocalServer32 检测、CoCreateInstance、Visible=False、Quit、进程识别、看门狗强制结束进程、结束后 COM 调用返回“RPC 服务器不可用”并被识别 —— 全部通过 |
| 真实打印链路（PrintController + SumatraPDF + GDI → CUPS-PDF），打印机 ≠ 默认打印机，份数 2、横向、双面长边、黑白 | PDF/PNG/JPG/BMP 由 SumatraPDF 打印成功；TIFF 在 Wine 中导致 SumatraPDF 崩溃 → 超时后结束进程并报告失败（不自动改用 GDI 重打，避免重复打印；单独调用 GDI 图片打印已验证可用）；TXT（UTF-8、GBK）GDI 打印成功（横向、2 份）；DOCX 因无 WPS 失败并给出含 ProgID 的明确提示，队列继续；批次“成功 6，失败 2”，失败项均符合预期 |
| 默认打印机 | 批次开始切换为目标打印机，结束后恢复为原默认打印机，恢复状态文件已删除 |
| 驱动不支持双面 | 记录“双面打印：当前打印机驱动不支持该设置，已使用驱动默认值” |
| GDI 中文文本 | 使用 TrueType 中文字体时正确输出；程序会核对字体是否真实存在并依次回退 |

独立代码审查（单独的审查代理）发现并已修复：复用用户已打开的 WPS 时会隐藏其窗口/关闭其文档；
`Quit` 卡住时默认打印机无法恢复；恢复状态文件可能被删除或被覆盖；启动失败的 WPS 进程残留；
深度检测/页数统计与打印并发驱动 WPS；PID 复用导致误判；日志重定向可能死锁；PDF 超时与看门狗不一致；
SumatraPDF 超时后 GDI 重打导致重复打印；`SetDefaultPrinter` 可能关闭“让 Windows 管理默认打印机”。

Wine 验证中发现并已修复的真实问题：
1. `pythoncom.CLSIDFromProgID` 在 pywin32 中不存在（改用 `pywintypes.IID`）—— 仅靠 mock 测试无法发现
2. `psutil.Process.name()` 对部分进程抛出非 psutil 异常，导致进程快照中断 —— 改用 Toolhelp32 快照
3. `subprocess.run(capture_output=True)` 在外部程序的子进程继承管道句柄时无限等待（超时无效）
4. 诊断工具在非中文代码页控制台输出中文时崩溃
5. pywin32 的 `SetPrinter(level 9)` 不接受 `pDevMode=None` —— 改用 `SetPrinterW`，失败时逐字段恢复
6. GDI 字体不存在时被静默替换导致中文变成方框 —— 增加字体存在性核对与回退

Wine 的局限（不代表 Windows 行为）：强制结束进程外 COM 服务器后 Wine 无法再次创建该服务器（Windows 的 SCM 会清理）；
Wine PostScript 驱动不能嵌入 OpenType(CFF) 字体；CUPS-PDF 忽略份数与色彩；Wine 不支持以 NULL 清除每用户 DEVMODE。

## C. 未验证项（需要在 Windows 10/11 + WPS 上完成）

- KWPS / KET / KWPP 的 COM 冒烟测试（`BDP-WPS-Diagnose.exe`）
- DOCX / XLSX / PPTX 以及原生 WPS / ET / DPS 通过 WPS 打印（`scripts/e2e_real_print.py --native`）
- WPS 各版本对命名参数、`PrintOut` 参数、`PrintOptions.ActivePrinter` 的实际支持情况
- 批量打印后 wps.exe / et.exe / wpp.exe 无残留（真实 WPS）
- 真实打印机的纸张、方向、双面、彩色、份数效果
- Windows 10/11 真机上的界面、拖拽、默认打印机恢复、`SetPrinterW` 清除每用户 DEVMODE

推荐的验证命令见 README “开发”一节；`.github/workflows/wps-e2e.yml` 可在 GitHub Windows runner 上安装 WPS 后自动执行
（Windows Server 环境，结果仅供参考）。
