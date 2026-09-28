# Batch Document Printer – WPS Edition v1.0.0

Windows + WPS Office 批量打印工具。只安装 WPS Office、不安装 Microsoft Office，也能批量打印
文字、表格、演示、PDF、图片和文本。基于 [batch-document-printer](https://github.com/icescat/batch-document-printer)
（作者 喵言喵语，MIT）二次开发。

## 下载与安装

1. 下载 `BatchDocumentPrinter-WPS-v1.0.0.zip`，可用 `SHA256SUMS.txt` 校验：
   `certutil -hashfile BatchDocumentPrinter-WPS-v1.0.0.zip SHA256`
2. 解压到任意目录（建议不要放在 `C:\Program Files`）
3. 运行 `BatchDocumentPrinter-WPS.exe`；如需检查 WPS COM，运行 `BDP-WPS-Diagnose.exe`

无需安装 Python、pip、Git 或 Visual Studio。

## 系统要求

- Windows 10 / Windows 11，64 位
- WPS Office for Windows（需要注册 COM 接口：`KWPS.Application`、`KET.Application`、`KWPP.Application`；
  WPS 2019 及以后的个人版/专业版正常安装后通常已注册，若未注册请在 WPS 配置工具 → 高级 → 兼容设置 中勾选
  “WPS Office 兼容第三方系统和软件”）
- Microsoft Office 非必需
- PDF、图片、文本打印不需要 WPS/Office

## 支持的格式

| 类型 | 扩展名 | 引擎 |
|------|--------|------|
| PDF | .pdf | SumatraPDF 3.4.6（随附） |
| 文字 | .doc .docx .wps | WPS 文字 KWPS.Application |
| 表格 | .xls .xlsx .et | WPS 表格 KET.Application |
| 演示 | .ppt .pptx .dps | WPS 演示 KWPP.Application |
| 图片 | .jpg .jpeg .png .bmp .tiff .tif .webp | SumatraPDF（备用 GDI） |
| 文本 | .txt | Windows GDI |

## 支持的打印参数

打印机选择 · 份数 1–999 · 纸张 A3/A4/A5/B5/Letter/Legal（可同步打印机纸张）·
方向 自动/纵向/横向 · 单面/双面长边/双面短边 · 彩色/黑白。
驱动不支持的设置不会导致失败，会记录“当前打印机驱动不支持该设置，已使用驱动默认值”。
Office 文档的纸张/方向默认遵循文档自身页面设置，可选择强制套用。

## 验证情况（请务必阅读）

开发与构建环境为 Linux，没有 Windows 桌面和 WPS Office。已完成与未完成的验证如下：

| 项目 | 状态 |
|------|------|
| 单元测试 + mock 端到端测试（111 项，含队列失败隔离、超时、默认打印机恢复、崩溃恢复、Backend 选择） | ✅ 通过（Linux CPython 3.11 与 Windows CPython 3.11 on Wine） |
| PyInstaller 打包产物启动、界面、环境检测、日志、诊断工具 | ✅ 在 Wine 中运行通过（非真实 Windows） |
| 真实 COM 生命周期代码（注册表检测、CoCreateInstance、隐藏、Quit、进程识别、看门狗结束进程） | ✅ 在 Wine 中使用其自带的进程外 COM 服务器验证（非 WPS） |
| PDF / PNG / JPG / BMP / TXT 真实打印链路（真实 pywin32 + SumatraPDF + GDI，打印到 CUPS-PDF 虚拟打印机） | ✅ 在 Wine + CUPS-PDF 中通过；默认打印机切换并恢复（TIFF：SumatraPDF 在 Wine 中崩溃，GDI 备用方案单独验证可用；Windows 上未验证） |
| KWPS / KET / KWPP COM 冒烟测试 | ⚠️ **未在真实 WPS 上验证** |
| DOCX / XLSX / PPTX / WPS / ET / DPS 通过 WPS 打印 | ⚠️ **未在真实 WPS 上验证** |
| Windows 10 / 11 真机运行 | ⚠️ **未验证** |
| 真实打印机的双面 / 彩色 / 纸张效果 | ⚠️ **未验证** |

建议首次使用时：运行 `BDP-WPS-Diagnose.exe` → 在“打印设置”中选择“Microsoft Print to PDF”试打几个文件 → 再使用实体打印机。
开发者可用 `scripts/e2e_real_print.py`（配合文件端口虚拟 PDF 打印机）在 Windows + WPS 上做完整自动验证。

## 已知限制

- 见上表：Office 文档打印路径尚未经过真实 WPS 验证
- WPS 演示的 ActivePrinter 不可写时会在批量打印期间临时更改 Windows 默认打印机（结束后自动恢复）
- 双面、彩色依赖打印机驱动读取每用户默认打印参数，个别驱动/WPS 版本可能不生效
- 启动打印时若 WPS 已在运行，WPS 可能复用该实例：不会被关闭，但超时保护无法强制结束它（程序会提示先关闭 WPS）
- 可执行文件未签名，个别杀毒软件可能误报

## 许可证

MIT（保留原作者 喵言喵语 的版权声明）。随附的 SumatraPDF 为 GPLv3，源码见
https://github.com/sumatrapdfreader/sumatrapdf/tree/3.4.6rel 。详见 LICENSE、NOTICE、licenses/。
