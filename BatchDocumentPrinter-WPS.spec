# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置（onedir，稳定优先）

产物: dist/BatchDocumentPrinter-WPS/
    BatchDocumentPrinter-WPS.exe   图形界面主程序
    BDP-WPS-Diagnose.exe           控制台诊断工具（WPS COM 冒烟测试，不打印）
    _internal/                     运行库、SumatraPDF、图标等

构建: pyinstaller --noconfirm --clean BatchDocumentPrinter-WPS.spec
"""
from PyInstaller.utils.hooks import collect_all

block_cipher = None

datas = [
    ('external/SumatraPDF/SumatraPDF.exe', 'external/SumatraPDF'),
    ('external/SumatraPDF/SumatraPDF-settings.txt', 'external/SumatraPDF'),
    ('resources/app_icon.ico', 'resources'),
]
binaries = []
hiddenimports = [
    'win32timezone', 'win32print', 'win32api', 'win32con', 'win32ui', 'win32file',
    'win32com', 'win32com.client', 'win32com.client.dynamic', 'pythoncom', 'pywintypes',
    'psutil', 'PIL.ImageWin', 'PIL.ImageSequence', 'PyPDF2', 'pptx', 'openpyxl',
    'tkinterdnd2',
]
for package in ('tkinterdnd2',):
    d, b, h = collect_all(package)
    datas += d
    binaries += b
    hiddenimports += h

excludes = ['xlwings', 'comtypes', 'numpy', 'pandas', 'matplotlib', 'scipy', 'IPython',
            'pytest', 'docx', 'lxml.html.clean']

gui = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=excludes,
    noarchive=False,
    cipher=block_cipher,
)
diag = Analysis(
    ['scripts/test_wps_com.py'],
    pathex=['.'],
    hiddenimports=['win32timezone', 'win32com.client.dynamic', 'pythoncom', 'pywintypes', 'psutil'],
    excludes=excludes,
    noarchive=False,
    cipher=block_cipher,
)

gui_pyz = PYZ(gui.pure, gui.zipped_data, cipher=block_cipher)
diag_pyz = PYZ(diag.pure, diag.zipped_data, cipher=block_cipher)

gui_exe = EXE(
    gui_pyz, gui.scripts, [],
    exclude_binaries=True,
    name='BatchDocumentPrinter-WPS',
    debug=False,
    strip=False,
    upx=False,            # 不压缩：减少杀毒软件误报，提高稳定性
    console=False,
    icon='resources/app_icon.ico',
    version='version_info.txt',
)
diag_exe = EXE(
    diag_pyz, diag.scripts, [],
    exclude_binaries=True,
    name='BDP-WPS-Diagnose',
    debug=False,
    strip=False,
    upx=False,
    console=True,
    icon='resources/app_icon.ico',
    version='version_info.txt',
)

coll = COLLECT(
    gui_exe, gui.binaries, gui.zipfiles, gui.datas,
    diag_exe, diag.binaries, diag.zipfiles, diag.datas,
    strip=False,
    upx=False,
    name='BatchDocumentPrinter-WPS',
)
