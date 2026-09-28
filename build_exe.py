#!/usr/bin/env python3
"""
Batch Document Printer – WPS Edition 发布构建脚本（需在 Windows 上运行）

步骤：运行单元测试 → PyInstaller(onedir) → 校验打包内容 → 复制文档与许可证
      → 生成 BatchDocumentPrinter-WPS-v<版本>.zip 与 SHA256SUMS.txt

用法:
    pip install -r requirements-dev.txt
    python build_exe.py [--skip-tests]
"""
import argparse
import hashlib
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from src.version import APP_ID, VERSION  # noqa: E402

DIST = ROOT / "dist"
APP_DIR = DIST / APP_ID
ZIP_NAME = f"{APP_ID}-v{VERSION}.zip"

REQUIRED_IN_BUNDLE = [
    f"{APP_ID}.exe",
    "BDP-WPS-Diagnose.exe",
    "_internal/external/SumatraPDF/SumatraPDF.exe",
    "_internal/resources/app_icon.ico",
    "_internal/win32/win32print.pyd",
    "_internal/Pythonwin/win32ui.pyd",
    "_internal/pywin32_system32/pythoncom311.dll",
    "_internal/pywin32_system32/pywintypes311.dll",
    "_internal/tkinterdnd2/tkdnd/win64",
    "_internal/psutil",
    "_internal/tkinterdnd2",
    "LICENSE",
    "NOTICE",
    "README.md",
    "licenses/SumatraPDF-COPYING.txt",
]
DOC_FILES = ["README.md", "LICENSE", "NOTICE", "CHANGELOG.md"]


def run(cmd):
    print("$", " ".join(str(c) for c in cmd), flush=True)
    subprocess.run(cmd, check=True, cwd=ROOT)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-tests", action="store_true")
    args = parser.parse_args()

    if not args.skip_tests:
        run([sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider", "tests"])

    shutil.rmtree(ROOT / "build", ignore_errors=True)
    shutil.rmtree(APP_DIR, ignore_errors=True)
    run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "BatchDocumentPrinter-WPS.spec"])

    for name in DOC_FILES:
        if (ROOT / name).exists():
            shutil.copy2(ROOT / name, APP_DIR / name)
    shutil.copytree(ROOT / "licenses", APP_DIR / "licenses", dirs_exist_ok=True)
    shutil.copy2(ROOT / "docs" / "RELEASE_NOTES.md", APP_DIR / "RELEASE_NOTES.md")

    missing = [item for item in REQUIRED_IN_BUNDLE if not (APP_DIR / item).exists()]
    if missing:
        print("打包内容缺失:", missing)
        return 1
    print("打包内容校验通过")

    zip_path = DIST / ZIP_NAME
    zip_path.unlink(missing_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for file in sorted(APP_DIR.rglob("*")):
            if file.is_file():
                zf.write(file, Path(APP_ID) / file.relative_to(APP_DIR))
    sums = DIST / "SHA256SUMS.txt"
    sums.write_text(f"{sha256(zip_path)}  {ZIP_NAME}\n", encoding="utf-8")
    print(f"生成: {zip_path} ({zip_path.stat().st_size / 1024 / 1024:.1f} MB)")
    print(sums.read_text(encoding="utf-8"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
