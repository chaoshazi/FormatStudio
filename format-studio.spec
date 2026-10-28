# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
for package in ("uvicorn", "fastapi", "pydantic", "pdfplumber", "pypdfium2", "reportlab", "docx", "openpyxl", "pptx"):
    hiddenimports += collect_submodules(package)
hiddenimports += [
    "app",
    "app.main",
    "app.capabilities",
    "app.config",
    "app.db",
    "app.formats",
    "app.presets",
    "app.utils",
    "app.services",
    "app.services.converter",
    "app.services.engines",
    "app.services.metadata",
    "app.services.pdf_conv",
    "app.services.pdf_render",
    "app.services.styles_service",
    "app.services.table_conv",
    "app.services.task_manager",
    "app.services.template_service",
    "app.services.text_conv",
]
hiddenimports += ["win32com.client", "pythoncom", "webview", "webview.platforms.edgechromium"]

datas = [
    ("web/dist", "web/dist"),
]

binaries = []
if __import__("pathlib").Path("vendor/libreoffice").exists():
    datas.append(("vendor/libreoffice", "vendor/libreoffice"))

a = Analysis(
    ["desktop.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "pandas"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FormatStudio",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="FormatStudio")