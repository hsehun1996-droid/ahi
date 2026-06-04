# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None

CTK_DIR = r"C:\Users\EX\AppData\Local\Programs\Python\Python313\Lib\site-packages\customtkinter"
CTKMB_DIR = r"C:\Users\EX\AppData\Local\Programs\Python\Python313\Lib\site-packages\CTkMessagebox"

datas = [
    (os.path.join(CTK_DIR, "assets"), "customtkinter/assets"),
    (os.path.join(CTKMB_DIR, "icons"), "CTkMessagebox/icons"),
    # PDF 내보내기용 Ghostscript (bin/lib만 — 폴더 최상위의 잡파일 제외)
    (os.path.join(SPECPATH, "ghostscript", "bin"), "ghostscript/bin"),
    (os.path.join(SPECPATH, "ghostscript", "lib"), "ghostscript/lib"),
    # CI 이미지 (창 아이콘 logo.ico, 브랜드 로고, 길통이 등)
    (os.path.join(SPECPATH, "concept art"), "concept art"),
    # UI/PDF 한글 폰트
    (os.path.join(SPECPATH, "fonts"), "fonts"),
]

hiddenimports = [
    "customtkinter",
    "CTkMessagebox",
    "PIL",
    "PIL.Image",
    "PIL.ImageTk",
    "PIL.ImageOps",
    "reportlab",
    "reportlab.pdfgen",
    "reportlab.pdfgen.canvas",
    "reportlab.lib.pagesizes",
    "reportlab.lib.utils",
    "openpyxl",
    "openpyxl.styles",
    "sqlite3",
    "tkinter",
    "tkinter.ttk",
    "tkinter.messagebox",
    "tkinter.filedialog",
    "tkinter.colorchooser",
    "tkinter.simpledialog",
    "tkinter.font",
    "darkdetect",
    "packaging",
]

a = Analysis(
    ["highway.py"],
    pathex=[r"C:\Users\EX\Desktop\project1"],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="고속도로_포장유지보수_관리",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(SPECPATH, "conceptart_icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="고속도로_포장유지보수_관리",
)
