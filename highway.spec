# -*- mode: python ; coding: utf-8 -*-
"""고속도로 포장유지보수 이력 관리 프로그램 — PyInstaller 빌드 스펙.

목표: 파이썬/패키지가 설치되지 않은 새 데스크탑에서도 모든 기능이 동작하도록
필요한 리소스를 한 폴더에 모두 번들링한다.

번들 포함 항목
──────────────
- customtkinter / CTkMessagebox 의 assets·icons (테마·아이콘) — 설치 경로 자동 탐색
- reportlab 데이터 (PDF 내보내기)
- ghostscript bin/lib (PDF 내 EPS 변환)         ← 프로젝트 동봉
- 한글 폰트 fonts/ (UI·PDF 한글 표시)            ← 프로젝트 동봉
- concept art/ (창 아이콘·로고·캐릭터 이미지)     ← 프로젝트 동봉
- templates/operation_plan_template.hwp (운영계획변경 한글 양식) ← 프로젝트 동봉

한글(HWP) 자동 작성은 실행 PC에 '한글'과 pywin32 가 설치돼 있어야 동작한다.
(exe 에는 win32com 등이 포함되지만, 한글 프로그램 자체는 대상 PC에 설치 필요)
"""
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# ── 프로젝트 루트 (spec 파일이 위치한 폴더) ────────────────────────────────
ROOT = SPECPATH


def _exists(*parts):
    """프로젝트 루트 기준 경로가 실제 존재할 때만 (src, dst) 튜플을 만든다."""
    src = os.path.join(ROOT, *parts)
    return src if os.path.exists(src) else None


# ── 1) 서드파티 패키지 리소스 자동 수집 (설치 경로 하드코딩 제거) ──────────
#   customtkinter 는 assets(테마 json·폰트)를 런타임에 파일로 읽으므로 반드시 동봉.
#   CTkMessagebox 는 icons(png)를 파일로 읽는다. reportlab 은 내장 폰트·데이터 사용.
datas = []
datas += collect_data_files("customtkinter")
datas += collect_data_files("CTkMessagebox")
datas += collect_data_files("reportlab")

# ── 2) 프로젝트 동봉 리소스 (있을 때만 추가) ───────────────────────────────
_bundle = [
    # PDF 내보내기용 Ghostscript (bin/lib 만 — 폴더 최상위 잡파일 제외)
    (_exists("ghostscript", "bin"), "ghostscript/bin"),
    (_exists("ghostscript", "lib"), "ghostscript/lib"),
    # CI 이미지 (창 아이콘 logo.ico, 브랜드 로고, 길통이 등)
    (_exists("concept art"), "concept art"),
    # UI/PDF 한글 폰트 (NotoSansKR-*, malgun)
    (_exists("fonts"), "fonts"),
    # 운영계획변경 한글(HWP) 양식 템플릿
    (_exists("templates"), "templates"),
]
datas += [(src, dst) for (src, dst) in _bundle if src]

# ── 3) hiddenimports (정적 분석으로 누락되기 쉬운 모듈 명시) ────────────────
hiddenimports = [
    "customtkinter",
    "CTkMessagebox",
    "PIL",
    "PIL.Image",
    "PIL.ImageTk",
    "PIL.ImageOps",
    "PIL.EpsImagePlugin",
    "reportlab",
    "reportlab.pdfgen",
    "reportlab.pdfgen.canvas",
    "reportlab.lib.pagesizes",
    "reportlab.lib.utils",
    "reportlab.pdfbase",
    "reportlab.pdfbase.pdfmetrics",
    "reportlab.pdfbase.ttfonts",
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
    # 프로젝트 로컬 모듈
    "constants",
    "utils",
    "canvas_utils",
    "dropdown_widget",
    "hwp_export",
    # 운영계획변경 한글(HWP) 자동화 (방법 A: COM) — 대상 PC에 한글+pywin32 필요
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
]
hiddenimports += collect_submodules("mixins")

a = Analysis(
    ["highway.py"],
    pathex=[ROOT],
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
    icon=os.path.join(ROOT, "conceptart_icon.ico"),
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
