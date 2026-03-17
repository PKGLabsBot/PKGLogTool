# -*- mode: python ; coding: utf-8 -*-
# LogTool PyInstaller spec
# 절대경로 없이 SPECPATH 기준으로 동작 - 어떤 PC에서도 빌드 가능
# 사용법: pyinstaller LogTool.spec

import os
from pathlib import Path

spec_dir = Path(SPECPATH)

# version.txt 에서 버전 읽기
version_file = spec_dir / "version.txt"
try:
    APP_VERSION = version_file.read_text(encoding="utf-8").strip()
except Exception:
    APP_VERSION = "1.0"

EXE_NAME = f"LogTool_v{APP_VERSION}"

# ============================================================
# pandas가 동적으로 로드하는 패키지는 PyInstaller가 자동 감지 못함
# xlsxwriter, openpyxl 등 명시적으로 포함 필요
# ============================================================
hidden_imports = [
    "xlsxwriter",
    "xlsxwriter.workbook",
    "xlsxwriter.worksheet",
    "xlsxwriter.chart",
    "xlsxwriter.utility",
    "pandas.io.excel._xlsxwriter",
    "pandas.io.excel._openpyxl",
    "pandas.io.formats.excel",
]

excludes = [
    "matplotlib",
    "scipy",
    "PIL",
    "IPython",
    "notebook",
    "pytest",
    "setuptools",
]

a = Analysis(
    [str(spec_dir / "logtool.py")],
    pathex=[str(spec_dir)],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=str(spec_dir / "file_version_info.txt"),
    icon=[str(spec_dir / "icon.ico")],
)
