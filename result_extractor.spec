# PyInstaller spec for Result Extractor (Windows .exe)
# Build on Windows: pyinstaller result_extractor.spec

# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# CustomTkinter themes/fonts (required for GUI to render correctly)
ctk_datas = collect_data_files('customtkinter')

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=ctk_datas,
    hiddenimports=[
        'result_extractor',
        'result_extractor.converter',
        'result_extractor.excel_writer',
        'result_extractor.pdf_reader',
        'result_extractor.config',
        'pdfplumber',
        'openpyxl',
        'pandas',
        'customtkinter',
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Result Extractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window (GUI app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
