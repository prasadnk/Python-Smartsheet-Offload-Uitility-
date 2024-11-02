# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['DownloadSmartsheet_V2.py'],
    pathex=[r'C:\Users\Prasad\OneDrive\Documents\SS_Py_OffLoad\DownloadSmartsheet_V2.py'],
    binaries=[],
    datas=[],
    hiddenimports=['smartsheet', 'pandas', 'openpyxl', 'tkinter'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SmartsheetDownloader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SmartsheetDownloader',
)
