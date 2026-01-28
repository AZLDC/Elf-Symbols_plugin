# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包設定檔 - Elf-Symbols_plugin

import os

script_dir = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['Elf-Symbols_plugin.pyw'],
    pathex=[script_dir],
    binaries=[],
    datas=[
        ('繁.png', '.'),
        ('簡.png', '.'),
        ('轉.png', '.'),
    ],
    hiddenimports=[
        'keyboard',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='Elf-Symbols_plugin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 無命令列視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='轉.png',  # 使用轉.png 作為 exe 圖示
    uac_admin=False,
)
