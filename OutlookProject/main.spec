# main.spec
# -*- mode: python ; coding: utf-8 -*-

import os

block_cipher = None

project_root = os.path.abspath('.')

a = Analysis(
    ['main.py'],
    pathex=[project_root],
    binaries=[],
    datas=[
        (os.path.join(project_root, 'controllers'), 'controllers'),
        (os.path.join(project_root, 'models'), 'models'),
        (os.path.join(project_root, 'config.json5'), '.'),
        (os.path.join(project_root, 'msedgedriver.exe'), '.')
    ],
    hiddenimports=['win32com', 'win32com.client', 'pywintypes'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=True  # このオプションを追加してアーカイブファイルの生成を防ぐ
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='.'
)
