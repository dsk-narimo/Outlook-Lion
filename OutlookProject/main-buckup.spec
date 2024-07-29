# -*- mode: python ; coding: utf-8 -*-

import os

# プロジェクトのルートディレクトリを取得
project_root = os.getcwd()

block_cipher = None

a = Analysis(
    [os.path.join(project_root, 'main.py')],
    pathex=[project_root],
    binaries=[],
    datas=[
        ('controllers', 'controllers'),  # controllersディレクトリをそのまま含める
        ('models', 'models'),  # modelsディレクトリをそのまま含める
        ('config.json5', '.'),  # config.json5ファイルを実行ファイルのルートに含める
        ('msedgedriver.exe', '.'),  # msedgedriver.exeファイルを実行ファイルのルートに含める
    ],
    hiddenimports=[],
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
    name='main',
    destdir='dist'
)
