# -*- mode: python ; coding: utf-8 -*-


block_cipher = None

version = 'v1.0.5'

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.ico', '.'),
        ('natural_air_system_manual.pdf', '.'),
        ('icons/*.png', 'icons'),
    ],
    excludes=['tests'],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name=f'natural-air-system_{version}',
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
    icon='app.ico',
)
