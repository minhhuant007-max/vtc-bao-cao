# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['G:\\APP LOC DU LIEU\\Tổng hợp\\VTC_Bao_Cao\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('G:\\APP LOC DU LIEU\\Tổng hợp\\VTC_Bao_Cao\\config', 'config')],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='VTC_Bao_Cao',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='VTC_Bao_Cao',
)
