# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
root = Path(SPECPATH)


a = Analysis(
    [str(root / 'installer_runner.py')],
    pathex=[],
    binaries=[(str(Path(DISTPATH) / 'VoucherPass.exe'), '.')],
    datas=[],
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
    a.binaries,
    a.datas,
    [],
    name='VoucherPass_Setup',
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
    icon=[str(root / 'assets' / 'VoucherPass.ico')],
)
