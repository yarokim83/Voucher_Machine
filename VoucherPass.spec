# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_all

from pathlib import Path
root = Path(SPECPATH)
datas = [(str(root / 'assets'), 'assets'), (str(root / 'version.txt'), '.')]
binaries = [(str(root / 'bin' / 'SumatraPDF.exe'), 'bin')]
hiddenimports = ['win32timezone']
for package in ('tkinterdnd2', 'pystray', 'pynput', 'selenium'):
    package_data, package_binaries, package_imports = collect_all(package)
    datas += package_data
    binaries += package_binaries
    hiddenimports += package_imports


a = Analysis(
    [str(root / 'main.py')],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'tensorflow', 'matplotlib', 'scipy', 'IPython', 'jupyter', 'sympy', 'pandas'],
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
    name='VoucherPass',
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
