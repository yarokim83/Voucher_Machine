# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_all

datas = [('c:\\Users\\baewoong.kim\\.gemini\\VoucherPass\\dist\\VoucherPass.ico', 'dist'), ('c:\\Users\\baewoong.kim\\.gemini\\VoucherPass\\dist\\app_icon.png', 'dist'), ('c:\\Users\\baewoong.kim\\.gemini\\VoucherPass\\dist\\voucherpass_icon.png', 'dist'), ('C:\\Users\\baewoong.kim\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python312\\site-packages\\tkinterdnd2', 'tkinterdnd2')]
binaries = [('c:\\Users\\baewoong.kim\\.gemini\\VoucherPass\\bin\\SumatraPDF.exe', 'bin')]
hiddenimports = []
datas += collect_data_files('tkinterdnd2')
tmp_ret = collect_all('pystray')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pynput')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('selenium')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    icon=['c:\\Users\\baewoong.kim\\.gemini\\VoucherPass\\dist\\VoucherPass.ico'],
)
