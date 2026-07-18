# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['word_header_footer_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Admin\\AppData\\Local\\Python\\pythoncore-3.14-64\\Lib\\site-packages\\tkinterdnd2', 'tkinterdnd2')],
    hiddenimports=['platform'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinterdnd2'],
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
    name='WordHeaderFooterCleaner',
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
)
