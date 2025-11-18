# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['csinfo_gui.py'],
    pathex=['.'],
    binaries=[],
    datas=[('csinfo', 'csinfo'), ('assets', 'assets')],
    hiddenimports=['csinfo', 'csinfo._impl'],
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
    name='csinfo',
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
    icon=['assets\\app.ico'],
)
