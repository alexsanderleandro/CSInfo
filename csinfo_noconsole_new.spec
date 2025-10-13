# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['csinfo_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('version.py', '.'), ('assets\\ico.png', 'assets')],
    hiddenimports=['platform', 'reportlab', 'reportlab.lib', 'reportlab.pdfgen', 'reportlab.lib.pagesizes', 'reportlab.platypus'],
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
    name='csinfo_noconsole_new',
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
