# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['csinfo_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/app.ico', 'assets'),
        ('assets/ico.png', 'assets')
    ],
    hiddenimports=[
        'csinfo',
        'csinfo._impl',
        'csinfo._core',
        'csinfo.network_discovery',
        'win32timezone',
        'win32api',
        'win32con',
        'win32security',
        'pythoncom',
        'pywintypes',
        'wmi',
        'psutil',
        'requests',
        'dotenv',
        'pymupdf',
        'PIL',
        'PIL._imaging',
        'PIL.Image',
        'PIL.ImageFile',
        'PIL.PngImagePlugin',
        'PIL.ImageDraw',
        'PIL.ImageFont',
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CSInfo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/app.ico',
)
