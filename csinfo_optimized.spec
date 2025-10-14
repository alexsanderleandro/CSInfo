# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

import os
import sys
import PyInstaller.__main__

# Add current directory to path
sys.path.append(os.getcwd())

# Large modules to exclude
excludes = [
    # Standard library
    'unittest', 'pydoc', 'doctest', 'test', 'pdb', 'email', 'http', 'xml',
    'html', 'wsgiref', 'curses', 'multiprocessing', 'asyncio', 'ensurepip',
    'distutils', 'pkg_resources', 'setuptools', 'pip', 'wheel',
    
    # GUI frameworks
    'tkinter', 'tcl', 'tk', 'PyQt5', 'PySide2', 'PySide6', 'PyQt6', 'wx',
    'kivy', 'pyglet', 'pygame', 'pyside', 'pyside2', 'pyside6',
    
    # Data science
    'numpy', 'pandas', 'scipy', 'matplotlib', 'seaborn', 'plotly', 'bokeh',
    'tensorflow', 'torch', 'theano', 'sklearn', 'statsmodels',
    
    # Web and networking
    'flask', 'django', 'tornado', 'requests_oauthlib', 'google', 'boto', 'boto3',
    'botocore', 's3transfer', 'urllib3', 'aiohttp', 'twisted', 'scrapy',
    
    # Databases
    'sqlalchemy', 'sqlite3', 'psycopg2', 'MySQLdb', 'pymongo', 'redis',
    
    # Development tools
    'pytest', 'nose', 'sphinx', 'pylint', 'pycodestyle', 'mypy', 'jedi',
    'ipython', 'jupyter', 'notebook', 'ipykernel', 'ipywidgets',
    
    # Unused image formats in PIL
    'PIL.ImageQt', 'PIL.ImageTk', 'PIL.ImageGrab', 'PIL.ImageGL',
    'PIL.ImageChops', 'PIL.ImageColor', 'PIL.ImageDraw', 'PIL.ImageEnhance',
    'PIL.ImageFile', 'PIL.ImageFilter', 'PIL.ImageFont', 'PIL.ImageMath',
    'PIL.ImageMorph', 'PIL.ImageOps', 'PIL.ImagePalette', 'PIL.ImagePath',
    'PIL.ImageQt', 'PIL.ImageSequence', 'PIL.ImageShow', 'PIL.ImageStat',
    'PIL.ImageTk', 'PIL.ImageWin', 'PIL.MpoImagePlugin', 'PIL.PaletteFile',
    'PIL.PdfParser', 'PIL.PsdImagePlugin', 'PIL.PyAccess',
    'PIL.SpiderImagePlugin', 'PIL.TiffImagePlugin', 'PIL.TiffTags',
    'PIL.WalImageFile', 'PIL.WebPImagePlugin', 'PIL.WmfImagePlugin',
    'PIL.XVThumbImagePlugin', 'PIL.XbmImagePlugin', 'PIL.XpmImagePlugin',
    'PIL._imagingcms', 'PIL._imagingft', 'PIL._imagingmorph', 'PIL._imagingtk',
    'PIL._webp',
]

a = Analysis(
    ['csinfo_gui.py'],
    pathex=[os.getcwd()],
    binaries=[],
    datas=[
        ('assets/app.ico', 'assets'),
        ('assets/ico.png', 'assets')
    ],
    hiddenimports=[
        'csinfo._impl',
        'csinfo._core',
        'csinfo.network_discovery',
        'win32timezone',
    ],
    hooksprefix=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=2,  # More aggressive optimization
)

# Optimize the PYZ archive
pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
    exclude_binaries=True
)

# Create the executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CSInfo',
    debug=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join('assets', 'app.ico'),
    bootloader_ignore_signals=False,
    ascii_mode=False,
    bootloader_append_signature=False,
)

# Don't create a console window
app = EXE(
    exe,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CSInfo',
    debug=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join('assets', 'app.ico'),
    bootloader_ignore_signals=False,
)
