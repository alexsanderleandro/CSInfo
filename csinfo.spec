# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files
import os

# collect all submodules and data files from the local `csinfo` package so
# the dynamically-loaded backend (`csinfo._impl`) is bundled by PyInstaller.
_hidden_csinfo = collect_submodules('csinfo')
_datas_csinfo = collect_data_files('csinfo')

# include top-level assets and version.py so they are available in sys._MEIPASS
# prefer bundling the PNG (used by the GUI and PDF header) and an ICO for the
# Windows executable icon. Also include version.py at the bundle root so
# _load_version() can find it when frozen.
extra_datas = []
try:
    extra_datas.append((os.path.join(project_root, 'assets', 'ico.png'), 'assets'))
except Exception:
    pass
try:
    extra_datas.append((os.path.join(project_root, 'assets', 'app.ico'), 'assets'))
except Exception:
    pass
try:
    # place version.py at the root of the extracted bundle so _load_version()
    # finds it as <MEIPASS>/version.py
    extra_datas.append((os.path.join(project_root, 'version.py'), '.'))
except Exception:
    pass
# merge with collected datas from the csinfo package
try:
    _datas_csinfo = list(_datas_csinfo) + extra_datas
except Exception:
    pass

# ensure project root is in pathex so Analysis can find the local package
try:
    project_root = os.path.abspath(os.path.dirname(__file__))
except NameError:
    # when PyInstaller executes the spec the __file__ variable may not be
    # defined; fallback to the current working directory
    project_root = os.path.abspath(os.getcwd())

a = Analysis(
    ['csinfo_gui.py'],
    pathex=[project_root],
    binaries=[],
    datas=_datas_csinfo,
    hiddenimports=_hidden_csinfo,
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
    # Use the ICO file for the Windows exe icon when available. We fall back to
    # not specifying an icon if it isn't present — PyInstaller on Windows
    # expects an .ico file for the executable icon. The PNG is still bundled
    # and used at runtime by Tkinter via iconphoto.
    icon=os.path.join(project_root, 'assets', 'app.ico'),
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
