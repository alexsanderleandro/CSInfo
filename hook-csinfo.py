from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = collect_all('csinfo')

# Add any additional hidden imports if needed
hiddenimports.extend([
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
])
