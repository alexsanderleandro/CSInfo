@echo off
setlocal enabledelayedexpansion

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Building optimized CSInfo...
pyinstaller ^
    --clean ^
    --onefile ^
    --windowed ^
    --icon=assets/app.ico ^
    --name CSInfo ^
    --exclude-module=matplotlib ^
    --exclude-module=numpy ^
    --exclude-module=pandas ^
    --exclude-module=scipy ^
    --exclude-module=PyQt5 ^
    --exclude-module=PySide2 ^
    --exclude-module=PySide6 ^
    --exclude-module=PyQt6 ^
    --exclude-module=tkinter ^
    --exclude-module=unittest ^
    --exclude-module=email ^
    --exclude-module=http ^
    --exclude-module=xml ^
    --exclude-module=html ^
    --exclude-module=asyncio ^
    --exclude-module=multiprocessing ^
    --upx-dir="C:\path\to\upx" ^
    --upx-exclude=vcruntime140.dll ^
    --hidden-import=win32timezone ^
    --hidden-import=win32api ^
    --hidden-import=win32con ^
    --hidden-import=win32security ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=wmi ^
    --hidden-import=psutil ^
    --hidden-import=requests ^
    --hidden-import=dotenv ^
    --hidden-import=pymupdf ^
    --hidden-import=PIL ^
    --add-data "assets/app.ico;assets" ^
    --add-data "assets/ico.png;assets" ^
    --add-binary "csinfo;csinfo" ^
    csinfo_gui.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Optimized build completed successfully!
    echo Executable: %CD%\dist\CSInfo.exe
    echo.
    echo Size: %~z0\dist\CSInfo.exe bytes
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%
)

pause
