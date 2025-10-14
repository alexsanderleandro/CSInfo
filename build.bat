@echo off
setlocal enabledelayedexpansion

:: Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

:: Build with PyInstaller
pyinstaller ^
    --clean ^
    --onefile ^
    --windowed ^
    --icon=assets/app.ico ^
    --name CSInfo ^
    --hidden-import=csinfo ^
    --hidden-import=csinfo._impl ^
    --hidden-import=csinfo._core ^
    --hidden-import=csinfo.network_discovery ^
    --hidden-import=win32timezone ^
    --hidden-import=win32api ^
    --hidden-import=win32con ^
    --hidden-import=win32security ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=wmi ^
    --add-data "assets/app.ico;assets" ^
    --add-data "assets/ico.png;assets" ^
    csinfo_gui.py

echo.
echo Build complete!
echo The executable is located at: %CD%\dist\CSInfo.exe
pause