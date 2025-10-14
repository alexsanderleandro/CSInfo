@echo off
setlocal enabledelayedexpansion

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Building CSInfo...
pyinstaller ^
    --clean ^
    --onefile ^
    --windowed ^
    --icon=assets/app.ico ^
    --name CSInfo ^
    --additional-hooks-dir=. ^
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
    --add-data "csinfo;csinfo" ^
    csinfo_gui.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build completed successfully!
    echo Executable: %CD%\dist\CSInfo.exe
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%
)

pause
