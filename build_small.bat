@echo off
setlocal enabledelayedexpansion

:: Clean previous builds
echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

:: Create a temporary directory for the build
mkdir build_temp

:: Copy only necessary files
echo Copying required files...
xcopy /E /I /Y csinfo build_temp\csinfo
xcopy /Y csinfo_gui.py build_temp\
mkdir build_temp\assets
xcopy /Y assets\app.ico build_temp\assets\
xcopy /Y assets\ico.png build_temp\assets\

:: Create a temporary requirements.txt
echo Creating minimal requirements...
echo wmi>build_temp\requirements.txt
echo psutil>>build_temp\requirements.txt
echo requests>>build_temp\requirements.txt
echo python-dotenv>>build_temp\requirements.txt
echo pymupdf>>build_temp\requirements.txt
echo Pillow>>build_temp\requirements.txt
echo pywin32>>build_temp\requirements.txt

:: Install requirements in a virtual environment
echo Setting up build environment...
cd build_temp
python -m venv venv
call venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt

:: Build with PyInstaller
echo Building optimized executable...
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
    --exclude-module=unittest ^
    --exclude-module=email ^
    --exclude-module=http ^
    --exclude-module=asyncio ^
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
    csinfo_gui.py

:: Copy the built executable to the main dist folder
if exist dist\CSInfo.exe (
    copy /Y dist\CSInfo.exe ..\dist\
    cd ..
    echo.
    echo Build completed successfully!
    echo Executable: %CD%\dist\CSInfo.exe
    echo.
    echo Final size: %~z0\dist\CSInfo.exe bytes
) else (
    cd ..
    echo.
    echo Build failed! Check the build log above for errors.
)

:: Clean up
rmdir /s /q build_temp

pause
