@echo off
chcp 65001 > nul

echo.
echo  ============================================================
echo   QXDM AUTO LOG - EXE Build
echo  ============================================================
echo.

python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [*] Installing PyInstaller...
    pip install pyinstaller --quiet
    if errorlevel 1 (
        echo [ERROR] PyInstaller install failed
        pause
        exit /b 1
    )
)

python -c "import win32com" 2>nul
if errorlevel 1 (
    echo [*] Installing pywin32...
    pip install pywin32 --quiet
    if errorlevel 1 (
        echo [ERROR] pywin32 install failed
        pause
        exit /b 1
    )
)

:: Read __version__ from qxdm_auto_log.py
python -c "f=open('qxdm_auto_log.py',encoding='utf-8').read(); import re; m=re.search('__version__.*?([0-9]+\.[0-9]+\.[0-9]+)',f); print(m.group(1) if m else '0.0.0')" > _ver.txt
set /p VERSION=<_ver.txt
del _ver.txt

echo [*] Version: %VERSION%
set EXE_NAME=qxdm_auto_log_v%VERSION%

if exist "dist\%EXE_NAME%.exe" (
    del /f /q "dist\%EXE_NAME%.exe"
)
if exist build (
    rmdir /s /q build
)

echo [*] Building EXE: %EXE_NAME%.exe
echo.

pyinstaller --onefile --windowed --name "%EXE_NAME%" ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    --hidden-import win32api ^
    --hidden-import pywintypes ^
    --collect-submodules win32com ^
    qxdm_auto_log.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo   Build complete!  dist\%EXE_NAME%.exe
echo  ============================================================
echo.

explorer dist

pause
