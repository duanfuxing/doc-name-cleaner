@echo off
chcp 65001 >nul
echo ============================================
echo   Build rename.py to exe (DocNameCleaner)
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+
    pause
    exit /b 1
)

:: Check Python version >= 3.10
python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" 2>nul
if errorlevel 1 (
    echo [ERROR] Python 3.10+ is required
    echo Current version:
    python --version
    pause
    exit /b 1
)

:: Install PyInstaller
echo [1/2] Installing PyInstaller ...
pip install pyinstaller -q
if errorlevel 1 (
    echo [ERROR] PyInstaller install failed
    pause
    exit /b 1
)

:: Build
echo [2/2] Building ...
pyinstaller --onefile --console --name DocNameCleaner --clean rename.py
if errorlevel 1 (
    echo [ERROR] Build failed
    pause
    exit /b 1
)

:: Clean up temp files
echo Cleaning up ...
rd /s /q build 2>nul
del DocNameCleaner.spec 2>nul

echo.
echo ============================================
echo   Build success!
echo   Output: dist\DocNameCleaner.exe
echo.
echo   Usage:
echo   1. Copy DocNameCleaner.exe and keywords.txt
echo      to the same folder
echo   2. Double-click DocNameCleaner.exe to run
echo ============================================
pause
