@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === DB View Export - build ===
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [1/2] installing pyinstaller ...
    python -m pip install --upgrade pyinstaller openpyxl
)
echo [2/2] building ...
python build.py
echo.
pause
