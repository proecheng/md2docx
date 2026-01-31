@echo off
chcp 65001 >nul
echo ===============================================
echo MD2DOCX - Build EXE
echo ===============================================
echo.

REM Check if pyinstaller is installed
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building...
echo.

pyinstaller --onefile --windowed --name "MD2DOCX" md2docx.py

echo.
echo ===============================================
if exist "dist\MD2DOCX.exe" (
    echo Build successful!
    echo.
    echo EXE location: dist\MD2DOCX.exe
    echo.
    echo Usage:
    echo   1. Double-click to open file dialog
    echo   2. Drag .md file onto the exe
    echo   3. Command line: MD2DOCX.exe file.md
) else (
    echo Build failed. Please check errors above.
)
echo ===============================================
pause
