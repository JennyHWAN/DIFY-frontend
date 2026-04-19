@echo off
:: ============================================================
::  Build script for SOC Report Generator Windows EXE
::  Run this file on a Windows machine to produce the .exe
::  Output: dist\SOC_Report_Generator\SOC_Report_Generator.exe
:: ============================================================

echo ============================================================
echo  SOC Report Generator - EXE Builder
echo ============================================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.9+ and add it to PATH.
    pause
    exit /b 1
)

echo [1/3] Installing / upgrading dependencies...
pip install -r requirements.txt
pip install pyinstaller --upgrade
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)

echo.
echo [2/3] Cleaning previous build artifacts...
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

echo.
echo [3/3] Building executable with PyInstaller...
pyinstaller app.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed. See output above.
    pause
    exit /b 1
)

:: Copy .env.example next to the exe so users can rename it to .env
copy /Y .env.example dist\SOC_Report_Generator\.env.example >nul 2>&1

echo.
echo ============================================================
echo  BUILD SUCCESSFUL
echo  Executable: dist\SOC_Report_Generator\SOC_Report_Generator.exe
echo.
echo  BEFORE RUNNING:
echo    1. Copy dist\SOC_Report_Generator to the target machine
echo    2. Rename .env.example to .env in that folder
echo    3. Fill in DIFY_API_BASE_URL and DIFY_API_KEY in .env
echo    4. Double-click SOC_Report_Generator.exe
echo ============================================================
echo.
pause
