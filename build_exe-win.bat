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
python -m pip install -r requirements.txt --trusted-host pypi.org --trusted-host files.pythonhosted.org
python -m pip install pyinstaller --upgrade --trusted-host pypi.org --trusted-host files.pythonhosted.org
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)

echo.
echo [2/3] Cleaning previous build artifacts...
if exist C:\DIFY_build\dist   rmdir /s /q C:\DIFY_build\dist
if exist C:\DIFY_build\work   rmdir /s /q C:\DIFY_build\work

echo.
echo [3/3] Building executable with PyInstaller...
:: Use a short output path to avoid Windows 260-char MAX_PATH limit.
:: The project is inside a long OneDrive path; Streamlit's nested asset
:: paths would otherwise exceed the limit when appended to dist\.
python -m PyInstaller app-win.spec --distpath C:\DIFY_build\dist --workpath C:\DIFY_build\work
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed. See output above.
    pause
    exit /b 1
)

:: Copy .env.example next to the exe so users can rename it to .env
copy /Y .env.example C:\DIFY_build\dist\SOC_Report_Generator\.env.example >nul 2>&1

:: Copy template files next to the exe (app.py reads them from sys.executable's folder when frozen)
echo Copying template files...
copy /Y template_index.xlsx C:\DIFY_build\dist\SOC_Report_Generator\template_index.xlsx >nul 2>&1
if exist AR_template (
    xcopy /E /I /Y AR_template C:\DIFY_build\dist\SOC_Report_Generator\AR_template >nul 2>&1
)
if exist MA_template (
    xcopy /E /I /Y MA_template C:\DIFY_build\dist\SOC_Report_Generator\MA_template >nul 2>&1
)

echo.
echo ============================================================
echo  BUILD SUCCESSFUL
echo  Output folder: C:\DIFY_build\dist\SOC_Report_Generator\
echo.
echo  TO DISTRIBUTE TO USERS:
echo    1. Copy the entire folder to the target machine
echo    2. Double-click SOC_Report_Generator.exe
echo    (API keys are baked in - no setup required by the user)
echo    (Templates are included in the folder alongside the exe)
echo ============================================================
echo.
pause
