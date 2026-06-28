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

echo [1/4] Installing / upgrading dependencies...
python -m pip install -r requirements.txt --trusted-host pypi.org --trusted-host files.pythonhosted.org
python -m pip install pyinstaller --upgrade --trusted-host pypi.org --trusted-host files.pythonhosted.org
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)

echo.
echo [2/4] Cleaning previous build artifacts...
if exist C:\DIFY_build\dist   rmdir /s /q C:\DIFY_build\dist
if exist C:\DIFY_build\work   rmdir /s /q C:\DIFY_build\work

echo.
echo [3/4] Building executable with PyInstaller...
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

:: ── Package into a Velopack installer + update bundle ──────────────────────
:: Produces C:\DIFY_build\Releases\SOC_Report_Generator-<ver>-Setup.exe (per-user,
:: no-admin install) plus the .nupkg used by the in-app auto-updater. Requires the
:: .NET SDK; `vpk` is the Velopack CLI (installed once as a global dotnet tool).
echo.
echo [4/4] Packaging installer with Velopack (vpk)...
dotnet tool install -g vpk >nul 2>&1
dotnet tool update  -g vpk >nul 2>&1
set /p APPVER=<VERSION
vpk pack --packId SOC_Report_Generator --packVersion %APPVER% ^
    --packDir C:\DIFY_build\dist\SOC_Report_Generator ^
    --mainExe SOC_Report_Generator.exe ^
    --packTitle "SOC Report Generator" ^
    --outputDir C:\DIFY_build\Releases
if errorlevel 1 (
    echo [ERROR] vpk pack failed. Is the .NET SDK installed? See output above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  BUILD SUCCESSFUL
echo  Installer: C:\DIFY_build\Releases\SOC_Report_Generator-%APPVER%-Setup.exe
echo.
echo  TO DISTRIBUTE TO USERS (first time):
echo    1. Send them the Setup.exe above
echo    2. They double-click it - installs to %%LocalAppData%% (no admin)
echo    (API keys + templates are baked in - no setup required)
echo.
echo  UPDATES AFTER THAT ARE AUTOMATIC:
echo    Publish a new GitHub Release (bump VERSION) and the app
echo    self-updates on next launch.
echo ============================================================
echo.
pause
