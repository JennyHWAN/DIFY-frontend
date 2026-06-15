@echo off
setlocal
:: ============================================================
::  Build script for SOC Report Generator Windows EXE
::  Run this file on a Windows machine to produce the .exe
::  Output: C:\DIFY_build\dist\SOC_Report_Generator\SOC_Report_Generator.exe
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

:: Local, non-OneDrive staging folder for the build source.
set "SRC=C:\DIFY_src"

echo [1/4] Staging source to a local folder (%SRC%)...
:: PyInstaller fails with "OSError: [Errno 22] Invalid argument" when it reads
:: source files that live on a OneDrive "Files On-Demand" path: online-only
:: placeholders are not real local files and the read fails. Copying the project
:: to a plain local folder forces OneDrive to download (hydrate) every file and
:: hands PyInstaller real files to read. It also keeps source paths short, which
:: avoids the Windows 260-char MAX_PATH limit during analysis.
if exist "%SRC%" rmdir /s /q "%SRC%"
robocopy "%~dp0." "%SRC%" /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP ^
    /XD .git venv .venv __pycache__ .idea build dist
:: robocopy uses a bitmask exit code: 0-7 = success, 8+ = real failure.
if errorlevel 8 (
    echo [ERROR] Failed to stage source to %SRC%.
    echo         If files are online-only, open OneDrive and choose
    echo         "Always keep on this device" for the project folder, then retry.
    pause
    exit /b 1
)

:: Everything from here runs against the local copy.
pushd "%SRC%"

echo.
echo [2/4] Installing / upgrading dependencies...
python -m pip install -r requirements.txt --trusted-host pypi.org --trusted-host files.pythonhosted.org
python -m pip install pyinstaller --upgrade --trusted-host pypi.org --trusted-host files.pythonhosted.org
if errorlevel 1 (
    echo [ERROR] pip install failed.
    popd
    pause
    exit /b 1
)

echo.
echo [3/4] Cleaning previous build artifacts...
if exist C:\DIFY_build\dist   rmdir /s /q C:\DIFY_build\dist
if exist C:\DIFY_build\work   rmdir /s /q C:\DIFY_build\work

echo.
echo [4/4] Building executable with PyInstaller...
:: Short output path also avoids the 260-char MAX_PATH limit when Streamlit's
:: deeply nested asset paths are appended to dist\.
python -m PyInstaller app-win.spec --distpath C:\DIFY_build\dist --workpath C:\DIFY_build\work
set "BUILD_ERR=%errorlevel%"
popd
if not "%BUILD_ERR%"=="0" (
    echo [ERROR] PyInstaller build failed. See output above.
    pause
    exit /b 1
)

:: Copy .env.example next to the exe so users can rename it to .env.
:: Source files come from the hydrated local copy in %SRC%, not OneDrive.
copy /Y "%SRC%\.env.example" C:\DIFY_build\dist\SOC_Report_Generator\.env.example >nul 2>&1

:: Copy template files next to the exe (app.py reads them from sys.executable's folder when frozen)
echo Copying template files...
copy /Y "%SRC%\template_index.xlsx" C:\DIFY_build\dist\SOC_Report_Generator\template_index.xlsx >nul 2>&1
if exist "%SRC%\AR_template" (
    xcopy /E /I /Y "%SRC%\AR_template" C:\DIFY_build\dist\SOC_Report_Generator\AR_template >nul 2>&1
)
if exist "%SRC%\MA_template" (
    xcopy /E /I /Y "%SRC%\MA_template" C:\DIFY_build\dist\SOC_Report_Generator\MA_template >nul 2>&1
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
endlocal
