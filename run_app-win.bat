@echo off
:: ============================================================
::  Quick-launch script (no build needed — requires Python)
::  Double-click this file to start the app in your browser.
:: ============================================================

echo ============================================================
echo  SOC Report Generator
echo ============================================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.9+ and add it to PATH.
    pause
    exit /b 1
)

:: Install dependencies if not already present
echo Checking dependencies...
pip show streamlit >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies (first run only)...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
)

:: Load .env if it exists
if exist .env (
    echo Loading .env configuration...
) else (
    echo [WARNING] No .env file found. You can enter the API key in the app sidebar.
)

echo.
echo Starting app at http://localhost:8501 ...
echo Press Ctrl+C in this window to stop the server.
echo.

:: Open browser after a short delay (start is non-blocking)
start "" /b cmd /c "timeout /t 3 >nul && start http://localhost:8501"

:: Run Streamlit
python -m streamlit run app.py --server.headless=true --browser.gatherUsageStats=false

pause
