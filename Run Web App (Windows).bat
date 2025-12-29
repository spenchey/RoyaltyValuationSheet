@echo off
REM Music Royalty Valuation Tool - Web App Launcher (Windows)
REM Double-click this file to start the web server

cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python from python.org
    pause
    exit /b 1
)

REM Check/install required packages
echo Checking dependencies...
pip install -r requirements.txt --quiet

echo.
echo ==================================================
echo   Starting Royalty Valuation Tool - AI-Powered
echo ==================================================
echo.
echo   Opening http://localhost:5001 in your browser...
echo   Press Ctrl+C to stop the server
echo.

REM Open browser after a short delay
start "" http://localhost:5001

REM Run the web app on port 5001 (5000 often used by other services)
set PORT=5001
python web_app.py

pause
