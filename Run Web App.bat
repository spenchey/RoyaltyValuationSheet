@echo off
REM Music Royalty Valuation Tool - Web Version
REM Double-click to start the web server

cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python from python.org
    pause
    exit /b 1
)

REM Check/install required packages
python -c "import flask, pandas, openpyxl" 2>nul || (
    echo Installing required packages...
    python -m pip install flask pandas openpyxl --quiet
)

REM Run the web app
python web_app.py

pause
