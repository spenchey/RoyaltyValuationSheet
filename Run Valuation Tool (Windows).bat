@echo off
REM Music Royalty Valuation Tool - Windows Launcher
REM Double-click this file to run the tool

cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python from python.org
    pause
    exit /b 1
)

REM Check/install required packages
python -c "import pandas, openpyxl" 2>nul || (
    echo Installing required packages...
    pip install pandas openpyxl --quiet
)

REM Run the tool
python royalty_valuation.py

pause
