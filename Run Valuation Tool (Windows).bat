@echo off
REM Music Royalty Valuation Tool - Windows Launcher (Desktop Version)
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
echo Checking dependencies...
pip install -r requirements.txt --quiet

REM Run the tool
python royalty_valuation.py

pause
