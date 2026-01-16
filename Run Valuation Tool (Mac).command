#!/bin/bash
# Music Royalty Valuation Tool - Mac Launcher
# Double-click this file to run the tool

cd "$(dirname "$0")"

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    osascript -e 'display alert "Python 3 Required" message "Please install Python 3 from python.org" as critical'
    exit 1
fi

# Check/install required packages
python3 -c "import pandas, openpyxl" 2>/dev/null || {
    echo "Installing required packages..."
    pip3 install pandas openpyxl --quiet
}

# Run the tool
python3 royalty_valuation.py
