#!/usr/bin/env python3
"""
Music Royalty Valuation Tool - Web Version
Run this file and open the URL in any browser (including on your phone).
"""

from flask import Flask, request, send_file, render_template_string
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import io
import re

app = Flask(__name__)

# HTML Template - Mobile-friendly
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Royalty Valuation Tool</title>
    <style>
        * {
            box-sizing: border-box;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
        }
        body {
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 500px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            padding: 30px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }
        h1 {
            margin: 0 0 10px 0;
            color: #333;
            font-size: 24px;
            text-align: center;
        }
        .subtitle {
            color: #666;
            text-align: center;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .upload-area {
            border: 2px dashed #ddd;
            border-radius: 12px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-bottom: 20px;
        }
        .upload-area:hover {
            border-color: #667eea;
            background: #f8f9ff;
        }
        .upload-area.dragover {
            border-color: #667eea;
            background: #f0f3ff;
        }
        .upload-icon {
            font-size: 48px;
            margin-bottom: 10px;
        }
        .upload-text {
            color: #666;
            margin-bottom: 10px;
        }
        .upload-hint {
            color: #999;
            font-size: 12px;
        }
        input[type="file"] {
            display: none;
        }
        .file-name {
            background: #f0f3ff;
            padding: 12px 16px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: none;
            align-items: center;
            gap: 10px;
        }
        .file-name.show {
            display: flex;
        }
        .file-name span {
            flex: 1;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .file-name button {
            background: none;
            border: none;
            color: #999;
            cursor: pointer;
            font-size: 18px;
        }
        .submit-btn {
            width: 100%;
            padding: 16px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        .submit-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }
        .submit-btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        .error {
            background: #fee;
            color: #c00;
            padding: 12px 16px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: none;
        }
        .error.show {
            display: block;
        }
        .success {
            background: #efe;
            color: #060;
            padding: 12px 16px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: none;
        }
        .success.show {
            display: block;
        }
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
        .loading.show {
            display: block;
        }
        .spinner {
            width: 40px;
            height: 40px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .instructions {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
        }
        .instructions h3 {
            font-size: 14px;
            color: #333;
            margin: 0 0 10px 0;
        }
        .instructions ol {
            margin: 0;
            padding-left: 20px;
            color: #666;
            font-size: 13px;
        }
        .instructions li {
            margin-bottom: 6px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Royalty Valuation Tool</h1>
        <p class="subtitle">Upload your earnings CSV to generate a DCF valuation</p>

        <div class="error" id="error"></div>
        <div class="success" id="success"></div>

        <form id="uploadForm" action="/process" method="post" enctype="multipart/form-data">
            <div class="upload-area" id="uploadArea">
                <div class="upload-icon">📊</div>
                <div class="upload-text">Tap to select your CSV file</div>
                <div class="upload-hint">or drag and drop here</div>
            </div>
            <input type="file" name="file" id="fileInput" accept=".csv,.xlsx">

            <div class="file-name" id="fileName">
                <span id="fileNameText"></span>
                <button type="button" id="clearFile">&times;</button>
            </div>

            <div class="loading" id="loading">
                <div class="spinner"></div>
                <div>Generating valuation...</div>
            </div>

            <button type="submit" class="submit-btn" id="submitBtn" disabled>
                Generate Valuation
            </button>
        </form>

        <div class="instructions">
            <h3>How it works:</h3>
            <ol>
                <li>Upload your royalty earnings CSV</li>
                <li>We'll analyze the yearly totals</li>
                <li>Download your complete DCF valuation spreadsheet</li>
            </ol>
        </div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileName = document.getElementById('fileName');
        const fileNameText = document.getElementById('fileNameText');
        const clearFile = document.getElementById('clearFile');
        const submitBtn = document.getElementById('submitBtn');
        const uploadForm = document.getElementById('uploadForm');
        const loading = document.getElementById('loading');
        const errorDiv = document.getElementById('error');
        const successDiv = document.getElementById('success');

        uploadArea.addEventListener('click', () => fileInput.click());

        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                updateFileName();
            }
        });

        fileInput.addEventListener('change', updateFileName);

        function updateFileName() {
            if (fileInput.files.length) {
                fileNameText.textContent = fileInput.files[0].name;
                fileName.classList.add('show');
                uploadArea.style.display = 'none';
                submitBtn.disabled = false;
                errorDiv.classList.remove('show');
            }
        }

        clearFile.addEventListener('click', () => {
            fileInput.value = '';
            fileName.classList.remove('show');
            uploadArea.style.display = 'block';
            submitBtn.disabled = true;
        });

        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            loading.classList.add('show');
            submitBtn.disabled = true;
            errorDiv.classList.remove('show');
            successDiv.classList.remove('show');

            const formData = new FormData(uploadForm);

            try {
                const response = await fetch('/process', {
                    method: 'POST',
                    body: formData
                });

                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = response.headers.get('X-Filename') || 'Valuation.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    a.remove();

                    successDiv.textContent = 'Valuation generated! Check your downloads.';
                    successDiv.classList.add('show');

                    // Reset form
                    fileInput.value = '';
                    fileName.classList.remove('show');
                    uploadArea.style.display = 'block';
                } else {
                    const text = await response.text();
                    throw new Error(text);
                }
            } catch (err) {
                errorDiv.textContent = err.message || 'Something went wrong. Please try again.';
                errorDiv.classList.add('show');
            } finally {
                loading.classList.remove('show');
                submitBtn.disabled = !fileInput.files.length;
            }
        });
    </script>
</body>
</html>
"""


def create_valuation_template(royalty_name, year_minus_3, year_minus_2, year_minus_1, ytd, base_year):
    """Creates the complete valuation template with data populated. Returns bytes."""

    wb = Workbook()
    ws = wb.active
    ws.title = "Valuation Model"

    # Define styles
    edit_font = Font(italic=True, color="0066CC")
    header_font = Font(bold=True, size=11)
    section_font = Font(bold=True, size=12)
    input_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    scenario_bear_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    scenario_base_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    scenario_bull_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    weighted_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # TITLE
    ws['A1'] = "MUSIC ROYALTY DCF VALUATION MODEL"
    ws['A1'].font = Font(bold=True, size=16)
    ws['A2'] = "Master Template with Weighted Scenario Analysis"
    ws['A2'].font = Font(italic=True, size=11, color="666666")

    # DATA INPUT SECTION
    ws['A4'] = "DATA INPUT"
    ws['A4'].font = section_font

    ws['A5'] = "Royalty Name/ID:"
    ws['B5'] = royalty_name
    ws['B5'].fill = input_fill
    ws['C5'] = "<- Edit"
    ws['C5'].font = edit_font

    ws['A7'] = "HISTORICAL ROYALTIES"
    ws['A7'].font = header_font

    ws['A8'] = "Year -3 Royalties"
    ws['B8'] = year_minus_3
    ws['B8'].fill = input_fill
    ws['B8'].number_format = '#,##0.00'
    ws['C8'] = "<- Edit"
    ws['C8'].font = edit_font

    ws['A9'] = "Year -2 Royalties"
    ws['B9'] = year_minus_2
    ws['B9'].fill = input_fill
    ws['B9'].number_format = '#,##0.00'
    ws['C9'] = "<- Edit"
    ws['C9'].font = edit_font

    ws['A10'] = "Year -1 Royalties"
    ws['B10'] = year_minus_1
    ws['B10'].fill = input_fill
    ws['B10'].number_format = '#,##0.00'
    ws['C10'] = "<- Edit"
    ws['C10'].font = edit_font

    ws['A11'] = "Current YTD Royalties"
    ws['B11'] = ytd
    ws['B11'].fill = input_fill
    ws['B11'].number_format = '#,##0.00'
    ws['C11'] = "<- Edit"
    ws['C11'].font = edit_font

    ws['A12'] = "3-Year Average"
    ws['B12'] = "=AVERAGE(B8:B10)"
    ws['B12'].number_format = '#,##0.00'

    ws['A13'] = "Base Year Royalties"
    ws['B13'] = base_year
    ws['B13'].fill = input_fill
    ws['B13'].number_format = '#,##0.00'
    ws['C13'] = "<- Edit (normalized starting CF)"
    ws['C13'].font = edit_font

    # KEY ASSUMPTIONS
    ws['A15'] = "KEY ASSUMPTIONS"
    ws['A15'].font = section_font

    ws['A16'] = "Growth Rate (Years 1-3)"
    ws['B16'] = 0.05
    ws['B16'].fill = input_fill
    ws['B16'].number_format = '0.0%'
    ws['C16'] = "<- Edit"
    ws['C16'].font = edit_font

    ws['A17'] = "Growth Rate (Years 4-5)"
    ws['B17'] = 0.03
    ws['B17'].fill = input_fill
    ws['B17'].number_format = '0.0%'
    ws['C17'] = "<- Edit"
    ws['C17'].font = edit_font

    ws['A18'] = "Discount Rate"
    ws['B18'] = 0.12
    ws['B18'].fill = input_fill
    ws['B18'].number_format = '0.0%'
    ws['C18'] = "<- Edit"
    ws['C18'].font = edit_font

    ws['A19'] = "Terminal Growth Rate"
    ws['B19'] = -0.05
    ws['B19'].fill = input_fill
    ws['B19'].number_format = '0.0%'
    ws['C19'] = "<- Edit (usually negative)"
    ws['C19'].font = edit_font

    # SCENARIO ANALYSIS
    ws['E4'] = "SCENARIO ANALYSIS"
    ws['E4'].font = section_font

    ws['F5'] = "Bear"
    ws['F5'].font = header_font
    ws['F5'].fill = scenario_bear_fill
    ws['F5'].alignment = Alignment(horizontal='center')
    ws['G5'] = "Base"
    ws['G5'].font = header_font
    ws['G5'].fill = scenario_base_fill
    ws['G5'].alignment = Alignment(horizontal='center')
    ws['H5'] = "Bull"
    ws['H5'].font = header_font
    ws['H5'].fill = scenario_bull_fill
    ws['H5'].alignment = Alignment(horizontal='center')

    ws['E6'] = "Base Year CF"
    ws['F6'] = "=B13*0.9"
    ws['G6'] = "=B13"
    ws['H6'] = "=B13*1.1"
    for col in ['F', 'G', 'H']:
        ws[f'{col}6'].number_format = '#,##0.00'

    ws['E7'] = "Growth (Yr 1-3)"
    ws['F7'] = "=B16-0.02"
    ws['G7'] = "=B16"
    ws['H7'] = "=B16+0.03"
    for col in ['F', 'G', 'H']:
        ws[f'{col}7'].number_format = '0.0%'

    ws['E8'] = "Growth (Yr 4-5)"
    ws['F8'] = "=B17-0.01"
    ws['G8'] = "=B17"
    ws['H8'] = "=B17+0.02"
    for col in ['F', 'G', 'H']:
        ws[f'{col}8'].number_format = '0.0%'

    ws['E9'] = "Discount Rate"
    ws['F9'] = "=B18+0.02"
    ws['G9'] = "=B18"
    ws['H9'] = "=B18"
    for col in ['F', 'G', 'H']:
        ws[f'{col}9'].number_format = '0.0%'

    ws['E10'] = "Terminal Growth"
    ws['F10'] = "=B19-0.02"
    ws['G10'] = "=B19"
    ws['H10'] = "=B19+0.02"
    for col in ['F', 'G', 'H']:
        ws[f'{col}10'].number_format = '0.0%'

    ws['E12'] = "Year 5 CF"
    for col, c in [('F', 'F'), ('G', 'G'), ('H', 'H')]:
        ws[f'{col}12'] = f"={c}6*(1+{c}7)^3*(1+{c}8)^2"
        ws[f'{col}12'].number_format = '#,##0.00'

    ws['E13'] = "Terminal Value"
    for col, c in [('F', 'F'), ('G', 'G'), ('H', 'H')]:
        ws[f'{col}13'] = f"={c}12*(1+{c}10)/({c}9-{c}10)"
        ws[f'{col}13'].number_format = '#,##0.00'

    ws['E14'] = "PV of Terminal"
    for col, c in [('F', 'F'), ('G', 'G'), ('H', 'H')]:
        ws[f'{col}14'] = f"={c}13/(1+{c}9)^5"
        ws[f'{col}14'].number_format = '#,##0.00'

    ws['E16'] = "Implied Value"
    ws['E16'].font = header_font
    for col, c in [('F', 'F'), ('G', 'G'), ('H', 'H')]:
        ws[f'{col}16'] = (
            f"={c}6/(1+{c}9)"
            f"+{c}6*(1+{c}7)/(1+{c}9)^2"
            f"+{c}6*(1+{c}7)^2/(1+{c}9)^3"
            f"+{c}6*(1+{c}7)^3*(1+{c}8)/(1+{c}9)^4"
            f"+{c}12/(1+{c}9)^5"
            f"+{c}14"
        )
        ws[f'{col}16'].number_format = '$#,##0.00'
        ws[f'{col}16'].font = Font(bold=True)

    ws['E17'] = "vs Base Case"
    ws['F17'] = "=F16/G16-1"
    ws['G17'] = "-"
    ws['H17'] = "=H16/G16-1"
    for col in ['F', 'H']:
        ws[f'{col}17'].number_format = '0.0%'

    # WEIGHTED AVERAGE VALUATION
    ws['E19'] = "WEIGHTED AVERAGE VALUATION"
    ws['E19'].font = section_font

    ws['E20'] = "Scenario Weights"
    ws['E20'].font = header_font
    ws['F20'] = "Bear Weight"
    ws['G20'] = "Base Weight"
    ws['H20'] = "Bull Weight"

    ws['F21'] = 0.25
    ws['F21'].fill = input_fill
    ws['F21'].number_format = '0%'
    ws['G21'] = 0.50
    ws['G21'].fill = input_fill
    ws['G21'].number_format = '0%'
    ws['H21'] = 0.25
    ws['H21'].fill = input_fill
    ws['H21'].number_format = '0%'
    ws['I21'] = "<- Edit weights (must = 100%)"
    ws['I21'].font = edit_font

    ws['E22'] = "Weight Check"
    ws['F22'] = "=F21+G21+H21"
    ws['F22'].number_format = '0%'
    ws['G22'] = '=IF(F22=1,"OK","ERROR: Must = 100%")'

    ws['E24'] = "WEIGHTED VALUATION"
    ws['E24'].font = Font(bold=True, size=12)
    ws['F24'] = "=F16*F21+G16*G21+H16*H21"
    ws['F24'].number_format = '$#,##0.00'
    ws['F24'].font = Font(bold=True, size=14)
    ws['F24'].fill = weighted_fill

    ws['E25'] = "Valuation Range"
    ws['F25'] = "=F16"
    ws['F25'].number_format = '$#,##0'
    ws['G25'] = "to"
    ws['H25'] = "=H16"
    ws['H25'].number_format = '$#,##0'

    ws['E26'] = "EV / Base Year CF"
    ws['F26'] = "=F24/B13"
    ws['F26'].number_format = '0.0x'

    ws['E27'] = "Payback Period (years)"
    ws['F27'] = "=F24/B13"
    ws['F27'].number_format = '0.0'

    # 5-YEAR DCF PROJECTION
    ws['A21'] = "5-YEAR DCF PROJECTION"
    ws['A21'].font = section_font

    headers = ["Year", "Base", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Terminal"]
    for i, h in enumerate(headers):
        col = get_column_letter(i + 1)
        ws[f'{col}22'] = h
        ws[f'{col}22'].font = header_font

    ws['A23'] = "Fiscal Year"
    ws['B23'] = datetime.now().year
    for i in range(1, 6):
        ws[f'{get_column_letter(i+2)}23'] = f"={get_column_letter(i+1)}23+1"
    ws['H23'] = "Perpetuity"

    ws['A24'] = "Royalty Income"
    ws['B24'] = "=B13"
    ws['C24'] = "=B24*(1+$B$16)"
    ws['D24'] = "=C24*(1+$B$16)"
    ws['E24'] = "=D24*(1+$B$16)"
    ws['F24'] = "=E24*(1+$B$17)"
    ws['G24'] = "=F24*(1+$B$17)"
    ws['H24'] = "=G24*(1+$B$19)"
    for col in 'BCDEFGH':
        ws[f'{col}24'].number_format = '#,##0.00'

    ws['A25'] = "Growth Rate"
    ws['B25'] = "-"
    ws['C25'] = "=$B$16"
    ws['D25'] = "=$B$16"
    ws['E25'] = "=$B$16"
    ws['F25'] = "=$B$17"
    ws['G25'] = "=$B$17"
    ws['H25'] = "=$B$19"
    for col in 'CDEFGH':
        ws[f'{col}25'].number_format = '0.0%'

    ws['A27'] = "Discount Factor"
    ws['B27'] = 1
    for i in range(1, 6):
        col = get_column_letter(i + 2)
        ws[f'{col}27'] = f"=1/(1+$B$18)^{i}"
        ws[f'{col}27'].number_format = '0.0000'
    ws['H27'] = "=G27"
    ws['H27'].number_format = '0.0000'

    ws['A28'] = "PV of Cash Flow"
    for col in ['C', 'D', 'E', 'F', 'G']:
        ws[f'{col}28'] = f"={col}24*{col}27"
        ws[f'{col}28'].number_format = '#,##0.00'

    # VALUATION SUMMARY
    ws['A30'] = "VALUATION SUMMARY"
    ws['A30'].font = section_font

    ws['A31'] = "Terminal Value (undiscounted)"
    ws['B31'] = "=H24/($B$18-$B$19)"
    ws['B31'].number_format = '#,##0.00'
    ws['C31'] = "Gordon Growth formula"
    ws['C31'].font = Font(italic=True, color="666666")

    ws['A32'] = "PV of Terminal Value"
    ws['B32'] = "=B31*G27"
    ws['B32'].number_format = '#,##0.00'

    ws['A34'] = "Sum of PV of Cash Flows"
    ws['B34'] = "=SUM(C28:G28)"
    ws['B34'].number_format = '#,##0.00'

    ws['A35'] = "PV of Terminal Value"
    ws['B35'] = "=B32"
    ws['B35'].number_format = '#,##0.00'

    ws['A36'] = "Enterprise Value"
    ws['B36'] = "=B34+B35"
    ws['B36'].number_format = '$#,##0.00'
    ws['B36'].font = Font(bold=True)

    ws['A38'] = "% from Cash Flows"
    ws['B38'] = "=B34/B36"
    ws['B38'].number_format = '0.0%'

    ws['A39'] = "% from Terminal Value"
    ws['B39'] = "=B35/B36"
    ws['B39'].number_format = '0.0%'

    # SENSITIVITY ANALYSIS 1
    ws['A41'] = "SENSITIVITY: Discount Rate vs Growth Rate (Years 1-3)"
    ws['A41'].font = section_font

    ws['A42'] = "Enterprise Value"
    ws['C42'] = "Growth Rate (Years 1-3)"
    ws['C42'].font = header_font

    growth_rates = [0.00, 0.02, 0.04, 0.06, 0.08, 0.10, 0.12]
    for i, gr in enumerate(growth_rates):
        col = get_column_letter(i + 3)
        ws[f'{col}43'] = gr
        ws[f'{col}43'].number_format = '0%'
        ws[f'{col}43'].font = header_font
        ws[f'{col}43'].alignment = Alignment(horizontal='center')

    ws['A44'] = "Discount"
    discount_rates = [0.08, 0.10, 0.12, 0.14, 0.16, 0.18]
    for i, dr in enumerate(discount_rates):
        row = 44 + i
        ws[f'B{row}'] = dr
        ws[f'B{row}'].number_format = '0%'
        ws[f'B{row}'].font = header_font

        for j in range(len(growth_rates)):
            col = get_column_letter(j + 3)
            formula = (
                f"=($B$13*(1+{col}$43)^3*(1+$B$17)^2*(1+$B$19)/($B{row}-$B$19))/(1+$B{row})^5"
                f"+$B$13/(1+$B{row})"
                f"+$B$13*(1+{col}$43)/(1+$B{row})^2"
                f"+$B$13*(1+{col}$43)^2/(1+$B{row})^3"
                f"+$B$13*(1+{col}$43)^3*(1+$B$17)/(1+$B{row})^4"
                f"+$B$13*(1+{col}$43)^3*(1+$B$17)^2/(1+$B{row})^5"
            )
            ws[f'{col}{row}'] = formula
            ws[f'{col}{row}'].number_format = '#,##0'

    ws['A45'] = "Rate"

    # SENSITIVITY ANALYSIS 2
    ws['A52'] = "SENSITIVITY: Discount Rate vs Terminal Growth Rate"
    ws['A52'].font = section_font

    ws['A53'] = "Enterprise Value"
    ws['C53'] = "Terminal Growth Rate"
    ws['C53'].font = header_font

    term_growth_rates = [-0.10, -0.07, -0.05, -0.03, 0.00, 0.02, 0.03]
    for i, tg in enumerate(term_growth_rates):
        col = get_column_letter(i + 3)
        ws[f'{col}54'] = tg
        ws[f'{col}54'].number_format = '0%'
        ws[f'{col}54'].font = header_font
        ws[f'{col}54'].alignment = Alignment(horizontal='center')

    ws['A55'] = "Discount"
    for i, dr in enumerate(discount_rates):
        row = 55 + i
        ws[f'B{row}'] = dr
        ws[f'B{row}'].number_format = '0%'
        ws[f'B{row}'].font = header_font

        for j in range(len(term_growth_rates)):
            col = get_column_letter(j + 3)
            formula = (
                f"=($B$13*(1+$B$16)^3*(1+$B$17)^2*(1+{col}$54)/($B{row}-{col}$54))/(1+$B{row})^5"
                f"+$B$13/(1+$B{row})"
                f"+$B$13*(1+$B$16)/(1+$B{row})^2"
                f"+$B$13*(1+$B$16)^2/(1+$B{row})^3"
                f"+$B$13*(1+$B$16)^3*(1+$B$17)/(1+$B{row})^4"
                f"+$B$13*(1+$B$16)^3*(1+$B$17)^2/(1+$B{row})^5"
            )
            ws[f'{col}{row}'] = formula
            ws[f'{col}{row}'].number_format = '#,##0'

    ws['A56'] = "Rate"

    # MODEL NOTES
    ws['A62'] = "MODEL NOTES"
    ws['A62'].font = section_font

    notes = [
        "* Green cells are INPUT cells - edit these with your royalty data",
        "* Royalties = pure cash flow (no costs modeled)",
        "* Terminal Value = Year 5 CF x (1+g) / (r-g) using Gordon Growth Model",
        "* Two-phase growth: Years 1-3 near-term, Years 4-5 mature growth",
        "* Weighted Valuation combines Bear/Base/Bull using your probability weights",
        "* Sensitivity tables show impact of key assumption changes"
    ]
    for i, note in enumerate(notes):
        ws[f'A{63+i}'] = note
        ws[f'A{63+i}'].font = Font(size=10, color="666666")

    # Column widths
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 26
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['I'].width = 30

    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def process_csv(file_storage):
    """Process uploaded CSV and return Excel bytes + filename."""

    # Read the file
    filename = file_storage.filename
    if filename.endswith('.xlsx'):
        df = pd.read_excel(file_storage)
    else:
        df = pd.read_csv(file_storage)

    # Find the amount column
    amount_col = None
    for col in ['payable_amount', 'amount', 'earnings', 'royalty']:
        if col in df.columns.str.lower().tolist():
            amount_col = [c for c in df.columns if c.lower() == col][0]
            break

    if amount_col is None:
        amount_cols = [c for c in df.columns if 'amount' in c.lower()]
        if amount_cols:
            amount_col = amount_cols[0]
        else:
            raise ValueError("Could not find an amount/earnings column in the CSV")

    # Find the year column
    year_col = None
    for col in ['distribution_year', 'year', 'date']:
        if col in df.columns.str.lower().tolist():
            year_col = [c for c in df.columns if c.lower() == col][0]
            break

    if year_col is None:
        raise ValueError("Could not find a year column in the CSV")

    # Sum by year
    yearly = df.groupby(year_col)[amount_col].sum().sort_index()

    # Get years
    current_year = datetime.now().year
    years_list = sorted(yearly.index)

    # Extract values
    ytd = yearly.get(current_year, 0)
    year_minus_1 = yearly.get(current_year - 1, 0)
    year_minus_2 = yearly.get(current_year - 2, 0)
    year_minus_3 = yearly.get(current_year - 3, 0)

    # If no current year data, shift
    if ytd == 0 and years_list:
        latest = max(years_list)
        ytd = yearly.get(latest, 0)
        year_minus_1 = yearly.get(latest - 1, 0)
        year_minus_2 = yearly.get(latest - 2, 0)
        year_minus_3 = yearly.get(latest - 3, 0)

    # Base year = most recent full year
    base_year = year_minus_1 if year_minus_1 > 0 else ytd

    # Generate output filename
    base_name = os.path.splitext(filename)[0]
    if 'listing' in base_name.lower():
        match = re.search(r'listing[-_]?(\d+)', base_name, re.IGNORECASE)
        if match:
            royalty_name = f"Listing {match.group(1)}"
        else:
            royalty_name = base_name
    else:
        royalty_name = base_name

    output_filename = f"{royalty_name} Valuation.xlsx"

    # Create the valuation
    excel_bytes = create_valuation_template(
        royalty_name=royalty_name,
        year_minus_3=year_minus_3,
        year_minus_2=year_minus_2,
        year_minus_1=year_minus_1,
        ytd=ytd,
        base_year=base_year
    )

    return excel_bytes, output_filename


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return 'No file uploaded', 400

    file = request.files['file']
    if file.filename == '':
        return 'No file selected', 400

    try:
        excel_bytes, output_filename = process_csv(file)

        response = send_file(
            excel_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )
        response.headers['X-Filename'] = output_filename
        return response

    except Exception as e:
        return str(e), 400


if __name__ == '__main__':
    import socket

    # Get local IP for mobile access
    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except:
        local_ip = '127.0.0.1'

    print("\n" + "="*50)
    print("  ROYALTY VALUATION TOOL - WEB VERSION")
    print("="*50)
    print(f"\n  Open in browser:")
    print(f"    - On this computer: http://localhost:5000")
    print(f"    - On your phone:    http://{local_ip}:5000")
    print(f"\n  (Make sure your phone is on the same WiFi)")
    print("\n  Press Ctrl+C to stop the server")
    print("="*50 + "\n")

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
