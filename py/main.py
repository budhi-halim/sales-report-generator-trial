import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import io
from datetime import datetime
from pyscript import ffi, window, document

# Section: Configuration Class
# Holds all configurable settings for columns, formats, and processing rules
class ColumnConfig:
    def __init__(self):
        self.main_column = 'B'  # Main column to check for data presence
        self.extract_columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']  # Columns to extract
        self.validation_column = 'K'  # Column to validate for blanks in rows 4-6
        self.key_map = {  # Mapping from column letters to data keys
            'B': 'area',
            'C': 'date',
            'D': 'customer_name',
            'E': 'invoice_no',
            'F': 'product_type',
            'G': 'product_name',
            'H': 'quantity',
            'I': 'unit_price'
        }
        self.output_headers = [  # Headers for the output sales sheet
            'Area', 'Tanggal', 'Periode', 'Customer', 'No. Faktur', 'Jenis', 'Produk', 'Jumlah',
            'Harga Satuan', 'Total (IDR)', 'Kurs', 'Total (USD)'
        ]
        self.preserve_upper = {'ABC', 'PT', 'CV', 'US'}  # Words to keep uppercase
        self.area_replacements = {  # Area abbreviations to full names
            'Bdg': 'Bandung',
            'Bgr': 'Bogor',
            'Bks': 'Bekasi',
            'Jkt': 'Jakarta',
            'Lpg': 'Lampung',
            'Mdn': 'Medan',
            'Tgr': 'Tangerang'
        }
        self.idr_format_2 = r'_-"Rp"* #,##0.00_ ;_-"Rp"* -#,##0.00_ ;_-"Rp"* "-"??_ ;_-@_-'
        self.idr_format_0 = r'_-"Rp"* #,##0_ ;_-"Rp"* -#,##0_ ;_-"Rp"* "-"??_ ;_-@_-'
        self.usd_format_2 = r'_-"$"* #,##0.00_ ;_-"$"* -#,##0.00_ ;_-"Rp"* "-"??_ ;_-@_-' 

# Section: Utility Functions
# Helper functions for data processing
def col_to_index(col):
    # Converts column letter to zero-based index
    return ord(col.upper()) - ord('A')

def proper_case(text, preserve_upper):
    # Applies proper casing to text, preserving specific uppercase words
    if not text:
        return ''
    words = str(text).strip().split()
    processed = []
    for word in words:
        upper_word = word.upper()
        if upper_word in preserve_upper:
            processed.append(upper_word)
        else:
            processed.append(word.capitalize())
    return ' '.join(processed)

def process_area(area, replacements, preserve_upper):
    # Processes area string with replacements and casing
    if not area:
        return ''
    area = proper_case(area.strip().split()[0], preserve_upper)
    return replacements.get(area, area)

# Section: Structure Validation
# Validates input file structure for blanks in validation column
def validate_structure(buffer, filename, config):
    # Checks if specified validation cells are blank
    engine = 'xlrd' if filename.endswith('.xls') else 'openpyxl'
    df = pd.read_excel(buffer, header=None, engine=engine)
    validation_idx = col_to_index(config.validation_column)
    if validation_idx >= df.shape[1]:
        return True  # Column doesn't exist, treat as valid
    cells = [df.iloc[3, validation_idx], df.iloc[4, validation_idx], df.iloc[5, validation_idx]]
    return all(pd.isna(cell) or str(cell).strip() == '' for cell in cells)

# Section: Date Validation
# Validates extracted dates are proper datetime
def validate_dates(data):
    # Checks if all dates in extracted data are valid datetime
    for row in data:
        date_val = row.get('date')
        if date_val and not isinstance(date_val, datetime):
            # Try force convert if str or num
            try:
                if isinstance(date_val, (str, int, float)):
                    pd.to_datetime(date_val)
                else:
                    return False
            except:
                return False
    return True

# Section: Data Extraction
# Extracts relevant data from input file
def extract_data(buffer, filename, config):
    # Reads and extracts data rows based on config, handling date conversion
    engine = 'xlrd' if filename.endswith('.xls') else 'openpyxl'
    df = pd.read_excel(buffer, header=None, engine=engine, skiprows=3)
    main_idx = col_to_index(config.main_column)
    extract_idxs = [col_to_index(col) for col in config.extract_columns]
    data = []
    for _, row in df.iterrows():
        if pd.isna(row[main_idx]) or str(row[main_idx]).strip() == '':
            continue
        row_data = {}
        for i, col in enumerate(config.extract_columns):
            val = row[extract_idxs[i]]
            if pd.isna(val):
                val = ''
            elif col == 'H':
                try:
                    val = float(val)
                except:
                    val = ''
            elif col == 'I':
                try:
                    val = float(val)
                except:
                    val = ''
            elif col == 'C':
                if isinstance(val, (str, int, float)):
                    try:
                        val = pd.to_datetime(val)
                    except:
                        val = val  # Keep original if fail, validate later
            elif col == 'E':
                val = str(val).rstrip('.0') if isinstance(val, float) else str(val)  # Force str, remove .0
            row_data[config.key_map[col]] = val
        data.append(row_data)
    return data

# Section: Data Processing
# Normalizes and processes extracted data
def process_data(combined_data, config):
    # Applies normalization and collects unique periods
    periods = set()
    for row in combined_data:
        row['area'] = process_area(row['area'], config.area_replacements, config.preserve_upper)
        row['customer_name'] = proper_case(row['customer_name'], config.preserve_upper)
        row['product_name'] = proper_case(row['product_name'], config.preserve_upper)
        date_value = row.get('date')
        if date_value and isinstance(date_value, datetime):
            periods.add(date_value.strftime('%Y-%m'))
        row['invoice_no'] = str(row['invoice_no']) if row.get('invoice_no') else ''
        row['product_type'] = str(row['product_type']) if row.get('product_type') else ''
        row['quantity'] = row['quantity'] if row.get('quantity') else ''
        row['unit_price'] = row['unit_price'] if row.get('unit_price') else ''
    return sorted(list(periods)), combined_data

# Section: Blank Check
# Checks for blank required fields
def check_blanks(combined_data):
    # Identifies blank cells in required fields
    warnings = []
    key_to_col = {
        'area': 'A',
        'date': 'B',
        'customer_name': 'D',
        'invoice_no': 'E',
        'product_type': 'F',
        'product_name': 'G',
        'quantity': 'H',
        'unit_price': 'I'
    }
    required_keys = list(key_to_col.keys())
    for i, row in enumerate(combined_data, 1):
        for key in required_keys:
            if not row.get(key):
                warnings.append(f"{key_to_col[key]}{i+1}")
    return warnings

# Section: Table Generation
# General function to generate Excel table from 2D data
def generate_table(sheet, headers, data_2d, ref, table_name, col_formats=None, row_formulas=None):
    # Writes headers and data to sheet, applies formats/formulas, creates table
    for c, header in enumerate(headers, 1):
        sheet.cell(1, c).value = header
    for r_idx, row_data in enumerate(data_2d, 0):
        r = r_idx + 2  # Data starts at row 2
        for c, val in enumerate(row_data, 1):
            sheet.cell(r, c).value = val
            if col_formats and col_formats.get(c):
                sheet.cell(r, c).number_format = col_formats[c]
        if row_formulas and row_formulas.get(r_idx):
            for col_idx, formula in row_formulas[r_idx].items():
                sheet.cell(r, col_idx).value = formula
                if col_formats and col_formats.get(col_idx):
                    sheet.cell(r, col_idx).number_format = col_formats[col_idx]
    tab = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    sheet.add_table(tab)

# Section: Main Processing Function
# Orchestrates file processing and output
def process_files(file_datas):
    message_div = document.getElementById('message')
    message_div.innerHTML = ''
    file_datas = file_datas.to_py()
    if not file_datas:
        message = 'No files uploaded.'
        message_div.innerHTML = message
        return {'message': message}
    config = ColumnConfig()
    invalid_files = []  # Structure issues
    invalid_date_files = []  # Date format issues
    combined_data = []
    for f in file_datas:
        name = f['name']
        buffer = io.BytesIO(bytes(f['data']))
        if not validate_structure(buffer, name, config):
            invalid_files.append(name)
            continue
        buffer.seek(0)
        data = extract_data(buffer, name, config)
        if not validate_dates(data):
            invalid_date_files.append(name)
            continue
        combined_data.extend(data)
    msg = ''
    if invalid_files:
        msg += '<p>These files do not follow the required format:</p><ul>' + ''.join(f'<li>{f}</li>' for f in invalid_files) + '</ul>'
    if invalid_date_files:
        if msg:
            msg += ''
        msg += '<p>These files have invalid date formatting:</p><ul>' + ''.join(f'<li>{f}</li>' for f in invalid_date_files) + '</ul>'
    if msg:
        message_div.innerHTML = msg
        return {'message': msg, 'type': 'error'}
    sorted_periods, combined_data = process_data(combined_data, config)
    blank_cells = check_blanks(combined_data)
    wb = Workbook()
    sheet1 = wb.active
    sheet1.title = 'Sales Data'
    # Build 2D data for sales
    sales_data_2d = []
    sales_row_formulas = {}
    sales_col_formats = {
        1: '@',
        2: 'dd/mm/yyyy',
        3: 'dd/mm/yyyy',
        4: '@',
        5: '@',
        6: '@',
        7: '@',
        8: '#,##0',
        9: config.idr_format_2,
        10: config.idr_format_2,
        11: config.idr_format_0,
        12: config.usd_format_2
    }
    for i, row in enumerate(combined_data):
        date_value = row.get('date') if isinstance(row.get('date'), datetime) else ''
        sales_data_2d.append([row['area'], date_value, '', row['customer_name'], row['invoice_no'], row['product_type'], row['product_name'], row['quantity'], row['unit_price'], '', '', ''])
        sales_row_formulas[i] = {
            3: f'=DATE(YEAR(B{i+2}), MONTH(B{i+2}), 1)',
            10: f'=H{i+2}*I{i+2}',
            11: f'=VLOOKUP(TEXT(C{i+2}, "YYYY-MM"), \'Exchange Rate\'!A:B, 2, FALSE)',
            12: f'=J{i+2}/K{i+2}'
        }
    generate_table(sheet1, config.output_headers, sales_data_2d, f"A1:L{len(combined_data) + 1}", "Sales_Data", sales_col_formats, sales_row_formulas)
    sheet2 = wb.create_sheet('Exchange Rate')
    # Build 2D data for exchange (Python-computed periods)
    exchange_headers = ['Period', 'Rate']
    exchange_data_2d = [[p, None] for p in sorted_periods]
    exchange_col_formats = {1: '@'}
    generate_table(sheet2, exchange_headers, exchange_data_2d, f"A1:B{len(sorted_periods) + 1}", "Exchange_Rate", exchange_col_formats)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    buffer = list(output.getvalue())
    if blank_cells:
        msg = '<p>Processing complete. Warning: Please check these empty cells in the output:</p><ul>' + ''.join(f'<li>{c}</li>' for c in blank_cells) + '</ul>'
        message_type = 'warning'
    else:
        msg = 'Processing complete. All data processed successfully.'
        message_type = 'success'
    message_div.innerHTML = msg
    return {'buffer': buffer, 'message': msg, 'type': message_type}

window.process_files = ffi.create_proxy(process_files)