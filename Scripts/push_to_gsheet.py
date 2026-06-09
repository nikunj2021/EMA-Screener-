"""
Push EMA Breadth results to Google Sheets.
Reads the last row from history Excel files and appends to Google Sheets.
"""

import os
import json
import gspread
import openpyxl
from google.oauth2.service_account import Credentials
from datetime import datetime, date

# ── CONFIG ───────────────────────────────────────────────────────────────────
SHEET_ID   = os.environ["GSHEET_ID"]
CREDS_JSON = os.environ["GSHEET_CREDS_JSON"]

HISTORY_FILES = {
    "Nifty50_EMA":  "Report/EMAReport/nifty_50_ema_breadth_history.xlsx",
    "Nifty250_EMA": "Report/EMAReport/Micro_250_ema_breadth_history.xlsx",
    "Nifty500_EMA": "Report/EMAReport/nifty_500_ema_breadth_history.xlsx",
}

# Sheet name inside each Excel history file (EMA Breadth History sheet)
EXCEL_SHEET_NAME = "EMA Breadth History"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── AUTH ─────────────────────────────────────────────────────────────────────
creds_dict = json.loads(CREDS_JSON)
creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
gc    = gspread.authorize(creds)
sh    = gc.open_by_key(SHEET_ID)

# ── HELPERS ───────────────────────────────────────────────────────────────────
def excel_date_to_str(value):
    """Convert Excel serial date number OR date/datetime object to DD-Mon-YYYY string."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d-%b-%Y")
    if isinstance(value, (int, float)):
        # Excel serial date: days since 1899-12-30
        try:
            from datetime import timedelta
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=int(value))).strftime("%d-%b-%Y")
        except Exception:
            return str(value)
    return str(value)

def get_or_create_worksheet(sh, title, headers):
    """Get existing worksheet or create with headers."""
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=20)
        ws.append_row(headers, value_input_option="USER_ENTERED")
        print(f"  Created new tab: {title}")
    return ws

def read_last_row(excel_path, sheet_name):
    """
    Read last data row from Excel history file.
    Returns a list: [date_str, scan_time, pct20, pct50, pct200, total]
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    # Print available sheets for debugging
    print(f"  Available sheets in {excel_path}: {wb.sheetnames}")

    if sheet_name not in wb.sheetnames:
        print(f"  WARNING: Sheet '{sheet_name}' not found! Using first sheet.")
        ws = wb.worksheets[0]
    else:
        ws = wb[sheet_name]

    rows = list(ws.iter_rows(values_only=True))
    print(f"  Total rows in sheet (incl. header): {len(rows)}")

    if len(rows) < 2:
        print(f"  No data rows found in {excel_path}")
        return None

    # Print header and last row for debugging
    print(f"  Header row : {rows[0]}")
    print(f"  Last row   : {rows[-1]}")

    last = rows[-1]

    # Extract by position: col0=Date, col1=ScanTime, col2=pct20, col3=pct50, col4=pct200, col5=total
    date_val  = excel_date_to_str(last[0])
    scan_time = str(last[1]) if last[1] is not None else ""
    pct20     = round(float(last[2]), 2) if last[2] is not None else ""
    pct50     = round(float(last[3]), 2) if last[3] is not None else ""
    pct200    = round(float(last[4]), 2) if last[4] is not None else ""
    total     = int(last[5]) if last[5] is not None else ""

    return [date_val, scan_time, pct20, pct50, pct200, total]

# ── PUSH EACH FILE ────────────────────────────────────────────────────────────
HEADERS = ["Date", "Scan Time", "Above EMA20 %", "Above EMA50 %", "Above EMA200 %", "Total Stocks"]

for tab_name, excel_path in HISTORY_FILES.items():
    print(f"\nProcessing: {tab_name}")

    if not os.path.exists(excel_path):
        print(f"  File not found: {excel_path} — skipping")
        continue

    row_data = read_last_row(excel_path, EXCEL_SHEET_NAME)
    if not row_data:
        continue

    ws = get_or_create_worksheet(sh, tab_name, HEADERS)

    # Duplicate check — skip if date already in column A
    existing_dates = ws.col_values(1)
    if row_data[0] in existing_dates:
        print(f"  Already exists: {row_data[0]} — skipping")
        continue

    ws.append_row(row_data, value_input_option="USER_ENTERED")
    print(f"  ✓ Appended: {row_data}")

print("\n✅ Google Sheets push complete.")
