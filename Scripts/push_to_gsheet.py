"""
Push EMA Breadth results to Google Sheets.
Reads the latest row from the history Excel and appends it to the sheet.
Tabs: Nifty50_EMA | Nifty250_EMA | Nifty500_EMA | Sectoral_EMA
"""

import os
import json
import gspread
import openpyxl
from google.oauth2.service_account import Credentials
from datetime import datetime

# ── CONFIG ──────────────────────────────────────────────────────────────────
SHEET_ID = os.environ["GSHEET_ID"]
CREDS_JSON = os.environ["GSHEET_CREDS_JSON"]

HISTORY_FILES = {
    "Nifty50_EMA":  "Report/EMAReport/nifty_50_ema_breadth_history.xlsx",
    "Nifty250_EMA": "Report/EMAReport/nifty_250_ema_breadth_history.xlsx",
    "Nifty500_EMA": "Report/EMAReport/nifty_500_ema_breadth_history.xlsx",
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── AUTH ─────────────────────────────────────────────────────────────────────
creds_dict = json.loads(CREDS_JSON)
creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_ID)

# ── HELPER ───────────────────────────────────────────────────────────────────
def get_or_create_worksheet(sh, title, headers):
    """Get existing sheet tab or create with headers."""
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=20)
        ws.append_row(headers, value_input_option="USER_ENTERED")
    return ws

def read_last_row_from_excel(path, sheet_name="EMA Breadth History"):
    """Read the last data row from the history Excel sheet."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return None
    headers = rows[0]
    last = rows[-1]
    return dict(zip(headers, last))

# ── PUSH EMA BREADTH HISTORIES ───────────────────────────────────────────────
EMA_HEADERS = ["Date", "Scan Time", "Above EMA20 %", "Above EMA50 %", "Above EMA200 %", "Total Stocks"]

for tab_name, excel_path in HISTORY_FILES.items():
    if not os.path.exists(excel_path):
        print(f"  Skipping {tab_name} — file not found: {excel_path}")
        continue

    row_dict = read_last_row_from_excel(excel_path)
    if not row_dict:
        print(f"  Skipping {tab_name} — no data rows")
        continue

    ws = get_or_create_worksheet(sh, tab_name, EMA_HEADERS)

    # Check if this date already exists (avoid duplicates)
    existing = ws.col_values(1)  # all dates in column A
    date_val = str(row_dict.get("Date", ""))
    if date_val in existing:
        print(f"  {tab_name}: {date_val} already exists — skipping")
        continue

    new_row = [
        date_val,
        str(row_dict.get("Scan Time", "")),
        row_dict.get("Above EMA20 %") or row_dict.get("pct20", ""),
        row_dict.get("Above EMA50 %") or row_dict.get("pct50", ""),
        row_dict.get("Above EMA200 %") or row_dict.get("pct200", ""),
        row_dict.get("Total", ""),
    ]
    ws.append_row(new_row, value_input_option="USER_ENTERED")
    print(f"  ✓ {tab_name}: appended {date_val}")

print("\nGoogle Sheets push complete.")