"""
Auto-Catchup Runner for NSE EMA Breadth & Sectoral Analysis
Checks existing Excel reports for the last recorded trading date, finds all
completed NSE trading days (Mon-Fri) missing up to the latest available session,
and runs EMA50D, EMA250D, EMA500D, and sectoral_ema_analysis for each missed date.
"""

import os
import sys
import subprocess
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import openpyxl
import yfinance as yf

IST = ZoneInfo("Asia/Kolkata")

HISTORY_FILES = [
    "Report/EMAReport/nifty_50_ema_breadth_history.xlsx",
    "Report/EMAReport/Micro_250_ema_breadth_history.xlsx",
    "Report/EMAReport/nifty_500_ema_breadth_history.xlsx",
]
SECTORAL_FILE = "Report/EMAReport/Sectoral_EMA_Analysis.xlsx"

EMA_SCRIPTS = [
    "Scripts/EMA50D.py",
    "Scripts/EMA250D.py",
    "Scripts/EMA500D.py",
]
SECTORAL_SCRIPT = "Scripts/sectoral_ema_analysis.py"


def today_ist() -> date:
    return datetime.now(IST).date()


def parse_report_date(val) -> date | None:
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if not val:
        return None
    raw = str(val).strip()
    for fmt in ("%d-%b-%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def get_last_recorded_date(xlsx_path: str, sheet_name: str = "EMA Breadth History") -> date | None:
    """Read the latest date recorded in column A of the history workbook."""
    if not os.path.exists(xlsx_path):
        return None
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        dates = []
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            d = parse_report_date(row[0])
            if d:
                dates.append(d)
        wb.close()
        return max(dates) if dates else None
    except Exception as e:
        print(f"  ⚠ Could not read {xlsx_path}: {e}")
        return None


def get_sectoral_recorded_date(xlsx_path: str) -> date | None:
    """Read the Price Date recorded in Sheet 2 (col C) or Sheet 1 title of Sectoral_EMA_Analysis.xlsx."""
    if not os.path.exists(xlsx_path):
        return None
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        if "Ticker Detail (Above EMA)" in wb.sheetnames:
            ws = wb["Ticker Detail (Above EMA)"]
            for row in ws.iter_rows(min_row=3, max_row=10, min_col=3, max_col=3, values_only=True):
                d = parse_report_date(row[0])
                if d:
                    wb.close()
                    return d
        wb.close()
    except Exception:
        pass
    return None


def fetch_nse_trading_days(start_date: date, end_date: date) -> list[date]:
    """
    Query benchmark NSE ticker (RELIANCE.NS) to get actual completed NSE trading days
    in [start_date, end_date]. Automatically excludes weekends & NSE holidays.
    """
    fetch_start = (start_date - timedelta(days=5)).strftime("%Y-%m-%d")
    fetch_end   = (end_date + timedelta(days=1)).strftime("%Y-%m-%d")
    raw = yf.download("RELIANCE.NS", start=fetch_start, end=fetch_end,
                      interval="1d", progress=False, auto_adjust=True)
    if raw is None or raw.empty:
        return []
    days = sorted({
        (idx.date() if hasattr(idx, "date") else idx)
        for idx in raw.index
    })
    return [d for d in days if start_date <= d <= end_date and d.weekday() < 5]


def run_script_for_date(script_path: str, target_dt: date):
    date_arg = target_dt.strftime("%Y-%m-%d")
    print(f"\n>>> Running {script_path} --date {date_arg}")
    res = subprocess.run([sys.executable, script_path, "--date", date_arg])
    if res.returncode != 0:
        print(f"\n❌ {script_path} failed for {date_arg} (exit code {res.returncode}).")
        sys.exit(res.returncode)


def main():
    now = datetime.now(IST)
    today = now.date()
    print("=" * 65)
    print(f"  NSE Market Report Catch-Up Check  |  {now.strftime('%d-%b-%Y %H:%M:%S IST')}")
    print("=" * 65)

    # 1. Check last recorded date across all 3 EMA history files
    last_dates = [get_last_recorded_date(p) for p in HISTORY_FILES]
    valid_dates = [d for d in last_dates if d is not None]

    # Look back up to 14 calendar days if history files don't exist yet
    oldest_recorded = min(valid_dates) if len(valid_dates) == len(HISTORY_FILES) else (today - timedelta(days=14))
    search_from = oldest_recorded + timedelta(days=1)

    print(f"  Recorded dates in EMA histories : {[str(d) for d in last_dates]}")
    print(f"  Checking completed NSE sessions : {search_from} → {today}")

    # 2. Fetch actual completed NSE trading days from Yahoo Finance
    trading_days = fetch_nse_trading_days(search_from, today)

    if not trading_days:
        # Check if Sectoral report alone is behind the latest EMA history date
        latest_ema_date = max(valid_dates) if valid_dates else None
        sec_date = get_sectoral_recorded_date(SECTORAL_FILE)
        if latest_ema_date and sec_date != latest_ema_date:
            print(f"  Sectoral report ({sec_date}) is behind ({latest_ema_date}) — updating...")
            run_script_for_date(SECTORAL_SCRIPT, latest_ema_date)
        else:
            print("\n✅ All reports are already up to date (or today is a weekend/NSE holiday). Nothing to run.")
        return

    print(f"  Missing trading day(s) to process: {[d.strftime('%d-%b-%Y') for d in trading_days]}")

    # 3. Backfill each missing weekday chronologically
    for idx, td in enumerate(trading_days):
        is_latest = (idx == len(trading_days) - 1)
        for script, recorded_dt in zip(EMA_SCRIPTS, last_dates):
            if recorded_dt is None or recorded_dt < td:
                run_script_for_date(script, td)

        # Sectoral report is a snapshot file (no history log), so generate it for the latest missing day
        if is_latest:
            run_script_for_date(SECTORAL_SCRIPT, td)

    print("\n✅ Catch-up and daily report generation complete!")


if __name__ == "__main__":
    main()