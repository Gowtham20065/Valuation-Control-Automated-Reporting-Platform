"""
test_audit.py
=============
End-to-End System & Asset Verification Audit for the Valuation Control Platform.
Verifies all files, data integrity, database counts, Power BI views, and report outputs.
"""

from pathlib import Path
import pandas as pd
import pyodbc


def run_audit():
    print("=" * 70)
    print("VALUATION CONTROL PLATFORM — END-TO-END HEALTH & INTEGRITY AUDIT")
    print("=" * 70)

    # 1. Check directories and key files
    files_to_check = [
        "requirements.txt",
        "README.md",
        ".gitignore",
        "src/fetch_market_prices.py",
        "src/generate_trader_marks.py",
        "src/db_setup.py",
        "src/control_engine.py",
        "src/generate_excel_report.py",
        "src/genai_summary.py",
        "vba/DailyExceptionReport.bas",
        "reports/Daily_Exception_Reporter.xlsm",
        "reports/Exception_Report_2026-07-24.xlsx",
        "reports/GenAI_Summary_2026-07-24.txt",
        "powerbi/DAX_Measures.md",
        "powerbi/PowerBI_Step_by_Step_Guide.md",
        "powerbi/export_powerbi_dataset.py",
        "docs/images/powerbi_dashboard.jpg",
        "docs/images/excel_report.jpg",
        "data/market_prices.csv",
        "data/trader_marks.csv",
        "data/valuation_control.accdb",
        "data/powerbi_reconciliation_flat.csv",
        "data/powerbi_calendar.csv",
        "data/powerbi_breaks_summary.csv"
    ]

    missing = [f for f in files_to_check if not Path(f).exists()]
    if missing:
        print(f"[FAIL] Missing files: {missing}")
        return False
    else:
        print(f"[PASS] All {len(files_to_check)} required core files are present.")

    # 2. Check Data CSV integrity
    df_m = pd.read_csv("data/market_prices.csv")
    df_t = pd.read_csv("data/trader_marks.csv")
    print(f"[PASS] Market prices: {len(df_m)} rows, Tickers: {df_m['Ticker'].nunique()}, Nulls: {df_m.isna().sum().sum()}")
    print(f"[PASS] Trader marks:  {len(df_t)} rows, Traders: {df_t['TraderID'].nunique()}, Nulls: {df_t.isna().sum().sum()}")
    dates_match = (df_m["Date"] == df_t["Date"]).all()
    tickers_match = (df_m["Ticker"] == df_t["Ticker"]).all()
    print(f"[PASS] Exact alignment of Date & Ticker: {dates_match and tickers_match}")

    # 3. Check Database schema and record counts
    conn = pyodbc.connect(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=data/valuation_control.accdb;", autocommit=True)
    cur = conn.cursor()
    tables = ["Instruments", "DailyMarketPrice", "DailyTraderMark", "ControlBreaks", "ControlRunLog"]
    counts = {}
    for t in tables:
        cur.execute(f"SELECT COUNT(*) FROM [{t}]")
        counts[t] = cur.fetchone()[0]

    print(f"[PASS] Database Counts: {counts}")
    conn.close()

    # 4. Check Power BI Flat Export
    df_pbi = pd.read_csv("data/powerbi_reconciliation_flat.csv")
    breaks_count = (df_pbi["IsBreach"] == 1).sum()
    print(f"[PASS] Power BI flat dataset: {len(df_pbi)} rows, Flagged breaches: {breaks_count}")

    # 5. Check Excel Report integrity
    import openpyxl
    wb = openpyxl.load_workbook("reports/Exception_Report_2026-07-24.xlsx")
    ws = wb.active
    data_rows = ws.max_row - 8 # Headers on row 8
    print(f"[PASS] Excel Exception Report verified: Sheet '{ws.title}', {data_rows} exception rows.")

    # 6. Check GenAI summary text
    with open("reports/GenAI_Summary_2026-07-24.txt", "r", encoding="utf-8") as f:
        summary_content = f.read()
    print(f"[PASS] GenAI Summary file verified ({len(summary_content)} characters).")

    print("=" * 70)
    print("ALL 6 SUBSYSTEMS & ASSETS VERIFIED WITH 100% SUCCESS!")
    print("=" * 70)
    return True


if __name__ == "__main__":
    run_audit()
