"""
export_powerbi_dataset.py
=========================
Prepares and exports an optimized analytics view for Power BI Desktop.
Extracts data from data/valuation_control.accdb, computes time-series
metrics (including 7-day rolling average breaks and daily break rates),
and saves:
- data/powerbi_breaks_summary.csv
- data/powerbi_calendar.csv
- data/powerbi_reconciliation_flat.csv

Provides mathematical cross-verification for DAX measures.
"""

from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd
import pyodbc


def get_connection(db_path: Path) -> pyodbc.Connection:
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str, autocommit=True)


def generate_powerbi_exports():
    project_root = Path(__file__).resolve().parent.parent
    db_path = project_root / "data" / "valuation_control.accdb"
    data_dir = project_root / "data"

    print(f"[*] Connecting to database: {db_path.name}")
    conn = get_connection(db_path)

    # 1. Full Reconciliation View (Market Prices + Trader Marks + Control Breaks)
    flat_query = """
        SELECT 
            m.PriceID,
            m.[Date] AS ValuationDate,
            i.InstrumentID,
            i.Ticker,
            i.AssetClass,
            i.Sector,
            m.MarketPrice,
            t.TraderMark,
            t.TraderID,
            ROUND(((t.TraderMark - m.MarketPrice) / m.MarketPrice) * 100, 2) AS CalculatedVariancePct,
            b.BreakID,
            b.Severity,
            b.Status,
            IIF(b.BreakID IS NULL, 0, 1) AS IsBreach
        FROM ((DailyMarketPrice m
        INNER JOIN Instruments i ON m.InstrumentID = i.InstrumentID)
        INNER JOIN DailyTraderMark t ON (m.InstrumentID = t.InstrumentID AND m.[Date] = t.[Date]))
        LEFT JOIN ControlBreaks b ON (m.InstrumentID = b.InstrumentID AND m.[Date] = b.[Date])
        ORDER BY m.[Date] ASC, i.Ticker ASC
    """
    df_flat = pd.read_sql(flat_query, conn)
    df_flat["ValuationDate"] = pd.to_datetime(df_flat["ValuationDate"]).dt.strftime("%Y-%m-%d")
    df_flat["Severity"] = df_flat["Severity"].fillna("In Tolerance")
    df_flat["Status"] = df_flat["Status"].fillna("N/A")

    flat_export_path = data_dir / "powerbi_reconciliation_flat.csv"
    df_flat.to_csv(flat_export_path, index=False)
    print(f"[+] Exported flat reconciliation dataset: {flat_export_path.name} ({len(df_flat)} rows)")

    # 2. Generate Calendar Dimension Table
    unique_dates = pd.to_datetime(sorted(df_flat["ValuationDate"].unique()))
    df_calendar = pd.DataFrame({"Date": unique_dates})
    df_calendar["Year"] = df_calendar["Date"].dt.year
    df_calendar["Month"] = df_calendar["Date"].dt.month
    df_calendar["MonthName"] = df_calendar["Date"].dt.strftime("%b")
    df_calendar["MonthYear"] = df_calendar["Date"].dt.strftime("%b %Y")
    df_calendar["DayOfWeek"] = df_calendar["Date"].dt.day_name()
    df_calendar["DayOfWeekNum"] = df_calendar["Date"].dt.dayofweek + 1
    df_calendar["Date"] = df_calendar["Date"].dt.strftime("%Y-%m-%d")

    cal_export_path = data_dir / "powerbi_calendar.csv"
    df_calendar.to_csv(cal_export_path, index=False)
    print(f"[+] Exported calendar dimension: {cal_export_path.name} ({len(df_calendar)} trading days)")

    # 3. Daily Time-Series Aggregations & 7-Day Rolling Average Verification
    df_daily = df_flat.groupby("ValuationDate").agg(
        TotalPositions=("PriceID", "count"),
        TotalBreaks=("IsBreach", "sum"),
        MajorBreaks=("Severity", lambda s: (s == "Major").sum()),
        MinorBreaks=("Severity", lambda s: (s == "Minor").sum()),
        AvgVariancePct=("CalculatedVariancePct", "mean")
    ).reset_index()

    df_daily["BreakRatePct"] = (df_daily["TotalBreaks"] / df_daily["TotalPositions"] * 100).round(2)
    # 7-day rolling average of breaks
    df_daily["Rolling7DayAvgBreaks"] = df_daily["TotalBreaks"].rolling(window=7, min_periods=1).mean().round(2)

    daily_export_path = data_dir / "powerbi_breaks_summary.csv"
    df_daily.to_csv(daily_export_path, index=False)
    print(f"[+] Exported daily aggregated time-series: {daily_export_path.name} ({len(df_daily)} days)")

    # 4. Verification Check
    print("\n" + "=" * 60)
    print("POWER BI DAX BENCHMARK CROSS-VERIFICATION")
    print("=" * 60)
    total_positions = len(df_flat)
    total_breaks = df_flat["IsBreach"].sum()
    overall_break_rate = (total_breaks / total_positions) * 100
    major_count = (df_flat["Severity"] == "Major").sum()
    minor_count = (df_flat["Severity"] == "Minor").sum()
    max_breaks_day = df_daily.loc[df_daily["TotalBreaks"].idxmax()]

    print(f"Total Positions Evaluated : {total_positions}")
    print(f"Total Breaks Flagged      : {total_breaks}")
    print(f"Overall Break Rate        : {overall_break_rate:.2f}%")
    print(f"Major Severity Breaks     : {major_count}")
    print(f"Minor Severity Breaks     : {minor_count}")
    print(f"Peak Breaks Date          : {max_breaks_day['ValuationDate']} ({max_breaks_day['TotalBreaks']} breaks)")
    print(f"Peak 7-Day Rolling Avg    : {df_daily['Rolling7DayAvgBreaks'].max():.2f} breaks/day")
    print("=" * 60 + "\n")

    conn.close()


if __name__ == "__main__":
    generate_powerbi_exports()
