"""
control_engine.py
=================
Automated Independent Price Verification (IPV) and Valuation Control Engine.

Core Functionality:
1. Reconciles daily trader position marks against independent market prices.
2. Calculates relative percentage variance:
      Variance% = ((TraderMark - MarketPrice) / MarketPrice) * 100
3. Applies institutional policy tolerance rules:
      - Threshold: +/- 3.0%
      - Minor Severity: 3.0% < |Variance%| <= 5.0%
      - Major Severity: |Variance%| > 5.0%
4. Populates the ControlBreaks exception ledger with audit status 'Open'.
5. Logs execution performance and audit metrics into ControlRunLog.
"""

import argparse
from datetime import datetime
from pathlib import Path
import time
import pandas as pd
import pyodbc

TOLERANCE_THRESHOLD = 3.0  # 3.0% variance threshold
MAJOR_SEVERITY_THRESHOLD = 5.0  # > 5.0% variance triggers major escalation


def get_db_connection(db_path: Path) -> pyodbc.Connection:
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str, autocommit=False)


def run_control_engine(db_path: Path, target_date: str | None = None) -> dict:
    """
    Executes the IPV control engine across historical records or for a specific date.
    """
    start_time = time.perf_counter()
    conn = get_db_connection(db_path)
    cursor = conn.cursor()

    print("\n" + "=" * 70)
    print("INDEPENDENT PRICE VERIFICATION (IPV) CONTROL ENGINE")
    print(f"Run Execution Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Policy Threshold: +/-{TOLERANCE_THRESHOLD}% | Major Threshold: >{MAJOR_SEVERITY_THRESHOLD}%")
    if target_date:
        print(f"Scope: Single Valuation Date [{target_date}]")
    else:
        print("Scope: Full 6-Month Historical Verification Range")
    print("=" * 70)

    # 1. Fetch matched prices and marks via relational join
    base_query = """
        SELECT 
            m.InstrumentID,
            i.Ticker,
            i.AssetClass,
            i.Sector,
            m.[Date],
            m.MarketPrice,
            t.TraderMark,
            t.TraderID
        FROM (DailyMarketPrice m
        INNER JOIN DailyTraderMark t ON (m.InstrumentID = t.InstrumentID AND m.[Date] = t.[Date]))
        INNER JOIN Instruments i ON m.InstrumentID = i.InstrumentID
    """
    
    if target_date:
        base_query += f" WHERE m.[Date] = #{target_date}#"
        
    base_query += " ORDER BY m.[Date] ASC, i.Ticker ASC"

    cursor.execute(base_query)
    rows = cursor.fetchall()
    total_evaluated = len(rows)

    if total_evaluated == 0:
        print("[!] No records found matching the query criteria.")
        conn.close()
        return {"total_evaluated": 0, "breaks_found": 0}

    print(f"[*] Reconciling {total_evaluated} mark-to-market position pairs...")

    # 2. Evaluate variance and severity rules
    breaks_to_insert = []
    
    for r in rows:
        inst_id = r.InstrumentID
        ticker = r.Ticker
        price_date = r.Date
        market_price = float(r.MarketPrice)
        trader_mark = float(r.TraderMark)
        trader_id = r.TraderID
        sector = r.Sector

        # Relative percentage variance formula
        # Variance% = ((TraderMark - MarketPrice) / MarketPrice) * 100
        variance_pct = round(((trader_mark - market_price) / market_price) * 100.0, 2)
        abs_variance = abs(variance_pct)

        if abs_variance > TOLERANCE_THRESHOLD:
            # Policy Breach Identified
            breach_flag = True
            if abs_variance > MAJOR_SEVERITY_THRESHOLD:
                severity = "Major"
            else:
                severity = "Minor"

            status = "Open"

            breaks_to_insert.append({
                "InstrumentID": inst_id,
                "Ticker": ticker,
                "Sector": sector,
                "TraderID": trader_id,
                "Date": price_date,
                "MarketPrice": market_price,
                "TraderMark": trader_mark,
                "VariancePct": variance_pct,
                "ToleranceThreshold": TOLERANCE_THRESHOLD,
                "BreachFlag": breach_flag,
                "Severity": severity,
                "Status": status
            })

    # 3. Clear previous breaks if performing a full back-run, or delete for target_date
    if target_date:
        cursor.execute(f"DELETE FROM ControlBreaks WHERE [Date] = #{target_date}#")
    else:
        cursor.execute("DELETE FROM ControlBreaks")
    conn.commit()

    # 4. Insert flagged breaks into ControlBreaks table
    if breaks_to_insert:
        insert_sql = """
            INSERT INTO ControlBreaks (
                InstrumentID, [Date], MarketPrice, TraderMark,
                VariancePct, ToleranceThreshold, BreachFlag, Severity, Status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        records = [
            (
                b["InstrumentID"],
                b["Date"],
                b["MarketPrice"],
                b["TraderMark"],
                b["VariancePct"],
                b["ToleranceThreshold"],
                b["BreachFlag"],
                b["Severity"],
                b["Status"]
            )
            for b in breaks_to_insert
        ]
        cursor.executemany(insert_sql, records)
        conn.commit()

    runtime_seconds = round(time.perf_counter() - start_time, 3)

    # 5. Log execution metrics into ControlRunLog
    run_timestamp = datetime.now()
    cursor.execute("""
        INSERT INTO ControlRunLog (RunTimestamp, BreaksFound, RuntimeSeconds)
        VALUES (?, ?, ?)
    """, (run_timestamp, len(breaks_to_insert), runtime_seconds))
    conn.commit()

    # 6. Generate Summary Diagnostics
    df_breaks = pd.DataFrame(breaks_to_insert) if breaks_to_insert else pd.DataFrame()
    
    print("\n[+] CONTROL EXECUTION COMPLETED")
    print(f"    - Positions Reconciled : {total_evaluated}")
    print(f"    - Clean Positions      : {total_evaluated - len(breaks_to_insert)} ({(total_evaluated - len(breaks_to_insert))/total_evaluated*100:.1f}%)")
    print(f"    - Total Breaks Flagged : {len(breaks_to_insert)} ({len(breaks_to_insert)/total_evaluated*100:.1f}%)")
    print(f"    - Engine Runtime       : {runtime_seconds} seconds")

    if not df_breaks.empty:
        minor_count = (df_breaks["Severity"] == "Minor").sum()
        major_count = (df_breaks["Severity"] == "Major").sum()
        print(f"\n    Breakdown by Severity:")
        print(f"      * Minor (3% - 5%) : {minor_count}")
        print(f"      * Major (> 5%)    : {major_count}")

        print(f"\n    Breakdown by Sector:")
        sector_counts = df_breaks.groupby("Sector")["Severity"].value_counts().unstack().fillna(0).astype(int)
        print(sector_counts.to_string())

        print(f"\n    Breakdown by Trader:")
        trader_counts = df_breaks.groupby("TraderID")["Severity"].value_counts().unstack().fillna(0).astype(int)
        print(trader_counts.to_string())

        print(f"\n    Sample Flagged Control Breaks:")
        display_sample = df_breaks[["Date", "Ticker", "TraderID", "MarketPrice", "TraderMark", "VariancePct", "Severity", "Status"]].head(6)
        # Format Date to YYYY-MM-DD
        display_sample["Date"] = pd.to_datetime(display_sample["Date"]).dt.strftime("%Y-%m-%d")
        print(display_sample.to_string(index=False))

    conn.close()

    return {
        "total_evaluated": total_evaluated,
        "breaks_found": len(breaks_to_insert),
        "runtime_seconds": runtime_seconds
    }


def main():
    parser = argparse.ArgumentParser(description="IPV Valuation Control Engine")
    parser.add_argument("--date", type=str, help="Specific date to evaluate (YYYY-MM-DD)", default=None)
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent.parent
    db_path = project_root / "data" / "valuation_control.accdb"

    if not db_path.exists():
        raise FileNotFoundError(f"Database not found at: {db_path}. Please run db_setup.py first.")

    run_control_engine(db_path, target_date=args.date)


if __name__ == "__main__":
    main()
