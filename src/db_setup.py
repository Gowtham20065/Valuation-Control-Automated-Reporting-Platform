"""
db_setup.py
===========
Initializes the MS Access relational database (data/valuation_control.accdb)
for the Valuation Control & Automated Reporting Platform.

Creates schema according to 3NF relational standards:
- Instruments (Golden Source / Master Reference Data)
- DailyMarketPrice (External Independent Benchmark Feed)
- DailyTraderMark (Front Office Position Marks)
- ControlBreaks (Exception Ledger populated by IPV Engine)
- ControlRunLog (Audit Trail for job executions)

Populates master reference data and ingests historical datasets.
"""

from pathlib import Path
import subprocess
import pandas as pd
import pyodbc

# Master Reference Data (Golden Source)
INSTRUMENTS_SEED = [
    {"Ticker": "AAPL", "AssetClass": "Equity", "Sector": "Technology"},
    {"Ticker": "MSFT", "AssetClass": "Equity", "Sector": "Technology"},
    {"Ticker": "GOOGL", "AssetClass": "Equity", "Sector": "Technology"},
    {"Ticker": "JPM", "AssetClass": "Equity", "Sector": "Financials"},
    {"Ticker": "GS", "AssetClass": "Equity", "Sector": "Financials"},
    {"Ticker": "UBS", "AssetClass": "Equity", "Sector": "Financials"},
    {"Ticker": "MS", "AssetClass": "Equity", "Sector": "Financials"},
    {"Ticker": "BAC", "AssetClass": "Equity", "Sector": "Financials"},
]


def create_access_database(db_path: Path):
    """
    Creates an empty .accdb database file via Windows ADOX Catalog COM object.
    """
    if db_path.exists():
        print(f"[*] Found existing database at: {db_path}")
        return

    print(f"[*] Initializing new MS Access database at: {db_path}")
    db_path.parent.mkdir(parents=True, exist_ok=True)

    # Use PowerShell to invoke ADOX.Catalog to create a valid, uncorrupted .accdb container
    ps_cmd = (
        f"$catalog = New-Object -ComObject ADOX.Catalog; "
        f"$catalog.Create('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{db_path}\";'); "
        f"[System.Runtime.Interopservices.Marshal]::ReleaseComObject($catalog) | Out-Null"
    )

    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0 and not db_path.exists():
        raise RuntimeError(f"Failed to create Access database file: {result.stderr}")

    print(f"[+] Successfully created blank database container: {db_path.name}")


def get_connection(db_path: Path) -> pyodbc.Connection:
    """
    Establishes an ODBC connection to the MS Access database.
    """
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str, autocommit=False)


def create_tables(conn: pyodbc.Connection):
    """
    Creates the relational schema with Primary and Foreign Keys.
    """
    cursor = conn.cursor()
    print("[*] Creating relational tables and integrity constraints...")

    # Drop existing tables in reverse dependency order if rebuilding
    existing_tables = [table.table_name for table in cursor.tables(tableType="TABLE")]
    for table_name in ["ControlBreaks", "DailyTraderMark", "DailyMarketPrice", "Instruments", "ControlRunLog"]:
        if table_name in existing_tables:
            cursor.execute(f"DROP TABLE [{table_name}]")
    conn.commit()

    # 1. Instruments (Master Reference Data / Golden Source)
    cursor.execute("""
        CREATE TABLE Instruments (
            InstrumentID AUTOINCREMENT PRIMARY KEY,
            Ticker VARCHAR(20) NOT NULL,
            AssetClass VARCHAR(50) NOT NULL,
            Sector VARCHAR(50) NOT NULL
        )
    """)

    # 2. DailyMarketPrice (Independent External Feed)
    cursor.execute("""
        CREATE TABLE DailyMarketPrice (
            PriceID AUTOINCREMENT PRIMARY KEY,
            InstrumentID INTEGER NOT NULL,
            [Date] DATETIME NOT NULL,
            MarketPrice DOUBLE NOT NULL,
            CONSTRAINT FK_Market_Instrument FOREIGN KEY (InstrumentID) REFERENCES Instruments(InstrumentID)
        )
    """)

    # 3. DailyTraderMark (Front Office Desk Marks)
    cursor.execute("""
        CREATE TABLE DailyTraderMark (
            MarkID AUTOINCREMENT PRIMARY KEY,
            InstrumentID INTEGER NOT NULL,
            [Date] DATETIME NOT NULL,
            TraderMark DOUBLE NOT NULL,
            TraderID VARCHAR(50) NOT NULL,
            CONSTRAINT FK_Trader_Instrument FOREIGN KEY (InstrumentID) REFERENCES Instruments(InstrumentID)
        )
    """)

    # 4. ControlBreaks (Exception Ledger populated by IPV Engine)
    cursor.execute("""
        CREATE TABLE ControlBreaks (
            BreakID AUTOINCREMENT PRIMARY KEY,
            InstrumentID INTEGER NOT NULL,
            [Date] DATETIME NOT NULL,
            MarketPrice DOUBLE NOT NULL,
            TraderMark DOUBLE NOT NULL,
            VariancePct DOUBLE NOT NULL,
            ToleranceThreshold DOUBLE NOT NULL,
            BreachFlag YESNO NOT NULL,
            Severity VARCHAR(20) NOT NULL,
            Status VARCHAR(20) NOT NULL,
            CONSTRAINT FK_Breaks_Instrument FOREIGN KEY (InstrumentID) REFERENCES Instruments(InstrumentID)
        )
    """)

    # 5. ControlRunLog (Job Execution Audit Trail)
    cursor.execute("""
        CREATE TABLE ControlRunLog (
            RunID AUTOINCREMENT PRIMARY KEY,
            RunTimestamp DATETIME NOT NULL,
            BreaksFound INTEGER NOT NULL,
            RuntimeSeconds DOUBLE NOT NULL
        )
    """)

    conn.commit()
    print("[+] All 5 relational tables created successfully.")


def seed_instruments(conn: pyodbc.Connection) -> dict[str, int]:
    """
    Populates Instruments table and returns a mapping of {Ticker: InstrumentID}.
    """
    cursor = conn.cursor()
    print("[*] Seeding Instruments master reference data...")

    for inst in INSTRUMENTS_SEED:
        cursor.execute(
            "INSERT INTO Instruments (Ticker, AssetClass, Sector) VALUES (?, ?, ?)",
            (inst["Ticker"], inst["AssetClass"], inst["Sector"])
        )
    conn.commit()

    # Query back generated InstrumentIDs
    cursor.execute("SELECT InstrumentID, Ticker FROM Instruments")
    ticker_to_id = {row.Ticker: row.InstrumentID for row in cursor.fetchall()}
    print(f"[+] Loaded {len(ticker_to_id)} instruments: {ticker_to_id}")
    return ticker_to_id


def load_market_prices(conn: pyodbc.Connection, csv_path: Path, ticker_to_id: dict[str, int]):
    """
    Loads market_prices.csv into DailyMarketPrice table with InstrumentID foreign keys.
    """
    print(f"[*] Ingesting market prices from: {csv_path.name}")
    df = pd.read_csv(csv_path)
    df["InstrumentID"] = df["Ticker"].map(ticker_to_id)

    if df["InstrumentID"].isna().any():
        unmapped = df[df["InstrumentID"].isna()]["Ticker"].unique()
        raise ValueError(f"Found unmapped tickers in market prices: {unmapped}")

    cursor = conn.cursor()
    records = [
        (int(row.InstrumentID), str(row.Date), float(row.MarketPrice))
        for row in df.itertuples(index=False)
    ]

    cursor.executemany(
        "INSERT INTO DailyMarketPrice (InstrumentID, [Date], MarketPrice) VALUES (?, ?, ?)",
        records
    )
    conn.commit()
    print(f"[+] Ingested {len(records)} records into DailyMarketPrice.")


def load_trader_marks(conn: pyodbc.Connection, csv_path: Path, ticker_to_id: dict[str, int]):
    """
    Loads trader_marks.csv into DailyTraderMark table with InstrumentID foreign keys.
    """
    print(f"[*] Ingesting trader marks from: {csv_path.name}")
    df = pd.read_csv(csv_path)
    df["InstrumentID"] = df["Ticker"].map(ticker_to_id)

    if df["InstrumentID"].isna().any():
        unmapped = df[df["InstrumentID"].isna()]["Ticker"].unique()
        raise ValueError(f"Found unmapped tickers in trader marks: {unmapped}")

    cursor = conn.cursor()
    records = [
        (int(row.InstrumentID), str(row.Date), float(row.TraderMark), str(row.TraderID))
        for row in df.itertuples(index=False)
    ]

    cursor.executemany(
        "INSERT INTO DailyTraderMark (InstrumentID, [Date], TraderMark, TraderID) VALUES (?, ?, ?, ?)",
        records
    )
    conn.commit()
    print(f"[+] Ingested {len(records)} records into DailyTraderMark.")


def verify_database(conn: pyodbc.Connection):
    """
    Runs diagnostic queries to verify data integrity and foreign key joins.
    """
    cursor = conn.cursor()
    print("\n" + "=" * 60)
    print("DATABASE INTEGRITY & VERIFICATION REPORT")
    print("=" * 60)

    tables = ["Instruments", "DailyMarketPrice", "DailyTraderMark", "ControlBreaks", "ControlRunLog"]
    for t in tables:
        cursor.execute(f"SELECT COUNT(*) FROM [{t}]")
        count = cursor.fetchone()[0]
        print(f"  Table [{t:18}]: {count:>6} rows")

    # Verify relational join between Market Prices and Instruments
    cursor.execute("""
        SELECT TOP 3 i.Ticker, i.Sector, m.[Date], m.MarketPrice
        FROM DailyMarketPrice m
        INNER JOIN Instruments i ON m.InstrumentID = i.InstrumentID
        ORDER BY m.[Date] ASC, i.Ticker ASC
    """)
    print("\nSample Relational Join (DailyMarketPrice -> Instruments):")
    for r in cursor.fetchall():
        print(f"  Ticker: {r[0]:<6} | Sector: {r[1]:<12} | Date: {str(r[2])[:10]} | MarketPrice: ${r[3]:.2f}")

    # Verify relational join between Trader Marks and Instruments
    cursor.execute("""
        SELECT TOP 3 i.Ticker, t.TraderID, t.[Date], t.TraderMark
        FROM DailyTraderMark t
        INNER JOIN Instruments i ON t.InstrumentID = i.InstrumentID
        ORDER BY t.[Date] ASC, i.Ticker ASC
    """)
    print("\nSample Relational Join (DailyTraderMark -> Instruments):")
    for r in cursor.fetchall():
        print(f"  Ticker: {r[0]:<6} | Trader: {r[1]:<14} | Date: {str(r[2])[:10]} | TraderMark: ${r[3]:.2f}")

    print("=" * 60 + "\n")


def main():
    project_root = Path(__file__).resolve().parent.parent
    data_dir = project_root / "data"
    db_path = data_dir / "valuation_control.accdb"
    market_prices_csv = data_dir / "market_prices.csv"
    trader_marks_csv = data_dir / "trader_marks.csv"

    # Step 1: Ensure MS Access database exists
    create_access_database(db_path)

    # Step 2: Connect via ODBC
    conn = get_connection(db_path)

    try:
        # Step 3: Create schema
        create_tables(conn)

        # Step 4: Seed Instruments reference data
        ticker_to_id = seed_instruments(conn)

        # Step 5: Ingest Market Prices & Trader Marks
        load_market_prices(conn, market_prices_csv, ticker_to_id)
        load_trader_marks(conn, trader_marks_csv, ticker_to_id)

        # Step 6: Verify integrity
        verify_database(conn)

    finally:
        conn.close()
        print("[+] Database connection closed.")


if __name__ == "__main__":
    main()
