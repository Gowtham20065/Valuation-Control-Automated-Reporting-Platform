"""
fetch_market_prices.py
======================
Fetches 6 months of daily historical closing prices for 8 designated
equity tickers (AAPL, MSFT, GOOGL, JPM, GS, UBS, MS, BAC) via yfinance.
Outputs a clean, normalized CSV to data/market_prices.csv.

Domain Context (IPV / Valuation Control):
In investment banks, independent market data is sourced from external,
unbiased providers (e.g., Bloomberg, Refinitiv, stock exchanges) to serve
as the benchmark against which internal trading desk marks are verified.
"""

import os
from pathlib import Path
import pandas as pd
import yfinance as yf

# Institutional portfolio scope
TICKERS = ["AAPL", "MSFT", "GOOGL", "JPM", "GS", "UBS", "MS", "BAC"]
PERIOD = "6mo"


def fetch_market_prices(tickers: list[str] = TICKERS, period: str = PERIOD) -> pd.DataFrame:
    """
    Downloads historical close prices from Yahoo Finance and formats them
    into a standardized long-format DataFrame: [Date, Ticker, MarketPrice].
    """
    print(f"[*] Sourcing independent market prices for: {', '.join(tickers)}")
    print(f"[*] Historical lookback period: {period}")

    # Use threads=False to avoid local SQLite cache lock collisions in yfinance
    raw_data = yf.download(
        tickers=tickers,
        period=period,
        interval="1d",
        threads=False,
        progress=False,
        auto_adjust=False,
    )

    if raw_data.empty:
        raise ValueError("No price data retrieved from market data provider.")

    # Extract Close prices
    if "Close" in raw_data:
        close_prices = raw_data["Close"]
    else:
        raise KeyError("Expected 'Close' column not found in downloaded data.")

    # Reshape from wide (Date x Ticker) to long (Date, Ticker, MarketPrice)
    df_long = close_prices.reset_index().melt(
        id_vars=["Date"],
        var_name="Ticker",
        value_name="MarketPrice"
    )

    # Clean date formatting (YYYY-MM-DD) and remove any timezone info
    df_long["Date"] = pd.to_datetime(df_long["Date"]).dt.strftime("%Y-%m-%d")

    # Round market prices to standard 2 decimal places
    df_long["MarketPrice"] = df_long["MarketPrice"].round(2)

    # Clean missing values if any
    df_long = df_long.dropna(subset=["MarketPrice"])

    # Sort deterministically by Date and Ticker
    df_long = df_long.sort_values(by=["Date", "Ticker"]).reset_index(drop=True)

    return df_long


def main():
    # Resolve output directory
    project_root = Path(__file__).resolve().parent.parent
    data_dir = project_root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    output_path = data_dir / "market_prices.csv"

    # Fetch data
    df_prices = fetch_market_prices()

    # Save to CSV
    df_prices.to_csv(output_path, index=False)

    # Summary diagnostics
    unique_dates = df_prices["Date"].nunique()
    total_records = len(df_prices)
    min_date = df_prices["Date"].min()
    max_date = df_prices["Date"].max()

    print(f"\n[+] Successfully saved market prices to: {output_path}")
    print(f"    - Date Range: {min_date} to {max_date} ({unique_dates} trading days)")
    print(f"    - Total Records: {total_records}")
    print(f"    - Tickers Count: {df_prices['Ticker'].nunique()} ({', '.join(df_prices['Ticker'].unique())})")
    print("\nSample records:")
    print(df_prices.head(8).to_string(index=False))


if __name__ == "__main__":
    main()
