"""
generate_trader_marks.py
========================
Generates synthetic trader position marks based on real market prices.
Adds realistic daily valuation noise (~1.5% std dev) and injects targeted
intentional breaks (5% to 8% variance) to simulate genuine valuation control
exceptions across desks.

Domain Context (IPV / Valuation Control):
Traders mark their books daily at end-of-day. While marks usually track
market prices closely (small bid-ask and timing noise), operational issues,
stale pricing, illiquidity, or aggressive marking can create discrepancies.
We inject deliberate breaks so the automated IPV engine, exception reports,
and executive dashboards have realistic exception events to identify and escalate.
"""

from pathlib import Path
import numpy as np
import pandas as pd

# Trader coverage assignments across institutional trading desks
TRADER_MAPPING = {
    "AAPL": "TDR_NY_TECH01",   # Senior Tech Trader (New York)
    "MSFT": "TDR_NY_TECH01",   # Senior Tech Trader (New York)
    "GOOGL": "TDR_NY_TECH02",  # Junior Tech Trader (New York)
    "JPM": "TDR_NY_FIN01",     # Senior Financials Trader (New York)
    "BAC": "TDR_NY_FIN01",     # Senior Financials Trader (New York)
    "MS": "TDR_NY_FIN01",      # Senior Financials Trader (New York)
    "GS": "TDR_LN_FIN02",      # Global Financials Trader (London)
    "UBS": "TDR_LN_FIN02",     # Global Financials Trader (London)
}

# Controlled reproducibility
RANDOM_SEED = 42
BASE_NOISE_STD = 0.015  # 1.5% standard deviation for normal market noise


def generate_trader_marks(
    market_prices_path: Path,
    output_path: Path,
    seed: int = RANDOM_SEED
) -> pd.DataFrame:
    """
    Reads real market prices, applies Gaussian noise, injects intentional
    valuation breaks (5-8%), and assigns desk trader IDs.
    """
    if not market_prices_path.exists():
        raise FileNotFoundError(f"Market prices file not found at: {market_prices_path}")

    df_market = pd.read_csv(market_prices_path)
    np.random.seed(seed)

    total_rows = len(df_market)
    print(f"[*] Reading {total_rows} market price records from: {market_prices_path.name}")

    # Generate baseline Gaussian noise: N(0, 0.015^2)
    noise = np.random.normal(loc=0.0, scale=BASE_NOISE_STD, size=total_rows)

    # Calculate initial synthetic marks
    df_marks = df_market.copy()
    df_marks["Noise"] = noise
    df_marks["TraderID"] = df_marks["Ticker"].map(TRADER_MAPPING)

    # Identify distinct dates and tickers for intentional break injection
    all_dates = sorted(df_marks["Date"].unique())
    
    # Select ~20 strategic incidents across the 6-month timeline
    # We include single-name spikes, plus a multi-stock event (e.g. market volatility day)
    # to give our GenAI and Power BI analytics realistic patterns to surface.
    np.random.seed(seed + 1)
    
    # Sample 18 random distinct date-ticker pairs for isolated breaks
    sampled_indices = np.random.choice(df_marks.index, size=18, replace=False)
    
    # Inject intentional major breaks (5% to 8% magnitude, both positive and negative)
    # Positive = trader marked higher than market (aggressive P&L mark)
    # Negative = trader marked lower than market (conservative or stale mark)
    for idx in sampled_indices:
        # Magnitude between 5.2% and 8.0%
        sign = 1 if np.random.rand() > 0.45 else -1
        magnitude = np.random.uniform(0.052, 0.080)
        df_marks.loc[idx, "Noise"] = sign * magnitude

    # Also inject a correlated desk event on a specific historical date (e.g., date index 80)
    # where TDR_NY_FIN01 had a model pricing error impacting JPM, BAC, MS
    event_date = all_dates[min(80, len(all_dates) - 1)]
    fin_event_indices = df_marks[
        (df_marks["Date"] == event_date) & 
        (df_marks["TraderID"] == "TDR_NY_FIN01")
    ].index

    for idx in fin_event_indices:
        df_marks.loc[idx, "Noise"] = np.random.uniform(0.055, 0.075)

    # Compute final TraderMark = MarketPrice * (1 + Noise), rounded to 2 decimals
    df_marks["TraderMark"] = (df_marks["MarketPrice"] * (1 + df_marks["Noise"])).round(2)

    # Format into final schema: Date, Ticker, TraderMark, TraderID
    df_output = df_marks[["Date", "Ticker", "TraderMark", "TraderID"]].copy()

    # Save to CSV
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df_output.to_csv(output_path, index=False)

    # --- Verification & Diagnostics ---
    # Compute absolute variance to verify distribution
    variance_pct = ((df_output["TraderMark"] - df_market["MarketPrice"]).abs() / df_market["MarketPrice"]) * 100
    
    within_tolerance = (variance_pct < 3.0).sum()
    minor_breaks = ((variance_pct >= 3.0) & (variance_pct <= 5.0)).sum()
    major_breaks = (variance_pct > 5.0).sum()

    print(f"\n[+] Successfully generated trader marks: {output_path}")
    print(f"    - Total Marks: {len(df_output)}")
    print(f"    - In Tolerance (< 3%): {within_tolerance} ({within_tolerance/total_rows*100:.1f}%)")
    print(f"    - Minor Breaks (3% - 5%): {minor_breaks} ({minor_breaks/total_rows*100:.1f}%)")
    print(f"    - Major Breaks (> 5%): {major_breaks} ({major_breaks/total_rows*100:.1f}%)")
    print(f"    - Total Breaks Flagged: {minor_breaks + major_breaks}")
    
    print("\nSample records:")
    print(df_output.head(8).to_string(index=False))

    return df_output


def main():
    project_root = Path(__file__).resolve().parent.parent
    market_prices_path = project_root / "data" / "market_prices.csv"
    output_path = project_root / "data" / "trader_marks.csv"

    generate_trader_marks(market_prices_path, output_path)


if __name__ == "__main__":
    main()
