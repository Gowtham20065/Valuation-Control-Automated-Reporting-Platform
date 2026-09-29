"""
genai_summary.py
================
Generative AI Proof-of-Concept (PoC) for Valuation Control & IPV.

Uses Google Gemini (gemini-flash-latest via google-genai SDK) to generate
a 3-4 sentence plain-English executive briefing for a Valuation Control
Manager, synthesizing daily exceptions, identifying correlated sector/trader
patterns, and recommending escalation actions.

Responsible AI Banking Framework:
- Framed as a Human-in-the-Loop (HITL) Decision-Support Pilot (SR 11-7 compliant).
- Low temperature (0.2) to prevent numerical hallucination.
- Multi-tier resilience: primary model alias 'gemini-flash-latest', failover to
  'gemini-flash-lite-latest', and deterministic rule-based analytical baseline.
"""

import argparse
import json
import os
from pathlib import Path
import warnings
import pandas as pd
import pyodbc

# Suppress minor library deprecation/AFC warnings
warnings.filterwarnings("ignore", category=UserWarning)

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

PRIMARY_MODEL = "gemini-flash-latest"
FAILOVER_MODEL = "gemini-flash-lite-latest"


def get_db_connection(db_path: Path) -> pyodbc.Connection:
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str, autocommit=True)


def fetch_breaks_for_date(db_path: Path, target_date: str) -> pd.DataFrame:
    conn = get_db_connection(db_path)
    query = f"""
        SELECT 
            b.BreakID,
            i.Ticker,
            i.AssetClass,
            i.Sector,
            Format(b.[Date], 'YYYY-MM-DD') AS ValuationDate,
            b.MarketPrice,
            b.TraderMark,
            b.VariancePct,
            b.Severity,
            t.TraderID,
            b.Status
        FROM ((ControlBreaks b
        INNER JOIN Instruments i ON b.InstrumentID = i.InstrumentID)
        INNER JOIN DailyTraderMark t ON (b.InstrumentID = t.InstrumentID AND b.[Date] = t.[Date]))
        WHERE b.[Date] = #{target_date}#
        ORDER BY ABS(b.VariancePct) DESC
    """
    df = pd.read_sql(query, conn)
    conn.close()
    return df


def generate_rule_based_summary(df: pd.DataFrame, target_date: str) -> str:
    """
    Deterministic analytical fallback engine (Zero-Hallucination Baseline).
    Produces an exact, rule-based 3-4 sentence briefing when API is unavailable.
    """
    total_breaks = len(df)
    major_count = (df["Severity"] == "Major").sum()
    minor_count = (df["Severity"] == "Minor").sum()

    # Outlier analysis
    max_idx = df["VariancePct"].abs().idxmax()
    outlier = df.loc[max_idx]

    # Concentration analysis
    top_sector = df["Sector"].value_counts().index[0]
    top_sector_count = df["Sector"].value_counts().iloc[0]
    top_trader = df["TraderID"].value_counts().index[0]
    top_trader_count = df["TraderID"].value_counts().iloc[0]

    s1 = (
        f"For valuation date {target_date}, the IPV control engine flagged {total_breaks} total exceptions, "
        f"comprising {major_count} Major (>5.0%) and {minor_count} Minor (3.0%–5.0%) breaches."
    )
    
    if top_trader_count > 1:
        s2 = (
            f"A high-risk concentration pattern was detected in the {top_sector} book, where trader {top_trader} "
            f"accounted for {top_trader_count} of the {total_breaks} breaks with correlated aggressive marks."
        )
    else:
        s2 = f"Breaches were primarily localized in the {top_sector} sector across individual desk positions."

    s3 = (
        f"The single largest valuation break was observed on {outlier['Ticker']} with a {outlier['VariancePct']:+.2f}% "
        f"variance (Trader Mark: ${outlier['TraderMark']:.2f} vs. Independent Market: ${outlier['MarketPrice']:.2f})."
    )

    s4 = (
        f"Product Control recommends immediate escalation to the {top_sector} Desk Head to challenge pricing models "
        f"and initiate formal justification reviews before market open."
    )

    return f"{s1} {s2} {s3} {s4}"


def generate_gemini_summary(df: pd.DataFrame, target_date: str, api_key: str | None = None) -> tuple[str, str]:
    """
    Calls Google Gemini using the google-genai SDK.
    Attempts primary model 'gemini-flash-latest', fails over to 'gemini-flash-lite-latest',
    and gracefully falls back to deterministic rule-based baseline if API is unavailable.
    Returns (summary_text, source_mode).
    """
    key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")

    if not key or not GENAI_AVAILABLE:
        return generate_rule_based_summary(df, target_date), "Analytical Rule-Based Baseline (Offline / No Key)"

    records = df[[
        "Ticker", "Sector", "TraderID", "MarketPrice", "TraderMark", "VariancePct", "Severity"
    ]].to_dict(orient="records")
    
    prompt_payload = json.dumps({
        "valuation_date": target_date,
        "total_breaks": len(df),
        "breaks": records
    }, indent=2)

    system_instruction = (
        "You are an Executive Valuation Control Analyst in Group Finance at a global investment bank (e.g., UBS). "
        "Your role is to write a concise, professional, exactly 3 to 4 sentence executive briefing for the Valuation Committee "
        "and Trading Desk Head summarizing today's flagged IPV control breaks.\n\n"
        "Strict Requirements:\n"
        "1. Sentence 1: State the total breaks flagged and the breakdown between Major (>5%) and Minor (3-5%) severity.\n"
        "2. Sentence 2: Identify structural patterns, such as trader concentration, sector correlation (e.g. Financials), or marking bias (+/-).\n"
        "3. Sentence 3: Highlight the single largest outlier breach (ticker name and exact variance percentage).\n"
        "4. Sentence 4: Provide an actionable escalation recommendation for Product Control.\n"
        "Constraint: Exactly 3 to 4 sentences. Highly professional financial tone. Never invent numbers not present in the payload."
    )

    client = genai.Client(api_key=key)

    # Attempt primary model, then failover model
    for model_name in [PRIMARY_MODEL, FAILOVER_MODEL]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=f"Summarize the following valuation control exceptions according to your instructions:\n\n{prompt_payload}",
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    temperature=0.2,
                    max_output_tokens=300,
                )
            )
            if response.text and response.text.strip():
                return response.text.strip(), f"Google Gemini Live API ({model_name})"
        except Exception as e:
            print(f"[!] Info: Model {model_name} returned temporary status ({e}). Attempting next tier...")

    # Final fallback to deterministic analytical baseline
    print("[!] Active API models unavailable. Activating deterministic analytical baseline.")
    return generate_rule_based_summary(df, target_date), "Analytical Rule-Based Baseline (Fallback)"


def main():
    parser = argparse.ArgumentParser(description="GenAI IPV Exception Summary Briefing (Proof-of-Concept)")
    parser.add_argument("--date", type=str, default="2026-07-24", help="Valuation date (YYYY-MM-DD)")
    parser.add_argument("--api-key", type=str, default=None, help="Google Gemini API Key (optional)")
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent.parent
    db_path = project_root / "data" / "valuation_control.accdb"
    reports_dir = project_root / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    print("\n" + "=" * 75)
    print("VALUATION CONTROL & IPV — GENAI EXECUTIVE SUMMARY (PROOF-OF-CONCEPT)")
    print(f"Valuation Date: {args.date} | Target Model: {PRIMARY_MODEL}")
    print("Governance Context: Human-in-the-Loop (HITL) Decision-Support Pilot (SR 11-7)")
    print("=" * 75)

    df_breaks = fetch_breaks_for_date(db_path, args.date)

    if df_breaks.empty:
        print(f"[*] No valuation control breaks identified for date: {args.date}.")
        print("[+] All trader marks passed within +/-3.0% policy threshold. No briefing required.")
        return

    print(f"[*] Analyzing {len(df_breaks)} flagged exceptions for date: {args.date}...")

    summary_text, source_mode = generate_gemini_summary(df_breaks, args.date, args.api_key)

    print(f"\n[+] Synthesis Engine: {source_mode}")
    print("-" * 75)
    print("EXECUTIVE BRIEFING FOR VALUATION COMMITTEE & DESK HEAD:")
    print("-" * 75)
    print(summary_text)
    print("-" * 75)

    # Save to text report
    out_file = reports_dir / f"GenAI_Summary_{args.date}.txt"
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(f"VALUATION CONTROL & IPV EXECUTIVE BRIEFING\n")
        f.write(f"Valuation Date: {args.date}\n")
        f.write(f"Engine: {source_mode}\n")
        f.write(f"Timestamp: {pd.Timestamp.now()}\n")
        f.write("=" * 60 + "\n\n")
        f.write(summary_text + "\n")

    print(f"\n[+] Saved executive briefing to: {out_file.name}")
    print("\n[i] Responsible AI Governance Note:")
    print("    In accordance with Bank Model Risk Management (SR 11-7) principles, this summary")
    print("    serves as decision-support commentary. A designated Product Controller must review")
    print("    and verify all figures against the golden source database before formal committee submission.\n")


if __name__ == "__main__":
    main()
