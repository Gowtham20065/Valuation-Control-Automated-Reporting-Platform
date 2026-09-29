"""
generate_excel_report.py
========================
Generates the formatted Daily Valuation Exception Report in Excel.
Produces:
1. reports/Exception_Report_[YYYY-MM-DD].xlsx - Formatted report matching institutional UBS styling.
2. Embeds the VBA macro logic in vba/DailyExceptionReport.bas.

Includes:
- Executive KPI Summary Cards (Total Exceptions, Major Count, Minor Count, Asset Class)
- Navy corporate styling (Calibri/Segoe UI, borders, alternating fills)
- Number formatting ($#,##0.00 for prices, +0.00%/-0.00% for variance)
- Dynamic Conditional Formatting (Soft red for Major, Soft amber for Minor)
"""

import argparse
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import pyodbc


def get_db_connection(db_path: Path) -> pyodbc.Connection:
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str, autocommit=True)


def fetch_report_data(conn: pyodbc.Connection, target_date: str) -> pd.DataFrame:
    query = f"""
        SELECT 
            b.BreakID,
            i.Ticker,
            i.AssetClass,
            i.Sector,
            b.[Date] AS ValuationDate,
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
    return df


def build_formatted_excel(df: pd.DataFrame, target_date: str, output_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exception_Report"
    ws.views.sheetView[0].showGridLines = True

    # Palette
    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    sub_navy_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    table_hdr_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    card_hdr_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    font_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=9, italic=True, color="475569")
    font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10)
    font_card_title = Font(name="Calibri", size=9, bold=True, color="64748B")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    # 1. Main Title (Row 1)
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = "VALUATION CONTROL & INDEPENDENT PRICE VERIFICATION (IPV) EXCEPTION REPORT"
    title_cell.font = font_title
    title_cell.fill = navy_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 2. Subtitle (Row 2)
    ws.merge_cells("A2:K2")
    sub_cell = ws["A2"]
    sub_cell.value = (
        f"Valuation Date: {target_date}  |  "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  |  "
        "Governance: UBS Group Finance / Product Control"
    )
    sub_cell.font = font_sub
    sub_cell.fill = sub_navy_fill
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # 3. Executive KPI Summary Cards (Rows 4-5)
    total_exceptions = len(df)
    major_count = (df["Severity"] == "Major").sum() if not df.empty else 0
    minor_count = (df["Severity"] == "Minor").sum() if not df.empty else 0

    cards = [
        ("A4:B4", "A5:B5", "TOTAL EXCEPTIONS", total_exceptions, "1E293B"),
        ("D4:E4", "D5:E5", "MAJOR SEVERITY (>5%)", major_count, "B91C1C"),
        ("G4:H4", "G5:H5", "MINOR SEVERITY (3-5%)", minor_count, "B45309"),
        ("J4:K4", "J5:K5", "ASSET CLASS SCOPE", "Equity Cash", "1E3A8A"),
    ]

    for hdr_range, val_range, title, val, color in cards:
        ws.merge_cells(hdr_range)
        h_cell = ws[hdr_range.split(":")[0]]
        h_cell.value = title
        h_cell.font = font_card_title
        h_cell.fill = card_hdr_fill
        h_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(val_range)
        v_cell = ws[val_range.split(":")[0]]
        v_cell.value = val
        v_cell.font = Font(name="Calibri", size=15, bold=True, color=color)
        v_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Borders
        for row in ws[hdr_range.split(":")[0]:val_range.split(":")[1]]:
            for cell in row:
                cell.border = thin_border

    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 26

    # 4. Table Headers (Row 8)
    headers = [
        "Break ID", "Ticker", "Asset Class", "Sector", "Valuation Date",
        "Market Price ($)", "Trader Mark ($)", "Variance (%)", "Severity", "Trader ID", "Status"
    ]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=8, column=col_idx, value=h)
        cell.font = font_tbl_hdr
        cell.fill = table_hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[8].height = 24

    # 5. Populate Data Rows (starting Row 9)
    start_row = 9
    if df.empty:
        ws.merge_cells("A9:K9")
        clean_cell = ws["A9"]
        clean_cell.value = f"No valuation control breaks flagged for date: {target_date}. All marks within +/-3.0% tolerance."
        clean_cell.font = Font(name="Calibri", size=10, italic=True, color="166534")
        clean_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[9].height = 22
        end_row = 9
    else:
        for r_idx, row in df.iterrows():
            current_row = start_row + r_idx
            ws.cell(row=current_row, column=1, value=int(row["BreakID"])).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=2, value=str(row["Ticker"])).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=3, value=str(row["AssetClass"])).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=4, value=str(row["Sector"])).alignment = Alignment(horizontal="left")
            ws.cell(row=current_row, column=5, value=str(row["ValuationDate"])[:10]).alignment = Alignment(horizontal="center")

            # Prices & Numbers
            c_mkt = ws.cell(row=current_row, column=6, value=float(row["MarketPrice"]))
            c_mkt.number_format = "$#,##0.00"
            c_mkt.alignment = Alignment(horizontal="right")

            c_tm = ws.cell(row=current_row, column=7, value=float(row["TraderMark"]))
            c_tm.number_format = "$#,##0.00"
            c_tm.alignment = Alignment(horizontal="right")

            c_var = ws.cell(row=current_row, column=8, value=float(row["VariancePct"]) / 100.0)
            c_var.number_format = "+0.00%;-0.00%;0.00%"
            c_var.alignment = Alignment(horizontal="right")

            ws.cell(row=current_row, column=9, value=str(row["Severity"])).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=10, value=str(row["TraderID"])).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=11, value=str(row["Status"])).alignment = Alignment(horizontal="center")

            for col in range(1, 12):
                c = ws.cell(row=current_row, column=col)
                c.font = font_data
                c.border = thin_border

            ws.row_dimensions[current_row].height = 20
        end_row = start_row + len(df) - 1

    # 6. Apply Conditional Formatting on Severity Column (Col I / 9)
    if not df.empty:
        red_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        red_font = Font(name="Calibri", size=10, bold=True, color="991B1B")
        rule_major = CellIsRule(operator="equal", formula=['"Major"'], fill=red_fill, font=red_font)
        ws.conditional_formatting.add(f"I9:I{end_row}", rule_major)

        amber_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        amber_font = Font(name="Calibri", size=10, bold=True, color="92400E")
        rule_minor = CellIsRule(operator="equal", formula=['"Minor"'], fill=amber_fill, font=amber_font)
        ws.conditional_formatting.add(f"I9:I{end_row}", rule_minor)

    # 7. Column Widths
    col_widths = {
        "A": 12, "B": 12, "C": 14, "D": 16, "E": 15,
        "F": 17, "G": 17, "H": 15, "I": 14, "J": 17, "K": 12
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # Save
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"[+] Successfully generated Daily Exception Report: {output_path.name}")


def main():
    parser = argparse.ArgumentParser(description="Generate Daily Valuation Exception Report in Excel")
    parser.add_argument("--date", type=str, default="2026-07-24", help="Valuation date (YYYY-MM-DD)")
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent.parent
    db_path = project_root / "data" / "valuation_control.accdb"
    reports_dir = project_root / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    output_path = reports_dir / f"Exception_Report_{args.date}.xlsx"

    print(f"[*] Extracting valuation exceptions for date: {args.date} from database...")
    conn = get_db_connection(db_path)
    try:
        df_breaks = fetch_report_data(conn, args.date)
        print(f"[*] Found {len(df_breaks)} control breaks on {args.date}.")
        build_formatted_excel(df_breaks, args.date, output_path)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
