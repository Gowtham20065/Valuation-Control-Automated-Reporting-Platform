# Power BI Executive Dashboard — Step-by-Step Build Guide

This guide walks you through building the **Independent Price Verification (IPV) Executive Monitoring Dashboard** in Power BI Desktop from scratch.

---

## Step 1: Connect to the Data Source

You can connect Power BI Desktop to the project in either of two ways:

### Option A: Direct MS Access Connection (Recommended for Relational Modeling)
1. Open **Power BI Desktop**.
2. Click **Get Data** $\rightarrow$ **More...** $\rightarrow$ **Access database** $\rightarrow$ **Connect**.
3. Browse to:  
   `Valuation-Control-Automated-Reporting-Platform\data\valuation_control.accdb`
4. In the Navigator window, select the following 4 tables:
   - `Instruments`
   - `DailyMarketPrice`
   - `DailyTraderMark`
   - `ControlBreaks`
5. Click **Load**.

### Option B: Quick CSV Import (Alternative)
1. Click **Get Data** $\rightarrow$ **Text/CSV**.
2. Select `data/powerbi_reconciliation_flat.csv` and `data/powerbi_calendar.csv`.
3. Click **Load**.

---

## Step 2: Establish the Star Schema Data Model

In Power BI Desktop, click the **Model view** icon on the left navigation bar:

1. **Verify Relationships:**
   - `Instruments[InstrumentID]` $\rightarrow$ `DailyMarketPrice[InstrumentID]` (1-to-many, Single cross-filter)
   - `Instruments[InstrumentID]` $\rightarrow$ `DailyTraderMark[InstrumentID]` (1-to-many, Single cross-filter)
   - `Instruments[InstrumentID]` $\rightarrow$ `ControlBreaks[InstrumentID]` (1-to-many, Single cross-filter)
2. **Add a Calendar Dimension (Date Table):**
   - Click **Modeling** tab $\rightarrow$ **New Table**, and enter:
     ```dax
     Calendar = 
     ADDCOLUMNS(
         CALENDAR(MIN(DailyMarketPrice[Date]), MAX(DailyMarketPrice[Date])),
         "Year", YEAR([Date]),
         "Month", FORMAT([Date], "MMM"),
         "MonthYear", FORMAT([Date], "MMM YYYY"),
         "MonthNumber", MONTH([Date]),
         "DayOfWeek", FORMAT([Date], "ddd")
     )
     ```
   - Connect `Calendar[Date]` $\rightarrow$ `DailyMarketPrice[Date]`.
   - Connect `Calendar[Date]` $\rightarrow$ `ControlBreaks[Date]`.

---

## Step 3: Create the DAX Measures Table

To keep the model clean and professional:
1. On the **Home** tab, click **Enter Data**.
2. Name the table `_Measures` and click **Load**.
3. In the `_Measures` table, right-click and select **New Measure** for each formula:

### Key Measures to Create:
1. **Total Positions Evaluated:**
   ```dax
   Total Positions Evaluated = COUNTROWS(DailyMarketPrice)
   ```
2. **Total Control Breaks:**
   ```dax
   Total Control Breaks = COUNTROWS(ControlBreaks)
   ```
3. **Break Rate %:**
   ```dax
   Break Rate % = DIVIDE([Total Control Breaks], [Total Positions Evaluated], 0)
   ```
   *(Set format to Percentage `0.0%`)*
4. **Major Breaches Count:**
   ```dax
   Major Breaches Count = CALCULATE(COUNTROWS(ControlBreaks), ControlBreaks[Severity] = "Major")
   ```
5. **Minor Breaches Count:**
   ```dax
   Minor Breaches Count = CALCULATE(COUNTROWS(ControlBreaks), ControlBreaks[Severity] = "Minor")
   ```
6. **Variance %:**
   ```dax
   Variance % = 
   DIVIDE(
       SUM(DailyTraderMark[TraderMark]) - SUM(DailyMarketPrice[MarketPrice]),
       SUM(DailyMarketPrice[MarketPrice]),
       0
   )
   ```
7. **7-Day Rolling Avg Breaks:**
   ```dax
   7-Day Rolling Avg Breaks = 
   VAR CurrentDate = MAX('Calendar'[Date])
   VAR RollingPeriod = DATESBETWEEN('Calendar'[Date], CurrentDate - 6, CurrentDate)
   RETURN
       DIVIDE(CALCULATE(COUNTROWS(ControlBreaks), RollingPeriod), 7, 0)
   ```

---

## Step 4: Visual Canvas Layout & Design

Set the canvas background to `#F8FAFC` (subtle off-white) or clean white.

### Visual 1: Header Banner (Top, Y=0, Height=60)
- Add a **Text Box** across the top:
  - Text: **VALUATION CONTROL & IPV EXECUTIVE MONITORING DASHBOARD**
  - Font: 16pt Bold, Navy `#1E3A8A`.
  - Subtitle: *Portfolio Scope: 8 Equities (Tech & Financials) | Benchmark: Exchange Closes | Governance: UBS Group Finance*

---

### Visual 2: Top KPI Ribbon (Rows 2, 4 Cards)
Use **Card (New)** or 4 standard **Card** visuals side-by-side:
1. **Card 1:** `Total Positions Evaluated` $\rightarrow$ **1,008**
2. **Card 2:** `Total Control Breaks` $\rightarrow$ **62**
3. **Card 3:** `Break Rate %` $\rightarrow$ **6.2%** (Callout color: `#1E3A8A`)
4. **Card 4:** `Major Breaches Count` $\rightarrow$ **22** (Callout color: `#B91C1C` Crimson Red)

---

### Visual 3: Trend Line & Column Chart (Center-Left)
- **Visual Type:** **Line and Clustered Column Chart**
- **X-Axis:** `Calendar[Date]`
- **Column Y-Axis:** `Total Control Breaks` (Bar color: `#94A3B8` Slate)
- **Line Y-Axis:** `7-Day Rolling Avg Breaks` (Line color: `#1E3A8A` Navy Blue, Stroke Width = 3)
- **Title:** *Daily Control Breaks & 7-Day Rolling Average*
- **Key Observation:** Shows the acute spike on **2026-07-24** (4 breaks on the Financials incident) and the rolling surge.

---

### Visual 4: Sector & Severity Breakdown (Center-Right)
- **Visual Type:** **Clustered Bar Chart**
- **Y-Axis:** `Instruments[Sector]` (`Financials`, `Technology`)
- **X-Axis:** `Total Control Breaks`
- **Legend:** `ControlBreaks[Severity]` (`Major`, `Minor`)
- **Colors:**
  - Major: `#EF4444` (Red)
  - Minor: `#F59E0B` (Amber)
- **Title:** *Exceptions by Sector & Severity Tier*
- **Key Observation:** Financials account for 36 breaks (12 Major, 24 Minor) vs. Technology with 26 breaks (10 Major, 16 Minor).

---

### Visual 5: Severity Distribution Donut Chart (Bottom-Left)
- **Visual Type:** **Donut Chart**
- **Legend:** `ControlBreaks[Severity]`
- **Values:** `Total Control Breaks`
- **Colors:**
  - Minor: `#F59E0B` (64.5% - 40 breaks)
  - Major: `#EF4444` (35.5% - 22 breaks)
- **Title:** *Severity Tier Distribution*

---

### Visual 6: Detailed Exception Ledger Table (Bottom-Right)
- **Visual Type:** **Table**
- **Columns:**
  - `ControlBreaks[Date]`
  - `Instruments[Ticker]`
  - `Instruments[Sector]`
  - `DailyTraderMark[TraderID]`
  - `ControlBreaks[MarketPrice]` (Format as `$#,##0.00`)
  - `ControlBreaks[TraderMark]` (Format as `$#,##0.00`)
  - `ControlBreaks[VariancePct]` (Format as `+0.00%;-0.00%`)
  - `ControlBreaks[Severity]`
- **Title:** *Exceptions Drill-Down Ledger*

---

### Visual 7: Interactive Slicers (Top-Right or Left Sidebar)
- **Date Slicer:** `Calendar[Date]` (Between slider)
- **Sector Slicer:** `Instruments[Sector]` (Dropdown: All, Financials, Technology)
- **Trader ID Slicer:** `DailyTraderMark[TraderID]` (Dropdown)

---

## Step 5: Verification Checklist

When your dashboard is built, cross-verify against these figures:
- [x] Total Positions: **1,008**
- [x] Total Breaks: **62**
- [x] Break Rate: **6.15%**
- [x] Major Breaches: **22**
- [x] Minor Breaches: **40**
- [x] Selecting date `2026-07-24` highlights 4 exceptions (3 Major on Financials, 1 Minor on Tech).
