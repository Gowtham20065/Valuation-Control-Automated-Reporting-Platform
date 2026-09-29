# DAX Measures Reference Guide — IPV & Valuation Control Dashboard

This document provides the exact DAX formulas for the Power BI executive dashboard, complete with their evaluation context, institutional rationale, and interview explanations.

---

### Measure 1: Variance % (Relative Mark-to-Market Variance)

```dax
Variance % = 
DIVIDE(
    SUM(DailyTraderMark[TraderMark]) - SUM(DailyMarketPrice[MarketPrice]),
    SUM(DailyMarketPrice[MarketPrice]),
    0
)
```

- **Format:** Percentage (`+0.00%;-0.00%;0.00%`)
- **Institutional Context:**  
  Measures the aggregate portfolio-level marking bias. A persistent positive variance across months indicates that trading desks are systematically marking above market consensus.
- **DAX Mechanics & Interview Talking Point:**  
  > *"I used the `DIVIDE()` function rather than the mathematical `/` operator. `DIVIDE()` handles division-by-zero internally, returning 0 or blank rather than raising an unhandled `#DIV/0!` exception on cards or charts if market price is missing."*

---

### Measure 2: 7-Day Rolling Average Breaks

```dax
7-Day Rolling Avg Breaks = 
VAR CurrentDate = MAX('Calendar'[Date])
VAR RollingPeriod = 
    DATESBETWEEN(
        'Calendar'[Date],
        CurrentDate - 6,
        CurrentDate
    )
RETURN
    DIVIDE(
        CALCULATE(
            COUNTROWS(ControlBreaks),
            RollingPeriod
        ),
        7,
        0
    )
```

- **Format:** Decimal Number (`0.00`)
- **Institutional Context:**  
  Daily break counts fluctuate due to market holidays, volatile earnings days, and auction timing. The 7-day rolling average acts as a low-pass filter, smoothing out daily volatility to highlight chronic, persistent control deterioration.
- **DAX Mechanics & Interview Talking Point:**  
  > *"This measure uses variable assignment and filter context manipulation:*  
  > *1. `VAR CurrentDate = MAX('Calendar'[Date])` captures the visual's current date coordinate.*  
  > *2. `DATESBETWEEN(..., CurrentDate - 6, CurrentDate)` defines a sliding 7-day window.*  
  > *3. `CALCULATE(..., RollingPeriod)` modifies the filter context via context transition, summing all breaks within that 7-day window regardless of the single-date row context.*  
  > *4. Finally, dividing by 7 normalizes the count into a rolling daily average rate."*

---

### Measure 3: Total Positions Evaluated

```dax
Total Positions Evaluated = 
COUNTROWS(DailyMarketPrice)
```

- **Format:** Whole Number (`#,##0`)
- **Institutional Context:**  
  The total population of mark-to-market pairs evaluated by the control engine (1,008 positions over the 6 months).

---

### Measure 4: Total Control Breaks

```dax
Total Control Breaks = 
COUNTROWS(ControlBreaks)
```

- **Format:** Whole Number (`#,##0`)
- **Institutional Context:**  
  Total count of positions breaching the $\pm 3.0\%$ policy threshold (62 exceptions).

---

### Measure 5: Overall Break Rate % (Key Risk Indicator - KRI)

```dax
Break Rate % = 
DIVIDE(
    [Total Control Breaks],
    [Total Positions Evaluated],
    0
)
```

- **Format:** Percentage (`0.00%`)
- **Institutional Context:**  
  The primary Key Risk Indicator (KRI) reported to the Chief Risk Officer (CRO). Across our 6-month historical baseline, this sits at **6.15%**.

---

### Measure 6: Major Severity Breaches (>5%)

```dax
Major Breaches Count = 
CALCULATE(
    COUNTROWS(ControlBreaks),
    ControlBreaks[Severity] = "Major"
)
```

- **Format:** Whole Number (`#,##0`)
- **Institutional Context:**  
  High-priority breaches requiring formal Valuation Committee escalation and potential P&L reserves (22 exceptions).

---

### Measure 7: Minor Severity Breaches (3-5%)

```dax
Minor Breaches Count = 
CALCULATE(
    COUNTROWS(ControlBreaks),
    ControlBreaks[Severity] = "Minor"
)
```

- **Format:** Whole Number (`#,##0`)
- **Institutional Context:**  
  Routine desk-level investigation items (40 exceptions).
