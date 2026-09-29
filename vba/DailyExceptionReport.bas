Attribute VB_Name = "DailyExceptionReport"
Option Explicit

' ==============================================================================
' Valuation Control & Automated Reporting Platform - IPV Exception Reporter
' Target Organization: Investment Banking / Group Finance / Product Control
' Description: Connects via ADODB to MS Access (valuation_control.accdb),
'              extracts ControlBreaks for a specified date, applies institutional
'              formatting and conditional severity highlighting, generates an
'              Executive KPI Summary, and auto-saves the formatted report.
' ==============================================================================

Public Sub RunDailyExceptionReport()
    Dim targetDate As String
    Dim dbPath As String
    Dim defaultDate As String
    
    ' Default to high-impact demonstration date or current date
    defaultDate = "2026-07-24"
    
    targetDate = InputBox( _
        "Enter Valuation Date to generate Exception Report (YYYY-MM-DD):", _
        "Valuation Control - Daily Exception Reporter", _
        defaultDate _
    )
    
    If Trim(targetDate) = "" Then
        MsgBox "Operation cancelled by user.", vbInformation, "Valuation Control"
        Exit Sub
    End If
    
    ' Resolve database path relative to this workbook or standard project path
    dbPath = ThisWorkbook.Path & "\..\data\valuation_control.accdb"
    If Dir(dbPath) = "" Then
        dbPath = ThisWorkbook.Path & "\data\valuation_control.accdb"
    End If
    
    If Dir(dbPath) = "" Then
        MsgBox "Database not found at expected path: " & vbCrLf & dbPath, vbCritical, "Database Connection Error"
        Exit Sub
    End If
    
    ' Generate report for the selected date
    Call GenerateReportForDate(dbPath, targetDate)
End Sub


Public Sub GenerateReportForDate(ByVal dbPath As String, ByVal targetDate As String)
    Dim conn As Object
    Dim rs As Object
    Dim sqlQuery As String
    Dim ws As Worksheet
    Dim reportWb As Workbook
    Dim lastRow As Long
    Dim exportPath As String
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. Establish ADODB Connection to MS Access Database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    ' 2. Formulate Relational Join Query for target date
    sqlQuery = "SELECT b.BreakID, i.Ticker, i.AssetClass, i.Sector, " & _
               "Format(b.[Date], 'YYYY-MM-DD') AS ValuationDate, " & _
               "b.MarketPrice, b.TraderMark, b.VariancePct, " & _
               "b.Severity, t.TraderID, b.Status " & _
               "FROM ((ControlBreaks b " & _
               "INNER JOIN Instruments i ON b.InstrumentID = i.InstrumentID) " & _
               "INNER JOIN DailyTraderMark t ON (b.InstrumentID = t.InstrumentID AND b.[Date] = t.[Date])) " & _
               "WHERE b.[Date] = #" & targetDate & "# " & _
               "ORDER BY ABS(b.VariancePct) DESC"
               
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn, 1, 1 ' adOpenKeyset, adLockReadOnly
    
    If rs.EOF And rs.BOF Then
        MsgBox "No valuation control breaks found for date: " & targetDate & vbCrLf & _
               "All marks were within +/-3.0% tolerance.", vbInformation, "Valuation Control - Clean Run"
        rs.Close
        conn.Close
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' 3. Create a fresh Workbook for the Daily Exception Report
    Set reportWb = Workbooks.Add(xlWBATWorksheet)
    Set ws = reportWb.Sheets(1)
    ws.Name = "Exception_Report"
    ws.Tab.Color = RGB(30, 58, 138) ' Navy Blue
    
    ' 4. Build Title and Header Block
    ws.Range("A1:K1").Merge
    ws.Range("A1").Value = "VALUATION CONTROL & INDEPENDENT PRICE VERIFICATION (IPV) EXCEPTION REPORT"
    With ws.Range("A1")
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 58, 138) ' Deep Navy Blue
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ws.Range("A2:K2").Merge
    ws.Range("A2").Value = "Valuation Date: " & targetDate & " | Report Generated: " & Format(Now, "YYYY-MM-DD HH:NN:SS") & " | Governance: UBS Group Finance / Product Control"
    With ws.Range("A2")
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = RGB(71, 85, 105)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlCenter
        .RowHeight = 20
    End With
    
    ' 5. Build Executive KPI Summary Cards (Rows 4 to 7)
    Call BuildExecutiveSummary(ws, rs, targetDate)
    
    ' 6. Write Data Table Headers (Row 9)
    Dim headers As Variant
    headers = Array("Break ID", "Ticker", "Asset Class", "Sector", "Valuation Date", _
                    "Market Price ($)", "Trader Mark ($)", "Variance (%)", "Severity", "Trader ID", "Status")
    
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(9, c + 1).Value = headers(c)
    Next c
    
    With ws.Range("A9:K9")
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(51, 65, 85) ' Slate Charcoal
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
    End With
    
    ' 7. Paste Recordset Data starting at Row 10
    ws.Range("A10").CopyFromRecordset rs
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 8. Apply Number Formatting and Alignment
    ws.Range("A10:A" & lastRow).HorizontalAlignment = xlCenter ' BreakID
    ws.Range("B10:B" & lastRow).HorizontalAlignment = xlCenter ' Ticker
    ws.Range("C10:C" & lastRow).HorizontalAlignment = xlCenter ' Asset Class
    ws.Range("D10:D" & lastRow).HorizontalAlignment = xlLeft   ' Sector
    ws.Range("E10:E" & lastRow).HorizontalAlignment = xlCenter ' Valuation Date
    
    ws.Range("F10:F" & lastRow).NumberFormat = "$#,##0.00"     ' Market Price
    ws.Range("G10:G" & lastRow).NumberFormat = "$#,##0.00"     ' Trader Mark
    ws.Range("H10:H" & lastRow).NumberFormat = "+0.00%;-0.00%;0.00%" ' Variance Pct
    
    ws.Range("I10:I" & lastRow).HorizontalAlignment = xlCenter ' Severity
    ws.Range("J10:J" & lastRow).HorizontalAlignment = xlCenter ' Trader ID
    ws.Range("K10:K" & lastRow).HorizontalAlignment = xlCenter ' Status
    
    ' Subtle cell borders
    With ws.Range("A9:K" & lastRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(203, 213, 225)
    End With
    
    ' 9. Dynamic Conditional Formatting (Red for Major, Yellow for Minor)
    Call ApplySeverityConditionalFormatting(ws, lastRow)
    
    ' Auto-fit columns with safety margins
    ws.Columns("A:K").AutoFit
    For c = 1 To 11
        If ws.Columns(c).ColumnWidth < 14 Then ws.Columns(c).ColumnWidth = 14
    Next c
    
    ' 10. Auto-Save Formatted Workbook
    exportPath = ThisWorkbook.Path & "\Exception_Report_" & targetDate & ".xlsx"
    reportWb.SaveAs Filename:=exportPath, FileFormat:=51 ' 51 = xlOpenXMLWorkbook (.xlsx)
    
    ' Clean up ADODB objects
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Daily Exception Report successfully generated and archived at:" & vbCrLf & _
           exportPath & vbCrLf & vbCrLf & _
           "Flagged Breaks: " & (lastRow - 9), vbInformation, "Valuation Control - Report Generated"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error generating exception report: " & Err.Description, vbCritical, "Macro Execution Error"
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Sub


Private Sub BuildExecutiveSummary(ByVal ws As Worksheet, ByVal rs As Object, ByVal targetDate As String)
    ' Card 1: Total Exceptions
    ws.Range("A4:B4").Merge
    ws.Range("A4").Value = "TOTAL EXCEPTIONS"
    ws.Range("A5:B5").Merge
    ws.Range("A5").Formula = "=COUNTA(A10:A100)"
    
    ' Card 2: Major Breaches (>5%)
    ws.Range("D4:E4").Merge
    ws.Range("D4").Value = "MAJOR SEVERITY (>5%)"
    ws.Range("D5:E5").Merge
    ws.Range("D5").Formula = "=COUNTIF(I10:I100, ""Major"")"
    
    ' Card 3: Minor Breaches (3-5%)
    ws.Range("G4:H4").Merge
    ws.Range("G4").Value = "MINOR SEVERITY (3-5%)"
    ws.Range("G5:H5").Merge
    ws.Range("G5").Formula = "=COUNTIF(I10:I100, ""Minor"")"
    
    ' Card 4: Breaches by Asset Class
    ws.Range("J4:K4").Merge
    ws.Range("J4").Value = "ASSET CLASS SCOPE"
    ws.Range("J5:K5").Merge
    ws.Range("J5").Value = "Equity Cash"
    
    ' Style Summary Headers
    Dim cardCols As Variant
    cardCols = Array("A4:B4", "D4:E4", "G4:H4", "J4:K4")
    Dim i As Long
    For i = 0 To UBound(cardCols)
        With ws.Range(cardCols(i))
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = RGB(100, 116, 139)
            .Interior.Color = RGB(248, 250, 252)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Style Summary Values
    Dim valCols As Variant
    valCols = Array("A5:B5", "D5:E5", "G5:H5", "J5:K5")
    For i = 0 To UBound(valCols)
        With ws.Range(valCols(i))
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 28
        End With
    Next i
    
    ' Specific Alert Colors for Metric Cards
    ws.Range("A5:B5").Font.Color = RGB(30, 41, 59)
    ws.Range("D5:E5").Font.Color = RGB(185, 28, 28) ' Crimson Red
    ws.Range("G5:H5").Font.Color = RGB(180, 83, 9)   ' Amber Yellow
    ws.Range("J5:K5").Font.Color = RGB(30, 58, 138)  ' Navy Blue
    
    ' Add Card Borders
    For i = 0 To UBound(cardCols)
        Dim rangeStr As String
        rangeStr = Left(cardCols(i), 1) & "4:" & Mid(cardCols(i), 4, 1) & "5"
        With ws.Range(rangeStr).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(226, 232, 240)
        End With
    Next i
End Sub


Private Sub ApplySeverityConditionalFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim severityRange As Range
    Dim fullDataRange As Range
    Dim cfMajor As FormatCondition
    Dim cfMinor As FormatCondition
    
    Set severityRange = ws.Range("I10:I" & lastRow)
    severityRange.FormatConditions.Delete
    
    ' Rule 1: Major Severity -> Soft Red Fill with Bold Dark Red Text
    Set cfMajor = severityRange.FormatConditions.Add(xlCellValue, xlEqual, "=""Major""")
    With cfMajor
        .Interior.Color = RGB(254, 226, 226) ' Light pastel red
        .Font.Color = RGB(153, 27, 27)       ' Dark crimson red
        .Font.Bold = True
    End With
    
    ' Rule 2: Minor Severity -> Soft Yellow Fill with Bold Amber Text
    Set cfMinor = severityRange.FormatConditions.Add(xlCellValue, xlEqual, "=""Minor""")
    With cfMinor
        .Interior.Color = RGB(254, 243, 199) ' Light pastel amber
        .Font.Color = RGB(146, 64, 14)       ' Dark amber brown
        .Font.Bold = True
    End With
End Sub
