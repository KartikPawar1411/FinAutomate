Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub LinkCell(ByVal dest As Range, ByVal src As Range)
    ' Keep source blanks blank; preserve an actual zero.
    Dim ref As String
    ref = "'" & src.Parent.Name & "'!" & src.Address(False, False)
    dest.Formula = "=IF(ISBLANK(" & ref & "),""""," & ref & ")"
    dest.Font.Color = RGB(0, 105, 64)
End Sub

Public Sub TrendChart(ByVal w As Worksheet, ByVal anchor As Range, ByVal tableRow As Long, ByVal n As Long, ByVal cols As Variant, ByVal label As String, ByVal percent As Boolean)
    Dim co As ChartObject, series As Series, c As Variant, i As Long
    If n < 1 Then Exit Sub
    Set co = w.ChartObjects.Add(anchor.Left + 2, anchor.Top + 2, anchor.Width - 4, anchor.Height - 4)
    With co.Chart
        .ChartType = xlLineMarkers
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        For Each c In cols
            i = i + 1: Set series = .SeriesCollection.NewSeries
            series.Name = CStr(w.Cells(tableRow, CLng(c)).Value2)
            series.XValues = w.Cells(tableRow + 1, 1).Resize(n)
            ' Native line charts can plot formula text as zero. Numeric-only helper
            ' ranges leave actual empty cells for unavailable observations.
            Dim point As Long, hx As Long, source As Range
            hx = 10 + CLng(c)
            For point = 1 To n
                Set source = w.Cells(tableRow + point, CLng(c))
                If VarType(source.Value2) = vbDouble Then w.Cells(tableRow + point, hx).Formula = "=" & source.Address
            Next
            w.Columns(hx).Hidden = True
            series.Values = w.Cells(tableRow + 1, hx).Resize(n)
            .PlotVisibleOnly = False
            series.Format.Line.ForeColor.RGB = IIf(i = 1, RGB(31, 75, 105), RGB(15, 142, 129))
            series.MarkerStyle = xlMarkerStyleCircle: series.MarkerSize = 5
        Next
        .HasTitle = True: .ChartTitle.Text = label: .ChartTitle.Font.Size = 12
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
        .Axes(xlValue).TickLabels.NumberFormat = IIf(percent, "0%", "#,##0")
        .DisplayBlanksAs = xlNotPlotted
    End With
End Sub

Public Sub RefreshDashboard()
    Dim w As Worksheet, fs As Worksheet, ratios As Worksheet, peers As Worksheet, ids As Object
    Dim r As Long, latest As Long, p As Long, id As String, k, n As Long, c As Long, cols, peerStart As Long, t As Double
    t = Timer: Set w = S("Dashboard"): Set fs = S("Financial_Statements"): Set ratios = S("Financial_Ratios")
    Set peers = S("Company_Comparison"): Set ids = Dict()
    id = CStr(w.Range("B4").Value2): ClearView w: ClearView peers
    Title w, "Company review", "Annual filings  /  original currency, amounts in millions"
    Title peers, "Peer comparison", "Latest accepted annual filing per company. Different dates, currencies and industries limit comparability."
    Headers peers, 4, Split("Company|Period end|Currency|Revenue (m)|Net margin|Debt / EBITDA|EV / EBITDA|P/E", "|")
    For r = 2 To LastR(fs): ids(CStr(fs.Cells(r, 1).Value2)) = r: Next
    If ids.Count = 0 Then w.Range("A4") = "No accepted filings. Inspect Clean_Data.": Exit Sub
    If Not ids.Exists(id) Then id = CStr(ids.Keys()(0))
    p = 4
    For Each k In ids.Keys
        r = ids(k): p = p + 1
        LinkCell peers.Cells(p, 1), fs.Cells(r, 1): LinkCell peers.Cells(p, 2), fs.Cells(r, 4)
        LinkCell peers.Cells(p, 3), fs.Cells(r, 5): LinkCell peers.Cells(p, 4), fs.Cells(r, 7)
        cols = Array(7, 13, 19, 20)
        For c = 0 To 3: LinkCell peers.Cells(p, 5 + c), ratios.Cells(r, cols(c)): Next
    Next
    peers.Range("B5:B" & p).NumberFormat = "yyyy-mm-dd": peers.Range("D5:D" & p).NumberFormat = AMOUNT
    peers.Range("E5:E" & p).NumberFormat = PCT: peers.Range("F5:H" & p).NumberFormat = MULT
    peers.Range("C5:C" & p).HorizontalAlignment = xlCenter
    peers.Range("D5:H" & p).HorizontalAlignment = xlRight
    peers.Calculate
    ThisWorkbook.Names.Add Name:="CompanyIDs", RefersTo:="=Company_Comparison!$A$5:$A$" & p
    latest = ids(id): w.Range("A4") = "Company selector": w.Range("B4") = id
    With w.Range("B4")
        .Interior.Color = RGB(220, 239, 239): .Font.Color = vbBlue
        .Validation.Delete: .Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=CompanyIDs"
    End With
    w.Range("C4:D4").Merge: w.Range("C4").Value2 = Literal(fs.Cells(latest, 2).Value2)
    w.Range("E4") = fs.Cells(latest, 3).Value2: w.Range("F4") = fs.Cells(latest, 5).Value2 & " millions"
    w.Range("G4:H4").Merge: w.Range("G4") = S("Settings").Range("B5").Value2 & " rejected in batch"
    w.Range("A3:H3").Merge
    w.Range("A3").Formula = "=IF(Settings!B4=""Current"","""",""Pending full workflow: views may reflect an earlier import."")"
    w.Range("A3").Font.Color = RGB(146, 76, 24)
    Headers w, 6, Split("Revenue (m)|EBITDA (m)|Net margin|Revenue YoY|Debt / EBITDA|Interest cover|EV / EBITDA|P/E", "|")
    LinkCell w.Cells(7, 1), fs.Cells(latest, 7): LinkCell w.Cells(7, 2), fs.Cells(latest, 13)
    cols = Array(7, 15, 13, 14, 19, 20)
    For c = 0 To 5: LinkCell w.Cells(7, 3 + c), ratios.Cells(latest, cols(c)): Next
    w.Range("A7:B7").NumberFormat = AMOUNT: w.Range("C7:D7").NumberFormat = PCT: w.Range("E7:H7").NumberFormat = MULT
    w.Range("A7:H7").Font.Size = 15: w.Rows(7).RowHeight = 31
    Headers w, 23, Split("Period|Revenue|EBITDA|Gross margin|Net margin|Debt / EBITDA|Current ratio|EV / EBITDA", "|")
    For r = 2 To LastR(fs)
        If fs.Cells(r, 1).Value2 = id And fs.Cells(r, 5).Value2 = fs.Cells(latest, 5).Value2 Then
            n = n + 1: p = 23 + n
            LinkCell w.Cells(p, 1), fs.Cells(r, 3): LinkCell w.Cells(p, 2), fs.Cells(r, 7)
            LinkCell w.Cells(p, 3), fs.Cells(r, 13)
            cols = Array(4, 7, 13, 10, 19)
            For c = 0 To 4: LinkCell w.Cells(p, c + 4), ratios.Cells(r, cols(c)): Next
        End If
    Next
    w.Range("B24:C" & p).NumberFormat = AMOUNT: w.Range("D24:E" & p).NumberFormat = PCT
    w.Range("F24:H" & p).NumberFormat = MULT: w.Calculate
    TrendChart w, w.Range("A9:D21"), 23, n, Array(2, 3), "Revenue and EBITDA  |  " & fs.Cells(latest, 5) & " m", False
    TrendChart w, w.Range("E9:H21"), 23, n, Array(4, 5), "Profit margins", True
    peerStart = p + 3: w.Cells(peerStart - 1, 1) = "Latest filings by company"
    Dim source As Range, target As Range
    Set source = peers.Range("A4:H" & LastR(peers)): Set target = w.Cells(peerStart, 1).Resize(source.Rows.Count, 8)
    source.Copy target: target.Value2 = source.Value2
    w.Cells(peerStart + source.Rows.Count + 1, 1) = "Each row uses its own period and currency. No FX conversion or cross-sector ranking."
    Audit "Dashboard", "OK", n, id & " | " & n & " annual filings in latest currency.", Took(t)
    w.Hyperlinks.Add w.Range("I4"), "", "'Valuation_Setup'!A1", , "Valuation workspace"
End Sub
