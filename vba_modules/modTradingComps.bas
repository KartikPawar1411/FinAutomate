Attribute VB_Name = "modTradingComps"
Option Explicit

Public Sub BuildRanges(ByVal targetRow As Long, ByVal last As Long)
    Dim w As Worksheet, summary As Worksheet, cfg As Worksheet, i As Long, j As Long, r As Long, refs As String, metric As String
    Dim methods, bases, stats, sref As String, ev As String, eq As String, bridge As String, shares As String
    Set w = VSheet("Trading_Comps"): ClearView w: Set cfg = S("Valuation_Setup")
    methods = Array("EV / revenue", "EV / EBITDA", "EV / EBIT", "P/E")
    bases = Array("G", "H", "I", "J")
    Headers w, 1, Split("Method|Peers|Min|P25|Median|P75|Max|Mean", "|")
    Headers w, 8, Split("Method|Quartile|Multiple|Target metric|Implied EV|Implied equity|Implied price", "|")
    For i = 0 To 3
        r = 2 + i: w.Cells(r, 1) = methods(i)
        refs = "Peer_Group!" & Chr$(71 + i) & "2:" & Chr$(71 + i) & last
        w.Cells(r, 2).Formula = "=COUNT(" & refs & ")"
        stats = Array("MIN(" & refs & ")", "PERCENTILE(" & refs & ",0.25)", "MEDIAN(" & refs & ")", "PERCENTILE(" & refs & ",0.75)", "MAX(" & refs & ")", "AVERAGE(" & refs & ")")
        For j = 0 To 5: w.Cells(r, j + 3).Formula = "=IF(B" & r & ">0," & stats(j) & ",""N/A"")": Next
        For j = 0 To 2
            r = 9 + i * 3 + j: w.Cells(r, 1) = methods(i): w.Cells(r, 2) = Array("P25", "Median", "P75")(j)
            w.Cells(r, 3).Formula = "=" & Chr$(68 + j) & i + 2
            metric = "Trading_Data!" & bases(i) & targetRow
            w.Cells(r, 4).Formula = "=IF(" & metric & ">0," & metric & ",""N/M"")"
            bridge = "Trading_Data!AB" & targetRow: shares = "Trading_Data!K" & targetRow
            If i < 3 Then
                ev = "C" & r & "*D" & r: eq = "IF(ISNUMBER(" & bridge & "),E" & r & "-" & bridge & ",""N/A"")"
            Else
                eq = "C" & r & "*D" & r: ev = "IF(ISNUMBER(" & bridge & "),F" & r & "+" & bridge & ",""N/A"")"
            End If
            w.Cells(r, 5).Formula = "=IFERROR(" & ev & ",""N/M"")": w.Cells(r, 6).Formula = "=IFERROR(" & eq & ",""N/M"")"
            w.Cells(r, 7).Formula = "=IFERROR(IF(AND(F" & r & ">0," & shares & ">0),F" & r & "/" & shares & ",""N/M""),""N/M"")"
        Next
    Next
    w.UsedRange.Columns.AutoFit: w.Range("C2:H5,C9:G20").NumberFormat = "#,##0.00"
    w.Calculate
    Set summary = VSheet("Valuation_Summary"): ClearView summary
    Title summary, CStr(cfg.Range("B5")) & " | valuation", "Valuation " & Format$(cfg.Range("B2"), "yyyy-mm-dd") & "   /   LTM to " & Format$(cfg.Range("B3"), "yyyy-mm-dd") & "   /   " & cfg.Range("B4") & " millions except per share"
    Headers summary, 4, Split("Method|Low / P25|Median|High / P75", "|")
    For i = 0 To 3
        r = 5 + i: summary.Cells(r, 1) = methods(i)
        For j = 0 To 2: summary.Cells(r, j + 2).Formula = "=Trading_Comps!G" & 9 + i * 3 + j: Next
    Next
    summary.Range("A10") = "DCF / share": summary.Range("B10").Formula = "=DCF!B23"
    summary.Range("C10") = "WACC": summary.Range("D10").Formula = "=DCF!B18": summary.Range("D10").NumberFormat = "0.00%"
    summary.Range("A11") = "Terminal growth": summary.Range("B11").Formula = "=Valuation_Setup!B17": summary.Range("B11").NumberFormat = "0.0%"
    summary.Range("C11") = "Terminal basis": summary.Range("D11").Formula = "=Valuation_Setup!B19"
    summary.Range("B5:D8,B10").NumberFormat = "0.00"
    summary.Calculate
    Dim chart As ChartObject, ser As Series, rg As Range, points As Long
    For i = 5 To 8
        If IsNumeric(summary.Cells(i, 2)) And IsNumeric(summary.Cells(i, 4)) Then
            points = points + 1: summary.Cells(points, 10) = summary.Cells(i, 1)
            summary.Cells(points, 11).Formula = "=B" & i: summary.Cells(points, 12).Formula = "=D" & i & "-B" & i
        End If
    Next
    If points > 0 Then
        Set rg = summary.Range("A13:H26")
        Set chart = summary.ChartObjects.Add(rg.Left, rg.Top, rg.Width, rg.Height)
        With chart.Chart
            .ChartType = xlBarStacked
            Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
            For i = 1 To 2
                Set ser = .SeriesCollection.NewSeries: ser.XValues = summary.Cells(1, 10).Resize(points): ser.Values = summary.Cells(1, 10 + i).Resize(points)
                If i = 1 Then ser.Format.Fill.Visible = msoFalse Else ser.Format.Fill.ForeColor.RGB = RGB(24, 118, 122)
                ser.Format.Line.Visible = msoFalse
            Next
            .HasTitle = True: .ChartTitle.Text = "Implied price per diluted share | peer interquartile range"
            .HasLegend = False: .PlotVisibleOnly = False
        End With
    End If
    summary.Columns("J:L").Hidden = True
    Dim data As Worksheet: Set data = S("Trading_Data")
    summary.Range("A28") = "LTM revenue": summary.Range("B28").Formula = "=Trading_Data!G" & targetRow
    summary.Range("C28") = "EBITDA margin": summary.Range("D28").Formula = "=IF(ISNUMBER(Trading_Data!Y" & targetRow & "),Trading_Data!Y" & targetRow & ",""N/A"")": summary.Range("D28").NumberFormat = "0.0%"
    summary.Range("A29:H29").Merge: summary.Range("A29").Formula = "=""Revenue growth: ""&IF(ISNUMBER(Trading_Data!Z" & targetRow & "),TEXT(Trading_Data!Z" & targetRow & ",""0.0%""),""unavailable"")&""   |   Net debt / EBITDA: ""&IF(ISNUMBER(Trading_Data!AA" & targetRow & "),TEXT(Trading_Data!AA" & targetRow & ",""0.00"")&""x"",""unavailable"")"
    summary.Range("A30:H30").Merge: summary.Range("A30").Formula = "=IFERROR(IF(ISNUMBER(Trading_Data!V" & targetRow & "),""EV / EBITDA relative to peer median: ""&TEXT(Trading_Data!V" & targetRow & "/Trading_Comps!E3-1,""+0.0%;-0.0%;0.0%""),""Peer comparison unavailable""),""Peer comparison unavailable"")"
    summary.Range("A31:H31").Merge: summary.Range("A31") = "Quote " & data.Cells(targetRow, 5) & " | " & data.Cells(targetRow, 30) & " | FX " & data.Cells(targetRow, 32)
    summary.Range("A32:H32").Merge: summary.Range("A32") = "Reconciliation: " & S("Quality_Control").Range("B3") & " imported rows. " & S("Quality_Control").Range("B5") & " quarantined."
    summary.Range("A33:H33").Merge: summary.Range("A33") = "Peer_Group documents the selection. LTM_Reconciliation lists source IDs. Equity bridge includes minority and preferred claims."
    summary.Range("A34:H34").Merge: summary.Range("A34") = "Illustrative model. Assumptions on Valuation_Setup. Current FX translates LTM amounts at a constant rate."
    summary.Range("B28").NumberFormat = "#,##0.0"
    Dim obs As Worksheet, outRow As Long
    Set obs = S("Observations"): outRow = 35
    For i = 2 To LastR(obs)
        If obs.Cells(i, 1) = cfg.Range("B5") Then
            summary.Cells(outRow, 1).Resize(1, 8).Merge
            summary.Cells(outRow, 1) = obs.Cells(i, 6).Value2
            outRow = outRow + 1
        End If
    Next
    summary.Cells(outRow, 1).Resize(1, 8).Merge
    summary.Cells(outRow, 1).Formula = "=""Forecast revenue growth ""&TEXT(Valuation_Setup!B6,""0.0%"")&"" | EBITDA margin ""&TEXT(Valuation_Setup!B7,""0.0%"")&"" | CapEx / revenue ""&TEXT(Valuation_Setup!B9,""0.0%"")&"" | tax ""&TEXT(Valuation_Setup!B11,""0.0%"")"
    summary.Hyperlinks.Add summary.Cells(outRow + 2, 1), "", "'Record_Lineage'!A1", , "Trace source records"
    With summary.PageSetup
        .PrintArea = "$A$1:$H$" & outRow: .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
End Sub
