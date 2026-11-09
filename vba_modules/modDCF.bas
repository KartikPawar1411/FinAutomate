Attribute VB_Name = "modDCF"
Option Explicit

Public Sub BuildDCF(ByVal targetRow As Long)
    Dim w As Worksheet, i As Long, col As Long, letter As String, prior As String, labels, a As String, b As String
    Set w = VSheet("DCF"): ClearView w
    Title w, "Discounted cash flow", "Five forecast years  /  end-year discounting  /  amounts in reporting currency millions"
    Headers w, 4, Split("Metric|Historical LTM|Forecast 1|Forecast 2|Forecast 3|Forecast 4|Forecast 5", "|")
    labels = Split("Revenue|EBITDA|D&A|EBIT|Cash tax on positive EBIT|Capital expenditures|Change in operating NWC|Unlevered free cash flow|Discount factor|PV of UFCF", "|")
    For i = 0 To UBound(labels): w.Cells(5 + i, 1) = labels(i): Next
    w.Range("B5").Formula = "=Trading_Data!G" & targetRow
    w.Range("B6").Formula = "=IF(ISNUMBER(Trading_Data!H" & targetRow & "),Trading_Data!H" & targetRow & ",""N/A"")"
    w.Range("B8").Formula = "=IF(ISNUMBER(Trading_Data!I" & targetRow & "),Trading_Data!I" & targetRow & ",""N/A"")"
    For col = 3 To 7
        letter = Chr$(64 + col): prior = Chr$(63 + col)
        w.Cells(5, col).Formula = "=" & prior & "5*(1+Valuation_Setup!B6)"
        w.Cells(6, col).Formula = "=" & letter & "5*Valuation_Setup!B7"
        w.Cells(7, col).Formula = "=" & letter & "5*Valuation_Setup!B8"
        w.Cells(8, col).Formula = "=" & letter & "6-" & letter & "7"
        w.Cells(9, col).Formula = "=MAX(" & letter & "8,0)*Valuation_Setup!B11"
        w.Cells(10, col).Formula = "=" & letter & "5*Valuation_Setup!B9"
        w.Cells(11, col).Formula = "=(" & letter & "5-" & prior & "5)*Valuation_Setup!B10"
        w.Cells(12, col).Formula = "=" & letter & "8-" & letter & "9+" & letter & "7-" & letter & "10-" & letter & "11"
        w.Cells(13, col).Formula = "=1/(1+$B$18)^" & col - 2
        w.Cells(14, col).Formula = "=" & letter & "12*" & letter & "13"
    Next
    labels = Split("CAPM cost of equity|After-tax cost of debt|WACC|Terminal value - growth|Terminal value - exit|Enterprise value|Equity value|Per diluted share", "|")
    For i = 0 To UBound(labels): w.Cells(16 + i, 1) = labels(i): Next
    w.Range("B16").Formula = "=Valuation_Setup!B12+Valuation_Setup!B14*Valuation_Setup!B13"
    w.Range("B17").Formula = "=Valuation_Setup!B15*(1-Valuation_Setup!B11)"
    w.Range("B18").Formula = "=B16*(1-Valuation_Setup!B16)+B17*Valuation_Setup!B16"
    w.Range("A24") = "Assumptions check"
    w.Range("B24").Formula = "=IF(AND(COUNT(Valuation_Setup!B6:B18)=13,Valuation_Setup!B6>-1,Valuation_Setup!B8>=0,Valuation_Setup!B9>=0,Valuation_Setup!B11>=0,Valuation_Setup!B11<=1,Valuation_Setup!B16>=0,Valuation_Setup!B16<=1,B18>0,Valuation_Setup!B17>-1),""PASS"",""CHECK INPUTS"")"
    w.Range("B19").Formula = "=IF(AND(B18>Valuation_Setup!B17,G12>0),G12*(1+Valuation_Setup!B17)/(B18-Valuation_Setup!B17),""N/M"")"
    w.Range("B20").Formula = "=IF(AND(G6>0,Valuation_Setup!B18>0),G6*Valuation_Setup!B18,""N/M"")"
    w.Range("B21").Formula = "=IFERROR(IF(B24=""PASS"",SUM(C14:G14)+G13*IF(Valuation_Setup!B19=""Growth"",B19,IF(Valuation_Setup!B19=""Exit"",B20,""N/M"")),""N/A""),""N/M"")"
    a = "Trading_Data!AB" & targetRow: b = "Trading_Data!K" & targetRow
    w.Range("B22").Formula = "=IFERROR(IF(ISNUMBER(" & a & "),B21-" & a & ",""N/A""),""N/M"")"
    w.Range("B23").Formula = "=IFERROR(IF(AND(" & b & ">0,B22>0),B22/" & b & ",""N/M""),""N/M"")"
    w.Range("A26") = "Perpetual growth sensitivity | price per share"
    For i = 1 To 5
        w.Cells(27, i + 1).Formula = "=Valuation_Setup!B17+(" & i - 3 & ")*0.005"
        w.Cells(27 + i, 1).Formula = "=$B$18+(" & i - 3 & ")*0.01"
    Next
    For i = 28 To 32
        For col = 2 To 6
            letter = Chr$(64 + col)
            w.Cells(i, col).Formula = "=IFERROR(IF(AND($A" & i & ">" & letter & "$27,$G$12>0,ISNUMBER(" & a & ")," & b & ">0),(NPV($A" & i & ",$C$12:$G$12)+$G$12*(1+" & letter & "$27)/($A" & i & "-" & letter & "$27)/(1+$A" & i & ")^5-" & a & ")/" & b & ",""N/M""),""N/M"")"
        Next
    Next
    w.Range("B5:G23,B28:F32").NumberFormat = "#,##0.00;(#,##0.00);0.00"
    w.Range("B16:B18,B27:F27,A28:A32").NumberFormat = "0.00%"
    w.Columns("A").ColumnWidth = 35: w.Columns("B:G").ColumnWidth = 16
    w.Calculate
End Sub
