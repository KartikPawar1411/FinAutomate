Attribute VB_Name = "modReporting"
Option Explicit

Public Sub BuildReport()
    Dim w As Worksheet, d As Worksheet, src As Worksheet, area As Range, dest As Range, r As Long, n As Long, finish As Long
    Dim mode As String, t As Double, memo As String, chartCol As Variant
    t = Timer: Set w = S("Report"): Set d = S("Dashboard"): mode = CStr(S("Home").Range("B4").Value2)
    ClearView w
    If LastR(S("Financial_Statements")) < 2 Then Fail "No accepted financial statements to report."
    If mode = "Comparison" Then
        Title w, "Annual peer comparison", "Financial review  /  one row per issuer, latest accepted fiscal year"
        Set src = S("Company_Comparison"): Set area = src.Range("A4:H" & LastR(src))
        Set dest = w.Range("A5").Resize(area.Rows.Count, 8): area.Copy dest: dest.Value2 = area.Value2
        finish = 6 + area.Rows.Count
        memo = "Amounts are in millions of the row currency. Filing dates and business models differ; no ranking or FX translation."
    Else
        Title w, CStr(d.Range("C4").Value2) & "  /  annual review", d.Range("B4").Value2 & "   |   " & d.Range("E4").Value2 & "   |   " & d.Range("F4").Value2
        Set area = d.Range("A6:H7"): Set dest = w.Range("A4:H5"): area.Copy dest: dest.Value2 = area.Value2
        n = 0
        Do While Len(CStr(d.Cells(24 + n, 1).Value2)) > 0: n = n + 1: Loop
        Set area = d.Range("A23").Resize(n + 1, 8): Set dest = w.Range("A21").Resize(n + 1, 8)
        area.Copy dest: dest.Value2 = area.Value2
        TrendChart w, w.Range("A7:D19"), 21, n, Array(2, 3), "Revenue and EBITDA  |  " & d.Range("F4").Value2, False
        TrendChart w, w.Range("E7:H19"), 21, n, Array(4, 5), "Profit margins", True
        finish = 23 + n
        Dim fs As Worksheet, latest As Long, financials, j As Long
        Set fs = S("Financial_Statements")
        For r = 2 To LastR(fs)
            If fs.Cells(r, 1).Value2 = d.Range("B4").Value2 Then latest = r
        Next
        Headers w, finish, Split("Assets|Equity|Cash|Total debt|EBIT|Net income|Gross profit|CapEx", "|")
        financials = Array(20, 29, 22, 28, 15, 19, 9, 35)
        For j = 0 To 7
            If IsEmpty(fs.Cells(latest, financials(j)).Value2) Then
                w.Cells(finish + 1, j + 1) = "N/A"
            Else
                w.Cells(finish + 1, j + 1).Value2 = fs.Cells(latest, financials(j)).Value2
            End If
        Next
        w.Cells(finish + 1, 1).Resize(1, 8).NumberFormat = AMOUNT
        finish = finish + 3
        memo = "Trends include accepted annual filings in the latest reporting currency. No FX translation. Missing years may interrupt comparability."
    End If
    w.Cells(finish, 1).Resize(1, 8).Merge: w.Cells(finish, 1) = "Basis and limitations": Band w.Cells(finish, 1).Resize(1, 8)
    Dim lines, i As Long
    lines = Array(memo, S("Settings").Range("B5").Value2 & " rejected in this batch. N/A = missing input or comparable prior year; N/M = invalid economic base.", "Management-style analytical output from supplied data; no independent assurance or investment recommendation.")
    For i = 0 To 2
        Set dest = w.Cells(finish + 1 + i, 1).Resize(1, 8): dest.Merge: dest.Cells(1) = lines(i)
        dest.WrapText = True: dest.RowHeight = 29
    Next
    With w.PageSetup
        .PrintArea = w.Range("A1:H" & finish + 3).Address: .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = IIf(finish < 44, 1, False)
        .CenterHorizontally = True: .LeftFooter = "FINAUTOMATE  /  FINANCIAL REVIEW": .RightFooter = "Page &P of &N"
    End With
    Audit "Report", "OK", 1, mode & " report generated.", Took(t)
End Sub
