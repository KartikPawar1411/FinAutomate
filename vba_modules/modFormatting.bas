Attribute VB_Name = "modFormatting"
Option Explicit

Public Sub Band(ByVal area As Range)
    With area
        .Interior.Color = RGB(26, 58, 74): .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 24: .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub Title(ByVal w As Worksheet, ByVal label As String, ByVal subtitle As String)
    w.Range("A1:H1").Merge: w.Range("A1").Value2 = Literal(label)
    With w.Range("A1:H1"): .Font.Size = 22: .Font.Bold = True: .Font.Color = RGB(25, 54, 70): .RowHeight = 39: End With
    w.Range("A2:H2").Merge: w.Range("A2").Value2 = subtitle
    With w.Range("A2:H2"): .Font.Color = RGB(22, 98, 96): .RowHeight = 23: End With
End Sub

Public Sub ClearView(ByVal w As Worksheet)
    Dim i As Long
    For i = w.ChartObjects.Count To 1 Step -1: w.ChartObjects(i).Delete: Next
    w.UsedRange.UnMerge: w.UsedRange.Clear
    w.Columns("A:H").ColumnWidth = 16
End Sub

Private Function ExternalRef(ByVal f As String) As Boolean
    Dim i As Long, inText As Boolean, ch As String
    For i = 1 To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then inText = Not inText
        If ch = "!" And Not inText Then ExternalRef = True: Exit Function
    Next
End Function

Public Sub ApplyIBFormatting()
    Dim w As Worksheet, n, c As Range, r As Long, bad As Long, t As Double, a As Range, nm As Variant, prior As Worksheet
    t = Timer
    Set prior = ThisWorkbook.ActiveSheet
    For Each n In Array("Raw_Data", "Clean_Data", "Financial_Statements", "Financial_Ratios", "Audit_Log")
        Set w = S(CStr(n)): r = LastR(w)
        If w.AutoFilterMode Then w.AutoFilterMode = False
        If r >= 1 Then w.Range("A1").CurrentRegion.AutoFilter
        If r > 1 And n <> "Audit_Log" Then
            If n = "Financial_Ratios" Then
                w.Range("D2:I" & r).NumberFormat = PCT: w.Range("J2:N" & r).NumberFormat = MULT
                w.Range("O2:Q" & r).NumberFormat = PCT: w.Range("R2:T" & r).NumberFormat = MULT
            ElseIf n <> "Raw_Data" Then
                w.Range("G2:AI" & r).NumberFormat = AMOUNT
            End If
            Set a = w.UsedRange.Offset(1, 0).Resize(r - 1)
            a.Font.Color = RGB(0, 58, 171)
            For Each c In a.Cells
                If c.HasFormula Then
                    If ExternalRef(c.Formula) Then c.Font.Color = RGB(0, 105, 64) Else c.Font.Color = vbBlack
                End If
                If IsError(c.Value2) Then bad = bad + 1: c.Interior.Color = RGB(255, 209, 200)
                If c.Column = 38 And c.Value2 = "Rejected" Then c.Interior.Color = RGB(255, 209, 200)
            Next
        End If
        w.UsedRange.Columns.AutoFit
        If n = "Audit_Log" Then w.Columns(5).ColumnWidth = 86
        If n = "Raw_Data" Or n = "Clean_Data" Or n = "Financial_Statements" Then w.Columns(39).ColumnWidth = 84
        ' Freeze panes are window-specific in Excel; activate only for this UI operation.
        w.Activate
        With ThisWorkbook.Windows(1): .DisplayGridlines = False: .FreezePanes = False: .SplitColumn = 1: .SplitRow = 1: .FreezePanes = True: End With
    Next
    prior.Activate
    If bad > 0 Then Fail bad & " spreadsheet formula errors. Inspect highlighted cells."
    Audit "Format", "OK", 0, "Financial formats applied; formula error scan complete.", Took(t)
End Sub
