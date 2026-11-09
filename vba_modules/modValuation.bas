Attribute VB_Name = "modValuation"
Option Explicit

Public Function VSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set VSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If VSheet Is Nothing Then
        Set VSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        VSheet.Name = name
    End If
End Function

Public Function InstallValuation() As String
    On Error GoTo bad
    Dim setup As Worksheet, r As Long, labels, vals, btn As Shape, n, addr As Range
    Set setup = VSheet("Valuation_Setup")
    If IsEmpty(setup.Range("A1")) Then
        setup.Range("A1") = "Valuation assumptions"
        labels = Split("Valuation date|Common period end|Reporting currency|Target issuer|Forecast revenue growth|EBITDA margin|D&A / revenue|CapEx / revenue|Incremental NWC / incremental revenue|Tax rate|Risk-free rate|Equity risk premium|Levered beta|Pre-tax debt cost|Target debt weight|Perpetual growth|Exit EBITDA multiple|Terminal method", "|")
        vals = Array(DateSerial(2024, 3, 31), DateSerial(2023, 12, 31), "USD", "ACME", 0.06, 0.22, 0.04, 0.065, 0.12, 0.25, 0.04, 0.055, 1.1, 0.055, 0.25, 0.025, 12, "Growth")
        For r = 0 To UBound(labels)
            setup.Cells(2 + r, 1) = labels(r): setup.Cells(2 + r, 2) = vals(r)
        Next
        setup.Range("B2:B3").NumberFormat = "yyyy-mm-dd": setup.Range("B6:B13,B15:B17").NumberFormat = "0.0%"
        setup.Range("B2:B19").Font.Color = vbBlue
        setup.Range("B19").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "Growth,Exit"
        setup.Range("A20") = "Peer group": setup.Range("B20") = "Industrial components": setup.Range("B20").Font.Color = vbBlue
        setup.Range("D2") = "Edit blue assumptions, then refresh. Values shown are synthetic."
        setup.Range("D3") = "Select peers and document reasons on Peer_Group."
        setup.Range("D4") = "Add verified CSVs to valuation/filings before refreshing."
        setup.Columns("A").ColumnWidth = 40: setup.Columns("B").ColumnWidth = 20: setup.Columns("D").ColumnWidth = 68
        For Each n In Array("RefreshValuation", "ExportValuationPDF", "SubmitReview", "TraceValuation")
            Set addr = setup.Cells(7 + setup.Shapes.Count * 2, 4)
            Set btn = setup.Shapes.AddShape(msoShapeRoundedRectangle, addr.Left, addr.Top, 185, 24)
            btn.OnAction = CStr(n): btn.TextFrame.Characters.Text = CStr(n)
        Next
        setup.Range("A22:A25").Value = Application.Transpose(Array("Review record ID", "Action", "Reviewer", "Reason / approved source"))
        setup.Range("B23").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "Review,Approve,Reject"
        S("Home").Hyperlinks.Add S("Home").Range("A21"), "", "'Valuation_Setup'!A1", , "Open valuation workspace"
    End If
    BuildValuation
    InstallValuation = "OK"
    Exit Function
bad:
    InstallValuation = Err.Description
End Function

Private Sub LoadTables()
    Dim name, rows As Collection, rec, a(), r As Long, c As Long, w As Worksheet, t As String
    For Each name In Array("Trading_Data", "LTM_Reconciliation", "Record_Lineage", "Quality_Control", "Valuation_History", "Review_History", "Observations")
        If CStr(name) = "Record_Lineage" Then
            BulkLineage
            GoTo nextTable
        End If
        Set rows = CsvRows(PathIn(ThisWorkbook.Path, "valuation\tables\" & name & ".csv"))
        rec = rows(1): ReDim a(1 To rows.Count, 1 To UBound(rec) + 1)
        For r = 1 To rows.Count
            rec = rows(r)
            For c = 0 To UBound(rec)
                t = rec(c)
                Dim numericField As Boolean
                Select Case CStr(name)
                    Case "Trading_Data": numericField = c >= 6 And c <= 27
                    Case "LTM_Reconciliation": numericField = c >= 6
                    Case "Quality_Control": numericField = c = 1 Or c >= 3
                    Case "Observations", "Valuation_History": numericField = c = 2 Or c = 3
                    Case Else: numericField = False
                End Select
                If r > 1 And numericField And Matches(t, "^-?[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?$") And Len(t) < 30 Then
                    a(r, c + 1) = Val(t)
                Else
                    a(r, c + 1) = Literal(t)
                End If
            Next
        Next
        Set w = VSheet(CStr(name)): w.UsedRange.Clear
        w.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value2 = a
        Band w.Range("A1").Resize(1, UBound(a, 2)): w.UsedRange.Columns.AutoFit
        For c = 1 To UBound(a, 2)
            If w.Columns(c).ColumnWidth > 55 Then w.Columns(c).ColumnWidth = 55
        Next
        If w.AutoFilterMode Then w.AutoFilterMode = False
        w.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).AutoFilter
nextTable:
    Next
End Sub

Private Sub BulkLineage()
    Dim w As Worksheet, query As QueryTable, types(1 To 22) As Variant, i As Long
    Set w = VSheet("Record_Lineage")
    For i = 1 To 22: types(i) = xlTextFormat: Next
    w.UsedRange.Clear
    Set query = w.QueryTables.Add("TEXT;" & PathIn(ThisWorkbook.Path, "valuation\tables\Record_Lineage.csv"), w.Range("A1"))
    With query
        .TextFilePlatform = 65001: .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True: .TextFileTabDelimiter = False
        .TextFileTextQualifier = xlTextQualifierDoubleQuote: .TextFileColumnDataTypes = types
        .AdjustColumnWidth = False: .Refresh BackgroundQuery:=False: .Delete
    End With
    Band w.Range("A1:V1"): w.Columns("A:V").ColumnWidth = 19
    If w.AutoFilterMode Then w.AutoFilterMode = False
    w.Range("A1").CurrentRegion.AutoFilter
End Sub

Private Function RunEngine(Optional ByVal reviewing As Boolean = False) As Boolean
    Dim w As Worksheet, cmd As String, cli As String, job As Object, cell As Range, arg As String, code As Long
    Set w = S("Valuation_Setup")
    If Not Matches(CStr(w.Range("B4")), "^[A-Z]{3}$") Then Fail "Enter a three-letter reporting currency."
    cli = PathIn(ThisWorkbook.Path, "scripts\valuation_data.py")
    cmd = "python """ & cli & """ --date " & Format$(w.Range("B2"), "yyyy-mm-dd") & " --end " & Format$(w.Range("B3"), "yyyy-mm-dd") & " --currency " & w.Range("B4")
    If reviewing Then
        cmd = cmd & " --review"
        For Each cell In w.Range("B22:B25")
            arg = CStr(cell.Value2)
            If arg = "" Or InStr(arg, """") Or InStr(arg, "%") Or Right$(arg, 1) = "\" Then Fail "Enter all review fields using plain text."
            cmd = cmd & " """ & arg & """"
        Next
    End If
    Set job = CreateObject("WScript.Shell")
    code = job.Run(cmd, 0, True)
    If code <> 0 Then Fail "Data preparation failed. Check valuation inputs and run scripts/valuation_data.py to see the diagnostic."
    RunEngine = True
End Function

Public Sub RefreshValuation()
    On Error GoTo bad
    If RunEngine() Then BuildValuation
    S("Valuation_Summary").Activate
    Exit Sub
bad:
    S("Valuation_Setup").Range("D5") = Err.Description: MsgBox Err.Description, vbExclamation
End Sub

Public Sub SubmitReview()
    On Error GoTo bad
    If RunEngine(True) Then BuildValuation
    Exit Sub
bad:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub BuildValuation()
    Dim oldEvents As Boolean, detail As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Application.EnableEvents = False
    RebuildModel
    Application.EnableEvents = oldEvents
    Exit Sub
failed:
    detail = Err.Description: Application.EnableEvents = oldEvents
    S("Valuation_Setup").Range("D5") = detail
    Fail detail
End Sub

Private Sub RebuildModel()
    Dim w As Worksheet, peer As Worksheet, data As Worksheet, group As Object, i As Long, p As Long, id As String, old, item, c As Long
    Dim cfg As Worksheet, target As String
    Set cfg = S("Valuation_Setup"): cfg.Range("D5") = "Refresh in progress"
    LoadTables
    Set data = S("Trading_Data"): Set peer = VSheet("Peer_Group"): Set group = Dict()
    target = CStr(cfg.Range("B5").Value2)
    Dim history As Worksheet, groupName As String, oldGroup As String, priorGroup As Object, entry As Long, policyKey As String
    Set history = VSheet("Peer_Decisions"): Set priorGroup = Dict()
    Headers history, 1, Split("Timestamp|Group|Issuer|Include|Reason|Workbook user", "|")
    groupName = Trim$(CStr(cfg.Range("B20"))): oldGroup = CStr(peer.Range("L1"))
    If groupName = "" Then Fail "Enter a peer group name."
    For i = 2 To LastR(history)
        policyKey = history.Cells(i, 2) & "|" & history.Cells(i, 3)
        priorGroup(policyKey) = Array(history.Cells(i, 4).Value2, history.Cells(i, 5).Value2)
    Next
    For i = 2 To LastR(peer)
        If oldGroup <> "" Then
            policyKey = oldGroup & "|" & peer.Cells(i, 1)
            Dim changed As Boolean: changed = True
            If priorGroup.Exists(policyKey) Then
                old = priorGroup(policyKey): changed = old(0) <> peer.Cells(i, 2) Or old(1) <> peer.Cells(i, 3)
            End If
            If changed Then
                entry = LastR(history) + 1
                history.Cells(entry, 1).Resize(1, 6).Value = Array(Now, Literal(oldGroup), Literal(peer.Cells(i, 1).Value2), Literal(peer.Cells(i, 2).Value2), Literal(peer.Cells(i, 3).Value2), Literal(Application.UserName))
                priorGroup(policyKey) = Array(peer.Cells(i, 2).Value2, peer.Cells(i, 3).Value2)
            End If
        End If
    Next
    peer.UsedRange.Clear
    Headers peer, 1, Split("Issuer|Include|Selection reason|Industry|Geography|Business model|EV / sales|EV / EBITDA|EV / EBIT|P/E", "|")
    For i = 2 To LastR(data)
        id = CStr(data.Cells(i, 1)): If id = target Then p = i
        peer.Cells(i, 1).Value2 = Literal(id)
        If priorGroup.Exists(groupName & "|" & id) Then
            old = priorGroup(groupName & "|" & id)
        Else
            old = Array("No", "Pending analyst review")
            If groupName = "Industrial components" And CStr(data.Cells(i, 30)) = "Synthetic market case v2" Then old = Array("Yes", "Synthetic component manufacturer")
        End If
        peer.Cells(i, 2).Value2 = Literal(old(0)): peer.Cells(i, 3).Value2 = Literal(old(1))
        For c = 4 To 6: peer.Cells(i, c).Value2 = Literal(data.Cells(i, c - 2).Value2): Next
        For c = 7 To 10
            Dim ref As String: ref = "Trading_Data!" & data.Cells(i, c + 14).Address
            peer.Cells(i, c).Formula = "=IF(AND(A" & i & "<>Valuation_Setup!B$5,B" & i & "=""Yes"",C" & i & "<>"""",ISNUMBER(" & ref & "))," & ref & ","""")"
        Next
    Next
    peer.Columns("A:J").AutoFit: peer.Range("B:C").Font.Color = vbBlue
    peer.Range("L1") = Literal(groupName)
    history.Columns("A:F").AutoFit
    If LastR(data) >= 2 Then
        ThisWorkbook.Names.Add Name:="ValuationIssuers", RefersTo:="=Trading_Data!$A$2:$A$" & LastR(data)
        cfg.Range("B5").Validation.Delete: cfg.Range("B5").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "=ValuationIssuers"
        peer.Range("B2:B" & LastR(data)).Validation.Delete
        peer.Range("B2:B" & LastR(data)).Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "Yes,No"
    End If
    If p = 0 Then
        VSheet("Valuation_Summary").UsedRange.ClearContents
        Fail "Target has no eligible financial and market data at the chosen dates."
    End If
    BuildDCF p
    BuildRanges p, LastR(data)
    cfg.Range("D5") = "Current as of " & Format$(cfg.Range("B2"), "yyyy-mm-dd")
    Audit "Valuation", "OK", LastR(data) - 1, target & " valuation refreshed.", 0
End Sub

Public Sub ExportValuationPDF()
    On Error GoTo bad
    If RunEngine() Then BuildValuation
    Dim path As String, folder As String
    folder = PathIn(ThisWorkbook.Path, "reports")
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    path = PathIn(folder, "Valuation_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf")
    If Dir$(path) <> "" Then Fail "A PDF already exists for this timestamp."
    S("Valuation_Summary").ExportAsFixedFormat xlTypePDF, path
    Audit "Valuation PDF", "OK", 1, path, 0: Exit Sub
bad:
    MsgBox Err.Description, vbExclamation
End Sub
