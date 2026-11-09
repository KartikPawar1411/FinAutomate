Attribute VB_Name = "modCleaning"
Option Explicit

Private Function ParseAmount(ByVal v As Variant, ByRef amountOut As Double) As Boolean
    Dim t As String
    If VarType(v) = vbDouble Or VarType(v) = vbLong Then
        amountOut = CDbl(v): ParseAmount = True: Exit Function
    End If
    t = Trim$(CStr(v))
    If Left$(t, 1) = "(" And Right$(t, 1) = ")" Then t = "-" & Mid$(t, 2, Len(t) - 2)
    If Not Matches(t, "^[+-]?([0-9]+|[0-9]{1,3}(,[0-9]{3})+)(\.[0-9]+)?$") Then Exit Function
    On Error GoTo bad
    amountOut = Val(Replace(t, ",", "")): ParseAmount = True
bad:
End Function

Public Sub CleanFinancialData()
    Dim w As Worksheet, a, v(), r As Long, c As Long, h, memo As String, problems As String, t As Double
    Dim val As Double, dt As Double, factor As Double, key As String, sig As String, part As String
    Dim seen As Object, variants As Object, collected As New Collection, out As New Collection, item
    Dim converted As Long, trimmed As Long, duplicateCount As Long
    Set w = S("Raw_Data"): t = Timer: h = Schema()
    If LastR(w) < 2 Then Fail "No input rows. Import a file or reset demo data."
    a = w.Range("A1").Resize(LastR(w), WIDTH).Value2: Set seen = Dict(): Set variants = Dict()
    For r = 2 To UBound(a, 1)
        ReDim v(1 To WIDTH): memo = "": problems = "": trimmed = 0: converted = 0
        For c = 1 To 37
            v(c) = a(r, c)
            If VarType(v(c)) = vbString Then
                v(c) = Trim$(Replace(v(c), ChrW(160), " "))
                If v(c) <> a(r, c) Then trimmed = trimmed + 1
            End If
        Next
        For c = 40 To WIDTH: v(c) = a(r, c): Next
        If trimmed > 0 Then Note memo, "Trimmed " & trimmed & " fields"
        If v(1) <> UCase$(CStr(v(1))) Then Note memo, "Uppercased identifier"
        v(1) = UCase$(CStr(v(1))): v(3) = UCase$(CStr(v(3))): v(5) = UCase$(CStr(v(5)))
        If Not Matches(CStr(v(1)), "^[A-Z0-9][A-Z0-9._-]{0,31}$") Then Note problems, "Invalid or missing Company_ID"
        If Len(CStr(v(2))) = 0 Then Note problems, "Missing Company_Name"
        If Not Matches(CStr(v(3)), "^FY[0-9]{4}$") Then Note problems, "Expected FYyyyy"
        If PeriodDate(CStr(v(4)), dt) Then
            v(4) = dt: Note memo, "Parsed end date"
            If v(3) <> "FY" & Year(CDate(dt)) Then Note problems, "Period/year mismatch"
        Else
            Note problems, "Invalid end date"
        End If
        If Not Matches(CStr(v(5)), "^[A-Z]{3}$") Then Note problems, "Invalid currency code"
        If InStr(1, "|USD|EUR|GBP|INR|CAD|AUD|JPY|CHF|CNY|HKD|SGD|", "|" & v(5) & "|", vbBinaryCompare) = 0 Then Note problems, "Unsupported currency"
        Select Case LCase$(CStr(v(6)))
            Case "actual", "units": factor = 0.000001
            Case "thousands": factor = 0.001
            Case "millions": factor = 1
            Case "billions": factor = 1000
            Case Else: factor = 0: Note problems, "Unknown unit"
        End Select
        If factor <> 0 Then
            If factor <> 1 Then Note memo, v(6) & " to millions; price per share unchanged"
            v(6) = "Millions"
        End If
        For c = 7 To 35
            If Len(CStr(v(c))) = 0 Then
                v(c) = Empty
                If c = 7 Then Note problems, "Missing Revenue"
            ElseIf ParseAmount(v(c), val) Then
                If VarType(v(c)) = vbString Then converted = converted + 1
                If c <> 31 And factor <> 0 Then val = val * factor
                v(c) = val
            Else
                Note problems, "Invalid " & h(c - 1)
            End If
        Next
        If converted > 0 Then Note memo, "Parsed " & converted & " numeric fields"
        If VarType(v(7)) = vbDouble Then
            If v(7) < 0 Then Note problems, "Negative revenue"
        End If
        CheckIdentities v, problems
        key = CStr(v(1)) & "|" & CStr(v(3)): sig = ""
        For c = 1 To 35
            part = CStr(v(c)): sig = sig & Len(part) & ":" & part
        Next
        If seen.Exists(sig) Then
            v(38) = "Duplicate": Note memo, "Duplicate of " & seen(sig): duplicateCount = duplicateCount + 1
        Else
            seen.Add sig, v(36)
            If variants.Exists(key) Then variants(key) = 2 Else variants.Add key, 1
            If problems = "" Then v(38) = "Accepted" Else v(38) = "Rejected"
        End If
        If problems <> "" Then Note memo, problems
        v(39) = memo: collected.Add v
    Next
    For Each item In collected
        key = CStr(item(1)) & "|" & CStr(item(3))
        If variants(key) > 1 And item(38) <> "Duplicate" Then
            item(38) = "Rejected": memo = CStr(item(39)): Note memo, "Conflicting filings for company/FY": item(39) = memo
        End If
        out.Add item
    Next
    CommitRows S("Clean_Data"), out, True
    S("Clean_Data").Columns(4).NumberFormat = "yyyy-mm-dd"
    Audit "Clean", "OK", out.Count, CStr(duplicateCount) & " exact duplicate(s) excluded from analysis. See Cleaning_Notes.", Took(t)
End Sub
