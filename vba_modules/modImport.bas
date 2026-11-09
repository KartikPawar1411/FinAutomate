Attribute VB_Name = "modImport"
Option Explicit
Private sourceTab As String

' All file decoding and header checks complete before a single append/replace.
Public Function ReadSource(ByVal path As String) As Collection
    Dim rows As Collection, rec, head, mapping As Object, canonical, ix(), j As Long, i As Long
    Dim payload As Collection, v(), field As String, ext As String, name As String
    Set payload = New Collection: name = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    If Len(Dir$(path)) = 0 Then Fail "Missing source: " & name
    If FileLen(path) = 0 Then Fail "Empty source: " & name
    ext = LCase$(Mid$(name, InStrRev(name, ".") + 1))
    Select Case ext
        Case "csv": Set rows = CsvRows(path)
        Case "xlsx": Set rows = SheetRows(path)
        Case Else: Fail "Expected CSV or XLSX: " & name
    End Select
    If rows.Count < 2 Then Fail "No data rows: " & name
    head = rows(1): canonical = Schema(): Set mapping = Dict()
    mapping.CompareMode = vbTextCompare
    For i = 0 To UBound(head)
        field = Trim$(CStr(head(i)))
        If mapping.Exists(field) Then Fail "Repeated header: " & field
        mapping.Add field, i
    Next
    ReDim ix(1 To 35)
    For j = 1 To 35
        ix(j) = -1
        If mapping.Exists(canonical(j - 1)) Then
            ix(j) = mapping(canonical(j - 1))
        ElseIf j <= 7 Then
            Fail "Missing column " & canonical(j - 1) & ": " & name
        End If
    Next
    Dim digest As String, batch As String
    digest = FileDigest(path): batch = NewBatchID()
    For i = 2 To rows.Count
        rec = rows(i)
        If UBound(rec) <> UBound(head) Then Fail name & ": wrong field count at record " & i
        If mapping.Exists("Statement_Basis") Then
            If LCase$(Trim$(CStr(rec(mapping("Statement_Basis"))))) <> "annual consolidated actuals" Then Fail name & ": unsupported statement basis at record " & i
        End If
        ReDim v(1 To WIDTH)
        For j = 1 To 35
            If ix(j) >= 0 Then v(j) = rec(ix(j))
        Next
        v(36) = name & " #" & i: v(37) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        v(40) = digest & "-" & i: v(41) = batch: v(42) = digest: v(43) = IIf(ext = "csv", "CSV", sourceTab)
        payload.Add v
    Next
    Set ReadSource = payload
End Function

Public Function CsvRows(ByVal path As String) As Collection
    Dim stream As Object, data As String, ch As String, token As String, fields As Collection, rows As New Collection
    Dim i As Long, inside As Boolean, closed As Boolean, items(), j As Long
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path: data = stream.ReadText: stream.Close
    Set fields = New Collection
    data = Replace(Replace(data, vbCrLf, vbLf), vbCr, vbLf)
    If Right$(data, 1) <> vbLf Then data = data & vbLf
    For i = 1 To Len(data)
        ch = Mid$(data, i, 1)
        If inside Then
            If ch = """" Then
                If Mid$(data, i + 1, 1) = """" Then
                    token = token & """": i = i + 1
                Else
                    inside = False: closed = True
                End If
            Else
                token = token & ch
            End If
        ElseIf ch = "," Or ch = vbLf Then
            If ch = vbLf And fields.Count = 0 And token = "" And Not closed Then GoTo advance
            fields.Add token: token = "": closed = False
            If ch = vbLf Then
                ReDim items(0 To fields.Count - 1)
                For j = 1 To fields.Count: items(j - 1) = fields(j): Next
                rows.Add items: Set fields = New Collection
            End If
        ElseIf ch = """" And token = "" And Not closed Then
            inside = True
        Else
            If closed Or ch = """" Then Fail "Malformed CSV quoting at record " & rows.Count + 1
            token = token & ch
        End If
advance:
    Next
    If inside Then Fail "Unclosed CSV quotation."
    Set CsvRows = rows
End Function

Private Function SheetRows(ByVal path As String) As Collection
    Dim b As Workbook, w As Worksheet, a, items(), r As Long, c As Long, last As Long, rows As New Collection
    Dim msg As String, opened As Workbook, val As Variant, is1904 As Boolean
    For Each opened In Application.Workbooks
        If StrComp(opened.FullName, path, vbTextCompare) = 0 Then Fail "Close the open input workbook before importing it."
    Next
    On Error GoTo bad
    Set b = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Password:="", WriteResPassword:="")
    Set w = b.Worksheets(1): is1904 = b.Date1904: sourceTab = w.Name
    last = w.Cells(1, w.Columns.Count).End(xlToLeft).Column
    If LastR(w) < 2 Or last < 2 Then Fail "Missing XLSX headers or rows."
    a = w.Range(w.Cells(1, 1), w.Cells(LastR(w), last)).Value2
    For r = 1 To UBound(a, 1)
        ReDim items(0 To last - 1)
        For c = 1 To last
            val = a(r, c)
            If IsError(val) Then val = "#INPUT_ERROR"
            If r > 1 And LCase$(Trim$(CStr(a(1, c)))) = "period_end_date" And VarType(val) = vbDouble Then
                val = Format$(CDate(val - 1462 * CLng(is1904)), "yyyy-mm-dd")
            End If
            items(c - 1) = val
        Next
        rows.Add items
    Next
    b.Close False: Set SheetRows = rows: Exit Function
bad:
    msg = Err.Description
    If Not b Is Nothing Then b.Close False
    Fail msg
End Function

Public Sub ImportPath(ByVal path As String)
    Dim rows As Collection, t As Double, msg As String
    t = Timer: On Error GoTo bad
    Set rows = ReadSource(path)
    CommitRows S("Raw_Data"), rows, False
    S("Settings").Range("B4").Value2 = "Pending refresh"
    Audit "Import", "OK", rows.Count, "Staged " & rows.Count & " source rows; refresh analysis.", Took(t)
    Exit Sub
bad:
    msg = Err.Description: Audit "Import", "ERROR", 0, msg, Took(t): Fail msg
End Sub

' Noninteractive automation boundary: catch VBA errors before returning to COM.
Public Function TryImportPath(ByVal path As String) As String
    On Error GoTo rejected
    ImportPath path
    Exit Function
rejected:
    TryImportPath = Err.Description
End Function

Public Sub PickFiles()
    Dim fd As FileDialog, path As Variant, failures As Long, saved As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select annual statement extracts": .AllowMultiSelect = True
        .Filters.Clear: .Filters.Add "Statement files", "*.csv;*.xlsx"
        If .Show <> -1 Then Exit Sub
    End With
    For Each path In fd.SelectedItems
        On Error Resume Next
        Err.Clear: ImportPath CStr(path)
        If Err.Number <> 0 Then failures = failures + 1: saved = Err.Description
        On Error GoTo 0
    Next
    If failures > 0 Then Fail CStr(failures) & " source file(s) rejected; see Audit_Log. " & saved
End Sub
