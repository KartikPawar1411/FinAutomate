Attribute VB_Name = "modProvenance"
Option Explicit

Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef buffer As Any) As Long

Public Function NewBatchID() As String
    Dim bytes(0 To 15) As Byte, i As Long
    If CoCreateGuid(bytes(0)) <> 0 Then Fail "Cannot create import batch ID."
    For i = 0 To 15: NewBatchID = NewBatchID & Right$("0" & Hex$(bytes(i)), 2): Next
End Function

Public Function FileDigest(ByVal path As String) As String
    Dim proc As Object, tool As String, lines, line, value As String
    tool = Environ$("SystemRoot") & "\System32\certutil.exe"
    Set proc = CreateObject("WScript.Shell").Exec("""" & tool & """ -hashfile """ & path & """ SHA256")
    lines = Split(proc.StdOut.ReadAll, vbLf)
    For Each line In lines
        value = LCase$(Trim$(Replace(CStr(line), vbCr, "")))
        If Matches(value, "^[a-f0-9]{64}$") Then FileDigest = value: Exit Function
    Next
    Fail "Could not fingerprint the source file."
End Function

Public Sub TraceValuation()
    Dim id As String, tbl As Worksheet, hit As Range, sourceIDs, ledger As Worksheet, criteria(), i As Long
    id = CStr(S("Valuation_Setup").Range("B5").Value2)
    Set tbl = S("Trading_Data")
    Set hit = tbl.Columns(1).Find(id, LookIn:=xlValues, LookAt:=xlWhole)
    If hit Is Nothing Then Exit Sub
    sourceIDs = Split(CStr(tbl.Cells(hit.Row, 31).Value2), " ")
    ReDim criteria(0 To UBound(sourceIDs))
    For i = 0 To UBound(sourceIDs): criteria(i) = Mid$(sourceIDs(i), 2): Next
    Set ledger = S("Record_Lineage")
    ledger.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:=criteria, Operator:=xlFilterValues
    ledger.Activate
End Sub
