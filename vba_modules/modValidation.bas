Attribute VB_Name = "modValidation"
Option Explicit

' Reconciliation tolerance in millions: max(0.000001, 0.01% of reported balance).
Private Sub Tie(ByVal lhs As Variant, ByVal first As Variant, ByVal second As Variant, ByVal sign As Long, ByVal label As String, ByRef problems As String)
    If VarType(lhs) <> vbDouble Or VarType(first) <> vbDouble Or VarType(second) <> vbDouble Then Exit Sub
    If Abs(lhs - first - sign * second) > Application.Max(0.000001, Abs(lhs) * 0.0001) Then Note problems, label
End Sub

Public Sub CheckIdentities(ByRef v As Variant, ByRef problems As String)
    Tie v(9), v(7), v(8), -1, "Gross profit does not reconcile", problems
    Tie v(20), v(25), v(29), 1, "Balance sheet does not balance", problems
    Tie v(19), v(17), v(18), -1, "Net income does not reconcile", problems
End Sub

Public Sub ValidateFinancialStatements()
    Dim w As Worksheet, a, v(), c As Long, r As Long, keep As New Collection, errors As Long, t As Double
    t = Timer: Set w = S("Clean_Data")
    If LastR(w) < 2 Then Fail "Clean data is empty."
    a = w.Range("A1").Resize(LastR(w), WIDTH).Value2
    For r = 2 To UBound(a, 1)
        If a(r, 38) = "Accepted" Then
            ReDim v(1 To WIDTH)
            For c = 1 To WIDTH: v(c) = a(r, c): Next
            keep.Add v
        ElseIf a(r, 38) = "Rejected" Then
            errors = errors + 1
        End If
    Next
    Set w = S("Financial_Statements"): CommitRows w, keep, True
    If keep.Count > 1 Then w.Range("A1").Resize(keep.Count + 1, WIDTH).Sort Key1:=w.Range("A2"), Order1:=xlAscending, Key2:=w.Range("D2"), Order2:=xlAscending, Header:=xlYes
    w.Columns(4).NumberFormat = "yyyy-mm-dd"
    S("Settings").Range("B5").Value2 = errors
    Audit "Validate", IIf(errors > 0, "REVIEW", "OK"), keep.Count, errors & " rejected; " & keep.Count & " accepted.", Took(t)
End Sub
