VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "Dashboard" Then Exit Sub
    If Intersect(Target, Sh.Range("B4")) Is Nothing Then Exit Sub
    Dim msg As String: On Error GoTo failed
    Application.EnableEvents = False
    RefreshDashboard
done:
    Application.EnableEvents = True
    Exit Sub
failed:
    msg = Err.Description: Audit "Company selector", "ERROR", 0, msg, 0: Resume done
End Sub
