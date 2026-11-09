Attribute VB_Name = "modLogging"
Option Explicit

Public Sub Audit(ByVal task As String, ByVal outcome As String, ByVal n As Long, ByVal detail As String, ByVal seconds As Double)
    Dim w As Worksheet, r As Long
    Set w = S("Audit_Log"): r = LastR(w) + 1
    w.Cells(r, 1).Value = Now
    w.Cells(r, 2).Resize(1, 5).Value = Array(task, outcome, n, Left$(detail, 30000), seconds)
    w.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    w.Cells(r, 6).NumberFormat = "0.000"
    S("Home").Range("B5").Value2 = Left$(task & " | " & outcome & " | " & detail, 180)
End Sub
