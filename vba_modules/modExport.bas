Attribute VB_Name = "modExport"
Option Explicit

Public Function TryExportPDF() As String
    Dim path As String
    On Error GoTo failed
    path = ExportCurrentPDF(): Exit Function
failed:
    TryExportPDF = Err.Description: Audit "Export PDF", "ERROR", 0, TryExportPDF, 0
End Function

Public Function ExportCurrentPDF() As String
    Dim folder As String, path As String, stem As String, suffix As Long, t As Double
    t = Timer: folder = PathIn(ThisWorkbook.Path, "reports")
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    If S("Home").Range("B4") = "Comparison" Then stem = "Comparison" Else stem = "Company_" & S("Dashboard").Range("B4").Value2
    stem = stem & "_" & Format$(Now, "yyyymmdd_hhnnss")
    path = PathIn(folder, stem & ".pdf")
    Do While Dir$(path) <> "": suffix = suffix + 1: path = PathIn(folder, stem & "_" & suffix & ".pdf"): Loop
    S("Report").ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If FileLen(path) < 1 Then Fail "Excel created an empty PDF."
    ExportCurrentPDF = path: Audit "Export PDF", "OK", 1, Mid$(path, InStrRev(path, Application.PathSeparator) + 1), Took(t)
End Function
