Attribute VB_Name = "modUtilities"
Option Explicit

Public Const WIDTH As Long = 43
Public Const AMOUNT As String = "#,##0.0;(#,##0.0);0.0"
Public Const PCT As String = "0.0%;(0.0%);0.0%"
Public Const MULT As String = "0.00""x"";(0.00""x"");0.00""x"""
Public Function S(ByVal n As String) As Worksheet
    Set S = ThisWorkbook.Worksheets(n)
End Function

Public Function Schema() As Variant
    Schema = Split("Company_ID,Company_Name,Reporting_Period,Period_End_Date,Currency,Units,Revenue,COGS,Gross_Profit,SGA,RnD,Other_OpEx,EBITDA,Depreciation_Amortization,EBIT,Interest_Expense,EBT,Income_Tax,Net_Income,Total_Assets,Current_Assets,Cash_Equivalents,Accounts_Receivable,Inventory,Total_Liabilities,Current_Liabilities,Long_Term_Debt,Total_Debt,Total_Equity,Shares_Outstanding,Stock_Price,Market_Cap,Enterprise_Value,Dividends_Paid,CapEx,Source_File,Import_Timestamp,Validation_Status,Cleaning_Notes", ",")
End Function

Public Function LastR(ByVal w As Worksheet) As Long
    Dim f As Range
    Set f = w.Cells.Find(What:="*", After:=w.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If f Is Nothing Then LastR = 1 Else LastR = f.Row
End Function

Public Sub Headers(ByVal w As Worksheet, ByVal rowNo As Long, ByVal names As Variant)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        w.Cells(rowNo, 1 + i - LBound(names)).Value2 = names(i)
    Next
    Band w.Cells(rowNo, 1).Resize(1, UBound(names) - LBound(names) + 1)
End Sub

Public Function Dict() As Object
    Set Dict = CreateObject("Scripting.Dictionary")
End Function

Public Sub Fail(ByVal detail As String)
    Err.Raise vbObjectError + 701, "FinAutomate", detail
End Sub

Public Function Literal(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then
        If Len(v) > 32000 Then Fail "Input cell exceeds 32,000 characters."
        If Len(v) > 0 Then Literal = "'" & v
    Else
        Literal = v
    End If
End Function

Public Function Matches(ByVal value As String, ByVal pattern As String) As Boolean
    Static rx As Object
    If rx Is Nothing Then Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern
    Matches = rx.Test(value)
End Function

Public Function Took(ByVal started As Double) As Double
    Took = Timer - started
    If Took < 0 Then Took = Took + 86400
End Function

Public Function PeriodDate(ByVal t As String, ByRef serial As Double) As Boolean
    Dim a, y As Long, m As Long, d As Long, dt As Date
    If Matches(t, "^[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}$") Then
        a = Split(t, "-"): y = CLng(a(0)): m = CLng(a(1)): d = CLng(a(2))
    ElseIf Matches(t, "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$") Then
        a = Split(t, "/"): y = CLng(a(2)): m = CLng(a(0)): d = CLng(a(1))
    Else
        Exit Function
    End If
    If y < 1901 Or y > 9998 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(y, m, d)
    If Day(dt) <> d Then Exit Function
    serial = CDbl(dt): PeriodDate = True
End Function

Public Sub Note(ByRef memo As String, ByVal msg As String)
    If memo <> "" Then memo = memo & "; "
    memo = memo & msg
End Sub

Public Sub CommitRows(ByVal w As Worksheet, ByVal records As Collection, ByVal replaceData As Boolean)
    Dim a(), v, r As Long, c As Long, first As Long, oldEnd As Long
    oldEnd = LastR(w): first = oldEnd + 1
    If replaceData Then first = 2
    If first + records.Count - 1 > w.Rows.Count Then Fail "Worksheet row limit reached."
    If records.Count > 0 Then
        ReDim a(1 To records.Count, 1 To WIDTH)
        For Each v In records
            r = r + 1
            For c = 1 To WIDTH
                a(r, c) = Literal(v(c))
            Next
        Next
        w.Cells(first, 1).Resize(records.Count, WIDTH).Value2 = a
    End If
    If replaceData And oldEnd >= first + records.Count Then w.Range(w.Cells(first + records.Count, 1), w.Cells(oldEnd, WIDTH)).ClearContents
End Sub

Public Function PathIn(ByVal folder As String, ByVal file As String) As String
    PathIn = folder & Application.PathSeparator & file
End Function
