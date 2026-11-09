Attribute VB_Name = "modMain"
Option Explicit

Public Sub InitializeApplication()
    Dim w As Worksheet, item, rowNo As Long, actions, macros, descriptions, i As Long, box As Shape, spot As Range
    Set w = S("Home"): ClearView w
    For i = w.Shapes.Count To 1 Step -1: w.Shapes(i).Delete: Next
    Title w, "FinAutomate", "INVESTMENT BANKING  /  Financial automation and reporting"
    w.Range("A4") = "Report type": w.Range("B4") = "Company"
    w.Range("B4").Validation.Delete: w.Range("B4").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "Company,Comparison"
    w.Range("B4").Interior.Color = RGB(219, 238, 241): w.Range("B4").Font.Color = vbBlue
    w.Range("C4:H4").Merge: w.Range("C4") = "  Choose an issuer on Dashboard, or switch to Comparison here."
    w.Range("A5") = "Last action": w.Range("B5:H5").Merge: w.Range("B5").WrapText = True: w.Rows(5).RowHeight = 34
    w.Range("A7:B7").Merge: w.Range("A7") = "Workflow": w.Range("C7:H7").Merge: w.Range("C7") = "What the action does": Band w.Range("A7:H7")
    actions = Split("Import financial data|Clean financial data|Validate statements|Calculate financial ratios|Apply banking formatting|Refresh dashboard|Generate financial report|Export PDF|View automation log|Reset demo data|Run full workflow", "|")
    macros = Split("RunImportData RunCleanData RunValidation RunCalculateRatios RunIBFormatting RunRefreshDashboard RunGenerateReport RunExportPDF RunViewLog RunResetDemoData RunFullPipeline", " ")
    descriptions = Split("Append one or more CSV or XLSX extracts; original files stay unchanged.|Normalize units and dates; retain source references and row decisions.|Reconcile identities and separate accepted filings from exceptions.|Build 17 formula-driven metrics for each accepted company-year.|Style inputs and references, freeze headers and find formula errors.|Rebuild issuer trends, KPI summaries and the peer table.|Prepare a company review or peer comparison using the report type above.|Refresh and save the selected report in the reports folder.|Inspect counts, elapsed time and errors for each step.|Back up this workbook, reload six synthetic issuers and run the workflow.|Rebuild analysis and report from the current raw data.", "|")
    For i = 0 To UBound(actions)
        rowNo = 8 + i: w.Rows(rowNo).RowHeight = 31
        Set spot = w.Cells(rowNo, 1).Resize(1, 2)
        Set box = w.Shapes.AddShape(msoShapeRoundedRectangle, spot.Left + 2, spot.Top + 3, spot.Width - 7, spot.Height - 6)
        box.Name = CStr(macros(i)): box.OnAction = CStr(macros(i)): box.TextFrame.Characters.Text = actions(i)
        box.TextFrame.Characters.Font.Color = vbWhite: box.TextFrame.Characters.Font.Bold = True
        box.TextFrame.HorizontalAlignment = xlHAlignCenter: box.TextFrame.VerticalAlignment = xlVAlignCenter
        box.Fill.ForeColor.RGB = IIf(i = 10, RGB(0, 108, 103), RGB(37, 72, 94)): box.Line.Visible = msoFalse
        w.Cells(rowNo, 3).Resize(1, 6).Merge: w.Cells(rowNo, 3) = descriptions(i)
        w.Cells(rowNo, 3).VerticalAlignment = xlCenter
        If i Mod 2 = 0 Then w.Cells(rowNo, 3).Resize(1, 6).Interior.Color = RGB(239, 245, 246)
    Next
    For Each item In Array("Raw_Data", "Clean_Data", "Financial_Statements"): Headers S(CStr(item)), 1, Schema(): Next
    For Each item In Array("Raw_Data", "Clean_Data", "Financial_Statements")
        S(CStr(item)).Range("AN1:AQ1").Value = Array("Record_ID", "Batch_ID", "Source_SHA256", "Source_worksheet")
        Band S(CStr(item)).Range("AN1:AQ1")
    Next
    Headers S("Audit_Log"), 1, Split("Timestamp|Macro|Outcome|Records|Details|Seconds", "|")
    Set w = S("Settings"): w.UsedRange.Clear
    Headers w, 1, Array("Setting", "Value")
    w.Range("A2") = "Input basis": w.Range("B2") = "Annual consolidated actuals"
    w.Range("A3") = "Output units": w.Range("B3") = "Millions in filing currency"
    w.Range("A4") = "Processing status": w.Range("B4") = "Ready for input"
    w.Range("A5") = "Rejected rows": w.Range("B5") = 0: w.Columns("A:B").AutoFit
    Set w = S("Instructions"): ClearView w: Title w, "Operator guide", "First run  /  use the saved synthetic example or import your own annual extracts."
    Dim lines
    lines = Array("1. Extract the full project before opening the workbook. Buttons require Excel Desktop for Windows.", "2. Start on Dashboard. Change the blue company ID to see another issuer. Use Home for all workflow controls.", "3. Import one annual, consolidated filing per company/FY. See templates/import_template.csv for the input fields.", "4. Use Run full workflow after imports. Filter Clean_Data by Rejected or Duplicate to trace excluded rows.", "5. Choose Company or Comparison on Home, then Export PDF. Exports are saved beside the workbook in reports/.", "6. A blank is unknown, not zero. N/A means missing input/history; N/M means a nonmeaningful denominator or multiple.", "7. See docs/ for formula definitions, currency/date rules, troubleshooting, testing and benchmark scope.")
    For i = 0 To UBound(lines): w.Cells(i + 4, 1).Resize(1, 8).Merge: w.Cells(i + 4, 1) = lines(i): w.Rows(i + 4).RowHeight = 26: Next
    S("Home").Activate
    ThisWorkbook.Windows(1).DisplayGridlines = False
End Sub

Public Sub LoadDemo()
    Dim files, f, rec, all As New Collection, part As Collection, started As Double
    started = Timer
    files = Split("company_acme.csv company_databridge.csv company_energyflow.csv company_globalbank.csv company_medlife.csv company_techvista.csv", " ")
    For Each f In files
        Set part = ReadSource(PathIn(PathIn(ThisWorkbook.Path, "sample_data"), CStr(f)))
        For Each rec In part: all.Add rec: Next
    Next
    CommitRows S("Raw_Data"), all, True
    S("Settings").Range("B4") = "Pending refresh"
    Audit "Load demo", "OK", all.Count, "Six source files loaded.", Took(started)
End Sub

Public Sub ProcessTo(ByVal stage As Long)
    Dim t As Double: t = Timer
    S("Settings").Range("B4") = "Pending refresh"
    CleanFinancialData
    If stage >= 2 Then ValidateFinancialStatements
    If stage >= 3 Then CalculateFinancialRatios
    If stage >= 4 Then RefreshDashboard
    If stage >= 5 Then
        ApplyIBFormatting
        BuildReport
        S("Settings").Range("B4") = "Current"
        Audit "Full workflow", "OK", LastR(S("Raw_Data")) - 1, S("Settings").Range("B5") & " rejected; analysis and report refreshed.", Took(t)
    End If
End Sub

Public Sub Perform(ByVal action As String)
    Dim oldScreen As Boolean, oldEvents As Boolean, t As Double, detail As String, snapshot As String
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents: t = Timer
    On Error GoTo failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Select Case action
        Case "Import": PickFiles
        Case "Clean": ProcessTo 1
        Case "Validate": ProcessTo 2
        Case "Ratios": ProcessTo 3
        Case "Format": ApplyIBFormatting
        Case "Dashboard": ProcessTo 5: S("Dashboard").Activate
        Case "Report": ProcessTo 5: S("Report").Activate
        Case "PDF": ProcessTo 5: detail = ExportCurrentPDF()
        Case "Log": S("Audit_Log").Activate
        Case "Reset"
            If MsgBox("Replace raw data with the sample portfolio? A workbook backup will be saved beside this file first.", vbYesNo + vbDefaultButton2, "Reset demo") <> vbYes Then GoTo done
            snapshot = PathIn(ThisWorkbook.Path, "Before_reset_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm")
            If Dir$(snapshot) <> "" Then Fail "A backup already exists for this second. Retry."
            ThisWorkbook.SaveCopyAs snapshot: LoadDemo: ProcessTo 5
        Case "Full": ProcessTo 5
    End Select
done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: Exit Sub
failed:
    detail = Err.Description
    Audit action, "ERROR", 0, detail, Took(t)
    MsgBox detail, vbExclamation, "FinAutomate": Resume done
End Sub

Public Sub RunImportData()
    Perform "Import"
End Sub

Public Sub RunCleanData()
    Perform "Clean"
End Sub

Public Sub RunValidation()
    Perform "Validate"
End Sub

Public Sub RunCalculateRatios()
    Perform "Ratios"
End Sub

Public Sub RunIBFormatting()
    Perform "Format"
End Sub

Public Sub RunRefreshDashboard()
    Perform "Dashboard"
End Sub

Public Sub RunGenerateReport()
    Perform "Report"
End Sub

Public Sub RunExportPDF()
    Perform "PDF"
End Sub

Public Sub RunViewLog()
    Perform "Log"
End Sub

Public Sub RunResetDemoData()
    Perform "Reset"
End Sub

Public Sub RunFullPipeline()
    Perform "Full"
End Sub
