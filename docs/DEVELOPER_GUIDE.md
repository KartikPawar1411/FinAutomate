# Build and maintain

For the expanded version, keep valuation/ and its staged tables beside the workbook. The automatic build imports all 15 standard modules and invokes InstallValuation after the annual pipeline. On a manual build, complete the annual steps below, open the VBA Immediate Window (Ctrl+G), enter `? InstallValuation()`, confirm it returns OK, then save. Annual ledgers now carry four provenance columns after the original 39. See [valuation architecture and runtime dependencies](VALUATION_PLATFORM.md) and the [release verification report](UPGRADE_REPORT.md).

## Native build
Requirements: Windows, Excel Desktop with VBA, PowerShell, and a writable project folder. The **checked-in template** is sufficient; Excel creates all native formulas, shapes and charts. Close the output workbook. From the project root:

```powershell
.\scripts\assemble_workbook.ps1 -OutputName FinAutomate.rebuilt.xlsm
```

The script uses a separate hidden Excel instance, opens the macro-free template, imports the 15 exported .bas files, inserts the event body into the existing workbook component, builds the dashboard from the samples, and smoke-tests before moving the result into place. Omitting OutputName replaces the configured release file after success. It does not edit registry keys, lower macro security, auto-launch the result, or require Python. For an automated build, the Excel profile must already permit VBProject automation. If organizational policy disallows it, use the manual build.

## Manual assembly without programmatic project access
1. Open FinAutomate_Template.xlsx and Save As **Excel Macro-Enabled Workbook (*.xlsm)** in the same project folder with sample_data beside it.
2. Alt+F11 → select that workbook's VBAProject → File → Import File (Ctrl+M). Import every .bas in vba_modules, exactly once. Do **not** import ThisWorkbook.cls as a standalone class.
3. Double-click **ThisWorkbook** under Microsoft Excel Objects. Copy the text from `Option Explicit` to the end of ThisWorkbook.cls into the code pane; omit VERSION/BEGIN/Attribute export metadata.
4. Debug → Compile VBAProject. In Excel, Alt+F8 → **InitializeApplication** → Run, then **LoadDemo**, then **RunFullPipeline**. Save. Initialization creates the actual Home buttons; do not reinitialize just to refresh a report.
5. To repair an individual control, right-click its shape → Assign Macro → select its Run... wrapper from this workbook. Or use Insert → Shapes, draw the shape, and assign a wrapper. Labels and assignments, in order: Import / RunImportData; Clean / RunCleanData; Validate / RunValidation; Ratios / RunCalculateRatios; Format / RunIBFormatting; Dashboard / RunRefreshDashboard; Report / RunGenerateReport; PDF / RunExportPDF; Log / RunViewLog; Reset / RunResetDemoData; Full workflow / RunFullPipeline.

## Architecture
modUtilities owns the 43-column annual ledger schema, date parsing, safe cell writes, array commits and worksheet references. modImport maps headers and parses sources before committing a single file. modCleaning normalizes records and indexes full signatures and issuer/FY keys. modValidation owns accounting tie-outs and the accepted statement projection. modFinancialRatios writes guarded formulas in a rectangular block. modDashboard owns selector, KPI links, chart ranges and peers; ThisWorkbook handles the single selector change event. modReporting builds the report snapshot, modExport writes the PDF, modFormatting styles financial sheets, modLogging appends stage outcomes, and modMain controls sequencing and restores UI state.

Business sheets use named worksheets, canonical headers and dynamic row extents rather than a fixed sample row count. The dashboard layout and schema column indexes are intentionally stable. Comparisons and prior-year references are regenerated after sorting accepted filings. Hidden helper columns in dashboard/report charts store only numeric observations so missing chart values are not drawn as zero. Source snapshots, audit outcomes and print snapshots serve different purposes.

The optional template.mjs is the layout source for the provided .xlsx shell. Re-rendering it requires Node and @oai/artifact-tool (or an ARTIFACT_TOOL_URL file URL for that package's ES module); routine Windows assembly does **not** require it. Keep exported VBA as the source of truth, run the build, then run both suites in [QA](QA_CHECKLIST.md). The saved workbook and three example PDFs are suitable GitHub release assets. Do not commit real client files, backups or their reports.

## Limits
Designed for analyst-sized batches; Excel's worksheet limit, available memory, validation loops and cell formatting constrain scale. The scale study covers up to 50,000 synthetic records and four quoted issuers. CSV staging is in memory per file. A selected batch can contain independently successful and failed files; review Audit_Log. Unknown input columns are not archived in Raw_Data. The audit ledger is editable inside Excel. No claim of security certification, historical market pricing, sector adjustment or independent audit is made.
