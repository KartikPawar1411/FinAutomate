# Troubleshooting

| Symptom | What to check |
|---|---|
| Buttons do nothing | Use the assembled .xlsm in Excel Desktop for Windows. Verify you are not in Protected View or using the macro-free template. Review the source and follow your organization's normal trusted-file policy. Do not enable all macros globally. |
| Build reports VBProject unavailable | Running the app does not require project access. Build manually from the .xlsx using the [developer guide](DEVELOPER_GUIDE.md); the build script does not change Trust Center settings. |
| Existing output workbook locked | Close it or build with a different -OutputName. The new workbook is staged before the old release is replaced. |
| Nothing imported | Check Audit_Log. Use UTF-8 CSV or XLSX, correct row-1 headers and rectangular rows. Close an XLSX already open in the active instance. A header-only or zero-byte file has no records. |
| Numbers or dates quarantined | Use ISO or US slash dates, point decimals, properly quoted comma-grouped values and a recognized unit. Currency symbols and text such as TBD are not financial amounts. Filter Clean_Data and look at the source record. |
| Latest year missing | The dashboard uses latest **accepted** filings. Check Rejected/duplicate records before assuming a stale chart. Run the full workflow after imports; model sheets are generated outputs. |
| N/A or N/M | Consult [FORMULAS.md](FORMULAS.md). A missing inventory is not zero. YoY requires a compatible prior year; negative earnings do not produce a positive P/E. |
| Several filings for one year are missing | Conflicting versions are withheld. Check all business fields, not just revenue and earnings. Do not approve one by changing a generated status. Rebuild from approved upstream records. |
| Missing sample directory | Extract sample_data beside the workbook. Loading a demo reads all six files first. |
| Export fails | Check write access and free space and that reports is a directory. Use a local writable project folder. Review the audit message; the source ledger is retained. |
| Manual build lacks selector events | The event code belongs in the existing ThisWorkbook component, not a separately imported class. Debug → Compile VBAProject and reopen with macros permitted. |
| Tests do not match after personal edits | Run the release tests on an unmodified demo build. The oracle deliberately expects the six published fixture files. |
