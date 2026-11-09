# Verification

Run both commands from the extracted project root:

```powershell
python .\scripts\test_workbook.py
.\scripts\test_excel.ps1
```

The first test uses only Python 3's standard library and the **saved, assembled** demo workbook. It parses the independent source CSVs, derives all 17 metrics, reads native Excel formula caches, checks each number or N/A/N/M label by issuer/year, and scans all worksheets for native cell errors. It fails on missing/extra filings. It is not a VBA compiler or a replacement for the second suite.

The PowerShell suite reopens a disposable copy in Windows Excel, runs the VBA engine, changes the selector with events enabled and exports real PDFs. Tests leave a randomly named FinAutomateQA folder in the user's temporary directory for diagnosis; its path is printed. The original application and source files are not modified. With -PublishEvidence the example PDFs, print-view PDFs and benchmark JSON in the project are refreshed.

| Release check | Observed result |
|---|---|
| Native build | Workbook initialized, 11 control shapes, two dashboard charts |
| Demo ledger | 21 source records; 20 accepted, one DataBridge year rejected |
| Model verification | 340 / 340 saved Excel results matched; zero native formula errors |
| Selector event | ACME → EnergyFlow changed issuer, currency and latest revenue to EUR 3,500m |
| Duplicate and conflict tests | Reimporting ACME did not inflate results; two distinct versions of one FY were both withheld |
| Error fixture | Invalid dates, absent identifiers, text and negative revenue rejected; equivalent units standardized |
| File validation | Unsupported, missing, header-only, zero-byte, malformed-width, open input workbook and missing required header all rejected without appending partial records |
| XLSX | Reordered headers, numeric date and scaled amount imported correctly; source hashes unchanged |
| Report generation | Two nonempty PDFs; one-page layout for each verified on the shipped demo |
| Export failure | A file blocking the export directory produced an error and left statement data intact |

## Manual release checklist
- Open **a disposable copy** from an extracted folder. Verify controls execute under normal per-file macro policy. With macros unavailable the populated sheets should still be viewable.
- Use the actual import picker: select a valid and invalid file together; inspect the per-file outcomes. Cancel the picker and confirm the raw ledger is unchanged.
- Change the issuer through the dropdown; confirm both chart titles/series and the report currency. Inspect an unavailable margin and confirm no plotted zero is substituted for missing data.
- Test a locked network folder and restricted PDF destination under the target user's permissions. Do not loosen organizational security policies. Check Home and Audit_Log for the underlying message.
- On a workbook copy, confirm Reset prompts; Cancel preserves raw data. Accept: verify the backup opens and source data is recoverable. Rename a demo file before reset to test a missing input; restore it afterward.
- Confirm blue input numbers, green cross-sheet formulas, black local formulas on the financial model. Verify formula text before/after formatting is identical.
- Export one issuer and the combined comparison, check every page at the target printer settings, reopen the saved workbook, and confirm selector and buttons still work.
