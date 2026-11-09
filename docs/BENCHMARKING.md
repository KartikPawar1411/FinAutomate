# Measure the workflow

Run from a local Windows folder with Excel Desktop. The integration suite creates its own workbook copy and runs actual VBA in a separate Excel instance.

```powershell
.\scripts\test_excel.ps1
```

The Stopwatch starts immediately before LoadDemo reads six files and stops immediately after the second successful PDF export. It includes cleaning, statement validation, all ratios, formatting, dashboard/chart refresh, individual and comparison report generation and PDF I/O. It excludes Excel startup, copying fixtures, screenshots and subsequent fault-injection tests. Counts come from the ledgers and files present. One issuer can have several company-years; do not describe 21 records as 21 companies.

See [benchmark.json](benchmark.json) for the published single-run time, build and counts. For a repeatable performance study, record machine, Excel build, storage and power mode; close other workbooks, run at least five repetitions and report each run, median and range. Do not extrapolate the six-file fixture to production volume. Use -PublishEvidence only if intentionally replacing the shipped benchmark, PDFs and print views.

**Manual baseline:** none is supplied. To measure one, use the same inputs, decisions and outputs; start before opening the first source and stop after saving both comparable PDFs. Record participant, date, scope, raw times and correctness checks. Only after measuring both equivalent workflows compute seconds saved = manual − automatic and percentage saved = (manual − automatic) / manual. Do not present an estimate as a measurement.
