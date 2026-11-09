# Résumé and interview notes

**FinAutomate — Investment Banking Excel Automation & Reporting**

Excel Desktop · VBA · PowerShell / Excel COM · Python regression testing

- Built an Excel/VBA trading-comps and five-year DCF platform with calendar-aligned LTM financials, four peer valuation methods, diluted EV-to-equity bridges and a WACC sensitivity analysis.
- Implemented publication-date filters, SHA-256 record lineage and a controlled filing review workflow. Verified 340 annual ratio results, 30 financial-engine checks and 17 native Excel valuation and integration checks.
- Benchmarked 50,000 synthetic financial records on Windows. The pipeline accepted 49,497 records and quarantined 503. Excel refreshed the valuation workbook in 5.32 seconds and exported its PDF in 0.63 seconds in that trial.

**One-minute demo:** Show ACME FY2023 interest coverage and trace 14.00x to EBIT of 1,330 and interest expense of 95. Switch to EnergyFlow and show the thousands-to-millions conversion without FX translation. Open the DataBridge rejection and explain why excluding FY2023 is the right analytical result. Finish on the two PDF exports and the measured benchmark.

**Interview tradeoff:** A source can match on sales and net income but differ on EV or debt. First compare all 35 canonical business fields; then quarantine differing versions with the same issuer/year. Removing both from downstream analysis until resolved sacrifices coverage to avoid quietly choosing a restatement.

**Evidence:** [published measurement](benchmark.json) · [two example reports](../reports/) · [test procedures and observations](QA_CHECKLIST.md). The demo is synthetic and the timing is a single machine-specific observation. Do not claim hours saved, production deployment or client use without supporting evidence.
