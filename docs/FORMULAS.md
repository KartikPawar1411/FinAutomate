# Metric dictionary

Amounts are normalized to millions of filing currency, shares to millions of shares. No FX or EV bridge is implied. Financial_Ratios contains native Excel formulas against the sorted accepted annual statements. The original stated market cap and enterprise value are used; no unsupported estimates are synthesized.

| Metric | Numerator / denominator or calculation |
|---|---|
| Gross margin | Gross profit / revenue |
| EBITDA margin | EBITDA / revenue |
| EBIT margin | EBIT / revenue |
| Net margin | Net income / revenue |
| ROE | Net income / mean of opening and closing equity; both balances must be positive |
| ROIC | EBIT × (1 − income tax / EBT) / opening invested capital. Opening capital = prior total debt + equity − cash equivalents. Requires positive EBT, positive opening capital and a filing-derived effective tax rate from 0% to 100%. Effective rather than marginal tax is a modeling limitation. |
| Current ratio | Current assets / current liabilities |
| Quick ratio | (Current assets − inventory) / current liabilities; inventory must be reported. This proxy does not separately remove prepayments. |
| Debt / equity | Total debt / total equity |
| Debt / EBITDA | Total debt / EBITDA |
| Interest coverage | EBIT / interest expense |
| Revenue YoY | Current revenue / prior revenue − 1 |
| EBITDA YoY | Current EBITDA / prior EBITDA − 1 |
| Net income YoY | Current net income / prior net income − 1 |
| EV / revenue | Reported enterprise value / revenue |
| EV / EBITDA | Reported enterprise value / EBITDA |
| P/E | Reported market capitalization / net income; no per-share rounding or inferred share count |

**Missing vs not meaningful.** `N/A` means an input or comparable prior fiscal year is unavailable. `N/M` means a required denominator is zero/negative, or an explicit economic guard fails: negative EV, market cap or debt as a multiple numerator, nonpositive opening/closing equity for ROE, or invalid tax base for ROIC. Negative earnings margins and negative interest coverage over positive interest remain numeric. Positive-to-negative earnings growth is shown; recovery from a negative base is N/M. These are text labels, not Excel errors.

**Periods.** Opening balances and growth require the same issuer, same currency and prior anniversary date (one-year date addition in either direction accommodates leap days). First available year, a missing year and changed reporting currency do not yield a prior-period metric. Amount trends may show observations separated by missing fiscal years; year labels identify the observations.

**Reconciliation.** Test gross profit = revenue − COGS, total assets = total liabilities + total equity, and net income = EBT − income tax when all relevant figures exist. Tolerance in normalized millions is max(0.000001, absolute left-hand side × 0.0001). Other statement identities and source EV are not certified by these checks.

**Worked check.** ACME FY2023: 1,330 / 95 = 14.00x coverage; 1,560 / 7,000 = 22.29% EBITDA margin; 926.25 / ((6,500 + 7,000)/2) = 13.72% ROE. FY2023 revenue growth = 7,000 / 6,200 − 1 = 12.90%. DataBridge FY2021 revenue is zero: its margin is N/M, not zero. TechVista's blank inventory makes its quick ratio N/A.

Methodology reference: [Aswath Damodaran, NYU Stern: financial measures and ratios](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/definitions.html). Precise denominator timing and exclusion rules above describe this implementation.
