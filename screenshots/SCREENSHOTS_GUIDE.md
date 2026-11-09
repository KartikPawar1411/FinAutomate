# Portfolio images

Home.png and Dashboard.png are rendered from actual Excel worksheet print views. ACME_Annual_Review.png and Peer_Comparison.png are first-page previews of the two native PDF exports. They contain synthetic company data.

To regenerate: run scripts/test_excel.ps1 -PublishEvidence from PowerShell with Excel Desktop. Then use Poppler, e.g. `pdftoppm -f 1 -l 1 -scale-to 1250 -singlefile -png screenshots/Dashboard.pdf screenshots/Dashboard`. Repeat for Home and the two reports under reports/.

For a live UI capture, open the assembled workbook at 85–100% zoom, select ACME on Dashboard, and capture the title, fiscal year, controls, KPI row and both charts. A second useful image is Clean_Data filtered to Rejected with Source_File and Cleaning_Notes visible. Do not include unrelated workbooks, client names or file paths.
