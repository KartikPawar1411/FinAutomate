[CmdletBinding()]
param([switch]$PublishEvidence)
$ErrorActionPreference = 'Stop'
$base = Split-Path $PSScriptRoot -Parent
$work = Join-Path ([IO.Path]::GetTempPath()) ('FinAutomateQA_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $work -ItemType Directory)
Copy-Item -LiteralPath (Join-Path $base 'FinAutomate.xlsm') -Destination $work
Copy-Item -LiteralPath (Join-Path $base 'sample_data') -Destination $work -Recurse
$path = Join-Path $work 'FinAutomate.xlsm'
function Assert($ok, $message) { if (!$ok) { throw $message }; Write-Host "PASS $message" }
function RowCount($sheet) { return $sheet.Range('AJ1048576').End(-4162).Row - 1 }
function RunMacro($method) { $x.Run($prefix+$method) }
function TryBad($file) {
    $old = RowCount $raw
    $caught = [string]$x.Run($prefix+'TryImportPath',$file)
    Assert ($caught -and (RowCount $raw) -eq $old) ("Reject without append: " + [IO.Path]::GetFileName($file))
}
function Snapshot($sheet,$bounds,$name) {
    $sheet.PageSetup.PrintArea=$bounds
    $sheet.PageSetup.Orientation=2; $sheet.PageSetup.Zoom=$false
    $sheet.PageSetup.FitToPagesWide=1; $sheet.PageSetup.FitToPagesTall=1
    $sheet.ExportAsFixedFormat(0,(Join-Path $work ($name+'.pdf')))
}
$x=$null; $w=$null
try {
    $x=New-Object -ComObject Excel.Application
    $x.Visible=$false; $x.DisplayAlerts=$false; $x.EnableEvents=$false
    $w=$x.Workbooks.Open($path,0,$false); $prefix="'"+$w.Name+"'!"
    $raw=$w.Sheets('Raw_Data'); $fs=$w.Sheets('Financial_Statements'); $dash=$w.Sheets('Dashboard')
    $panel=$w.Sheets('Home'); $cfg=$w.Sheets('Settings'); $report=$w.Sheets('Report')
    $clock=[Diagnostics.Stopwatch]::StartNew()
    RunMacro 'LoadDemo'
    $x.Run($prefix+'ProcessTo',5)
    $p1=RunMacro 'ExportCurrentPDF'
    $panel.Range('B4')='Comparison'; RunMacro 'BuildReport'; $p2=RunMacro 'ExportCurrentPDF'
    $clock.Stop()
    $measure=[ordered]@{excel_version=$x.Version;excel_build=$x.Build;scope='Load six CSVs, clean, validate, compute, format, refresh dashboard, generate two report types and export both PDFs; excludes Excel startup';automated_seconds=[math]::Round($clock.Elapsed.TotalSeconds,3);source_files=6;source_rows=(RowCount $raw);accepted_rows=(RowCount $fs);rejected_rows=[int]$cfg.Range('B5').Value2;pdfs_generated=@(Get-ChildItem (Join-Path $work 'reports') -Filter '*.pdf').Count;manual_seconds=$null}
    Assert ($measure.source_rows -eq 21 -and $measure.accepted_rows -eq 20 -and $measure.rejected_rows -eq 1) 'Demo reconciles 21 = 20 accepted + 1 rejected'
    Assert ($measure.pdfs_generated -eq 2 -and (Get-Item $p1).Length -gt 1000 -and (Get-Item $p2).Length -gt 1000) 'Two native PDF exports'
    Assert ($panel.Shapes.Count -eq 11 -and $dash.ChartObjects().Count -eq 2) 'Home buttons and native charts'
    $x.EnableEvents=$true
    $dash.Range('B4')='ENFL'
    Assert ($dash.Range('C4').Text -eq 'EnergyFlow Ltd' -and $dash.Range('A7').Value2 -eq 3500 -and $dash.Range('F4').Text -eq 'EUR millions') 'Selector event switches issuer, currency and normalized revenue'
    $dash.Range('B4')='ACME'
    $x.EnableEvents=$false
    if ($PublishEvidence) {
        Snapshot $panel 'A1:H18' 'Home'
        Snapshot $dash 'A1:H38' 'Dashboard'
    }
    $hash=(Get-FileHash (Join-Path $work 'sample_data/company_acme.csv')).Hash
    $x.Run($prefix+'ImportPath',(Join-Path $work 'sample_data/company_acme.csv'))
    $x.Run($prefix+'ProcessTo',2)
    Assert ((RowCount $fs) -eq 20 -and (RowCount $raw) -eq 25) 'Repeated import leaves financial results unchanged'
    $x.Run($prefix+'ImportPath',(Join-Path $work 'sample_data/error_test_data.csv'))
    $x.Run($prefix+'ProcessTo',3)
    Assert ($cfg.Range('B5').Value2 -eq 7 -and (RowCount $fs) -eq 23) 'Bad dates, absent IDs, text and negative sales quarantined; mixed units accepted after normalization'
    $h=(Get-Content (Join-Path $work 'sample_data/company_acme.csv'))[0]
    $line=(Get-Content (Join-Path $work 'sample_data/company_acme.csv'))[1]
    $conflict=Join-Path $work 'restatement.csv'
    @($h,($line -replace '23900,','23901,')) | Set-Content -LiteralPath $conflict -Encoding utf8
    $x.Run($prefix+'ImportPath',$conflict); $x.Run($prefix+'ProcessTo',2)
    Assert ((RowCount $fs) -eq 22 -and $cfg.Range('B5').Value2 -eq 9) 'Restatement conflict blocks both distinct versions'
    $bad=Join-Path $work 'bad.csv'
    @('Company_ID,Revenue','CASE,100') | Set-Content -LiteralPath $bad
    TryBad $bad
    @($h,$line,'incomplete,row') | Set-Content -LiteralPath $bad
    TryBad $bad
    TryBad (Join-Path $work 'sample_data/empty_file.csv')
    TryBad (Join-Path $work 'sample_data/zero_bytes.csv')
    TryBad (Join-Path $work 'absent.csv')
    $other=Join-Path $work 'unsupported.txt'; 'test' | Set-Content $other; TryBad $other
    $inputBook=$x.Workbooks.Add(); $inputSheet=$inputBook.Worksheets.Item(1)
    $head=@('Revenue','Company_Name','Company_ID','Period_End_Date','Units','Reporting_Period','Currency')
    $vals=@(1200000,'XLSX example','XLS',([datetime]'2024-12-31').ToOADate(),'Thousands','FY2024','USD')
    for ($i=0;$i -lt $head.Count;$i++) { $inputSheet.Cells.Item(1,$i+1)=$head[$i]; $inputSheet.Cells.Item(2,$i+1)=$vals[$i] }
    $inputSheet.Cells.Item(2,4).NumberFormat='yyyy-mm-dd'
    $xlsx=Join-Path $work 'typed.xlsx'; $inputBook.SaveAs($xlsx,51)
    TryBad $xlsx
    $inputBook.Close($false)
    $xhash=(Get-FileHash $xlsx).Hash
    $x.Run($prefix+'ImportPath',$xlsx); $x.Run($prefix+'ProcessTo',3)
    $last=$fs.Range('A1048576').End(-4162).Row
    Assert ($fs.Cells.Item($last,1).Text -eq 'XLS' -and $fs.Cells.Item($last,7).Value2 -eq 1200 -and $fs.Cells.Item($last,4).Text -eq '2024-12-31') 'Reordered XLSX headers, typed date and thousand-unit conversion'
    Assert ($xhash -eq (Get-FileHash $xlsx).Hash -and $hash -eq (Get-FileHash (Join-Path $work 'sample_data/company_acme.csv')).Hash) 'Source files remain byte-identical'
    $failDir=Join-Path $work 'blocked_export'
    [void](New-Item -ItemType Directory -Path $failDir)
    $w.SaveAs((Join-Path $failDir 'FinAutomate.xlsm'),52)
    $blockedPath=Join-Path $failDir 'reports'
    'File occupies export folder path' | Set-Content -LiteralPath $blockedPath
    $exportError=$x.Run($prefix+'TryExportPDF')
    Assert (![string]::IsNullOrEmpty($exportError) -and (RowCount $fs) -eq 23) 'Export failure reported without changing statement data'
    # Leave the disposable failure fixture for investigation.
    if ($PublishEvidence) {
        Copy-Item -LiteralPath $p1 -Destination (Join-Path $base 'reports/ACME_Annual_Review.pdf') -Force
        Copy-Item -LiteralPath $p2 -Destination (Join-Path $base 'reports/Peer_Comparison.pdf') -Force
        $measure | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $base 'docs/benchmark.json') -Encoding utf8
        Copy-Item -LiteralPath (Join-Path $work 'Home.pdf') -Destination (Join-Path $base 'screenshots/Home.pdf') -Force
        Copy-Item -LiteralPath (Join-Path $work 'Dashboard.pdf') -Destination (Join-Path $base 'screenshots/Dashboard.pdf') -Force
    }
    $measure | ConvertTo-Json
} finally {
    if ($w) { $w.Close($false) }
    if ($x) { $x.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x) }
    Write-Host "Disposable QA files: $work"
}
