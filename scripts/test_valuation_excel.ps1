[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$checks=[Collections.Generic.List[object]]::new()
function Verify($name,$expected,$actual) {
    if ($expected -is [double]) { $ok=[Math]::Abs($expected-[double]$actual) -lt 0.000001 }
    else { $ok=($expected -eq $actual) }
    $checks.Add([ordered]@{test=$name;expected=$expected;actual=$actual;passed=$ok})
    if (!$ok) { Write-Warning "$name expected $expected, got $actual" }
}
$x=$null;$w=$null
try {
    $x=New-Object -ComObject Excel.Application
    $x.DisplayAlerts=$false;$x.EnableEvents=$false
    $w=$x.Workbooks.Open((Join-Path $root 'FinAutomate.xlsm'),0,$true)
    $prefix="'FinAutomate.xlsm'!"
    $cfg=$w.Sheets('Valuation_Setup');$dcf=$w.Sheets('DCF');$summary=$w.Sheets('Valuation_Summary')
    $trading=$w.Sheets('Trading_Data');$peers=$w.Sheets('Peer_Group');$comps=$w.Sheets('Trading_Comps')
    $x.Calculate()
    $wacc=(.04+1.1*.055)*.75+.055*(1-.25)*.25
    $rev=7000.;$pv=0.;$flow=0
    foreach ($year in 1..5) {
        $prev=$rev;$rev*=1.06
        $flow=$rev*.18*.75+$rev*.04-$rev*.065-($rev-$prev)*.12
        $pv+=$flow/[math]::Pow(1+$wacc,$year)
    }
    $ent=$pv+$flow*1.025/($wacc-.025)/[math]::Pow(1+$wacc,5)
    Verify 'CAPM WACC' $wacc $dcf.Range('B18').Value2
    Verify 'Growth DCF EV' $ent $dcf.Range('B21').Value2
    Verify 'DCF implied share price' (($ent-2685)/500) $dcf.Range('B23').Value2
    Verify 'Sensitivity center matches independent DCF' (($ent-2685)/500) $dcf.Range('D30').Value2
    Verify 'Price decreases with higher WACC' $true ($dcf.Range('D29').Value2 -gt $dcf.Range('D31').Value2)
    Verify 'Quoted diluted EV bridge' 35185.0 $trading.Range('T2').Value2
    Verify 'Target excluded from peer count' 3 $comps.Range('B2').Value2
    $pe=@((71.5*390/(5460*.14625)),(58.5*625/(8750*.10725)),(84.5*410/(5740*.169))) | Sort-Object
    Verify 'Peer median P/E' ([double]$pe[1]) $comps.Range('E5').Value2
    Verify 'Implied equity from peer P/E' ([double]$pe[1]*910) $comps.Range('F19').Value2
    Verify 'Football field chart present' 1 $summary.ChartObjects().Count
    $summary.ExportAsFixedFormat(0,(Join-Path $root 'reports/Valuation_Case.pdf'))
    $peers.Range('B3')='No';$x.Calculate()
    Verify 'Excluding peer changes statistics' 2 $comps.Range('B2').Value2
    $cfg.Range('B19')='Exit';$x.Calculate()
    Verify 'Exit-multiple terminal EV' ($pv+$rev*.22*12/[math]::Pow(1+$wacc,5)) $dcf.Range('B21').Value2
    $cfg.Range('B19')='Growth';$cfg.Range('B17')=.15;$x.Calculate()
    Verify 'Terminal growth above WACC' 'N/M' $dcf.Range('B21').Text
    $ratio=$w.Sheets('Financial_Ratios').Range('K2')
    $formula=$ratio.Formula
    $x.Run($prefix+'ApplyIBFormatting')
    Verify 'Formatting preserves formula' $formula $ratio.Formula
    $tmp=Join-Path ([IO.Path]::GetTempPath()) ('FinValTest_'+[guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory $tmp)
    $csv=Join-Path $tmp 'literal.csv'
    @('Company_ID,Company_Name,Reporting_Period,Period_End_Date,Currency,Units,Revenue','LIT,=1+2,FY2023,2023-12-31,USD,Millions,100') | Set-Content $csv
    $x.Run($prefix+'ImportPath',$csv)
    $raw=$w.Sheets('Raw_Data');$row=$raw.Range('A1048576').End(-4162).Row
    Verify 'Formula-like CSV field stays literal' $false $raw.Cells.Item($row,2).HasFormula
    Verify 'Annual source hash matches input bytes' ((Get-FileHash $csv).Hash.ToLower()) $raw.Cells.Item($row,42).Value2
    $enc=$x.Workbooks.Add();$enc.Worksheets.Item(1).Range('A1')='Revenue';$encPath=Join-Path $tmp 'protected.xlsx'
    $enc.SaveAs($encPath,51,'qa-password');$enc.Close($false)
    $errorText=$x.Run($prefix+'TryImportPath',$encPath)
    Verify 'Encrypted XLSX rejected without append' $true ([bool]$errorText -and $raw.Range('A1048576').End(-4162).Row -eq $row)
    $checks | ConvertTo-Json -Depth 3 | Set-Content (Join-Path $root 'docs/valuation_excel_tests.json')
    if ($checks.Where({!$_.passed}).Count) { throw 'Native valuation assertions failed' }
    Write-Host "$($checks.Count) native assertions passed"
} finally {
    if ($w) {$w.Close($false)}
    if ($x) {$x.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
}
