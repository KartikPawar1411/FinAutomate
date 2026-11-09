[CmdletBinding()]
param([int[]]$Sizes=@(1000,10000,50000))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$results=@()
Add-Type -Namespace FinancialQA -Name Window -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint process);'
foreach ($size in $Sizes) {
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('ValuationScale_'+[guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory $dir)
    $path=Join-Path $dir 'FinAutomate.xlsm'
    Copy-Item -LiteralPath (Join-Path $root 'FinAutomate.xlsm') -Destination $path
    $metrics=(python (Join-Path $PSScriptRoot 'stage_scale.py') $size (Join-Path $dir 'valuation')) | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) {throw 'Data preparation failed'}
    $x=$null;$w=$null
    try {
        $x=New-Object -ComObject Excel.Application
        $x.DisplayAlerts=$false;$x.ScreenUpdating=$false;$x.EnableEvents=$false
        $w=$x.Workbooks.Open($path,0,$false)
        $sw=[Diagnostics.Stopwatch]::StartNew()
        $status=$x.Run("'FinAutomate.xlsm'!InstallValuation")
        if ($status -ne 'OK') {throw $status}
        $sw.Stop();$refresh=$sw.Elapsed.TotalSeconds
        $sheet=$w.Sheets('Valuation_Summary')
        $sw.Restart();$sheet.ExportAsFixedFormat(0,(Join-Path $dir 'valuation.pdf'));$sw.Stop()
        $export=$sw.Elapsed.TotalSeconds
        $w.Save()
        [uint32]$excelId=0
        [void][FinancialQA.Window]::GetWindowThreadProcessId([intptr]$x.Hwnd,[ref]$excelId)
        $metrics | Add-Member -NotePropertyMembers @{excel_peak_MiB=[math]::Round((Get-Process -Id $excelId).PeakWorkingSet64/1MB,1);refresh_workbook_seconds=$refresh;pdf_seconds=$export;workbook_bytes=(Get-Item $path).Length;excel_build=$x.Build}
        $results+=$metrics
        $results | ConvertTo-Json -Depth 4 -AsArray | Set-Content (Join-Path $root 'docs/scale_excel.json')
        Write-Host "Completed $size records; Excel refresh $refresh s, PDF $export s."
    } finally {
        if ($w) {$w.Close($false)}
        if ($x) {$x.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
    }
}
