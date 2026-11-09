[CmdletBinding()]
param([string]$OutputName = 'FinAutomate.xlsm')
$ErrorActionPreference = 'Stop'
$base = Split-Path $PSScriptRoot -Parent
if ($OutputName -ne [IO.Path]::GetFileName($OutputName) -or [IO.Path]::GetExtension($OutputName) -ne '.xlsm') { throw 'Use an .xlsm filename only.' }
$stage = Join-Path $base ("build_" + [guid]::NewGuid().ToString('N') + '.xlsm')
$destination = Join-Path $base $OutputName
$x = $null; $w = $null
try {
    $x = New-Object -ComObject Excel.Application
    $x.DisplayAlerts = $false; $x.EnableEvents = $false; $x.Visible = $false
    $w = $x.Workbooks.Open((Join-Path $base 'FinAutomate_Template.xlsx'), 0, $true)
    $project = $w.VBProject
    if (!$project -or !$project.VBComponents.Count) { throw 'VBProject access unavailable. Use manual assembly in docs/DEVELOPER_GUIDE.md.' }
    $w.SaveAs($stage, 52)
    foreach ($f in Get-ChildItem (Join-Path $base 'vba_modules') -Filter '*.bas' | Sort-Object Name) {
        $null = $project.VBComponents.Import($f.FullName)
    }
    $cls = [IO.File]::ReadAllText((Join-Path $base 'vba_modules/ThisWorkbook.cls'))
    $code = $cls.Substring($cls.IndexOf('Option Explicit'))
    $project.VBComponents.Item($w.CodeName).CodeModule.AddFromString($code)
    $prefix = "'" + $w.Name + "'!"
    Write-Host 'Initializing Home controls'
    $x.Run($prefix+'InitializeApplication')
    Write-Host 'Importing demo'
    $x.Run($prefix+'LoadDemo')
    Write-Host 'Building analysis'
    $x.Run($prefix+'ProcessTo',5)
    Write-Host 'Building valuation workspace from staged financial tables'
    $result=$x.Run($prefix+'InstallValuation')
    if ($result -ne 'OK') { throw "Valuation build: $result" }
    if ($w.Sheets('Home').Shapes.Count -ne 11 -or $w.Sheets('Dashboard').ChartObjects().Count -ne 2 -or $w.Sheets('Settings').Range('B4').Text -ne 'Current') { throw 'Assembly smoke test failed.' }
    $w.Sheets('Home').Activate()
    $w.Save(); $w.Close($false); $w=$null
    Move-Item -LiteralPath $stage -Destination $destination -Force
    Write-Host "Built $destination"
} finally {
    if ($w) { $w.Close($false) }
    if ($x) { $x.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x) }
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage }
}
