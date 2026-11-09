# Capture only this disposable Excel window. No desktop or other window is recorded.
$ErrorActionPreference='Stop'
$base=Split-Path $PSScriptRoot -Parent
$temp=Join-Path ([IO.Path]::GetTempPath()) ('FinDemo_'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory $temp)
foreach($name in 'FinAutomate.xlsm','scripts','valuation','sample_data') {Copy-Item -LiteralPath (Join-Path $base $name) -Destination $temp -Recurse}
[void](New-Item -ItemType Directory (Join-Path $temp 'reports'))
$frames=Join-Path $temp 'frames';[void](New-Item -ItemType Directory $frames)
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace FinCapture -Name Native -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags); [System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out System.Drawing.Rectangle r);' -ReferencedAssemblies System.Drawing.Primitives
function Frame($index,$sheet,$cell,$zoom) {
    $s=$w.Sheets($sheet);$s.Activate();$x.Goto($s.Range($cell),$true);$x.ActiveWindow.Zoom=$zoom
    Start-Sleep -Milliseconds 350
    $rect=[Drawing.Rectangle]::Empty
    [void][FinCapture.Native]::GetWindowRect([intptr]$x.Hwnd,[ref]$rect)
    $bmp=[Drawing.Bitmap]::new(($rect.Width-$rect.X),($rect.Height-$rect.Y))
    $g=[Drawing.Graphics]::FromImage($bmp);$dc=$g.GetHdc()
    try {$ok=[FinCapture.Native]::PrintWindow([intptr]$x.Hwnd,$dc,2)} finally {$g.ReleaseHdc($dc)}
    if (!$ok) {throw 'Window recording unavailable'}
    $bmp.Save((Join-Path $frames ($index+'.png')));$g.Dispose();$bmp.Dispose()
}
$x=$null;$w=$null
try {
    $x=New-Object -ComObject Excel.Application
    $x.DisplayAlerts=$false;$x.Visible=$true;$x.EnableEvents=$false;$x.WindowState=-4143
    $x.Width=930;$x.Height=650
    $w=$x.Workbooks.Open((Join-Path $temp 'FinAutomate.xlsm'),0,$false)
    $prefix="'FinAutomate.xlsm'!"
    $x.Run($prefix+'ImportPath',(Join-Path $temp 'sample_data/company_acme.csv'))
    Frame '01' 'Home' 'A1' 82
    $x.Run($prefix+'ProcessTo',2)
    Frame '02' 'Quality_Control' 'A1' 82
    $ledger=$w.Sheets('Record_Lineage');$rid=$null
    for ($r=2;$r -le $ledger.UsedRange.Rows.Count;$r++) {if ($ledger.Cells.Item($r,8).Text -eq 'REVIEWCASE') {$rid=$ledger.Cells.Item($r,1).Value2;break}}
    $cfg=$w.Sheets('Valuation_Setup');$cfg.Range('B22')=$rid;$cfg.Range('B23')='Review';$cfg.Range('B24')='Demo analyst';$cfg.Range('B25')='Reviewed source reference in the synthetic fixture'
    $x.Run($prefix+'SubmitReview');$cfg.Range('B23')='Approve';$x.Run($prefix+'SubmitReview')
    Frame '03' 'Review_History' 'A1' 80
    $cfg.Range('B5')='ARCO';$x.Run($prefix+'RefreshValuation')
    Frame '04' 'Valuation_Setup' 'A1' 82
    Frame '05' 'LTM_Reconciliation' 'A1' 78
    Frame '06' 'Peer_Group' 'A1' 78
    Frame '07' 'DCF' 'A1' 78
    $x.Run($prefix+'TraceValuation')
    Frame '08' 'Record_Lineage' 'H1' 76
    $x.Run($prefix+'ExportValuationPDF')
    Frame '09' 'Valuation_Summary' 'A1' 72
    if (!(Get-ChildItem (Join-Path $temp 'reports') -Filter '*.pdf')) {throw 'Export did not complete'}
    Write-Host "Frame directory $frames"
} finally {
    if ($w) {$w.Close($false)}
    if ($x) {$x.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($x)}
}
