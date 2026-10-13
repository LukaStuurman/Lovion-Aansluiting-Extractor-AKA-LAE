param([string]$ImagePath)
$ErrorActionPreference='Stop'
. (Resolve-Path (Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1')).Path
if([string]::IsNullOrWhiteSpace($ImagePath)){
    Write-Output 'OCR resource stress test requires a local screenshot via -ImagePath.'
    exit 0
}
for($i=0;$i -lt 3;$i++){$null=Get-LovionOcrSnapshot -ImagePath $ImagePath -FullWindow}
$engine=$script:LovionOcrEngine
if($null -eq $engine){throw 'TEST FAILED: OCR engine was not created'}
[System.GC]::Collect();[System.GC]::WaitForPendingFinalizers();[System.GC]::Collect()
$process=[System.Diagnostics.Process]::GetCurrentProcess();$process.Refresh();$baseline=$process.PrivateMemorySize64
for($i=0;$i -lt 80;$i++){
    $snapshot=Get-LovionOcrSnapshot -ImagePath $ImagePath -FullWindow
    if($null -eq $snapshot){throw "TEST FAILED: OCR snapshot $i failed"}
    if(-not [object]::ReferenceEquals($engine,$script:LovionOcrEngine)){throw 'TEST FAILED: OCR engine was recreated'}
}
[System.GC]::Collect();[System.GC]::WaitForPendingFinalizers();[System.GC]::Collect()
$process.Refresh();$growth=$process.PrivateMemorySize64-$baseline
if($growth -gt 256MB){throw ('TEST FAILED: native OCR memory grew by {0:N1} MB' -f ($growth/1MB))}
Write-Output ('OCR resource lifetime: PASS; 80 repeated full-window images, one engine, growth {0:N1} MB.' -f ($growth/1MB))
