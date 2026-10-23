$ErrorActionPreference = 'Stop'
. (Resolve-Path (Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1')).Path
function Update-StatusBar {}
function Write-LovionBatchLog { param([string]$Message) }
function Assert-Test { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw "TEST FAILED: $Message" } }
$fixtureDirectory = Join-Path ([IO.Path]::GetTempPath()) ('LovionResumeTest-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixtureDirectory)
$sourcePath = Join-Path $fixtureDirectory 'failed.json'
$script:StationTestReportPath = Join-Path $fixtureDirectory 'resumed.json'
$script:StationTestRequested = @('001.001','001.002','001.003')
$records = @(
    [pscustomobject]@{Station='001.001';Transformer=1;TransformerCount=2;Status='in bedrijf';Outcome='imported';ExpectedCount=2;ImportedCount=2}
    [pscustomobject]@{Station='001.001';Transformer=2;TransformerCount=2;Status='gepland';Outcome='planned';ExpectedCount=0;ImportedCount=0}
    [pscustomobject]@{Station='001.002';Transformer=1;TransformerCount=2;Status='in bedrijf';Outcome='imported';ExpectedCount=1;ImportedCount=1}
)
$rows = @(
    [pscustomobject]@{SourceFile='001.001';'Master Asset ID'='111111111111';SourceType='LovionStationBatch';SourcePath='';FinalRdX='100';FinalRdY='200'}
    [pscustomobject]@{SourceFile='001.001';'Master Asset ID'='222222222222';SourceType='LovionStationBatch';SourcePath='';FinalRdX='110';FinalRdY='210'}
    [pscustomobject]@{SourceFile='001.002';'Master Asset ID'='333333333333';SourceType='LovionStationBatch';SourcePath='';FinalRdX='120';FinalRdY='220'}
    [pscustomobject]@{SourceFile='001.002';'Master Asset ID'='444444444444';SourceType='LovionStationBatch';SourcePath='';FinalRdX='130';FinalRdY='230'}
)
$fixture = [ordered]@{State='failed';RequestedStations=$script:StationTestRequested;Records=$records;ImportedRows=4;Error='Test interruption'}
$fixture | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $sourcePath -Encoding UTF8
$rows | Export-Csv -LiteralPath ([IO.Path]::ChangeExtension($sourcePath,'.csv')) -NoTypeInformation -Encoding UTF8
$completed = @(Restore-LovionStationTestReports -Paths @($sourcePath))
Assert-Test ($completed.Count -eq 1 -and $completed[0] -eq '001.001') 'only whole stations are retained'
Assert-Test ($script:Table.Rows.Count -eq 2 -and $script:StationTestRecords.Count -eq 2) 'partial station and tail batch are excluded'
Assert-Test ($script:Table.Rows[0]['SourceFile'] -eq '001.001') 'original station label is retained'
Assert-Test ($script:StationTestRecords[1].Outcome -eq 'planned') 'planned note survives resuming'
Write-LovionStationTestCheckpoint -State 'running'
$checkpoint = Get-Content -LiteralPath $script:StationTestReportPath -Raw | ConvertFrom-Json
$savedCsvPath = [IO.Path]::ChangeExtension($script:StationTestReportPath,'.csv')
$savedHeader = [string](Get-Content -LiteralPath $savedCsvPath -TotalCount 1)
Assert-Test ($savedHeader -match '"SourceFile";"SourceType"') 'new checkpoint CSV uses Excel-compatible semicolons'
$saved = @(Import-LovionCsvRows -Path $savedCsvPath)
Assert-Test ($checkpoint.ImportedRows -eq 2 -and $saved.Count -eq 2) 'checkpoint saves exactly the retained grid'
Assert-Test ($saved[0].FinalRdX -eq '100' -and $saved[0].SourceFile -eq '001.001') 'checkpoint preserves existing values'
$rejected = $false
try { Restore-LovionStationTestReports -Paths @($sourcePath,$sourcePath) | Out-Null } catch { $rejected = $true }
Assert-Test $rejected 'duplicate source stations are rejected'
Assert-Test ($script:Table.Rows.Count -eq 2) 'a rejected restore does not append rows'
$rows[0].SourceFile = '999.999'
$rows | Export-Csv -LiteralPath ([IO.Path]::ChangeExtension($sourcePath,'.csv')) -NoTypeInformation -Encoding UTF8
$rejected = $false
try { Restore-LovionStationTestReports -Paths @($sourcePath) | Out-Null } catch { $rejected = $true }
Assert-Test $rejected 'mismatched CSV identity is rejected'
Write-Output ('Station resume and preserved checkpoint: PASS (fixtures: ' + $fixtureDirectory + ')')
