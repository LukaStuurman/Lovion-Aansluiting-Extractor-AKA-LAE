param([string]$ImagePath)
$ErrorActionPreference = 'Stop'
. (Resolve-Path (Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1')).Path
function Assert-Test { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw "TEST FAILED: $Message" } }
function New-TestWord {
    param([string]$Text, [double]$X, [double]$Y, [double]$Width, [double]$Height = 16)
    return [pscustomobject]@{ Text=$Text; X=$X; Y=$Y; Width=$Width; Height=$Height; CenterX=$X+$Width/2; CenterY=$Y+$Height/2 }
}
$words = @(
    New-TestWord 'Master' 995 688 54 18
    New-TestWord 'Asset' 1053 688 42 18
    New-TestWord 'ID' 1100 688 16 18
    New-TestWord 'Status' 1308 688 38 18
    New-TestWord 'Aantal' 1360 688 40 18
    New-TestWord '1111111111111111111' 995 718 156
    New-TestWord 'gepland' 1308 718 45
    New-TestWord '2222222222222222222' 995 740 156
    New-TestWord 'in' 1308 740 11
    New-TestWord 'bedrijf' 1322 740 36
)
$snapshot = [pscustomobject]@{ Width=1920; Height=1080; Words=$words; Lines=@([pscustomobject]@{ Text='Master Asset ID Status Aantal'; Words=@($words[0..4]) }); rawText='' }
$first = Get-LovionTransformerMetadataFromSnapshot -Snapshot $snapshot -TransformerIndex 0
$second = Get-LovionTransformerMetadataFromSnapshot -Snapshot $snapshot -TransformerIndex 1
Assert-Test ($first.IsPlanned -and $first.Status -eq 'gepland') 'first row is classified as planned'
Assert-Test (-not $second.IsPlanned -and $second.Status -eq 'in bedrijf') 'second row is classified independently as operating'
Assert-Test ($first.AssetId -ne $second.AssetId) 'row identities are kept separate'
Write-Output 'Transformer row status: PASS'

if ($ImagePath) {
    $actual = Get-LovionOcrSnapshot -ImagePath $ImagePath -TransformerTable -OcrScale 3
    $firstActual = Get-LovionTransformerMetadataFromSnapshot -Snapshot $actual -TransformerIndex 0
    $secondActual = Get-LovionTransformerMetadataFromSnapshot -Snapshot $actual -TransformerIndex 1
    if ($null -eq $firstActual -or $null -eq $secondActual) {
        $actual.Words | Where-Object { $_.CenterY -gt 640 -and $_.CenterY -lt 790 } | Select-Object Text,X,Y,CenterX,CenterY | Format-Table -AutoSize
    }
    Assert-Test ($null -ne $firstActual -and $firstActual.IsPlanned) 'live screenshot first row is planned'
    Assert-Test ($null -ne $secondActual -and $secondActual.Status -eq 'in bedrijf') 'live screenshot second row is operating'
    Write-Output 'Live 159.528 screenshot statuses: PASS'
}

function Assert-LovionAutomationNotCancelled {}
function Set-LovionBatchStatus { param([string]$Message) }
function Write-LovionBatchLog { param([string]$Message) }
function Add-LovionStationTestRecord { param($Station,$TransformerIndex,$TransformerCount,$Status,$Outcome,$ExpectedCount,$ImportedCount) }
function Open-LovionConnectionsForStation {
    param($StationNumber,$StationIndex,$StationCount,$TransformerIndex,$KnownTransformerCount,[switch]$ReuseCurrentStationResults)
    $script:Calls += [pscustomobject]@{ Index=$TransformerIndex; Reuse=[bool]$ReuseCurrentStationResults }
    $planned = $TransformerIndex -eq $script:PlannedIndex
    return [pscustomobject]@{ InitialResultCount=2; ManualRequired=$false; SkipReason=if($planned){'gepland'}else{''}; ConnectionCount=12; TransformerStatus=if($planned){'gepland'}else{'in bedrijf'} }
}
function Invoke-LovionBatchImport { param($SourcePrefix,[switch]$SkipCountdown,[switch]$UseExactSourceName,$SourceType) $script:Imports += 1; return [pscustomobject]@{ ImportedCount=12 } }
function Close-LovionStationResultTabs { $script:Closes += 1 }
$script:PlannedIndex=0; $script:Calls=@(); $script:Imports=0; $script:Closes=0
$result = Invoke-LovionStationsImportCore -StationNumbers @('159.528')
Assert-Test ($result.ImportedCount -eq 12 -and $script:Imports -eq 1) 'only the working transformer is imported'
Assert-Test ($result.PlannedTransformers.Count -eq 1 -and $result.PlannedTransformers[0] -match '159.528.*1/2.*gepland') 'planned transformer is retained for the final message'
Assert-Test ($script:Calls.Count -eq 2 -and $script:Calls[1].Reuse) 'working transformer still reuses the same station results'
Assert-Test ($script:Closes -eq 1) 'planned transformer does not close nonexistent tabs'
$script:PlannedIndex=1; $script:Calls=@(); $script:Imports=0; $script:Closes=0
$result = Invoke-LovionStationsImportCore -StationNumbers @('159.528')
Assert-Test ($result.ImportedCount -eq 12 -and $script:Imports -eq 1 -and $script:Closes -eq 1) 'a planned second row does not block the operating first row'
Assert-Test ($result.PlannedTransformers[0] -match '2/2.*gepland') 'planned second row is retained for the final message'
$script:PlannedIndex=-1; $script:Calls=@(); $script:Imports=0; $script:Closes=0
$result = Invoke-LovionStationsImportCore -StationNumbers @('159.539')
Assert-Test ($result.ImportedCount -eq 24 -and $script:Imports -eq 2 -and $script:Closes -eq 2) 'two operating transformers are both imported'
Assert-Test ($result.PlannedTransformers.Count -eq 0) 'operating transformers are not marked as planned'
Write-Output 'Planned transformer continuation and two operating transformers: PASS'
