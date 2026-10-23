$ErrorActionPreference = 'Stop'

$sourcePath = Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1'
. (Resolve-Path -LiteralPath $sourcePath).Path

function Assert-Test {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw "TEST FAILED: $Message"
    }
}

function Assert-LovionAutomationNotCancelled {}
function Set-LovionWindowActive { return $true }
function Write-LovionBatchLog { param([string]$Message) }

# 1. The transformer row is anchored to the lower-right Master Asset ID header,
#    not to the generic connection-grid coordinates.
$wordSpecs = @(
    @{ Text = 'Master'; X = 995; Y = 688; Width = 54; Height = 18 },
    @{ Text = 'Asset'; X = 1053; Y = 688; Width = 42; Height = 18 },
    @{ Text = 'ID'; X = 1100; Y = 688; Width = 16; Height = 18 }
)
$words = @($wordSpecs | ForEach-Object {
    [pscustomobject]@{
        Text = $_.Text
        X = [double]$_.X
        Y = [double]$_.Y
        Width = [double]$_.Width
        Height = [double]$_.Height
        CenterX = [double]$_.X + ([double]$_.Width / 2.0)
        CenterY = [double]$_.Y + ([double]$_.Height / 2.0)
    }
})
$script:MockOcrSnapshot = [pscustomobject]@{
    Width = 1920
    Height = 1080
    rawText = 'Aantal geladen: 2 Gefilterd: 2 Geselecteerd: 0 Master Asset ID'
    Lines = @([pscustomobject]@{ Text = 'Master Asset ID'; Words = $words })
    Words = $words
}
function Get-LovionOcrSnapshot {
    param([string]$ImagePath, [switch]$UseLovionWindow, [switch]$FullWindow)
    return $script:MockOcrSnapshot
}
$firstPoint = Wait-LovionTransformerResultRowPoint -VisibleRowIndex 0 -TimeoutSeconds 2 -PollMilliseconds 1
$secondPoint = Wait-LovionTransformerResultRowPoint -VisibleRowIndex 1 -TimeoutSeconds 2 -PollMilliseconds 1
Assert-Test ($firstPoint.X -gt 900 -and $firstPoint.Y -gt 600) 'transformer row 1 is in the lower-right result grid'
Assert-Test ($firstPoint.Y -eq 726) 'transformer row 1 uses the expected row offset'
Assert-Test ($secondPoint.Y -eq 748) 'transformer row 2 uses the expected row pitch'
Write-Output 'Transformer OCR anchor: PASS'

# 2. Geselecteerd is in the bottom-right status bar and must be read from the
#    full window. This also verifies that the workflow waits for the selection
#    to become visible instead of failing during Citrix rendering latency.
$script:SelectedReads = 0
$script:FullWindowReads = 0
$script:TableHeaderReads = 0
function Get-LovionScreenCounters {
    param([switch]$FullWindow, [switch]$UseLovionWindow, [switch]$TableHeader)
    if ($TableHeader) {
        $script:TableHeaderReads += 1
        return [pscustomobject]@{ loaded=$null; filtered=$null; selected=$null; rawText='' }
    }
    $script:SelectedReads += 1
    if ($FullWindow) { $script:FullWindowReads += 1 }
    $selected = if ($script:SelectedReads -eq 1) { 0 } else { 1 }
    return [pscustomobject]@{
        loaded = 2
        filtered = 2
        selected = $selected
        rawText = "Aantal geladen: 2 Gefilterd: 2 Geselecteerd: $selected"
    }
}
$selected = Wait-LovionSelectedCount -TargetCount 1 -TimeoutSeconds 3
Assert-Test ($selected -eq 1) 'selected counter reaches 1'
Assert-Test ($script:SelectedReads -eq $script:FullWindowReads -and $script:TableHeaderReads -gt 0) 'bottom-right query counter is read using full-window fallback after checking the table band'
Write-Output 'Selected counter wait: PASS'

function Get-LovionScreenCounters {
    param([switch]$FullWindow, [switch]$UseLovionWindow, [switch]$TableHeader)
    Assert-Test ($UseLovionWindow -and $TableHeader -and -not $FullWindow) 'connection counter targets the named Lovion table band'
    return [pscustomobject]@{ loaded=199; filtered=199; selected=99; rawText='Aantal geladen: 199 Gefilterd: 199 Geselecteerd: 99' }
}
Assert-Test ((Get-LovionSelectedCount) -eq 99) 'last batch of 99 rows is read from the upper counter band'
Write-Output 'Upper connection counter (99 rows): PASS'

# 3. The second transformer reuses the already loaded station result page.
function Set-LovionBatchStatus { param([string]$Message) }
function Open-LovionConnectionsForStation {
    param(
        [string]$StationNumber, [int]$StationIndex, [int]$StationCount,
        [int]$TransformerIndex = 0, [Nullable[int]]$KnownTransformerCount,
        [switch]$ReuseCurrentStationResults
    )
    $script:OpenCalls += [pscustomobject]@{
        Station = $StationNumber
        TransformerIndex = $TransformerIndex
        Reuse = [bool]$ReuseCurrentStationResults
        KnownCount = if ($null -eq $KnownTransformerCount) { $null } else { [int]$KnownTransformerCount }
    }
    if ($TransformerIndex -gt 0 -and (-not $ReuseCurrentStationResults -or $KnownTransformerCount -ne 2)) {
        throw 'Transformer 2 did not reuse the current station results.'
    }
    return [pscustomobject]@{ InitialResultCount = 2; ManualRequired = $false }
}
function Invoke-LovionBatchImport {
    param([string]$SourcePrefix, [switch]$SkipCountdown, [switch]$UseExactSourceName, [string]$SourceType)
    $script:ImportCalls += 1
    return [pscustomobject]@{ ImportedCount = 1 }
}
function Close-LovionStationResultTabs { $script:CloseCalls += 1 }

$script:OpenCalls = @()
$script:ImportCalls = 0
$script:CloseCalls = 0
$result = Invoke-LovionStationsImportCore -StationNumbers @('017.545')
Assert-Test ($script:OpenCalls.Count -eq 2) 'two transformer cycles are opened for 017.545'
Assert-Test (-not $script:OpenCalls[0].Reuse) 'transformer 1 performs the initial query'
Assert-Test ($script:OpenCalls[1].Reuse) 'transformer 2 reuses the existing result page'
Assert-Test ($script:OpenCalls[1].KnownCount -eq 2) 'transformer 2 receives the known count'
Assert-Test ($script:ImportCalls -eq 2) 'both transformer cycles are imported'
Assert-Test ($script:CloseCalls -eq 2) 'each transformer cycle closes before the next one'
Write-Output 'Station result reuse: PASS'

# 4. The next station must still click EXPLORE and Start query if Citrix OCR
#    misses a correctly typed station number. A failed query itself must abort.
$script:RecognizeStationInput = $true
$script:StationQueryClicks = New-Object System.Collections.Generic.List[string]
$script:StationQueryLogs = New-Object System.Collections.Generic.List[string]
$script:StationQueryResultReady = $true
function Wait-LovionScreenPattern {
    param([string]$Pattern, [int]$TimeoutSeconds, [string]$Description)
    Assert-Test ($TimeoutSeconds -le 5) 'station input OCR retry is bounded'
    if ($script:RecognizeStationInput) { return [pscustomobject]@{ rawText='017.525' } }
    return $null
}
function Write-LovionBatchLog { param([string]$Message) $script:StationQueryLogs.Add($Message) }
function Get-LovionScreenCounters {
    param([switch]$FullWindow)
    Assert-Test $FullWindow 'the pre-query snapshot covers the full Lovion window'
    return [pscustomobject]@{ rawText='Previous station result'; loaded=1; filtered=1 }
}
function Invoke-LovionStationOcrClick {
    param($TargetText,$Description,$TimeoutSeconds,$MinXRatio,$MaxXRatio,$MinYRatio,$MaxYRatio,$FallbackRecordedX,$FallbackRecordedY,$FallbackAfterSeconds,$AfterClickMilliseconds)
    $script:StationQueryClicks.Add([string]$TargetText)
}
function Wait-LovionInitialStationQueryResults {
    param([string]$PreviousRawText,[int]$TimeoutSeconds)
    Assert-Test ($PreviousRawText -eq 'Previous station result') 'query uses a pre-click snapshot'
    if (-not $script:StationQueryResultReady) { return $null }
    return [pscustomobject]@{ rawText='New station result'; loaded=1; filtered=1 }
}
$first = Invoke-LovionStationQueryAfterInput -StationNumber '017.525'
$script:RecognizeStationInput = $false
$second = Invoke-LovionStationQueryAfterInput -StationNumber '017.526'
Assert-Test ($first.loaded -eq 1 -and $second.loaded -eq 1) 'both station queries return results'
Assert-Test (($script:StationQueryClicks.ToArray() -join '|') -eq 'EXPLORE|Start query|EXPLORE|Start query') 'second station must reopen Explore and start its own query'
Assert-Test (@($script:StationQueryLogs | Where-Object { $_ -match '017\.526.*niet door OCR bevestigd' }).Count -eq 1) 'unverified station input must be logged'
$script:StationQueryResultReady = $false
$queryFailed = $false
try { $null = Invoke-LovionStationQueryAfterInput -StationNumber '017.527' }
catch { $queryFailed = $_.Exception.Message -like '*konden na Start query niet betrouwbaar worden gelezen*' }
Assert-Test $queryFailed 'an unverified query result must not be imported'
Write-Output 'Second station OCR continuation: PASS'
Write-Output 'All Lovion station workflow tests: PASS'
