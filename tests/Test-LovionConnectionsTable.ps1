$ErrorActionPreference = 'Stop'
. (Resolve-Path (Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1')).Path

function Assert-Test {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "TEST FAILED: $Message" }
}

function New-TestSnapshot {
    param([string]$Text, [object]$Loaded = 64, [object]$Filtered = $null)
    return [pscustomobject]@{ rawText = $Text; loaded = $Loaded; filtered = $Filtered; selected = 1 }
}

$snapshot = New-TestSnapshot 'LS Aansluiting Aansluit methode Gebruiks-doel Doorlaat waarde Aantal geladen:64'
$evidence = Get-LovionConnectionsTableEvidence $snapshot
Assert-Test $evidence.Ready 'OCR spaces and punctuation in connection headers are tolerated'
Assert-Test ($evidence.Headers.Count -eq 3) 'all three connection columns are recognized'
Assert-Test (-not (Get-LovionConnectionsTableEvidence (New-TestSnapshot 'LS Stroomtransformatorgroep Master Asset ID Type Status')).Ready) 'transformer query results are rejected'
Assert-Test (-not (Get-LovionConnectionsTableEvidence (New-TestSnapshot 'Aansluitmethode Gebruiksdoel' $null $null)).Ready) 'headers without a row counter are rejected'
Assert-Test (-not (Get-LovionConnectionsTableEvidence $null).Ready) 'missing OCR is rejected'
Assert-Test (Get-LovionConnectionsTableEvidence (New-TestSnapshot 'Aansluitmethode Gebruiksdoel' 0)).Ready 'zero connections is a valid loaded table'

# Reproduce the live failure: compact adjacent digits were read as a letter.
# The counter retry must reread pixels at a larger scale, never guess 64.
$script:CounterRetryReads = @()
$script:EnlargedText = 'Aantal geladen: 64 Gefilterd: 64 Geselecteerd: 1'
function Get-LovionOcrSnapshot {
    param([string]$ImagePath, [switch]$UseLovionWindow, [switch]$FullWindow, [switch]$TableHeader, [switch]$CountersOnly, [int]$OcrScale = 1)
    $script:CounterRetryReads += [pscustomobject]@{ CounterPane = [bool]$CountersOnly; Scale = $OcrScale; Targeted = [bool]$UseLovionWindow }
    $text = if ($CountersOnly) { $script:EnlargedText } else { 'Aansluitmethode Gebruiksdoel Aantal geladen: unreadable Gefilterd: unreadable Geselecteerd: 1' }
    return [pscustomobject]@{ rawText = $text }
}
$counters = Get-LovionScreenCounters -UseLovionWindow -TableHeader
Assert-Test ($counters.loaded -eq 64 -and $counters.filtered -eq 64) 'enlarged counter retry recovers both exact 64 counts'
Assert-Test ($script:CounterRetryReads.Count -eq 2 -and $script:CounterRetryReads[1].CounterPane -and $script:CounterRetryReads[1].Scale -eq 2 -and $script:CounterRetryReads[1].Targeted) 'retry enlarges only the known Lovion counter pane'
$script:EnlargedText = 'Aantal geladen: unreadable Gefilterd: unreadable Geselecteerd: 1'
$counters = Get-LovionScreenCounters -UseLovionWindow -TableHeader
Assert-Test ($null -eq $counters.loaded -and $null -eq $counters.filtered) 'unreadable counters remain unknown rather than guessed'

function Assert-LovionAutomationNotCancelled {}
function Set-LovionWindowActive { return $true }
function Write-LovionBatchLog { param([string]$Message) }
$script:CounterReads = @()
$script:TopReady = $true
function Get-LovionScreenCounters {
    param([switch]$UseLovionWindow, [switch]$FullWindow, [switch]$TableHeader)
    $script:CounterReads += [pscustomobject]@{ Targeted = [bool]$UseLovionWindow; Full = [bool]$FullWindow; Top = [bool]$TableHeader }
    if ($FullWindow -or $script:TopReady) { return New-TestSnapshot 'Aansluitmethode Gebruiksdoel Aantal geladen: 64' }
    return New-TestSnapshot 'Aansluitmethode Gebruiksdoel' $null $null
}
$result = Wait-LovionConnectionsTable -TimeoutSeconds 1
Assert-Test ($result.loaded -eq 64) 'top-band recognition returns the 64 loaded connections'
Assert-Test ($script:CounterReads.Count -eq 1 -and $script:CounterReads[0].Targeted -and $script:CounterReads[0].Top) 'recognition reads the known Lovion table band, not the foreground window'
$script:CounterReads = @()
$script:TopReady = $false
$result = Wait-LovionConnectionsTable -TimeoutSeconds 1
Assert-Test ($result.loaded -eq 64) 'old layout counters fall back to full-window OCR'
Assert-Test ($script:CounterReads.Count -eq 2 -and $script:CounterReads[1].Targeted -and $script:CounterReads[1].Full) 'fallback still targets the known Lovion window'
Write-Output 'Lovion connection table recognition: PASS'
