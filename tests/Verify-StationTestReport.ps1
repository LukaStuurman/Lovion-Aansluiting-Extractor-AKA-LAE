param(
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [Parameter(Mandatory=$true)][string]$StationListPath,
    [switch]$AllowPartial
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
function Assert-Audit { param([bool]$Condition,[string]$Message) if(-not $Condition){throw "AUDIT FAILED: $Message"} }
$expected=@([System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $StationListPath).Path) -split '[,;\s]+' | Where-Object {$_} | Select-Object -Unique)
$report=Get-Content -LiteralPath $ReportPath -Raw|ConvertFrom-Json
Assert-Audit ([string]::IsNullOrWhiteSpace($report.Error)) $report.Error
Assert-Audit ($report.RequestedStations.Count -eq $expected.Count -and @($expected | Where-Object {$_ -notin $report.RequestedStations}).Count -eq 0) 'requested station list differs from the supplied list'
if(-not $AllowPartial){Assert-Audit ($report.State -eq 'complete') 'the live station run is not complete'}
$records=@($report.Records)
$csvPath=[System.IO.Path]::ChangeExtension((Resolve-Path -LiteralPath $ReportPath).Path,'.csv')
$rows=if(Test-Path -LiteralPath $csvPath){@(Import-Csv -LiteralPath $csvPath)}else{@()}
$sum=0
$offset=0
foreach($record in $records){
    Assert-Audit ($record.Station -in $expected) 'unexpected station was imported'
    Assert-Audit ($record.ExpectedCount -eq $record.ImportedCount) ("counter/import mismatch for {0} transformer {1}" -f $record.Station,$record.Transformer)
    if($record.Outcome -eq 'planned'){
        Assert-Audit ($record.Status -eq 'gepland' -and $record.ImportedCount -eq 0) 'planned transformer was imported or not correctly classified'
        continue
    }
    Assert-Audit ($record.Outcome -eq 'imported' -and $record.Status -eq 'in bedrijf') 'operating transformer outcome/status is unverified'
    $count=[int]$record.ImportedCount
    if($count -gt 0){
        $batch=@($rows|Select-Object -Skip $offset -First $count)
        Assert-Audit ($batch.Count -eq $count) 'CSV is missing imported connection rows'
        Assert-Audit (@($batch|Where-Object {$_.SourceFile -ne $record.Station}).Count -eq 0) 'CSV station identity differs from the import record'
        $ids=@($batch|ForEach-Object {$_.'Master Asset ID'})
        Assert-Audit (@($ids|Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count -eq 0) 'imported rows are missing Master Asset ID'
        Assert-Audit (@($ids|Select-Object -Unique).Count -eq $count) 'duplicate connection IDs within one transformer import'
    }
    $sum+=$count
    $offset+=$count
}
Assert-Audit ($sum -eq $report.ImportedRows -and $rows.Count -eq $sum) 'report total and saved CSV row count differ'
$completed=@()
$twoOperating=@()
foreach($station in $expected){
    $stationRecords=@($records|Where-Object {$_.Station -eq $station})
    if($stationRecords.Count -eq 0){if($AllowPartial){continue};throw "AUDIT FAILED: station $station was not tested"}
    $trafoCount=[int]($stationRecords|Measure-Object TransformerCount -Maximum).Maximum
    $indexes=@($stationRecords|Select-Object -ExpandProperty Transformer -Unique)
    $covered=$indexes.Count -eq $trafoCount -and @((1..$trafoCount)|Where-Object {$_ -notin $indexes}).Count -eq 0
    if(-not $covered){if($AllowPartial){continue};throw "AUDIT FAILED: station $station is missing a transformer cycle"}
    Assert-Audit ($stationRecords.Count -eq $trafoCount) 'duplicate transformer records'
    $completed+=$station
    if(@($stationRecords|Where-Object {$_.Outcome -eq 'imported'}).Count -ge 2){$twoOperating+=$station}
}
if(-not $AllowPartial){Assert-Audit ($completed.Count -eq $expected.Count) 'not all stations were fully tested'}
[pscustomobject]@{
    Complete=($report.State -eq 'complete' -and $completed.Count -eq $expected.Count)
    StationsVerified=$completed.Count
    StationsRequested=$expected.Count
    ImportedRows=$sum
    PlannedTransformers=@($records|Where-Object {$_.Outcome -eq 'planned'}).Count
    StationsWithMultipleOperatingTransformers=$twoOperating
    SavedCsv=$csvPath
}|ConvertTo-Json -Depth 4
