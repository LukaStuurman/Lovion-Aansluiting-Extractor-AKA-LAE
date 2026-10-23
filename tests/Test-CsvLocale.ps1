$ErrorActionPreference = 'Stop'
. (Resolve-Path (Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1')).Path

function Assert-Test {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "TEST FAILED: $Message" }
}

$testDir = Join-Path ([System.IO.Path]::GetTempPath()) ('LovionCsvLocale-' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($testDir)
try {
    # Mock the grid-to-object conversion without touching geocoding or WFS.
    $script:CsvFixture = @(
        [pscustomobject][ordered]@{
            Station = '044.504'
            Naam = 'Made; Breda, west'
            Coordinaat = '123,45'
            Opmerking = "eerste regel`ntweede regel"
            Citaat = 'Hij zei "Hallo"'
        }
        [pscustomobject][ordered]@{
            Station = '017.525'
            Naam = 'Één aansluiting'
            Coordinaat = '987,65'
            Opmerking = 'normaal'
            Citaat = ''
        }
    )
    function Convert-DataTableRowsToObjects {
        param([System.Data.DataTable]$Table)
        return $script:CsvFixture
    }

    $dutchPath = Join-Path $testDir 'Nederlandse export.csv'
    Export-GridToCsv -Path $dutchPath
    $header = [string](Get-Content -LiteralPath $dutchPath -TotalCount 1)
    Assert-Test ($header -match '"Station";"Naam";"Coordinaat"') 'Dutch CSV uses semicolons between columns'

    $bytes = [System.IO.File]::ReadAllBytes($dutchPath)
    Assert-Test ($bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) 'CSV has UTF-8 BOM for Excel'

    $actual = @(Import-LovionCsvRows -Path $dutchPath)
    Assert-Test ($actual.Count -eq 2) 'Dutch CSV imports exactly two rows'
    Assert-Test ($actual[0].Station -ceq '044.504') 'station formatting survives a CSV roundtrip'
    Assert-Test ($actual[0].Naam -ceq 'Made; Breda, west') 'quoted separator characters survive'
    Assert-Test ($actual[0].Coordinaat -ceq '123,45') 'decimal comma string is unchanged'
    Assert-Test ($actual[0].Opmerking -ceq "eerste regel`ntweede regel") 'multiline quoted field survives'
    Assert-Test ($actual[0].Citaat -ceq 'Hij zei "Hallo"') 'escaped quotes survive'
    Assert-Test ($actual[1].Naam -ceq 'Één aansluiting') 'Unicode survives'

    # Existing exports and older station-resume reports use comma delimiters.
    $legacyPath = Join-Path $testDir 'old export.csv'
    $script:CsvFixture | Export-Csv -LiteralPath $legacyPath -Delimiter ',' -NoTypeInformation -Encoding UTF8
    $old = @(Import-LovionCsvRows -Path $legacyPath)
    Assert-Test ($old.Count -eq 2 -and $old[0].Naam -ceq 'Made; Breda, west') 'legacy comma CSV still imports'
    Assert-Test ($old[0].Station -ceq '044.504') 'legacy station code survives'

    Write-Output 'Dutch CSV export and backwards-compatible CSV import: PASS'
}
finally {
    Remove-Item -LiteralPath $testDir -Force -Recurse -ErrorAction SilentlyContinue
}
