$ErrorActionPreference = 'Stop'

$sourcePath = Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1'
. (Resolve-Path -LiteralPath $sourcePath).Path

$exportPath = Join-Path $env:TEMP ('lovion-macro-test-' + [guid]::NewGuid().ToString('N') + '.xml')
try {
    $xmlText = Get-LovionMacroXml
    [xml]$macro = $xmlText
    if ($macro.DocumentElement.Name -ne 'Reports') {
        throw 'XML-root is geen Reports.'
    }
    if ($macro.Reports.Report.Localizations.Localization.externalName -ne 'Trafo zoeker Macro') {
        throw 'Macronaam klopt niet.'
    }
    if ($macro.Reports.Report.rwoProviderId -ne 'electricity') {
        throw 'Lovion-provider klopt niet.'
    }
    if ($macro.Reports.Report.rwoTypeName -ne 'elec_e_lv_i_transformer_group') {
        throw 'Lovion-type klopt niet.'
    }
    if ($macro.Reports.LovionFileFormat.Version.minRelease -ne '7.2.1') {
        throw 'Minimum Lovion-release klopt niet.'
    }
    if ($macro.Reports.Report.And.Field.Report.And.Field.Report.And.Field.Report.And.Field.Report.And.Field.name -ne 'number') {
        throw 'Station-parameter ontbreekt.'
    }

    Export-LovionMacroXml -Path $exportPath | Out-Null
    [byte[]]$sourceBytes = Get-LovionMacroBytes
    [byte[]]$exportBytes = [System.IO.File]::ReadAllBytes($exportPath)
    if (-not [System.Linq.Enumerable]::SequenceEqual($sourceBytes, $exportBytes)) {
        throw 'De gedownloade XML wijkt af van de ingebouwde macrobron.'
    }

    Write-Output 'Lovion macro source/export: PASS'
}
finally {
    if (Test-Path -LiteralPath $exportPath) {
        Remove-Item -LiteralPath $exportPath -Force
    }
}
