$ErrorActionPreference = 'Stop'

$sourcePath = Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1'
. (Resolve-Path -LiteralPath $sourcePath).Path

function Get-ControlTexts {
    param([System.Windows.Forms.Control]$Control)

    $texts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace([string]$Control.Text)) {
        $texts.Add([string]$Control.Text)
    }
    foreach ($child in @($Control.Controls)) {
        foreach ($text in @(Get-ControlTexts -Control $child)) {
            $texts.Add($text)
        }
    }
    if ($Control -is [System.Windows.Forms.ToolStrip]) {
        foreach ($item in @($Control.Items)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item.Text)) {
                $texts.Add([string]$item.Text)
            }
            if ($item -is [System.Windows.Forms.ToolStripDropDownItem]) {
                foreach ($dropDownItem in @($item.DropDownItems)) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$dropDownItem.Text)) {
                        $texts.Add([string]$dropDownItem.Text)
                    }
                }
            }
        }
    }
    return @($texts)
}

$form = Build-MainForm
try {
    $texts = @(Get-ControlTexts -Control $form)
    foreach ($removedText in @('PDOK', 'Enexis', 'Terms', 'Herlees', 'Street Terms')) {
        if ($texts -contains $removedText) {
            throw "Removed GUI control is still visible: $removedText"
        }
    }
    foreach ($requiredText in @('Klembord', 'Plakken', 'Stations', 'Lovion macro')) {
        if ($texts -notcontains $requiredText) {
            throw "Required GUI control is missing: $requiredText"
        }
    }
    Write-Output 'GUI surface button removal: PASS'
}
finally {
    $form.Dispose()
}
