$ErrorActionPreference='Stop'
. (Resolve-Path (Join-Path $PSScriptRoot '..\LovionCoordinateTool.Gui.ps1')).Path
function Assert-Test {param([bool]$Condition,[string]$Message) if(-not $Condition){throw "TEST FAILED: $Message"}}
[Windows.Forms.Application]::EnableVisualStyles()
$form=Build-MainForm
try {
    for($i=0;$i -lt 98;$i++){Add-TableColumn -Name ('FixtureColumn'+$i)}
    for($i=0;$i -lt 2641;$i++){
        $row=$script:Table.NewRow()
        $row['SourceFile']='001.001'
        $row['RecordTitle']='Synthetic connection '+$i
        $script:Table.Rows.Add($row)
    }
    $form.Show()
    [Windows.Forms.Application]::DoEvents()
    $grid=$script:Grid
    $fullRowCount=$grid.Rows.Count
    Assert-Test ($fullRowCount -eq 2642) 'all 2641 data rows and the new-row placeholder remain in the grid'
    $provider=$grid.AccessibilityObject
    $childCount=$provider.GetChildCount()
    Assert-Test ($childCount -gt 2 -and $childCount -lt 100) 'only displayed rows, headers and scrollbars are enumerated'
    for($i=0;$i -lt $childCount;$i++){
        $child=$provider.GetChild($i)
        Assert-Test ($null -ne $child) 'enumerated accessibility child exists'
        # Simulate a legacy accessibility client visiting every visible cell.
        for($j=0;$j -lt $child.GetChildCount();$j++){$null=$child.GetChild($j)}
    }
    $firstName=$provider.GetChild(1).Name
    $grid.FirstDisplayedScrollingRowIndex=2600
    [Windows.Forms.Application]::DoEvents()
    Assert-Test ($provider.GetChildCount() -lt 100) 'enumeration stays bounded after scrolling'
    Assert-Test ($provider.GetChild(1).Name -ne $firstName) 'scrolling exposes the real newly displayed row'
    Assert-Test ($null -ne $provider.GetFocused()) 'keyboard focus remains accessible'
    Assert-Test ($grid.Rows.Count -eq $fullRowCount -and $script:Table.Rows.Count -eq 2641) 'accessibility enumeration does not remove or page the data'
    Assert-Test ($null -eq $provider.GetChild(-1) -and $null -eq $provider.GetChild($provider.GetChildCount())) 'out-of-range child requests return null'
    Write-Output 'Large grid accessibility virtualization: PASS (2641 rows retained)'
} finally { $form.Close();$form.Dispose() }
