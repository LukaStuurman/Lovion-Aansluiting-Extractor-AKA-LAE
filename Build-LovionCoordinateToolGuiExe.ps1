[CmdletBinding()]
param(
    [string]$OutputDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path -Path $scriptRoot -ChildPath 'dist'
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
}

$portableDirectory = Join-Path -Path $OutputDirectory -ChildPath 'portable'
if (-not [System.IO.Directory]::Exists($portableDirectory)) {
    [void][System.IO.Directory]::CreateDirectory($portableDirectory)
}

$exePath = Join-Path -Path $OutputDirectory -ChildPath 'LovionCoordinateWorkbench.exe'
if (Test-Path -Path $exePath) {
    Remove-Item -Path $exePath -Force
}

$scriptPath = Join-Path $scriptRoot 'LovionCoordinateTool.Gui.ps1'
$appDirectory = Join-Path -Path $OutputDirectory -ChildPath 'app'
if (-not [System.IO.Directory]::Exists($appDirectory)) {
    [void][System.IO.Directory]::CreateDirectory($appDirectory)
}

$launcherSource = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    public static int Main()
    {
        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var scriptPath = Path.Combine(baseDirectory, "app", "LovionCoordinateTool.Gui.ps1");
        if (!File.Exists(scriptPath))
        {
            MessageBox.Show(
                "Het applicatiebestand ontbreekt:\n" + scriptPath,
                "Lovion Coordinate Tool",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 2;
        }

        var powerShellPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
        if (!File.Exists(powerShellPath))
        {
            powerShellPath = "powershell.exe";
        }

        var processStartInfo = new ProcessStartInfo
        {
            FileName = powerShellPath,
            Arguments = "-NoProfile -Sta -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + scriptPath + "\"",
            WorkingDirectory = baseDirectory,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (var process = Process.Start(processStartInfo))
        {
            process.WaitForExit();
            return process.ExitCode;
        }
    }
}
"@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition $launcherSource -Language CSharp -ReferencedAssemblies System.Windows.Forms -OutputAssembly $exePath -OutputType WindowsApplication

if (Test-Path -Path (Join-Path $OutputDirectory 'LovionCoordinateTool.Gui.ps1')) {
    Remove-Item -Path (Join-Path $OutputDirectory 'LovionCoordinateTool.Gui.ps1') -Force
}
Copy-Item -Path (Join-Path $scriptRoot 'README.md') -Destination (Join-Path $OutputDirectory 'README.md') -Force
[System.IO.File]::Copy($scriptPath, (Join-Path $appDirectory 'LovionCoordinateTool.Gui.ps1'), $true)
[System.IO.File]::Copy($exePath, (Join-Path $portableDirectory 'LovionCoordinateWorkbench.exe'), $true)
[System.IO.File]::Copy((Join-Path $scriptRoot 'README.md'), (Join-Path $portableDirectory 'README.md'), $true)
$portableAppDirectory = Join-Path -Path $portableDirectory -ChildPath 'app'
if (-not [System.IO.Directory]::Exists($portableAppDirectory)) {
    [void][System.IO.Directory]::CreateDirectory($portableAppDirectory)
}
[System.IO.File]::Copy($scriptPath, (Join-Path $portableAppDirectory 'LovionCoordinateTool.Gui.ps1'), $true)

Write-Host "GUI EXE gebouwd: $exePath"
