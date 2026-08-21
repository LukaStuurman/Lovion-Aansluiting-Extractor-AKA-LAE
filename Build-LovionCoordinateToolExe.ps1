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

$exePath = Join-Path -Path $OutputDirectory -ChildPath 'LovionCoordinateTool.exe'
if (Test-Path -Path $exePath) {
    Remove-Item -Path $exePath -Force
}

$scriptPath = Join-Path $scriptRoot 'Get-LovionCoordinates.ps1'
$scriptBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-Content -Path $scriptPath -Raw -Encoding UTF8)))

$launcherSource = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

internal static class Program
{
    private static string Quote(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "\"\"";
        }

        if (!value.Any(char.IsWhiteSpace) && !value.Contains("\""))
        {
            return value;
        }

        return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }

    public static int Main(string[] args)
    {
        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var embeddedScript = Encoding.UTF8.GetString(Convert.FromBase64String("$scriptBase64"));
        var tempScriptPath = Path.Combine(Path.GetTempPath(), "LovionCoordinateTool_" + Guid.NewGuid().ToString("N") + ".ps1");
        File.WriteAllText(tempScriptPath, embeddedScript, Encoding.UTF8);

        var powerShellPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
        if (!File.Exists(powerShellPath))
        {
            powerShellPath = "powershell.exe";
        }

        var forwardedArguments = args.Length == 0
            ? string.Empty
            : " " + string.Join(" ", args.Select(Quote));

        var processStartInfo = new ProcessStartInfo
        {
            FileName = powerShellPath,
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File " + Quote(tempScriptPath) + forwardedArguments,
            WorkingDirectory = baseDirectory,
            UseShellExecute = false
        };

        try
        {
            using (var process = Process.Start(processStartInfo))
            {
                process.WaitForExit();
                return process.ExitCode;
            }
        }
        finally
        {
            try
            {
                if (File.Exists(tempScriptPath))
                {
                    File.Delete(tempScriptPath);
                }
            }
            catch
            {
            }
        }
    }
}
"@

Add-Type -TypeDefinition $launcherSource -Language CSharp -OutputAssembly $exePath -OutputType ConsoleApplication

Copy-Item -Path $scriptPath -Destination (Join-Path $OutputDirectory 'Get-LovionCoordinates.ps1') -Force
Copy-Item -Path (Join-Path $scriptRoot 'README.md') -Destination (Join-Path $OutputDirectory 'README.md') -Force

Write-Host "EXE gebouwd: $exePath"
