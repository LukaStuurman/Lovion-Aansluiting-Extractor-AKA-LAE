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

$exePath = Join-Path -Path $OutputDirectory -ChildPath 'LovionCoordinateWorkbench.exe'
if (Test-Path -Path $exePath) {
    Remove-Item -Path $exePath -Force
}

$scriptPath = Join-Path $scriptRoot 'LovionCoordinateTool.Gui.ps1'
$scriptText = Get-Content -Path $scriptPath -Raw -Encoding UTF8
$macroPath = Join-Path $scriptRoot 'assets\Trafo zoeken macro rapport.xml'
if (-not (Test-Path -LiteralPath $macroPath -PathType Leaf)) {
    throw "Het Lovion-macrobestand ontbreekt: $macroPath"
}
$macroBytes = [System.IO.File]::ReadAllBytes($macroPath)
$macroBase64 = [Convert]::ToBase64String($macroBytes)

# De standalone EXE voert het ingebedde script uit vanuit een tijdelijke map.
# Laat configuratie en logs toch naast de EXE terechtkomen in plaats van in %TEMP%.
$originalDirectoryLine = '$script:ScriptDirectory = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }'
$embeddedDirectoryLine = '$script:ScriptDirectory = if (-not [string]::IsNullOrWhiteSpace($env:LOVION_APP_BASE_DIRECTORY)) { $env:LOVION_APP_BASE_DIRECTORY } elseif ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }'
if (-not $scriptText.Contains($originalDirectoryLine)) {
    throw 'Kan ScriptDirectory-regel niet vinden in LovionCoordinateTool.Gui.ps1.'
}
$scriptText = $scriptText.Replace($originalDirectoryLine, $embeddedDirectoryLine)
$originalMacroLine = '$script:EmbeddedLovionMacroXmlBase64 = $null'
$embeddedMacroLine = '$script:EmbeddedLovionMacroXmlBase64 = ''' + $macroBase64 + ''''
if (-not $scriptText.Contains($originalMacroLine)) {
    throw 'Kan de macro-embedregel niet vinden in LovionCoordinateTool.Gui.ps1.'
}
$scriptText = $scriptText.Replace($originalMacroLine, $embeddedMacroLine)

# Comprimeer het script voordat het in de launcher wordt ingebed. Dit houdt de C#-bron
# en de uiteindelijke EXE compact, terwijl er geen los .ps1-bestand hoeft te worden meegeleverd.
$scriptBytes = [System.Text.Encoding]::UTF8.GetBytes($scriptText)
$compressedStream = New-Object System.IO.MemoryStream
$gzipStream = New-Object System.IO.Compression.GZipStream($compressedStream, [System.IO.Compression.CompressionMode]::Compress, $true)
try {
    $gzipStream.Write($scriptBytes, 0, $scriptBytes.Length)
}
finally {
    $gzipStream.Dispose()
}
$compressedScriptBase64 = [Convert]::ToBase64String($compressedStream.ToArray())
$compressedStream.Dispose()

$launcherSource = @"
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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

    private static byte[] DecompressScript()
    {
        var compressed = Convert.FromBase64String("$compressedScriptBase64");
        using (var input = new MemoryStream(compressed))
        using (var gzip = new GZipStream(input, CompressionMode.Decompress))
        using (var output = new MemoryStream())
        {
            gzip.CopyTo(output);
            return output.ToArray();
        }
    }

    [STAThread]
    public static int Main(string[] args)
    {
        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var tempDirectory = Path.Combine(Path.GetTempPath(), "LovionCoordinateWorkbench");
        Directory.CreateDirectory(tempDirectory);
        var tempScriptPath = Path.Combine(tempDirectory, "LovionCoordinateTool.Gui_" + Guid.NewGuid().ToString("N") + ".ps1");

        try
        {
            File.WriteAllBytes(tempScriptPath, DecompressScript());

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
                Arguments = "-NoProfile -Sta -WindowStyle Hidden -ExecutionPolicy Bypass -File " + Quote(tempScriptPath) + forwardedArguments,
                WorkingDirectory = baseDirectory,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            processStartInfo.EnvironmentVariables["LOVION_APP_BASE_DIRECTORY"] = baseDirectory;

            using (var process = Process.Start(processStartInfo))
            {
                process.WaitForExit();
                return process.ExitCode;
            }
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                "Lovion Coordinate Workbench kon niet worden gestart:\n" + exception.Message,
                "Lovion Coordinate Workbench",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 2;
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition $launcherSource -Language CSharp -ReferencedAssemblies System.Windows.Forms -OutputAssembly $exePath -OutputType WindowsApplication

Write-Host "Standalone GUI EXE gebouwd: $exePath"
