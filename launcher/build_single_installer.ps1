$ErrorActionPreference = "Stop"

function Test-IsExcludedDirectory {
  param(
    [string]$Name
  )

  if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
  $blocked = @(".git", ".vs", "node_modules", "launcher", "dist")
  if ($Name.StartsWith("_inspect_", [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
  return $blocked -contains $Name
}

function Test-IsExcludedFile {
  param(
    [string]$Name
  )

  if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
  $extension = [System.IO.Path]::GetExtension($Name)
  $blocked = @(".ps1", ".cmd", ".bat", ".exe", ".pdb", ".cs")
  return $blocked -contains $extension.ToLowerInvariant()
}

function Copy-PluginTree {
  param(
    [string]$SourceRoot,
    [string]$DestinationRoot
  )

  if (-not (Test-Path -LiteralPath $SourceRoot)) {
    throw "Source root not found: $SourceRoot"
  }

  New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null

  Get-ChildItem -LiteralPath $SourceRoot -Force | ForEach-Object {
    if ($_.PSIsContainer) {
      if (Test-IsExcludedDirectory -Name $_.Name) {
        return
      }
      Copy-PluginTree -SourceRoot $_.FullName -DestinationRoot (Join-Path $DestinationRoot $_.Name)
      return
    }

    if (Test-IsExcludedFile -Name $_.Name) {
      return
    }

    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $DestinationRoot $_.Name) -Force
  }
}

function Get-ManifestField {
  param(
    [string]$ManifestText,
    [string]$Field
  )

  $pattern = '"' + [regex]::Escape($Field) + '"\s*:\s*"([^"]*)"'
  $match = [regex]::Match($ManifestText, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if ($match.Success) {
    return $match.Groups[1].Value
  }
  return ""
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pluginRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
$manifestPath = Join-Path $pluginRoot "manifest.json"

if (-not (Test-Path -LiteralPath $manifestPath)) {
  throw "manifest.json not found at plugin root: $pluginRoot"
}

$manifestText = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8
$pluginId = Get-ManifestField -ManifestText $manifestText -Field "id"
$pluginVersion = Get-ManifestField -ManifestText $manifestText -Field "version"

if ([string]::IsNullOrWhiteSpace($pluginId)) {
  throw "manifest.json is missing 'id'"
}

$distDir = Join-Path $scriptRoot "dist"
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

$tempRoot = Join-Path $env:TEMP ("psd_export_installer_build_" + [guid]::NewGuid().ToString("N"))
$payloadRoot = Join-Path $tempRoot "payload"
$zipPath = Join-Path $tempRoot "payload.zip"

try {
  New-Item -ItemType Directory -Path $payloadRoot -Force | Out-Null
  Copy-PluginTree -SourceRoot $pluginRoot -DestinationRoot $payloadRoot

  Compress-Archive -Path (Join-Path $payloadRoot "*") -DestinationPath $zipPath -Force
  $payloadBytes = [System.IO.File]::ReadAllBytes($zipPath)
  $payloadBase64 = [System.Convert]::ToBase64String($payloadBytes)

  $chunkSize = 12000
  $chunks = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $payloadBase64.Length; $i += $chunkSize) {
    $len = [Math]::Min($chunkSize, $payloadBase64.Length - $i)
    $chunks.Add($payloadBase64.Substring($i, $len))
  }
  $payloadLiteral = ($chunks | ForEach-Object { '"' + $_ + '"' }) -join "`r`n            + "

  $safePluginId = $pluginId.Replace("\", "\\").Replace('"', '\"')
  $safeVersionRaw = if ($null -eq $pluginVersion) { "" } else { $pluginVersion }
  $safeVersion = $safeVersionRaw.Replace("\", "\\").Replace('"', '\"')

  $installerSource = @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

internal static class PsdExportSingleInstaller
{
    private const string PluginId = "$safePluginId";
    private const string PluginVersion = "$safeVersion";
    private static readonly string[] HostNames = { "PHSP", "PHSPBETA" };
    private static readonly string[] InstallBuckets = { "Developer", "External" };
    private static readonly string PayloadBase64 =
            $payloadLiteral;

    public static int Main(string[] args)
    {
        try
        {
            var uninstall = HasArg(args, "uninstall", "/uninstall", "-u");
            var status = HasArg(args, "status", "/status", "-s");
            var noLaunch = HasArg(args, "--no-launch", "/no-launch");
            var preferBeta = HasArg(args, "--beta", "/beta");

            Console.WriteLine("PSD Export Pipeline Installer");
            Console.WriteLine("Plugin Id : " + PluginId);
            Console.WriteLine("Version   : " + (string.IsNullOrEmpty(PluginVersion) ? "(unknown)" : PluginVersion));
            Console.WriteLine();

            if (status)
            {
                PrintStatus();
                return 0;
            }

            if (uninstall)
            {
                Uninstall();
                return 0;
            }

            var tempRoot = ExtractPayload();
            try
            {
                var payloadRoot = Path.Combine(tempRoot, "payload");
                Install(payloadRoot);
            }
            finally
            {
                TryDeleteDirectory(tempRoot);
            }

            if (!noLaunch)
            {
                LaunchPhotoshop(preferBeta);
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Installer failed: " + ex.Message);
            return 1;
        }
    }

    private static bool HasArg(IReadOnlyList<string> args, params string[] names)
    {
        if (args == null || names == null) return false;
        for (var i = 0; i < args.Count; i++)
        {
            var value = (args[i] ?? string.Empty).Trim();
            for (var j = 0; j < names.Length; j++)
            {
                if (string.Equals(value, names[j], StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        return false;
    }

    private static string ExtractPayload()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), "PSDExportSingleInstaller_" + Guid.NewGuid().ToString("N"));
        var zipPath = Path.Combine(tempRoot, "payload.zip");
        var payloadRoot = Path.Combine(tempRoot, "payload");
        Directory.CreateDirectory(tempRoot);
        Directory.CreateDirectory(payloadRoot);

        var bytes = Convert.FromBase64String(PayloadBase64);
        File.WriteAllBytes(zipPath, bytes);
        ZipFile.ExtractToDirectory(zipPath, payloadRoot);
        return tempRoot;
    }

    private static void Install(string payloadRoot)
    {
        if (!Directory.Exists(payloadRoot))
        {
            throw new InvalidOperationException("Payload root does not exist: " + payloadRoot);
        }

        var installs = 0;
        foreach (var versionDir in EnumerateHostVersionDirectories())
        {
            foreach (var bucket in InstallBuckets)
            {
                var bucketDir = Path.Combine(versionDir.FullName, bucket);
                Directory.CreateDirectory(bucketDir);
                var destination = Path.Combine(bucketDir, PluginId);
                Console.WriteLine("Installing -> " + destination);
                MirrorDirectory(payloadRoot, destination);
                installs += 1;
            }
        }

        var localExternalRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Adobe",
            "UXP",
            "Plugins",
            "External"
        );
        Directory.CreateDirectory(localExternalRoot);
        var version = string.IsNullOrWhiteSpace(PluginVersion) ? "0.0.0" : PluginVersion.Trim();
        var localExternalFolderName = PluginId + "_" + version;
        var localExternalTarget = Path.Combine(localExternalRoot, localExternalFolderName);
        DeleteLocalExternalPluginCopies(localExternalRoot);
        Console.WriteLine("Installing -> " + localExternalTarget);
        MirrorDirectory(payloadRoot, localExternalTarget);
        installs += 1;

        if (installs == 0)
        {
            Console.WriteLine("No PHSP/PHSPBETA version folders found.");
            Console.WriteLine("Please open Photoshop once, then run installer again.");
            return;
        }

        Console.WriteLine();
        Console.WriteLine("Install completed (" + installs + " targets).");
    }

    private static void Uninstall()
    {
        var removed = 0;
        foreach (var versionDir in EnumerateHostVersionDirectories())
        {
            foreach (var bucket in InstallBuckets)
            {
                var target = Path.Combine(versionDir.FullName, bucket, PluginId);
                if (!Directory.Exists(target))
                {
                    continue;
                }
                Console.WriteLine("Removing -> " + target);
                Directory.Delete(target, true);
                removed += 1;
            }
        }

        var localExternalRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Adobe",
            "UXP",
            "Plugins",
            "External"
        );
        if (Directory.Exists(localExternalRoot))
        {
            foreach (var dir in new DirectoryInfo(localExternalRoot).GetDirectories())
            {
                if (string.Equals(dir.Name, PluginId, StringComparison.OrdinalIgnoreCase)
                    || dir.Name.StartsWith(PluginId + "_", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Removing -> " + dir.FullName);
                    Directory.Delete(dir.FullName, true);
                    removed += 1;
                }
            }
        }

        Console.WriteLine();
        Console.WriteLine("Uninstall completed. Removed folders: " + removed);
    }

    private static void DeleteLocalExternalPluginCopies(string localExternalRoot)
    {
        if (!Directory.Exists(localExternalRoot))
        {
            return;
        }

        foreach (var dir in new DirectoryInfo(localExternalRoot).GetDirectories())
        {
            if (string.Equals(dir.Name, PluginId, StringComparison.OrdinalIgnoreCase)
                || dir.Name.StartsWith(PluginId + "_", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Removing old local external copy -> " + dir.FullName);
                Directory.Delete(dir.FullName, true);
            }
        }
    }

    private static void PrintStatus()
    {
        var found = new List<string>();
        foreach (var versionDir in EnumerateHostVersionDirectories())
        {
            foreach (var bucket in InstallBuckets)
            {
                var target = Path.Combine(versionDir.FullName, bucket, PluginId);
                if (Directory.Exists(target))
                {
                    found.Add(target);
                }
            }
        }

        var localExternalRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Adobe",
            "UXP",
            "Plugins",
            "External"
        );
        if (Directory.Exists(localExternalRoot))
        {
            foreach (var dir in new DirectoryInfo(localExternalRoot).GetDirectories())
            {
                if (string.Equals(dir.Name, PluginId, StringComparison.OrdinalIgnoreCase)
                    || dir.Name.StartsWith(PluginId + "_", StringComparison.OrdinalIgnoreCase))
                {
                    found.Add(dir.FullName);
                }
            }
        }

        if (found.Count == 0)
        {
            Console.WriteLine("Plugin is not currently installed.");
            return;
        }

        Console.WriteLine("Installed locations:");
        for (var i = 0; i < found.Count; i++)
        {
            Console.WriteLine("  " + found[i]);
        }
    }

    private static IEnumerable<DirectoryInfo> EnumerateHostVersionDirectories()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var roots = HostNames
            .Select((host) => Path.Combine(appData, "Adobe", "UXP", "PluginsStorage", host))
            .Where(Directory.Exists);

        var all = new List<DirectoryInfo>();
        foreach (var root in roots)
        {
            var rootInfo = new DirectoryInfo(root);
            foreach (var dir in rootInfo.GetDirectories())
            {
                if (ParseIntegerSafe(dir.Name) <= 0)
                {
                    continue;
                }
                all.Add(dir);
            }
        }

        return all.OrderByDescending((dir) => ParseIntegerSafe(dir.Name));
    }

    private static int ParseIntegerSafe(string value)
    {
        int parsed;
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
        {
            return parsed;
        }
        return -1;
    }

    private static void MirrorDirectory(string sourceRoot, string destinationRoot)
    {
        if (Directory.Exists(destinationRoot))
        {
            Directory.Delete(destinationRoot, true);
        }

        Directory.CreateDirectory(destinationRoot);
        CopyRecursive(new DirectoryInfo(sourceRoot), new DirectoryInfo(destinationRoot));
    }

    private static void CopyRecursive(DirectoryInfo sourceDir, DirectoryInfo destinationDir)
    {
        var files = sourceDir.GetFiles();
        foreach (var file in files)
        {
            var targetPath = Path.Combine(destinationDir.FullName, file.Name);
            file.CopyTo(targetPath, true);
        }

        var dirs = sourceDir.GetDirectories();
        foreach (var dir in dirs)
        {
            var targetSubdir = destinationDir.CreateSubdirectory(dir.Name);
            CopyRecursive(dir, targetSubdir);
        }
    }

    private static void LaunchPhotoshop(bool preferBeta)
    {
        var photoshopExe = FindPhotoshopExecutable(preferBeta);
        if (string.IsNullOrEmpty(photoshopExe))
        {
            Console.WriteLine("Photoshop executable not found under Program Files\\Adobe.");
            return;
        }

        Console.WriteLine("Starting Photoshop: " + photoshopExe);
        Process.Start(new ProcessStartInfo
        {
            FileName = photoshopExe,
            UseShellExecute = true,
        });
    }

    private static string FindPhotoshopExecutable(bool preferBeta)
    {
        var adobeRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "Adobe"
        );
        if (!Directory.Exists(adobeRoot))
        {
            return string.Empty;
        }

        var bestPath = string.Empty;
        var bestScore = int.MinValue;
        foreach (var dir in new DirectoryInfo(adobeRoot).GetDirectories("Adobe Photoshop*"))
        {
            var exePath = Path.Combine(dir.FullName, "Photoshop.exe");
            if (!File.Exists(exePath))
            {
                continue;
            }

            var isBeta = dir.Name.IndexOf("beta", StringComparison.OrdinalIgnoreCase) >= 0;
            var year = 0;
            var match = Regex.Match(dir.Name, "(20\\d{2})");
            if (match.Success)
            {
                int.TryParse(match.Groups[1].Value, out year);
            }

            var score = year;
            if (preferBeta && isBeta) score += 10000;
            if (!preferBeta && isBeta) score -= 1000;

            if (score > bestScore)
            {
                bestScore = score;
                bestPath = exePath;
            }
        }

        return bestPath;
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }
        catch
        {
        }
    }
}
"@

  $sourceOut = Join-Path $distDir "PsdExportSingleInstaller.generated.cs"
  Set-Content -LiteralPath $sourceOut -Value $installerSource -Encoding UTF8

  $outputExe = Join-Path $distDir ("PSDExportPipelineInstaller_" + ($pluginVersion -replace '[^0-9A-Za-z\.-]', '_') + ".exe")
  if (Test-Path -LiteralPath $outputExe) {
    Remove-Item -LiteralPath $outputExe -Force
  }

  Add-Type `
    -TypeDefinition $installerSource `
    -Language CSharp `
    -ReferencedAssemblies @("System.IO.Compression.dll", "System.IO.Compression.FileSystem.dll") `
    -OutputAssembly $outputExe `
    -OutputType ConsoleApplication

  Write-Host "Single installer build succeeded:"
  Write-Host "  $outputExe"
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
