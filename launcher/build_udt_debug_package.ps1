$ErrorActionPreference = "Stop"

function Test-IsExcludedDirectory {
  param([string]$Name)

  if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
  $blocked = @(".git", ".vs", "node_modules", "launcher", "dist")
  if ($Name.StartsWith("_inspect_", [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
  return $blocked -contains $Name
}

function Test-IsExcludedFile {
  param([string]$Name)

  if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
  $extension = [System.IO.Path]::GetExtension($Name).ToLowerInvariant()
  $blocked = @(".ps1", ".cmd", ".bat", ".exe", ".pdb", ".cs")
  return $blocked -contains $extension
}

function Copy-PluginTree {
  param(
    [string]$SourceRoot,
    [string]$DestinationRoot
  )

  New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null

  Get-ChildItem -LiteralPath $SourceRoot -Force | ForEach-Object {
    if ($_.PSIsContainer) {
      if (Test-IsExcludedDirectory -Name $_.Name) { return }
      Copy-PluginTree -SourceRoot $_.FullName -DestinationRoot (Join-Path $DestinationRoot $_.Name)
      return
    }

    if (Test-IsExcludedFile -Name $_.Name) { return }
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
  if ($match.Success) { return $match.Groups[1].Value }
  return ""
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pluginRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
$distDir = Join-Path $scriptRoot "dist"
$manifestPath = Join-Path $pluginRoot "manifest.json"

if (-not (Test-Path -LiteralPath $manifestPath)) {
  throw "manifest.json not found: $manifestPath"
}

$manifestText = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8
$pluginVersion = Get-ManifestField -ManifestText $manifestText -Field "version"
if ([string]::IsNullOrWhiteSpace($pluginVersion)) {
  $pluginVersion = "unknown"
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

$packageName = "PSDExportPipeline_UDT_Debug_$pluginVersion"
$tempRoot = Join-Path $env:TEMP ("psd_export_udt_debug_" + [guid]::NewGuid().ToString("N"))
$packageRoot = Join-Path $tempRoot $packageName
$zipPath = Join-Path $distDir "$packageName.zip"

try {
  Copy-PluginTree -SourceRoot $pluginRoot -DestinationRoot $packageRoot

  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }

  Compress-Archive -Path $packageRoot -DestinationPath $zipPath -Force
  Write-Host "UDT debug package build succeeded:"
  Write-Host "  $zipPath"
  Write-Host ""
  Write-Host "Unzip it, then UXP Developer Tools > Add Plugin should point to:"
  Write-Host "  <unzipped>\\$packageName"
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
