$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pluginRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
$manifestPath = Join-Path $pluginRoot "manifest.json"
$distDir = Join-Path $scriptRoot "dist"

if (-not (Test-Path -LiteralPath $manifestPath)) {
  throw "manifest.json not found: $manifestPath"
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$version = if ($manifest.version) { $manifest.version } else { "unknown" }
$safeVersion = $version -replace '[^0-9A-Za-z\.-]', '_'

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

$tempRoot = Join-Path $env:TEMP ("psd_export_updater_" + [guid]::NewGuid().ToString("N"))
$packageRoot = Join-Path $tempRoot "PSDExportUpdater"
$zipPath = Join-Path $distDir "PSDExportUpdater_$safeVersion.zip"

try {
  New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $scriptRoot "PSDExportUpdater.cmd") -Destination $packageRoot -Force
  Copy-Item -LiteralPath (Join-Path $scriptRoot "update_from_git.ps1") -Destination $packageRoot -Force

  $readme = @"
PSD Export Pipeline Updater

1. Run PSDExportUpdater.cmd.
2. Choose "Set Git repo URL" the first time.
3. Choose branch/tag.
4. Choose "Update + build CCX + install".

Settings are stored at:
%APPDATA%\PSDExportPipeline\updater.json

Requirements:
- Git for Windows
- Adobe Creative Cloud Desktop / UnifiedPluginInstallerAgent
"@
  Set-Content -LiteralPath (Join-Path $packageRoot "README_UPDATER.txt") -Value $readme -Encoding ASCII

  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }
  Compress-Archive -Path $packageRoot -DestinationPath $zipPath -Force
  Write-Host "Updater package build succeeded:"
  Write-Host "  $zipPath"
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
