$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pluginRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
$manifestPath = Join-Path $pluginRoot "manifest.json"
$distDir = Join-Path $scriptRoot "dist"

if (-not (Test-Path -LiteralPath $manifestPath)) {
  throw "manifest.json not found: $manifestPath"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scriptRoot "build_ccx_package.ps1")
if ($LASTEXITCODE -ne 0) {
  throw "build_ccx_package.ps1 failed"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scriptRoot "build_manager.ps1")
if ($LASTEXITCODE -ne 0) {
  throw "build_manager.ps1 failed"
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$version = if ($manifest.version) { $manifest.version } else { "unknown" }
$safeVersion = $version -replace '[^0-9A-Za-z\.-]', '_'

$managerExe = Join-Path $distDir "PSDExportManager.exe"
$ccx = Get-ChildItem -LiteralPath $distDir -Filter "PSDExportPipeline_*.ccx" |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1

if (-not (Test-Path -LiteralPath $managerExe)) {
  throw "Manager exe not found: $managerExe"
}
if (-not $ccx) {
  throw "CCX package not found under $distDir"
}

$tempRoot = Join-Path $env:TEMP ("psd_export_manager_" + [guid]::NewGuid().ToString("N"))
$packageRoot = Join-Path $tempRoot "PSDExportManager"
$zipPath = Join-Path $distDir "PSDExportManager_$safeVersion.zip"

try {
  New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
  Copy-Item -LiteralPath $managerExe -Destination $packageRoot -Force
  Copy-Item -LiteralPath $ccx.FullName -Destination $packageRoot -Force

  $readme = @"
PSD Export Manager

1. Run PSDExportManager.exe.
2. Click Install Current Version for first install.
3. Click Check Updates to compare with GitHub.
4. Click Update and Install when an update is available.

Requirements:
- Git for Windows for update checks/downloads.
- Adobe Creative Cloud Desktop / UnifiedPluginInstallerAgent for CCX install.
"@
  Set-Content -LiteralPath (Join-Path $packageRoot "README_MANAGER.txt") -Value $readme -Encoding ASCII

  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }
  Compress-Archive -Path $packageRoot -DestinationPath $zipPath -Force
  Write-Host "Manager package build succeeded:"
  Write-Host "  $zipPath"
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
