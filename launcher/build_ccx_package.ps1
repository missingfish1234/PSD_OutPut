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
  $blocked = @(".ps1", ".cmd", ".bat", ".exe", ".pdb", ".cs", ".zip", ".ccx")
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

function Get-UpiaPath {
  $candidates = @(
    (Join-Path ${env:ProgramFiles} "Common Files\Adobe\Adobe Desktop Common\RemoteComponents\UPI\UnifiedPluginInstallerAgent\UnifiedPluginInstallerAgent.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Common Files\Adobe\Adobe Desktop Common\RemoteComponents\UPI\UnifiedPluginInstallerAgent\UnifiedPluginInstallerAgent.exe")
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
      return $candidate
    }
  }

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

$safeVersion = $pluginVersion -replace '[^0-9A-Za-z\.-]', '_'
$packageName = "PSDExportPipeline_$safeVersion"
$tempRoot = Join-Path $env:TEMP ("psd_export_ccx_" + [guid]::NewGuid().ToString("N"))
$packageRoot = Join-Path $tempRoot "payload"
$zipPath = Join-Path $tempRoot "$packageName.zip"
$ccxPath = Join-Path $distDir "$packageName.ccx"
$cmdPath = Join-Path $distDir "Install_$packageName.cmd"
$upiaPath = Get-UpiaPath

try {
  Copy-PluginTree -SourceRoot $pluginRoot -DestinationRoot $packageRoot

  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }
  if (Test-Path -LiteralPath $ccxPath) {
    Remove-Item -LiteralPath $ccxPath -Force
  }

  Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force
  Move-Item -LiteralPath $zipPath -Destination $ccxPath -Force

$cmdText = @"
@echo off
setlocal
set "CCX=%~dp0$packageName.ccx"
if not "%~1"=="" set "CCX=%~1"
set "UPIA=$upiaPath"
if not exist "%CCX%" (
  echo Missing CCX: "%CCX%"
  pause
  exit /b 1
)
if not exist "%UPIA%" (
  set "UPIA=%ProgramFiles%\Common Files\Adobe\Adobe Desktop Common\RemoteComponents\UPI\UnifiedPluginInstallerAgent\UnifiedPluginInstallerAgent.exe"
)
if not exist "%UPIA%" (
  set "UPIA=%ProgramFiles(x86)%\Common Files\Adobe\Adobe Desktop Common\RemoteComponents\UPI\UnifiedPluginInstallerAgent\UnifiedPluginInstallerAgent.exe"
)
if not exist "%UPIA%" (
  echo UnifiedPluginInstallerAgent.exe was not found. Install or update Adobe Creative Cloud Desktop, then run this again.
  pause
  exit /b 1
)
echo Installing "%CCX%" with:
echo   "%UPIA%"
"%UPIA%" /install "%CCX%"
if errorlevel 1 (
  echo.
  echo First install command failed. Trying --install syntax...
  "%UPIA%" --install "%CCX%"
)
echo.
echo Done. Restart Photoshop and check Plugins / UXP panel list.
pause
"@
  Set-Content -LiteralPath $cmdPath -Value $cmdText -Encoding ASCII

  Write-Host "CCX package build succeeded:"
  Write-Host "  $ccxPath"
  Write-Host "Installer helper:"
  Write-Host "  $cmdPath"
  if ([string]::IsNullOrWhiteSpace($upiaPath)) {
    Write-Host ""
    Write-Host "Warning: UnifiedPluginInstallerAgent.exe was not found on this computer."
    Write-Host "Install or update Adobe Creative Cloud Desktop before running the helper cmd."
  }
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
