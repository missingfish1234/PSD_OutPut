param(
  [Parameter(Mandatory = $true)]
  [string]$CcxPath,
  [string]$PluginId = ""
)

$ErrorActionPreference = "Stop"

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

function Copy-TreeFiles {
  param(
    [string]$SourceRoot,
    [string]$DestinationRoot
  )

  $files = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -Force
  foreach ($file in $files) {
    $relative = $file.FullName.Substring($SourceRoot.Length).TrimStart('\', '/')
    $destination = Join-Path $DestinationRoot $relative
    $destinationDir = Split-Path -Parent $destination
    if (-not (Test-Path -LiteralPath $destinationDir)) {
      New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
  }
}

if (-not (Test-Path -LiteralPath $CcxPath)) {
  throw "CCX not found: $CcxPath"
}

$tempRoot = Join-Path $env:TEMP ("psd_export_sync_" + [guid]::NewGuid().ToString("N"))
$payloadRoot = Join-Path $tempRoot "payload"

try {
  New-Item -ItemType Directory -Path $payloadRoot -Force | Out-Null
  $zipPath = Join-Path $tempRoot "payload.zip"
  Copy-Item -LiteralPath $CcxPath -Destination $zipPath -Force
  Expand-Archive -LiteralPath $zipPath -DestinationPath $payloadRoot -Force

  $manifestPath = Join-Path $payloadRoot "manifest.json"
  if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "manifest.json not found inside CCX: $CcxPath"
  }

  $manifestText = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($PluginId)) {
    $PluginId = Get-ManifestField -ManifestText $manifestText -Field "id"
  }
  $version = Get-ManifestField -ManifestText $manifestText -Field "version"
  if ([string]::IsNullOrWhiteSpace($PluginId)) {
    throw "Plugin id not found in manifest."
  }

  $storageRoot = Join-Path $env:APPDATA "Adobe\UXP\PluginsStorage"
  $synced = 0
  foreach ($product in @("PHSP", "PHSPBETA")) {
    $productRoot = Join-Path $storageRoot $product
    if (-not (Test-Path -LiteralPath $productRoot)) { continue }

    Get-ChildItem -LiteralPath $productRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
      foreach ($bucket in @("Developer", "External")) {
        $target = Join-Path $_.FullName (Join-Path $bucket $PluginId)
        if (-not (Test-Path -LiteralPath $target)) { continue }

        Copy-TreeFiles -SourceRoot $payloadRoot -DestinationRoot $target
        $targetManifest = Join-Path $target "manifest.json"
        $targetVersion = ""
        if (Test-Path -LiteralPath $targetManifest) {
          $targetVersion = Get-ManifestField -ManifestText (Get-Content -LiteralPath $targetManifest -Raw -Encoding UTF8) -Field "version"
        }
        Write-Host "Synced UXP cache: $target ($targetVersion)"
        $script:synced += 1
      }
    }
  }

  if ($synced -eq 0) {
    Write-Host "No existing Photoshop UXP cache folder was found for $PluginId. Restart Photoshop after UPIA install."
  } else {
    Write-Host "UXP cache sync completed for $PluginId $version. Restart Photoshop before opening the panel."
  }
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
