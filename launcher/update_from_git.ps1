param(
  [string]$RepoUrl = "",
  [string]$Ref = "",
  [switch]$Install,
  [switch]$BuildCcx,
  [switch]$NonInteractive
)

$ErrorActionPreference = "Stop"
$DEFAULT_REPO_URL = "https://github.com/missingfish1234/PSD_OutPut.git"
$DEFAULT_REF = "main"

function Get-ScriptRoot {
  return Split-Path -Parent $MyInvocation.ScriptName
}

function Get-PluginRoot {
  return [System.IO.Path]::GetFullPath((Join-Path (Get-ScriptRoot) ".."))
}

function Get-StateRoot {
  $root = Join-Path $env:APPDATA "PSDExportPipeline"
  New-Item -ItemType Directory -Path $root -Force | Out-Null
  return $root
}

function Get-ConfigPath {
  return Join-Path (Get-StateRoot) "updater.json"
}

function Read-Config {
  $path = Get-ConfigPath
  if (-not (Test-Path -LiteralPath $path)) {
    return [pscustomobject]@{
      repoUrl = $DEFAULT_REPO_URL
      ref = $DEFAULT_REF
      worktree = (Join-Path (Get-StateRoot) "repo")
    }
  }

  try {
    $config = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $config.worktree) {
      $config | Add-Member -NotePropertyName worktree -NotePropertyValue (Join-Path (Get-StateRoot) "repo")
    }
    if (-not $config.repoUrl) {
      $config | Add-Member -NotePropertyName repoUrl -NotePropertyValue $DEFAULT_REPO_URL
    }
    if (-not $config.ref) {
      $config.ref = $DEFAULT_REF
    }
    return $config
  } catch {
    return [pscustomobject]@{
      repoUrl = $DEFAULT_REPO_URL
      ref = $DEFAULT_REF
      worktree = (Join-Path (Get-StateRoot) "repo")
    }
  }
}

function Save-Config {
  param([object]$Config)
  $Config | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Get-ConfigPath) -Encoding UTF8
}

function Test-GitAvailable {
  try {
    git --version | Out-Null
    return $true
  } catch {
    return $false
  }
}

function Test-IsGitWorktree {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath (Join-Path $Path ".git"))) {
    return $false
  }
  return $true
}

function Invoke-Git {
  param(
    [string]$WorkingDirectory,
    [string[]]$Arguments
  )

  Push-Location $WorkingDirectory
  try {
    & git @Arguments
    if ($LASTEXITCODE -ne 0) {
      throw "git $($Arguments -join ' ') failed with exit code $LASTEXITCODE"
    }
  } finally {
    Pop-Location
  }
}

function Resolve-WorkingCopy {
  param([object]$Config)

  $localRoot = Get-PluginRoot
  if (Test-IsGitWorktree -Path $localRoot) {
    return $localRoot
  }

  if ([string]::IsNullOrWhiteSpace($Config.repoUrl)) {
    return ""
  }

  return $Config.worktree
}

function Ensure-WorkingCopy {
  param([object]$Config)

  if (-not (Test-GitAvailable)) {
    throw "Git was not found. Install Git for Windows first."
  }

  $localRoot = Get-PluginRoot
  if (Test-IsGitWorktree -Path $localRoot) {
    return $localRoot
  }

  if ([string]::IsNullOrWhiteSpace($Config.repoUrl)) {
    throw "Repo URL is not configured."
  }

  $worktree = $Config.worktree
  if (-not (Test-Path -LiteralPath $worktree)) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $worktree) -Force | Out-Null
    & git clone $Config.repoUrl $worktree
    if ($LASTEXITCODE -ne 0) {
      throw "git clone failed with exit code $LASTEXITCODE"
    }
  }

  if (-not (Test-IsGitWorktree -Path $worktree)) {
    throw "Configured worktree is not a Git repository: $worktree"
  }

  return $worktree
}

function Get-CurrentVersion {
  param([string]$Root)
  $manifestPath = Join-Path $Root "manifest.json"
  if (-not (Test-Path -LiteralPath $manifestPath)) {
    return "(manifest missing)"
  }
  try {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    return $manifest.version
  } catch {
    return "(manifest unreadable)"
  }
}

function Show-Status {
  param([object]$Config)
  $worktree = Resolve-WorkingCopy -Config $Config
  Write-Host ""
  Write-Host "PSD Export Pipeline Updater"
  Write-Host "Repo URL : $($Config.repoUrl)"
  Write-Host "Ref      : $($Config.ref)"
  Write-Host "Worktree : $worktree"
  if ($worktree -and (Test-IsGitWorktree -Path $worktree)) {
    Push-Location $worktree
    try {
      Write-Host "Version  : $(Get-CurrentVersion -Root $worktree)"
      Write-Host "Branch   : $((& git rev-parse --abbrev-ref HEAD) 2>$null)"
      Write-Host "Commit   : $((& git rev-parse --short HEAD) 2>$null)"
      Write-Host "Changes  :"
      & git status --short
    } finally {
      Pop-Location
    }
  } else {
    Write-Host "Status   : no Git worktree yet"
  }
  Write-Host ""
}

function Set-RepoUrlInteractive {
  param([object]$Config)
  $value = Read-Host "Git repo URL"
  if (-not [string]::IsNullOrWhiteSpace($value)) {
    $Config.repoUrl = $value.Trim()
    Save-Config -Config $Config
  }
}

function Select-RefInteractive {
  param([object]$Config)

  if ([string]::IsNullOrWhiteSpace($Config.repoUrl)) {
    Write-Host "Set repo URL first."
    return
  }

  Write-Host "Fetching remote branches/tags..."
  $rows = @()
  $remote = & git ls-remote --heads --tags $Config.repoUrl
  if ($LASTEXITCODE -ne 0) {
    Write-Host "Unable to list remote refs. You can still type one manually."
  } else {
    foreach ($line in $remote) {
      $parts = $line -split "\s+"
      if ($parts.Count -lt 2) { continue }
      $name = $parts[1] -replace "^refs/heads/", "" -replace "^refs/tags/", "tags/"
      if ($name.EndsWith("^{}")) { continue }
      $rows += $name
    }
  }

  $rows = $rows | Select-Object -Unique | Select-Object -First 30
  for ($i = 0; $i -lt $rows.Count; $i += 1) {
    Write-Host ("{0,2}. {1}" -f ($i + 1), $rows[$i])
  }
  Write-Host " 0. Type manually"
  $choice = Read-Host "Choose ref"
  $number = 0
  if ([int]::TryParse($choice, [ref]$number) -and $number -gt 0 -and $number -le $rows.Count) {
    $Config.ref = $rows[$number - 1]
  } else {
    $manual = Read-Host "Branch or tag name"
    if (-not [string]::IsNullOrWhiteSpace($manual)) {
      $Config.ref = $manual.Trim()
    }
  }
  Save-Config -Config $Config
}

function Update-WorkingCopy {
  param([object]$Config)

  $worktree = Ensure-WorkingCopy -Config $Config
  $refName = if ([string]::IsNullOrWhiteSpace($Config.ref)) { "main" } else { $Config.ref.Trim() }
  Push-Location $worktree
  try {
    & git fetch --all --tags --prune
    if ($LASTEXITCODE -ne 0) { throw "git fetch failed" }

    if ($refName.StartsWith("tags/")) {
      $tagName = $refName.Substring(5)
      & git checkout "refs/tags/$tagName"
      if ($LASTEXITCODE -ne 0) { throw "git checkout tag failed" }
    } else {
      & git checkout $refName
      if ($LASTEXITCODE -ne 0) {
        & git checkout -b $refName "origin/$refName"
        if ($LASTEXITCODE -ne 0) { throw "git checkout branch failed" }
      }

      & git pull --ff-only
      if ($LASTEXITCODE -ne 0) { throw "git pull --ff-only failed. Commit/stash local changes, then try again." }
    }

    Write-Host "Updated to version $(Get-CurrentVersion -Root $worktree), commit $((& git rev-parse --short HEAD) 2>$null)"
  } finally {
    Pop-Location
  }

  return $worktree
}

function Build-Ccx {
  param([string]$Root)
  $script = Join-Path $Root "launcher\build_ccx_package.ps1"
  if (-not (Test-Path -LiteralPath $script)) {
    throw "CCX build script not found: $script"
  }
  & powershell -NoProfile -ExecutionPolicy Bypass -File $script
  if ($LASTEXITCODE -ne 0) {
    throw "CCX build failed with exit code $LASTEXITCODE"
  }
}

function Get-LatestCcx {
  param([string]$Root)
  $dist = Join-Path $Root "launcher\dist"
  if (-not (Test-Path -LiteralPath $dist)) {
    return ""
  }
  $file = Get-ChildItem -LiteralPath $dist -Filter "PSDExportPipeline_*.ccx" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if ($file) { return $file.FullName }
  return ""
}

function Install-Ccx {
  param([string]$Root)
  $ccx = Get-LatestCcx -Root $Root
  if ([string]::IsNullOrWhiteSpace($ccx)) {
    throw "No CCX package found. Build CCX first."
  }
  $cmd = Join-Path (Split-Path -Parent $ccx) ("Install_" + [System.IO.Path]::GetFileNameWithoutExtension($ccx) + ".cmd")
  if (Test-Path -LiteralPath $cmd) {
    & cmd.exe /c "`"$cmd`" `"$ccx`""
    return
  }
  throw "Install helper not found: $cmd"
}

function Run-Interactive {
  $config = Read-Config
  while ($true) {
    Show-Status -Config $config
    Write-Host "1. Set Git repo URL"
    Write-Host "2. Select branch/tag"
    Write-Host "3. Update from Git"
    Write-Host "4. Build CCX package"
    Write-Host "5. Install latest CCX"
    Write-Host "6. Update + build CCX + install"
    Write-Host "0. Exit"
    $choice = Read-Host "Choose"
    switch ($choice) {
      "1" { Set-RepoUrlInteractive -Config $config; $config = Read-Config }
      "2" { Select-RefInteractive -Config $config; $config = Read-Config }
      "3" { Update-WorkingCopy -Config $config | Out-Null }
      "4" { Build-Ccx -Root (Ensure-WorkingCopy -Config $config) }
      "5" { Install-Ccx -Root (Ensure-WorkingCopy -Config $config) }
      "6" {
        $root = Update-WorkingCopy -Config $config
        Build-Ccx -Root $root
        Install-Ccx -Root $root
      }
      "0" { return }
      default { Write-Host "Unknown choice." }
    }
  }
}

$config = Read-Config
if (-not [string]::IsNullOrWhiteSpace($RepoUrl)) {
  $config.repoUrl = $RepoUrl
}
if (-not [string]::IsNullOrWhiteSpace($Ref)) {
  $config.ref = $Ref
}
Save-Config -Config $config

if ($NonInteractive) {
  $root = Update-WorkingCopy -Config $config
  if ($BuildCcx -or $Install) {
    Build-Ccx -Root $root
  }
  if ($Install) {
    Install-Ccx -Root $root
  }
  exit 0
}

Run-Interactive
