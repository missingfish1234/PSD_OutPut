param(
  [string]$RepoUrl = "https://github.com/missingfish1234/PSD_OutPut.git",
  [string]$Branch = "main",
  [string]$Message = ""
)

$ErrorActionPreference = "Stop"

function Get-PluginRoot {
  $scriptRoot = Split-Path -Parent $MyInvocation.ScriptName
  return [System.IO.Path]::GetFullPath((Join-Path $scriptRoot ".."))
}

function Ensure-Git {
  try {
    git --version | Out-Null
  } catch {
    throw "Git was not found. Install Git for Windows first."
  }
}

function Ensure-IgnoreFile {
  param([string]$Root)
  $ignorePath = Join-Path $Root ".gitignore"
  if (Test-Path -LiteralPath $ignorePath) {
    return
  }

  @"
launcher/dist/
_inspect_*/
_inspect_unity_*/
node_modules/
.DS_Store
Thumbs.db
*.log
.vs/
.vscode/
"@ | Set-Content -LiteralPath $ignorePath -Encoding ASCII
}

function Invoke-Git {
  param([string[]]$Arguments)
  & git @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "git $($Arguments -join ' ') failed with exit code $LASTEXITCODE"
  }
}

$root = Get-PluginRoot
Ensure-Git
Ensure-IgnoreFile -Root $root

Push-Location $root
try {
  if (-not (Test-Path -LiteralPath (Join-Path $root ".git"))) {
    Invoke-Git @("init")
  }

  Invoke-Git @("branch", "-M", $Branch)

  $remoteUrl = ""
  try {
    $remoteUrl = (& git remote get-url origin 2>$null)
  } catch {
    $remoteUrl = ""
  }

  if ([string]::IsNullOrWhiteSpace($remoteUrl)) {
    Invoke-Git @("remote", "add", "origin", $RepoUrl)
  } elseif ($remoteUrl.Trim() -ne $RepoUrl) {
    Invoke-Git @("remote", "set-url", "origin", $RepoUrl)
  }

  Invoke-Git @("add", "-A")
  $status = (& git status --short)
  if (-not $status) {
    Write-Host "No local changes to publish."
  } else {
    if ([string]::IsNullOrWhiteSpace($Message)) {
      $manifestPath = Join-Path $root "manifest.json"
      $version = "unknown"
      if (Test-Path -LiteralPath $manifestPath) {
        try {
          $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
          if ($manifest.version) { $version = $manifest.version }
        } catch {
          $version = "unknown"
        }
      }
      $Message = "Publish PSD Export Pipeline $version"
    }
    Invoke-Git @("commit", "-m", $Message)
  }

  Invoke-Git @("push", "-u", "origin", $Branch)
  Write-Host "Published to $RepoUrl ($Branch)."
} finally {
  Pop-Location
}
