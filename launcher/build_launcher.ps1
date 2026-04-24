$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceFile = Join-Path $scriptRoot "PsdExportLauncher.cs"
$distDir = Join-Path $scriptRoot "dist"
$outputExe = Join-Path $distDir "PSDExportLauncher.exe"

if (-not (Test-Path -LiteralPath $sourceFile)) {
  throw "Source file not found: $sourceFile"
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

if (Test-Path -LiteralPath $outputExe) {
  Remove-Item -LiteralPath $outputExe -Force
}

$sourceCode = Get-Content -LiteralPath $sourceFile -Raw -Encoding UTF8
Add-Type -TypeDefinition $sourceCode -Language CSharp -OutputAssembly $outputExe -OutputType ConsoleApplication

Write-Host "Build succeeded:"
Write-Host "  $outputExe"
