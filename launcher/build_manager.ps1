$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceFile = Join-Path $scriptRoot "PsdExportManager.cs"
$distDir = Join-Path $scriptRoot "dist"
$outputExe = Join-Path $distDir "PSDExportManager.exe"

if (-not (Test-Path -LiteralPath $sourceFile)) {
  throw "Source file not found: $sourceFile"
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

if (Test-Path -LiteralPath $outputExe) {
  Remove-Item -LiteralPath $outputExe -Force
}

$sourceCode = Get-Content -LiteralPath $sourceFile -Raw -Encoding UTF8
Add-Type `
  -TypeDefinition $sourceCode `
  -Language CSharp `
  -ReferencedAssemblies @("System.Windows.Forms.dll", "System.Drawing.dll", "System.IO.Compression.dll", "System.IO.Compression.FileSystem.dll") `
  -OutputAssembly $outputExe `
  -OutputType WindowsApplication

Write-Host "Manager build succeeded:"
Write-Host "  $outputExe"
