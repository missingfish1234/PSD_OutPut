@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%build_launcher.ps1"
if errorlevel 1 (
  echo Build failed.
  exit /b 1
)
"%SCRIPT_DIR%dist\PSDExportLauncher.exe" install-run --plugin-root "%SCRIPT_DIR%.."
endlocal
