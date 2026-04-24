@echo off
setlocal
set "SCRIPT=%~dp0update_from_git.ps1"
if not exist "%SCRIPT%" (
  echo Missing updater script: "%SCRIPT%"
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
pause
