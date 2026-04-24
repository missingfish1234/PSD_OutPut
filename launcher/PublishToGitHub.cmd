@echo off
setlocal
set "SCRIPT=%~dp0publish_to_github.ps1"
if not exist "%SCRIPT%" (
  echo Missing publish script: "%SCRIPT%"
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
pause
