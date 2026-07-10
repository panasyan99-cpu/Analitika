@echo off
setlocal
cd /d "%~dp0"
if not exist ".github\workflows" mkdir ".github\workflows"
copy /Y "build-installer.yml" ".github\workflows\build-installer.yml" >nul
if errorlevel 1 (
  echo ERROR: Could not install workflow.
  pause
  exit /b 1
)
echo Workflow installed successfully:
echo .github\workflows\build-installer.yml
pause
