@echo off
setlocal
cd /d "%~dp0"
title Analitika 1.1.0 - Build Release

echo =====================================================
echo   Analitika 1.1.0 - Windows build and installer
echo =====================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo ERROR: Python 3.12+ was not found.
    echo Install Python from https://www.python.org/downloads/windows/
    echo Enable "Add python.exe to PATH" during installation.
    pause
    exit /b 1
  )
  set "PY=python"
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/5] Creating virtual environment...
  %PY% -m venv .venv || goto :failed
)

set "VPY=.venv\Scripts\python.exe"
set "PIP=.venv\Scripts\pip.exe"

echo [2/5] Installing dependencies...
"%VPY%" -m pip install --upgrade pip || goto :failed
"%PIP%" install -r requirements.txt || goto :failed

echo [3/5] Validating source...
"%VPY%" -m py_compile Analitika.py src\report.py src\updater.py || goto :failed

echo [4/5] Building application...
"%VPY%" -m PyInstaller --noconfirm --clean Analitika.spec || goto :failed

echo [5/5] Building installer...
set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo Inno Setup 6 is not installed.
  echo The portable application is ready in dist\Analitika\
  echo Install Inno Setup from https://jrsoftware.org/isdl.php and rerun this file.
  pause
  exit /b 0
)
"%ISCC%" installer_inno_setup.iss || goto :failed

echo.
echo SUCCESS: installer_output\Analitika_Setup_1.1.0.exe
pause
exit /b 0

:failed
echo.
echo BUILD FAILED. Review the error above.
pause
exit /b 1
