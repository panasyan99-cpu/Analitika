@echo off
setlocal
cd /d "%~dp0"
title Analitika 1.1.0 Builder
if not exist .venv (
  py -3 -m venv .venv || goto :fail
)
call .venv\Scripts\activate.bat || goto :fail
python -m pip install --upgrade pip || goto :fail
pip install -r requirements.txt || goto :fail
pyinstaller --noconfirm --clean Analitika.spec || goto :fail
where iscc >nul 2>nul
if errorlevel 1 (
  echo Inno Setup not found. Install it from https://jrsoftware.org/isdl.php
  pause
  exit /b 2
)
iscc installer_inno_setup.iss || goto :fail
echo.
echo READY: installer_output\Analitika_Setup_1.1.0.exe
pause
exit /b 0
:fail
echo BUILD FAILED
pause
exit /b 1
