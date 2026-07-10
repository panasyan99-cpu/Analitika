@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title Analitika - EXE build

set "LOG=%~dp0build_log.txt"
echo =========================================
echo   ANALITIKA v1.0 - EXE BUILD
echo =========================================
echo Project folder: %CD%
echo Log file: %LOG%
echo.

echo ==== BUILD START %DATE% %TIME% ==== > "%LOG%"
echo Project: %CD% >> "%LOG%"

set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if "%PY%"=="" (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)

if "%PY%"=="" (
    echo ERROR: Python was not found.
    echo Install Python 3.11 or 3.12 and enable "Add Python to PATH".
    echo Download: https://www.python.org/downloads/windows/
    echo Python was not found. >> "%LOG%"
    goto FAIL
)

echo Python command: %PY%
%PY% --version
%PY% --version >> "%LOG%" 2>&1

if exist ".venv" (
    echo Removing old .venv...
    rmdir /s /q ".venv" >> "%LOG%" 2>&1
)

echo.
echo [1/5] Creating local virtual environment...
%PY% -m venv .venv >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL

set "VPY=%~dp0.venv\Scripts\python.exe"
if not exist "%VPY%" (
    echo ERROR: venv python not found: %VPY%
    echo venv python missing >> "%LOG%"
    goto FAIL
)

echo.
echo [2/5] Installing dependencies...
"%VPY%" -m pip install --upgrade pip >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL
"%VPY%" -m pip install -r requirements.txt >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL

echo.
echo [3/5] Checking Python files...
"%VPY%" -m py_compile Analitika.py src\report.py >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL

echo.
echo [4/5] Building application with PyInstaller...
"%VPY%" -m PyInstaller --noconfirm --clean Analitika.spec >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL

if not exist "dist\Analitika" (
    echo ERROR: dist\Analitika was not created.
    echo dist folder missing >> "%LOG%"
    goto FAIL
)

echo.
echo [5/5] Preparing release folder...
if exist "release" rmdir /s /q "release" >> "%LOG%" 2>&1
mkdir "release\Analitika\Reports" >> "%LOG%" 2>&1
mkdir "release\Analitika\Output" >> "%LOG%" 2>&1
mkdir "release\Analitika\logs" >> "%LOG%" 2>&1
xcopy /E /I /Y "dist\Analitika" "release\Analitika\app" >> "%LOG%" 2>&1
copy /Y "README_FOR_USERS.txt" "release\Analitika\README.txt" >> "%LOG%" 2>&1
copy /Y "assets\logo.png" "release\Analitika\logo.png" >> "%LOG%" 2>&1
copy /Y "run_analitika.bat" "release\Analitika\run_analitika.bat" >> "%LOG%" 2>&1
copy /Y "installer_inno_setup.iss" "release\Analitika\installer_inno_setup.iss" >> "%LOG%" 2>&1

echo.
echo =========================================
echo BUILD OK.
echo Release folder:
echo   release\Analitika
echo Start file:
echo   release\Analitika\run_analitika.bat
echo =========================================
echo ==== BUILD OK %DATE% %TIME% ==== >> "%LOG%"
goto END

:FAIL
echo.
echo =========================================
echo BUILD FAILED.
echo Send me this file: build_log.txt
echo =========================================
echo ==== BUILD FAILED %DATE% %TIME% ==== >> "%LOG%"

:END
echo.
echo Press any key to close this window...
pause >nul
endlocal
