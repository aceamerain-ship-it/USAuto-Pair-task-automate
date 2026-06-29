@echo off
setlocal EnableDelayedExpansion

:: ================================================================
::  USAUTO Pair Project Automation  —  Option A (Bundled Python)
::  Self-contained: downloads portable Python on first run only.
::  After setup, runs fully offline with no system dependencies.
::
::  Usage:  USAUTO_Pair_A.bat  [input.xlsx]  [output.xlsx]
::          Defaults: input.xlsx -> output.xlsx (same folder)
:: ================================================================

title USAUTO Pair Project Automation

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "PY_DIR=%SCRIPT_DIR%\python"
set "PY_EXE=%PY_DIR%\python.exe"
set "PY_VER=3.12.10"
set "PY_URL=https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip"
set "PY_ZIP=%TEMP%\python_embed.zip"
set "PIP_URL=https://bootstrap.pypa.io/get-pip.py"
set "PIP_SCRIPT=%TEMP%\get-pip.py"
set "PS_SETUP=%TEMP%\usauto_setup.ps1"

:: ── Arguments ────────────────────────────────────────────────────
if "%~1"=="" (
    set "INPUT_FILE=%SCRIPT_DIR%\input.xlsx"
) else (
    set "INPUT_FILE=%~f1"
)
if "%~2"=="" (
    set "OUTPUT_FILE=%SCRIPT_DIR%\output.xlsx"
) else (
    set "OUTPUT_FILE=%~f2"
)

:: ── Banner ───────────────────────────────────────────────────────
echo.
echo   ============================================================
echo    USAUTO Pair Project Automation  -  Bundled Python Edition
echo   ============================================================
echo.

:: ── Check input file ─────────────────────────────────────────────
if not exist "%INPUT_FILE%" (
    echo   ERROR: Input file not found:
    echo          %INPUT_FILE%
    echo.
    pause & exit /b 1
)

:: ── Check Python worker script ───────────────────────────────────
if not exist "%SCRIPT_DIR%\usauto_pair.py" (
    echo   ERROR: usauto_pair.py not found in:
    echo          %SCRIPT_DIR%
    echo.
    pause & exit /b 1
)

:: ── Setup: download portable Python if not present ───────────────
if exist "%PY_EXE%" goto :PYTHON_READY

echo   First-run setup: installing bundled Python %PY_VER%...
echo   (Internet required this one time only)
echo.

:: Check PowerShell
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo   ERROR: PowerShell not found.
    pause & exit /b 1
)

:: Create python directory
if not exist "%PY_DIR%" mkdir "%PY_DIR%"

:: Build the PS1 file line by line via PowerShell Add-Content (no redirect needed)
:: This avoids ALL cmd.exe redirect + spaces-in-path issues
del /q "%PS_SETUP%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '$ErrorActionPreference = ''Stop'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''  [1/5] Downloading Python embeddable package...'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'try {'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '    (New-Object Net.WebClient).DownloadFile(''%PY_URL%'', ''%PY_ZIP%'')'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '    Write-Host ''         Download OK.'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '} catch { Write-Host (''  ERROR: '' + $_.Exception.Message); exit 1 }'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''  [2/5] Extracting Python...'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Expand-Archive -Path ''%PY_ZIP%'' -DestinationPath ''%PY_DIR%'' -Force'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Remove-Item ''%PY_ZIP%'' -Force -ErrorAction SilentlyContinue'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''         Extracted OK.'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''  [3/5] Configuring Python path...'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '$pth = Get-ChildItem ''%PY_DIR%\python*._pth'' | Select-Object -First 1'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'if ($pth) { $c = Get-Content $pth.FullName; $c = $c -replace ''#import site'', ''import site''; Set-Content $pth.FullName $c; Write-Host (''         Configured: '' + $pth.Name) } else { Write-Host ''         WARNING: ._pth not found'' }'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''  [4/5] Installing pip...'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '(New-Object Net.WebClient).DownloadFile(''%PIP_URL%'', ''%PIP_SCRIPT%'')'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '& ''%PY_EXE%'' ''%PIP_SCRIPT%'' --no-warn-script-location -q'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'if ($LASTEXITCODE -ne 0) { Write-Host ''  ERROR: pip install failed.''; exit 1 }'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Remove-Item ''%PIP_SCRIPT%'' -Force -ErrorAction SilentlyContinue'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''         pip installed.'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''  [5/5] Installing openpyxl...'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' '& ''%PY_EXE%'' -m pip install openpyxl --no-warn-script-location -q'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'if ($LASTEXITCODE -ne 0) { Write-Host ''  ERROR: openpyxl install failed.''; exit 1 }'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''         openpyxl installed.'''"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Content '%PS_SETUP%' 'Write-Host ''  Setup complete!'''"

:: Verify PS1 was created
if not exist "%PS_SETUP%" (
    echo   ERROR: Failed to create setup script at %PS_SETUP%
    pause & exit /b 1
)

:: Run the setup
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SETUP%"
set "PS_ERR=%errorlevel%"
del /q "%PS_SETUP%" >nul 2>&1

if %PS_ERR% neq 0 (
    echo.
    echo   ERROR: Python setup failed. See messages above.
    echo   Delete the "%PY_DIR%" folder and re-run to reinstall.
    echo.
    pause & exit /b 1
)

echo.

:PYTHON_READY

:: ── Verify Python + openpyxl ─────────────────────────────────────
"%PY_EXE%" -c "import openpyxl; print('  Python ready  -  openpyxl ' + openpyxl.__version__)"
if %errorlevel% neq 0 (
    echo.
    echo   ERROR: Python check failed.
    echo   Delete the "%PY_DIR%" folder and re-run to reinstall.
    echo.
    pause & exit /b 1
)

:: ── Run ──────────────────────────────────────────────────────────
echo.
echo   Input  : %INPUT_FILE%
echo   Output : %OUTPUT_FILE%
echo.

"%PY_EXE%" "%SCRIPT_DIR%\usauto_pair.py" "%INPUT_FILE%" "%OUTPUT_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo   !! Process failed. See errors above. !!
    echo.
    pause & exit /b 1
)

echo.
echo   Press any key to exit...
pause >nul
endlocal
