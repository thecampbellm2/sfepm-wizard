@echo off
title SFEPM Takeoff Wizard — Build & Deploy
setlocal enabledelayedexpansion

echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║   SFEPM Takeoff Wizard — Build ^& Deploy          ║
echo  ╚══════════════════════════════════════════════════╝
echo.

:: ── Find Python ──────────────────────────────────────────────────────────────
set PYTHON=
py --version > nul 2>&1
if not errorlevel 1 ( set PYTHON=py & goto :found_python )
python --version > nul 2>&1
if not errorlevel 1 ( set PYTHON=python & goto :found_python )
echo  ERROR: Python not found.
pause & exit /b 1

:found_python
echo  Python found via: %PYTHON%
echo.

:: ── Auto-detect latest wizard script ─────────────────────────────────────────
set SCRIPT=
for /f "delims=" %%i in ('dir /b /o:-n "SFEPM_Takeoff_Wizard_v*.py" 2^>nul') do (
    if not defined SCRIPT set SCRIPT=%%i
)
if not defined SCRIPT (
    echo  ERROR: No SFEPM_Takeoff_Wizard_v*.py found in this folder.
    pause & exit /b 1
)
echo  Script found: %SCRIPT%

:: ── Extract version from script ──────────────────────────────────────────────
set VERSION=
for /f "tokens=*" %%a in ('powershell -NoProfile -Command "(Select-String -Path '%SCRIPT%' -Pattern 'VERSION = ""v([\d.]+)""').Matches.Groups[1].Value"') do set VERSION=%%a
echo  Version: %VERSION%
echo.

:: ── Write version.txt ────────────────────────────────────────────────────────
echo %VERSION%> version.txt

:: ── Install compatible PyInstaller ───────────────────────────────────────────
echo  Checking PyInstaller...
%PYTHON% -m pip install "pyinstaller>=6.10" --upgrade --quiet

:: ── Clean previous build ─────────────────────────────────────────────────────
echo  Cleaning previous build...
if exist "dist\SFEPM_Takeoff_Wizard.exe" del /f /q "dist\SFEPM_Takeoff_Wizard.exe"
if exist "build" rmdir /s /q "build"
if exist "SFEPM_Takeoff_Wizard.spec" del /f /q "SFEPM_Takeoff_Wizard.spec"

:: ── Build ────────────────────────────────────────────────────────────────────
echo  Building exe — this takes about 30 seconds...
echo.
%PYTHON% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "SFEPM_Takeoff_Wizard" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.utils ^
    --hidden-import openpyxl.worksheet.datavalidation ^
    --collect-all tkinter ^
    %SCRIPT%

if not exist "dist\SFEPM_Takeoff_Wizard.exe" (
    echo.
    echo  BUILD FAILED — check the output above for errors.
    pause & exit /b 1
)

echo.
echo  Build successful.
echo.

:: ── Deploy to GitHub ─────────────────────────────────────────────────────────
echo  Deploying to GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1" -Script "%SCRIPT%" -Version "%VERSION%"

echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║   ALL DONE                                       ║
echo  ║                                                  ║
echo  ║   Remember to update changelog.txt before        ║
echo  ║   the next release                               ║
echo  ╚══════════════════════════════════════════════════╝
echo.
pause
