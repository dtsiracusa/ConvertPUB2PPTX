@echo off
cd /d "%~dp0"
setlocal enabledelayedexpansion
title Publisher to PowerPoint Converter Tool

:: ============================================================
:: run_migration.bat
:: Verifies companion files then runs test_convert2pptx.py
:: Place this file in the same folder as the Python script.
:: ============================================================

REM ANSI color codes (Windows 10+)
for /f %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
set "GREEN=!ESC![32m"
set "RED=!ESC![31m"
set "RESET=!ESC![0m"

:: Resolve the directory this .bat file lives in
set "SCRIPT_DIR=%~dp0"

:: Resolve Python executable
set "PYTHON_EXE=python"

:: Define expected files
set "PYTHON_SCRIPT=%SCRIPT_DIR%test_convert2pptx.py"
set "TEMPLATE_FILE=%SCRIPT_DIR%Pub-Pow Blank Template.potx"
set "CSV_FILE=%SCRIPT_DIR%test_batch.csv"

:: ============================================================
:: PRE-LAUNCH CHECKS — Verify companion files
:: ============================================================
echo.
echo ======================================================================
echo       Publisher to PowerPoint Converter Tool - Pre-Launch Checks
echo ======================================================================
echo.

set "PASS_COUNT=0"
set "FAIL_COUNT=0"

:: -- test_convert2pptx.py --
if exist "%PYTHON_SCRIPT%" (
    echo   !GREEN![PASS]!RESET!  test_convert2pptx.py
    set /a PASS_COUNT+=1
) else (
    echo   !RED![FAIL]!RESET!  test_convert2pptx.py
    set /a FAIL_COUNT+=1
)

:: -- Pub-Pow Blank Template.potx --
if exist "%TEMPLATE_FILE%" (
    echo   !GREEN![PASS]!RESET!  Pub-Pow Blank Template.potx
    set /a PASS_COUNT+=1
) else (
    echo   !RED![FAIL]!RESET!  Pub-Pow Blank Template.potx
    set /a FAIL_COUNT+=1
)

:: -- test_batch.csv --
if exist "%CSV_FILE%" (
    echo   !GREEN![PASS]!RESET!  test_batch.csv
    set /a PASS_COUNT+=1
) else (
    echo   !RED![FAIL]!RESET!  test_batch.csv
    set /a FAIL_COUNT+=1
)

:: ============================================================
:: SUMMARY
:: ============================================================
echo.
echo ======================================================================
echo       Summary:  !PASS_COUNT! Passed  /  !FAIL_COUNT! Failed
echo ======================================================================
echo.

:: Abort if any file is missing
if !FAIL_COUNT! GTR 0 (
    echo  One or more required files are missing from:
    echo    !SCRIPT_DIR!
    echo.
    echo  Please ensure all files are present and try again.
    echo.
    goto :eof_pause
)

:: ============================================================
:: LAUNCH CONVERSION SCRIPT
:: ============================================================
echo  All checks passed. Launching conversion script...
echo.
echo ============================================================
echo  Script : test_convert2pptx.py
echo  Folder : %SCRIPT_DIR%
echo ============================================================
echo.

%PYTHON_EXE% "%PYTHON_SCRIPT%"
set "SCRIPT_EXIT=%errorlevel%"

if !SCRIPT_EXIT! neq 0 (
    echo.
    echo  [ERROR] Conversion script failed with exit code !SCRIPT_EXIT!.
    echo          Please review any errors above.
    echo.
    goto :eof_pause
)

:: ============================================================
:: DONE
:: ============================================================
echo.
echo ======================================================================
echo       Conversion script has finished successfully.
echo ======================================================================
echo.
:eof_pause
echo.
pause
exit /b 0

