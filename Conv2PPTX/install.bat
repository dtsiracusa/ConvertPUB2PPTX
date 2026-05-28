@echo off
set "BANG=!"
setlocal EnableExtensions EnableDelayedExpansion
title Install ^& Setup - Publisher to PowerPoint Converter

REM ANSI color codes (Windows 10+)
for /f %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
set "GREEN=!ESC![32m"
set "RED=!ESC![31m"
set "RESET=!ESC![0m"

REM ============================================================
REM CONFIG
REM ============================================================
set "SCRIPT_NAME=manual_convert2pptx.py"
REM Optional: if you distribute a default template file alongside the script, set this
REM set "TEMPLATE_NAME=SOP Powerpoint Template.potx"

set "BASE_DIR=%~dp0"
set "BASE_DIR=%BASE_DIR:~0,-1%"

echo ============================================================
echo Install ^& Setup Script - Publisher to PowerPoint Converter
echo  Step 1: Detect Python + pip
echo  Step 2: Verify/Install required Python packages
echo  Step 3: Verify Microsoft Publisher + PowerPoint installed
echo  Step 4: Configure Publisher macro security (Trust Center)
echo  Step 5: Verify script file exists
echo  Step 6: Create CSV of Files for Test
echo  Step 7: Summary
echo ============================================================
echo.

REM Track pass/fail
set "PY_OK=0"
set "PIP_OK=0"
set "WIN32_OK=0"
set "PUB_OK=0"
set "PPT_OK=0"
set "SCRIPT_OK=0"
set "COM_OK=0"
set "VBA_OK=0"
set "GRAB_OK=0"

REM ============================================================
REM STEP 1 — CHECK FOR PYTHON
REM ============================================================
echo [1/7] Checking for Python installation...
echo.

python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo !RED![!BANG!]!RESET! Python was NOT found on this machine.
    echo.
    echo Python is required to run this converter.
    echo Opening the Microsoft Store to download Python...
    echo.
    start "" "ms-windows-store://pdp/?ProductId=9pnrbtzxmb4z"
    echo.
    echo After installing Python, re-run this batch file.
    echo If Python is installed but this still fails, check:
    echo Settings ^> Apps ^> Advanced app settings ^> App execution aliases
    echo and enable Python aliases.
    echo.
    pause
    goto :SUMMARY
) ELSE (
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do set "PYVER=%%v"
    echo !GREEN![OK]!RESET! Python is installed: !PYVER!
    set "PY_OK=1"
)
    
    set "PYTHON_EXE=python"
    for /f "tokens=*" %%E in ('where python 2^>nul') do (
        set "PYTHON_EXE=%%E"
        goto :PY_RESOLVED
    )
    :PY_RESOLVED

echo.

echo Checking for pip...
python -m pip --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 goto :PIP_MISSING

for /f "tokens=*" %%p in ('python -m pip --version 2^>^&1') do set "PIPVER=%%p"
echo !GREEN![OK]!RESET! pip detected: !PIPVER!
set "PIP_OK=1"
goto :PIP_DONE

:PIP_MISSING
echo !RED![!BANG!]!RESET! pip was NOT detected for this Python installation.
echo Attempting to bootstrap pip (ensurepip)...
python -m ensurepip --upgrade >nul 2>&1
python -m pip --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo !RED![ERROR]!RESET! pip is still not available. Reinstall Python with pip enabled.
    echo.
    pause
    goto :SUMMARY
)
for /f "tokens=*" %%p in ('python -m pip --version 2^>^&1') do set "PIPVER=%%p"
echo !GREEN![OK]!RESET! pip is now available: !PIPVER!
set "PIP_OK=1"

:PIP_DONE
echo.

REM ============================================================
REM STEP 2 — CHECK AND INSTALL/UPDATE REQUIRED PACKAGES
REM ============================================================
echo [2/7] Checking required Python libraries...
echo.

echo Checking pywin32 (win32com / pythoncom)...
python -c "import win32com.client, pythoncom" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo !RED![!BANG!]!RESET! pywin32 not found. Installing/Upgrading...
    python -m pip install --upgrade pywin32
    IF %ERRORLEVEL% NEQ 0 (
        echo !RED![ERROR]!RESET! Failed to install pywin32. Check internet access and permissions.
        echo.
        pause
        goto :SUMMARY
    )
    echo !GREEN![OK]!RESET! pywin32 installed.
    echo Running pywin32 post-install script...
    python -m pywin32_postinstall -install >nul 2>&1
    python -c "import win32com.client, pythoncom" >nul 2>&1
    IF %ERRORLEVEL% NEQ 0 (
        echo !RED![ERROR]!RESET! pywin32 import still failing after install.
        echo.
        pause
        goto :SUMMARY
    ) ELSE (
        echo !GREEN![OK]!RESET! pywin32 import verified.
        set "WIN32_OK=1"
    )
) ELSE (
    echo !GREEN![OK]!RESET! pywin32 already installed. Checking for updates...
    python -m pip install --upgrade pywin32 >nul 2>&1
    echo !GREEN![OK]!RESET! pywin32 verified/up to date.
    set "WIN32_OK=1"
)
echo.

REM ============================================================
REM STEP 3 — CHECK FOR MICROSOFT PUBLISHER + POWERPOINT
REM ============================================================
echo [3/7] Checking for Microsoft Publisher and PowerPoint...
echo.

REM --- Detect Publisher EXE ---
set "PUB_EXE="

REM Method 1: App Paths registry (MSI installs)
for %%K in (
    "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSPUB.EXE"
    "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\MSPUB.EXE"
) do (
    if not defined PUB_EXE (
        for /f "tokens=2,*" %%A in ('reg query %%~K /ve 2^>nul ^| find /i "(Default)"') do set "PUB_EXE=%%B"
    )
)

REM Method 2: Common Click-to-Run paths (Microsoft 365 / Office 2019+)
if not defined PUB_EXE (
    for %%P in (
        "%ProgramFiles%\Microsoft Office\root\Office16\MSPUB.EXE"
        "%ProgramFiles(x86)%\Microsoft Office\root\Office16\MSPUB.EXE"
        "%ProgramFiles%\Microsoft Office\Office16\MSPUB.EXE"
        "%ProgramFiles(x86)%\Microsoft Office\Office16\MSPUB.EXE"
    ) do (
        if not defined PUB_EXE (
            if exist "%%~P" set "PUB_EXE=%%~P"
        )
    )
)

REM Method 3: where command (checks PATH)
if not defined PUB_EXE (
    for /f "tokens=*" %%W in ('where MSPUB.EXE 2^>nul') do (
        if not defined PUB_EXE set "PUB_EXE=%%W"
    )
)

if defined PUB_EXE (
    echo !GREEN![OK]!RESET! Publisher found: !PUB_EXE!
    set "PUB_OK=1"
) else (
    echo !RED![!BANG!]!RESET! Publisher executable not detected.
)

REM --- Detect PowerPoint EXE ---
set "PPT_EXE="

REM Method 1: App Paths registry
for %%K in (
    "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"
    "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"
) do (
    if not defined PPT_EXE (
        for /f "tokens=2,*" %%A in ('reg query %%~K /ve 2^>nul ^| find /i "(Default)"') do set "PPT_EXE=%%B"
    )
)

REM Method 2: Common Click-to-Run paths
if not defined PPT_EXE (
    for %%P in (
        "%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE"
        "%ProgramFiles(x86)%\Microsoft Office\root\Office16\POWERPNT.EXE"
        "%ProgramFiles%\Microsoft Office\Office16\POWERPNT.EXE"
        "%ProgramFiles(x86)%\Microsoft Office\Office16\POWERPNT.EXE"
    ) do (
        if not defined PPT_EXE (
            if exist "%%~P" set "PPT_EXE=%%~P"
        )
    )
)

REM Method 3: where command
if not defined PPT_EXE (
    for /f "tokens=*" %%W in ('where POWERPNT.EXE 2^>nul') do (
        if not defined PPT_EXE set "PPT_EXE=%%W"
    )
)

if defined PPT_EXE (
    echo !GREEN![OK]!RESET! PowerPoint found: !PPT_EXE!
    set "PPT_OK=1"
) else (
    echo !RED![!BANG!]!RESET! PowerPoint executable not detected.
)

echo.

REM --- COM automation sanity test with 30-second timeout ---
REM Skip if both EXEs were already found
if "!PUB_OK!"=="1" if "!PPT_OK!"=="1" (
    echo Skipping COM check -- both executables found.
    set "COM_OK=1"
    goto :COM_DONE
)

echo Running COM automation check (Publisher.Application + PowerPoint.Application)...
echo (This may take up to 30 seconds. If a sign-in or activation dialog appears,
echo  complete it so the check can finish.)
echo.

REM Build temp Python script that runs the COM check in a subprocess with timeout
echo import sys, subprocess > "%TEMP%\_com_check.py"
echo inner = "import win32com.client as w; a=w.Dispatch('Publisher.Application'); a.Quit(); b=w.Dispatch('PowerPoint.Application'); b.Quit()" >> "%TEMP%\_com_check.py"
echo try: >> "%TEMP%\_com_check.py"
echo     r = subprocess.run([sys.executable, '-c', inner], timeout=30, capture_output=True) >> "%TEMP%\_com_check.py"
echo     sys.exit(r.returncode) >> "%TEMP%\_com_check.py"
echo except subprocess.TimeoutExpired: >> "%TEMP%\_com_check.py"
echo     sys.exit(2) >> "%TEMP%\_com_check.py"

python "%TEMP%\_com_check.py"
set "COM_RESULT=!ERRORLEVEL!"
del "%TEMP%\_com_check.py" >nul 2>&1

IF "!COM_RESULT!"=="0" (
    echo !GREEN![OK]!RESET! COM automation test passed.
    set "COM_OK=1"
    REM If EXE detection missed them but COM works, mark as passed
    if "!PUB_OK!"=="0" (
        echo !GREEN![OK]!RESET! Publisher confirmed via COM automation.
        set "PUB_OK=1"
    )
    if "!PPT_OK!"=="0" (
        echo !GREEN![OK]!RESET! PowerPoint confirmed via COM automation.
        set "PPT_OK=1"
    )
) ELSE IF "!COM_RESULT!"=="2" (
    echo !RED![!BANG!]!RESET! COM automation test timed out.
    taskkill /F /IM MSPUB.EXE >nul 2>&1
    taskkill /F /IM POWERPNT.EXE >nul 2>&1
    echo.
    echo This usually means Publisher or PowerPoint displayed a dialog
    echo (sign-in, activation, etc.) that blocked automation.
    echo Open Publisher manually, complete any dialogs, then re-run this installer.
) ELSE (
    echo !RED![!BANG!]!RESET! COM automation test failed.
    echo Publisher and/or PowerPoint may not be installed,
    echo or Office COM registration is broken.
)

:COM_DONE
echo.

REM Final warning if anything is still not detected
if "!PUB_OK!"=="0" (
    echo !RED![WARNING]!RESET! Publisher could not be confirmed. The converter requires it.
    echo.
)
if "!PPT_OK!"=="0" (
    echo !RED![WARNING]!RESET! PowerPoint could not be confirmed. The converter requires it.
    echo.
)
if "!COM_OK!"=="0" (
    echo !RED![WARNING]!RESET! The converter relies on Office COM automation.
    echo Install Microsoft Office with BOTH Publisher and PowerPoint.
    echo.
)

REM ============================================================
REM STEP 4 — PUBLISHER TRUST CENTER (MACRO SECURITY)
REM ============================================================
echo [4/7] Checking Publisher macro security (Trust Center)...
echo.
set "REG_POLICY=HKCU\Software\Policies\Microsoft\Office\16.0\Publisher\Security"
set "REG_USER=HKCU\Software\Microsoft\Office\16.0\Publisher\Security"
set "VBA_KEY="
set "VBA_CUR="

REM --- Check policy path first (takes precedence if set by Group Policy) ---
for /f "tokens=3" %%V in ('reg query "!REG_POLICY!" /v VBAWarnings 2^>nul ^| find /i "VBAWarnings"') do (
    set "VBA_CUR=%%V"
    set "VBA_KEY=!REG_POLICY!"
)

REM --- Fall back to user path ---
if not defined VBA_KEY (
    for /f "tokens=3" %%V in ('reg query "!REG_USER!" /v VBAWarnings 2^>nul ^| find /i "VBAWarnings"') do (
        set "VBA_CUR=%%V"
        set "VBA_KEY=!REG_USER!"
    )
)

REM --- If neither exists, default behavior is 2 (Disable with notification); target the user path ---
if not defined VBA_KEY (
    set "VBA_CUR=0x2"
    set "VBA_KEY=!REG_USER!"
)

REM --- Already correct? ---
if "!VBA_CUR!"=="0x4" (
    echo !GREEN![OK]!RESET! VBAWarnings is already set to 4 (Disable all without notification^).
    set "VBA_OK=1"
    goto :VBA_DONE
)

REM --- Change to 4 ---
echo Current VBAWarnings: !VBA_CUR! -- changing to 4 (Disable all without notification^)...
reg add "!VBA_KEY!" /v VBAWarnings /t REG_DWORD /d 4 /f >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo !RED![!BANG!]!RESET! Failed to update VBAWarnings. This may require admin privileges or a Group Policy override.
    goto :VBA_DONE
)
echo !GREEN![OK]!RESET! VBAWarnings changed to 4 (Disable all macros without notification^).
set "VBA_OK=1"

:VBA_DONE
echo.

REM ============================================================
REM STEP 5 — VERIFY SCRIPT FILE EXISTS
REM ============================================================
echo [5/7] Checking for converter script in this folder...
echo.

if exist "!BASE_DIR!\!SCRIPT_NAME!" (
    echo !GREEN![OK]!RESET! Found: !BASE_DIR!\!SCRIPT_NAME!
    set "SCRIPT_OK=1"
) else (
    echo !RED![!BANG!]!RESET! Missing script: !BASE_DIR!\!SCRIPT_NAME!
    echo     Place the Python converter script in this same folder,
    echo     or update SCRIPT_NAME at the top of this .bat file.
)
echo.

REM Optional template check if you set TEMPLATE_NAME above
if defined TEMPLATE_NAME (
    echo Checking template file...
    if exist "!BASE_DIR!\!TEMPLATE_NAME!" (
        echo !GREEN![OK]!RESET! Found template: !BASE_DIR!\!TEMPLATE_NAME!
    ) else (
        echo !RED![!BANG!]!RESET! Template not found: !BASE_DIR!\!TEMPLATE_NAME!
        echo The Python script can still run, but will fall back to a blank presentation.
    )
    echo.
)

REM ============================================================
REM STEP 6 — GENERATE DYNAMIC FILE PATH CSV
REM ============================================================
echo [6/7] Generating dynamic file path CSV...
echo.

call :run_python "%BASE_DIR%\grabber.py"
set GRABBER_EXIT=!ERRORLEVEL!

IF !GRABBER_EXIT! NEQ 0 (
    echo.
    echo !RED![!BANG!]!RESET! [WARNING] grabber.py encountered an issue ^(exit code !GRABBER_EXIT!^).
) else (
    echo !GREEN![OK]!RESET!      CSV successfully generated.
    set "GRAB_OK=1"
)
:end
REM ============================================================
REM STEP 7 — SUMMARY
REM ============================================================
:SUMMARY
echo [7/7] Setup summary:
echo.

call :SHOW_RESULT "Python" "!PY_OK!"
call :SHOW_RESULT "pip" "!PIP_OK!"
call :SHOW_RESULT "pywin32" "!WIN32_OK!"
call :SHOW_RESULT "Publisher (EXE detect)" "!PUB_OK!"
call :SHOW_RESULT "PowerPoint (EXE detect)" "!PPT_OK!"
call :SHOW_RESULT "Office COM automation" "!COM_OK!"
call :SHOW_RESULT "Publisher macro security" "!VBA_OK!"
call :SHOW_RESULT "CSV of File Paths Made" "!GRAB_OK!"
call :SHOW_RESULT "Converter script present" "!SCRIPT_OK!"

echo.
echo ============================================================
echo Next steps:
echo - If anything failed above, fix it and re-run this batch file.
echo - If everything needed passed, run your converter script:
echo     python "!SCRIPT_NAME!"
echo ============================================================
echo.
echo Opening test_batch folder...
powershell -NoProfile -Command ^
    "$shell = New-Object -ComObject Shell.Application;" ^
    "$shell.Open('!BASE_DIR!\test_batch');" ^
    "Start-Sleep -Milliseconds 800;" ^
    "$win = $shell.Windows() | Where-Object { $_.LocationURL -like '*test_batch*' } | Select-Object -First 1;" ^
    "if ($win) { $win.Left=200; $win.Top=100; $win.Width=900; $win.Height=600 }"
echo This window will stay open. Press any key to close it.
pause >nul
endlocal


REM ============================================================
REM HELPER: SHOW PASS/FAIL
REM ============================================================
:SHOW_RESULT
set "LBL=%~1"
set "VAL=%~2"
if "!VAL!"=="1" (
    echo !GREEN![PASS]!RESET! !LBL!
) else (
    echo !RED![FAIL]!RESET! !LBL!
)
exit /b 0

:: ---------------------------------------------------------------
:: Subroutine :run_python
::
:: Calls the resolved PYTHON_EXE with all passed arguments. We
:: prefer an absolute path discovered by `where python` because
:: that bypasses the Microsoft Store app-execution alias. The
:: Store alias is a stub under %LOCALAPPDATA%\Microsoft\WindowsApps
:: that can close the parent console window when the hosted
:: Python process exits, which is exactly the bug we saw with
:: grabber.py. If we only ever have the alias available, we fall
:: back to plain `python` but route through `cmd /c` to limit the
:: blast radius.
:: ---------------------------------------------------------------
:run_python
if "!PYTHON_EXE!"=="python" (
    cmd /c python %*
) else (
    "!PYTHON_EXE!" %*
)
exit /b %ERRORLEVEL%