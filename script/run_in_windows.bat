@echo off

title TreeX4GD Launcher
echo =================================================
echo TreeX4GD - Python Script Loader for Windows
echo =================================================
echo.

REM --------------------------------------------------
REM 0. Pre-Setup (Run before DelayedExpansion)
REM --------------------------------------------------
cd /d "%~dp0"

setlocal EnableDelayedExpansion

REM --------------------------------------------------
REM 1. Setup Environment Variables
REM --------------------------------------------------
echo [i] Initializing launcher...
set "SCRIPT_FILE=TreeX4GD.py"
set "VENV_DIR=venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"
set "REQUIREMENTS=requirements.txt"
echo.

REM --------------------------------------------------
REM 2. Python Script Presence Check
REM --------------------------------------------------
if not exist "%SCRIPT_FILE%" (
    echo [ISSUE] "%SCRIPT_FILE%" not found in the current directory.
    pause
    exit /b 1
)

REM --------------------------------------------------
REM 3. Check System Python
REM --------------------------------------------------
echo [i] Checking for system Python...
python --version >nul 2>&1
if !errorlevel! equ 0 (
    set "SYS_PYTHON=python"
    echo [OK] Found Python via 'python'
) else (
    py --version >nul 2>&1
    if !errorlevel! equ 0 (
        set "SYS_PYTHON=py"
        echo [OK] Found Python via 'py'
    ) else (
        echo [ISSUE] Python is not installed or not in PATH.
        echo [^>] Please install Python 3.11 or newer and ensure it is added to PATH.
        pause
        exit /b 1
    )
)
echo.

REM --------------------------------------------------
REM 4. Check if Tkinter is available
REM --------------------------------------------------
echo [i] Verifying Tkinter availability...
%SYS_PYTHON% -c "import tkinter" >nul 2>&1
if !errorlevel! neq 0 (
    echo [ERROR] 'tkinter' module is missing.
    echo [^>] Reinstall Python and enable 'tcl/tk and IDLE'.
    pause
    exit /b 1
)
echo [OK] Tkinter is available.
echo.

REM --------------------------------------------------
REM 5. Virtual Environment Validation & Creation if Needed
REM --------------------------------------------------
set "NEED_VENV=false"
echo [i] Checking python virtual environment...

REM 5.1. Check if Venv folder exists and is functional
if exist "%VENV_DIR%\" (
    echo [i] VENV detected. Validating...
    "%VENV_PYTHON%" -c "import sys; sys.exit(0)" >nul 2>&1
    if !errorlevel! neq 0 (
        echo [ISSUE] Virtual environment appears corrupted.
        echo [i] Removing venv folder for clean rebuild...
        rd /s /q "%VENV_DIR%"
        set "NEED_VENV=true"
    ) else (
        echo [OK] Python virtual environment is apparently healthy.
        echo.
    )
) else (
    echo [i] Virtual environment not created yet...
    set "NEED_VENV=true"
)

REM 5.2. Create Venv if not present or corrupted
if "!NEED_VENV!"=="true" (
    echo [i] Creating virtual environment...
    %SYS_PYTHON% -m venv "%VENV_DIR%"
    if !errorlevel! neq 0 (
        echo [ERROR] Failed to create virtual environment.
        echo [^>] Check permissions and Python installation.
        pause
        exit /b 1
    )

    if not exist "%VENV_PYTHON%" (
        echo [ISSUE] Venv creation failed: python.exe not found in the Venv.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created successfully.
    echo.
)

REM --------------------------------------------------
REM 6. Dependency Check & Installation
REM --------------------------------------------------
echo [i] Checking dependencies...

REM 6.1. Check if requirements.txt exists
if not exist "%REQUIREMENTS%" (
    echo [ERROR] File '%REQUIREMENTS%' is missing.
    echo [^>] This file is required to proceed.
    pause
    exit /b 1
)

REM 6.2. Check Dependencies (Silently)
REM Exits with 1 (True) if missing, 0 (False) if satisfied.
"%VENV_PYTHON%" -c "import os, sys, re; from importlib.metadata import distributions; installed = {d.metadata['Name'].lower().replace('-', '_') for d in distributions()}; missing = any((m := re.match(r'^([a-zA-Z0-9_\-\.]+)', l.split('#')[0].strip())) and m.group(1).lower().replace('-', '_') not in installed for l in open(os.environ['REQUIREMENTS'], encoding='utf-8')); sys.exit(1 if missing else 0)" >nul 2>&1
set "MISSING_DEPS=!errorlevel!"

REM 6.3. Check Internet Connectivity (Silently)
REM Uses Python socket to verify an outward connection without triggering SSL certificate errors.
"%VENV_PYTHON%" -c "import socket; socket.create_connection(('www.google.com', 443), timeout=3)" >nul 2>&1
if !errorlevel! equ 0 ( set "IS_ONLINE=true" ) else ( set "IS_ONLINE=false" )

REM 6.4. Decision Logic
if "!MISSING_DEPS!" neq "0" (
    REM Case: Dependencies MISSING
    if "!IS_ONLINE!"=="false" (
        echo [ERROR] Dependencies are missing, and you are OFFLINE.
        echo [^>] We cannot install required libraries without internet.
        echo [^>] Please connect to the network and try again.
        pause
        exit /b 1
    ) else (
        echo [i] Missing requirements detected. Installing...
        echo ----------------------------------------
        "%VENV_PYTHON%" -m pip install --upgrade pip --quiet
        "%VENV_PYTHON%" -m pip install -r "%REQUIREMENTS%"
        if !errorlevel! neq 0 (
            echo ----------------------------------------
            echo [ERROR] Dependency installation failed.
            pause
            exit /b 1
        )
        echo ----------------------------------------
        echo [OK] Dependencies installed.
    )
) else (
    REM Case: Dependencies PRESENT
    if "!IS_ONLINE!"=="false" (
        echo [WARN] YOU ARE OFFLINE
        echo [i] All required libraries are installed, so the app will launch.
        echo [^>] However, features requiring internet access will not work,
        echo     but you can still process local JSON files.
        echo ----------------------------------------
        timeout /t 3 >nul
    ) else (
        echo [OK] Dependencies satisfied ^& Internet connected.
    )
)
echo.

REM --------------------------------------------------
REM 7. Launch Application
REM --------------------------------------------------
echo [i] Launching %SCRIPT_FILE%...
echo ----------------------------------------

"%VENV_PYTHON%" "%SCRIPT_FILE%"
set "EXIT_CODE=!errorlevel!"

echo ----------------------------------------
if !EXIT_CODE! neq 0 (
    echo [ERROR] Application exited with error code !EXIT_CODE!.
    pause
) else (
    echo [OK] Application finished successfully.
    echo.
    echo Exiting in 3 seconds...
    timeout /t 3 >nul
)

exit /b !EXIT_CODE!