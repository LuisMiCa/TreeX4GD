@echo off
title TreeX4GD Builder
echo =================================================
echo TreeX4GD - Windows Executable Builder
echo =================================================
echo.

REM --------------------------------------------------
REM 1. Setup & Path Resolution 
REM --------------------------------------------------
echo [i] Initializing build environment...

REM Safe directory changing without DelayedExpansion
cd /d "%~dp0"
set "BUILD_DIR=%~dp0"

REM Resolve Project Root and Script Directory
pushd ..
set "PROJECT_ROOT=%CD%"
set "APP_SCRIPT_DIR=%PROJECT_ROOT%\script"
popd

REM Define Output paths
set "TMP_DIR=%BUILD_DIR%TMP"
set "DIST_DIR=%BUILD_DIR%TreeX4GD-build-output"
set "VENV_PYTHON=%APP_SCRIPT_DIR%\venv\Scripts\python.exe"
set "REQUIREMENTS=%APP_SCRIPT_DIR%\requirements.txt"

REM Define Source Assets
set "SRC_SPEC=%BUILD_DIR%TreeX4GD.spec"
set "SRC_ICON=%APP_SCRIPT_DIR%\TreeX4GD.ico"
set "SRC_PYTHON=%APP_SCRIPT_DIR%\TreeX4GD.py"
set "SRC_TPL_HTML=%APP_SCRIPT_DIR%\TreeX4GD-template.html"
set "SRC_HELP_HTML=%APP_SCRIPT_DIR%\TreeX4GD-help.html"
set "SRC_IMAGES_DIR=%APP_SCRIPT_DIR%\TreeX4GD-help-images"

echo [i] Project Root: "%PROJECT_ROOT%"
echo [i] Script Dir:   "%APP_SCRIPT_DIR%"
echo.

REM --------------------------------------------------
REM 2. Pre-Flight Path Check
REM --------------------------------------------------
echo [i] Checking project path for problematic characters...
REM Characters '[ ]' can break PyInstaller spec parsing
echo "%PROJECT_ROOT%" | findstr /c:"[" >nul
if not errorlevel 1 goto :BAD_PATH
echo "%PROJECT_ROOT%" | findstr /c:"]" >nul
if not errorlevel 1 goto :BAD_PATH
REM other characters that may break batch commands: & ! ^ ( )
REM not checking those for now, but could be added if needed...

echo [OK] Path is safe for PyInstaller.
echo.
goto :PATH_OK

:BAD_PATH
echo.
echo =================================================
echo [FATAL] INVALID FOLDER NAME DETECTED
echo =================================================
echo PyInstaller may incorrectly parse the project path if it
echo contains certain special characters, such as square brackets [ ].
echo Current Path: 
echo "%PROJECT_ROOT%"
echo.
echo Please rename your project folder to something without those characters,
echo or move it to a different location if the characters are part of a parent folder name. 
echo =================================================
echo.
pause
exit /b 1

:PATH_OK
REM --------------------------------------------------
REM 3. Virtual Environment & Dependency Check
REM --------------------------------------------------
echo [i] Verifying virtual environment and dependencies...

REM 3.1. Check if Python Executable exists
if not exist "%VENV_PYTHON%" (
    echo [ISSUE] Virtual environment not found at:
    echo         "%VENV_PYTHON%"
    goto :FAIL_VENV
)

REM 3.2. Check if requirements.txt exists
if not exist "%REQUIREMENTS%" (
    echo [ISSUE] Requirements file missing at:
    echo         "%REQUIREMENTS%"
    goto :FAIL_VENV
)

REM 3.3. Check if dependencies are installed
REM Safe Inline Python packages Check (using os.environ to avoid quote/path breaks)
"%VENV_PYTHON%" -c "import os, sys, re; from importlib.metadata import distributions; installed = {d.metadata['Name'].lower().replace('-', '_') for d in distributions()}; missing = any((m := re.match(r'^([a-zA-Z0-9_\-\.]+)', l.split('#')[0].strip())) and m.group(1).lower().replace('-', '_') not in installed for l in open(os.environ['REQUIREMENTS'], encoding='utf-8')); sys.exit(1 if missing else 0)" >nul 2>&1

if errorlevel 1 (
    echo [ISSUE] Dependencies are missing or Python failed to initialize.
    goto :FAIL_VENV
)

echo [OK] Virtual environment is healthy and dependencies are satisfied.
echo.
goto :PROCEED_BUILD

REM 3.4. Handle Virtual Environment Issues
:FAIL_VENV
echo.
echo [^>] Please run the main application launcher first to prepare the environment.
echo [^>] Run "run_in_windows.bat" located in:
echo      "%APP_SCRIPT_DIR%"
echo.
pause
exit /b 1

:PROCEED_BUILD
REM --------------------------------------------------
REM 4. PyInstaller Availability & Internet Check
REM --------------------------------------------------
echo [i] Checking PyInstaller...

"%VENV_PYTHON%" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo [i] PyInstaller not found. Installation required.
    echo [i] Verifying internet connection...
    
    REM Check Internet Connectivity (Silently via Python socket)
    "%VENV_PYTHON%" -c "import socket; socket.create_connection(('www.google.com', 443), timeout=3)" >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] You are OFFLINE.
        echo [^>] PyInstaller needs to be downloaded to build the executable, 
        echo     but no internet connection was detected.
        echo [^>] Please connect to the network and try again.
        pause
        exit /b 1
    )
    
    echo [OK] Internet connected. Installing PyInstaller...
    echo ----------------------------------------
    "%VENV_PYTHON%" -m pip install --upgrade pip --quiet
    "%VENV_PYTHON%" -m pip install pyinstaller
    echo ----------------------------------------
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller.
        pause
        exit /b 1
    )
    echo [OK] PyInstaller installed successfully.
) else (
    echo [OK] PyInstaller is already installed.
)
echo.

REM --------------------------------------------------
REM 5. Prepare Temporary Build Directory
REM --------------------------------------------------
echo [i] Preparing temporary workspace...

if exist "%TMP_DIR%" (
    echo [i] Cleaning previous temp files...
    rd /s /q "%TMP_DIR%"
)
mkdir "%TMP_DIR%"

echo [OK] Workspace ready at:
echo        "%TMP_DIR%"
echo.

REM --------------------------------------------------
REM 6. Pre-Flight Check & Asset Copy
REM --------------------------------------------------
echo [i] Verifying source assets...

set "MISSING_FILES=false"

if not exist "%SRC_SPEC%" set "MISSING_FILES=true" & echo [ISSUE] Missing: "%SRC_SPEC%"
if not exist "%SRC_ICON%" set "MISSING_FILES=true" & echo [ISSUE] Missing: "%SRC_ICON%"
if not exist "%SRC_PYTHON%" set "MISSING_FILES=true" & echo [ISSUE] Missing: "%SRC_PYTHON%"
if not exist "%SRC_TPL_HTML%" set "MISSING_FILES=true" & echo [ISSUE] Missing: "%SRC_TPL_HTML%"
if not exist "%SRC_HELP_HTML%" set "MISSING_FILES=true" & echo [ISSUE] Missing: "%SRC_HELP_HTML%"
if not exist "%SRC_IMAGES_DIR%\" set "MISSING_FILES=true" & echo [ISSUE] Missing Folder: "%SRC_IMAGES_DIR%"

if "%MISSING_FILES%"=="true" (
    echo.
    echo [FATAL] Required build files are missing.
    echo [^>] Please ensure all files listed above exist and try again.
    pause
    exit /b 1
)

echo [OK] All source files found.
echo.
echo [i] Copying files to workspace...

REM --- Perform Copy ---
copy /Y "%SRC_SPEC%" "%TMP_DIR%\" >nul
copy /Y "%SRC_ICON%" "%TMP_DIR%\" >nul
copy /Y "%SRC_PYTHON%" "%TMP_DIR%\" >nul
copy /Y "%SRC_TPL_HTML%" "%TMP_DIR%\" >nul
copy /Y "%SRC_HELP_HTML%" "%TMP_DIR%\" >nul
xcopy "%SRC_IMAGES_DIR%" "%TMP_DIR%\TreeX4GD-help-images\" /E /I /Y /Q >nul

echo [OK] Assets copied successfully.
echo.

REM --------------------------------------------------
REM 7. Build Windows Executable
REM --------------------------------------------------
echo [i] Running PyInstaller...
echo ----------------------------------------

cd /d "%TMP_DIR%"
"%VENV_PYTHON%" -m PyInstaller --clean "TreeX4GD.spec"

if errorlevel 1 (
    echo.
    echo ----------------------------------------
    echo [ERROR] Build failed. See output above.
    echo [i] Temporary files kept in "%TMP_DIR%" for inspection.
    pause
    exit /b 1
)

echo ----------------------------------------
echo [OK] PyInstaller finished.
echo.

REM --------------------------------------------------
REM 8. Process Output
REM --------------------------------------------------
echo [i] Processing build artifacts...

set "BUILT_EXE=%TMP_DIR%\dist\TreeX4GD.exe"

if not exist "%BUILT_EXE%" (
    echo [ERROR] TreeX4GD.exe was not found in the dist folder.
    echo [^>] Try running the build again. Check PyInstaller output for clues.
    pause
    exit /b 1
)

REM Prepare Final Destination
if exist "%DIST_DIR%" (
    rd /s /q "%DIST_DIR%"
)
mkdir "%DIST_DIR%"

REM Move the file
move "%BUILT_EXE%" "%DIST_DIR%\" >nul
echo [OK] Binary moved to output directory.

REM --------------------------------------------------
REM 9. Cleanup and Exit
REM --------------------------------------------------
echo [i] Cleaning up temporary workspace...
cd /d "%BUILD_DIR%"
rd /s /q "%TMP_DIR%"

echo.
echo =================================================
echo [OK] BUILD COMPLETE!
echo =================================================
echo.
echo [^>] Executable location:
echo       "%DIST_DIR%\TreeX4GD.exe"
echo.
echo Press any key to close...
pause >nul
exit /b 0