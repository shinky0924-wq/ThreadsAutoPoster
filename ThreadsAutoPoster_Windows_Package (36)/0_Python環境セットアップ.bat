@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title Threads AutoPoster - Python Setup

echo ========================================================
echo   Threads AutoPoster - Python Environment Setup Tool
echo ========================================================
echo.

set "PY_EXE="
set "PY_ARGS="

python -c "import sys" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_EXE=python"
    goto :check_packages
)

py -3 -c "import sys" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_EXE=py"
    set "PY_ARGS=-3"
    goto :check_packages
)

py -c "import sys" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_EXE=py"
    goto :check_packages
)

for %%V in (313 312 311 310 39 38) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" -c "import sys" >nul 2>&1
        if !errorlevel! equ 0 (
            set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
            goto :check_packages
        )
    )
)

for %%V in (313 312 311 310 39 38) do (
    if exist "%ProgramFiles%\Python%%V\python.exe" (
        "%ProgramFiles%\Python%%V\python.exe" -c "import sys" >nul 2>&1
        if !errorlevel! equ 0 (
            set "PY_EXE=%ProgramFiles%\Python%%V\python.exe"
            goto :check_packages
        )
    )
)

:: Python not found -> install via winget or browser
echo [1/3] Python was NOT detected or not working on this PC.
echo Attempting automatic installation of Python 3.12 via winget...
winget install --id Python.Python.3.12 --exact --accept-package-agreements --accept-source-agreements
if !errorlevel! equ 0 (
    echo [OK] Python installed successfully!
    for %%V in (312 313 311 310) do (
        if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
            set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        )
    )
    if not defined PY_EXE (
        py -3 -c "import sys" >nul 2>&1
        if !errorlevel! equ 0 (
            set "PY_EXE=py"
            set "PY_ARGS=-3"
        )
    )
    if not defined PY_EXE set "PY_EXE=python"
) else (
    echo.
    echo [NOTICE] winget is not available or installation was cancelled.
    echo Opening python.org in your browser to download the installer...
    start https://www.python.org/downloads/
    echo.
    echo IMPORTANT: When installing manually:
    echo            Check the box: [x] Add python.exe to PATH
    echo.
    pause
    exit /b 1
)

:check_packages
echo [1/3] Python confirmed:
"%PY_EXE%" %PY_ARGS% --version
echo Location: %PY_EXE% %PY_ARGS%
echo.

echo [2/3] Installing required packages (openpyxl, playwright, pillow, pyinstaller)...
"%PY_EXE%" %PY_ARGS% -m pip install --upgrade pip >nul 2>&1
"%PY_EXE%" %PY_ARGS% -m pip install --user -r requirements.txt
if !errorlevel! neq 0 (
    "%PY_EXE%" %PY_ARGS% -m pip install -r requirements.txt
)

echo.
echo [3/3] Setting up Playwright Chromium browser...
"%PY_EXE%" %PY_ARGS% -m playwright install chromium

echo.
echo ========================================================
echo   [SUCCESS] Environment Setup Completed!
echo ========================================================
echo.
echo You can now:
echo   - Run "start_gui.bat" to launch the GUI app directly
echo   - Run "build_exe.bat" to create ThreadsAutoPoster.exe
echo ========================================================
echo.
pause
