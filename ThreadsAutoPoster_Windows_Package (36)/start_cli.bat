@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title Threads AutoPoster - CLI Runner

set "PY_EXE="
set "PY_ARGS="

python -c "import sys" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_EXE=python"
    goto :run_cli
)

py -3 -c "import sys" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_EXE=py"
    set "PY_ARGS=-3"
    goto :run_cli
)

py -c "import sys" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_EXE=py"
    goto :run_cli
)

for %%V in (313 312 311 310 39 38) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto :run_cli
    )
)

echo [ERROR] Python not found!
echo Please run 0_Python環境セットアップ.bat or install Python from https://www.python.org/downloads/
pause
exit /b 1

:run_cli
echo Running Threads AutoPoster (CLI)...
"%PY_EXE%" %PY_ARGS% threads_auto_poster_cli.py
pause
