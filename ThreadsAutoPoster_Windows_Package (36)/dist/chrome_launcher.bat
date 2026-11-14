@echo off
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"
title Chrome Custom Profile Launcher

if not exist "chrome_launcher.py" goto :try_exe

set "PY_EXE="
set "PY_ARGS="

py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 goto :set_py_launcher

py -c "import sys" >nul 2>&1
if not errorlevel 1 goto :set_py_plain

python -c "import sys" >nul 2>&1
if not errorlevel 1 goto :set_python_cmd

python3 -c "import sys" >nul 2>&1
if not errorlevel 1 goto :set_python3_cmd

if exist "python\python.exe" goto :set_local_py
if exist "venv\Scripts\python.exe" goto :set_venv_py
if exist ".venv\Scripts\python.exe" goto :set_dotvenv_py

if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" goto :set_py314_local
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" goto :set_py313_local
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" goto :set_py312_local
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" goto :set_py311_local
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" goto :set_py310_local
if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe" goto :set_py39_local

if exist "%ProgramFiles%\Python314\python.exe" goto :set_py314_pf
if exist "%ProgramFiles%\Python313\python.exe" goto :set_py313_pf
if exist "%ProgramFiles%\Python312\python.exe" goto :set_py312_pf
if exist "%ProgramFiles%\Python311\python.exe" goto :set_py311_pf
if exist "%ProgramFiles%\Python310\python.exe" goto :set_py310_pf

if exist "C:\Python313\python.exe" goto :set_py313_root
if exist "C:\Python312\python.exe" goto :set_py312_root
if exist "C:\Python311\python.exe" goto :set_py311_root
if exist "C:\Python310\python.exe" goto :set_py310_root

if exist "%USERPROFILE%\anaconda3\python.exe" goto :set_anaconda_user
if exist "%USERPROFILE%\miniconda3\python.exe" goto :set_miniconda_user
if exist "%LOCALAPPDATA%\anaconda3\python.exe" goto :set_anaconda_local
if exist "%LOCALAPPDATA%\miniconda3\python.exe" goto :set_miniconda_local

assoc .py 2>nul | findstr /i "Python" >nul 2>&1
if not errorlevel 1 goto :run_py_assoc

goto :try_exe

:set_py_launcher
set "PY_EXE=py"
set "PY_ARGS=-3"
goto :run_py

:set_py_plain
set "PY_EXE=py"
goto :run_py

:set_python_cmd
set "PY_EXE=python"
goto :run_py

:set_python3_cmd
set "PY_EXE=python3"
goto :run_py

:set_local_py
set "PY_EXE=%CD%\python\python.exe"
goto :run_py

:set_venv_py
set "PY_EXE=%CD%\venv\Scripts\python.exe"
goto :run_py

:set_dotvenv_py
set "PY_EXE=%CD%\.venv\Scripts\python.exe"
goto :run_py

:set_py314_local
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
goto :run_py

:set_py313_local
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
goto :run_py

:set_py312_local
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
goto :run_py

:set_py311_local
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
goto :run_py

:set_py310_local
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
goto :run_py

:set_py39_local
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
goto :run_py

:set_py314_pf
set "PY_EXE=%ProgramFiles%\Python314\python.exe"
goto :run_py

:set_py313_pf
set "PY_EXE=%ProgramFiles%\Python313\python.exe"
goto :run_py

:set_py312_pf
set "PY_EXE=%ProgramFiles%\Python312\python.exe"
goto :run_py

:set_py311_pf
set "PY_EXE=%ProgramFiles%\Python311\python.exe"
goto :run_py

:set_py310_pf
set "PY_EXE=%ProgramFiles%\Python310\python.exe"
goto :run_py

:set_py313_root
set "PY_EXE=C:\Python313\python.exe"
goto :run_py

:set_py312_root
set "PY_EXE=C:\Python312\python.exe"
goto :run_py

:set_py311_root
set "PY_EXE=C:\Python311\python.exe"
goto :run_py

:set_py310_root
set "PY_EXE=C:\Python310\python.exe"
goto :run_py

:set_anaconda_user
set "PY_EXE=%USERPROFILE%\anaconda3\python.exe"
goto :run_py

:set_miniconda_user
set "PY_EXE=%USERPROFILE%\miniconda3\python.exe"
goto :run_py

:set_anaconda_local
set "PY_EXE=%LOCALAPPDATA%\anaconda3\python.exe"
goto :run_py

:set_miniconda_local
set "PY_EXE=%LOCALAPPDATA%\miniconda3\python.exe"
goto :run_py

:run_py
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
"%PY_EXE%" %PY_ARGS% "chrome_launcher.py" %*
if errorlevel 1 pause
exit /b %errorlevel%

:run_py_assoc
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
call "chrome_launcher.py" %*
if errorlevel 1 pause
exit /b %errorlevel%

:try_exe
if exist "chrome_launcher.exe" goto :run_exe

set "CSC_EXE="
if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC_EXE=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC_EXE=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC_EXE=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC_EXE=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not defined CSC_EXE goto :no_runtime
if not exist "ChromeLauncher.cs" goto :no_runtime

echo chrome_launcher.exe を生成しています...
"%CSC_EXE%" /nologo /target:exe /codepage:65001 /utf8output /out:"chrome_launcher.exe" "ChromeLauncher.cs"
if exist "chrome_launcher.exe" goto :run_exe

:no_runtime
echo =========================================
echo [エラー] Python または実行環境が見つかりませんでした。
echo chrome_launcher.py を直接実行できる場合はそちらをご利用ください。
echo =========================================
pause
exit /b 1

:run_exe
"chrome_launcher.exe" %*
if errorlevel 1 goto :exe_blocked
exit /b 0

:exe_blocked
echo =========================================
echo [エラー] Windows の Device Guard / セキュリティポリシーにより chrome_launcher.exe がブロックされました。
echo Python 経由で chrome_launcher.py を直接起動してください。
echo =========================================
pause
exit /b 1
