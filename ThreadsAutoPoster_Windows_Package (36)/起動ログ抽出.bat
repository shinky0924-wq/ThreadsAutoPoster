@echo off
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"
title Threads AutoPoster - 起動ログ抽出・診断ツール

set "LOG_FILE=%~dp0startup_log.txt"
set "DESK_LOG=%USERPROFILE%\Desktop\ThreadsAutoPoster_起動ログ.txt"

echo ========================================================
echo   Threads AutoPoster - 起動ログ抽出・診断ツール
echo ========================================================
echo.
echo 起動エラーの原因を特定するため、詳細ログをデスクトップへ抽出しています...
echo.

echo ======================================================== > "%LOG_FILE%"
echo   Threads AutoPoster - Startup Diagnostic Log Report >> "%LOG_FILE%"
echo   Date: %date% %time% >> "%LOG_FILE%"
echo   Folder: %~dp0 >> "%LOG_FILE%"
echo   OS: >> "%LOG_FILE%"
ver >> "%LOG_FILE%" 2>&1
echo ======================================================== >> "%LOG_FILE%"

echo --- [1/4] Checking Required Files --- >> "%LOG_FILE%"
if exist "%~dp0threads_gui_app.py" (
    echo [OK] threads_gui_app.py: Found >> "%LOG_FILE%"
) else (
    echo [NG] threads_gui_app.py: NOT Found (Please Extract All from ZIP!) >> "%LOG_FILE%"
)
if exist "%~dp0ThreadsAutoPoster.exe" (
    echo [OK] ThreadsAutoPoster.exe: Found >> "%LOG_FILE%"
) else (
    echo [NG] ThreadsAutoPoster.exe: NOT Found >> "%LOG_FILE%"
)
if exist "%~dp0アカウント情報.xlsx" (
    echo [OK] Account file (アカウント情報.xlsx): Found >> "%LOG_FILE%"
) else (
    echo [WARN] Account file (アカウント情報.xlsx): NOT Found in this folder >> "%LOG_FILE%"
)

echo. >> "%LOG_FILE%"
echo --- [2/4] Python Environment Check --- >> "%LOG_FILE%"
python -c "import sys; print('PATH python:', sys.executable, sys.version)" >> "%LOG_FILE%" 2>&1
py -3 -c "import sys; print('py -3:', sys.executable, sys.version)" >> "%LOG_FILE%" 2>&1

set "TEST_PY="
python -c "import sys" >nul 2>&1 && set "TEST_PY=python"
if not defined TEST_PY py -3 -c "import sys" >nul 2>&1 && set "TEST_PY=py -3"
if not defined TEST_PY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "TEST_PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined TEST_PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "TEST_PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined TEST_PY if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "TEST_PY=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not defined TEST_PY if exist "%ProgramFiles%\Python312\python.exe" set "TEST_PY=%ProgramFiles%\Python312\python.exe"

if defined TEST_PY (
    echo [OK] Selected Python: %TEST_PY% >> "%LOG_FILE%"
    echo --- Library Import Tests --- >> "%LOG_FILE%"
    %TEST_PY% -c "import tkinter; print('[OK] tkinter: Working')" >> "%LOG_FILE%" 2>&1
    %TEST_PY% -c "import openpyxl; print('[OK] openpyxl: Working')" >> "%LOG_FILE%" 2>&1
    %TEST_PY% -c "import playwright; print('[OK] playwright: Working')" >> "%LOG_FILE%" 2>&1
) else (
    echo [FATAL] No working Python found on this PC. >> "%LOG_FILE%"
)

echo. >> "%LOG_FILE%"
echo --- [3/4] Recent Crash Logs --- >> "%LOG_FILE%"
if exist "%~dp0crash_log.txt" (
    type "%~dp0crash_log.txt" >> "%LOG_FILE%" 2>&1
) else (
    echo No crash_log.txt found. >> "%LOG_FILE%"
)

echo. >> "%LOG_FILE%"
echo --- [4/4] GUI Application Test Launch --- >> "%LOG_FILE%"
if defined TEST_PY if exist "%~dp0threads_gui_app.py" (
    echo Executing: %TEST_PY% "%~dp0threads_gui_app.py" >> "%LOG_FILE%"
    %TEST_PY% "%~dp0threads_gui_app.py" >> "%LOG_FILE%" 2>&1
    echo Exit Code: %errorlevel% >> "%LOG_FILE%"
)

copy /y "%LOG_FILE%" "%DESK_LOG%" >nul 2>&1

echo ========================================================
echo   [OK] 起動ログの抽出が完了しました（文字化け防止版）
echo ========================================================
echo.
echo 保存場所:
echo   ・デスクトップ: %DESK_LOG%
echo   ・フォルダ内:   %LOG_FILE%
echo.
echo メモ帳でログを開きます...
start notepad.exe "%LOG_FILE%"
echo.
echo 何かキーを押すとこの画面を閉じます...
pause
