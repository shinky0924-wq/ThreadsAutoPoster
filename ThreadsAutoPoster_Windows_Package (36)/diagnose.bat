@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title Threads AutoPoster - 原因診断＆自動修復

if exist "%~dp0diagnose.ps1" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnose.ps1"
    if !errorlevel! equ 0 exit /b 0
)

echo ========================================================
echo   Threads AutoPoster - 原因診断＆自動修復
echo ========================================================
echo.
echo [1/3] ファイルの存在チェック...
if exist "threads_gui_app.py" (echo   [OK] threads_gui_app.py : 存在します) else (echo   [NG] threads_gui_app.py : 見つかりません)
if exist "ThreadsAutoPoster.exe" (echo   [OK] ThreadsAutoPoster.exe : 存在します) else (echo   [NG] ThreadsAutoPoster.exe : 見つかりません)
if exist "アカウント情報.xlsx" (echo   [OK] アカウント情報.xlsx : 存在します) else (echo   [WARN] アカウント情報.xlsx : 見つかりません)
echo.

echo [2/3] Python環境とライブラリの自動修復中 (--user)...
set "DO_PY="
for %%P in (python.exe py.exe "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%ProgramFiles%\Python312\python.exe") do (
    if not defined DO_PY (
        %%P -c "import sys" >nul 2>&1
        if !errorlevel! equ 0 set "DO_PY=%%P"
    )
)

if defined DO_PY (
    echo   Pythonを検出: !DO_PY!
    !DO_PY! -m pip install --user openpyxl playwright
    !DO_PY! -m playwright install chromium
    echo.
    echo [OK] ライブラリの修復が完了しました！
    echo ThreadsAutoPoster.exe をダブルクリックして起動してください。
) else (
    echo   [NG] Pythonが見つかりませんでした。
    echo   「0_他のPCでもすぐ動く環境作成.bat」を実行してください。
)

echo.
pause
