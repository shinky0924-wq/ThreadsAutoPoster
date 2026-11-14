@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion
cd /d "%~dp0"
title Threads AutoPoster - 他PC用自動環境セットアップ

echo ========================================================
echo   Threads AutoPoster - 他PC用自動環境セットアップ
echo ========================================================
echo.
echo このPCで Threads AutoPoster を動かすための準備を全自動で行います。
echo 【Python確認・Playwrightブラウザ導入・Excel読み込み準備】
echo.

set "PY_EXE="
set "PY_ARGS="

py -3 -c "import sys; sys.exit(0 if 'WindowsApps' not in sys.executable else 1)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    set "PY_ARGS=-3"
    goto :python_ready
)

python -c "import sys; sys.exit(0 if 'WindowsApps' not in sys.executable else 1)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=python"
    goto :python_ready
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"

if not defined PY_EXE if exist "%ProgramFiles%\Python312\python.exe" set "PY_EXE=%ProgramFiles%\Python312\python.exe"
if not defined PY_EXE if exist "%ProgramFiles%\Python313\python.exe" set "PY_EXE=%ProgramFiles%\Python313\python.exe"
if not defined PY_EXE if exist "%ProgramFiles%\Python311\python.exe" set "PY_EXE=%ProgramFiles%\Python311\python.exe"
if not defined PY_EXE if exist "%ProgramFiles%\Python310\python.exe" set "PY_EXE=%ProgramFiles%\Python310\python.exe"

if defined PY_EXE goto :python_ready

:install_python
echo [1/3] このPCにPythonが検出されませんでした。
echo Windowsパッケージマネージャー 【winget】 で自動インストールします...
winget install --id Python.Python.3.12 --exact --accept-package-agreements --accept-source-agreements
if %errorlevel% equ 0 (
    echo [OK] Pythonのインストールが完了しました！
    set "PY_EXE=python"
    goto :python_ready
)

echo.
echo [案内] ブラウザでPython公式サイトを開きます。
echo インストーラー実行時は必ず最初の画面で『Add python.exe to PATH』にチェックを入れてください。
start https://www.python.org/downloads/
echo.
echo インストール完了後、この画面で何かキーを押すと続行します...
pause
set "PY_EXE=python"

:python_ready
echo.
echo [1/3] Python環境を確認しました:
"%PY_EXE%" %PY_ARGS% --version
echo.

echo [2/3] 必要なライブラリ 【openpyxl, playwright, pillow】 をインストール中...
"%PY_EXE%" %PY_ARGS% -m pip install --upgrade pip >nul 2>&1
"%PY_EXE%" %PY_ARGS% -m pip install --user openpyxl playwright pillow
if %errorlevel% neq 0 (
    "%PY_EXE%" %PY_ARGS% -m pip install openpyxl playwright pillow
)

echo.
echo [3/3] ブラウザエンジン 【Playwright Chromium】 をセットアップ中...
"%PY_EXE%" %PY_ARGS% -m playwright install chromium

echo.
echo ========================================================
echo   【準備完了】このPCの環境構築がすべて完了しました！
echo ========================================================
echo.
echo これでこのPCでもすぐに自動投稿を実行できます:
echo   ・「ThreadsAutoPoster.exe」をダブルクリックして起動
echo   ・または「start_gui.bat」から即座に起動可能
echo.
echo ========================================================
echo.
echo 処理が完了しました。何かキーを押すと画面を閉じます...
pause
exit /b 0
