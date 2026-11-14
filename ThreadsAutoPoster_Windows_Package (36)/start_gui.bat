@echo off
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"
title Threads AutoPoster Launcher

set "LOG_FILE=%~dp0startup_log.txt"

:: 1. Check if running inside temporary unextracted folder
echo %~dp0 | findstr /i "Temp Temporary" >nul
if %errorlevel% equ 0 (
    echo ========================================================
    echo   【重要】ZIPファイルが展開（解凍）されていません！
    echo ========================================================
    echo.
    echo ZIPファイルを開いたまま直接実行すると、他のファイルが読み込めません。
    echo ダウンロードしたZIPファイルを右クリックし、
    echo 「すべて展開」を選択して解凍してから実行してください。
    echo ========================================================
    echo.
    pause
    exit /b 1
)

:: 2. Locate threads_gui_app.py
if exist "%~dp0threads_gui_app.py" goto :script_ok
if exist "%~dp0ThreadsAutoPoster\threads_gui_app.py" (
    cd /d "%~dp0ThreadsAutoPoster"
    goto :script_ok
)

echo ========================================================
echo   【エラー】threads_gui_app.py が見つかりません
echo ========================================================
echo.
echo ZIPファイル内のすべてのファイルが同じフォルダに解凍されているか確認してください。
echo ========================================================
echo.
pause
exit /b 1

:script_ok
set "PY_EXE="
set "PY_ARGS="
set "PYW_EXE="

:: 3. Detect Python environment
py -3 -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    set "PY_ARGS=-3"
    pyw -3 -c "import sys" >nul 2>&1
    if %errorlevel% equ 0 set "PYW_EXE=pyw"
    goto :python_found
)

python -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=python"
    goto :python_found
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" set "PYW_EXE=%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe"
    goto :python_found
)

if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe" set "PYW_EXE=%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe"
    goto :python_found
)

if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\pythonw.exe" set "PYW_EXE=%LOCALAPPDATA%\Programs\Python\Python311\pythonw.exe"
    goto :python_found
)

if exist "%ProgramFiles%\Python312\python.exe" (
    set "PY_EXE=%ProgramFiles%\Python312\python.exe"
    if exist "%ProgramFiles%\Python312\pythonw.exe" set "PYW_EXE=%ProgramFiles%\Python312\pythonw.exe"
    goto :python_found
)

if exist "%ProgramFiles%\Python313\python.exe" (
    set "PY_EXE=%ProgramFiles%\Python313\python.exe"
    if exist "%ProgramFiles%\Python313\pythonw.exe" set "PYW_EXE=%ProgramFiles%\Python313\pythonw.exe"
    goto :python_found
)

for %%P in ("%LOCALAPPDATA%\Programs\Python\Python310\python.exe" "C:\Python312\python.exe" "C:\Python311\python.exe") do (
    if not defined PY_EXE if exist %%P set "PY_EXE=%%~P"
)

if not defined PY_EXE goto :err_no_python

:python_found
:: 4. Resolve pythonw.exe (Windowless GUI runner)
if not defined PYW_EXE (
    for /f "delims=" %%I in ('"%PY_EXE%" %PY_ARGS% -c "import sys, os; w = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'); print(w if os.path.exists(w) else '')" 2^>nul') do set "PYW_EXE=%%I"
)

:: 5. Check required packages (quick silent check)
"%PY_EXE%" %PY_ARGS% -c "import openpyxl; import playwright; import PIL" >nul 2>&1
if %errorlevel% neq 0 (
    echo ライブラリを確認・自動導入しています (初回のみ約1分)...
    "%PY_EXE%" %PY_ARGS% -m pip install openpyxl playwright pillow >nul 2>&1
    if %errorlevel% neq 0 (
        "%PY_EXE%" %PY_ARGS% -m pip install --user openpyxl playwright pillow >nul 2>&1
    )
    "%PY_EXE%" %PY_ARGS% -m playwright install chromium >nul 2>&1
)

:: 6. Launch GUI App and close CMD window immediately
if defined PYW_EXE (
    start "" "%PYW_EXE%" %PY_ARGS% "threads_gui_app.py"
) else (
    start "" "%PY_EXE%" %PY_ARGS% "threads_gui_app.py"
)
exit /b 0

:err_no_python
echo ========================================================
echo   【確認】このPCに Python が見つかりませんでした
echo ========================================================
echo.
echo Threads AutoPoster の実行には Python が必要です。
echo.
echo   [1] Python 3.12 を自動インストールする (推奨)
echo   [2] Python公式サイト (python.org) をブラウザで開く
echo   [3] 終了
echo ========================================================
echo.
set /p USER_CHOICE="選択肢の番号を入力してください [1, 2, 3]: "

if "%USER_CHOICE%"=="1" (
    echo.
    echo Python 3.12 をインストールしています...
    winget install --id Python.Python.3.12 --exact --accept-package-agreements --accept-source-agreements
    echo.
    echo インストールが完了しました。再度「start_gui.bat」を実行してください。
    pause
    exit /b 0
)

if "%USER_CHOICE%"=="2" (
    start https://www.python.org/downloads/
    exit /b 0
)

exit /b 1
