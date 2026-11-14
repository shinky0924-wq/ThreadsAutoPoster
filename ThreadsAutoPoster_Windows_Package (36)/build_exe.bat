@echo off
setlocal
cd /d "%~dp0"
title Threads AutoPoster - EXE Builder

echo ========================================================
echo   Threads AutoPoster - EXE Builder
echo ========================================================
echo.

if not exist "%~dp0threads_gui_app.py" goto :err_no_script

echo [1/4] Preparing dist folder and copying required files...
if not exist "dist" mkdir "dist"
if not exist "dist\ThreadsAutoPoster" mkdir "dist\ThreadsAutoPoster"

copy /y *.xlsx "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y *.png "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y *.vbs "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "threads_gui_app.py" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "threads_auto_poster_cli.py" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "chrome_launcher.py" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "chrome_launcher.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "ChromeLauncher.cs" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "requirements.txt" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "0_他のPCでもすぐ動く環境作成.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "0_Python環境セットアップ.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "start_gui.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "他のPCでの利用手順_マニュアル.txt" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "セキュリティ警告解除.bat" "dist\ThreadsAutoPoster\" >nul 2>&1

if not exist "dist\ThreadsAutoPoster\images" mkdir "dist\ThreadsAutoPoster\images" >nul 2>&1
if exist "images" xcopy /e /i /y "images" "dist\ThreadsAutoPoster\images" >nul 2>&1

copy /y *.xlsx "dist\" >nul 2>&1
copy /y *.png "dist\" >nul 2>&1
copy /y *.vbs "dist\" >nul 2>&1
copy /y "threads_gui_app.py" "dist\" >nul 2>&1
copy /y "start_gui.bat" "dist\" >nul 2>&1
copy /y "他のPCでの利用手順_マニュアル.txt" "dist\" >nul 2>&1
copy /y "セキュリティ警告解除.bat" "dist\" >nul 2>&1

echo [OK] Required files and settings copied to dist folder.
echo.

echo [2/4] Compiling Windows Native Launcher [ThreadsAutoPoster.exe]...
set "CSC_EXE="
if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC_EXE=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC_EXE=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not defined CSC_EXE goto :skip_csc
if not exist "Launcher.cs" goto :skip_csc

"%CSC_EXE%" /nologo /target:winexe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.dll /out:"dist\ThreadsAutoPoster\ThreadsAutoPoster.exe" "Launcher.cs" > build_csc.log 2>&1
if exist "dist\ThreadsAutoPoster\ThreadsAutoPoster.exe" (
    copy /y "dist\ThreadsAutoPoster\ThreadsAutoPoster.exe" "dist\ThreadsAutoPoster.exe" >nul 2>&1
    echo [OK] Native Windows Launcher [ThreadsAutoPoster.exe] generated successfully!
) else (
    echo [INFO] csc compilation note:
    type build_csc.log 2>nul
)
if exist "ChromeLauncher.cs" (
    "%CSC_EXE%" /nologo /target:exe /codepage:65001 /utf8output /out:"dist\ThreadsAutoPoster\chrome_launcher.exe" "ChromeLauncher.cs" >> build_csc.log 2>&1
    if exist "dist\ThreadsAutoPoster\chrome_launcher.exe" (
        copy /y "dist\ThreadsAutoPoster\chrome_launcher.exe" "dist\chrome_launcher.exe" >nul 2>&1
        copy /y "dist\ThreadsAutoPoster\chrome_launcher.exe" "chrome_launcher.exe" >nul 2>&1
        echo [OK] Custom Profile Launcher [chrome_launcher.exe] generated successfully!
    )
)
:skip_csc
echo.

echo [3/4] Checking Python environment and installing libraries...
set "PY_EXE="
set "PY_ARGS="

python -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=python"
    goto :python_ok
)

py -3 -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    set "PY_ARGS=-3"
    goto :python_ok
)

py -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    set "PY_EXE=py"
    goto :python_ok
)

if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not defined PY_EXE if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
if not defined PY_EXE if exist "C:\Program Files\Python312\python.exe" set "PY_EXE=C:\Program Files\Python312\python.exe"
if not defined PY_EXE if exist "C:\Program Files\Python313\python.exe" set "PY_EXE=C:\Program Files\Python313\python.exe"
if not defined PY_EXE if exist "C:\Program Files\Python311\python.exe" set "PY_EXE=C:\Program Files\Python311\python.exe"
if not defined PY_EXE if exist "C:\Program Files\Python310\python.exe" set "PY_EXE=C:\Program Files\Python310\python.exe"

:python_ok
if not defined PY_EXE goto :no_py_installer

echo Python detected: %PY_EXE% %PY_ARGS%
echo Installing / updating required libraries (openpyxl, playwright)...
"%PY_EXE%" %PY_ARGS% -m pip install openpyxl playwright >nul 2>&1
echo [OK] Libraries verified.
goto :build_finish

:no_py_installer
echo [INFO] Python was not detected yet. Native launcher package has been prepared in dist.

:build_finish
echo.
echo [4/4] Finalizing file deployment in dist folder...
copy /y *.xlsx "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y *.png "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y *.vbs "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "threads_gui_app.py" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "threads_auto_poster_cli.py" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "chrome_launcher.py" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "chrome_launcher.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "ChromeLauncher.cs" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "0_他のPCでもすぐ動く環境作成.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "start_gui.bat" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "他のPCでの利用手順_マニュアル.txt" "dist\ThreadsAutoPoster\" >nul 2>&1
copy /y "セキュリティ警告解除.bat" "dist\ThreadsAutoPoster\" >nul 2>&1

copy /y *.xlsx "dist\" >nul 2>&1
copy /y *.png "dist\" >nul 2>&1
copy /y *.vbs "dist\" >nul 2>&1
copy /y "threads_gui_app.py" "dist\" >nul 2>&1
copy /y "threads_auto_poster_cli.py" "dist\" >nul 2>&1
copy /y "chrome_launcher.py" "dist\" >nul 2>&1
copy /y "chrome_launcher.bat" "dist\" >nul 2>&1
copy /y "ChromeLauncher.cs" "dist\" >nul 2>&1
copy /y "start_gui.bat" "dist\" >nul 2>&1
copy /y "他のPCでの利用手順_マニュアル.txt" "dist\" >nul 2>&1
copy /y "セキュリティ警告解除.bat" "dist\" >nul 2>&1
if exist "dist\ThreadsAutoPoster\ThreadsAutoPoster.exe" (
    copy /y "dist\ThreadsAutoPoster\ThreadsAutoPoster.exe" "dist\" >nul 2>&1
)

echo.
echo ========================================================
echo   [SUCCESS] dist folder build is COMPLETE!
echo ========================================================
echo.
echo   Location:
echo      dist\ThreadsAutoPoster\
echo.
echo   Included Files:
echo      - ThreadsAutoPoster.exe (Main Launcher)
echo      - アカウント情報.xlsx
echo      - threads_main_account.xlsx
echo      - start_gui.bat
echo      - 0_他のPCでもすぐ動く環境作成.bat
echo      - 他のPCでの利用手順_マニュアル.txt
echo.
echo   [How to use on other PCs]:
echo      Copy the "dist\ThreadsAutoPoster" folder to your other PC.
echo      On the other PC, run "0_他のPCでもすぐ動く環境作成.bat" once.
echo ========================================================
echo.
echo Process complete. Press any key to close this window...
pause
exit /b 0

:err_no_script
echo ========================================================
echo   [ERROR] threads_gui_app.py was not found!
echo ========================================================
echo.
echo   Cause: You may have opened build_exe.bat directly
echo          from inside the ZIP file without extracting it.
echo.
echo   Solution:
echo     1. Right-click the downloaded ZIP file.
echo     2. Click "Extract All..." (すべて展開).
echo     3. Open the extracted folder and run build_exe.bat again.
echo.
echo ========================================================
echo.
echo Press any key to close this window...
pause
exit /b 1
