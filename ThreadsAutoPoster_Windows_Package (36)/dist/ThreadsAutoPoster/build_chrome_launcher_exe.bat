@echo off
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"
title chrome_launcher.exe Build Tool

echo ===============================================================
echo   chrome_launcher.exe ビルドツール
echo ===============================================================
echo.

set "CSC_EXE="
if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC_EXE=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC_EXE=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC_EXE=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC_EXE if exist "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC_EXE=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not defined CSC_EXE goto :try_pyinstaller
if not exist "ChromeLauncher.cs" goto :try_pyinstaller

echo [1/1] Windows標準C#コンパイラで chrome_launcher.exe を生成しています...
"%CSC_EXE%" /nologo /target:exe /codepage:65001 /utf8output /out:"chrome_launcher.exe" "ChromeLauncher.cs"
if exist "chrome_launcher.exe" goto :build_ok

:try_pyinstaller
echo [INFO] PyInstaller によるビルドを試行します...
py -3 -m pip install pyinstaller >nul 2>&1
py -3 -m PyInstaller --onefile --console --name chrome_launcher chrome_launcher.py
if exist "dist\chrome_launcher.exe" goto :copy_dist_exe

python -m pip install pyinstaller >nul 2>&1
python -m PyInstaller --onefile --console --name chrome_launcher chrome_launcher.py
if exist "dist\chrome_launcher.exe" goto :copy_dist_exe

echo [ERROR] chrome_launcher.exe の生成に失敗しました。
pause
exit /b 1

:copy_dist_exe
copy /y "dist\chrome_launcher.exe" ".\chrome_launcher.exe" >nul
goto :build_ok

:build_ok
echo.
echo ===============================================================
echo   作成完了！ chrome_launcher.exe が生成されました。
echo ===============================================================
pause
exit /b 0
