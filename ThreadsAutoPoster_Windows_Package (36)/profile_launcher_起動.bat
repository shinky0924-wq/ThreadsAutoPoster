@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Threads Profile Launcher (profile_launcher)

echo ===============================================================
echo   Threads Profile Launcher (profile_launcher) 起動中...
echo ===============================================================
echo.

if exist "python_embed\python.exe" (
    "python_embed\python.exe" profile_launcher.py
    exit /b
)

python profile_launcher.py
if %ERRORLEVEL% NEQ 0 (
    py profile_launcher.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Pythonの実行に失敗しました。
    echo 同梱の「0_他のPCでもすぐ動く環境作成.bat」を実行してください。
    pause
)
