@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Threads Profile Editor (プロフィール編集ツール)

if exist "プロフィール編集.exe" (
    start "" "プロフィール編集.exe"
    exit /b
)
if exist "profile_editor.exe" (
    start "" "profile_editor.exe"
    exit /b
)

python profile_editor.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 起動に失敗しました。Python環境とPlaywrightを確認してください。
    pause
)
