@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Build Profile Launcher EXE

echo ===============================================================
echo   Profile Launcher EXE ビルドツール (PyInstaller)
echo ===============================================================
echo.

pip install pyinstaller
pyinstaller --onefile --noconsole --name profile_launcher profile_launcher.py
copy dist\profile_launcher.exe .\profile_launcher.exe /Y
copy dist\profile_launcher.exe .\profile_luncher.exe /Y

echo.
echo ===============================================================
echo   ✅ ビルド完了！ profile_launcher.exe / profile_luncher.exe が作成されました。
echo ===============================================================
pause
