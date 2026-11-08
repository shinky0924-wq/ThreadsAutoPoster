@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Build Profile Editor EXE

echo ===============================================================
echo   Threads Profile Editor (プロフィール編集.exe) ビルドツール
echo ===============================================================
echo.

pip install pyinstaller
pyinstaller --onefile --noconsole --name "プロフィール編集" profile_editor.py
copy "dist\プロフィール編集.exe" ".\プロフィール編集.exe" /Y
copy "dist\プロフィール編集.exe" ".\profile_editor.exe" /Y

echo.
echo ===============================================================
echo   ✅ ビルド完了！ プロフィール編集.exe が正常に作成されました。
echo ===============================================================
pause
