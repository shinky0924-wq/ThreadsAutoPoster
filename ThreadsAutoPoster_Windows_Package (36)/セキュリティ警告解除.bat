@echo off
cd /d "%~dp0"
title Threads AutoPoster - Security Unblocker
echo ========================================================
echo   Threads AutoPoster - Security Unblocker
echo ========================================================
echo.
echo Removing Windows Internet download flag (Mark of the Web)...
powershell -NoProfile -Command "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File"
echo.
echo [OK] All files in this folder have been unblocked!
echo You can now safely run build_exe.bat or start_gui.bat.
echo.
pause
