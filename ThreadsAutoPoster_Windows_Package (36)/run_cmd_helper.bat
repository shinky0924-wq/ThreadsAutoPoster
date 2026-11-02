@echo off
cd /d "%~dp0"
title Threads AutoPoster - Console
echo ========================================================
echo   Threads AutoPoster - Console Helper
echo   (This window is guaranteed to stay open)
echo ========================================================
echo.
echo Running build_exe.bat...
echo.
call "%~dp0build_exe.bat"
echo.
echo ========================================================
echo Build script finished. This prompt stays open for you.
echo ========================================================
cmd /k
