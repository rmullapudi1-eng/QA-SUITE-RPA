@echo off
title QA Suite RPA

echo Killing any existing instances...
taskkill /F /IM python.exe >nul 2>&1
timeout /t 2 >nul

cd /d "%~dp0"

echo Starting QA Suite App...
echo Open http://localhost:8052 in your browser
echo.

"C:\Program Files\Python311\python.exe" app_ui.py

echo.
echo App stopped. Press any key to close.
pause
