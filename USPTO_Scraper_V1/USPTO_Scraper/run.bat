@echo off
title USPTO Scraper
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo The scraper is not set up yet.
    echo Double-click setup.bat first, wait for "Setup complete!", then run this again.
    echo.
    pause
    exit /b 1
)

echo Starting the scraper. Your browser will open automatically.
echo Keep this window open while using it. Closing it stops the app.
echo.
".venv\Scripts\python.exe" app.py
echo.
echo The scraper has stopped.
pause
