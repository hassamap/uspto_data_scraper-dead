@echo off
title USPTO Scraper
cd /d "%~dp0"

rem --- Find Python (tries "python" first, then the "py" launcher) ---
set "PY="
python --version >nul 2>&1
if not errorlevel 1 set "PY=python"
if not defined PY (
    py -3 --version >nul 2>&1
    if not errorlevel 1 set "PY=py -3"
)
if not defined PY (
    echo.
    echo Python was not found.
    echo Install Python 3 from https://www.python.org/downloads/
    echo During install, tick "Add python.exe to PATH", then run this file again.
    echo.
    pause
    exit /b 1
)

rem --- One-time setup: packages and the Playwright browser ---
if not exist ".setup_done" (
    echo First-time setup: installing packages. This may take a few minutes...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Package installation failed. Check your internet connection and try again.
        pause
        exit /b 1
    )
    echo Installing the test browser...
    %PY% -m playwright install chromium
    if errorlevel 1 (
        echo.
        echo Browser installation failed. Check your internet connection and try again.
        pause
        exit /b 1
    )
    echo setup complete> ".setup_done"
)

echo.
echo Starting the scraper. Your browser will open automatically.
echo Keep this window open while using it. Closing it stops the app.
echo.
%PY% app.py
echo.
echo The scraper has stopped.
pause
