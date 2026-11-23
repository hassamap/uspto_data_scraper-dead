@echo off
title USPTO Scraper - Setup
cd /d "%~dp0"

echo ============================================================
echo  USPTO Scraper - one-time setup
echo  Installs everything the scraper needs into this folder.
echo  No administrator rights are needed.
echo  It can take several minutes. Keep this window open.
echo ============================================================
echo.

rem --- Find Python (tries "python" first, then the "py" launcher) ---
set "PY="
python --version >nul 2>&1
if not errorlevel 1 set "PY=python"
if not defined PY (
    py -3 --version >nul 2>&1
    if not errorlevel 1 set "PY=py -3"
)
if not defined PY (
    echo [ERROR] Python was not found on this computer.
    echo.
    echo  1. Go to https://www.python.org/downloads/ and install Python 3.
    echo  2. During installation, tick "Add python.exe to PATH".
    echo  3. Run setup.bat again.
    echo.
    pause
    exit /b 1
)
echo [1/3] Found Python:
%PY% --version
echo.

echo [2/3] Creating a private Python environment in this folder...
if not exist ".venv\Scripts\python.exe" (
    %PY% -m venv .venv
    if errorlevel 1 (
        echo.
        echo [ERROR] Could not create the Python environment.
        echo  Reinstall Python from python.org and tick "Add python.exe to PATH".
        echo.
        pause
        exit /b 1
    )
)
echo Environment ready.
echo.

echo [3/3] Installing packages and the scraper browser...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul 2>&1
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Package installation failed.
    echo  Check your internet connection, then run setup.bat again.
    echo.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 (
    echo.
    echo [ERROR] Browser installation failed.
    echo  Check your internet connection, then run setup.bat again.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Setup complete!
echo  Next time, double-click run.bat to start the scraper.
echo ============================================================
echo.
pause
