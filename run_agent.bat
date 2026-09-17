@echo off
title Certificate Generation Agent

echo ============================================================
echo         CERTIFICATE GENERATION AGENT FOR WINDOWS
echo ============================================================
echo.

where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not added to PATH.
    echo Please install Python 3.10+ and add it to PATH.
    pause
    exit /b 1
)

echo [INFO] Checking required Python dependencies...
python -c "import docx, openpyxl" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Installing required dependencies...
    python -m pip install -r requirements.txt
)

echo [OK] Dependencies verified.
echo.

if "%~1"=="" (
    echo [INFO] Launching Certificate Generation GUI...
    python gui.py
) else (
    echo [INFO] Running in Command-Line Mode...
    python mail_merge.py %*
)