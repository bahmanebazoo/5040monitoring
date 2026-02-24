@echo off
chcp 65001 >nul 2>&1
title 5040 Monitoring - Analytics Pipeline

echo ══════════════════════════════════════════════════════════
echo   5040 Monitoring - Quality ^& Analytics Pipeline
echo ══════════════════════════════════════════════════════════
echo.

REM ─── تنظیم مسیر ───
cd /d "%~dp0"
cd 5040monitoring-version-1

REM ─── بررسی وجود پایتون ───
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found! Please install Python and add it to PATH.
    echo Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [INFO] Python found:
python --version
echo.

REM --- Setup venv ---
if not exist "venv" (
    echo [..] Creating virtual environment...
    python -m venv venv
    echo [OK] Virtual environment created.
)

call venv\Scripts\activate.bat

REM ─── نصب وابستگی‌ها ───
echo [INFO] Installing / checking dependencies...
pip install openpyxl pandas numpy jdatetime rapidfuzz thefuzz --quiet
if %errorlevel% neq 0 (
    echo [WARNING] Some dependencies may have failed to install.
    echo [INFO] Trying to continue anyway...
)
echo [OK] Dependencies are ready.
echo.

REM ─── اجرای برنامه اصلی ───
echo [INFO] Starting the monitoring pipeline...
echo ──────────────────────────────────────────────────────────
echo.

python main.py

echo.
if %errorlevel% equ 0 (
    echo ══════════════════════════════════════════════════════════
    echo   [DONE] Pipeline completed successfully!
    echo ══════════════════════════════════════════════════════════
) else (
    echo ══════════════════════════════════════════════════════════
    echo   [FAILED] An error occurred during execution.
    echo   Check the messages above for details.
    echo ══════════════════════════════════════════════════════════
)

echo.
pause
