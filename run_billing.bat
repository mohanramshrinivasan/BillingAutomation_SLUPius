@echo off
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed.
    pause
    exit /b 1
)

python billing_prepare.py
echo.
pause