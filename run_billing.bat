
Paste:

```bat
@echo off
title Billing Automation

echo ===============================
echo Running Billing Automation...
echo ===============================

cd /d "%~dp0"

echo Checking Python...

python --version

if errorlevel 1 (
    echo Trying py command...
    py --version

    if errorlevel 1 (
        echo ERROR: Python not found
        echo Please reinstall Python and CHECK "Add Python to PATH"
        pause
        exit /b
    )

    py billing_automation.py
) else (
    python billing_automation.py
)

echo.
echo DONE - Check Outlook Drafts
pause