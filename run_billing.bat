@echo off
cd /d "%~dp0"

title Billing Automation Tool

echo Running Billing Automation...
echo --------------------------------

if exist "C:\Users\mshrinivasan\AppData\Local\Programs\Python\Python313\python.exe" (
    "C:\Users\mshrinivasan\AppData\Local\Programs\Python\Python313\python.exe" billing_prepare.py
) else (
    py billing_prepare.py
)

if %errorlevel% neq 0 (
    echo.
    echo Billing Automation failed.
    pause
    exit /b %errorlevel%
)

echo.
echo --------------------------------
echo Billing Automation completed successfully.
pause