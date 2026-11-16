@echo off
title BSNL IP Billing and SLA Reconciliation System
echo ========================================================
echo   BSNL Telecom Infrastructure (IP-1) Reconciliation
echo   Automated Invoice and Outage SLA Verification
echo ========================================================
echo.

:: 1. Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not added to PATH.
    echo Please install Python 3.10+ from https://www.python.org/
    echo IMPORTANT: Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b
)

echo [OK] Python detected.
echo.

:: 2. Check and install dependencies if needed
echo Checking and installing required packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [WARNING] Encountered issues installing packages. Attempting to run anyway...
)

echo.
echo ========================================================
echo Starting application in your default web browser...
echo (You can close this window when you are done)
echo ========================================================
echo.

python -m streamlit run app.py
pause
