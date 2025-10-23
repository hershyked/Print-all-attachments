@echo off
REM Quick Install Script - Builds and Installs in One Step
REM Run this script to build and install the add-in automatically

echo ========================================
echo Print All Attachments - Quick Install
echo ========================================
echo.
echo This script will:
echo   1. Build the add-in
echo   2. Install it to your system
echo   3. Register it with Outlook
echo.

REM Check for administrator privileges
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo ERROR: This script requires administrator privileges!
    echo.
    echo Right-click this file and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo Step 1: Building the add-in...
echo.
call build.bat
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Build failed! Please check the errors above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo.
echo Step 2: Installing the add-in...
echo.

REM Run PowerShell installation script
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Force

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Installation failed! Please check the errors above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Quick Install Complete!
echo ========================================
echo.
echo Please restart Outlook to see the Print Attachments button.
echo.
pause
