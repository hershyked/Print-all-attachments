@echo off
REM Uninstall script for Print All Attachments Outlook Add-in

echo ========================================
echo Print All Attachments - Uninstaller
echo ========================================
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

echo This will remove Print All Attachments from your system.
echo.
set /p confirm="Are you sure you want to uninstall? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo Uninstall cancelled.
    exit /b 0
)

echo.
echo Uninstalling...
echo.

REM Run PowerShell uninstall
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Uninstall

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Uninstallation encountered errors. Please check the messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Uninstallation Complete!
echo ========================================
echo.
echo Please restart Outlook for changes to take effect.
echo.
pause
