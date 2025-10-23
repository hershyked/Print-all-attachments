@echo off
REM Check Prerequisites for Print All Attachments

echo ========================================
echo Print All Attachments - Prerequisites Check
echo ========================================
echo.
echo Checking your system for required components...
echo.

set ALL_OK=1

REM Check Windows version
echo [1/5] Checking Windows version...
ver | findstr /i "Windows" >nul
if %ERRORLEVEL% EQU 0 (
    echo   [OK] Windows detected
) else (
    echo   [FAIL] Windows not detected
    set ALL_OK=0
)
echo.

REM Check for .NET Framework 4.7.2
echo [2/5] Checking .NET Framework 4.7.2+...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo   [OK] .NET Framework 4+ detected
) else (
    echo   [WARNING] .NET Framework 4.7.2+ might not be installed
    echo   Download from: https://dotnet.microsoft.com/download/dotnet-framework
    set ALL_OK=0
)
echo.

REM Check for Visual Studio / MSBuild
echo [3/5] Checking for MSBuild (Visual Studio)...
set FOUND_VS=0

if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe" set FOUND_VS=1
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe" set FOUND_VS=1

if %FOUND_VS% EQU 1 (
    echo   [OK] Visual Studio / MSBuild found
) else (
    echo   [FAIL] Visual Studio not found
    echo   Required for building from source
    echo   Download from: https://visualstudio.microsoft.com/downloads/
    echo   Install with "Office/SharePoint development" workload
    set ALL_OK=0
)
echo.

REM Check for Outlook
echo [4/5] Checking for Microsoft Outlook...
set FOUND_OUTLOOK=0

if exist "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" set FOUND_OUTLOOK=1
if exist "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" set FOUND_OUTLOOK=1
if exist "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE" set FOUND_OUTLOOK=1
if exist "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE" set FOUND_OUTLOOK=1
if exist "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE" set FOUND_OUTLOOK=1
if exist "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" set FOUND_OUTLOOK=1

if %FOUND_OUTLOOK% EQU 1 (
    echo   [OK] Microsoft Outlook found
) else (
    echo   [WARNING] Microsoft Outlook not found in common locations
    echo   This add-in requires Outlook Desktop (2013 or later)
    set ALL_OK=0
)
echo.

REM Check for Administrator privileges
echo [5/5] Checking administrator privileges...
net session >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo   [OK] Running with administrator privileges
) else (
    echo   [WARNING] Not running as administrator
    echo   Installation will require administrator privileges
)
echo.

echo ========================================
if %ALL_OK% EQU 1 (
    echo RESULT: All prerequisites met! 
    echo.
    echo You can now run:
    echo   - quick-install.bat  (build and install in one step)
    echo   - build.bat          (build only)
    echo.
) else (
    echo RESULT: Some prerequisites are missing
    echo.
    echo Please install the missing components listed above.
    echo.
)
echo ========================================
echo.
pause
