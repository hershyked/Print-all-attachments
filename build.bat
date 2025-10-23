@echo off
REM Build script for Print All Attachments Outlook Add-in
REM This script simplifies building the add-in without opening Visual Studio

echo ========================================
echo Print All Attachments - Build Script
echo ========================================
echo.

REM Check if running from the correct directory
if not exist "PrintAllAttachments.sln" (
    echo ERROR: PrintAllAttachments.sln not found!
    echo Please run this script from the repository root directory.
    echo.
    pause
    exit /b 1
)

echo Looking for MSBuild...
echo.

REM Try to find MSBuild in common locations
set MSBUILD_PATH=

REM Visual Studio 2022
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe
    goto :found_msbuild
)
if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe
    goto :found_msbuild
)
if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe
    goto :found_msbuild
)

REM Visual Studio 2019
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe
    goto :found_msbuild
)
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe
    goto :found_msbuild
)
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe
    goto :found_msbuild
)

REM Visual Studio 2017
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe
    goto :found_msbuild
)
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe
    goto :found_msbuild
)

REM Try using vswhere to locate MSBuild
where /q vswhere
if %ERRORLEVEL% EQU 0 (
    for /f "usebackq tokens=*" %%i in (`vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`) do (
        set MSBUILD_PATH=%%i
        goto :found_msbuild
    )
)

REM MSBuild not found
echo ERROR: MSBuild not found!
echo.
echo Please install one of the following:
echo   1. Visual Studio 2017 or later with .NET desktop development workload
echo   2. Build Tools for Visual Studio
echo.
echo Download from: https://visualstudio.microsoft.com/downloads/
echo.
pause
exit /b 1

:found_msbuild
echo Found MSBuild: %MSBUILD_PATH%
echo.

REM Check for NuGet
echo Looking for NuGet...
where /q nuget
if %ERRORLEVEL% NEQ 0 (
    echo NuGet not found in PATH. Attempting to download...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile 'nuget.exe'}"
    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Failed to download NuGet!
        echo Please download NuGet manually from https://www.nuget.org/downloads
        pause
        exit /b 1
    )
    set NUGET_PATH=nuget.exe
) else (
    set NUGET_PATH=nuget
)

echo Found NuGet
echo.

REM Restore NuGet packages
echo Restoring NuGet packages...
"%NUGET_PATH%" restore PrintAllAttachments.sln
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: NuGet restore failed!
    pause
    exit /b 1
)
echo.

REM Build the solution
echo Building solution (Release configuration)...
echo.
"%MSBUILD_PATH%" PrintAllAttachments.sln /p:Configuration=Release /p:Platform="Any CPU" /v:m
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo.
    echo Please check the error messages above.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo BUILD SUCCESSFUL!
echo ========================================
echo.
echo Build output location:
echo   PrintAllAttachments\bin\Release\
echo.
echo Next steps:
echo   1. Run install.ps1 to install the add-in
echo   2. Or manually copy files to your desired location
echo.
pause
