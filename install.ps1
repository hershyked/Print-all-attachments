# Installation script for Print All Attachments Outlook Add-in
# Run this script as Administrator after building the add-in

param(
    [Parameter(Mandatory=$false)]
    [string]$InstallPath = "$env:ProgramFiles\PrintAllAttachments",
    
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to register the add-in in registry
function Register-OutlookAddIn {
    param([string]$dllPath)
    
    Write-Host "Registering add-in in registry..." -ForegroundColor Cyan
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments"
    
    # Create registry key if it doesn't exist
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    # Set registry values
    Set-ItemProperty -Path $regPath -Name "Description" -Value "Print All Attachments Add-in" -Type String
    Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Print All Attachments" -Type String
    Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $regPath -Name "Manifest" -Value "$dllPath.manifest|vstolocal" -Type String
    
    Write-Host "Registry entries created successfully." -ForegroundColor Green
}

# Function to unregister the add-in
function Unregister-OutlookAddIn {
    Write-Host "Unregistering add-in from registry..." -ForegroundColor Cyan
    
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments"
    
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force
        Write-Host "Registry entries removed successfully." -ForegroundColor Green
    } else {
        Write-Host "Add-in was not registered." -ForegroundColor Yellow
    }
}

# Main script
Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Print All Attachments - Installation" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
if (-not (Test-Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Handle uninstall
if ($Uninstall) {
    Write-Host "Uninstalling Print All Attachments..." -ForegroundColor Yellow
    Write-Host ""
    
    # Unregister from registry
    Unregister-OutlookAddIn
    
    # Remove files if they exist
    if (Test-Path $InstallPath) {
        Write-Host "Removing files from $InstallPath..." -ForegroundColor Cyan
        Remove-Item -Path $InstallPath -Recurse -Force
        Write-Host "Files removed successfully." -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Uninstallation complete!" -ForegroundColor Green
    Write-Host "Please restart Outlook for changes to take effect." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 0
}

# Check if build output exists
$buildPath = Join-Path $PSScriptRoot "PrintAllAttachments\bin\Release"
if (-not (Test-Path $buildPath)) {
    Write-Host "ERROR: Build output not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Expected location: $buildPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please run build.bat first to build the add-in." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if DLL exists
$dllPath = Join-Path $buildPath "PrintAllAttachments.dll"
if (-not (Test-Path $dllPath)) {
    Write-Host "ERROR: PrintAllAttachments.dll not found in build output!" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if Outlook is running
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess -and -not $Force) {
    Write-Host "WARNING: Outlook is currently running!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please close Outlook before installing the add-in." -ForegroundColor Yellow
    Write-Host "Or use the -Force parameter to install anyway (not recommended)." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Installing to: $InstallPath" -ForegroundColor Cyan
Write-Host ""

# Create installation directory
if (Test-Path $InstallPath) {
    if (-not $Force) {
        $response = Read-Host "Installation directory already exists. Overwrite? (y/N)"
        if ($response -ne "y" -and $response -ne "Y") {
            Write-Host "Installation cancelled." -ForegroundColor Yellow
            exit 0
        }
    }
    Write-Host "Removing existing installation..." -ForegroundColor Cyan
    Remove-Item -Path $InstallPath -Recurse -Force
}

Write-Host "Creating installation directory..." -ForegroundColor Cyan
New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null

# Copy files
Write-Host "Copying files..." -ForegroundColor Cyan
Copy-Item -Path "$buildPath\*" -Destination $InstallPath -Recurse -Force

# Register the add-in
$installedDllPath = Join-Path $InstallPath "PrintAllAttachments.dll"
Register-OutlookAddIn -dllPath $installedDllPath

Write-Host ""
Write-Host "=======================================" -ForegroundColor Green
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installation location: $InstallPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Start Microsoft Outlook" -ForegroundColor White
Write-Host "  2. Go to File > Options > Add-ins" -ForegroundColor White
Write-Host "  3. Select 'COM Add-ins' from the Manage dropdown" -ForegroundColor White
Write-Host "  4. Click 'Go...'" -ForegroundColor White
Write-Host "  5. Ensure 'PrintAllAttachments' is checked" -ForegroundColor White
Write-Host "  6. Click OK and restart Outlook" -ForegroundColor White
Write-Host ""
Write-Host "The 'Print Attachments' button should appear in the Outlook ribbon." -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
