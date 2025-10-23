# Installation Guide

## 🚀 Quick Install (NEW - Easiest Method!)

**Install in 5 minutes with automated scripts - no Visual Studio UI needed!**

### Prerequisites
- Windows 10 or later
- Visual Studio 2017+ with Office/SharePoint development workload installed
  - Don't worry! You don't need to know how to use Visual Studio - the scripts do everything

### Installation Steps

1. **Download the repository**:
   - Click the green "Code" button on GitHub → Download ZIP
   - Or clone: `git clone https://github.com/hershyked/Print-all-attachments.git`
   - Extract to a folder (e.g., `C:\PrintAllAttachments`)

2. **Run the installer**:
   - Navigate to the extracted folder
   - Right-click `quick-install.bat`
   - Select "Run as administrator"
   - Wait for build and installation (5-10 minutes)

3. **Enable in Outlook**:
   - Open Outlook
   - Go to **File > Options > Add-ins**
   - Select "COM Add-ins" from the Manage dropdown → Click **Go...**
   - Check the box next to **PrintAllAttachments** → Click **OK**

4. **Restart Outlook** - Done! ✨

The "Print Attachments" button should now appear in your Outlook ribbon.

### What the Script Does
- ✅ Automatically finds MSBuild
- ✅ Builds the add-in (Release configuration)
- ✅ Copies files to Program Files
- ✅ Registers with Outlook
- ✅ Provides clear next steps

---

## Quick Install (Recommended - No Build Tools Required!)

**This is the easiest way to install the add-in. You don't need Visual Studio or any build tools!**

### Step-by-Step Installation

1. **Download the pre-built release:**
   - Go to the [Releases](https://github.com/hershyked/Print-all-attachments/releases) page
   - Click on the latest release (highest version number)
   - Download `PrintAllAttachments-Release.zip`

2. **Extract the files:**
   - Right-click the downloaded ZIP file
   - Select "Extract All..."
   - Choose a location (e.g., `C:\PrintAllAttachments`)
   - Click "Extract"

3. **Install the add-in:**
   - Open the extracted folder
   - Look for `setup.exe` (if available) and run it
   - OR manually register the DLL (see Manual Installation section below)
   - Follow any installation prompts

4. **Enable in Outlook:**
   - Open Microsoft Outlook
   - Go to **File > Options > Add-ins**
   - At the bottom, select "COM Add-ins" from the Manage dropdown
   - Click **Go...**
   - Check the box next to **PrintAllAttachments**
   - Click **OK**

5. **Restart Outlook:**
   - Close Outlook completely
   - Open Outlook again
   - The "Print Attachments" button should appear in the ribbon

**Note:** Pre-built releases are automatically built and tested by our CI/CD pipeline on GitHub Actions.

### Manual Installation (Alternative Method)

If a setup.exe is not available in the release, you can manually register the add-in:

1. **Extract the release files** to a permanent location (e.g., `C:\Program Files\PrintAllAttachments\`)
   - Do NOT delete these files after installation - they need to stay on your computer

2. **Register the add-in manually:**
   - Open Command Prompt as Administrator
   - Navigate to the extracted folder
   - Run the following command (adjust path as needed):
     ```cmd
     reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments" /v "Description" /t REG_SZ /d "Print All Attachments Add-in" /f
     reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments" /v "FriendlyName" /t REG_SZ /d "Print All Attachments" /f
     reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments" /v "LoadBehavior" /t REG_DWORD /d 3 /f
     reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments" /v "Manifest" /t REG_SZ /d "C:\Program Files\PrintAllAttachments\PrintAllAttachments.dll.manifest|vstolocal" /f
     ```
   - **Note:** Replace `C:\Program Files\PrintAllAttachments\` with your actual installation path

3. **Trust the add-in location:**
   - Open **File Explorer**
   - Navigate to the folder containing the DLL
   - Right-click the folder > Properties > Security
   - Ensure your user account has Read and Execute permissions

4. **Restart Outlook** to load the add-in

## Building from Source

### Prerequisites

Before building, ensure you have:

1. **Windows 10 or later**
2. **Microsoft Visual Studio 2017 or later** with:
   - Office/SharePoint development workload
   - .NET desktop development workload
3. **Microsoft Outlook** (2013 or later)
4. **.NET Framework 4.7.2** or later

### Installation Steps

#### Step 1: Install Visual Studio

If you don't have Visual Studio:

1. Download [Visual Studio Community](https://visualstudio.microsoft.com/downloads/) (free)
2. During installation, select:
   - "Office/SharePoint development" workload
   - ".NET desktop development" workload
3. Complete the installation

#### Step 2: Clone the Repository

```bash
git clone https://github.com/hershyked/Print-all-attachments.git
cd Print-all-attachments
```

#### Step 3: Open and Build the Project

1. Open `PrintAllAttachments.sln` in Visual Studio
2. In Visual Studio:
   - Select **Release** configuration from the dropdown (top toolbar)
   - Click **Build > Build Solution** (or press `Ctrl+Shift+B`)
3. Wait for the build to complete

#### Step 4: Publish the Add-in

1. Right-click the **PrintAllAttachments** project in Solution Explorer
2. Select **Publish**
3. Choose a publish location (e.g., a folder on your desktop)
4. Click **Finish**
5. Click **Publish Now**

#### Step 5: Install the Add-in

1. Navigate to the publish folder you specified
2. Run `setup.exe`
3. If you see a security warning, click **Install** (this is safe - it's your own code)
4. Wait for installation to complete
5. Close Outlook if it's running
6. Restart Outlook

#### Step 6: Verify Installation

1. Open Outlook
2. Click on any mail folder
3. Look for the **"Attachments"** group in the ribbon
4. You should see a **"Print Attachments"** button

## Development Installation

For development and testing:

1. Open the solution in Visual Studio
2. Press **F5** to build and run
3. Outlook will launch with the add-in loaded
4. Make changes to the code as needed
5. Press **Shift+F5** to stop debugging

## Troubleshooting Installation

### Visual Studio is missing Office development tools

1. Open Visual Studio Installer
2. Click **Modify** next to your Visual Studio installation
3. Check the **Office/SharePoint development** workload
4. Click **Modify** to install

### The add-in doesn't appear after installation

1. Close Outlook completely (check Task Manager to ensure no Outlook processes are running)
2. Open Outlook
3. Go to **File > Options > Add-ins**
4. At the bottom, select **COM Add-ins** from the Manage dropdown
5. Click **Go...**
6. Ensure **PrintAllAttachments** is checked
7. Click **OK**

### Security warnings during installation

This is normal for VSTO add-ins. The warning appears because:
- The add-in is not from the Microsoft Store
- It requires access to Outlook

To proceed:
1. Click **More info** on the warning
2. Click **Install anyway**

### The add-in is listed as disabled

1. Go to **File > Options > Add-ins**
2. Look in the **Disabled Application Add-ins** section
3. If you see **PrintAllAttachments**, select **Disabled Items** from the Manage dropdown
4. Click **Go...**
5. Select the add-in and click **Enable**
6. Restart Outlook

## Uninstallation

### Using the Uninstall Script (Easiest)

```bash
# Right-click and "Run as administrator"
uninstall.bat
```

Or using PowerShell:
```powershell
# Run PowerShell as Administrator
.\install.ps1 -Uninstall
```

### Windows 10/11

1. Open **Settings**
2. Go to **Apps > Installed apps**
3. Find **PrintAllAttachments**
4. Click the three dots and select **Uninstall**
5. Follow the prompts

### Alternative Method

1. Open **Control Panel**
2. Go to **Programs > Programs and Features**
3. Find **PrintAllAttachments**
4. Click **Uninstall**
5. Follow the prompts

## Manual Registry Cleanup (Advanced)

If you need to completely remove all traces:

**⚠️ Warning: Editing the registry can be dangerous. Proceed with caution.**

1. Press `Win+R`, type `regedit`, press Enter
2. Navigate to:
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\PrintAllAttachments
   ```
3. Delete the `PrintAllAttachments` key if it exists
4. Restart Outlook

## System Requirements

- **OS**: Windows 7 SP1 or later (Windows 10/11 recommended)
- **Office**: Microsoft Outlook 2013 or later (Desktop version)
- **Framework**: .NET Framework 4.7.2 or later
- **RAM**: 2 GB minimum (4 GB recommended)
- **Disk Space**: 50 MB for installation

## Getting Help

If you encounter issues during installation:

1. Check the [Troubleshooting](#troubleshooting-installation) section above
2. Review the [README](README.md) for additional information
3. Open an issue on [GitHub](https://github.com/hershyked/Print-all-attachments/issues)
