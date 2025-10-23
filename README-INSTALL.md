# Quick Installation Guide

**⚡ Get the add-in installed in just a few minutes!**

This guide shows you the fastest way to get Print All Attachments working in Outlook.

## 📊 Choose Your Installation Method

```
┌─────────────────────────────────────────────────────────┐
│  Which installation method is right for you?            │
└─────────────────────────────────────────────────────────┘
                         │
         ┌───────────────┴───────────────┐
         │                               │
         v                               v
   ┌─────────┐                     ┌──────────┐
   │ Release │                     │ Building │
   │Available?│                     │ Required?│
   └─────────┘                     └──────────┘
         │ YES                           │ YES
         v                               v
  ┌──────────────┐              ┌────────────────┐
  │ OPTION 1     │              │ OPTION 2       │
  │ Pre-built    │              │ Quick Install  │
  │ Release      │              │ Script         │
  │ (2 minutes)  │              │ (5 minutes)    │
  └──────────────┘              └────────────────┘
         │                               │
         v                               v
  Download ZIP                    Run quick-install.bat
         │                               │
         v                               v
  Run install.ps1                 Automated build
         │                               │
         v                               v
  Enable in Outlook               Enable in Outlook
         │                               │
         └───────────┬───────────────────┘
                     v
              ✅ READY TO USE!
```

## Prerequisites

Before you start, make sure you have:
- ✅ Windows 10 or later
- ✅ Microsoft Outlook (Desktop version - 2013 or later)
- ✅ .NET Framework 4.7.2 or later (usually pre-installed on Windows 10+)

## 🚀 Quick Install (Easiest)

**⚡ Install in 5 minutes with automated scripts!**

### Installation Flow

```
1. Download Repository
   ↓
2. Check Prerequisites (optional)
   Run: check-prerequisites.bat
   ↓
3. Run Quick Install
   Run as Admin: quick-install.bat
   ↓
4. Enable in Outlook
   File > Options > Add-ins > COM Add-ins
   ↓
5. Restart Outlook
   ↓
✅ Ready to Print Attachments!
```

### What You Need

Before you start, make sure you have:
- ✅ Windows 10 or later
- ✅ Microsoft Outlook (Desktop version - 2013 or later)
- ✅ .NET Framework 4.7.2 or later (usually pre-installed on Windows 10+)

## Option 1: Download Pre-built Release (Easiest)

**If available, this is the fastest method!**

1. **Download**: Go to [Releases](https://github.com/hershyked/Print-all-attachments/releases) and download the latest `PrintAllAttachments-Release.zip`

2. **Extract**: Unzip to a permanent location (e.g., `C:\Program Files\PrintAllAttachments`)

3. **Install**: Run `install.ps1` as Administrator
   - Right-click `install.ps1`
   - Select "Run with PowerShell" (as Administrator)
   - Follow the prompts

4. **Enable in Outlook**:
   - Open Outlook
   - File > Options > Add-ins > Manage: COM Add-ins > Go...
   - Check "PrintAllAttachments" > OK
   - Restart Outlook

**Done!** The "Print Attachments" button should appear in your Outlook ribbon.

## Option 2: One-Click Build and Install (No Visual Studio UI Needed)

**If no pre-built release is available, use this method:**

### What You Need
- Visual Studio 2017 or later (Community edition is free) with:
  - Office/SharePoint development workload
  - .NET desktop development workload

**💡 Tip**: Run `check-prerequisites.bat` to verify your system has everything needed!

**Don't worry!** You don't need to know how to use Visual Studio - the script handles everything.

### Installation Steps

1. **Download the repository**:
   - Click the green "Code" button on GitHub
   - Select "Download ZIP"
   - Extract to a folder (e.g., `C:\PrintAllAttachments`)

2. **(Optional) Check prerequisites**:
   - Double-click `check-prerequisites.bat`
   - Review the results
   - Install any missing components

3. **Run the quick install script**:
   - Open the extracted folder
   - Right-click `quick-install.bat`
   - Select "Run as administrator"
   - Wait for the build and installation to complete

4. **Enable in Outlook**:
   - Open Outlook
   - File > Options > Add-ins > Manage: COM Add-ins > Go...
   - Check "PrintAllAttachments" > OK
   - Restart Outlook

**That's it!** ✨

### What the Script Does
The `quick-install.bat` script automatically:
- ✅ Finds MSBuild on your system
- ✅ Downloads NuGet if needed
- ✅ Restores dependencies
- ✅ Builds the add-in
- ✅ Copies files to Program Files
- ✅ Registers the add-in with Outlook

## Option 3: Manual Build (Advanced Users)

If you prefer to build manually:

1. **Build**:
   ```bash
   # Double-click or run from command line
   build.bat
   ```

2. **Install**:
   ```powershell
   # Run PowerShell as Administrator
   .\install.ps1
   ```

## Troubleshooting

### Build Issues

#### "MSBuild not found" error
- **Cause**: Visual Studio not installed or not in expected location
- **Fix**: 
  1. Install [Visual Studio](https://visualstudio.microsoft.com/downloads/) (Community edition is free)
  2. During installation, select these workloads:
     - ✅ Office/SharePoint development
     - ✅ .NET desktop development
  3. Run `build.bat` again

#### "NuGet restore failed" error
- **Cause**: Network issues or NuGet not accessible
- **Fix**: 
  ```bash
  # Download NuGet manually
  # Then run from command prompt:
  nuget.exe restore PrintAllAttachments.sln
  ```

#### "Build failed" with reference errors
- **Cause**: Missing Office development components
- **Fix**: 
  1. Open Visual Studio Installer
  2. Click "Modify" on your installation
  3. Ensure "Office/SharePoint development" is checked
  4. Click "Modify" to install missing components

### Installation Issues

#### "Cannot run scripts" error in PowerShell
- **Cause**: Visual Studio not installed or missing components
- **Fix**: Install Visual Studio with Office/SharePoint development workload
  - Download from: https://visualstudio.microsoft.com/downloads/

### "Cannot run scripts" error in PowerShell
- **Cause**: PowerShell execution policy restriction
- **Fix**: Run this command in PowerShell (as Administrator):
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Add-in doesn't appear in Outlook
1. Close Outlook completely (check Task Manager)
2. File > Options > Add-ins
3. If in "Disabled Items", enable it
4. If not listed, run `install.ps1` again

### "Access denied" error
- **Cause**: Not running as Administrator
- **Fix**: Right-click the script and select "Run as administrator"

## Uninstallation

To remove the add-in:

```powershell
# Run PowerShell as Administrator
.\install.ps1 -Uninstall
```

Or manually:
1. Outlook > File > Options > Add-ins > COM Add-ins > Go...
2. Uncheck "PrintAllAttachments"
3. Delete the installation folder (`C:\Program Files\PrintAllAttachments`)

## Need Help?

- 📖 Read the [Full Documentation](README.md)
- ❓ Check the [FAQ](FAQ.md)
- 🐛 [Report an Issue](https://github.com/hershyked/Print-all-attachments/issues)

## Time Estimates

| Method | Time Required | Prerequisites |
|--------|--------------|---------------|
| Pre-built Release | **2-3 minutes** | None (just Outlook) |
| Quick Install Script | **5-10 minutes** | Visual Studio installed |
| Manual Build | **10-15 minutes** | Visual Studio + knowledge |

---

**Ready to print those attachments?** 🖨️ Choose your method above and get started!
