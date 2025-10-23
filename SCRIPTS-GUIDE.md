# Installation Scripts Summary

This document describes all the installation scripts and tools available for Print All Attachments.

## Quick Reference

| Script | Purpose | Requires Admin | Time |
|--------|---------|---------------|------|
| `check-prerequisites.bat` | Check if your system has all requirements | No | 30 sec |
| `quick-install.bat` | Build and install in one step | Yes | 5-10 min |
| `build.bat` | Build the add-in only | No | 3-5 min |
| `install.ps1` | Install after building | Yes | 1 min |
| `uninstall.bat` | Remove the add-in | Yes | 1 min |

## Detailed Descriptions

### check-prerequisites.bat

**Purpose**: Verify your system meets all requirements before installation

**What it checks**:
- ✅ Windows version
- ✅ .NET Framework 4.7.2+
- ✅ Visual Studio / MSBuild
- ✅ Microsoft Outlook
- ✅ Administrator privileges

**When to use**:
- Before attempting to build from source
- When troubleshooting build issues
- To verify installation requirements

**Usage**:
```bash
# Just double-click the file, or:
check-prerequisites.bat
```

**Output**: Color-coded report showing what's installed and what's missing

---

### quick-install.bat

**Purpose**: One-click solution to build and install the add-in

**What it does**:
1. Runs `build.bat` to compile the add-in
2. Runs `install.ps1` to install it
3. Provides next steps for enabling in Outlook

**Requirements**:
- Visual Studio 2017+ with Office development tools
- Administrator privileges

**When to use**:
- When you want the fastest installation from source
- When no pre-built release is available
- For first-time installations

**Usage**:
```bash
# Right-click and select "Run as administrator"
quick-install.bat
```

**Time**: 5-10 minutes depending on your system

---

### build.bat

**Purpose**: Build the add-in from source code

**What it does**:
1. Locates MSBuild automatically
2. Downloads NuGet if needed
3. Restores NuGet packages
4. Builds the Release configuration
5. Reports success or detailed errors

**Requirements**:
- Visual Studio 2017+ with Office development tools
- Internet connection (for NuGet)

**When to use**:
- When you want to build without installing
- When testing code changes
- When you need just the compiled files

**Usage**:
```bash
# Just double-click the file, or:
build.bat
```

**Output**: Build artifacts in `PrintAllAttachments/bin/Release/`

---

### install.ps1

**Purpose**: Install the built add-in to your system

**What it does**:
1. Copies files to Program Files (or custom location)
2. Registers add-in in Windows Registry
3. Configures Outlook to load the add-in

**Requirements**:
- Administrator privileges
- Already-built add-in (run `build.bat` first)

**When to use**:
- After building with `build.bat`
- When reinstalling after uninstalling
- When installing to a custom location

**Usage**:
```powershell
# Run PowerShell as Administrator, then:

# Standard installation
.\install.ps1

# Custom installation path
.\install.ps1 -InstallPath "C:\MyPath"

# Force install (even if Outlook is running - not recommended)
.\install.ps1 -Force

# Uninstall
.\install.ps1 -Uninstall
```

**Options**:
- `-InstallPath`: Custom installation directory
- `-Uninstall`: Remove the add-in
- `-Force`: Skip safety checks (not recommended)

---

### uninstall.bat

**Purpose**: Easy uninstallation of the add-in

**What it does**:
1. Asks for confirmation
2. Removes registry entries
3. Deletes installed files
4. Provides instructions to restart Outlook

**Requirements**:
- Administrator privileges

**When to use**:
- When removing the add-in completely
- When troubleshooting installation issues
- Before installing a new version

**Usage**:
```bash
# Right-click and select "Run as administrator"
uninstall.bat
```

**Effect**: Complete removal of the add-in from your system

---

## Installation Workflows

### First-Time Installation (Recommended)

```
1. check-prerequisites.bat  → Verify system
2. quick-install.bat        → Build and install
3. Enable in Outlook        → File > Options > Add-ins
4. Restart Outlook          → Done!
```

### Build and Install Separately

```
1. build.bat               → Build the add-in
2. install.ps1            → Install (as Administrator)
3. Enable in Outlook      → File > Options > Add-ins
4. Restart Outlook        → Done!
```

### Reinstallation

```
1. uninstall.bat          → Remove old version
2. quick-install.bat      → Install new version
3. Restart Outlook        → Done!
```

### Update Existing Installation

```
1. build.bat              → Build new version
2. install.ps1 -Force     → Overwrite existing
3. Restart Outlook        → Done!
```

## Troubleshooting Script Issues

### "MSBuild not found" (build.bat)

**Problem**: Visual Studio not installed or not detected

**Solutions**:
1. Run `check-prerequisites.bat` to verify installation
2. Install Visual Studio Community (free) with Office development workload
3. Manually specify MSBuild path by editing `build.bat` (advanced)

### "Cannot run scripts" (install.ps1)

**Problem**: PowerShell execution policy restriction

**Solutions**:
```powershell
# Option 1: Run with bypass (one-time)
powershell -ExecutionPolicy Bypass -File install.ps1

# Option 2: Change policy (permanent)
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### "Access denied" (any admin script)

**Problem**: Not running as administrator

**Solutions**:
1. Right-click script → "Run as administrator"
2. Open Command Prompt as Admin, then run script
3. Check User Account Control (UAC) settings

### "Build succeeded but install failed"

**Problem**: Outlook is running or file permissions issue

**Solutions**:
1. Close Outlook completely (check Task Manager)
2. Run `install.ps1 -Force` (use with caution)
3. Check that `PrintAllAttachments/bin/Release/` contains DLL files

## Best Practices

### DO:
- ✅ Run `check-prerequisites.bat` before first installation
- ✅ Close Outlook before installing
- ✅ Run install scripts as Administrator
- ✅ Keep installation files in a permanent location
- ✅ Read error messages carefully

### DON'T:
- ❌ Delete files after installation (they're needed)
- ❌ Install while Outlook is running
- ❌ Run scripts from temporary folders
- ❌ Ignore prerequisite warnings
- ❌ Skip the "Enable in Outlook" step

## Need More Help?

- 📖 [README-INSTALL.md](README-INSTALL.md) - Full installation guide
- 📖 [BUILD.md](BUILD.md) - Detailed build instructions
- 📖 [FAQ.md](FAQ.md) - Common questions
- 🐛 [Report an issue](https://github.com/hershyked/Print-all-attachments/issues)

---

**Version**: 1.0  
**Last Updated**: October 2024
