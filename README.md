# Print All Attachments - Outlook Add-in

[![Build and Release](https://github.com/hershyked/Print-all-attachments/actions/workflows/build.yml/badge.svg)](https://github.com/hershyked/Print-all-attachments/actions/workflows/build.yml)
[![GitHub release](https://img.shields.io/github/v/release/hershyked/Print-all-attachments)](https://github.com/hershyked/Print-all-attachments/releases)

An Outlook VSTO Add-in that allows you to print all attachments from multiple selected emails with a single click.

**✨ NEW: One-click installation scripts - install in minutes without Visual Studio!**

## 🚀 Quick Install

**Want to install right now?** Choose your method:

1. **Pre-built Release** (2 minutes) - Download and run from [Releases](https://github.com/hershyked/Print-all-attachments/releases)
2. **One-Click Script** (5 minutes) - Run `quick-install.bat` as Administrator (requires Visual Studio)
3. **Manual Build** (10 minutes) - Traditional Visual Studio build process

📖 **See [README-INSTALL.md](README-INSTALL.md) for step-by-step instructions**

## Overview

This add-in adds a "Print Attachments" button to the Outlook ribbon. When you select multiple emails and click this button, it will:
- Extract all attachments from the selected emails
- Send each attachment to your default printer
- Show you a summary of how many attachments were printed

## Features

- ✅ Print attachments from multiple selected emails at once
- ✅ Works with any file type that has a default print handler
- ✅ Provides feedback on success/errors
- ✅ Automatically cleans up temporary files
- ✅ Simple one-click operation

## Requirements

### For End Users (Just Installing)

- **Windows Operating System** (Windows 10 or later recommended)
- **Microsoft Outlook** (Desktop version - 2013 or later recommended)
- **.NET Framework 4.7.2** or later (usually already installed on Windows 10+)

**Note:** You do NOT need Visual Studio or any development tools! Just download the pre-built release.

### For Developers (Building from Source)

- **Windows Operating System**
- **Microsoft Outlook** (Desktop version - 2013 or later recommended)
- **Microsoft Visual Studio** (2017 or later) with Office/SharePoint development workload
- **.NET Framework 4.7.2** or later

## Installation

### 🚀 Option 1: Quick Install (Recommended - 5 minutes)

**New! One-click installation with automated scripts:**

1. **Download the repository** (or clone it)
2. **Run as Administrator**: `quick-install.bat`
3. **Enable in Outlook**: File > Options > Add-ins > COM Add-ins > Check "PrintAllAttachments"
4. **Restart Outlook** and you're done! ✨

**Requirements**: Visual Studio 2017+ with Office/SharePoint development workload
**Note**: You don't need to know how to use Visual Studio - the script does everything!

📖 **Full instructions**: See [README-INSTALL.md](README-INSTALL.md)

### 📦 Option 2: Pre-built Release (2 minutes - Easiest)

**When available, this is the fastest method:**

1. **Download**: Go to [Releases](https://github.com/hershyked/Print-all-attachments/releases) → Download `PrintAllAttachments-Release.zip`
2. **Extract**: Unzip to a permanent location (e.g., `C:\Program Files\PrintAllAttachments`)
3. **Install**: Run `install.ps1` as Administrator
4. **Enable in Outlook**: File > Options > Add-ins > COM Add-ins > Check "PrintAllAttachments"
5. **Restart Outlook**

### 🛠️ Option 3: Manual Build (Traditional Method)

**For developers or if you prefer manual control:**

#### Building with Scripts:
```bash
# Build only (run from repository root)
build.bat

# Install after building (run as Administrator)
powershell -ExecutionPolicy Bypass -File install.ps1
```

#### Building with Visual Studio:
1. Install [Visual Studio](https://visualstudio.microsoft.com/) with Office/SharePoint development workload
2. Clone the repository:
   ```bash
   git clone https://github.com/hershyked/Print-all-attachments.git
   cd Print-all-attachments
   ```
3. Open `PrintAllAttachments.sln` in Visual Studio
4. Set configuration to "Release"
5. Build > Build Solution (`Ctrl+Shift+B`)
6. Run `install.ps1` as Administrator to install

### 🔧 Option 4: Development Installation

For developers who want to debug or modify the code:

1. Open the solution in Visual Studio
2. Press `F5` to build and run with debugging
3. Outlook will launch with the add-in loaded automatically
4. Make changes and press `F5` again to test

## Usage

1. **Open Microsoft Outlook**

2. **Select emails:** 
   - In your inbox or any mail folder, select one or more emails that contain attachments
   - You can select multiple emails by holding `Ctrl` and clicking each email

3. **Print attachments:**
   - Look for the "Attachments" group in the Outlook ribbon (on the "Mail" tab)
   - Click the "Print Attachments" button
   - The add-in will process all attachments from the selected emails

4. **Review results:**
   - A message box will appear showing:
     - Number of emails processed
     - Number of attachments printed
     - Any errors encountered

## How It Works

1. The add-in retrieves all selected emails from the active Outlook window
2. For each email, it extracts all attachments to a temporary folder
3. Each attachment is sent to your default printer using Windows' default print handler
4. After printing, the temporary files are automatically deleted
5. A summary dialog shows the results

## Supported File Types

The add-in can print any file type that:
- Has a default application associated with it in Windows
- That application supports the "print" verb

Common supported formats include:
- PDF files (`.pdf`)
- Microsoft Office documents (`.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`)
- Images (`.jpg`, `.png`, `.bmp`, `.tiff`)
- Text files (`.txt`)

**Note:** Some file types may open their associated application briefly during printing. This is normal behavior.

## Troubleshooting

### The add-in doesn't appear in Outlook

1. Check if the add-in is enabled:
   - File > Options > Add-ins
   - Look for "PrintAllAttachments" in the list
   - If it's in the "Disabled Items" list, enable it

2. Check if the add-in is installed:
   - In the Add-ins dialog, check "Manage: COM Add-ins" at the bottom
   - Click "Go..."
   - Ensure "PrintAllAttachments" is checked

### Attachments are not printing

1. Verify your default printer is set up correctly
2. Check that you have the necessary applications to open the file types
3. Some file types may require their associated application to be closed for printing to work

### Permission Issues

- The add-in needs permission to access your Outlook data
- If prompted, grant the necessary permissions

## Development

### Project Structure

```
PrintAllAttachments/
├── PrintAllAttachments.csproj      # Project file
├── ThisAddIn.cs                    # Main add-in class
├── ThisAddIn.Designer.cs           # Designer file
├── PrintAttachmentsRibbon.cs       # Ribbon UI implementation
├── PrintAttachmentsRibbon.Designer.cs  # Ribbon designer
├── PrintAttachmentsRibbon.resx     # Ribbon resources
└── Properties/
    └── AssemblyInfo.cs             # Assembly metadata
```

### Key Components

- **ThisAddIn.cs**: Entry point for the VSTO add-in
- **PrintAttachmentsRibbon.cs**: Contains the UI button and main printing logic
- **PrintFile()**: Method that handles printing individual files

### Customization

You can customize the add-in by modifying:
- Button appearance in `PrintAttachmentsRibbon.Designer.cs`
- Printing behavior in the `PrintFile()` method
- Error handling and user feedback messages

## Security Considerations

- The add-in only accesses attachments from emails you explicitly select
- Temporary files are stored in your Windows temp folder and automatically deleted
- No data is sent to external servers
- The add-in requires trust to run in Outlook (standard for VSTO add-ins)

## License

This project is open source. Feel free to use, modify, and distribute as needed.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Support

For issues, questions, or suggestions, please open an issue on the GitHub repository.

## Changelog

### Version 1.0.0
- Initial release
- Print attachments from multiple selected emails
- Basic error handling and user feedback