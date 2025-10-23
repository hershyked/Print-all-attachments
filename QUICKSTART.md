# Quick Start Guide

Get started with Print All Attachments in 5 minutes! This is your **simple solution** for printing attachments from Outlook.

## For Users (Just Want to Use It)

### 🚀 NEW: One-Click Install Method (5 minutes)

1. **Download** the repository (click green "Code" button > Download ZIP)
2. **Extract** to a folder
3. **Right-click** `quick-install.bat` > "Run as administrator"
4. **Enable** in Outlook (File > Options > Add-ins > COM Add-ins)
5. **Restart** Outlook

**Prerequisites**: Visual Studio with Office development tools (free Community edition works)

### Step 1: Download (1 minute)
1. Go to [Releases](https://github.com/hershyked/Print-all-attachments/releases)
2. Download the latest `PrintAllAttachments-Release.zip`
   - **No need to download setup.exe** - pre-built binaries are in the ZIP

### Step 2: Install (2 minutes)
1. Extract the ZIP file to a permanent location (e.g., `C:\Program Files\PrintAllAttachments\`)
2. Look for `setup.exe` in the extracted folder and run it
   - OR follow manual installation in [INSTALLATION.md](INSTALLATION.md) if no setup.exe
3. Enable the add-in in Outlook (File > Options > Add-ins > COM Add-ins)
4. Restart Outlook

### Step 3: Use It! (30 seconds)
1. In Outlook, select emails with attachments (look for the 📎 icon)
2. Click "Print Attachments" button in the ribbon
3. Confirm if printing many files
4. Done! Check your printer

**That's it!** 🎉 No more opening each email individually!

**Note:** Pre-built binaries are automatically built and tested with GitHub Actions - no Visual Studio needed!

## For Developers (Want to Build It)

### Prerequisites
- Windows 10+
- Visual Studio 2017+ with Office development tools
- Outlook 2013+

### Quick Build (5 minutes)

```bash
# Clone the repo
git clone https://github.com/hershyked/Print-all-attachments.git
cd Print-all-attachments

# Open in Visual Studio
start PrintAllAttachments.sln

# In Visual Studio:
# 1. Press Ctrl+Shift+B to build
# 2. Press F5 to run with Outlook
```

## Common First-Time Issues

### "Add-in not showing"
→ File > Options > Add-ins > Manage COM Add-ins > Check "PrintAllAttachments"

### "Nothing printing"
→ Check default printer is set and has paper

### "Security warning"
→ Click "Install anyway" - it's safe if from official source

## What It Looks Like

```
Outlook Ribbon:
┌─────────────────────────────────────────┐
│ Mail | View | ... | [Attachments]       │
│                     └─ Print Attachments │
└─────────────────────────────────────────┘
```

## How It Works

```
1. SELECT → 2. CLICK → 3. PRINT
   Emails      Button      Result
   
📧 Email 1   ┌─────────┐   🖨️ Attachment1.pdf
📧 Email 2   │ Print   │   🖨️ Attachment2.docx
📧 Email 3 →│Attach-  │ → 🖨️ Attachment3.xlsx
   ...       │ ments   │      ...
📧 Email N   └─────────┘   🖨️ AttachmentN.png
```

## Next Steps

- 📖 Read the [Full Documentation](README.md)
- 🎯 See [Usage Examples](USAGE.md)
- ❓ Check the [FAQ](FAQ.md)
- 🐛 Report [Issues](https://github.com/hershyked/Print-all-attachments/issues)

## Tips for Success

✅ **DO:**
- Set your default printer before using
- Start with a few emails to test
- Check printer has paper and ink

❌ **DON'T:**
- Select hundreds of emails at once (start small)
- Click the button multiple times (be patient)
- Expect ZIP files to print (they won't)

## Support

Need help? Check:
1. [FAQ](FAQ.md) - Common questions
2. [Troubleshooting](INSTALLATION.md#troubleshooting-installation) - Fix issues
3. [GitHub Issues](https://github.com/hershyked/Print-all-attachments/issues) - Ask questions

---

**Ready?** Download and start printing! 🚀
