# Quick Start Guide

Get started with Print All Attachments in 5 minutes!

## For Users (Just Want to Use It)

### Step 1: Download (1 minute)
1. Go to [Releases](https://github.com/hershyked/Print-all-attachments/releases)
2. Download the latest `setup.exe`

### Step 2: Install (2 minutes)
1. Run `setup.exe`
2. Click "Install" when prompted
3. Wait for installation to complete
4. Close and restart Outlook

### Step 3: Use It! (30 seconds)
1. In Outlook, select emails with attachments
2. Click "Print Attachments" button in the ribbon
3. Done! Check your printer

**That's it!** 🎉

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
