# Usage Guide

## Quick Start

1. **Select emails** - In Outlook, select one or more emails that contain attachments
2. **Click the button** - Find and click the "Print Attachments" button in the ribbon
3. **Done!** - All attachments will be sent to your default printer

## Detailed Instructions

### Step 1: Select Emails with Attachments

In your Outlook inbox or any mail folder:

- **Single email**: Click on an email that has attachments
- **Multiple emails**: 
  - Hold `Ctrl` and click on each email you want to select
  - OR: Click the first email, hold `Shift`, click the last email to select a range
  - OR: Use `Ctrl+A` to select all emails in the folder

**Tip**: Look for the paperclip icon 📎 to identify emails with attachments.

### Step 2: Print Attachments

1. Look at the Outlook ribbon (the toolbar at the top)
2. Find the **"Attachments"** group (it should be on the "Mail" tab)
3. Click the **"Print Attachments"** button

### Step 3: Review Results

After clicking the button:

1. The add-in will process all selected emails
2. For each email, it will extract and print all attachments
3. A dialog box will appear showing:
   - Number of emails processed
   - Number of attachments printed
   - Any errors that occurred

**Example output:**
```
Processed 5 email(s).
Printed 7 attachment(s).
```

## Understanding the Process

Here's what happens when you click "Print Attachments":

1. **Selection**: The add-in retrieves all emails you selected
2. **Extraction**: For each email, all attachments are saved to a temporary folder
3. **Printing**: Each attachment is sent to your default printer
4. **Cleanup**: Temporary files are automatically deleted
5. **Feedback**: You receive a summary of the operation

## Supported File Types

The add-in can print any file that:
- Has a default application in Windows
- That application supports printing

### Commonly Supported Formats

✅ **Documents**
- PDF (`.pdf`)
- Word (`.doc`, `.docx`)
- Excel (`.xls`, `.xlsx`)
- PowerPoint (`.ppt`, `.pptx`)
- Text files (`.txt`)

✅ **Images**
- JPEG (`.jpg`, `.jpeg`)
- PNG (`.png`)
- BMP (`.bmp`)
- TIFF (`.tiff`, `.tif`)
- GIF (`.gif`)

✅ **Other**
- Web pages (`.html`, `.htm`)
- Rich Text (`.rtf`)

### Unsupported or Problematic Files

❌ **May Not Print**
- Compressed archives (`.zip`, `.rar`, `.7z`)
- Executable files (`.exe`, `.msi`)
- Video files (`.mp4`, `.avi`, `.mov`)
- Audio files (`.mp3`, `.wav`)

## Tips and Best Practices

### 1. Verify Your Printer

Before using the add-in:
- Make sure your printer is on and connected
- Check that it's set as the default printer
- Ensure it has paper and ink/toner

**To set default printer:**
1. Windows Settings > Devices > Printers & scanners
2. Select your printer
3. Click "Manage" > "Set as default"

### 2. Preview Before Printing

If you want to verify what will be printed:
1. Manually open one email
2. Save an attachment
3. Open it to see how it looks when printed
4. Then use the add-in for batch printing

### 3. Handle Large Batches

For many emails with attachments:
- **Select in smaller groups** - Process 10-20 emails at a time
- **Monitor printer queue** - Check Windows print spooler if needed
- **Wait between batches** - Give your printer time to process

### 4. Dealing with Different File Types

Some applications may briefly appear when printing:
- **PDF files**: Adobe Reader or default PDF viewer may flash
- **Office files**: Word/Excel may open briefly in the background
- **This is normal** - The application is processing the print job

### 5. Managing Print Settings

The add-in uses your default print settings. To adjust:

1. Open the file type in its native application
2. Go to Print settings (Ctrl+P)
3. Set your preferences (color, orientation, paper size)
4. Save as default for that application
5. These settings will be used when the add-in prints

## Common Scenarios

### Scenario 1: Daily Delivery Notes

**Problem**: You receive 20 delivery note emails every morning, each with one PDF attachment.

**Solution**:
1. Open your inbox
2. Select all delivery note emails (use a search or filter if needed)
3. Click "Print Attachments"
4. All PDFs print automatically

**Time saved**: Instead of 5 minutes of repetitive clicking, it takes 5 seconds!

### Scenario 2: Weekly Reports

**Problem**: Every Friday you need to print reports from multiple departments.

**Solution**:
1. Create a folder for weekly reports
2. Move all report emails to that folder
3. Select all emails in the folder
4. Click "Print Attachments"

### Scenario 3: Meeting Preparation

**Problem**: You need to print all meeting materials from various emails.

**Solution**:
1. Search for emails with a specific subject or date
2. Select relevant emails from the search results
3. Click "Print Attachments"
4. All meeting documents print together

## Keyboard Shortcuts

Speed up your workflow:

- `Ctrl+A` - Select all emails in current view
- `Ctrl+Click` - Add individual emails to selection
- `Shift+Click` - Select a range of emails
- `Alt` then arrow keys - Navigate ribbon with keyboard

## Troubleshooting Common Issues

### Nothing prints after clicking the button

**Check:**
- Is your printer turned on and connected?
- Do the emails actually have attachments?
- Did you get an error message?

### Some attachments didn't print

**Possible causes:**
- Unsupported file type (see [Supported File Types](#supported-file-types))
- Application not installed (e.g., trying to print .docx without Word)
- Printer ran out of paper mid-job

**Solution**: Check the summary dialog for specific errors.

### Application windows keep popping up

**This is normal** for some file types. The associated application needs to open the file to print it. The add-in tries to hide these windows, but some may briefly appear.

### Prints are going to the wrong printer

**Solution**: Set your desired printer as default:
1. Windows Settings > Devices > Printers & scanners
2. Select the printer you want
3. Click "Set as default"

### "Access Denied" or permission errors

**Causes:**
- Outlook security settings
- Antivirus blocking the add-in
- Corrupted temporary files

**Solutions:**
1. Add the add-in to your antivirus whitelist
2. Run Outlook as administrator (one time test)
3. Clear your Windows temp folder

## Advanced Usage

### Processing Specific Senders

1. Use Outlook search: `from:sender@example.com`
2. Select results
3. Print attachments

### Date Range Processing

1. Use Outlook search with date filters
2. Or create a search folder with date criteria
3. Select and print

### Using with Rules

You can create Outlook rules to:
1. Move certain emails to a specific folder
2. Periodically select all in that folder
3. Print attachments in batch

## Privacy and Security

### What the add-in accesses:
- ✅ Only emails you explicitly select
- ✅ Only attachments from those emails

### What the add-in does NOT do:
- ❌ Does not access emails you haven't selected
- ❌ Does not upload or send any data anywhere
- ❌ Does not store attachments permanently
- ❌ Does not read email content (only attachments)

### Temporary files:
- Saved to: `%TEMP%\OutlookAttachments_[random-id]\`
- Automatically deleted after printing
- No permanent storage of your files

## Getting Help

If you need assistance:

1. **Check documentation**:
   - This guide
   - [README](README.md)
   - [Installation Guide](INSTALLATION.md)

2. **Common issues**:
   - See Troubleshooting section above
   - Check printer settings
   - Verify file type support

3. **Still stuck?**
   - Open an issue on [GitHub](https://github.com/hershyked/Print-all-attachments/issues)
   - Include:
     - What you were trying to do
     - What happened instead
     - Any error messages
     - File types involved

## Feedback

Have suggestions for improvement? We'd love to hear from you!

- Feature requests: Open an issue on GitHub
- Bug reports: Open an issue with details
- General feedback: Discussions on GitHub

---

**Happy printing! 🖨️**
