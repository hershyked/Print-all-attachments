# Frequently Asked Questions (FAQ)

## General Questions

### What is Print All Attachments?
Print All Attachments is a Microsoft Outlook add-in that allows you to print all attachments from multiple selected emails with a single click.

### Is this free?
Yes, this is an open-source project available under the MIT License.

### What versions of Outlook are supported?
The add-in works with Microsoft Outlook 2013 and later, including Microsoft 365 Outlook.

### Does this work with Outlook.com or the web version?
No, this is a VSTO add-in designed for the desktop version of Outlook on Windows only.

### Is my data safe?
Yes. The add-in:
- Only accesses emails you explicitly select
- Does not send any data to external servers
- Temporarily saves attachments locally and deletes them after printing
- Does not store or transmit your personal information

## Installation Questions

### Where can I download the add-in?
You can download the latest release from the [Releases page](https://github.com/hershyked/Print-all-attachments/releases) or build it from source.

### Do I need Visual Studio to use this?
No, you only need Visual Studio if you want to build the add-in from source. For regular use, just run the installer.

### Why do I see a security warning during installation?
This is normal for VSTO add-ins. The warning appears because the add-in is not from the Microsoft Store and requires access to Outlook. If you built it from source or downloaded from our official repository, it's safe to proceed.

### The add-in doesn't appear after installation. What should I do?
1. Restart Outlook completely
2. Check File > Options > Add-ins
3. Select "COM Add-ins" from the Manage dropdown
4. Click "Go..." and ensure PrintAllAttachments is checked
5. If it's in the "Disabled Items" list, enable it

### Can I install this on macOS?
No, VSTO add-ins are Windows-only. This add-in requires Windows and desktop Outlook.

## Usage Questions

### How do I use the add-in?
1. Select one or more emails with attachments
2. Click the "Print Attachments" button in the ribbon
3. All attachments will be sent to your default printer

### Can I select which attachments to print?
Currently, the add-in prints all attachments from all selected emails. Selective printing may be added in a future version.

### Can I choose which printer to use?
Currently, attachments are sent to your default printer. Custom printer selection may be added in a future version.

### What file types can be printed?
Any file type that has a default application in Windows that supports printing:
- PDF files
- Microsoft Office documents (Word, Excel, PowerPoint)
- Images (JPEG, PNG, BMP, TIFF, etc.)
- Text files
- And more

### What happens if an email has no attachments?
The add-in simply skips that email and continues with the others.

### Can I print inline images?
No, the add-in only processes file attachments, not inline images embedded in the email body.

## Troubleshooting

### Nothing happens when I click the button
1. Verify you have selected at least one email
2. Check if the selected emails have attachments (look for the paperclip icon)
3. Ensure your default printer is set up and available

### Some files didn't print
Common reasons:
- The file type is not supported (e.g., ZIP files)
- The required application is not installed (e.g., PDF reader for PDFs)
- The printer ran out of paper or had an error
Check the summary dialog for specific error messages.

### I see applications opening briefly
This is normal for some file types. Applications like Adobe Reader or Microsoft Word need to open the file to print it. The add-in tries to hide these windows, but they may briefly appear.

### Print quality is poor
The add-in uses the default print settings for each application. To adjust:
1. Open the file type in its native application
2. Configure print settings (File > Print)
3. Set those as defaults for that application

### The add-in is slow with many emails
Processing time depends on:
- Number of emails selected
- Number and size of attachments
- Printer speed
- System performance

For best results, process 10-20 emails at a time.

### I'm getting "Access Denied" errors
Possible causes:
1. Outlook security settings - check macro security settings
2. Antivirus blocking - add the add-in to your whitelist
3. Windows permissions - ensure your user account can write to the temp folder

### Temporary files are not being deleted
The add-in should automatically clean up. If files remain:
1. Check: `%TEMP%` folder for `OutlookAttachments_*` folders
2. You can safely delete these manually
3. This might indicate the add-in crashed - check Outlook's error logs

## Technical Questions

### What technology is this built with?
- Microsoft Visual Studio Tools for Office (VSTO)
- C# programming language
- .NET Framework 4.7.2
- Microsoft Office Interop libraries

### Can I customize the code?
Yes! It's open source. See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

### How does the printing work?
The add-in:
1. Extracts attachments to a temporary folder
2. Uses Windows Shell Execute with the "print" verb
3. The associated application handles the actual printing
4. Cleans up temporary files afterward

### Why use VSTO instead of Office.js?
VSTO provides:
- Better access to Outlook's object model
- Ability to interact with the Windows printer subsystem
- No server or web hosting requirements
- Better performance for local operations

### Can this be converted to a web add-in?
Printing from web add-ins is limited due to browser security restrictions. VSTO is the better choice for this use case.

## Advanced Topics

### Can I automate this with macros or rules?
Not directly, but you could:
1. Create Outlook rules to move emails to a specific folder
2. Periodically select all in that folder
3. Print attachments manually

### Can I integrate this with other tools?
The add-in is standalone, but you could:
- Modify the source code to add integrations
- Use Windows scripting to automate Outlook selection
- Contact us about custom development

### How do I debug issues?
1. Check Outlook's Trust Center settings
2. Look for error messages in Event Viewer
3. Enable VSTO logging in registry
4. Build and run in Visual Studio for detailed debugging

### Can I contribute code?
Absolutely! See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

### Where can I request features?
Open a feature request on [GitHub Issues](https://github.com/hershyked/Print-all-attachments/issues).

## Performance

### How many emails can I process at once?
There's no hard limit, but for best performance:
- Recommended: 10-20 emails at a time
- Maximum tested: 50+ emails
- Depends on: attachment sizes, printer speed, system resources

### Does it work with large attachments?
Yes, but:
- Large files take longer to print
- Ensure sufficient disk space in temp folder
- Be patient - don't click the button multiple times

### Will this slow down Outlook?
No, the processing happens when you click the button. It doesn't run in the background or affect normal Outlook performance.

## Privacy & Security

### What data does the add-in collect?
None. It doesn't collect, store, or transmit any data.

### Where are attachments saved temporarily?
In your Windows temp folder: `%TEMP%\OutlookAttachments_[unique-id]\`

### Can others on my network see what I'm printing?
Standard network printer security applies. The add-in doesn't add any additional exposure.

### Is the code audited for security?
The code is:
- Open source (you can review it)
- Scanned with CodeQL (no issues found)
- Following Microsoft's security best practices

## Getting Help

### Where can I get more help?
1. Read the [README](README.md)
2. Check [USAGE.md](USAGE.md) for detailed instructions
3. Review [INSTALLATION.md](INSTALLATION.md) for setup help
4. Search existing [GitHub Issues](https://github.com/hershyked/Print-all-attachments/issues)
5. Open a new issue if your question isn't answered

### How do I report a bug?
Open a bug report on [GitHub Issues](https://github.com/hershyked/Print-all-attachments/issues/new) using the bug report template.

### Can I get commercial support?
This is a community project. For commercial support or custom development, contact the repository owner.

---

**Didn't find your answer?** [Open an issue](https://github.com/hershyked/Print-all-attachments/issues/new) with your question!
