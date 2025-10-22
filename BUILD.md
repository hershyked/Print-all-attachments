# Building and Testing Guide

This guide covers how to build, test, and debug the Print All Attachments add-in.

## For End Users: No Building Required!

**⭐ If you just want to use the add-in, you don't need to build it yourself!**

We provide pre-built releases that you can download and install directly:
- Go to the [Releases page](https://github.com/hershyked/Print-all-attachments/releases)
- Download the latest `PrintAllAttachments-Release.zip`
- See [INSTALLATION.md](INSTALLATION.md) for installation instructions

The following sections are for developers who want to:
- Contribute to the project
- Modify the source code
- Debug issues
- Build from source

## Automated Builds

This project uses GitHub Actions to automatically build the add-in on every commit. You can:

1. **Download build artifacts:**
   - Go to the [Actions](https://github.com/hershyked/Print-all-attachments/actions) tab
   - Click on any successful workflow run
   - Download the build artifacts (Debug or Release)

2. **View build status:**
   - The build status badge shows if the latest build passed
   - Builds run on Windows using MSBuild
   - Both Debug and Release configurations are built

3. **Automatic releases:**
   - When a version tag (e.g., `v1.0.0`) is pushed, a release is automatically created
   - Release artifacts are attached to the GitHub release
   - Users can download pre-built binaries from the Releases page

## Prerequisites

### Required Software

1. **Operating System**
   - Windows 10 or later (Windows 11 recommended)
   - 64-bit recommended

2. **Development Tools**
   - [Visual Studio 2017 or later](https://visualstudio.microsoft.com/)
   - Required workloads:
     - Office/SharePoint development
     - .NET desktop development

3. **Runtime Requirements**
   - .NET Framework 4.7.2 or later
   - Microsoft Outlook 2013 or later

### Installing Visual Studio with Required Components

1. Download Visual Studio (Community edition is free)
2. Run the installer
3. Select these workloads:
   - ✅ Office/SharePoint development
   - ✅ .NET desktop development
4. Click "Install" and wait for completion

## Building from Source

### Step 1: Clone the Repository

```bash
git clone https://github.com/hershyked/Print-all-attachments.git
cd Print-all-attachments
```

### Step 2: Open the Solution

```bash
# Windows Command Prompt
start PrintAllAttachments.sln

# Or double-click PrintAllAttachments.sln in File Explorer
```

### Step 3: Restore Dependencies

Visual Studio should automatically restore dependencies. If not:

1. Right-click the solution in Solution Explorer
2. Select "Restore NuGet Packages"

### Step 4: Build the Project

#### Using Visual Studio IDE

1. Select build configuration:
   - **Debug**: For development and testing
   - **Release**: For production deployment

2. Build the solution:
   - Menu: Build → Build Solution
   - Keyboard: `Ctrl+Shift+B`
   - Right-click solution → Build

#### Using Command Line (MSBuild)

```bash
# Navigate to solution directory
cd /path/to/Print-all-attachments

# Build Debug configuration
msbuild PrintAllAttachments.sln /p:Configuration=Debug

# Build Release configuration
msbuild PrintAllAttachments.sln /p:Configuration=Release
```

### Step 5: Verify Build

Check the output:
```
Build succeeded.
    0 Warning(s)
    0 Error(s)

Time Elapsed 00:00:05.12
```

Build artifacts location:
- Debug: `PrintAllAttachments/bin/Debug/`
- Release: `PrintAllAttachments/bin/Release/`

## Running and Debugging

### Debug Mode (F5)

1. **Start Debugging**:
   - Press `F5` or click Debug → Start Debugging
   - Outlook will launch with the add-in loaded

2. **What Happens**:
   - Visual Studio compiles the add-in
   - Registers it temporarily with Outlook
   - Launches Outlook with debugging attached
   - You can set breakpoints and inspect variables

3. **Stop Debugging**:
   - Press `Shift+F5` or click Debug → Stop Debugging
   - Outlook will close

### Run Without Debugging (Ctrl+F5)

- Starts Outlook without debugger attached
- Faster startup
- Can't hit breakpoints

### Setting Breakpoints

1. Click in the left margin next to a line of code
2. Or press `F9` on a line
3. When execution hits the breakpoint, code pauses
4. Inspect variables, step through code

Example locations for breakpoints:
```csharp
// In PrintAttachmentsRibbon.cs
private void btnPrintAttachments_Click(object sender, RibbonControlEventArgs e)
{
    // Set breakpoint here to debug button clicks
    try
    {
        Outlook.Application outlookApp = Globals.ThisAddIn.Application;
        // Set breakpoint here to inspect Outlook app
```

## Testing

Since this is a VSTO add-in, testing is primarily manual. Automated testing is limited but possible.

### Manual Testing Checklist

#### Basic Functionality Tests

- [ ] **Add-in Loads**
  1. Start Outlook
  2. Verify button appears in ribbon
  3. Button should be in "Mail" tab under "Attachments" group

- [ ] **Single Email with One Attachment**
  1. Select an email with one PDF attachment
  2. Click "Print Attachments"
  3. Verify attachment prints
  4. Check success message shows "1 email, 1 attachment"

- [ ] **Single Email with Multiple Attachments**
  1. Select an email with 3+ attachments
  2. Click "Print Attachments"
  3. Verify all attachments print
  4. Check success message shows correct counts

- [ ] **Multiple Emails**
  1. Select 5 emails with various attachments
  2. Click "Print Attachments"
  3. Verify all attachments from all emails print
  4. Check summary is accurate

- [ ] **Email with No Attachments**
  1. Select an email without attachments
  2. Click "Print Attachments"
  3. Should process without error
  4. Success message should show 0 attachments

#### File Type Tests

Test with different file types:

- [ ] **PDF Files** (.pdf)
  - Create test email with PDF attachment
  - Verify prints correctly

- [ ] **Word Documents** (.doc, .docx)
  - Create test email with Word attachment
  - Verify prints correctly

- [ ] **Excel Spreadsheets** (.xls, .xlsx)
  - Create test email with Excel attachment
  - Verify prints correctly

- [ ] **Images** (.jpg, .png)
  - Create test email with image attachments
  - Verify images print

- [ ] **Unsupported Files** (.zip, .exe)
  - Create test email with ZIP file
  - Should show error for that attachment
  - Should continue with other attachments

#### Error Handling Tests

- [ ] **No Selection**
  1. Don't select any emails
  2. Click "Print Attachments"
  3. Should show "Please select emails" message

- [ ] **Printer Offline**
  1. Turn off printer or set offline
  2. Try to print attachments
  3. Should handle gracefully (may show printer errors)

- [ ] **Large Attachments**
  1. Email with 10+ MB attachment
  2. Should print (may take time)
  3. Should clean up temp files

#### Performance Tests

- [ ] **Many Emails**
  - Test with 20 emails
  - Should complete within reasonable time
  - Should not freeze Outlook

- [ ] **Large Batch**
  - Test with 50+ attachments
  - Monitor memory usage
  - Verify temp files cleaned up

### Creating Test Data

#### Script to Create Test Emails

```vba
' VBA script to create test emails in Outlook
' Tools → Macro → Visual Basic Editor → Paste this code

Sub CreateTestEmails()
    Dim outlookApp As Outlook.Application
    Dim mailItem As Outlook.MailItem
    Dim i As Integer
    
    Set outlookApp = Application
    
    ' Create 5 test emails
    For i = 1 To 5
        Set mailItem = outlookApp.CreateItem(olMailItem)
        With mailItem
            .Subject = "Test Email " & i & " with Attachments"
            .Body = "This is a test email for Print All Attachments add-in."
            .To = "test@example.com"
            
            ' Add test files as attachments
            ' Replace with your test file paths
            ' .Attachments.Add "C:\TestFiles\test.pdf"
            ' .Attachments.Add "C:\TestFiles\test.docx"
            
            .Save
            .Move Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
        End With
    Next i
    
    MsgBox "Created " & i - 1 & " test emails"
End Sub
```

### Debugging Tips

#### View Debug Output

In Visual Studio:
1. View → Output (Ctrl+Alt+O)
2. Set "Show output from:" to "Debug"
3. Add debug statements in code:

```csharp
System.Diagnostics.Debug.WriteLine("Processing email: " + mailItem.Subject);
System.Diagnostics.Debug.WriteLine("Attachment count: " + mailItem.Attachments.Count);
```

#### Inspect Variables

When stopped at a breakpoint:
1. Hover over variables to see values
2. Use "Locals" window (Debug → Windows → Locals)
3. Use "Watch" window to monitor specific variables

#### Common Issues During Development

**Issue**: Add-in doesn't appear in Outlook
- **Solution**: Check File → Options → Add-ins → COM Add-ins
- Ensure it's enabled

**Issue**: Changes not reflecting
- **Solution**: 
  1. Stop debugging
  2. Clean solution (Build → Clean Solution)
  3. Rebuild (Build → Rebuild Solution)
  4. Start debugging again

**Issue**: "Cannot start Outlook" error
- **Solution**: 
  1. Close all Outlook processes (Task Manager)
  2. Try again

**Issue**: Breakpoints not hitting
- **Solution**:
  1. Verify you're in Debug mode
  2. Check "Tools → Options → Debugging → Symbols"
  3. Rebuild solution

## Publishing

### Creating a Release Build

1. Select "Release" configuration
2. Build → Build Solution
3. Output: `PrintAllAttachments/bin/Release/`

### Publishing with ClickOnce

1. Right-click project → Properties
2. Click "Publish" tab
3. Configure:
   - **Publishing Folder**: Where to save installer
   - **Installation Folder**: Where users install from
   - **Prerequisites**: .NET Framework 4.7.2

4. Click "Publish Wizard"
5. Follow steps:
   - Choose publish location
   - Choose how users install (from disk, website, etc.)
   - Update settings

6. Click "Finish"

### Output Files

After publishing:
```
publish/
├── setup.exe          # Main installer
├── PrintAllAttachments.vsto  # Add-in manifest
├── Application Files/
│   └── [version]/
│       ├── PrintAllAttachments.dll
│       └── ...
└── publish.htm        # Installation instructions
```

### Signing (Recommended)

For production deployment, sign your add-in:

1. **Create or obtain certificate**:
   - Project → Properties → Signing
   - Create Test Certificate (dev only)
   - Or use real certificate for production

2. **Sign the manifest**:
   - Check "Sign the ClickOnce manifests"
   - Select certificate
   - Enter password

## Continuous Integration

### GitHub Actions Workflow

This project uses GitHub Actions for continuous integration and automated builds.

**Workflow File:** `.github/workflows/build.yml`

**What it does:**
- Automatically builds the project on every push and pull request
- Builds both Debug and Release configurations
- Uploads build artifacts for download
- Creates GitHub releases with pre-built binaries when version tags are pushed

**Build artifacts:**
- Available for 30 days (Debug) or 90 days (Release)
- Can be downloaded from the Actions tab
- Include all DLLs, manifests, and related files

**Creating a release:**
1. Tag a commit with a version number:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
2. GitHub Actions will automatically:
   - Build the Release configuration
   - Create a new GitHub release
   - Attach build artifacts to the release
   - Generate release notes

**Viewing build status:**
- Check the Actions tab on GitHub
- Look for the workflow run corresponding to your commit
- View logs to debug build issues

## Troubleshooting Build Issues

### "Project could not be loaded"
- Ensure Visual Studio has Office development tools installed
- Verify .NET Framework 4.7.2 SDK is installed

### "Reference to Microsoft.Office.Interop.Outlook could not be resolved"
- Install Microsoft Office / Outlook
- Or install Office PIAs (Primary Interop Assemblies)

### Build succeeds but add-in doesn't work
- Check target framework matches Outlook's requirements
- Verify all dependencies are included
- Check Event Viewer for .NET runtime errors

## Performance Profiling

### Using Visual Studio Profiler

1. Debug → Performance Profiler
2. Select profiling tools:
   - CPU Usage
   - Memory Usage
3. Start profiling with Outlook
4. Perform typical operations
5. Stop profiling
6. Analyze results

### Monitoring Memory

Watch for:
- Memory leaks (increasing usage over time)
- Large object allocations
- Unreleased COM objects

## Code Quality

### Code Analysis

Enable code analysis:
1. Project Properties → Code Analysis
2. Check "Enable Code Analysis"
3. Run: Analyze → Run Code Analysis

### StyleCop (Optional)

For code style consistency:
1. Install StyleCop.Analyzers NuGet package
2. Configure rules
3. Fix warnings

## Documentation Updates

When making changes:
1. Update XML comments in code
2. Update README if user-facing changes
3. Update CHANGELOG with version notes
4. Update this guide if build process changes

## Getting Help

If you encounter build issues:
1. Check [FAQ](FAQ.md)
2. Search existing [GitHub Issues](https://github.com/hershyked/Print-all-attachments/issues)
3. Check Visual Studio Output window for detailed errors
4. Open a new issue with:
   - Error messages
   - Build output
   - Visual Studio version
   - Steps to reproduce

---

**Happy Building!** 🛠️
