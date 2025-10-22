# Contributing to Print All Attachments

Thank you for considering contributing to this project! This document provides guidelines for contributing.

## Ways to Contribute

- Report bugs
- Suggest new features
- Improve documentation
- Submit code improvements
- Write tests

## Getting Started

1. Fork the repository
2. Clone your fork:
   ```bash
   git clone https://github.com/YOUR-USERNAME/Print-all-attachments.git
   ```
3. Create a branch for your changes:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Development Setup

### Prerequisites

- Windows 10 or later
- Visual Studio 2017 or later with:
  - Office/SharePoint development workload
  - .NET desktop development workload
- Microsoft Outlook (2013 or later)
- .NET Framework 4.7.2 or later

### Building the Project

1. Open `PrintAllAttachments.sln` in Visual Studio
2. Restore NuGet packages (should happen automatically)
3. Build the solution (Ctrl+Shift+B)

### Running and Debugging

1. Set the configuration to Debug
2. Press F5 to start debugging
3. Outlook will launch with the add-in loaded
4. Test your changes
5. Press Shift+F5 to stop debugging

## Code Style Guidelines

### C# Conventions

- Use 4 spaces for indentation (not tabs)
- Use meaningful variable names
- Add XML comments for public methods
- Follow Microsoft C# coding conventions

### Example:

```csharp
/// <summary>
/// Prints a file using the default application
/// </summary>
/// <param name="filePath">Full path to the file</param>
/// <returns>True if successful</returns>
private bool PrintFile(string filePath)
{
    // Implementation
}
```

## Making Changes

### Before You Start

1. Check existing issues and pull requests
2. Open an issue to discuss major changes
3. Keep changes focused and minimal

### Code Changes

1. Write clean, readable code
2. Add comments for complex logic
3. Update documentation if needed
4. Test your changes thoroughly

### Testing

Since this is a VSTO add-in:

1. **Manual Testing Required**
   - Test with different file types
   - Test with single and multiple emails
   - Test error scenarios
   - Test on different Outlook versions if possible

2. **Test Checklist**
   - [ ] Add-in loads in Outlook
   - [ ] Button appears in ribbon
   - [ ] Prints PDF attachments
   - [ ] Prints Office document attachments
   - [ ] Prints image attachments
   - [ ] Handles emails with no attachments
   - [ ] Handles multiple selected emails
   - [ ] Shows appropriate error messages
   - [ ] Cleans up temporary files

### Documentation

Update documentation for:
- New features
- Changed behavior
- New configuration options
- Known limitations

## Submitting Changes

### Commit Messages

Write clear commit messages:

```
Add feature to filter printable file types

- Added file extension validation
- Updated error messages
- Added documentation for supported types
```

**Format:**
- First line: Brief summary (50 chars or less)
- Blank line
- Detailed description if needed
- Reference issue numbers: `Fixes #123`

### Pull Request Process

1. **Update your branch:**
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

2. **Push to your fork:**
   ```bash
   git push origin feature/your-feature-name
   ```

3. **Create Pull Request:**
   - Go to the repository on GitHub
   - Click "New Pull Request"
   - Select your branch
   - Fill in the PR template

4. **PR Description should include:**
   - What changed and why
   - How to test the changes
   - Screenshots if UI changed
   - Related issue numbers

5. **Wait for review:**
   - Address reviewer comments
   - Make requested changes
   - Push updates to your branch

## Reporting Bugs

### Before Reporting

1. Check if the issue already exists
2. Test with the latest version
3. Verify it's not a configuration issue

### Bug Report Template

```markdown
**Description**
A clear description of the bug

**To Reproduce**
1. Go to '...'
2. Click on '...'
3. Select '...'
4. See error

**Expected Behavior**
What should happen

**Actual Behavior**
What actually happened

**Screenshots**
If applicable

**Environment:**
- Windows Version: [e.g., Windows 11]
- Outlook Version: [e.g., Outlook 2021]
- Add-in Version: [e.g., 1.0.0]

**Additional Context**
Any other relevant information
```

## Suggesting Features

### Feature Request Template

```markdown
**Is your feature request related to a problem?**
A clear description of the problem

**Proposed Solution**
How you think it should work

**Alternatives Considered**
Other approaches you've thought about

**Additional Context**
Mockups, examples, etc.
```

## Code Review Process

### What Reviewers Look For

- Code quality and readability
- Performance implications
- Security considerations
- Error handling
- Documentation updates
- Breaking changes

### Response Time

- We aim to review PRs within 1 week
- Urgent fixes may be reviewed faster
- Complex changes may take longer

## Project Structure

```
Print-all-attachments/
├── PrintAllAttachments/          # Main project
│   ├── PrintAllAttachments.csproj
│   ├── ThisAddIn.cs              # Add-in entry point
│   ├── PrintAttachmentsRibbon.cs # UI and logic
│   └── Properties/
│       └── AssemblyInfo.cs
├── PrintAllAttachments.sln       # Solution file
├── README.md                     # Main documentation
├── INSTALLATION.md               # Install guide
├── USAGE.md                      # User guide
└── CONTRIBUTING.md               # This file
```

## Areas for Contribution

### High Priority

- **Testing**: Add unit tests and integration tests
- **Error Handling**: Improve error messages and recovery
- **Performance**: Optimize for large batches
- **Compatibility**: Test with different Outlook versions

### Medium Priority

- **Features**:
  - Print queue management
  - Custom printer selection
  - Print settings dialog
  - Progress bar for large batches
  
- **Documentation**:
  - Video tutorials
  - FAQ section
  - Troubleshooting guide expansion

### Low Priority

- **UI Enhancements**:
  - Custom icons
  - Keyboard shortcuts
  - Context menu integration

## Questions?

- Open an issue with the "question" label
- Check existing documentation
- Review closed issues for similar questions

## License

By contributing, you agree that your contributions will be licensed under the same license as the project.

## Code of Conduct

### Our Standards

- Be respectful and inclusive
- Welcome newcomers
- Accept constructive criticism
- Focus on what's best for the community

### Unacceptable Behavior

- Harassment or discrimination
- Trolling or insulting comments
- Publishing private information
- Unprofessional conduct

### Enforcement

Project maintainers have the right to remove, edit, or reject:
- Comments, commits, code, issues, and PRs
- That don't align with this Code of Conduct

## Recognition

Contributors will be recognized in:
- README contributors section (to be added)
- Release notes for significant contributions

## Thank You!

Your contributions make this project better for everyone. We appreciate your time and effort! 🎉
