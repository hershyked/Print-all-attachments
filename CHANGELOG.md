# Changelog

All notable changes to the Print All Attachments Outlook Add-in will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- Fixed `quick-install.bat` failing with "'build.bat' is not recognized" error when run from different directories

### Planned Features
- Custom printer selection
- Print settings configuration dialog
- Progress bar for large batches
- File type filtering options
- Print preview option
- Statistics tracking

## [1.0.0] - 2025-01-XX

### Added
- Initial release of Print All Attachments add-in
- Ribbon button in Outlook Mail tab
- Support for printing attachments from multiple selected emails
- Automatic extraction of attachments to temporary folder
- Print functionality using Windows default print handlers
- User feedback dialog with operation summary
- Error handling and reporting
- Automatic cleanup of temporary files
- Support for common file formats:
  - PDF documents
  - Microsoft Office documents (Word, Excel, PowerPoint)
  - Image files (JPEG, PNG, BMP, TIFF, GIF)
  - Text files
  - HTML files
- Comprehensive documentation:
  - README with overview
  - Installation guide
  - Usage guide
  - Contributing guidelines
- Visual Studio project structure for VSTO add-in
- .NET Framework 4.7.2 targeting
- Code security validation (CodeQL clean)

### Security
- Temporary file handling with automatic cleanup
- No external data transmission
- Access limited to user-selected emails only
- Uses Windows temp directory with unique folder names

## Version History

### Version Numbering

This project uses semantic versioning:
- **MAJOR** version for incompatible API changes
- **MINOR** version for backwards-compatible new features
- **PATCH** version for backwards-compatible bug fixes

### Release Notes

#### 1.0.0 - Initial Release

**Highlights:**
- First stable release
- Core functionality: print attachments from multiple emails
- Full documentation suite
- Security validated

**Requirements:**
- Windows 7 SP1 or later
- Microsoft Outlook 2013 or later
- .NET Framework 4.7.2 or later

**Known Limitations:**
- Requires default printer to be set
- Some file types may require associated application to be installed
- Applications may briefly appear during printing
- Print settings use application defaults (cannot be customized per print)

**Testing:**
- Manual testing on Windows 10/11
- Tested with Outlook 2016, 2019, and Microsoft 365
- Tested with PDF, Office documents, and image files
- Security analysis passed (CodeQL)

---

## Upgrade Guide

### From Future Version to 1.0.0
Not applicable - this is the initial release.

### Future Upgrades
Upgrade instructions will be provided here as new versions are released.

## Support

For issues, questions, or feature requests, please:
1. Check the [README](README.md)
2. Review this changelog
3. Search existing [GitHub Issues](https://github.com/hershyked/Print-all-attachments/issues)
4. Open a new issue if needed

## Contributors

Thank you to everyone who has contributed to this project!

- Initial development and documentation

---

**Note:** Dates use ISO 8601 format (YYYY-MM-DD)

[Unreleased]: https://github.com/hershyked/Print-all-attachments/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/hershyked/Print-all-attachments/releases/tag/v1.0.0
