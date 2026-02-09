# EmmaNeigh

**Transaction Management & Document Automation**

EmmaNeigh is a comprehensive desktop application for law firms handling M&A deals, financing transactions, and complex closings. Automate signature packet creation, document execution, time tracking, and more.

---

## Download & Install

### Windows
1. Go to [Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download either:
   - **`EmmaNeigh-Setup.exe`** - Installer with auto-updates (recommended)
   - **`EmmaNeigh-Portable.exe`** - No installation needed
3. Run and start processing documents!

### Mac
1. Go to [Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download the `.dmg` file
3. Open it and drag EmmaNeigh to your Applications folder

**No dependencies. No account required. Just download and run.**

---

## Features

### Closing Documentation

#### 1. Create Signature Packets
Extract signature pages from PDFs and Word documents, organized by signer.

- Drag & drop PDF or DOCX files
- Automatically detects signature blocks (BY:, Name:, signature tables)
- Creates individual packets per signer
- **Format preservation**: PDF in = PDF out, DOCX in = DOCX out
- Generates Excel tracking sheets and master index

#### 2. Create Execution Versions
Merge signed pages from DocuSign back into original documents.

- Upload original documents (PDF or DOCX)
- Upload signed PDF from DocuSign
- Automatically matches and merges signature pages
- Handles DocuSign PDF restrictions
- **Format preservation**: Outputs match input format

#### 3. Signature Blocks from Checklist
Generate signature blocks from transaction checklists.

- Parse Excel/CSV transaction checklists
- Parse incumbency certificates
- Auto-generate properly formatted signature blocks
- Insert into PDF or DOCX documents

### Document Processing

#### 4. Collate Documents
Merge track changes from multiple reviewers into a single document.

- Upload base document and reviewer versions
- Combines all tracked changes
- Preserves formatting and comments

### Project Management

#### 5. Email Search
Parse and search Outlook email exports.

- Upload Outlook CSV export
- Search by keyword, sender, date
- Privacy controls (blur sensitive info)
- Prepare for checklist integration

### Time Management

#### 6. Activity Summary
Generate time tracking summaries from emails and calendar.

- Upload email CSV and calendar exports (ICS/CSV)
- Auto-categorize by matter/client
- Generate daily/weekly activity summaries
- Timeline view of your day

---

## What's New in v5.0

### Redesigned Navigation
- New sidebar with collapsible categories
- Organized into: Closing, Document Processing, Project Management, Time Management
- Wider window (800px) for better workspace

### Format Preservation
- DOCX documents now output as DOCX (not converted to PDF)
- PDF documents continue to output as PDF
- Works across Signature Packets and Execution Versions

### Usage History Tracking
- Track all document processing operations
- View recent activity and statistics
- Export history to CSV
- All data stored locally

### Auto-Update Support (Windows Installer)
- Automatic update checks
- Download progress indicator
- One-click install and restart

---

## Security & Privacy

EmmaNeigh is designed for handling confidential legal documents:

| Feature | Description |
|---------|-------------|
| **100% Local** | All processing happens on your computer |
| **No Cloud** | Your files never leave your machine |
| **No Tracking** | Zero analytics or telemetry |
| **No Account** | No login or registration required |
| **Portable** | Runs from anywhere, leaves no trace |

Your documents stay on your computer. Always.

---

## How It Works

### Signature Detection
EmmaNeigh scans each page looking for signature block patterns:
- "BY:" fields with name detection
- Signature tables with name columns
- "Name:" and "Title:" fields

### Document Matching
For execution versions, EmmaNeigh uses:
- Document name matching from footers ("Signature Page to [Document]")
- Signer name matching
- Fuzzy text matching for accuracy

### Format Preservation
- Input format is detected automatically
- PDF inputs produce PDF outputs
- DOCX inputs produce DOCX outputs
- Mixed inputs are handled separately

---

## FAQ

**Q: Do I need to install anything else?**
A: No. Just download and run. Everything is bundled.

**Q: Does it work offline?**
A: Yes. EmmaNeigh only connects to the internet to check for updates.

**Q: What file formats are supported?**
A: PDF and Microsoft Word (.docx). Scanned documents without OCR may not work.

**Q: Is my data sent anywhere?**
A: No. All processing is 100% local on your machine.

**Q: What's the difference between Setup and Portable?**
A: Setup installs and supports auto-updates. Portable requires no installation.

**Q: Can I use this on my work computer?**
A: Yes. The portable version requires no admin rights.

---

## Version History

### v5.0.1 (February 2026)
- Format preservation: DOCX in = DOCX out, PDF in = PDF out
- Direct release uploads (improved build reliability)

### v5.0.0 (February 2026)
- Complete UI redesign with sidebar navigation
- Usage history tracking with SQLite
- Email CSV parsing and search
- Time tracking / activity summaries
- Auto-update support for Windows installer
- Collate documents feature
- Signature blocks from checklist

### v4.x (February 2026)
- Added DOCX support for signature detection
- Document collation feature
- Redline comparison (removed in v5.0)

### v3.x (February 2026)
- Execution version improvements
- Better DocuSign handling

### v2.x (February 2026)
- Desktop application with GUI
- Drag & drop file upload
- ZIP file downloads

### v1.0.0 (January 2026)
- Initial release (command-line version)

---

## Technical Details

### Built With
- **Electron** - Cross-platform desktop framework
- **Python** - Document processing (PyMuPDF, python-docx, pandas)
- **sql.js** - Local SQLite for history tracking

### System Requirements
- Windows 10/11 or macOS 10.15+
- 4GB RAM recommended
- 200MB disk space

---

## Support

For issues or feature requests, please [open an issue](https://github.com/raamtambe/EmmaNeigh/issues) on GitHub.

---

## License

MIT License - See [LICENSE](LICENSE) for details.

**Made for lawyers, by lawyers (with help from AI).**
