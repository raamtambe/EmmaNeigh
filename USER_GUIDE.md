# EmmaNeigh User Guide

## For Attorneys and Staff

This guide is for people who need to **use** EmmaNeigh. No technical knowledge required.

---

## What This Tool Does

EmmaNeigh automates document processing for M&A transactions and complex closings:

1. **Signature Packets** - Extract signature pages from documents, organized by signer
2. **Execution Versions** - Merge signed DocuSign pages back into original documents
3. **Signature Blocks** - Generate signature blocks from transaction checklists
4. **Document Collation** - Merge track changes from multiple reviewers
5. **Email Search** - Parse and search Outlook email exports
6. **Time Tracking** - Generate activity summaries from emails and calendar

---

## Getting Started

### Download & Install

**Windows:**
1. Go to [GitHub Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download **EmmaNeigh-Setup.exe** (recommended) or **EmmaNeigh-Portable.exe**
3. Run the installer or portable executable

**Mac:**
1. Download the **.dmg** file from Releases
2. Open and drag EmmaNeigh to Applications

### First Launch

When you open EmmaNeigh, you'll see a sidebar with four categories:
- **Closing** - Signature packets, execution versions, signature blocks
- **Document Processing** - Document collation
- **Project Management** - Email search
- **Time Management** - Activity summaries

---

## Feature Guides

### 1. Create Signature Packets

**What it does:** Extracts signature pages from PDFs/DOCX and creates individual packets per signer.

**How to use:**
1. Click "Create Sig Packets" in the sidebar
2. Drag & drop your document files (PDF or DOCX)
3. Click "Create Signature Packets"
4. Download the ZIP file when complete

**Output:**
- One PDF/DOCX packet per signer (format matches input)
- Excel tracking sheets
- Master signature index

### 2. Create Execution Versions

**What it does:** Merges signed pages from DocuSign back into original documents.

**How to use:**
1. Click "Execution Versions" in the sidebar
2. Upload original documents (PDF or DOCX)
3. Upload the signed PDF from DocuSign
4. Click "Create Execution Versions"
5. Download the executed documents

**Note:** Output format matches input - DOCX in = DOCX out, PDF in = PDF out.

### 3. Signature Blocks from Checklist

**What it does:** Generates signature blocks from a transaction checklist and incumbency certificates.

**How to use:**
1. Click "Sig Blocks from Checklist" in the sidebar
2. Upload your transaction checklist (Excel)
3. Upload incumbency certificates
4. Map entities to their roles
5. Upload documents to add signature pages to
6. Click "Generate Signature Blocks"

### 4. Collate Documents

**What it does:** Merges track changes from multiple reviewer versions into a single document.

**How to use:**
1. Click "Collate Documents" in the sidebar
2. Upload the base document
3. Upload reviewer versions
4. Click "Collate"
5. Download the merged document

### 5. Email Search

**What it does:** Parses Outlook CSV exports for searching and analysis.

**How to use:**
1. Click "Email Search" in the sidebar
2. Export emails from Outlook as CSV
3. Upload the CSV file
4. Search by keyword, sender, or date

### 6. Activity Summary

**What it does:** Generates time tracking summaries from email and calendar data.

**How to use:**
1. Click "Activity Summary" in the sidebar
2. Upload email CSV (from Outlook export)
3. Upload calendar file (ICS or CSV)
4. Select date range
5. View activity breakdown by matter/client

---

## Tips for Best Results

### Document Preparation
- Use searchable PDFs (with text layer, not scanned images)
- Ensure signature blocks have "BY:" and "Name:" fields
- Standard signature table formats work best

### File Formats
- **PDF** - Standard PDFs with text layers
- **DOCX** - Microsoft Word documents
- **Format preservation** - Input format = Output format

### Quality Control
- Review the master signature index after processing
- Spot-check a few packets before distributing
- Verify all signers were detected

---

## Common Questions

**Q: Do I need to install anything else?**
A: No. Just download and run. Everything is bundled.

**Q: Does it work offline?**
A: Yes. Only the update check requires internet.

**Q: What if a signer isn't detected?**
A: Check that the PDF has standard "BY:" and "Name:" fields. Scanned documents without OCR won't work.

**Q: Can two people have the same name?**
A: They'll be combined into one packet. You may need to manually separate if they're different people.

**Q: Is my data sent anywhere?**
A: No. All processing is 100% local on your machine.

---

## Viewing History

Click the **History** button in the top-right to see:
- Recent operations
- Statistics by feature
- Export history to CSV

---

## Updating

**Windows Installer (EmmaNeigh-Setup.exe):**
- Automatic update notifications
- Click "Download" when prompted
- Click "Install & Restart" when ready

**Portable/Mac:**
- Check [Releases](https://github.com/raamtambe/EmmaNeigh/releases) for new versions
- Download and replace the old version

---

## Getting Help

**For issues or bugs:**
- Open a GitHub issue: https://github.com/raamtambe/EmmaNeigh/issues
- Include: what you tried, what happened, any error messages

---

## Security

- All files stay on your computer
- No data is transmitted
- Safe for confidential client documents
- Complies with attorney-client privilege requirements

---

**Last Updated:** February 2026
