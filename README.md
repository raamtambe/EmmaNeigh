# EmmaNeigh

**Signature Packet Automation for M&A Transactions**

EmmaNeigh automates the tedious process of extracting and organizing signature pages from transaction documents. Built for law firms handling M&A deals, financing transactions, and other complex closings.

---

## Download & Install

### Windows
1. Go to [Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download `EmmaNeigh-Portable.exe`
3. Double-click to run - that's it!

### Mac
1. Go to [Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download the `.dmg` file
3. Open it and drag EmmaNeigh to your Applications folder
4. Double-click to run

**No installation required. No dependencies. Just download and run.**

---

## Features

### 1. Create Signature Packets
Extract signature pages from multiple PDFs and organize them by signer.

**How to use:**
1. Open EmmaNeigh
2. Click "Create Signature Packets"
3. Drag & drop your PDF files (or click to browse)
4. Click "Create Signature Packets"
5. Watch the horse run while processing
6. Download your ZIP file with all packets

**What you get:**
- Individual PDF packets for each signer (e.g., `signature_packet - JOHN SMITH.pdf`)
- Excel tracking sheets
- Master signature index

### 2. Create Execution Version
Merge signed pages from DocuSign back into the original document.

**The problem:** DocuSign returns locked PDFs that can't be edited.

**The solution:** EmmaNeigh automatically unlocks DocuSign PDFs and merges the signed pages back into your original document.

**How to use:**
1. Open EmmaNeigh
2. Click "Create Execution Version"
3. Upload the original PDF (without signature pages)
4. Upload the signed PDF from DocuSign
5. Enter the page number where signatures should be inserted
6. Click "Create Execution Version"
7. Download your final execution version

---

## Security & Privacy

EmmaNeigh is designed for handling confidential legal documents:

| Feature | Description |
|---------|-------------|
| **100% Local** | All processing happens on your computer |
| **No Internet** | App never connects to the internet |
| **No Cloud** | Your files never leave your machine |
| **No Tracking** | Zero analytics or telemetry |
| **No Account** | No login or registration required |
| **Portable** | Runs from anywhere, leaves no trace |

Your documents stay on your computer. Always.

---

## Screenshots

### Main Menu
Choose between creating signature packets or execution versions.

### Processing View
Watch the animated horse while your documents are processed. Real-time status updates show exactly what's happening.

### Results
Download all your signature packets as a single ZIP file.

---

## How It Works

### Signature Detection
EmmaNeigh scans each page looking for signature block patterns:
- "BY:" fields
- "Name:" fields
- Signature lines

When found, it extracts the signer's name and groups all their signature pages together.

### DocuSign Unlocking
DocuSign adds permission restrictions to signed PDFs. EmmaNeigh removes these restrictions so pages can be merged into execution versions.

---

## FAQ

**Q: Do I need to install anything else?**
A: No. Just download and run. Everything is bundled.

**Q: Does it work offline?**
A: Yes. EmmaNeigh never connects to the internet.

**Q: What PDF formats are supported?**
A: Standard PDFs with text layers. Scanned documents without OCR may not work.

**Q: Is my data sent anywhere?**
A: No. All processing is 100% local on your machine.

**Q: Can I use this on my work computer?**
A: Yes. It's a portable app that requires no installation or admin rights.

---

## Version History

### v2.1.0 (February 2026)
- Simplified installation - just download and run
- Automated builds for Windows and Mac
- Updated documentation

### v2.0.0 (February 2026)
- Desktop application with modern GUI
- Drag & drop file upload
- Running horse animation
- Execution Version creator
- ZIP file downloads

### v1.0.0 (January 2026)
- Initial release (command-line version)

See [CHANGELOG.md](CHANGELOG.md) for full details.

---

## Support

For issues or feature requests, please [open an issue](https://github.com/raamtambe/EmmaNeigh/issues) on GitHub.

---

## License

MIT License - See [LICENSE](LICENSE) for details.

**Made for lawyers, by lawyers (with help from AI).**
