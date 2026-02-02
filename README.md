# EmmaNeigh - Signature Packet Automation

## Overview

EmmaNeigh is an automated signature packet tool designed for M&A transactions and financing deals. It scans transaction documents (PDFs) and automatically creates individualized signature packets for each signatory party, eliminating the tedious manual process of extracting and organizing signature pages.

**Now with a beautiful desktop application!** Version 2.0 introduces a modern GUI with drag-and-drop file upload, animated progress indicators, and new features like Execution Version creation.

---

## What's New in v2.0.0

### Desktop Application
- **Modern GUI** - Beautiful, intuitive interface built with React
- **Drag & Drop** - Simply drag PDF files into the app
- **Running Horse Animation** - Fun visual feedback during processing
- **Cross-Platform** - Works on Windows, Mac, and Linux
- **Portable** - No installation required, runs from any folder

### New Features
- **Execution Version Creator** - Merge signed DocuSign pages back into original documents
- **ZIP Download** - Get all signature packets in a single downloadable ZIP file
- **Real-time Progress** - See exactly what's happening as files are processed
- **DocuSign PDF Unlocking** - Automatically removes DocuSign restrictions

### Security (Unchanged)
- **100% Local Processing** - All data stays on your machine
- **No Network Calls** - Works completely offline
- **No Telemetry** - Zero tracking or data collection

---

## Quick Start

### Desktop App (Recommended - v2.0)

1. Navigate to `desktop-app/` folder
2. Run setup (first time only):
   ```bash
   npm install
   cd frontend && npm install && cd ..
   pip3 install -r python/requirements.txt
   ```
3. Launch the app:
   ```bash
   npm run dev
   ```
4. Use the beautiful GUI to upload files and create signature packets!

### Command Line (Legacy - v1.0)

For the original drag-and-drop batch file experience:
1. Download the release ZIP
2. Drag a folder of PDFs onto `run_signature_packets.bat`
3. Find output in `signature_packets_output/` folder

---

## Features

### Create Signature Packets
Extract signature pages from transaction documents and organize by signer.

**How it works:**
1. Upload multiple PDF documents
2. Tool scans for signature pages (looks for "BY:", "Name:" fields)
3. Groups pages by individual signer
4. Generates one PDF packet per signer
5. Creates Excel tracking sheets
6. Download everything as a ZIP file

**Example Output:**
```
signature_packets_output/
├── packets/
│   ├── signature_packet - JOHN SMITH.pdf
│   ├── signature_packet - JANE DOE.pdf
│   └── signature_packet - ABC HOLDINGS LLC.pdf
└── tables/
    ├── MASTER_SIGNATURE_INDEX.xlsx
    └── (individual signer tables)
```

### Create Execution Version (NEW in v2.0)
Merge signed pages back into original documents after DocuSign signing.

**The Problem:** DocuSign returns locked/protected PDFs that can't be edited.

**The Solution:** EmmaNeigh automatically:
1. Unlocks the DocuSign PDF restrictions
2. Extracts the signed pages
3. Merges them into your original document
4. Creates the final execution version

---

## Installation

### Prerequisites
- Node.js 18+ (for desktop app)
- Python 3.10+
- npm

### Setup

```bash
# Clone the repository
git clone https://github.com/raamtambe/EmmaNeigh.git
cd EmmaNeigh/desktop-app

# Install dependencies
npm install
cd frontend && npm install && cd ..
pip3 install -r python/requirements.txt

# Run the app
npm run dev
```

### Building for Distribution

```bash
# Build frontend
npm run build:react

# Build Python processor
pip3 install pyinstaller
cd python && pyinstaller --onefile --name processor main.py && cd ..

# Build portable app
npm run dist:win   # Windows
npm run dist:mac   # macOS
npm run dist:linux # Linux
```

---

## Project Structure

```
EmmaNeigh/
├── desktop-app/              # NEW: Desktop application (v2.0)
│   ├── electron/             # Electron main process
│   ├── frontend/             # React UI
│   └── python/               # PDF processors
├── run_signature_packets.bat # Legacy: v1.0 launcher
├── README.md
├── CHANGELOG.md
└── (documentation files)
```

---

## Security & Privacy

EmmaNeigh is designed for law firm environments handling sensitive client documents:

| Guarantee | Description |
|-----------|-------------|
| **Local Processing** | All document processing happens on your machine |
| **No Network** | Zero internet connections, works offline |
| **No Cloud** | Files never leave your local disk |
| **No Telemetry** | No analytics, tracking, or data collection |
| **No Admin Rights** | Runs as a portable application |
| **Auditable** | Full source code available for review |

---

## Troubleshooting

### Desktop App Issues

**"npm: command not found"**
- Install Node.js from https://nodejs.org/

**"python3: command not found"**
- Install Python from https://python.org/

**App won't start**
- Make sure port 3000 is available
- Try `npm run dev` again

### Processing Issues

**"No signers detected"**
- Ensure PDFs have standard signature blocks with "BY:" and "Name:" fields
- Check that PDFs are not scanned images (need OCR first)

**"Cannot unlock DocuSign PDF"**
- The PDF may be password-protected (not just permission-restricted)
- Contact the sender for the password

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for full version history.

### v2.0.0 (February 2026)
- Desktop application with modern GUI
- Running horse animation
- Execution Version creator
- ZIP file download
- Cross-platform support

### v1.0.0 (January 2026)
- Initial release
- Signature packet generation
- Drag-and-drop batch launcher
- Excel tracking tables

---

## Contributing

Contributions welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

**Areas for Improvement:**
- Enhanced signer detection for non-standard formats
- DocuSign API integration
- Initial-only page detection
- Cover sheet generation

---

## License

MIT License - See [LICENSE](LICENSE) for details.

---

## Contact

**Project Maintainer:** Raam Tambe
**GitHub:** https://github.com/raamtambe/EmmaNeigh

For issues, questions, or feature requests, please open a GitHub issue.
