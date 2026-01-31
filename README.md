# EmmaNeigh - Signature Packet Builder

## Overview

EmmaNeigh is an automated signature packet generation tool designed for M&A transactions and financing deals. It scans transaction documents (PDFs) and automatically creates individualized signature packets for each signatory party, eliminating the tedious manual process of extracting and organizing signature pages.

**Key Features:**
- Automatic detection of signature pages across multiple documents
- Individual-focused packet generation (organized by person, not entity)
- Combines all signature requirements for each person across all documents
- Works entirely offline on locked-down corporate computers
- No installation required - portable Python runtime included
- Client data never leaves local machine

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [How It Works](#how-it-works)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Output Structure](#output-structure)
6. [Architecture](#architecture)
7. [Contributing](#contributing)
8. [Security & Privacy](#security--privacy)
9. [Troubleshooting](#troubleshooting)
10. [Changelog](#changelog)

---

## Quick Start

**For End Users (Lawyers/Staff):**

1. Download the latest release ZIP file
2. Unzip to a location on your computer
3. Drag a folder containing transaction PDFs onto `run_signature_packets.bat`
4. Wait for processing to complete
5. Find signature packets in `signature_packets_output/` folder

**That's it!** No installation, no Python, no command line required.

---

## How It Works

### The Problem

In M&A transactions, creating signature packets is a time-consuming manual task:
- Attorneys must review hundreds of pages across 20+ documents
- Each signatory needs only their specific signature pages
- Manual extraction is error-prone and tedious
- Missing a signature page can delay closings

### The Solution

EmmaNeigh automates this process:

1. **Scans all PDFs** in a selected folder
2. **Detects signature blocks** using intelligent pattern matching
3. **Identifies individual signers** by parsing "Name:" fields
4. **Extracts signature pages** for each person
5. **Generates combined packets** - one PDF per signer with all their pages
6. **Creates Excel tables** for tracking and quality control

### Example

If John Smith needs to sign:
- Credit Agreement (page 87)
- Guaranty (page 12)
- Security Agreement (page 45)

EmmaNeigh creates:
- `signature_packet - JOHN SMITH.pdf` (3 pages)
- `signature_packet - JOHN SMITH.xlsx` (tracking table)

---

## Installation

### For End Users

**No installation required!** 

Download and unzip. The tool includes:
- Portable Python runtime (no system install)
- All required libraries pre-packaged
- Simple BAT file launcher

### For Developers/Contributors

If you want to modify the code:

```bash
# Clone the repository
git clone https://github.com/raamtambe/EmmaNeigh.git
cd EmmaNeigh

# The portable Python is already included
# No additional setup needed
```

---

## Usage

### Basic Usage (Drag & Drop)

1. Gather all transaction PDFs in one folder
2. Drag that folder onto `run_signature_packets.bat`
3. A console window opens showing progress
4. Processing completes with message: "Signature packets saved to: [path]"

### Advanced Usage (Command Line)

```batch
python\python.exe src\build_signature_packets.py "C:\path\to\pdf\folder"
```

---

## Output Structure

After processing, you'll find:

```
your_pdf_folder/
├── signature_packets_output/
│   ├── packets/
│   │   ├── signature_packet - JOHN SMITH.pdf
│   │   ├── signature_packet - JANE DOE.pdf
│   │   └── signature_packet - ABC HOLDINGS LLC.pdf
│   └── tables/
│       ├── MASTER_SIGNATURE_INDEX.xlsx
│       ├── signature_packet - JOHN SMITH.xlsx
│       └── signature_packet - JANE DOE.xlsx
```

### File Descriptions

**PDF Packets:**
- One file per individual signer
- Contains only the pages that person must sign
- Pages maintain original formatting and quality
- Organized in document order

**Excel Tables:**
- `MASTER_SIGNATURE_INDEX.xlsx` - Complete list of all signature obligations
- Individual signer tables - Detailed page references for each person

---

## Architecture

### Design Principles

1. **No Cloud Dependencies** - All processing happens locally
2. **Portable Runtime** - Includes Python environment, no system install
3. **Person-Centric Logic** - Organizes by individual signers, not entities
4. **Fail-Safe Detection** - Prefers precision over recall to avoid false positives

### Signer Detection Strategy

The tool identifies signers using a two-tier approach:

**Tier 1 (Preferred):**
- Locate "BY:" signature markers
- Look downward for explicit "Name:" field
- Extract the person's name from that field

**Tier 2 (Fallback):**
- If no "Name:" field exists
- Look for probable person names near "BY:" marker
- Filter out entity names (LLC, INC, CORP, etc.)
- Apply name normalization to avoid duplicates

**Name Normalization:**
- Convert to uppercase
- Remove punctuation
- Collapse whitespace
- Ensures "John Smith", "JOHN SMITH", "John  Smith" are treated as one person

### File Structure

```
EmmaNeigh/
├── run_signature_packets.bat    # User-facing launcher
├── python/                       # Portable Python runtime
│   ├── python.exe
│   └── (Python libraries)
├── src/
│   └── build_signature_packets.py  # Core logic
└── docs/
    ├── README.md
    ├── USER_GUIDE.md
    └── SECURITY_AND_PRIVACY.md
```

---

## Contributing

### For Code Contributors

This is an internal tool, but improvements are welcome:

1. **Fork the repository**
2. **Create a feature branch:**
   ```bash
   git checkout -b feature/improve-signer-detection
   ```
3. **Make your changes** to `src/build_signature_packets.py`
4. **Test thoroughly** with real transaction documents
5. **Submit a pull request** with clear description

### Areas for Improvement

- Enhanced signer detection for non-standard formats
- Support for initial-only pages
- Capacity detection (Borrower, Guarantor, Lender)
- Cover sheet generation
- Duplicate name handling (same name, different people)

---

## Security & Privacy

### Data Handling Guarantees

✅ **All processing is local** - No network calls, no cloud uploads  
✅ **No telemetry** - Tool does not "phone home"  
✅ **No data persistence** - Files are read, processed, and output locally  
✅ **Auditable source code** - Open for security review  
✅ **No dependencies on external services**  

### Compliance Notes

- **Client Confidentiality:** All client files remain on local machine
- **Work Product:** Generated packets are derivative work product
- **No Internet Required:** Can run on air-gapped machines
- **Portable:** No system modifications, fully reversible

### IT/Security Review

This tool is designed to be reviewed by:
- Knowledge Management
- IT Security
- Practice group technology liaisons

Source code is available for full audit.

---

## Troubleshooting

### Common Issues

#### "No module named 'fitz'" or similar errors
- **Cause:** Portable Python libraries not properly installed
- **Solution:** Ensure you're using the complete ZIP package, not just the .bat file

#### "No signers detected"
- **Cause:** Signature blocks don't match expected format
- **Solution:** Check that PDFs have standard signature blocks with "BY:" and "Name:" fields

#### Multiple packets for same person with slight name variations
- **Cause:** Inconsistent capitalization or punctuation in source PDFs
- **Solution:** Manually consolidate or improve normalization in code

#### Tool runs but output folder is empty
- **Cause:** PDFs may be scanned images without text layer
- **Solution:** OCR the PDFs first using Adobe Acrobat

---

## Changelog

### Version 1.0.0 (Initial Release)

**Features:**
- Automatic signature page detection
- Individual-focused packet generation
- Drag-and-drop BAT launcher
- Excel tracking tables
- Master signature index
- Portable Python runtime

**Design Decisions:**
- Chose BAT over GUI exe to avoid tkinter/Tcl requirements on locked machines
- Used portable Python to eliminate installation requirements
- Prioritized person-centric organization over entity-centric
- Focused on precision (avoiding false positives) over recall

---

## License

MIT License - Internal tool for Kirkland & Ellis LLP

---

## Contact & Support

**Project Maintainer:** Raam Tambe  
**GitHub:** https://github.com/raamtambe/EmmaNeigh

For issues, questions, or feature requests, please open a GitHub issue.

---

## Acknowledgments

This tool was developed through iterative design with AI assistance, focusing on:
- Law firm IT constraints (locked-down Windows environments)
- Deal workflow requirements (closing-ready accuracy)
- User experience (minimal technical knowledge required)
- Security and compliance (local-only processing)

Special thanks to all attorneys and staff who provide feedback to improve signature packet workflows.
