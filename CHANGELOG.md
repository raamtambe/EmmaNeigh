# Changelog

All notable changes to the EmmaNeigh Signature Packet Builder will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.0] - 2026-02-01

### Added

#### Desktop Application
- **Modern GUI** - Beautiful, intuitive desktop interface built with Electron + React
- **Main Menu** - Easy navigation between features with visual feature cards
- **Drag & Drop Upload** - Simply drag PDF files into the application
- **Running Horse Animation** - Fun animated horse with real-time status messages during processing
- **Progress Bar** - Visual progress indicator showing processing status
- **ZIP Download** - Download all signature packets as a single ZIP file

#### New Feature: Execution Version Creator
- Merge signed pages back into original documents
- Automatic DocuSign PDF unlocking (removes permission restrictions)
- Specify exact page insertion point
- Creates final execution version ready for filing

#### Technical Improvements
- **Cross-Platform Support** - Works on Windows, Mac, and Linux
- **Portable Application** - No installation required, runs from any folder
- **Real-time Progress** - JSON-based progress streaming from Python to UI
- **Modular Python Code** - Refactored processors for better maintainability

### Changed
- Project structure reorganized with new `desktop-app/` directory
- Python code refactored into modular processors
- Documentation updated for v2.0

### Technical Stack
- **Frontend**: React 18 + Vite + Tailwind CSS
- **Desktop Shell**: Electron 28
- **Backend**: Python 3.10+ with PyMuPDF
- **IPC**: Electron IPC with JSON progress streaming

### Security
- Maintained all v1.0 security guarantees
- 100% local processing (no network calls)
- No telemetry or data collection
- Works on air-gapped machines

---

## [1.0.0] - 2026-01-31

### Added
- Initial release of Signature Packet Builder
- Automatic signature page detection across multiple PDFs
- Person-centric packet generation (organized by individual signer)
- Name normalization to handle variations
- Excel table output for quality control
- Master signature index showing all obligations
- Drag-and-drop BAT launcher for Windows
- Portable Python runtime (no installation required)
- Comprehensive documentation (README, User Guide, Architecture, Security)

### Core Features
- Signature block detection using keyword-based approach
- Two-tier signer extraction:
  - Tier 1: Explicit "Name:" field parsing
  - Tier 2: Proximity-based fallback for non-standard formats
- Entity name filtering (excludes LLC, INC, CORP, etc.)
- Multi-document aggregation per signer
- Multi-signer support on single pages
- Local-only processing (no network calls)

### Documentation
- README.md - Project overview and quick start
- USER_GUIDE.md - Step-by-step instructions for non-technical users
- ARCHITECTURE.md - Technical design and implementation details
- SECURITY_AND_PRIVACY.md - Security guarantees and compliance
- CONTRIBUTING.md - Guidelines for contributors

### Technical Specifications
- Python 3.11.2 (portable/embeddable)
- Dependencies: PyMuPDF, pandas, openpyxl
- Platform: Windows (locked-down corporate environments)
- Processing: Fully offline, local-only

### Known Limitations
- Does not detect initial-only pages
- Does not extract capacity fields (Borrower/Guarantor)
- Limited support for non-standard signature blocks
- Requires OCR for scanned PDFs
- No duplicate name disambiguation (same-named individuals)

---

## Future Versions

### Version 2.1.0 (Planned)
- [ ] DocuSign API integration for sending packets
- [ ] Initial-only page detection
- [ ] Capacity field extraction
- [ ] Batch processing mode

### Version 2.2.0 (Conceptual)
- [ ] Cover sheet generation
- [ ] Duplicate name disambiguation
- [ ] Custom output templates
- [ ] Processing history/log

---

## Version History Summary

| Version | Date       | Highlights                                      |
|---------|------------|-------------------------------------------------|
| 2.0.0   | 2026-02-01 | Desktop app, GUI, Execution Version creator     |
| 1.0.0   | 2026-01-31 | Initial release with core functionality         |

---

**Maintainer:** Raam Tambe
**Last Updated:** February 1, 2026
