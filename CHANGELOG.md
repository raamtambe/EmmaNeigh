# Changelog

All notable changes to EmmaNeigh will be documented in this file.

---

## [5.0.1] - 2026-02-09

### Added
- **Format preservation** - DOCX inputs now produce DOCX outputs, PDF inputs produce PDF outputs
- Direct release uploads (bypasses GitHub artifact storage quota)

### Changed
- Updated build workflow for more reliable releases

---

## [5.0.0] - 2026-02-09

### Added

#### New Features
- **Email Search** - Parse and search Outlook CSV exports
- **Time Tracking** - Generate activity summaries from emails and calendar
- **Document Collation** - Merge track changes from multiple reviewers
- **Usage History** - Track all operations with local SQLite database

#### UI Improvements
- **Sidebar Navigation** - New accordion-style collapsible categories
- **Wider Window** - 800px width for better workspace
- **History Modal** - View and export usage statistics

#### Auto-Update (Windows Installer)
- Automatic update checking
- Download progress indicator
- One-click install and restart

### Changed
- Complete navigation redesign with 4 categories:
  - Closing (Sig Packets, Execution Versions, Sig Blocks)
  - Document Processing (Collate)
  - Project Management (Email Search)
  - Time Management (Activity Summary)

### Removed
- Redline comparison feature (deprecated)

---

## [4.1.x] - 2026-02

### Added
- MS Word (.docx) support for all processors
- Document collation feature

---

## [3.x] - 2026-02

### Added
- Signature blocks from checklist workflow
- Incumbency certificate parsing
- Better DocuSign page matching

---

## [2.2.0] - 2026-02-03

### Changed
- Clean, professional UI
- Inter font, SVG icons, slate color palette

---

## [2.0.0] - 2026-02-01

### Added
- Desktop GUI with Electron
- Drag & drop file upload
- Execution Version creator
- ZIP downloads

---

## [1.0.0] - 2026-01-31

### Added
- Initial release
- Signature page detection
- Person-centric packets
- Excel tracking

---

## Version Summary

| Version | Date | Highlights |
|---------|------|------------|
| 5.0.1 | 2026-02-09 | Format preservation (DOCX in = DOCX out) |
| 5.0.0 | 2026-02-09 | Major UI redesign, email/time tracking, history |
| 4.1.x | 2026-02 | DOCX support, collation |
| 3.x | 2026-02 | Signature blocks from checklist |
| 2.x | 2026-02 | Desktop GUI |
| 1.0.0 | 2026-01-31 | Initial release |

---

**Maintainer:** Raam Tambe
