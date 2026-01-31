# Changelog

All notable changes to the EmmaNeigh Signature Packet Builder will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Planned
- Initial-only page detection
- Capacity field extraction (Borrower, Guarantor, Lender)
- Duplicate name disambiguation
- Cover sheet generation per packet
- Configuration file for custom signature keywords

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

### Design Decisions
- BAT launcher over EXE (avoids tkinter/PyInstaller complexity)
- Person-centric vs. entity-centric organization
- Name normalization for consistency
- Precision over recall in signer detection

### Known Limitations
- Does not detect initial-only pages
- Does not extract capacity fields (Borrower/Guarantor)
- Limited support for non-standard signature blocks
- Requires OCR for scanned PDFs
- No duplicate name disambiguation (same-named individuals)

---

## Development History

### Context
EmmaNeigh was developed iteratively through conversation with AI assistance, focusing on:
- Real-world law firm constraints (locked IT environments)
- Deal workflow requirements (closing-ready accuracy)
- User experience (minimal technical knowledge required)
- Security compliance (attorney-client privilege, no data leakage)

### Key Milestones
1. **Initial concept:** Automated signature packet creation
2. **Architecture choice:** Portable Python + BAT launcher
3. **Detection strategy:** Two-tier signer extraction
4. **Person-centric pivot:** Organize by individual, not entity
5. **Documentation:** Comprehensive guides for all audiences

### Evolution Notes
- **v0.1 (concept):** "UNKNOWN SIGNER" placeholders
- **v0.2 (refinement):** Entity name filtering
- **v0.3 (pivot):** Person-centric organization
- **v1.0 (release):** Production-ready with full documentation

---

## Future Versions

### Version 1.1.0 (Planned)
**Focus:** Enhanced detection and usability

#### Planned Features
- [ ] Initial-only page detection
- [ ] Capacity field extraction
- [ ] Improved error messages
- [ ] Self-check/validation mode
- [ ] Progress indicator during processing

#### Under Consideration
- [ ] Cover sheet generation
- [ ] Configurable signature keywords
- [ ] Dry-run mode (preview without creating output)

### Version 1.2.0 (Conceptual)
**Focus:** Advanced features

#### Potential Features
- [ ] Duplicate name handling (disambiguation prompts)
- [ ] DocuSign envelope ordering
- [ ] Batch processing mode
- [ ] Custom output templates

### Version 2.0.0 (Long-term)
**Focus:** Intelligence and integration

#### Visionary Features
- [ ] Machine learning for non-standard signature blocks
- [ ] Multi-language support
- [ ] Integration with deal management systems
- [ ] Automated quality control checks

---

## Maintenance Notes

### Bug Fix Policy
- **Critical bugs** (crashes, data loss): Immediate patch release
- **High-priority bugs** (incorrect output): Within 1 week
- **Medium bugs** (usability issues): Next minor release
- **Low-priority bugs** (cosmetic): Batched in future releases

### Deprecation Policy
- Breaking changes require major version bump
- Deprecated features announced in advance
- Backward compatibility maintained where possible

---

## Release Notes Template

For future releases, use this format:

```markdown
## [X.Y.Z] - YYYY-MM-DD

### Added
- New features

### Changed
- Changes to existing functionality

### Deprecated
- Features planned for removal

### Removed
- Removed features

### Fixed
- Bug fixes

### Security
- Security improvements
```

---

## Version History Summary

| Version | Date       | Highlights                                |
|---------|------------|-------------------------------------------|
| 1.0.0   | 2026-01-31 | Initial release with core functionality   |

---

**Note:** This changelog is maintained manually. All changes should be documented here before release.

**Maintainer:** Raam Tambe  
**Last Updated:** January 31, 2026
