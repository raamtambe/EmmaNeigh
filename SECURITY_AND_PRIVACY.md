# Security and Privacy Policy

## EmmaNeigh - Transaction Management & Document Automation

**Last Updated:** February 2026

---

## Overview

This document outlines the security and privacy practices of EmmaNeigh. It is intended for:
- IT Security teams
- Knowledge Management
- Risk Management
- Concerned users

---

## Core Security Principles

### 1. Local-Only Processing

**Guarantee:** All document processing occurs entirely on the user's local machine.

**Implementation:**
- Documents are processed in-memory
- No files are uploaded to any server
- No cloud storage integration
- Results saved locally to user-specified location

### 2. No Data Transmission

**Guarantee:** Client files never leave the local machine.

**What is NOT done:**
- No cloud uploads
- No telemetry on document contents
- No crash reporting with document data
- No external API calls for document processing

**Only exception:** Optional update check to GitHub API (version number only, no document data).

### 3. No Hidden Data Storage

**Guarantee:** The tool only stores usage history locally.

**What IS stored (locally):**
- Usage history (feature used, timestamp, counts)
- User preferences

**What is NOT stored:**
- Document contents
- File names
- Signer names
- Client information

**Location:**
- Windows: `%APPDATA%/emmaneigh/`
- Mac: `~/Library/Application Support/emmaneigh/`

---

## Privacy Guarantees

### Client Confidentiality

**Attorney-Client Privilege Protected:**
- Documents remain under attorney's control
- No third-party access
- Processing equivalent to using any local application
- No different from using Microsoft Word

### Personal Data Handling

**What Data is Processed (in memory only):**
- Signer names (extracted from documents)
- Document content (for signature detection)
- Email metadata (if using email features)

**What Data is NOT Collected:**
- User identity (no login required)
- Firm name
- Client names
- Document contents (not transmitted)
- Usage patterns (not transmitted)

---

## Technical Security

### Dependency Isolation

**Electron Application:**
- Self-contained application
- Python processors bundled inside
- No system Python required
- No PATH modifications
- No registry changes (portable version)

**Key Libraries:**
- PyMuPDF - PDF parsing (BSD license)
- python-docx - Word documents (MIT license)
- pandas - Data processing (BSD license)
- sql.js - Local SQLite (MIT license)

### File System Access

**Read Access:**
- Only files explicitly selected by user
- No file system scanning
- No access to other directories

**Write Access:**
- Output to user-specified location
- History database in AppData
- No system file modifications

### Network Access

**Update Checks (optional):**
- Connects to GitHub Releases API
- Sends: Current version number only
- Receives: Latest version info
- Can be disabled by blocking GitHub

**No Other Network Access:**
- Document processing is 100% offline
- No external services for any feature
- Works on air-gapped machines

---

## Compliance

### Regulatory Alignment

**GDPR:**
- Data minimization: Only processes what's needed
- Purpose limitation: Clear, defined features
- Storage limitation: No cloud retention
- User control: All data stays local

**Attorney Ethics (ABA Model Rules):**
- Maintains confidentiality (Rule 1.6)
- Competent use of technology (Rule 1.1)
- Supervision (transparent operation)

---

## Audit and Verification

### Source Code Transparency

**Repository:** https://github.com/raamtambe/EmmaNeigh

**Key Areas to Review:**
- `desktop-app/main.js` - Electron main process
- `desktop-app/python/` - Python processors
- `.github/workflows/` - Build process

### Security Review Checklist

For IT/Security teams:

**Static Analysis:**
- [ ] Review for network library usage
- [ ] Verify no obfuscated code
- [ ] Check dependencies for vulnerabilities

**Dynamic Analysis:**
- [ ] Monitor network traffic (should be minimal/none)
- [ ] Check file system access patterns
- [ ] Verify no background processes

---

## Safe Usage Guidelines

### Best Practices

**Before Use:**
- Download only from official GitHub releases
- Verify file hash if provided
- Run on firm-approved machines

**During Use:**
- Only process authorized documents
- Review output before distribution
- Use on secure network

**After Use:**
- Store output securely
- Follow document retention policies
- Delete temporary files when done

---

## Incident Response

### Reporting Security Concerns

1. Open GitHub issue (for non-sensitive issues)
2. Email maintainer directly (for sensitive issues)
3. Include: Description, steps to reproduce, impact

### Response Timeline

- Acknowledgment: Within 24 hours
- Assessment: Within 48 hours
- Fix: Based on severity

---

## Updates

### Security Patches

- Critical issues addressed immediately
- Updates distributed via GitHub Releases
- Users notified through update mechanism

### Version Verification

Check your version:
- About section in application
- GitHub Releases for latest

---

## Conclusion

EmmaNeigh is designed with privacy and security as foundational principles:

- **No data leaves the local machine**
- **No external services for document processing**
- **Transparent, auditable source code**
- **Minimal system footprint**
- **Compliant with attorney confidentiality obligations**

---

## Contact

**Security Questions:** Open GitHub issue or contact maintainer
**Source Code:** https://github.com/raamtambe/EmmaNeigh

**Maintainer:** Raam Tambe

---

**Document Version:** 2.0
**Last Review:** February 2026
