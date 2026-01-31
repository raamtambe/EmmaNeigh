# Security and Privacy Policy

## EmmaNeigh - Signature Packet Builder

**Last Updated:** January 2026

---

## Overview

This document outlines the security and privacy practices of the EmmaNeigh signature packet builder. It is intended for:
- Knowledge Management
- IT Security teams
- Risk Management
- Practice group technology liaisons
- Concerned users

---

## Core Security Principles

### 1. Local-Only Processing

**Guarantee:** All document processing occurs entirely on the user's local machine.

**Implementation:**
- No network connections initiated
- No HTTP/HTTPS requests
- No socket connections
- No DNS lookups
- No API calls

**Code Verification:**
```python
# Examine src/build_signature_packets.py
# No import of: requests, urllib, socket, http
# Only local file I/O operations
```

### 2. No Data Transmission

**Guarantee:** Client files never leave the local machine.

**What is NOT done:**
- ❌ No cloud uploads
- ❌ No telemetry
- ❌ No crash reporting to external servers
- ❌ No usage analytics
- ❌ No "phone home" functionality

**Evidence:**
- No API keys in source code
- No external service configurations
- Complete source code available for audit

### 3. No Data Persistence Beyond Output

**Guarantee:** The tool does not create hidden logs, caches, or databases.

**Processing Flow:**
1. Read PDFs from specified folder
2. Process in memory
3. Write output to specified folder
4. Exit cleanly

**No Hidden Storage:**
- No AppData folder usage
- No Registry modifications
- No temporary files in system directories
- Output only in user-visible location

---

## Privacy Guarantees

### Client Confidentiality

**Attorney-Client Privilege Protected:**
- Documents remain under attorney's physical control
- No third-party access or exposure
- Processing equivalent to reading files locally
- No different from using Adobe Acrobat

**Work Product Doctrine:**
- Generated signature packets are derivative work product
- Created in the course of legal representation
- Protected by work product immunity

### Personal Data Handling

**What Data is Processed:**
- Signer names (extracted from PDFs)
- Document names
- Page numbers

**What Data is Stored:**
- Output PDFs (signature page images)
- Excel tables (metadata only)
- All stored in user-specified output folder

**What Data is NOT Collected:**
- User identity
- Firm name
- Client name
- Document contents beyond signature pages
- Usage patterns
- Error logs (beyond local console)

---

## Technical Security Measures

### 1. Dependency Isolation

**Portable Python Runtime:**
- Self-contained in `python/` folder
- Does not modify system Python
- No PATH modifications
- No Registry changes
- Fully removable by deleting folder

**Library Provenance:**
- PyMuPDF: BSD-licensed, widely audited
- pandas: BSD-licensed, industry standard
- openpyxl: MIT-licensed, mature project

### 2. File System Access

**Read Access:**
- Limited to user-specified input folder
- Standard Python file I/O (`open()`, `fitz.open()`)
- No privileged file access

**Write Access:**
- Limited to output subfolder within input directory
- Creates: `signature_packets_output/`
- Does not modify original PDFs

**No System Access:**
- Does not read system files
- Does not access other user directories
- Does not enumerate drives
- Does not search for files

### 3. Execution Model

**No Elevated Privileges:**
- Runs with user's normal permissions
- No admin rights required
- No UAC prompts
- No system modifications

**Process Isolation:**
- Single-threaded execution
- No subprocess spawning (except internal Python)
- No external program invocation
- Clean process termination

---

## Compliance Considerations

### Regulatory Alignment

**GDPR (if applicable):**
- ✅ Data minimization: Only processes signature pages
- ✅ Purpose limitation: Only for signature packet creation
- ✅ Storage limitation: No data retention beyond session
- ✅ Integrity and confidentiality: Local-only processing

**CCPA (if applicable):**
- ✅ No "sale" of personal information
- ✅ No sharing with third parties
- ✅ No profiling or tracking

**Attorney Ethics Rules:**
- ✅ Maintains confidentiality (ABA Model Rule 1.6)
- ✅ Competent representation (due diligence on tools)
- ✅ Supervision of subordinates (transparent operation)

### Industry Standards

**NIST Cybersecurity Framework:**
- **Identify:** Clear data flow documentation
- **Protect:** No external exposure
- **Detect:** Transparent logging
- **Respond:** Source code available for incident analysis
- **Recover:** No persistent state to recover from

---

## Audit and Verification

### Source Code Transparency

**Repository:** https://github.com/raamtambe/EmmaNeigh

**Key Files to Review:**
1. `src/build_signature_packets.py` - Core logic
2. `run_signature_packets.bat` - Launcher
3. `docs/ARCHITECTURE.md` - Technical design

**What to Look For:**
- Network-related imports (there are none)
- External API calls (there are none)
- Data transmission code (there is none)

### Security Review Checklist

For IT/Security teams:

**Static Analysis:**
- [ ] Review source code for network libraries
- [ ] Verify no encrypted/obfuscated code
- [ ] Check dependencies for known vulnerabilities
- [ ] Confirm no auto-update mechanism

**Dynamic Analysis:**
- [ ] Monitor network traffic during execution (should be zero)
- [ ] Check file system access (should be limited to input folder)
- [ ] Verify no registry changes
- [ ] Confirm no background processes

**Behavioral Testing:**
- [ ] Run on sample PDFs and verify output
- [ ] Test error conditions (malformed PDFs)
- [ ] Verify graceful failure (no crashes)

---

## Threat Model

### What This Tool Protects Against

✅ **Accidental Data Leakage:**
- No cloud uploads
- No third-party services
- No clipboard access

✅ **Unauthorized Access:**
- No network exposure
- No listening ports
- No external communication

### What This Tool Does NOT Protect Against

⚠️ **Local Machine Compromise:**
- If user's computer is infected with malware, that malware could access files
- This is a general OS security issue, not tool-specific

⚠️ **User Error:**
- Tool cannot prevent users from emailing wrong packets
- Tool cannot verify signer identities

⚠️ **Physical Security:**
- Output files are as secure as the machine they're on
- Users must follow firm policies for file storage

---

## Incident Response

### If a Security Concern Arises

**Reporting:**
1. Open GitHub issue: https://github.com/raamtambe/EmmaNeigh/issues
2. Email maintainer with details
3. Describe: what happened, what was expected, any error messages

**Response Timeline:**
- Acknowledgment: Within 24 hours
- Initial assessment: Within 48 hours
- Fix/mitigation: Depends on severity

**Disclosure:**
- Security issues will be addressed promptly
- Updates will be pushed as new releases
- Users will be notified of critical issues

---

## Safe Usage Guidelines

### Best Practices for Users

**Before Running:**
1. ✅ Verify source (download from official GitHub releases)
2. ✅ Check ZIP file integrity
3. ✅ Run on firm-approved machines only

**During Use:**
1. ✅ Only process files you have authorization to access
2. ✅ Review output before distributing to signers
3. ✅ Delete output when no longer needed

**After Use:**
1. ✅ Store signature packets securely
2. ✅ Follow firm document retention policies
3. ✅ Report any unusual behavior

### What NOT to Do

❌ **Do not:**
- Modify the Python code without understanding it
- Run on untrusted PDF files from unknown sources
- Share output folders via insecure channels
- Leave sensitive output on shared drives

---

## Data Retention

**Tool Behavior:**
- Processes data during execution only
- Writes output to user-specified location
- No automatic deletion
- No automatic backup

**User Responsibility:**
- Output management (saving, deleting)
- Compliance with firm retention policies
- Secure disposal when appropriate

---

## Updates and Versioning

### How Updates are Distributed

**No Automatic Updates:**
- Tool does not check for updates
- Tool does not download updates
- User must manually download new versions

**Version Checking:**
- Users can check GitHub releases for new versions
- Release notes will describe changes

**Security Patches:**
- Critical security issues will be addressed immediately
- Patch releases will be clearly marked

---

## Third-Party Dependencies

### Library Security

**PyMuPDF (fitz):**
- Purpose: PDF parsing and page extraction
- License: BSD
- Security: Mature, actively maintained
- Alternatives reviewed: PyPDF2, pdfplumber

**pandas:**
- Purpose: Data manipulation
- License: BSD
- Security: Industry standard, well-audited
- Used by: Fortune 500, academia, government

**openpyxl:**
- Purpose: Excel file creation
- License: MIT
- Security: Widely used, stable

### Dependency Verification

Users can verify library versions:
```bash
python\python.exe -m pip list
```

Expected output includes:
- PyMuPDF (fitz)
- pandas
- openpyxl

---

## Conclusion

EmmaNeigh is designed with privacy and security as foundational principles:

✅ **No data leaves the local machine**  
✅ **No external services or APIs**  
✅ **Transparent, auditable source code**  
✅ **Minimal system footprint**  
✅ **Compliant with attorney confidentiality obligations**

For law firms with strict security policies, this tool provides:
- Productivity automation
- Complete data control
- Audit trail capability
- No external dependencies

---

## Contact

**Security Questions:** Open GitHub issue or contact maintainer

**Code Audit Requests:** Source code available at https://github.com/raamtambe/EmmaNeigh

**Maintainer:** Raam Tambe

---

**Document Version:** 1.0  
**Last Review:** January 2026  
**Next Review:** Annually or upon significant changes
