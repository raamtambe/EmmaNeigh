# Getting Started with EmmaNeigh

**Quick navigation for all users**

---

## I want to...

### Use the tool (End User)

👉 **Read:** [USER_GUIDE.md](USER_GUIDE.md)

**Quick start:**
1. Download latest release ZIP
2. Extract anywhere
3. Drag folder of PDFs onto `run_signature_packets.bat`
4. Find output in `signature_packets_output/`

**Need help?** See [USER_GUIDE.md](USER_GUIDE.md) for detailed step-by-step instructions.

---

### Review security/privacy (IT/Security)

👉 **Read:** [SECURITY_AND_PRIVACY.md](SECURITY_AND_PRIVACY.md)

**Key facts:**
- ✅ All processing is local
- ✅ No network calls
- ✅ No data transmission
- ✅ Auditable source code

**Deployment approved at:** Kirkland & Ellis LLP

---

### Install or deploy (IT/Admin)

👉 **Read:** [INSTALLATION.md](INSTALLATION.md)

**Deployment options:**
- Self-service download
- Centralized network location
- Individual workstation deployment

**Requirements:** Windows 10+, no admin rights needed

---

### Understand how it works (Technical)

👉 **Read:** [ARCHITECTURE.md](ARCHITECTURE.md)

**Core technologies:**
- Python 3.11.2 (portable)
- PyMuPDF for PDF parsing
- Two-tier signer detection
- Name normalization

---

### Contribute or modify (Developer)

👉 **Read:** [CONTRIBUTING.md](CONTRIBUTING.md)

**Setup:**
```bash
git clone https://github.com/raamtambe/EmmaNeigh.git
pip install pymupdf pandas openpyxl
```

**Areas for contribution:**
- Signer detection improvements
- New features (see roadmap)
- Documentation
- Testing

---

### See what's new (Everyone)

👉 **Read:** [CHANGELOG.md](CHANGELOG.md)

**Latest version:** 1.0.0 (January 31, 2026)

---

## Documentation Structure

```
docs/
├── README.md                    # Project overview
├── USER_GUIDE.md                # For attorneys/staff
├── INSTALLATION.md              # Setup and deployment
├── ARCHITECTURE.md              # Technical design
├── SECURITY_AND_PRIVACY.md      # Security guarantees
├── CONTRIBUTING.md              # For contributors
└── CHANGELOG.md                 # Version history
```

---

## Common Questions

### What does this tool do?

Creates individualized signature packets for M&A deals. Each person gets one PDF with only their signature pages from all documents.

### Do I need to install Python?

No. Everything needed is included in the ZIP file.

### Is it secure for client files?

Yes. All processing happens locally on your computer. No data is transmitted.

### Can I use this on my work laptop?

Yes. Designed for locked-down corporate Windows environments.

### How do I get help?

- End users: See [USER_GUIDE.md](USER_GUIDE.md)
- Technical issues: Open a GitHub issue
- Security questions: See [SECURITY_AND_PRIVACY.md](SECURITY_AND_PRIVACY.md)

---

## Project Information

**Name:** EmmaNeigh  
**Purpose:** Signature packet automation for M&A transactions  
**Version:** 1.0.0  
**License:** MIT  
**Maintainer:** Raam Tambe  
**Repository:** https://github.com/raamtambe/EmmaNeigh

---

## Quick Links

- **Download:** [GitHub Releases](https://github.com/raamtambe/EmmaNeigh/releases)
- **Report Bug:** [Open Issue](https://github.com/raamtambe/EmmaNeigh/issues)
- **Request Feature:** [Open Issue](https://github.com/raamtambe/EmmaNeigh/issues)
- **Email Maintainer:** See GitHub profile

---

## Next Steps

**For end users:**
1. Download the latest release
2. Read [USER_GUIDE.md](USER_GUIDE.md)
3. Try on test PDFs
4. Use on real deals

**For IT/security:**
1. Review [SECURITY_AND_PRIVACY.md](SECURITY_AND_PRIVACY.md)
2. Read [INSTALLATION.md](INSTALLATION.md)
3. Test in isolated environment
4. Approve for deployment

**For developers:**
1. Read [ARCHITECTURE.md](ARCHITECTURE.md)
2. Review [CONTRIBUTING.md](CONTRIBUTING.md)
3. Set up development environment
4. Pick an issue or feature

---

**Welcome to EmmaNeigh! Let's automate those signature packets.**
