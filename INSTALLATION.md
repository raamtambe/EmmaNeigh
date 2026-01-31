# Installation Guide

## EmmaNeigh - Signature Packet Builder

This guide covers installation for different user types and scenarios.

---

## For End Users (Lawyers/Staff)

### Requirements

**Minimal:**
- Windows 10 or later
- Ability to download and unzip files
- PDF documents to process

**No installation required:**
- ❌ No Python installation
- ❌ No admin rights
- ❌ No command-line knowledge

### Step 1: Download

1. Go to [GitHub Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Find the latest version (e.g., v1.0.0)
3. Download the ZIP file: `Signature_Packet_Tool_v1.0.zip`

### Step 2: Extract

1. Right-click the downloaded ZIP
2. Select "Extract All..."
3. Choose destination (Desktop or Documents recommended)
4. Click "Extract"

### Step 3: Verify Contents

After extraction, you should have:

```
Signature_Packet_Tool/
├── run_signature_packets.bat    ← This is what you'll use
├── python/                       ← Don't touch this
│   ├── python.exe
│   └── (various files)
├── src/
│   └── build_signature_packets.py
└── docs/
    └── USER_GUIDE.md             ← Read this!
```

### Step 4: First Run

**Test with sample PDFs:**
1. Create a test folder with 2-3 transaction PDFs
2. Drag the folder onto `run_signature_packets.bat`
3. Watch the console window
4. Check the output folder when complete

**If successful:**
- You'll see: "Signature packets saved to: [path]"
- Output appears in `signature_packets_output/`

**If errors:**
- Read the error message
- Check [Troubleshooting](#troubleshooting)
- Contact support if needed

---

## For IT/Security Review

### Deployment Options

**Option 1: Self-Service Download**
- Post ZIP on internal SharePoint/KM site
- Users download and unzip as needed
- No IT involvement required

**Option 2: Centralized Installation**
- Extract to shared network location
- Users run from network path
- Single copy to maintain

**Option 3: Individual Workstation Deployment**
- IT extracts to C:\Tools\ or similar
- Create desktop shortcut to BAT file
- Still portable, just centrally managed

### Security Verification

Before deployment, IT should:

1. **Scan ZIP for malware**
   - Standard antivirus scan
   - Should be clean

2. **Review source code**
   - Check `src/build_signature_packets.py`
   - Verify no network calls
   - Confirm no system modifications

3. **Test in isolated environment**
   - Run on sample PDFs
   - Monitor network traffic (should be zero)
   - Verify file system access (limited to input folder)

4. **Document approval**
   - Add to approved software list
   - Note version number
   - Set review schedule

### Whitelisting

**If application whitelisting is enforced:**

Add to whitelist:
- `python.exe` (in tool directory)
- `run_signature_packets.bat`

**PowerShell execution policy:**
- BAT file does not require PowerShell
- Should work regardless of execution policy

### Network Considerations

**No network access required:**
- Tool works offline
- No ports opened
- No external connections

**Firewall:**
- No firewall rules needed
- Can run on air-gapped machines

---

## For Developers/Contributors

### Development Environment Setup

**Prerequisites:**
- Git installed
- Python 3.11+ installed
- Text editor or IDE
- GitHub account

### Step 1: Clone Repository

```bash
git clone https://github.com/raamtambe/EmmaNeigh.git
cd EmmaNeigh
```

### Step 2: Set Up Development Environment

**Option A: Virtual Environment (Recommended)**
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate

pip install pymupdf pandas openpyxl
```

**Option B: Global Installation**
```bash
pip install pymupdf pandas openpyxl
```

### Step 3: Test Installation

```bash
python src/build_signature_packets.py path/to/test/pdfs
```

### Step 4: Set Up Git

```bash
git config user.name "Your Name"
git config user.email "your.email@example.com"
```

### Step 5: Create Feature Branch

```bash
git checkout -b feature/your-feature-name
```

**Now you're ready to develop!**

---

## For Package Maintainers

### Building a Distribution Package

**What's included in the distribution:**
1. Portable Python runtime
2. Pre-installed libraries
3. Source code
4. BAT launcher
5. Documentation

### Build Process

**1. Prepare Portable Python:**
```bash
# Download Python embeddable package
# From: https://www.python.org/downloads/windows/
# File: Windows embeddable package (64-bit)

# Extract to python/ directory
# Enable site-packages (edit python311._pth)
# Install pip
# Install dependencies
```

**2. Install Dependencies:**
```bash
cd python
python.exe -m pip install pymupdf pandas openpyxl
```

**3. Create Distribution Structure:**
```
Signature_Packet_Tool/
├── run_signature_packets.bat
├── python/
│   ├── python.exe
│   ├── python311.dll
│   └── (all files from embeddable package + site-packages)
├── src/
│   └── build_signature_packets.py
└── docs/
    ├── README.md
    ├── USER_GUIDE.md
    └── SECURITY_AND_PRIVACY.md
```

**4. Create ZIP:**
```bash
zip -r Signature_Packet_Tool_v1.0.0.zip Signature_Packet_Tool/
```

**5. Upload to GitHub Releases:**
- Create release tag: v1.0.0
- Upload ZIP
- Add release notes from CHANGELOG.md

---

## Platform-Specific Notes

### Windows

**Works on:**
- Windows 10 (all editions)
- Windows 11
- Windows Server 2019+

**Known issues:**
- OneDrive folders with long paths may cause issues (move to C:\)
- Some antivirus may quarantine (false positive, whitelist the folder)

### macOS

**Not officially supported** for end-user execution.

**For developers:**
- Can develop and test code
- Cannot build Windows executable
- Use Windows VM or separate machine for building distribution

### Linux

**Not officially supported** for end-user execution.

**For developers:**
- Can develop and test Python code
- Cannot build Windows executable
- Use Wine or Windows VM for testing BAT launcher

---

## Troubleshooting

### Common Issues During Installation

**Issue: "Windows protected your PC" warning**
- **Cause:** Unsigned executable
- **Solution:** Click "More info" → "Run anyway"
- **Better:** Download from official GitHub releases only

**Issue: ZIP won't extract**
- **Cause:** File corrupted during download
- **Solution:** Re-download the ZIP file

**Issue: Antivirus blocks python.exe**
- **Cause:** False positive (portable Python is uncommon)
- **Solution:** Whitelist the entire tool folder

**Issue: BAT file opens and closes immediately**
- **Cause:** Error before "Press any key" is reached
- **Solution:** 
  - Right-click BAT → Edit
  - Add `pause` before `exit /b`
  - See error message

### Verification Tests

**Test 1: Python runs**
```batch
python\python.exe --version
# Should print: Python 3.11.2
```

**Test 2: Libraries installed**
```batch
python\python.exe -m pip list
# Should show: PyMuPDF, pandas, openpyxl
```

**Test 3: Script runs**
```batch
python\python.exe src\build_signature_packets.py
# Should print usage error (no folder provided)
```

---

## Uninstallation

### For End Users

**Complete removal:**
1. Delete the extracted folder
2. Delete any shortcuts
3. Delete output folders if desired

**That's it!** No registry entries, no system modifications.

### For IT Deployments

**Centralized installations:**
1. Remove from shared location
2. Notify users
3. Clean up any shortcuts

**No system cleanup needed:**
- Tool leaves no traces in:
  - Registry
  - AppData
  - System directories
  - PATH environment variable

---

## Upgrade Process

### For End Users

**To upgrade to a new version:**

1. Download new version ZIP
2. Extract to **new folder** (don't overwrite old version)
3. Test with sample PDFs
4. Once verified, can delete old version
5. Update any shortcuts

**Why separate folders:**
- Allows rollback if issues
- Can compare outputs
- No risk to working version

### For IT Deployments

**Version management:**
1. Test new version in isolated environment
2. Deploy to pilot group
3. Gather feedback
4. Roll out to all users
5. Archive old version (don't delete immediately)

---

## Support

**Installation issues:**
- Check [USER_GUIDE.md](USER_GUIDE.md) first
- See [Troubleshooting](#troubleshooting) above
- Open GitHub issue with details

**Security review questions:**
- See [SECURITY_AND_PRIVACY.md](SECURITY_AND_PRIVACY.md)
- Contact IT liaison for firm approval process

**Development questions:**
- See [CONTRIBUTING.md](CONTRIBUTING.md)
- Open GitHub issue
- Email maintainer

---

## Version Information

**Current Version:** 1.0.0  
**Release Date:** January 31, 2026  
**Python Version:** 3.11.2 (portable)  
**Platform:** Windows 10+

**Check for updates:**
- GitHub Releases: https://github.com/raamtambe/EmmaNeigh/releases
- Watch repository for notifications
