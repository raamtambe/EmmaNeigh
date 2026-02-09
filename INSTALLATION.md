# Installation Guide

## EmmaNeigh - Transaction Management & Document Automation

---

## For End Users

### Windows

**Option 1: Installer (Recommended)**
1. Go to [GitHub Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download **EmmaNeigh-Setup.exe**
3. Run the installer
4. Follow the prompts (can choose install location)
5. Launch from Start Menu or Desktop shortcut

**Benefits:**
- Supports automatic updates
- Standard Windows installation
- Easy uninstall via Control Panel

**Option 2: Portable**
1. Download **EmmaNeigh-Portable.exe**
2. Save anywhere (Desktop, USB drive, etc.)
3. Double-click to run

**Benefits:**
- No installation required
- No admin rights needed
- Runs from any location

### Mac

1. Download the **.dmg** file from Releases
2. Open the DMG
3. Drag EmmaNeigh to Applications
4. Double-click to run

**First launch:** Right-click and select "Open" if macOS shows a security warning.

---

## Requirements

**Minimal:**
- Windows 10/11 or macOS 10.15+
- 4GB RAM
- 200MB disk space

**No additional requirements:**
- No Python installation needed
- No admin rights needed (for portable)
- No internet required (except for updates)

---

## For IT/Security Review

### Deployment Options

**Option 1: Self-Service**
- Post download link on internal portal
- Users download and run as needed
- No IT involvement required

**Option 2: Centralized**
- Deploy installer via SCCM/Intune
- Or place portable EXE on network share
- Single version to maintain

### Security Verification

Before deployment, verify:

1. **Scan for malware** - Standard antivirus scan
2. **Review behavior** - Monitor network traffic (should be zero except update checks)
3. **Test in isolation** - Run on sample documents first

### Network Requirements

- **Outbound:** GitHub API for update checks only (optional)
- **Inbound:** None
- **Firewall:** No special rules needed
- **Proxy:** Works behind corporate proxy

### File System Access

- Reads: User-selected files only
- Writes: Output to user-specified location
- No access to: System files, other user directories, registry

---

## Auto-Update (Windows Installer Only)

The installed version checks for updates on launch:

1. Connects to GitHub Releases API
2. Compares version numbers
3. Shows banner if update available
4. User clicks "Download" to get update
5. User clicks "Install & Restart" to apply

**To disable:** Updates are not forced; users can ignore the banner.

---

## Uninstallation

### Windows Installer
- Control Panel → Programs → Uninstall EmmaNeigh
- Or Settings → Apps → EmmaNeigh → Uninstall

### Windows Portable
- Delete the .exe file
- Delete output folders if desired
- No registry entries to clean

### Mac
- Drag EmmaNeigh from Applications to Trash
- Empty Trash

**Note:** User data (history database) is stored in:
- Windows: `%APPDATA%/emmaneigh/`
- Mac: `~/Library/Application Support/emmaneigh/`

Delete these folders for complete removal.

---

## Troubleshooting

### Windows Protected Your PC
- Click "More info" → "Run anyway"
- This happens because the app is not code-signed

### Mac Security Warning
- Right-click the app → Open → Open
- Or: System Preferences → Security → Open Anyway

### App Won't Start
- Ensure you have Windows 10+ or macOS 10.15+
- Try running as Administrator (Windows)
- Check antivirus hasn't quarantined the file

### Slow Performance
- Large files (100+ MB PDFs) may take longer
- Close other applications to free memory
- Check disk space

---

## Version Information

**Current Version:** 5.0.1
**Platforms:** Windows 10+, macOS 10.15+

**Check for updates:**
- GitHub Releases: https://github.com/raamtambe/EmmaNeigh/releases

---

**Last Updated:** February 2026
