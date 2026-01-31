# EmmaNeigh User Guide

## For Attorneys and Staff

This guide is for people who need to **use** the Signature Packet Builder, not modify it. No technical knowledge required.

---

## What This Tool Does

**In Plain English:**

You give it a folder full of transaction documents (PDFs).  
It automatically creates one signature packet for each person who needs to sign.  
Each packet contains only the pages that person must sign.

**Example:**

If you have:
- Credit Agreement (100 pages, 5 people sign)
- Guaranty (20 pages, 3 people sign)  
- Security Agreement (50 pages, 4 people sign)

The tool creates:
- One PDF for each unique person
- Each PDF has only their signature pages from all documents
- One Excel table per person showing which documents they're signing

---

## Step-by-Step Instructions

### Step 1: Download the Tool

1. Go to: [GitHub Releases](https://github.com/raamtambe/EmmaNeigh/releases)
2. Download the latest ZIP file (e.g., `Signature_Packet_Tool_v1.0.zip`)
3. Save it somewhere you can find it (Desktop or Documents folder)

### Step 2: Unzip the File

1. Right-click the ZIP file
2. Choose "Extract All..."
3. Pick a location and click "Extract"

You now have a folder with:
- `run_signature_packets.bat` (this is what you'll click)
- `python` folder (don't touch this)
- `docs` folder (optional reading)

### Step 3: Prepare Your PDFs

1. Create a new folder anywhere on your computer
2. Put ALL transaction PDFs into that folder
3. Make sure PDFs have text (not just scanned images)

**Tip:** Name the folder something clear like "Smith Deal Closing Docs"

### Step 4: Run the Tool

**Option A: Drag & Drop (Easiest)**
1. Click and drag your PDF folder
2. Drop it onto `run_signature_packets.bat`
3. A black window opens (don't close it!)

**Option B: Double-Click**
1. Double-click `run_signature_packets.bat`
2. When prompted, paste the path to your PDF folder
3. Press Enter

### Step 5: Wait for Processing

You'll see messages like:
```
Processing PDFs in:
C:\Users\You\Desktop\Smith Deal Closing Docs

Scanning Credit_Agreement.pdf
Scanning Guaranty.pdf
Built packet for JOHN SMITH
Built packet for JANE DOE

DONE.
Signature packets saved to:
C:\Users\You\Desktop\Smith Deal Closing Docs\signature_packets_output
```

**Don't close the window until it says "DONE."**

### Step 6: Find Your Output

Go to your original PDF folder. You'll now see a new folder:

```
signature_packets_output/
├── packets/         ← PDFs ready to send
└── tables/          ← Excel files for tracking
```

---

## Understanding the Output

### PDF Packets (in `packets/` folder)

**File naming:**
- `signature_packet - JOHN SMITH.pdf`
- `signature_packet - JANE DOE.pdf`

**What's inside:**
- Only the pages that person needs to sign
- Pages are in the order they appeared in the original documents
- Original formatting preserved (no weird reformatting)

### Excel Tables (in `tables/` folder)

**MASTER_SIGNATURE_INDEX.xlsx**
- Shows every signature obligation across all documents
- Use this for quality control

**Individual Tables (e.g., `signature_packet - JOHN SMITH.xlsx`)**
- Lists exactly which documents and pages John Smith must sign
- Columns: Signer Name | Document | Page

---

## What to Do Next

### Quality Control (Recommended)

1. Open `MASTER_SIGNATURE_INDEX.xlsx`
2. Quickly scan to ensure all signers were detected
3. Spot-check 2-3 PDF packets to confirm pages are correct

### Send to Signers

1. Email each person their PDF packet
2. Include their Excel table if they want page references
3. Track signatures using the master index

---

## Common Questions

### Q: Do I need to install Python?
**No.** Everything you need is in the ZIP file.

### Q: Does this upload my files anywhere?
**No.** All processing happens on your computer. No internet required.

### Q: Can I use this on my work laptop?
**Yes.** It works on locked-down corporate computers with no admin rights.

### Q: What if two people have the same name?
The tool will create one packet. You'll need to manually separate if they're different people.

### Q: Can I run this multiple times?
**Yes.** Each run creates a fresh output folder. Previous outputs are not deleted.

### Q: What if it doesn't detect a signer?
Check that the PDF has a standard signature block with "BY:" and "Name:" fields. Scanned PDFs without text won't work.

### Q: Can I use this for non-M&A deals?
**Yes.** Any transaction with signature pages works (real estate, litigation settlements, etc.).

---

## When to Contact IT/Support

You should contact support if:
- The tool crashes immediately
- Error messages you don't understand
- Signature pages are consistently missed
- You need to process hundreds of documents at once

**Not bugs:**
- Scanned PDFs without OCR (get them OCRed first)
- Unusual signature block formats (tool is tuned for standard blocks)

---

## Tips for Best Results

### ✅ Do This:
- Use searchable PDFs (with text layer)
- Ensure signature blocks have "Name:" fields
- Keep all PDFs in one folder
- Review the master index after each run

### ❌ Avoid This:
- Mixing scanned and native PDFs (OCR the scanned ones first)
- Running on partially downloaded files
- Interrupting the process mid-run
- Renaming files during processing

---

## Security Reminder

- **All files stay on your computer**
- **No data is transmitted**
- **Safe to use with confidential client documents**
- **Complies with attorney-client privilege requirements**

---

## Getting Help

**For basic questions:** See [Troubleshooting](#troubleshooting) in main README

**For bugs or issues:**
- Open a GitHub issue: https://github.com/raamtambe/EmmaNeigh/issues
- Include: what you tried, what happened, any error messages

**For feature requests:**
- Contact the project maintainer
- Describe your use case and why it matters

---

## Version Information

Check which version you're using:
- Look at the ZIP filename (e.g., `v1.0.zip`)
- Check the `docs/` folder for a VERSION file

Always use the latest version for bug fixes and improvements.

---

## Legal Notice

This is an internal productivity tool. It:
- Does not provide legal advice
- Does not replace attorney judgment
- Should be used for administrative tasks only
- Is not a substitute for document review

Always review signature packets before sending to clients or counterparties.
