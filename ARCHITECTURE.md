# EmmaNeigh - Technical Architecture

## Overview

This document explains the technical design, implementation choices, and architecture of EmmaNeigh. It's intended for developers who want to understand, modify, or extend the tool.

---

## Design Philosophy

### Core Principles

1. **Zero Dependencies on User's System**
   - Portable Python runtime included
   - All libraries pre-packaged
   - No system PATH modifications
   - No registry changes

2. **Offline-First Architecture**
   - No network calls
   - No cloud dependencies
   - No telemetry
   - Client data never leaves local machine

3. **Fail-Safe Detection**
   - Prioritize precision over recall
   - Better to miss a signature than create a false positive
   - Explicit is better than implicit

4. **Person-Centric Organization**
   - Group by individual signer, not entity
   - Combine all obligations per person across all documents
   - Support multiple signers on same page

---

## Technology Stack

### Core Components

**Python 3.11.2** (Portable/Embeddable)
- Chosen for: wide library support, PDF manipulation capabilities
- Deployment: Embeddable ZIP distribution (no installer)
- Constraints: No tkinter/GUI libraries on locked-down systems

**Key Libraries:**
- **PyMuPDF (fitz)** - PDF parsing and page extraction
- **pandas** - Data manipulation and Excel output
- **openpyxl** - Excel file creation
- **re** - Regular expressions for name normalization

### Distribution Model

**BAT Launcher** (not EXE)
- Avoids PyInstaller/tkinter complications
- Works on locked corporate Windows machines
- Provides visible logging for debugging
- Easier to audit and understand

---

## System Architecture

### High-Level Flow

```
┌─────────────────┐
│  User Action    │  Drag folder onto .bat
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  BAT Launcher   │  Validates folder, calls Python
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  PDF Scanner    │  Iterate through all PDFs
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Page Classifier │  Identify signature pages
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Signer Extractor│  Parse individual names
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Name Normalizer │  Canonicalize names
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Packet Builder  │  Group pages by person
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Output Generator│  Create PDFs and Excel files
└─────────────────┘
```

---

## Core Algorithms

### 1. Signature Page Detection

**Purpose:** Identify which pages contain signature blocks

**Algorithm:**
```python
def is_signature_page(text):
    """
    Returns True if page contains a signature block
    """
    SIGNATURE_KEYWORDS = [
        "IN WITNESS WHEREOF",
        "BY:",
        "NAME:",
        "TITLE:",
        "DATE:",
        "SIGNATURE",
        "________________"
    ]
    
    hits = sum(1 for keyword in SIGNATURE_KEYWORDS 
               if keyword in text.upper())
    
    return hits >= 2  # Require multiple markers
```

**Rationale:**
- Single keyword could be false positive
- Standard signature blocks have multiple consistent markers
- Uppercase comparison handles case variations

**Known Limitations:**
- Non-standard signature blocks may be missed
- Initial-only blocks not currently detected
- Scanned PDFs without OCR won't have text layer

### 2. Signer Extraction

**Purpose:** Extract individual person names from signature blocks

**Two-Tier Approach:**

**Tier 1: Explicit Name Field (Preferred)**
```python
def extract_from_name_field(lines, by_index):
    """
    Look downward from BY: for explicit Name: field
    """
    for j in range(1, 7):  # Search next 6 lines
        if by_index + j >= len(lines):
            break
        candidate = lines[by_index + j]
        if candidate.upper().startswith("NAME:"):
            # Extract everything after "Name:"
            return normalize_name(candidate.split(":", 1)[1])
    return None
```

**Tier 2: Proximity-Based Fallback**
```python
def extract_from_proximity(lines, by_index):
    """
    If no Name: field, look for probable person name
    """
    for j in range(1, 7):
        if by_index + j >= len(lines):
            break
        candidate = normalize_name(lines[by_index + j])
        if is_probable_person(candidate):
            return candidate
    return None
```

**Filtering Entity Names:**
```python
def is_probable_person(name):
    """
    Heuristic: persons have 2-4 words, no entity suffixes
    """
    ENTITY_TERMS = ["LLC", "INC", "CORP", "LP", "LLP", "TRUST"]
    
    if any(term in name for term in ENTITY_TERMS):
        return False
    
    word_count = len(name.split())
    return 2 <= word_count <= 4
```

**Why This Works:**
- Tier 1 matches 90%+ of financing documents
- Tier 2 catches edge cases
- Entity filtering prevents "ABC Holdings LLC" from becoming a packet

### 3. Name Normalization

**Purpose:** Ensure same person gets one packet despite variations

**Algorithm:**
```python
def normalize_name(name):
    """
    Canonicalize name to avoid duplicates
    """
    name = name.upper()                    # "John Smith" → "JOHN SMITH"
    name = re.sub(r"[.,]", "", name)       # Remove punctuation
    name = re.sub(r"\s+", " ", name)       # Collapse whitespace
    name = name.strip()                     # Trim edges
    return name
```

**Handles:**
- `"John Smith"` → `"JOHN SMITH"`
- `"JOHN  SMITH"` → `"JOHN SMITH"` (extra spaces)
- `"John Smith,"` → `"JOHN SMITH"` (trailing comma)
- `"john smith"` → `"JOHN SMITH"` (lowercase)

**Known Edge Cases:**
- Different people with same name (e.g., two John Smiths)
- Misspellings in source PDFs
- Hyphenated names with inconsistent hyphens

---

## Data Model

### Internal Representation

**Row Structure:**
```python
{
    "Signer Name": "JOHN SMITH",      # Normalized
    "Document": "Credit_Agreement.pdf",
    "Page": 87
}
```

**Master Table:**
- One row per signature obligation
- Sorted by: Signer Name, Document, Page
- Used for: QC, audit trail, completeness checks

**Grouping Logic:**
```python
for signer, group in df.groupby("Signer Name"):
    # group contains all pages for this person
    # across all documents
```

### Output Files

**PDF Packets:**
- Filename: `signature_packet - {NORMALIZED_NAME}.pdf`
- Contents: Extracted pages in document order
- Format: Exact copy of original pages (no reformatting)

**Excel Tables:**
- Master index: All obligations
- Per-person tables: Filtered view for that signer
- Columns: Signer Name | Document | Page

---

## File System Design

### Directory Structure

```
Tool Distribution:
Signature_Packet_Tool/
├── run_signature_packets.bat    # Entry point
├── python/                       # Portable runtime
│   ├── python.exe
│   ├── python311.dll
│   ├── Lib/                      # Standard library
│   └── site-packages/            # Installed packages
└── src/
    └── build_signature_packets.py

After Processing:
user_pdf_folder/
├── document1.pdf
├── document2.pdf
└── signature_packets_output/     # Auto-created
    ├── packets/
    │   └── signature_packet - JOHN SMITH.pdf
    └── tables/
        ├── MASTER_SIGNATURE_INDEX.xlsx
        └── signature_packet - JOHN SMITH.xlsx
```

### Path Handling

**Input Folder:**
```python
INPUT_DIR = sys.argv[1]  # Passed from BAT file
```

**Output Folders:**
```python
OUTPUT_BASE = os.path.join(INPUT_DIR, "signature_packets_output")
OUTPUT_PDF_DIR = os.path.join(OUTPUT_BASE, "packets")
OUTPUT_TABLE_DIR = os.path.join(OUTPUT_BASE, "tables")

os.makedirs(OUTPUT_PDF_DIR, exist_ok=True)
os.makedirs(OUTPUT_TABLE_DIR, exist_ok=True)
```

**Why This Design:**
- Output alongside input (easy to find)
- Namespace separation (`signature_packets_output/`)
- Idempotent (can run multiple times safely)

---

## Error Handling

### Validation Strategy

**Pre-Execution Checks:**
```python
if not os.path.isdir(INPUT_DIR):
    raise RuntimeError(f"Invalid folder: {INPUT_DIR}")

if len(rows) == 0:
    raise RuntimeError("No signers detected")
```

**Runtime Logging:**
- Print each document as it's scanned
- Print each packet as it's built
- Final message with output location

### Known Failure Modes

1. **No signers detected**
   - Cause: Non-standard signature blocks
   - Mitigation: Adjust SIGNATURE_KEYWORDS or detection logic

2. **Duplicate packets (name collisions)**
   - Cause: Two different people with identical names
   - Current behavior: Combined into one packet
   - Future: Add disambiguation logic

3. **Missing pages**
   - Cause: Scanned PDFs without text layer
   - Mitigation: User must OCR files first

---

## Performance Characteristics

### Complexity Analysis

**Time Complexity:**
- PDF scanning: O(n × m) where n = documents, m = avg pages
- Signer extraction: O(p) where p = signature pages
- Packet building: O(s × p) where s = unique signers

**Space Complexity:**
- In-memory: O(p) - signature page metadata only
- Disk I/O: Streaming page extraction (constant memory)

### Benchmarks

**Typical Deal:**
- 25 documents
- 500 total pages
- 100 signature pages
- 20 unique signers

**Processing Time:** ~15-30 seconds on modern laptop

**Scaling:**
- Handles 100+ documents comfortably
- Hundreds of signers is fine (linear scaling)
- Bottleneck is PDF parsing, not Python logic

---

## Security Considerations

### Threat Model

**Assumptions:**
- User has legitimate access to PDFs
- PDFs may contain confidential client data
- Processing happens on potentially locked-down corporate machines

**Guarantees:**
1. No network I/O
2. No file system access outside specified folder
3. No logging to external systems
4. Source code is auditable

### Data Flow Diagram

```
┌──────────────┐
│  User PDFs   │  (Input)
└──────┬───────┘
       │
       ▼
┌──────────────┐
│  Read Files  │  (Local disk only)
└──────┬───────┘
       │
       ▼
┌──────────────┐
│   Process    │  (In-memory)
└──────┬───────┘
       │
       ▼
┌──────────────┐
│ Write Output │  (Same directory)
└──────────────┘

No external systems touched
```

---

## Extending the Tool

### Adding New Features

**1. Initial-Only Page Detection**

Add to signature keywords:
```python
INITIAL_KEYWORDS = ["Initial:", "Initials:", "_______ (initials)"]
```

Modify classification:
```python
def is_initial_page(text):
    return any(kw in text for kw in INITIAL_KEYWORDS)
```

**2. Capacity Extraction (Borrower/Guarantor)**

After extracting signer name, look upward for capacity:
```python
def extract_capacity(lines, by_index):
    """
    Look upward for 'BORROWER:', 'GUARANTOR:', etc.
    """
    for j in range(1, 10):
        if by_index - j < 0:
            break
        line = lines[by_index - j].upper()
        if any(cap in line for cap in ["BORROWER", "GUARANTOR", "LENDER"]):
            return line.strip()
    return "UNKNOWN"
```

**3. Cover Sheets**

Generate a cover page per packet:
```python
from reportlab.pdfgen import canvas

def create_cover_sheet(signer_name, obligations):
    """
    Create PDF cover page listing all obligations
    """
    # Use reportlab to generate first page
    # Then prepend to packet
```

### Testing Recommendations

**Test Cases:**

1. **Standard signature blocks** (most common)
2. **Multiple signers per page**
3. **Same person across multiple documents**
4. **Entity vs. individual disambiguation**
5. **Edge cases**: unusual formatting, OCR'd text
6. **Performance**: large deals (100+ docs)

**Test Data:**
- Create sanitized sample PDFs
- Include various signature block formats
- Test name normalization edge cases

---

## Deployment

### Building a Release

1. **Ensure portable Python is packaged:**
   ```bash
   cd EmmaNeigh
   # Verify python/ folder contains full runtime
   ```

2. **Create ZIP:**
   ```bash
   zip -r Signature_Packet_Tool_v1.0.zip \
       run_signature_packets.bat \
       python/ \
       src/ \
       docs/
   ```

3. **Test on clean machine:**
   - Unzip
   - Run on sample PDFs
   - Verify output

4. **Upload to GitHub Releases:**
   - Tag version (v1.0.0)
   - Upload ZIP
   - Add release notes

### Version Numbering

Follow Semantic Versioning:
- **Major (1.0.0):** Breaking changes
- **Minor (1.1.0):** New features
- **Patch (1.0.1):** Bug fixes

---

## Future Roadmap

### Planned Enhancements

**Short Term:**
- Initial-only page detection
- Capacity field extraction
- Better duplicate name handling
- Self-check/validation mode

**Medium Term:**
- Cover sheet generation
- DocuSign envelope ordering
- Batch processing mode
- Configuration file for custom keywords

**Long Term:**
- Machine learning for non-standard blocks
- Integration with deal management systems
- Automated quality control checks
- Multi-language support

---

## Lessons Learned

### Design Decisions That Worked

1. **BAT over EXE** - Avoided tkinter complexity, better debuggability
2. **Person-centric logic** - Matches how lawyers actually work
3. **Portable Python** - Bypasses IT restrictions
4. **Name normalization** - Handles real-world inconsistencies

### Design Decisions to Reconsider

1. **No GUI** - Could add optional web UI for power users
2. **Simple heuristics** - Could use ML for better detection
3. **Excel output** - Could support CSV or JSON for automation

---

## Contributing Guidelines

### Code Standards

- **Type hints** preferred
- **Docstrings** for all functions
- **Comments** for non-obvious logic
- **Error messages** must be user-friendly

### Pull Request Process

1. Fork repository
2. Create feature branch
3. Add tests if applicable
4. Update documentation
5. Submit PR with clear description

### Code Review Criteria

- Does it solve a real problem?
- Is it maintainable?
- Does it preserve security guarantees?
- Is documentation updated?

---

## References

- [PyMuPDF Documentation](https://pymupdf.readthedocs.io/)
- [pandas Documentation](https://pandas.pydata.org/docs/)
- [Python Embeddable Package](https://docs.python.org/3/using/windows.html#the-embeddable-package)

---

**Last Updated:** January 2026  
**Maintainer:** Raam Tambe
