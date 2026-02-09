# EmmaNeigh - Technical Architecture

## Overview

This document explains the technical design and architecture of EmmaNeigh. Intended for developers who want to understand, modify, or extend the tool.

---

## Technology Stack

### Frontend (Electron)

- **Electron 28** - Cross-platform desktop framework
- **HTML/CSS/JavaScript** - UI rendering
- **sql.js** - In-browser SQLite for history

### Backend (Python Processors)

- **Python 3.11** - Core processing language
- **PyMuPDF (fitz)** - PDF parsing and manipulation
- **python-docx** - Word document processing
- **pandas** - Data manipulation
- **openpyxl** - Excel file creation

### Build System

- **PyInstaller** - Bundle Python to executables
- **electron-builder** - Package Electron app
- **GitHub Actions** - CI/CD pipeline

---

## Application Architecture

```
┌─────────────────────────────────────────┐
│           Electron Main Process          │
│  ┌─────────────────────────────────┐    │
│  │         main.js                  │    │
│  │  - Window management             │    │
│  │  - IPC handlers                  │    │
│  │  - SQLite database               │    │
│  │  - Auto-updater                  │    │
│  └─────────────────────────────────┘    │
└───────────────┬─────────────────────────┘
                │ IPC
┌───────────────▼─────────────────────────┐
│          Electron Renderer               │
│  ┌─────────────────────────────────┐    │
│  │         index.html               │    │
│  │  - Sidebar navigation            │    │
│  │  - Feature tabs                  │    │
│  │  - Progress indicators           │    │
│  │  - File uploads                  │    │
│  └─────────────────────────────────┘    │
└───────────────┬─────────────────────────┘
                │ spawn
┌───────────────▼─────────────────────────┐
│          Python Processors               │
│  ┌─────────────────────────────────┐    │
│  │  signature_packets.py            │    │
│  │  execution_version.py            │    │
│  │  sigblock_workflow.py            │    │
│  │  document_collator.py            │    │
│  │  email_csv_parser.py             │    │
│  │  time_tracker.py                 │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

---

## Directory Structure

```
EmmaNeigh/
├── desktop-app/
│   ├── main.js              # Electron main process
│   ├── index.html           # UI (HTML/CSS/JS)
│   ├── package.json         # Electron dependencies
│   ├── python/
│   │   ├── signature_packets.py
│   │   ├── execution_version.py
│   │   ├── sigblock_workflow.py
│   │   ├── sigblock_generator.py
│   │   ├── checklist_parser.py
│   │   ├── incumbency_parser.py
│   │   ├── document_collator.py
│   │   ├── email_csv_parser.py
│   │   ├── time_tracker.py
│   │   └── requirements.txt
│   └── resources/
│       ├── win/             # Windows executables
│       └── mac/             # macOS executables
├── .github/
│   └── workflows/
│       └── build.yml        # CI/CD pipeline
├── README.md
├── CHANGELOG.md
├── USER_GUIDE.md
├── INSTALLATION.md
├── ARCHITECTURE.md
└── SECURITY_AND_PRIVACY.md
```

---

## Core Algorithms

### Signature Detection

Detects signature blocks using pattern matching:

```python
# Keywords indicating signature content
BY_MARKER = "BY:"
NAME_FIELD = "NAME:"
TITLE_FIELD = "TITLE:"

# Table-based detection
NAME_HEADERS = ["NAME", "PRINTED NAME", "SIGNATORY"]
SIGNATURE_HEADERS = ["SIGNATURE", "SIGN", "BY"]
```

**Two-tier approach:**
1. **BY:/Name: blocks** - Traditional signature format
2. **Signature tables** - Grid-based signature pages

### Signer Extraction

```python
def extract_signer(text):
    # 1. Look for explicit "Name:" field
    if "NAME:" in text:
        return parse_name_field(text)
    
    # 2. Look for probable person name near "BY:"
    return find_person_near_by_marker(text)

def is_probable_person(name):
    # Filter out entities
    ENTITY_TERMS = ["LLC", "INC", "CORP", "LP"]
    if any(term in name for term in ENTITY_TERMS):
        return False
    return 2 <= word_count <= 4
```

### Name Normalization

```python
def normalize_name(name):
    name = name.upper()           # JOHN SMITH
    name = re.sub(r"[.,]", "", name)  # Remove punctuation
    name = re.sub(r"\s+", " ", name)  # Collapse whitespace
    return name.strip()
```

### Format Preservation

```python
def get_output_format(input_path):
    ext = os.path.splitext(input_path)[1].lower()
    return ext  # '.pdf' or '.docx'

# DOCX in = DOCX out
# PDF in = PDF out
```

---

## IPC Communication

### Electron ↔ Python

```javascript
// main.js - IPC handler
ipcMain.handle('create-signature-packets', async (event, files) => {
    // Write config to temp file
    const configPath = path.join(tempDir, 'config.json');
    fs.writeFileSync(configPath, JSON.stringify({ files }));
    
    // Spawn Python process
    const proc = spawn(processorPath, ['--config', configPath]);
    
    // Parse JSON output
    proc.stdout.on('data', (data) => {
        const msg = JSON.parse(data);
        if (msg.type === 'progress') {
            event.sender.send('progress', msg);
        }
    });
});
```

### Python Output Protocol

```python
def emit(msg_type, **kwargs):
    """Output JSON message to stdout."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)

# Usage:
emit("progress", percent=50, message="Processing...")
emit("result", success=True, outputPath="/path/to/output")
emit("error", message="Something went wrong")
```

---

## Build Process

### GitHub Actions Workflow

```yaml
jobs:
  build-windows:
    - Install Python dependencies
    - Build Python executables (PyInstaller)
    - Install npm dependencies
    - Copy executables to resources/win
    - Build Electron app (electron-builder)
    - Upload to Release

  build-mac:
    - Same process for macOS
    - Upload .dmg to Release
```

### PyInstaller Configuration

```bash
pyinstaller --onefile --name signature_packets signature_packets.py
```

### Electron Builder Configuration

```json
{
  "win": {
    "target": ["nsis", "portable"]
  },
  "mac": {
    "target": "dmg"
  },
  "extraResources": [
    {
      "from": "resources/${os}",
      "to": "processor"
    }
  ]
}
```

---

## Data Storage

### SQLite History Database

```sql
-- Location: %APPDATA%/emmaneigh/history.db

CREATE TABLE usage_history (
    id INTEGER PRIMARY KEY,
    timestamp TEXT,
    feature TEXT,
    input_count INTEGER,
    output_count INTEGER,
    duration_ms INTEGER
);
```

### Access Pattern

```javascript
// Using sql.js (in-memory SQLite)
const SQL = await initSqlJs();
const db = new SQL.Database(existingData);

db.run(`
    INSERT INTO usage_history (timestamp, feature, ...)
    VALUES (?, ?, ...)
`, [new Date().toISOString(), 'signature_packets', ...]);
```

---

## Performance

### Typical Processing Times

| Document Size | Pages | Processing Time |
|--------------|-------|-----------------|
| 10 PDFs, 100 pages | ~15 sec |
| 50 PDFs, 500 pages | ~45 sec |
| 100+ PDFs | 1-2 min |

### Memory Usage

- Electron: ~100-150 MB
- Python processor: Varies with document size
- Large PDFs: May spike to 500 MB+

---

## Extending the Tool

### Adding a New Processor

1. Create `desktop-app/python/new_feature.py`
2. Implement `emit()` protocol for progress/results
3. Add IPC handler in `main.js`
4. Add UI tab in `index.html`
5. Update `build.yml` to include in build

### Modifying Signature Detection

Edit `signature_packets.py`:
- `extract_person_signers()` - Main detection logic
- `NAME_HEADERS` / `SIGNATURE_HEADERS` - Keywords
- `is_probable_person()` - Entity filtering

---

## Contributing

### Code Standards

- Python: Type hints, docstrings, PEP 8
- JavaScript: ES6+, clear variable names
- Comments for non-obvious logic

### Pull Request Process

1. Fork repository
2. Create feature branch
3. Add tests if applicable
4. Update documentation
5. Submit PR with description

---

## References

- [Electron Documentation](https://www.electronjs.org/docs)
- [PyMuPDF Documentation](https://pymupdf.readthedocs.io/)
- [python-docx Documentation](https://python-docx.readthedocs.io/)

---

**Last Updated:** February 2026
**Maintainer:** Raam Tambe
