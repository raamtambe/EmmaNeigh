# EmmaNeigh Desktop Application

**Signature Packet Automation for M&A Transactions**

A cross-platform desktop application that automates the creation of signature packets and execution versions for legal documents.

## Features

### 1. Create Signature Packets
- Upload multiple PDF documents
- Automatically detect signature pages
- Group pages by signer
- Generate individual signature packets
- Export as ZIP file with Excel tracking sheets

### 2. Create Execution Version
- Merge signed pages back into original documents
- Automatically unlock DocuSign protected PDFs
- Specify exact page insertion point
- Creates final execution version PDF

### 3. DocuSign Integration (Coming Soon)
- Send signature packets directly via DocuSign
- Track signature status

## Security & Privacy

- **100% Local Processing** - All document processing happens on your machine
- **No Network Connections** - The application never connects to the internet
- **No Cloud Storage** - Your documents stay on your local disk
- **No Telemetry** - No analytics, tracking, or data collection
- **No Admin Rights Required** - Runs as a portable application

## Installation

### Option 1: Portable Application (Recommended)

1. Download `EmmaNeigh-Portable.zip` from the releases
2. Extract to any folder
3. Double-click `EmmaNeigh.exe` (Windows) or `EmmaNeigh.app` (Mac)

### Option 2: Build from Source

#### Prerequisites
- Node.js 18+
- Python 3.10+
- npm or yarn

#### Steps

1. Clone the repository:
```bash
git clone https://github.com/raamtambe/EmmaNeigh.git
cd EmmaNeigh/desktop-app
```

2. Install Node.js dependencies:
```bash
npm install
cd frontend && npm install && cd ..
```

3. Install Python dependencies:
```bash
cd python
pip install -r requirements.txt
cd ..
```

4. Run in development mode:
```bash
npm run dev
```

5. Build for production:
```bash
# Build frontend
npm run build:react

# Build Python processor (requires PyInstaller)
npm run build:python

# Build Electron app
npm run dist
```

## Project Structure

```
desktop-app/
├── electron/           # Electron main process
│   ├── main.js         # Entry point
│   └── preload.js      # IPC bridge
├── frontend/           # React UI
│   └── src/
│       ├── pages/      # Main views
│       └── components/ # Reusable components
├── python/             # PDF processing
│   ├── main.py         # Entry point
│   └── processors/     # Processing modules
└── package.json
```

## Development

### Running locally

```bash
# Start the development server
npm run dev

# This will:
# 1. Start Vite dev server for React (port 3000)
# 2. Launch Electron pointing to the dev server
```

### Building

```bash
# Windows portable
npm run dist:win

# macOS DMG
npm run dist:mac

# Linux AppImage
npm run dist:linux
```

## Technical Details

### Architecture

- **Frontend**: React + Vite + Tailwind CSS
- **Desktop Shell**: Electron
- **PDF Processing**: Python with PyMuPDF
- **IPC Communication**: Electron IPC + JSON streams

### How Signature Detection Works

1. Scans each PDF page for signature markers:
   - "BY:" field
   - "NAME:" field
   - Signature block patterns

2. Extracts signer names using a two-tier approach:
   - Tier 1: Look for explicit "Name:" field
   - Tier 2: Identify probable person names near "BY:"

3. Groups pages by normalized signer name

### How Execution Version Works

1. Opens original PDF (without signature pages)
2. Opens signed PDF from DocuSign
3. Unlocks DocuSign restrictions (permission-only, not password)
4. Merges pages at specified insertion point
5. Saves as new execution version PDF

## License

MIT License - See LICENSE file for details.

## Support

For questions or issues, please open a GitHub issue or contact the maintainer.
