# EmmaNeigh Desktop App - Setup Instructions

## Quick Fix for npm Permissions

If you see npm permission errors, run this command in Terminal:

```bash
sudo chown -R $(whoami) ~/.npm
```

Then try installing again.

## Installation Steps

### 1. Install Main Dependencies

```bash
cd desktop-app
npm install
```

### 2. Install Frontend Dependencies

```bash
cd frontend
npm install
cd ..
```

### 3. Install Python Dependencies

```bash
cd python
pip3 install -r requirements.txt
cd ..
```

### 4. Run in Development Mode

```bash
npm run dev
```

This will:
1. Start the React development server on http://localhost:3000
2. Launch the Electron app pointing to that server

### 5. Build for Distribution

#### Build Frontend
```bash
npm run build:react
```

#### Build Python Processor (requires PyInstaller)
```bash
pip3 install pyinstaller
cd python
pyinstaller --onefile --name processor main.py
cd ..
```

#### Build Electron App
```bash
# For Windows portable
npm run dist:win

# For macOS
npm run dist:mac

# For Linux
npm run dist:linux
```

## Troubleshooting

### "Cannot find module" errors
Make sure to run `npm install` in both the root and frontend directories.

### Python not found
Make sure Python 3 is installed and available as `python3` in your PATH.

### Electron doesn't start
Check that port 3000 is not in use by another application.

## Project Structure

```
desktop-app/
├── electron/
│   ├── main.js        # Electron entry point
│   └── preload.js     # Secure IPC bridge
├── frontend/
│   ├── src/
│   │   ├── App.jsx
│   │   ├── pages/
│   │   │   ├── MainMenu.jsx
│   │   │   ├── SignaturePackets.jsx
│   │   │   └── ExecutionVersion.jsx
│   │   └── components/
│   │       ├── FileUpload.jsx
│   │       ├── HorseAnimation.jsx
│   │       └── ProgressBar.jsx
│   └── package.json
├── python/
│   ├── main.py
│   ├── processors/
│   │   ├── signature_packets.py
│   │   └── execution_version.py
│   └── requirements.txt
└── package.json
```
