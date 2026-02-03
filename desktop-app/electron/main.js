const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const { spawn } = require('child_process');

// Determine if we're in development or production
const isDev = !app.isPackaged;

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    icon: path.join(__dirname, '../assets/icon.png'),
    title: 'EmmaNeigh - Signature Packet Automation',
  });

  // Load the app
  if (isDev) {
    mainWindow.loadURL('http://localhost:3000');
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '../frontend/dist/index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// Get path to Python processor executable
function getPythonProcessorPath() {
  if (isDev) {
    // In development, use Python directly
    return null;
  }

  // In production, use bundled executable
  const resourcesPath = process.resourcesPath;
  const platform = process.platform;

  if (platform === 'win32') {
    return path.join(resourcesPath, 'python-processor', 'signature_packets.exe');
  } else if (platform === 'darwin') {
    return path.join(resourcesPath, 'python-processor', 'signature_packets');
  }

  return null;
}

// Run Python processor
function runPythonProcessor(args, progressCallback) {
  return new Promise((resolve, reject) => {
    const processorPath = getPythonProcessorPath();

    if (!processorPath) {
      reject(new Error('Python processor not found. Please use the built application.'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error(`Processor not found at: ${processorPath}`));
      return;
    }

    // Make executable on Mac
    if (process.platform === 'darwin') {
      try {
        fs.chmodSync(processorPath, '755');
      } catch (e) {
        // Ignore chmod errors
      }
    }

    const proc = spawn(processorPath, args);

    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => {
      const text = data.toString();
      stdout += text;

      // Parse progress messages (JSON lines)
      const lines = text.split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress' && progressCallback) {
            progressCallback(msg.stage, msg.percent, msg.message);
          }
        } catch (e) {
          // Not JSON, ignore
        }
      }
    });

    proc.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    proc.on('close', (code) => {
      if (code === 0) {
        // Find the last JSON result line
        const lines = stdout.split('\n').filter(l => l.trim());
        for (let i = lines.length - 1; i >= 0; i--) {
          try {
            const result = JSON.parse(lines[i]);
            if (result.type === 'result') {
              resolve(result);
              return;
            }
          } catch (e) {
            // Not JSON, continue
          }
        }
        reject(new Error('No result from processor'));
      } else {
        reject(new Error(stderr || `Process exited with code ${code}`));
      }
    });

    proc.on('error', (err) => {
      reject(err);
    });
  });
}

// IPC Handlers

// Select files dialog
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });
  return result.filePaths;
});

// Select folder dialog
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
  });
  return result.filePaths[0] || null;
});

// Save file dialog
ipcMain.handle('save-file', async (event, defaultName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultName,
    filters: [
      { name: 'ZIP Archive', extensions: ['zip'] },
      { name: 'PDF Document', extensions: ['pdf'] },
    ],
  });
  return result.filePath || null;
});

// Process signature packets
ipcMain.handle('process-signature-packets', async (event, filePaths) => {
  try {
    const tempDir = path.join(app.getPath('temp'), 'emmaneigh-' + Date.now());

    // Create temp directory
    fs.mkdirSync(tempDir, { recursive: true });

    // Copy files to temp directory
    for (const filePath of filePaths) {
      const fileName = path.basename(filePath);
      fs.copyFileSync(filePath, path.join(tempDir, fileName));
    }

    // Progress callback
    const progressCallback = (stage, percent, message) => {
      mainWindow.webContents.send('progress', {
        type: 'progress',
        stage,
        percent,
        message
      });
    };

    // Run Python processor
    const result = await runPythonProcessor(
      ['signature-packets', tempDir],
      progressCallback
    );

    if (result.success) {
      // Create ZIP file from output
      const outputDir = result.outputPath;
      const zipPath = path.join(app.getPath('temp'), `EmmaNeigh-Output-${Date.now()}.zip`);

      await createZipFromDirectory(outputDir, zipPath);
      return {
        success: true,
        zipPath,
        packetsCreated: result.packetsCreated,
        packets: result.packets
      };
    } else {
      return { success: false, error: result.error || 'Processing failed' };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// Create execution version
ipcMain.handle('create-execution-version', async (event, originalPath, signedPath, insertAfter) => {
  try {
    // Progress callback
    const progressCallback = (stage, percent, message) => {
      mainWindow.webContents.send('progress', {
        type: 'progress',
        stage,
        percent,
        message
      });
    };

    // Run Python processor
    const result = await runPythonProcessor(
      ['execution-version', originalPath, signedPath, String(insertAfter)],
      progressCallback
    );

    return result;
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// Copy file to user-selected location
ipcMain.handle('copy-file', async (event, sourcePath, destPath) => {
  try {
    fs.copyFileSync(sourcePath, destPath);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// Helper function to create ZIP from directory
function createZipFromDirectory(sourceDir, outputPath) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => resolve(outputPath));
    archive.on('error', (err) => reject(err));

    archive.pipe(output);
    archive.directory(sourceDir, false);
    archive.finalize();
  });
}
