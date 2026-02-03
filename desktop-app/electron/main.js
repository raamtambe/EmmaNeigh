const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const pdfProcessor = require('./pdf-processor');

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

    // Process using JavaScript PDF processor
    const result = await pdfProcessor.processSignaturePackets(tempDir, progressCallback);

    if (result.success) {
      // Create ZIP file from output
      const outputDir = result.outputPath;
      const zipPath = path.join(app.getPath('temp'), `EmmaNeigh-Output-${Date.now()}.zip`);

      await createZipFromDirectory(outputDir, zipPath);
      return { success: true, zipPath, ...result };
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

    const result = await pdfProcessor.createExecutionVersion(
      originalPath,
      signedPath,
      insertAfter,
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
