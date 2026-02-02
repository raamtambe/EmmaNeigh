const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');
const archiver = require('archiver');

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

// Get the path to the Python processor
function getPythonPath() {
  if (isDev) {
    // In development, use the Python script directly
    return {
      executable: 'python3',
      script: path.join(__dirname, '../python/main.py'),
      useScript: true
    };
  } else {
    // In production, use the bundled executable
    const resourcePath = process.resourcesPath;
    const processorPath = path.join(resourcePath, 'python', 'processor');
    return {
      executable: processorPath,
      useScript: false
    };
  }
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
  return new Promise((resolve, reject) => {
    const pythonConfig = getPythonPath();
    const tempDir = path.join(app.getPath('temp'), 'emmaneigh-' + Date.now());

    // Create temp directory
    fs.mkdirSync(tempDir, { recursive: true });

    // Copy files to temp directory
    filePaths.forEach(filePath => {
      const fileName = path.basename(filePath);
      fs.copyFileSync(filePath, path.join(tempDir, fileName));
    });

    let args;
    let executable;

    if (pythonConfig.useScript) {
      executable = pythonConfig.executable;
      args = [pythonConfig.script, 'signature-packets', tempDir];
    } else {
      executable = pythonConfig.executable;
      args = ['signature-packets', tempDir];
    }

    const pythonProcess = spawn(executable, args);

    let outputData = '';

    pythonProcess.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(line => line.trim());
      lines.forEach(line => {
        try {
          const progress = JSON.parse(line);
          mainWindow.webContents.send('progress', progress);
        } catch (e) {
          // Not JSON, ignore
        }
      });
    });

    pythonProcess.stderr.on('data', (data) => {
      console.error('Python stderr:', data.toString());
    });

    pythonProcess.on('close', async (code) => {
      if (code === 0) {
        // Create ZIP file from output
        const outputDir = path.join(tempDir, 'signature_packets_output');
        const zipPath = path.join(app.getPath('temp'), `EmmaNeigh-Output-${Date.now()}.zip`);

        try {
          await createZipFromDirectory(outputDir, zipPath);
          resolve({ success: true, zipPath });
        } catch (err) {
          reject({ success: false, error: err.message });
        }
      } else {
        reject({ success: false, error: `Process exited with code ${code}` });
      }
    });

    pythonProcess.on('error', (err) => {
      reject({ success: false, error: err.message });
    });
  });
});

// Create execution version
ipcMain.handle('create-execution-version', async (event, originalPath, signedPath, insertAfter) => {
  return new Promise((resolve, reject) => {
    const pythonConfig = getPythonPath();

    let args;
    let executable;

    if (pythonConfig.useScript) {
      executable = pythonConfig.executable;
      args = [pythonConfig.script, 'execution-version', originalPath, signedPath, insertAfter.toString()];
    } else {
      executable = pythonConfig.executable;
      args = ['execution-version', originalPath, signedPath, insertAfter.toString()];
    }

    const pythonProcess = spawn(executable, args);

    let outputPath = '';

    pythonProcess.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(line => line.trim());
      lines.forEach(line => {
        try {
          const result = JSON.parse(line);
          if (result.type === 'progress') {
            mainWindow.webContents.send('progress', result);
          } else if (result.type === 'complete') {
            outputPath = result.outputPath;
          }
        } catch (e) {
          // Not JSON, ignore
        }
      });
    });

    pythonProcess.stderr.on('data', (data) => {
      console.error('Python stderr:', data.toString());
    });

    pythonProcess.on('close', (code) => {
      if (code === 0 && outputPath) {
        resolve({ success: true, outputPath });
      } else {
        reject({ success: false, error: `Process exited with code ${code}` });
      }
    });

    pythonProcess.on('error', (err) => {
      reject({ success: false, error: err.message });
    });
  });
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
