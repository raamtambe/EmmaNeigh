const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const archiver = require('archiver');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 700,
    minWidth: 500,
    minHeight: 600,
    resizable: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    titleBarStyle: 'hiddenInset',
    backgroundColor: '#0a0a0a'
  });

  mainWindow.loadFile('index.html');

  // Remove menu bar
  mainWindow.setMenuBarVisibility(false);
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// Get the processor path
function getProcessorPath(processorName = 'signature_packets') {
  const isDev = !app.isPackaged;

  if (isDev) {
    // Development - won't work without building Python first
    return null;
  }

  // Production - bundled processor
  const resourcesPath = process.resourcesPath;
  if (process.platform === 'win32') {
    return path.join(resourcesPath, 'processor', `${processorName}.exe`);
  } else {
    return path.join(resourcesPath, 'processor', processorName);
  }
}

// Process signature packets
ipcMain.handle('process-folder', async (event, folderPath) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath();

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    // Make executable on Mac
    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    const proc = spawn(processorPath, [folderPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('progress', msg);
          } else if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            reject(new Error(msg.message));
          }
        } catch (e) {}
      }
    });

    proc.stderr.on('data', (data) => {
      console.error('stderr:', data.toString());
    });

    proc.on('close', (code) => {
      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Processing failed with code ' + code));
      }
    });

    proc.on('error', (err) => {
      reject(err);
    });
  });
});

// Create ZIP from output folder
ipcMain.handle('create-zip', async (event, outputPath) => {
  const zipPath = path.join(app.getPath('temp'), `EmmaNeigh-Output-${Date.now()}.zip`);

  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => resolve(zipPath));
    archive.on('error', reject);

    archive.pipe(output);
    archive.directory(outputPath, false);
    archive.finalize();
  });
});

// Save ZIP to user location
ipcMain.handle('save-zip', async (event, zipPath) => {
  const { filePath } = await dialog.showSaveDialog(mainWindow, {
    defaultPath: 'EmmaNeigh-Signature-Packets.zip',
    filters: [{ name: 'ZIP Archive', extensions: ['zip'] }]
  });

  if (filePath) {
    fs.copyFileSync(zipPath, filePath);
    return filePath;
  }
  return null;
});

// Open folder in explorer
ipcMain.handle('open-folder', async (event, folderPath) => {
  shell.openPath(folderPath);
});

// Select folder dialog
ipcMain.handle('select-folder', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });
  return filePaths[0] || null;
});

// Select single PDF file
ipcMain.handle('select-pdf', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'PDF Documents', extensions: ['pdf'] }]
  });
  return filePaths[0] || null;
});

// Process execution version
ipcMain.handle('process-execution-version', async (event, originalsFolder, signedPdfPath) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath('execution_version');

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    // Make executable on Mac
    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    const proc = spawn(processorPath, [originalsFolder, signedPdfPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('progress', msg);
          } else if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            reject(new Error(msg.message));
          }
        } catch (e) {}
      }
    });

    proc.stderr.on('data', (data) => {
      console.error('stderr:', data.toString());
    });

    proc.on('close', (code) => {
      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Processing failed with code ' + code));
      }
    });

    proc.on('error', (err) => {
      reject(err);
    });
  });
});
