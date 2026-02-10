const { app, BrowserWindow, ipcMain, dialog, shell, safeStorage } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const archiver = require('archiver');

// SQLite for usage history
let db = null;
let SQL = null;
const historyDbPath = path.join(app.getPath('userData'), 'emmaneigh_history.db');

async function initDatabase() {
  try {
    const initSqlJs = require('sql.js');
    SQL = await initSqlJs();

    // Try to load existing database
    if (fs.existsSync(historyDbPath)) {
      const buffer = fs.readFileSync(historyDbPath);
      db = new SQL.Database(buffer);
    } else {
      db = new SQL.Database();
    }

    // Create tables if they don't exist
    db.run(`
      CREATE TABLE IF NOT EXISTS usage_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
        user_name TEXT,
        feature TEXT,
        action TEXT,
        input_count INTEGER DEFAULT 0,
        output_count INTEGER DEFAULT 0,
        duration_ms INTEGER DEFAULT 0
      )
    `);

    db.run(`
      CREATE TABLE IF NOT EXISTS user_profile (
        id TEXT PRIMARY KEY,
        name TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        telemetry_enabled INTEGER DEFAULT 0
      )
    `);

    saveDatabase();
    console.log('Database initialized at:', historyDbPath);
  } catch (e) {
    console.error('Failed to initialize database:', e);
  }
}

function saveDatabase() {
  if (db) {
    try {
      const data = db.export();
      const buffer = Buffer.from(data);
      fs.writeFileSync(historyDbPath, buffer);
    } catch (e) {
      console.error('Failed to save database:', e);
    }
  }
}

// API Key storage (encrypted with safeStorage)
const apiKeyPath = path.join(app.getPath('userData'), 'api_key.enc');

function getApiKey() {
  try {
    // First try encrypted storage
    if (fs.existsSync(apiKeyPath)) {
      const encrypted = fs.readFileSync(apiKeyPath);
      if (safeStorage.isEncryptionAvailable()) {
        return safeStorage.decryptString(encrypted);
      }
    }
    // Fallback to environment variable
    return process.env.ANTHROPIC_API_KEY || null;
  } catch (e) {
    console.error('Failed to get API key:', e);
    return process.env.ANTHROPIC_API_KEY || null;
  }
}

function setApiKey(apiKey) {
  try {
    if (safeStorage.isEncryptionAvailable()) {
      const encrypted = safeStorage.encryptString(apiKey);
      fs.writeFileSync(apiKeyPath, encrypted);
      return true;
    }
    return false;
  } catch (e) {
    console.error('Failed to set API key:', e);
    return false;
  }
}

function deleteApiKey() {
  try {
    if (fs.existsSync(apiKeyPath)) {
      fs.unlinkSync(apiKeyPath);
    }
    return true;
  } catch (e) {
    console.error('Failed to delete API key:', e);
    return false;
  }
}

// Auto-updater (only in production)
let autoUpdater;
if (app.isPackaged) {
  try {
    autoUpdater = require('electron-updater').autoUpdater;
  } catch (e) {
    console.log('electron-updater not available');
  }
}

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 700,
    minWidth: 700,
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

app.whenReady().then(async () => {
  // Initialize database first
  await initDatabase();

  createWindow();

  // Setup auto-updater with enhanced progress reporting
  if (autoUpdater) {
    autoUpdater.autoDownload = false; // Don't auto-download, let user decide
    autoUpdater.checkForUpdates();

    autoUpdater.on('update-available', (info) => {
      if (mainWindow) {
        mainWindow.webContents.send('update-available', info);
      }
    });

    autoUpdater.on('download-progress', (progress) => {
      if (mainWindow) {
        mainWindow.webContents.send('update-progress', {
          percent: Math.round(progress.percent),
          bytesPerSecond: progress.bytesPerSecond,
          transferred: progress.transferred,
          total: progress.total
        });
      }
    });

    autoUpdater.on('update-downloaded', () => {
      if (mainWindow) {
        mainWindow.webContents.send('update-downloaded');
      }
    });

    autoUpdater.on('error', (err) => {
      if (mainWindow) {
        mainWindow.webContents.send('update-error', err.message);
      }
    });
  }
});

// IPC handler to start download
ipcMain.handle('download-update', async () => {
  if (autoUpdater) {
    await autoUpdater.downloadUpdate();
    return true;
  }
  return false;
});

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
// Accepts either a folder path string (legacy) or an object { folder: string } or { files: string[] }
ipcMain.handle('process-folder', async (event, input) => {
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

    let args;
    // Handle legacy string input (folder path)
    if (typeof input === 'string') {
      args = [input];
    } else if (input.folder) {
      args = [input.folder];
    } else if (input.files) {
      // Write config to temp file for multi-file input
      const configPath = path.join(app.getPath('temp'), `packets-config-${Date.now()}.json`);
      fs.writeFileSync(configPath, JSON.stringify({ files: input.files }));
      args = ['--config', configPath];
    } else {
      reject(new Error('Invalid input: must be folder path or { folder } or { files }'));
      return;
    }

    const proc = spawn(processorPath, args);
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
ipcMain.handle('save-zip', async (event, zipPath, suggestedName) => {
  const defaultName = suggestedName || 'EmmaNeigh-Output.zip';
  const { filePath } = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultName,
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

// Select Excel/CSV file
ipcMain.handle('select-spreadsheet', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'csv'] }]
  });
  return filePaths[0] || null;
});

// Select multiple files
ipcMain.handle('select-files', async (event, extensions) => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Documents', extensions: extensions || ['pdf', 'docx'] }]
  });
  return filePaths || [];
});

// Select multiple PDF files
ipcMain.handle('select-pdfs-multiple', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'PDF Documents', extensions: ['pdf'] }]
  });
  return filePaths || [];
});

// Parse checklist
ipcMain.handle('parse-checklist', async (event, checklistPath) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath('checklist_parser');

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    const proc = spawn(processorPath, [checklistPath]);
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

    proc.on('close', (code) => {
      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Checklist parsing failed'));
      }
    });

    proc.on('error', reject);
  });
});

// Parse incumbency certificate
ipcMain.handle('parse-incumbency', async (event, incPath) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath('incumbency_parser');

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    const proc = spawn(processorPath, [incPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            reject(new Error(msg.message));
          }
        } catch (e) {}
      }
    });

    proc.on('close', (code) => {
      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Incumbency parsing failed'));
      }
    });

    proc.on('error', reject);
  });
});

// Process signature block workflow
ipcMain.handle('process-sigblocks', async (event, config) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath('sigblock_workflow');

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    // Write config to temp file
    const configPath = path.join(app.getPath('temp'), `sigblock-config-${Date.now()}.json`);
    fs.writeFileSync(configPath, JSON.stringify(config));

    const proc = spawn(processorPath, [configPath]);
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

    proc.on('close', (code) => {
      // Clean up config file
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Signature block processing failed'));
      }
    });

    proc.on('error', reject);
  });
});

// Process execution version
// originalsInput can be a folder path string (legacy) or { folder: string } or { files: string[] }
ipcMain.handle('process-execution-version', async (event, originalsInput, signedPdfPath) => {
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

    let args;
    let configPath = null;
    // Handle different input types
    if (typeof originalsInput === 'string') {
      // Legacy: folder path string
      args = [originalsInput, signedPdfPath];
    } else if (originalsInput.folder) {
      args = [originalsInput.folder, signedPdfPath];
    } else if (originalsInput.files) {
      // Write config to temp file
      configPath = path.join(app.getPath('temp'), `exec-config-${Date.now()}.json`);
      fs.writeFileSync(configPath, JSON.stringify({
        files: originalsInput.files,
        signed_pdf: signedPdfPath
      }));
      args = ['--config', configPath];
    } else {
      reject(new Error('Invalid input: must be folder path or { folder } or { files }'));
      return;
    }

    const proc = spawn(processorPath, args);
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

// Quit app (for disclaimer decline)
ipcMain.handle('quit-app', () => {
  app.quit();
});

// Install update and restart
ipcMain.handle('install-update', () => {
  if (autoUpdater) {
    autoUpdater.quitAndInstall();
  }
});

// Select single DOCX file
ipcMain.handle('select-docx', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Word Documents', extensions: ['docx'] }]
  });
  return filePaths[0] || null;
});

// Select multiple DOCX files
ipcMain.handle('select-docx-multiple', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Word Documents', extensions: ['docx'] }]
  });
  return filePaths || [];
});

// Collate documents
ipcMain.handle('collate-documents', async (event, config) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath('document_collator');

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    // Write config to temp file
    const configPath = path.join(app.getPath('temp'), `collate-config-${Date.now()}.json`);
    fs.writeFileSync(configPath, JSON.stringify(config));

    const proc = spawn(processorPath, [configPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('collate-progress', msg);
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
      // Clean up config file
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Document collation failed with code ' + code));
      }
    });

    proc.on('error', reject);
  });
});

// Redline documents
ipcMain.handle('redline-documents', async (event, config) => {
  return new Promise((resolve, reject) => {
    const processorPath = getProcessorPath('document_redline');

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Processor not found: ' + processorPath));
      return;
    }

    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    // Write config to temp file
    const configPath = path.join(app.getPath('temp'), `redline-config-${Date.now()}.json`);
    fs.writeFileSync(configPath, JSON.stringify(config));

    const proc = spawn(processorPath, [configPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('redline-progress', msg);
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
      // Clean up config file
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Document redline failed with code ' + code));
      }
    });

    proc.on('error', reject);
  });
});

// ========== EMAIL CSV PARSING ==========

ipcMain.handle('parse-email-csv', async (event, csvPath) => {
  const processorPath = getProcessorPath('email_csv_parser');

  if (!processorPath) {
    // In development, parse CSV directly with Node.js
    try {
      const csvContent = fs.readFileSync(csvPath, 'utf-8');
      const lines = csvContent.split('\n');
      const headers = lines[0].split(',').map(h => h.trim().toLowerCase());

      const emails = [];
      for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;

        // Simple CSV parsing (doesn't handle quoted commas perfectly)
        const values = lines[i].split(',');
        const email = {
          subject: '',
          from: '',
          to: '',
          date_sent: null,
          date_received: null,
          body: ''
        };

        headers.forEach((header, idx) => {
          const value = values[idx] ? values[idx].trim() : '';
          if (header.includes('subject')) email.subject = value;
          else if (header.includes('from')) email.from = value;
          else if (header.includes('to')) email.to = value;
          else if (header.includes('body') || header.includes('content')) email.body = value;
          else if (header.includes('sent')) email.date_sent = value;
          else if (header.includes('received') || header === 'date') email.date_received = value;
        });

        emails.push(email);
      }

      const uniqueSenders = new Set(emails.map(e => e.from).filter(f => f));
      const dates = emails.map(e => e.date_sent || e.date_received).filter(d => d).sort();

      return {
        success: true,
        emails: emails,
        summary: {
          total_emails: emails.length,
          unique_senders: uniqueSenders.size,
          date_range: {
            earliest: dates[0] || null,
            latest: dates[dates.length - 1] || null
          }
        }
      };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  // Production: Use Python processor
  const configPath = path.join(app.getPath('temp'), `email_config_${Date.now()}.json`);
  const config = {
    action: 'parse',
    csv_path: csvPath
  };
  fs.writeFileSync(configPath, JSON.stringify(config));

  return new Promise((resolve, reject) => {
    const proc = spawn(processorPath, [configPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('email-progress', msg);
          } else if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            reject(new Error(msg.message));
          }
        } catch (e) {}
      }
    });

    proc.on('close', (code) => {
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Email parsing failed with code ' + code));
      }
    });

    proc.on('error', reject);
  });
});

// ========== NATURAL LANGUAGE EMAIL SEARCH ==========

ipcMain.handle('nl-search-emails', async (event, config) => {
  const processorPath = getProcessorPath('email_nl_search');

  // Get the API key
  const apiKey = config.api_key || getApiKey();

  if (!apiKey) {
    return { success: false, error: 'No API key configured. Please add your Claude API key in Settings.' };
  }

  // In development without processor, make direct API call
  if (!processorPath) {
    try {
      const https = require('https');

      const emails = config.emails || [];
      const query = config.query || '';

      if (!query) {
        return { success: false, error: 'No query provided' };
      }

      // Prepare email context (limit to 100 emails, truncate bodies)
      const emailContext = emails.slice(0, 100).map((email, i) => ({
        index: i,
        from: email.from || 'Unknown',
        to: email.to || '',
        subject: email.subject || '(No Subject)',
        body_preview: (email.body || '').substring(0, 300),
        date: email.date_received || email.date_sent || '',
        attachments: email.attachments || '',
        has_attachments: email.has_attachments || false
      }));

      const prompt = `You are an email assistant analyzing a database of emails from a legal transaction.

User Question: ${query}

Email Database (${emailContext.length} emails):
${JSON.stringify(emailContext, null, 2)}

Please analyze these emails and answer the user's question. Be specific and cite relevant emails by their index number.

Respond with a JSON object containing:
{
    "answer": "Your detailed answer to the question",
    "relevant_email_indices": [0, 5, 12],
    "confidence": 0.85,
    "summary": "One-sentence summary of your finding"
}

Respond ONLY with the JSON object, no other text.`;

      const requestData = JSON.stringify({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 1024,
        messages: [{ role: 'user', content: prompt }]
      });

      return new Promise((resolve) => {
        const options = {
          hostname: 'api.anthropic.com',
          path: '/v1/messages',
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01',
            'Content-Length': Buffer.byteLength(requestData)
          }
        };

        const req = https.request(options, (res) => {
          let body = '';
          res.on('data', chunk => body += chunk);
          res.on('end', () => {
            try {
              if (res.statusCode !== 200) {
                resolve({ success: false, error: `API error: ${res.statusCode}` });
                return;
              }

              const response = JSON.parse(body);
              const responseText = response.content[0].text.trim();

              // Try to parse JSON from response
              let result;
              try {
                // Handle markdown code blocks
                let jsonText = responseText;
                if (jsonText.startsWith('```')) {
                  const lines = jsonText.split('\n');
                  const jsonLines = [];
                  let inJson = false;
                  for (const line of lines) {
                    if (line.startsWith('```json')) { inJson = true; continue; }
                    if (line.startsWith('```')) { inJson = false; continue; }
                    if (inJson) jsonLines.push(line);
                  }
                  jsonText = jsonLines.join('\n');
                }
                result = JSON.parse(jsonText);
              } catch (e) {
                result = { answer: responseText, relevant_email_indices: [], confidence: 0.5 };
              }

              resolve({
                success: true,
                answer: result.answer || 'No answer provided',
                relevant_email_indices: result.relevant_email_indices || [],
                confidence: result.confidence || 0.5,
                summary: result.summary || '',
                query: query
              });
            } catch (e) {
              resolve({ success: false, error: `Parse error: ${e.message}` });
            }
          });
        });

        req.on('error', (e) => {
          resolve({ success: false, error: `Network error: ${e.message}` });
        });

        req.setTimeout(60000, () => {
          req.destroy();
          resolve({ success: false, error: 'Request timed out' });
        });

        req.write(requestData);
        req.end();
      });
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  // Production: spawn Python subprocess
  if (app.isPackaged && !fs.existsSync(processorPath)) {
    return { success: false, error: 'NL search processor not found' };
  }

  if (process.platform !== 'win32' && app.isPackaged) {
    try { fs.chmodSync(processorPath, '755'); } catch (e) {}
  }

  const configPath = path.join(app.getPath('temp'), `nl_search_${Date.now()}.json`);
  const configData = {
    emails: config.emails || [],
    query: config.query || '',
    api_key: apiKey
  };
  fs.writeFileSync(configPath, JSON.stringify(configData));

  return new Promise((resolve, reject) => {
    let args;
    if (app.isPackaged) {
      args = [configPath];
    } else {
      args = [path.join(__dirname, 'python', 'email_nl_search.py'), configPath];
    }

    const proc = spawn(processorPath, args);

    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'progress') {
            event.sender.send('nl-search-progress', msg);
          }
        } catch (e) {}
      }
    });

    proc.stderr.on('data', (data) => {
      console.error('NL search stderr:', data.toString());
    });

    proc.on('close', (code) => {
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('NL search failed with code ' + code));
      }
    });

    proc.on('error', reject);
  });
});

// ========== TIME TRACKING ==========

ipcMain.handle('generate-time-summary', async (event, config) => {
  const processorPath = getProcessorPath('time_tracker');

  // In development, use simplified local processing
  if (!processorPath) {
    try {
      const emails = config.emails || [];
      const period = config.period || 'day';

      // Simple summary without calendar
      const today = new Date();
      const todayStr = today.toISOString().split('T')[0];

      // Filter emails for today/this week
      const filteredEmails = emails.filter(email => {
        const dateStr = email.date_sent || email.date_received;
        if (!dateStr) return false;
        try {
          const emailDate = new Date(dateStr);
          if (period === 'day') {
            return emailDate.toISOString().split('T')[0] === todayStr;
          } else {
            const weekAgo = new Date(today);
            weekAgo.setDate(weekAgo.getDate() - 7);
            return emailDate >= weekAgo;
          }
        } catch (e) {
          return false;
        }
      });

      // Simple categorization
      const byMatter = {};
      filteredEmails.forEach(email => {
        const subject = email.subject || '';
        // Simple extraction - look for [brackets] or first few words
        let matter = 'General';
        const bracketMatch = subject.match(/\[([^\]]+)\]/);
        if (bracketMatch) {
          matter = bracketMatch[1];
        } else if (subject.startsWith('Re:') || subject.startsWith('RE:')) {
          matter = subject.substring(4, 30).trim() || 'General';
        }

        if (!byMatter[matter]) {
          byMatter[matter] = { emails: 0, minutes: 0 };
        }
        byMatter[matter].emails++;
        byMatter[matter].minutes += 3; // Assume 3 min per email
      });

      const totalMinutes = filteredEmails.length * 3;
      const mattersArray = Object.entries(byMatter).map(([name, data]) => ({
        name,
        hours: Math.round(data.minutes / 60 * 10) / 10,
        percent: Math.round(data.minutes / Math.max(totalMinutes, 1) * 100),
        emails_sent: Math.floor(data.emails / 2),
        emails_received: Math.ceil(data.emails / 2),
        meetings: 0
      })).sort((a, b) => b.hours - a.hours);

      return {
        success: true,
        summary: {
          period,
          date: todayStr,
          total_active_hours: Math.round(totalMinutes / 60 * 10) / 10,
          total_meetings: 0,
          total_emails: filteredEmails.length,
          by_matter: mattersArray,
          timeline: []
        }
      };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  // Production: Use Python processor
  const configPath = path.join(app.getPath('temp'), `timetrack_config_${Date.now()}.json`);
  fs.writeFileSync(configPath, JSON.stringify(config));

  return new Promise((resolve, reject) => {
    const proc = spawn(processorPath, [configPath]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('timetrack-progress', msg);
          } else if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            reject(new Error(msg.message));
          }
        } catch (e) {}
      }
    });

    proc.on('close', (code) => {
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Time tracking failed with code ' + code));
      }
    });

    proc.on('error', reject);
  });
});

// ========== API KEY HANDLERS ==========

// Get stored API key
ipcMain.handle('get-api-key', async () => {
  const key = getApiKey();
  return { success: true, hasKey: !!key };
});

// Set API key (stores encrypted)
ipcMain.handle('set-api-key', async (event, apiKey) => {
  if (!apiKey || typeof apiKey !== 'string') {
    return { success: false, error: 'Invalid API key' };
  }
  const result = setApiKey(apiKey);
  return { success: result, error: result ? null : 'Encryption not available' };
});

// Delete stored API key
ipcMain.handle('delete-api-key', async () => {
  const result = deleteApiKey();
  return { success: result };
});

// Test API key by making a minimal API call
ipcMain.handle('test-api-key', async (event, apiKey) => {
  try {
    const https = require('https');

    return new Promise((resolve) => {
      const data = JSON.stringify({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 10,
        messages: [{ role: 'user', content: 'Hi' }]
      });

      const options = {
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(data)
        }
      };

      const req = https.request(options, (res) => {
        let body = '';
        res.on('data', chunk => body += chunk);
        res.on('end', () => {
          if (res.statusCode === 200) {
            resolve({ success: true, message: 'API key is valid' });
          } else if (res.statusCode === 401) {
            resolve({ success: false, error: 'Invalid API key' });
          } else {
            resolve({ success: false, error: `API error: ${res.statusCode}` });
          }
        });
      });

      req.on('error', (e) => {
        resolve({ success: false, error: `Network error: ${e.message}` });
      });

      req.setTimeout(10000, () => {
        req.destroy();
        resolve({ success: false, error: 'Request timed out' });
      });

      req.write(data);
      req.end();
    });
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get API key for use (returns actual key, only for internal use)
ipcMain.handle('get-api-key-value', async () => {
  return getApiKey();
});

// ========== USAGE HISTORY HANDLERS ==========

// Log usage event
ipcMain.handle('log-usage', async (event, data) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    db.run(`
      INSERT INTO usage_history (user_name, feature, action, input_count, output_count, duration_ms)
      VALUES (?, ?, ?, ?, ?, ?)
    `, [
      data.user_name || 'Guest',
      data.feature || 'unknown',
      data.action || 'process',
      data.input_count || 0,
      data.output_count || 0,
      data.duration_ms || 0
    ]);
    saveDatabase();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get recent usage history
ipcMain.handle('get-history', async (event, limit = 20) => {
  if (!db) return [];

  try {
    const results = db.exec(`
      SELECT id, timestamp, user_name, feature, action, input_count, output_count, duration_ms
      FROM usage_history
      ORDER BY timestamp DESC
      LIMIT ?
    `, [limit]);

    if (results.length === 0) return [];

    const columns = results[0].columns;
    return results[0].values.map(row => {
      const obj = {};
      columns.forEach((col, i) => obj[col] = row[i]);
      return obj;
    });
  } catch (e) {
    console.error('Error getting history:', e);
    return [];
  }
});

// Get usage statistics
ipcMain.handle('get-usage-stats', async () => {
  if (!db) return {};

  try {
    // Total counts by feature
    const featureStats = db.exec(`
      SELECT feature, COUNT(*) as count, SUM(output_count) as total_output
      FROM usage_history
      GROUP BY feature
    `);

    // Recent activity (last 7 days)
    const recentStats = db.exec(`
      SELECT DATE(timestamp) as date, COUNT(*) as count
      FROM usage_history
      WHERE timestamp >= datetime('now', '-7 days')
      GROUP BY DATE(timestamp)
      ORDER BY date DESC
    `);

    // Total lifetime stats
    const totals = db.exec(`
      SELECT COUNT(*) as total_operations, SUM(output_count) as total_outputs
      FROM usage_history
    `);

    const stats = {
      by_feature: {},
      by_date: {},
      total_operations: 0,
      total_outputs: 0
    };

    if (featureStats.length > 0) {
      featureStats[0].values.forEach(row => {
        stats.by_feature[row[0]] = { count: row[1], outputs: row[2] || 0 };
      });
    }

    if (recentStats.length > 0) {
      recentStats[0].values.forEach(row => {
        stats.by_date[row[0]] = row[1];
      });
    }

    if (totals.length > 0 && totals[0].values.length > 0) {
      stats.total_operations = totals[0].values[0][0] || 0;
      stats.total_outputs = totals[0].values[0][1] || 0;
    }

    return stats;
  } catch (e) {
    console.error('Error getting stats:', e);
    return {};
  }
});

// Export history to CSV
ipcMain.handle('export-history-csv', async () => {
  if (!db) return null;

  try {
    const results = db.exec(`
      SELECT timestamp, user_name, feature, action, input_count, output_count, duration_ms
      FROM usage_history
      ORDER BY timestamp DESC
    `);

    if (results.length === 0) return null;

    const headers = results[0].columns.join(',');
    const rows = results[0].values.map(row => row.join(',')).join('\n');
    const csv = headers + '\n' + rows;

    const { filePath } = await dialog.showSaveDialog(mainWindow, {
      defaultPath: `emmaneigh-history-${new Date().toISOString().split('T')[0]}.csv`,
      filters: [{ name: 'CSV', extensions: ['csv'] }]
    });

    if (filePath) {
      fs.writeFileSync(filePath, csv);
      return filePath;
    }
    return null;
  } catch (e) {
    console.error('Error exporting history:', e);
    return null;
  }
});

// Clear history (for privacy)
ipcMain.handle('clear-history', async () => {
  if (!db) return { success: false };

  try {
    db.run('DELETE FROM usage_history');
    saveDatabase();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Generic file picker
ipcMain.handle('select-file', async (event, options) => {
  const filters = options?.filters || [{ name: 'All Files', extensions: ['*'] }];
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: filters
  });
  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }
  return result.filePaths[0];
});

// Open file with default application
ipcMain.handle('open-file', async (event, filePath) => {
  try {
    await shell.openPath(filePath);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Update Checklist - parse emails and update checklist status
ipcMain.handle('update-checklist', async (event, config) => {
  const { checklistPath, emailPath } = config;

  // Create temp output folder
  const outputFolder = path.join(app.getPath('temp'), 'emmaneigh_checklist_' + Date.now());
  fs.mkdirSync(outputFolder, { recursive: true });

  // Get processor path using existing helper
  const processorPath = getProcessorPath('checklist_updater');

  if (!processorPath) {
    return { success: false, error: 'Checklist updater not available' };
  }

  if (app.isPackaged && !fs.existsSync(processorPath)) {
    return { success: false, error: 'Processor not found: ' + processorPath };
  }

  // Set executable permission on Mac/Linux
  if (process.platform !== 'win32' && app.isPackaged) {
    try { fs.chmodSync(processorPath, '755'); } catch (e) {}
  }

  return new Promise((resolve) => {
    let args;
    if (app.isPackaged) {
      args = [checklistPath, emailPath, outputFolder];
    } else {
      args = [path.join(__dirname, 'python', 'checklist_updater.py'), checklistPath, emailPath, outputFolder];
    }

    mainWindow.webContents.send('checklist-progress', { message: 'Analyzing emails...', percent: 20 });

    const proc = spawn(processorPath, args);
    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    proc.stderr.on('data', (data) => {
      stderr += data.toString();
      console.error('Checklist updater stderr:', data.toString());
    });

    proc.on('close', (code) => {
      mainWindow.webContents.send('checklist-progress', { message: 'Complete', percent: 100 });

      if (code !== 0) {
        resolve({ success: false, error: stderr || 'Process exited with code ' + code });
        return;
      }

      try {
        const result = JSON.parse(stdout);
        resolve({
          success: result.success,
          outputPath: result.output_path,
          itemsUpdated: result.items_updated,
          error: result.error
        });
      } catch (e) {
        resolve({ success: false, error: 'Failed to parse result: ' + e.message });
      }
    });

    proc.on('error', (err) => {
      resolve({ success: false, error: err.message });
    });
  });
});

// Generate Punchlist - extract open items from checklist
ipcMain.handle('generate-punchlist', async (event, config) => {
  const { checklistPath, statusFilters } = config;

  // Create temp output folder
  const outputFolder = path.join(app.getPath('temp'), 'emmaneigh_punchlist_' + Date.now());
  fs.mkdirSync(outputFolder, { recursive: true });

  // Get processor path using existing helper
  const processorPath = getProcessorPath('punchlist_generator');

  if (!processorPath) {
    return { success: false, error: 'Punchlist generator not available' };
  }

  if (app.isPackaged && !fs.existsSync(processorPath)) {
    return { success: false, error: 'Processor not found: ' + processorPath };
  }

  // Set executable permission on Mac/Linux
  if (process.platform !== 'win32' && app.isPackaged) {
    try { fs.chmodSync(processorPath, '755'); } catch (e) {}
  }

  return new Promise((resolve) => {
    let args;
    const filtersJson = JSON.stringify(statusFilters || ['pending', 'review', 'signature']);

    if (app.isPackaged) {
      args = [checklistPath, outputFolder, filtersJson];
    } else {
      args = [path.join(__dirname, 'python', 'punchlist_generator.py'), checklistPath, outputFolder, filtersJson];
    }

    mainWindow.webContents.send('punchlist-progress', { message: 'Analyzing checklist...', percent: 30 });

    const proc = spawn(processorPath, args);
    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    proc.stderr.on('data', (data) => {
      stderr += data.toString();
      console.error('Punchlist generator stderr:', data.toString());
    });

    proc.on('close', (code) => {
      mainWindow.webContents.send('punchlist-progress', { message: 'Complete', percent: 100 });

      if (code !== 0) {
        resolve({ success: false, error: stderr || 'Process exited with code ' + code });
        return;
      }

      try {
        const result = JSON.parse(stdout);
        resolve({
          success: result.success,
          outputPath: result.output_path,
          itemCount: result.item_count,
          categories: result.categories,
          error: result.error
        });
      } catch (e) {
        resolve({ success: false, error: 'Failed to parse result: ' + e.message });
      }
    });

    proc.on('error', (err) => {
      resolve({ success: false, error: err.message });
    });
  });
});

// Save database on app quit
app.on('before-quit', () => {
  saveDatabase();
});
