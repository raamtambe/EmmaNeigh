const { app, BrowserWindow, ipcMain, dialog, shell, safeStorage } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const archiver = require('archiver');
const crypto = require('crypto');
const nodemailer = require('nodemailer');

// Password hashing functions
function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.pbkdf2Sync(password, salt, 100000, 64, 'sha512').toString('hex');
  return `${salt}:${hash}`;
}

function verifyPassword(password, storedHash) {
  const [salt, hash] = storedHash.split(':');
  const verifyHash = crypto.pbkdf2Sync(password, salt, 100000, 64, 'sha512').toString('hex');
  return hash === verifyHash;
}

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
      CREATE TABLE IF NOT EXISTS users (
        id TEXT PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        email TEXT,
        password_hash TEXT NOT NULL,
        security_question TEXT,
        security_answer_hash TEXT,
        display_name TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        last_login DATETIME,
        api_key_encrypted TEXT,
        telemetry_enabled INTEGER DEFAULT 0
      )
    `);

    // Migrate existing users table if it doesn't have new columns
    try {
      db.run(`ALTER TABLE users ADD COLUMN email TEXT`);
    } catch (e) { /* Column may already exist */ }
    try {
      db.run(`ALTER TABLE users ADD COLUMN security_question TEXT`);
    } catch (e) { /* Column may already exist */ }
    try {
      db.run(`ALTER TABLE users ADD COLUMN security_answer_hash TEXT`);
    } catch (e) { /* Column may already exist */ }
    // 2FA columns
    try {
      db.run(`ALTER TABLE users ADD COLUMN email_verified INTEGER DEFAULT 0`);
    } catch (e) { /* Column may already exist */ }
    try {
      db.run(`ALTER TABLE users ADD COLUMN two_factor_enabled INTEGER DEFAULT 0`);
    } catch (e) { /* Column may already exist */ }

    // Create verification codes table for 2FA
    db.run(`
      CREATE TABLE IF NOT EXISTS verification_codes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT NOT NULL,
        code TEXT NOT NULL,
        type TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        expires_at DATETIME NOT NULL,
        used INTEGER DEFAULT 0,
        FOREIGN KEY (user_id) REFERENCES users(id)
      )
    `);

    // SMTP settings for 2FA email
    db.run(`
      CREATE TABLE IF NOT EXISTS app_settings (
        key TEXT PRIMARY KEY,
        value TEXT
      )
    `);

    db.run(`
      CREATE TABLE IF NOT EXISTS usage_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
        user_id TEXT,
        feature TEXT,
        action TEXT,
        input_count INTEGER DEFAULT 0,
        output_count INTEGER DEFAULT 0,
        duration_ms INTEGER DEFAULT 0,
        FOREIGN KEY (user_id) REFERENCES users(id)
      )
    `);

    saveDatabase();
    console.log('Database initialized at:', historyDbPath);

    // Set Claude model env var for Python processors
    initClaudeModel();
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

// ========== VERSION ENFORCEMENT ==========
// Configuration for mandatory updates with grace period
const VERSION_ENFORCEMENT = {
  // Set this to a version string to require users to update
  // null = no enforcement, users can stay on any version
  minimumVersion: null,  // e.g., '5.2.0' when you want to enforce

  // Grace period in days - how long users can delay the update
  gracePeriodDays: 7,

  // Message shown to users during grace period
  graceMessage: 'A required update is available. Please update within {days} days to continue using EmmaNeigh.',

  // Message shown when grace period expires
  blockedMessage: 'This version of EmmaNeigh is no longer supported. Please update to continue.',

  // Features to disable during grace period (empty = all features work)
  // Options: 'packets', 'execution', 'sigblocks', 'collate', 'email', 'time'
  disabledFeaturesDuringGrace: [],

  // URL to check for version requirements (optional - for remote control)
  // If set, app will fetch this URL to get minimumVersion dynamically
  remoteConfigUrl: null  // e.g., 'https://raw.githubusercontent.com/raamtambe/EmmaNeigh/main/version-config.json'
};

// Store for version enforcement state
const versionEnforcementPath = path.join(app.getPath('userData'), 'version_enforcement.json');

function getVersionEnforcementState() {
  try {
    if (fs.existsSync(versionEnforcementPath)) {
      return JSON.parse(fs.readFileSync(versionEnforcementPath, 'utf8'));
    }
  } catch (e) {}
  return { firstWarningDate: null, updateDismissed: false };
}

function saveVersionEnforcementState(state) {
  try {
    fs.writeFileSync(versionEnforcementPath, JSON.stringify(state));
  } catch (e) {
    console.error('Failed to save version enforcement state:', e);
  }
}

function compareVersions(v1, v2) {
  // Compare semantic versions: returns -1 if v1 < v2, 0 if equal, 1 if v1 > v2
  const parts1 = v1.split('.').map(Number);
  const parts2 = v2.split('.').map(Number);

  for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
    const p1 = parts1[i] || 0;
    const p2 = parts2[i] || 0;
    if (p1 < p2) return -1;
    if (p1 > p2) return 1;
  }
  return 0;
}

async function checkVersionEnforcement() {
  const currentVersion = app.getVersion();
  let minimumVersion = VERSION_ENFORCEMENT.minimumVersion;

  // Optionally fetch remote config for dynamic control
  if (VERSION_ENFORCEMENT.remoteConfigUrl) {
    try {
      const https = require('https');
      const response = await new Promise((resolve, reject) => {
        https.get(VERSION_ENFORCEMENT.remoteConfigUrl, (res) => {
          let data = '';
          res.on('data', chunk => data += chunk);
          res.on('end', () => resolve(data));
        }).on('error', reject);
      });
      const remoteConfig = JSON.parse(response);
      if (remoteConfig.minimumVersion) {
        minimumVersion = remoteConfig.minimumVersion;
      }
    } catch (e) {
      console.log('Could not fetch remote version config:', e.message);
    }
  }

  // No enforcement if minimumVersion is not set
  if (!minimumVersion) {
    return { enforced: false, status: 'none' };
  }

  // Check if current version meets minimum
  if (compareVersions(currentVersion, minimumVersion) >= 0) {
    // Version is OK, clear any previous enforcement state
    saveVersionEnforcementState({ firstWarningDate: null, updateDismissed: false });
    return { enforced: false, status: 'ok' };
  }

  // Version is below minimum - check grace period
  const state = getVersionEnforcementState();
  const now = new Date();

  if (!state.firstWarningDate) {
    // First time seeing this - start grace period
    state.firstWarningDate = now.toISOString();
    saveVersionEnforcementState(state);
  }

  const firstWarning = new Date(state.firstWarningDate);
  const daysSinceWarning = Math.floor((now - firstWarning) / (1000 * 60 * 60 * 24));
  const daysRemaining = Math.max(0, VERSION_ENFORCEMENT.gracePeriodDays - daysSinceWarning);

  if (daysRemaining > 0) {
    // Still in grace period
    return {
      enforced: true,
      status: 'grace',
      daysRemaining,
      message: VERSION_ENFORCEMENT.graceMessage.replace('{days}', daysRemaining),
      minimumVersion,
      currentVersion,
      disabledFeatures: VERSION_ENFORCEMENT.disabledFeaturesDuringGrace
    };
  } else {
    // Grace period expired - block the app
    return {
      enforced: true,
      status: 'blocked',
      daysRemaining: 0,
      message: VERSION_ENFORCEMENT.blockedMessage,
      minimumVersion,
      currentVersion,
      disabledFeatures: ['all']  // Block everything
    };
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

    // Explicitly set feed URL for reliable updates
    autoUpdater.setFeedURL({
      provider: 'github',
      owner: 'raamtambe',
      repo: 'EmmaNeigh'
    });

    // Log for debugging
    autoUpdater.on('checking-for-update', () => {
      console.log('Checking for updates...');
    });

    autoUpdater.checkForUpdates().catch(err => {
      console.error('Failed to check for updates:', err);
    });

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

    autoUpdater.on('update-not-available', (info) => {
      if (mainWindow) {
        mainWindow.webContents.send('update-not-available', info);
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

// Manual check for updates
ipcMain.handle('check-for-updates', async () => {
  if (!autoUpdater) {
    return { success: false, error: 'Auto-updater not available (development mode)' };
  }

  try {
    await autoUpdater.checkForUpdates();
    return { success: true, checking: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get app version
ipcMain.handle('get-app-version', () => {
  return app.getVersion();
});

// Check version enforcement status
ipcMain.handle('check-version-enforcement', async () => {
  return await checkVersionEnforcement();
});

// Dismiss update warning (user acknowledges but continues)
ipcMain.handle('dismiss-update-warning', async () => {
  const state = getVersionEnforcementState();
  state.updateDismissed = true;
  saveVersionEnforcementState(state);
  return { success: true };
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

  // Production - unified processor executable
  // All modules are bundled into a single emna_processor binary
  const resourcesPath = process.resourcesPath;
  if (process.platform === 'win32') {
    return path.join(resourcesPath, 'processor', 'emna_processor.exe');
  } else {
    return path.join(resourcesPath, 'processor', 'emna_processor');
  }
}


// Process signature packets
// Accepts either a folder path string (legacy) or an object { folder: string } or { files: string[] }
ipcMain.handle('process-folder', async (event, input) => {
  return new Promise((resolve, reject) => {
    // Send initial progress
    mainWindow.webContents.send('progress', { percent: 0, message: 'Initializing signature packet processor...' });

    const moduleName = 'signature_packets';
    const processorPath = getProcessorPath(moduleName);

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
      // If we have output_format, use config file
      if (input.output_format) {
        const configPath = path.join(app.getPath('temp'), `packets-config-${Date.now()}.json`);
        fs.writeFileSync(configPath, JSON.stringify({
          folder: input.folder,
          output_format: input.output_format
        }));
        args = ['--config', configPath];
      } else {
        args = [input.folder];
      }
    } else if (input.files) {
      // Write config to temp file for multi-file input
      const configPath = path.join(app.getPath('temp'), `packets-config-${Date.now()}.json`);
      fs.writeFileSync(configPath, JSON.stringify({
        files: input.files,
        output_format: input.output_format || 'preserve'
      }));
      args = ['--config', configPath];
    } else {
      reject(new Error('Invalid input: must be folder path or { folder } or { files }'));
      return;
    }

    const proc = spawn(processorPath, [moduleName, ...args]);
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

// Create ZIP from specific files (not a whole directory)
ipcMain.handle('create-zip-files', async (event, filePaths) => {
  const zipPath = path.join(app.getPath('temp'), `EmmaNeigh-Output-${Date.now()}.zip`);

  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => resolve(zipPath));
    archive.on('error', reject);

    archive.pipe(output);
    for (const fp of filePaths) {
      if (fs.existsSync(fp)) {
        archive.file(fp, { name: path.basename(fp) });
      }
    }
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

// Generate packet shell - combined signature packet with all pages
ipcMain.handle('generate-packet-shell', async (event, input) => {
  return new Promise((resolve, reject) => {
    const moduleName = 'packet_shell_generator';
    const processorPath = getProcessorPath(moduleName);

    if (!processorPath) {
      reject(new Error('Development mode - please build the app first'));
      return;
    }

    if (!fs.existsSync(processorPath)) {
      reject(new Error('Packet shell generator not found: ' + processorPath));
      return;
    }

    // Make executable on Mac
    if (process.platform === 'darwin') {
      try { fs.chmodSync(processorPath, '755'); } catch (e) {}
    }

    // Write config to temp file
    const configPath = path.join(app.getPath('temp'), `shell-config-${Date.now()}.json`);
    const outputDir = path.join(app.getPath('temp'), `packet-shell-${Date.now()}`);

    fs.writeFileSync(configPath, JSON.stringify({
      files: input.files,
      output_format: input.output_format || 'both',
      output_dir: outputDir
    }));

    const args = ['--config', configPath];
    const proc = spawn(processorPath, [moduleName, ...args]);
    let result = null;

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (msg.type === 'progress') {
            mainWindow.webContents.send('shell-progress', msg);
          } else if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            reject(new Error(msg.message));
          }
        } catch (e) {}
      }
    });

    proc.stderr.on('data', (data) => {
      console.error('shell stderr:', data.toString());
    });

    proc.on('close', (code) => {
      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Packet shell generation failed with code ' + code));
      }
    });

    proc.on('error', (err) => {
      reject(err);
    });
  });
});

// Select folder dialog
ipcMain.handle('select-folder', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });
  return filePaths[0] || null;
});

// Select multiple documents (PDF or DOCX)
ipcMain.handle('select-documents-multiple', async () => {
  const { filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Documents', extensions: ['pdf', 'docx'] }]
  });
  return filePaths || [];
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
    const moduleName = 'checklist_parser';
    const processorPath = getProcessorPath(moduleName);

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

    const proc = spawn(processorPath, [moduleName, checklistPath]);
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
    const moduleName = 'incumbency_parser';
    const processorPath = getProcessorPath(moduleName);

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

    const proc = spawn(processorPath, [moduleName, incPath]);
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
    const moduleName = 'sigblock_workflow';
    const processorPath = getProcessorPath(moduleName);

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

    const proc = spawn(processorPath, [moduleName, configPath]);
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
// signedPath can be a folder path (new two-folder workflow) or a single PDF file (legacy)
ipcMain.handle('process-execution-version', async (event, originalsInput, signedPath) => {
  return new Promise((resolve, reject) => {
    const moduleName = 'execution_version';
    const processorPath = getProcessorPath(moduleName);

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

    // Determine signed path type (folder vs file)
    const signedIsFolder = fs.existsSync(signedPath) && fs.statSync(signedPath).isDirectory();

    // Handle different input types
    if (typeof originalsInput === 'string') {
      // Legacy: folder path string - pass directly
      args = [originalsInput, signedPath];
    } else if (originalsInput.folder) {
      // Object with folder property
      args = [originalsInput.folder, signedPath];
    } else if (originalsInput.files) {
      // List of individual files - need config file
      configPath = path.join(app.getPath('temp'), `exec-config-${Date.now()}.json`);
      const config = {
        files: originalsInput.files
      };
      // Use signed_folder or signed_pdf based on path type
      if (signedIsFolder) {
        config.signed_folder = signedPath;
      } else {
        config.signed_pdf = signedPath;
      }
      fs.writeFileSync(configPath, JSON.stringify(config));
      args = ['--config', configPath];
    } else {
      reject(new Error('Invalid input: must be folder path or { folder } or { files }'));
      return;
    }

    const proc = spawn(processorPath, [moduleName, ...args]);
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
    const moduleName = 'document_collator';
    const processorPath = getProcessorPath(moduleName);

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

    const proc = spawn(processorPath, [moduleName, configPath]);
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
    const moduleName = 'document_redline';
    const processorPath = getProcessorPath(moduleName);

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

    const proc = spawn(processorPath, [moduleName, configPath]);
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
  const emailModuleName = 'email_csv_parser';
  const processorPath = getProcessorPath(emailModuleName);

  if (!processorPath) {
    // In development, parse CSV directly with Node.js
    try {
      // Read with BOM stripping
      let csvContent = fs.readFileSync(csvPath, 'utf-8');
      // Strip BOM (UTF-8 BOM: \uFEFF, also handle other BOMs)
      if (csvContent.charCodeAt(0) === 0xFEFF) {
        csvContent = csvContent.slice(1);
      }

      // Proper CSV parsing that handles quoted fields and multiline values
      function parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          if (char === '"') {
            if (inQuotes && line[i + 1] === '"') {
              current += '"';
              i++;
            } else {
              inQuotes = !inQuotes;
            }
          } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
          } else {
            current += char;
          }
        }
        result.push(current.trim());
        return result;
      }

      // Split into lines but handle multiline quoted fields
      function splitCSVLines(content) {
        const lines = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < content.length; i++) {
          const char = content[i];
          if (char === '"') {
            inQuotes = !inQuotes;
            current += char;
          } else if ((char === '\n' || char === '\r') && !inQuotes) {
            if (char === '\r' && content[i + 1] === '\n') i++;
            if (current.trim()) lines.push(current);
            current = '';
          } else {
            current += char;
          }
        }
        if (current.trim()) lines.push(current);
        return lines;
      }

      const lines = splitCSVLines(csvContent);
      // Clean and normalize headers - strip BOM, quotes, whitespace
      const headers = parseCSVLine(lines[0]).map(h =>
        h.replace(/^\uFEFF/, '').replace(/^["']|["']$/g, '').trim().toLowerCase()
      );

      const emails = [];
      for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        const email = {
          subject: '',
          from: '',
          to: '',
          cc: '',
          date_sent: null,
          date_received: null,
          body: '',
          attachments: '',
          has_attachments: false
        };

        headers.forEach((header, idx) => {
          const value = values[idx] ? values[idx].trim() : '';
          if (header.includes('subject') || header === 'title') email.subject = value;
          else if (header.includes('from') || header === 'sender') email.from = value;
          else if (header.includes('to') || header === 'recipient') email.to = value;
          else if (header === 'cc' || header.includes('carbon')) email.cc = value;
          else if (header.includes('body') || header.includes('content') || header === 'message') email.body = value;
          else if (header.includes('sent') || header === 'send date') email.date_sent = value;
          else if (header.includes('received') || header === 'date' || header === 'receive date') email.date_received = value;
          else if (header === 'has attachments') {
            // Boolean column from Outlook - just TRUE/FALSE
            email.has_attachments = !!(value && value.toLowerCase() !== 'no' && value.toLowerCase() !== 'false' && value !== '0');
          } else if (header.includes('attachment')) {
            // Actual attachment filename(s)
            email.attachments = value;
            email.has_attachments = !!(value && value.toLowerCase() !== 'no' && value.toLowerCase() !== 'false' && value !== '0');
          }
        });

        emails.push(email);
      }

      const uniqueSenders = new Set(emails.map(e => e.from).filter(f => f));
      const dates = emails.map(e => e.date_sent || e.date_received).filter(d => d).sort();
      const withAttachments = emails.filter(e => e.has_attachments).length;

      return {
        success: true,
        emails: emails,
        summary: {
          total_emails: emails.length,
          unique_senders: uniqueSenders.size,
          with_attachments: withAttachments,
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
    const proc = spawn(processorPath, [emailModuleName, configPath]);
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
  const nlModuleName = 'email_nl_search';
  const processorPath = getProcessorPath(nlModuleName);

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
        model: getClaudeModel(),
        max_tokens: 1024,
        messages: [{ role: 'user', content: prompt }]
      });

      return new Promise((resolve) => {
        const options = {
          hostname: 'api.anthropic.com',
          path: '/v1/messages',
          method: 'POST',
          rejectUnauthorized: false,  // Allow self-signed certs (corporate proxies)
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
      args = [nlModuleName, configPath];
    } else {
      args = [nlModuleName, path.join(__dirname, 'python', 'email_nl_search.py'), configPath];
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
  const timeModuleName = 'time_tracker';
  const processorPath = getProcessorPath(timeModuleName);

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
    const proc = spawn(processorPath, [timeModuleName, configPath]);
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
        model: getClaudeModel(),
        max_tokens: 10,
        messages: [{ role: 'user', content: 'Hi' }]
      });

      const options = {
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        rejectUnauthorized: false,  // Allow self-signed certs (corporate proxies)
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
            // Try to extract meaningful error from response
            let errMsg = `API error: ${res.statusCode}`;
            try {
              const parsed = JSON.parse(body);
              if (parsed.error && parsed.error.message) {
                errMsg = `API error (${res.statusCode}): ${parsed.error.message}`;
              }
            } catch(e) {}
            resolve({ success: false, error: errMsg });
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

// ========== USER ACCOUNT HANDLERS ==========

// Create new user account
ipcMain.handle('create-user', async (event, { username, password, displayName, email, securityQuestion, securityAnswer }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username || !password) {
    return { success: false, error: 'Username and password are required' };
  }

  if (!email || !email.includes('@')) {
    return { success: false, error: 'A valid email address is required for two-factor authentication' };
  }

  if (password.length < 4) {
    return { success: false, error: 'Password must be at least 4 characters' };
  }

  try {
    const id = crypto.randomUUID();
    const passwordHash = hashPassword(password);

    // Hash security answer if provided
    let securityAnswerHash = null;
    if (securityQuestion && securityAnswer) {
      securityAnswerHash = hashPassword(securityAnswer.toLowerCase().trim());
    }

    db.run(`INSERT INTO users (id, username, email, password_hash, security_question, security_answer_hash, display_name, two_factor_enabled) VALUES (?, ?, ?, ?, ?, ?, ?, 1)`,
      [id, username.toLowerCase().trim(), email.trim(), passwordHash, securityQuestion || null, securityAnswerHash, displayName || username]);
    saveDatabase();

    return { success: true, userId: id };
  } catch (e) {
    if (e.message && e.message.includes('UNIQUE')) {
      return { success: false, error: 'Username already exists' };
    }
    return { success: false, error: e.message };
  }
});

// Login with username/password (step 1 of login - may require 2FA)
ipcMain.handle('login-user', async (event, { username, password }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username || !password) {
    return { success: false, error: 'Username and password are required' };
  }

  try {
    const result = db.exec(`SELECT id, username, password_hash, display_name, api_key_encrypted, email, two_factor_enabled FROM users WHERE username = '${username.toLowerCase().trim()}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const row = result[0].values[0];
    const [id, uname, passwordHash, displayName, apiKeyEnc, email, twoFactorEnabled] = row;

    if (!verifyPassword(password, passwordHash)) {
      return { success: false, error: 'Invalid password' };
    }

    // 2FA is mandatory for all users with an email address
    if (email) {
      // Generate and send verification code
      const code = Math.floor(100000 + Math.random() * 900000).toString();
      const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString();

      // Invalidate existing codes
      db.run(`UPDATE verification_codes SET used = 1 WHERE user_id = '${id}' AND type = 'login' AND used = 0`);

      // Store new code
      db.run(`INSERT INTO verification_codes (user_id, code, type, expires_at) VALUES (?, ?, ?, ?)`,
        [id, code, 'login', expiresAt]);
      saveDatabase();

      // Send verification email
      const emailResult = await sendVerificationEmail(email, code, 'login');

      // Mask email for display
      const maskedEmail = email.replace(/(.{2})(.*)(@.*)/, '$1***$3');

      const response = {
        success: true,
        requires2FA: true,
        userId: id,
        email: maskedEmail
      };

      // If SMTP not configured, include code for development/fallback
      if (!emailResult.sent) {
        response._devCode = code;
        response._smtpNote = emailResult.reason || 'Email not sent';
      }

      return response;
    }

    // No email on account - require email to be set (2FA mandatory)
    return {
      success: false,
      error: 'Email address required for login. Please contact your administrator to add an email to your account.'
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Complete login after 2FA verification
ipcMain.handle('complete-2fa-login', async (event, { userId, code }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!code || code.length !== 6) {
    return { success: false, error: 'Invalid verification code' };
  }

  try {
    // Find valid code
    const codeResult = db.exec(`
      SELECT id, expires_at FROM verification_codes
      WHERE user_id = '${userId}'
      AND code = '${code}'
      AND type = 'login'
      AND used = 0
      ORDER BY created_at DESC
      LIMIT 1
    `);

    if (codeResult.length === 0 || codeResult[0].values.length === 0) {
      return { success: false, error: 'Invalid verification code' };
    }

    const [codeId, expiresAt] = codeResult[0].values[0];

    // Check expiration
    if (new Date(expiresAt) < new Date()) {
      return { success: false, error: 'Verification code has expired' };
    }

    // Mark code as used
    db.run(`UPDATE verification_codes SET used = 1 WHERE id = ${codeId}`);

    // Get user info and complete login
    const userResult = db.exec(`SELECT id, username, display_name, api_key_encrypted FROM users WHERE id = '${userId}'`);

    if (userResult.length === 0 || userResult[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const [id, username, displayName, apiKeyEnc] = userResult[0].values[0];

    // Update last login
    db.run(`UPDATE users SET last_login = datetime('now') WHERE id = '${id}'`);
    saveDatabase();

    return {
      success: true,
      user: { id, username, displayName: displayName || username, hasApiKey: !!apiKeyEnc }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get user statistics (for admin tracking)
ipcMain.handle('get-user-stats', async () => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    // Total registered users
    const totalResult = db.exec(`SELECT COUNT(*) FROM users`);
    const totalUsers = totalResult.length > 0 ? totalResult[0].values[0][0] : 0;

    // Users active in last 30 days
    const activeResult = db.exec(`SELECT COUNT(*) FROM users WHERE last_login >= datetime('now', '-30 days')`);
    const activeUsers = activeResult.length > 0 ? activeResult[0].values[0][0] : 0;

    // Users active in last 7 days
    const weeklyResult = db.exec(`SELECT COUNT(*) FROM users WHERE last_login >= datetime('now', '-7 days')`);
    const weeklyActive = weeklyResult.length > 0 ? weeklyResult[0].values[0][0] : 0;

    // User list with last login (no sensitive data)
    const usersResult = db.exec(`SELECT username, display_name, email, last_login, created_at FROM users ORDER BY last_login DESC`);
    const users = usersResult.length > 0 ? usersResult[0].values.map(row => ({
      username: row[0],
      displayName: row[1],
      email: row[2] ? row[2].replace(/(.{2})(.*)(@.*)/, '$1***$3') : null,
      lastLogin: row[3],
      createdAt: row[4]
    })) : [];

    return {
      success: true,
      totalUsers,
      activeUsers,
      weeklyActive,
      users
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Save API key for logged-in user
ipcMain.handle('set-user-api-key', async (event, { userId, apiKey }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    let encrypted = null;
    if (apiKey && safeStorage.isEncryptionAvailable()) {
      const encBuffer = safeStorage.encryptString(apiKey);
      encrypted = encBuffer.toString('base64');
    } else if (apiKey) {
      // Fallback: store as plain text (not ideal but functional)
      encrypted = apiKey;
    }

    db.run(`UPDATE users SET api_key_encrypted = '${encrypted || ''}' WHERE id = '${userId}'`);
    saveDatabase();

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get API key for logged-in user
ipcMain.handle('get-user-api-key', async (event, { userId }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    const result = db.exec(`SELECT api_key_encrypted FROM users WHERE id = '${userId}'`);

    if (result.length === 0 || !result[0].values[0] || !result[0].values[0][0]) {
      return { success: true, apiKey: null };
    }

    const encrypted = result[0].values[0][0];

    // Try to decrypt
    if (safeStorage.isEncryptionAvailable() && encrypted.length > 50) {
      try {
        const encBuffer = Buffer.from(encrypted, 'base64');
        const apiKey = safeStorage.decryptString(encBuffer);
        return { success: true, apiKey };
      } catch (decryptErr) {
        // Might be stored as plain text, return as-is
        return { success: true, apiKey: encrypted };
      }
    }

    // Return as-is (plain text fallback)
    return { success: true, apiKey: encrypted };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get user by ID (for session restore)
ipcMain.handle('get-user-by-id', async (event, { userId }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    const result = db.exec(`SELECT id, username, display_name, api_key_encrypted FROM users WHERE id = '${userId}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const row = result[0].values[0];
    const [id, uname, displayName, apiKeyEnc] = row;

    return {
      success: true,
      user: { id, username: uname, displayName: displayName || uname, hasApiKey: !!apiKeyEnc }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get security question for a user (for password reset)
ipcMain.handle('get-security-question', async (event, { username }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username) {
    return { success: false, error: 'Username is required' };
  }

  try {
    const result = db.exec(`SELECT security_question FROM users WHERE username = '${username.toLowerCase().trim()}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const question = result[0].values[0][0];
    if (!question) {
      return { success: false, error: 'No security question set for this account' };
    }

    // Return human-readable question
    const questions = {
      'pet': 'What was the name of your first pet?',
      'city': 'In what city were you born?',
      'school': 'What was the name of your first school?',
      'mother': "What is your mother's maiden name?"
    };

    return { success: true, question: questions[question] || question };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Reset password with security answer
ipcMain.handle('reset-password', async (event, { username, securityAnswer, newPassword }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username || !securityAnswer || !newPassword) {
    return { success: false, error: 'All fields are required' };
  }

  if (newPassword.length < 4) {
    return { success: false, error: 'Password must be at least 4 characters' };
  }

  try {
    const result = db.exec(`SELECT security_answer_hash FROM users WHERE username = '${username.toLowerCase().trim()}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const storedHash = result[0].values[0][0];
    if (!storedHash) {
      return { success: false, error: 'No security question set for this account' };
    }

    if (!verifyPassword(securityAnswer.toLowerCase().trim(), storedHash)) {
      return { success: false, error: 'Incorrect security answer' };
    }

    // Update password
    const newPasswordHash = hashPassword(newPassword);
    db.run(`UPDATE users SET password_hash = '${newPasswordHash}' WHERE username = '${username.toLowerCase().trim()}'`);
    saveDatabase();

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// ========== 2FA / EMAIL VERIFICATION HANDLERS ==========

// Generate a 6-digit verification code
function generateVerificationCode() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

// Get SMTP settings from database
function getSmtpSettings() {
  if (!db) return null;
  try {
    const keys = ['smtp_host', 'smtp_port', 'smtp_user', 'smtp_pass', 'smtp_from'];
    const settings = {};
    for (const key of keys) {
      const result = db.exec(`SELECT value FROM app_settings WHERE key = '${key}'`);
      if (result.length > 0 && result[0].values.length > 0) {
        settings[key] = result[0].values[0][0];
      }
    }
    // Decrypt password if stored
    if (settings.smtp_pass && safeStorage.isEncryptionAvailable()) {
      try {
        settings.smtp_pass = safeStorage.decryptString(Buffer.from(settings.smtp_pass, 'base64'));
      } catch (e) {
        // If decryption fails, try using as plain text
      }
    }
    return settings;
  } catch (e) {
    console.error('Failed to get SMTP settings:', e);
    return null;
  }
}

// Send verification email using nodemailer SMTP
async function sendVerificationEmail(email, code, type) {
  const subject = type === 'login'
    ? 'EmmaNeigh - Login Verification Code'
    : type === 'email_verify'
    ? 'EmmaNeigh - Verify Your Email'
    : 'EmmaNeigh - Verification Code';

  const htmlBody = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 400px; margin: 0 auto; padding: 20px;">
      <h2 style="color: #333; text-align: center;">EmmaNeigh</h2>
      <p style="color: #555; text-align: center;">Your verification code is:</p>
      <div style="background: #f5f5f5; border-radius: 8px; padding: 20px; text-align: center; margin: 20px 0;">
        <span style="font-size: 32px; letter-spacing: 8px; font-weight: bold; color: #333;">${code}</span>
      </div>
      <p style="color: #888; font-size: 13px; text-align: center;">This code expires in 10 minutes.</p>
      <p style="color: #888; font-size: 12px; text-align: center;">If you didn't request this code, please ignore this email.</p>
    </div>`;

  const textBody = `Your EmmaNeigh verification code is: ${code}\n\nThis code expires in 10 minutes.\n\nIf you didn't request this code, please ignore this email.`;

  // Get SMTP settings
  const smtp = getSmtpSettings();

  if (!smtp || !smtp.smtp_host || !smtp.smtp_user || !smtp.smtp_pass) {
    console.log(`[2FA] SMTP not configured. Verification code for ${email}: ${code}`);
    // Return true but with a flag indicating SMTP isn't set up
    return { sent: false, reason: 'smtp_not_configured', code };
  }

  try {
    const transporter = nodemailer.createTransport({
      host: smtp.smtp_host,
      port: parseInt(smtp.smtp_port) || 587,
      secure: parseInt(smtp.smtp_port) === 465,
      auth: {
        user: smtp.smtp_user,
        pass: smtp.smtp_pass
      }
    });

    await transporter.sendMail({
      from: smtp.smtp_from || smtp.smtp_user,
      to: email,
      subject: subject,
      text: textBody,
      html: htmlBody
    });

    console.log(`[2FA] Verification email sent to ${email}`);
    return { sent: true };
  } catch (e) {
    console.error('[2FA] Failed to send email:', e.message);
    return { sent: false, reason: e.message, code };
  }
}

// Request a verification code (for login 2FA or email verification)
ipcMain.handle('request-verification-code', async (event, { userId, email, type }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!userId && !email) {
    return { success: false, error: 'User ID or email required' };
  }

  try {
    let userEmail = email;
    let resolvedUserId = userId;

    // If we have userId, get the user's email
    if (userId && !email) {
      const result = db.exec(`SELECT email FROM users WHERE id = '${userId}'`);
      if (result.length === 0 || result[0].values.length === 0) {
        return { success: false, error: 'User not found' };
      }
      userEmail = result[0].values[0][0];
    }

    // If we have email but no userId, find the user
    if (email && !userId) {
      const result = db.exec(`SELECT id FROM users WHERE email = '${email}'`);
      if (result.length > 0 && result[0].values.length > 0) {
        resolvedUserId = result[0].values[0][0];
      }
    }

    if (!userEmail) {
      return { success: false, error: 'No email address on file. Please add an email to enable 2FA.' };
    }

    // Generate code
    const code = generateVerificationCode();
    const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 minutes

    // Invalidate any existing unused codes for this user/type
    if (resolvedUserId) {
      db.run(`UPDATE verification_codes SET used = 1 WHERE user_id = '${resolvedUserId}' AND type = '${type}' AND used = 0`);
    }

    // Store the code
    db.run(`
      INSERT INTO verification_codes (user_id, code, type, expires_at)
      VALUES (?, ?, ?, ?)
    `, [resolvedUserId || 'pending', code, type, expiresAt]);

    saveDatabase();

    // Send the email
    const emailResult = await sendVerificationEmail(userEmail, code, type);

    // Mask email for display
    const maskedEmail = userEmail.replace(/(.{2})(.*)(@.*)/, '$1***$3');

    const response = {
      success: true,
      message: `Verification code sent to ${maskedEmail}`,
      email: maskedEmail,
      emailSent: emailResult.sent
    };

    // If SMTP is not configured, include the code so user can still verify
    if (!emailResult.sent) {
      response._devCode = code;
      response.message = emailResult.reason === 'smtp_not_configured'
        ? `SMTP not configured. Code: ${code} (Configure SMTP in Settings to send emails)`
        : `Email failed: ${emailResult.reason}. Code: ${code}`;
    }

    return response;
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Verify a code
ipcMain.handle('verify-code', async (event, { userId, code, type }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!code || code.length !== 6) {
    return { success: false, error: 'Invalid verification code format' };
  }

  try {
    // Find valid code
    const result = db.exec(`
      SELECT id, user_id, expires_at FROM verification_codes
      WHERE code = '${code}'
      AND type = '${type}'
      AND used = 0
      AND (user_id = '${userId}' OR user_id = 'pending')
      ORDER BY created_at DESC
      LIMIT 1
    `);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'Invalid or expired verification code' };
    }

    const [codeId, codeUserId, expiresAt] = result[0].values[0];

    // Check expiration
    if (new Date(expiresAt) < new Date()) {
      return { success: false, error: 'Verification code has expired' };
    }

    // Mark code as used
    db.run(`UPDATE verification_codes SET used = 1 WHERE id = ${codeId}`);

    // If this is email verification, mark the email as verified
    if (type === 'email_verify' && userId) {
      db.run(`UPDATE users SET email_verified = 1 WHERE id = '${userId}'`);
    }

    saveDatabase();

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Enable/disable 2FA for a user
ipcMain.handle('toggle-2fa', async (event, { userId, enable }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    // Check if user has verified email
    const result = db.exec(`SELECT email, email_verified FROM users WHERE id = '${userId}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const [email, emailVerified] = result[0].values[0];

    if (enable && !email) {
      return { success: false, error: 'Please add an email address before enabling 2FA' };
    }

    if (enable && !emailVerified) {
      return { success: false, error: 'Please verify your email address before enabling 2FA' };
    }

    db.run(`UPDATE users SET two_factor_enabled = ${enable ? 1 : 0} WHERE id = '${userId}'`);
    saveDatabase();

    return { success: true, twoFactorEnabled: enable };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get user's 2FA status
ipcMain.handle('get-2fa-status', async (event, { userId }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    const result = db.exec(`SELECT email, email_verified, two_factor_enabled FROM users WHERE id = '${userId}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const [email, emailVerified, twoFactorEnabled] = result[0].values[0];

    return {
      success: true,
      email: email || null,
      emailVerified: !!emailVerified,
      twoFactorEnabled: !!twoFactorEnabled
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Update user email
ipcMain.handle('update-user-email', async (event, { userId, email }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!email || !email.includes('@')) {
    return { success: false, error: 'Invalid email address' };
  }

  try {
    // Check if email is already used by another user
    const existing = db.exec(`SELECT id FROM users WHERE email = '${email}' AND id != '${userId}'`);
    if (existing.length > 0 && existing[0].values.length > 0) {
      return { success: false, error: 'Email address is already in use' };
    }

    // Update email and reset verification status
    db.run(`UPDATE users SET email = '${email}', email_verified = 0, two_factor_enabled = 0 WHERE id = '${userId}'`);
    saveDatabase();

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// ========== APP SETTINGS HANDLERS ==========

ipcMain.handle('save-setting', async (event, { key, value }) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    db.run(`INSERT OR REPLACE INTO app_settings (key, value) VALUES ('${key}', '${(value || '').replace(/'/g, "''")}')`);
    saveDatabase();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('get-setting', async (event, key) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    const result = db.exec(`SELECT value FROM app_settings WHERE key = '${key}'`);
    if (result.length > 0 && result[0].values.length > 0) {
      return { success: true, value: result[0].values[0][0] };
    }
    return { success: true, value: null };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get the configured Claude model (or default)
function getClaudeModel() {
  if (!db) return 'claude-sonnet-4-6';
  try {
    const result = db.exec(`SELECT value FROM app_settings WHERE key = 'claude_model'`);
    if (result.length > 0 && result[0].values.length > 0 && result[0].values[0][0]) {
      const model = result[0].values[0][0];
      // Also set env var for Python processors
      process.env.CLAUDE_MODEL = model;
      return model;
    }
  } catch (e) {}
  return 'claude-sonnet-4-20250514';
}

// Initialize model env var on startup (called after DB init)
function initClaudeModel() {
  process.env.CLAUDE_MODEL = getClaudeModel();
}

// ========== SMTP SETTINGS HANDLERS ==========

// Save SMTP settings
ipcMain.handle('save-smtp-settings', async (event, { host, port, user, pass, from }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    // Encrypt password if possible
    let encryptedPass = pass;
    if (pass && safeStorage.isEncryptionAvailable()) {
      encryptedPass = safeStorage.encryptString(pass).toString('base64');
    }

    const settings = { smtp_host: host, smtp_port: port, smtp_user: user, smtp_pass: encryptedPass, smtp_from: from || user };
    for (const [key, value] of Object.entries(settings)) {
      db.run(`INSERT OR REPLACE INTO app_settings (key, value) VALUES ('${key}', '${(value || '').replace(/'/g, "''")}' )`);
    }
    saveDatabase();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get SMTP settings (without password)
ipcMain.handle('get-smtp-settings', async (event) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    const smtp = getSmtpSettings();
    return {
      success: true,
      host: smtp?.smtp_host || '',
      port: smtp?.smtp_port || '587',
      user: smtp?.smtp_user || '',
      from: smtp?.smtp_from || '',
      hasPassword: !!(smtp?.smtp_pass)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Test SMTP connection
ipcMain.handle('test-smtp', async (event, { host, port, user, pass, from }) => {
  try {
    const transporter = nodemailer.createTransport({
      host: host,
      port: parseInt(port) || 587,
      secure: parseInt(port) === 465,
      auth: { user, pass }
    });

    await transporter.verify();
    return { success: true, message: 'SMTP connection successful' };
  } catch (e) {
    return { success: false, error: `SMTP test failed: ${e.message}` };
  }
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
  const clModuleName = 'checklist_updater';
  const processorPath = getProcessorPath(clModuleName);

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
    const apiKey = getApiKey();  // Get API key for LLM-powered matching

    if (app.isPackaged) {
      args = [clModuleName, checklistPath, emailPath, outputFolder];
      if (apiKey) args.push(apiKey);  // Pass API key as 4th argument
    } else {
      args = [clModuleName, path.join(__dirname, 'python', 'checklist_updater.py'), checklistPath, emailPath, outputFolder];
      if (apiKey) args.push(apiKey);  // Pass API key as 4th argument
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
  const plModuleName = 'punchlist_generator';
  const processorPath = getProcessorPath(plModuleName);

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
    const apiKey = getApiKey();  // Get API key for LLM-powered categorization

    if (app.isPackaged) {
      args = [plModuleName, checklistPath, outputFolder, filtersJson];
      if (apiKey) args.push(apiKey);  // Pass API key as 4th argument
    } else {
      args = [plModuleName, path.join(__dirname, 'python', 'punchlist_generator.py'), checklistPath, outputFolder, filtersJson];
      if (apiKey) args.push(apiKey);  // Pass API key as 4th argument
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
