const { app, BrowserWindow, ipcMain, dialog, shell, safeStorage } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn, spawnSync } = require('child_process');
const archiver = require('archiver');
const crypto = require('crypto');
const nodemailer = require('nodemailer');
const os = require('os');
let initFirebase = () => false;
let loadFirebaseConfig = () => null;
let saveFirebaseConfig = () => false;
let logToFirestore = async () => false;
let batchLogToFirestore = async () => 0;

try {
  ({ initFirebase, loadFirebaseConfig, saveFirebaseConfig, logToFirestore, batchLogToFirestore } = require('./firebase-config'));
} catch (e) {
  console.error('Firebase module unavailable. Centralized logging will be disabled:', e.message);
}

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
const pendingLogsPath = path.join(app.getPath('userData'), 'pending_logs.json');
const APP_VERSION = require('./package.json').version;
const MACHINE_ID = crypto.createHash('sha256').update(os.hostname()).digest('hex').substring(0, 12);
const DEFAULT_AI_PROVIDER = 'ollama';
const DEFAULT_CLAUDE_MODEL = 'claude-sonnet-4-20250514';
const DEFAULT_OPENAI_MODEL = 'gpt-4o-mini';
const DEFAULT_HARVEY_BASE_URL = 'https://api.harvey.ai';
const DEFAULT_OLLAMA_BASE_URL = 'http://127.0.0.1:11434';
const DEFAULT_LMSTUDIO_BASE_URL = 'http://127.0.0.1:1234';
const DEFAULT_OLLAMA_MODEL = 'llama3.1:8b';
const DEFAULT_LMSTUDIO_MODEL = 'local-model';

// ========== CENTRALIZED ACTIVITY LOGGING (Firebase Firestore) ==========

/**
 * Log user activity to Firebase Firestore.
 * Falls back to offline queue if Firestore is unavailable.
 */
function logUserActivity(username, action) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: username || 'unknown',
    action: action,
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  logToFirestore('activity_logs', entry).then(success => {
    if (!success) {
      queuePendingLog('activity_logs', entry);
    }
  }).catch(() => {
    queuePendingLog('activity_logs', entry);
  });
}

/**
 * Log feature usage to Firebase Firestore.
 * Falls back to offline queue if Firestore is unavailable.
 */
function logUsageToFirestore(data) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: data.user_name || 'unknown',
    feature: data.feature || 'unknown',
    action: data.action || 'process',
    engine: data.engine || null,
    input_count: data.input_count || 0,
    output_count: data.output_count || 0,
    duration_ms: data.duration_ms || 0,
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  logToFirestore('usage_logs', entry).then(success => {
    if (!success) {
      queuePendingLog('usage_logs', entry);
    }
  }).catch(() => {
    queuePendingLog('usage_logs', entry);
  });
}

function logFeedbackToFirestore(data) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: String(data.username || 'unknown'),
    request: String(data.request || '').trim(),
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  logToFirestore('user_feedback', entry).then(success => {
    if (!success) {
      queuePendingLog('user_feedback', entry);
    }
  }).catch(() => {
    queuePendingLog('user_feedback', entry);
  });
}

// ========== OFFLINE QUEUE ==========

/**
 * Queue a log entry for later sync when offline.
 */
function queuePendingLog(collectionName, entry) {
  try {
    let pending = [];
    if (fs.existsSync(pendingLogsPath)) {
      pending = JSON.parse(fs.readFileSync(pendingLogsPath, 'utf8'));
    }
    pending.push({ collection: collectionName, data: entry });
    fs.writeFileSync(pendingLogsPath, JSON.stringify(pending, null, 2));
  } catch (e) {
    console.error('Failed to queue pending log:', e.message);
  }
}

/**
 * Flush all pending logs to Firestore on startup.
 */
async function syncPendingLogs() {
  if (!fs.existsSync(pendingLogsPath)) return;

  try {
    const pending = JSON.parse(fs.readFileSync(pendingLogsPath, 'utf8'));
    if (!pending || pending.length === 0) return;

    console.log(`Syncing ${pending.length} pending log(s) to Firestore...`);

    // Group by collection
    const grouped = {};
    for (const item of pending) {
      if (!grouped[item.collection]) grouped[item.collection] = [];
      grouped[item.collection].push(item.data);
    }

    let totalSynced = 0;
    const remaining = [];

    for (const [collName, docs] of Object.entries(grouped)) {
      const written = await batchLogToFirestore(collName, docs);
      totalSynced += written;
      // Keep any that failed to write
      if (written < docs.length) {
        const failed = docs.slice(written);
        for (const doc of failed) {
          remaining.push({ collection: collName, data: doc });
        }
      }
    }

    // Update pending file with any remaining entries
    if (remaining.length > 0) {
      fs.writeFileSync(pendingLogsPath, JSON.stringify(remaining, null, 2));
    } else {
      fs.unlinkSync(pendingLogsPath);
    }

    console.log(`Synced ${totalSynced} log(s). ${remaining.length} remaining.`);
  } catch (e) {
    console.error('Failed to sync pending logs:', e.message);
  }
}

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
    initAIProvider();
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

function getApiKeyFromEnvironment() {
  return (
    process.env.ANTHROPIC_API_KEY ||
    process.env.OPENAI_API_KEY ||
    process.env.HARVEY_API_KEY ||
    null
  );
}

function decodeApiKeyFromFile(rawValue) {
  if (!rawValue || rawValue.length === 0) return null;
  try {
    if (safeStorage.isEncryptionAvailable()) {
      return safeStorage.decryptString(rawValue);
    }
  } catch (decryptError) {
    // Fall through and try plain text decode (legacy/fallback path).
  }

  const plain = rawValue.toString('utf8').trim();
  return plain || null;
}

function getApiKey() {
  try {
    // First try file storage.
    if (fs.existsSync(apiKeyPath)) {
      const raw = fs.readFileSync(apiKeyPath);
      const stored = decodeApiKeyFromFile(raw);
      if (stored) return stored;
    }
    // Fallback to environment variable.
    return getApiKeyFromEnvironment();
  } catch (e) {
    console.error('Failed to get API key:', e);
    return getApiKeyFromEnvironment();
  }
}

function setApiKey(apiKey) {
  try {
    const normalizedKey = String(apiKey || '').trim();
    if (!normalizedKey) return false;

    if (safeStorage.isEncryptionAvailable()) {
      const encrypted = safeStorage.encryptString(normalizedKey);
      fs.writeFileSync(apiKeyPath, encrypted);
      return true;
    }

    // Fallback for environments where keychain-backed encryption is unavailable.
    fs.writeFileSync(apiKeyPath, normalizedKey, { encoding: 'utf8' });
    return true;
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

function normalizeClaudeModel(model) {
  const value = (model || '').trim();
  if (!value) return DEFAULT_CLAUDE_MODEL;

  // Backward compatibility with previous internal aliases.
  if (value === 'claude-sonnet-4-6') {
    return 'claude-sonnet-4-20250514';
  }

  return value;
}

function normalizeAIProvider(provider) {
  const value = (provider || '').trim().toLowerCase();
  if (!value) return DEFAULT_AI_PROVIDER;
  if (value === 'chatgpt') return 'openai';
  if (value === 'claude') return 'anthropic';
  if (value === 'ollama' || value === 'lmstudio' || value === 'anthropic' || value === 'openai' || value === 'harvey') return value;
  if (value === 'lm studio' || value === 'lm-studio') return 'lmstudio';
  return DEFAULT_AI_PROVIDER;
}

function isLocalProvider(provider) {
  const normalized = normalizeAIProvider(provider);
  return normalized === 'ollama' || normalized === 'lmstudio';
}

function providerRequiresApiKey(provider) {
  const normalized = normalizeAIProvider(provider);
  return normalized === 'anthropic' || normalized === 'openai' || normalized === 'harvey';
}

function inferProviderFromApiKey(apiKey) {
  const value = String(apiKey || '').trim().toLowerCase();
  if (!value) return null;

  if (value.startsWith('sk-ant-')) return 'anthropic';
  if (value.startsWith('sk-proj-') || value.startsWith('sk-') || value.startsWith('sess-')) return 'openai';
  if (value.startsWith('harvey_') || value.startsWith('hv_')) return 'harvey';

  return null;
}

function resolveProviderForApiKey(provider, apiKey) {
  const preferred = normalizeAIProvider(provider);
  if (isLocalProvider(preferred)) return preferred;
  const inferred = inferProviderFromApiKey(apiKey);
  return inferred || preferred;
}

function resolveAiCallContext({ apiKey, requestedProvider }) {
  const normalizedKey = String(apiKey || '').trim();
  const preferredProvider = normalizeAIProvider(requestedProvider);
  const inferredProvider = inferProviderFromApiKey(normalizedKey);
  const provider = resolveProviderForApiKey(preferredProvider, normalizedKey);
  const providerName = getAIProviderDisplayName(provider);
  const providerAutoDetected = !isLocalProvider(preferredProvider) && !!(inferredProvider && inferredProvider !== preferredProvider);
  const providerAutoDetectNote = providerAutoDetected
    ? `Selected provider did not match API key format; using ${providerName} automatically.`
    : null;

  return {
    apiKey: normalizedKey,
    provider,
    providerName,
    providerAutoDetected,
    providerAutoDetectNote
  };
}

function getAIProviderDisplayName(provider) {
  switch (normalizeAIProvider(provider)) {
    case 'ollama':
      return 'Ollama';
    case 'lmstudio':
      return 'LM Studio';
    case 'openai':
      return 'OpenAI';
    case 'harvey':
      return 'Harvey';
    default:
      return 'Anthropic';
  }
}

function parseJsonSafe(text) {
  try {
    return JSON.parse(text);
  } catch (_) {
    return null;
  }
}

function extractJsonResponse(text) {
  const raw = String(text || '').trim();
  if (!raw) return null;

  let candidate = raw;
  if (candidate.startsWith('```')) {
    const lines = candidate.split('\n');
    const jsonLines = [];
    let inJson = false;
    for (const line of lines) {
      if (line.startsWith('```json')) { inJson = true; continue; }
      if (line.startsWith('```')) {
        if (inJson) break;
        continue;
      }
      if (inJson) jsonLines.push(line);
    }
    if (jsonLines.length) {
      candidate = jsonLines.join('\n');
    }
  }

  return parseJsonSafe(candidate);
}

function parseProcessorJsonOutput(stdoutText) {
  const raw = String(stdoutText || '').trim();
  if (!raw) return null;

  const direct = parseJsonSafe(raw);
  if (direct && typeof direct === 'object') return direct;

  const fromExtract = extractJsonResponse(raw);
  if (fromExtract && typeof fromExtract === 'object') return fromExtract;

  const lines = raw.split('\n').map(line => line.trim()).filter(Boolean);
  for (let i = lines.length - 1; i >= 0; i--) {
    const parsed = parseJsonSafe(lines[i]);
    if (parsed && typeof parsed === 'object') return parsed;
  }

  const firstBrace = raw.indexOf('{');
  const lastBrace = raw.lastIndexOf('}');
  if (firstBrace !== -1 && lastBrace > firstBrace) {
    const sliced = raw.slice(firstBrace, lastBrace + 1);
    const parsed = parseJsonSafe(sliced);
    if (parsed && typeof parsed === 'object') return parsed;
  }

  return null;
}

function extractOpenAIText(messageContent) {
  if (typeof messageContent === 'string') return messageContent;
  if (Array.isArray(messageContent)) {
    return messageContent
      .map(part => {
        if (typeof part === 'string') return part;
        if (part && typeof part.text === 'string') return part.text;
        return '';
      })
      .filter(Boolean)
      .join('\n');
  }
  return '';
}

function buildMultipartFormData(fields) {
  const boundary = `----EmmaNeigh${Date.now()}${Math.floor(Math.random() * 100000)}`;
  const parts = [];

  for (const [name, value] of Object.entries(fields || {})) {
    parts.push(
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="${name}"\r\n\r\n` +
      `${value == null ? '' : String(value)}\r\n`
    );
  }
  parts.push(`--${boundary}--\r\n`);

  return {
    boundary,
    body: Buffer.from(parts.join(''), 'utf8')
  };
}

function requestHttps({ baseUrl, path: requestPath, method = 'GET', headers = {}, body = null, timeoutMs = 60000 }) {
  const http = require('http');
  const https = require('https');
  const url = new URL(requestPath, baseUrl);
  const transport = url.protocol === 'http:' ? http : https;
  const payload = body == null ? null : (Buffer.isBuffer(body) ? body : Buffer.from(String(body), 'utf8'));

  return new Promise((resolve) => {
    const req = transport.request({
      hostname: url.hostname,
      port: url.port || (url.protocol === 'https:' ? 443 : 80),
      path: `${url.pathname}${url.search}`,
      method,
      rejectUnauthorized: false, // Required for corporate SSL inspection proxies
      headers: payload
        ? { ...headers, 'Content-Length': Buffer.byteLength(payload) }
        : headers
    }, (res) => {
      let raw = '';
      res.on('data', chunk => raw += chunk);
      res.on('end', () => {
        resolve({
          statusCode: res.statusCode || 0,
          rawBody: raw,
          jsonBody: parseJsonSafe(raw)
        });
      });
    });

    req.on('error', (err) => {
      resolve({ statusCode: 0, rawBody: '', jsonBody: null, networkError: err.message });
    });

    req.setTimeout(timeoutMs, () => {
      req.destroy();
      resolve({ statusCode: 0, rawBody: '', jsonBody: null, networkError: 'Request timed out' });
    });

    if (payload) req.write(payload);
    req.end();
  });
}

function getErrorDetail(response) {
  if (!response) return 'Unknown error';
  if (response.networkError) return `Network error: ${response.networkError}`;
  const msg = response.jsonBody?.error?.message || response.jsonBody?.message;
  if (msg) return String(msg);
  return (response.rawBody || '').substring(0, 300) || `API error: ${response.statusCode}`;
}

function isLikelyModelError(response) {
  const statusCode = response?.statusCode || 0;
  const detail = getErrorDetail(response).toLowerCase();
  if (statusCode === 404) return true;
  if (statusCode === 400 && detail.includes('model')) return true;
  return detail.includes('model') && (
    detail.includes('not found') ||
    detail.includes('invalid') ||
    detail.includes('unsupported') ||
    detail.includes('available') ||
    detail.includes('access')
  );
}

async function callAnthropicPrompt({ apiKey, prompt, maxTokens }) {
  const modelCandidates = Array.from(new Set([
    getClaudeModel(),
    DEFAULT_CLAUDE_MODEL,
    'claude-3-5-sonnet-20241022',
    'claude-3-5-haiku-20241022'
  ]));

  let lastResponse = null;
  for (const modelName of modelCandidates) {
    const payload = JSON.stringify({
      model: modelName,
      max_tokens: maxTokens,
      messages: [{ role: 'user', content: prompt }]
    });

    const response = await requestHttps({
      baseUrl: 'https://api.anthropic.com',
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: payload,
      timeoutMs: 90000
    });

    lastResponse = response;
    if (response.statusCode !== 200) {
      if (isLikelyModelError(response)) continue;
      return { success: false, error: `Anthropic error (${response.statusCode}): ${getErrorDetail(response)}` };
    }

    const text = response.jsonBody?.content?.[0]?.text;
    if (!text) {
      return { success: false, error: 'Anthropic response was missing message content.' };
    }
    return { success: true, text, modelUsed: modelName };
  }

  return {
    success: false,
    error: `No supported Claude model is available for this API key. Last error: ${getErrorDetail(lastResponse)}`
  };
}

async function fetchOpenAICompatibleModels({ baseUrl, apiKey }) {
  const headers = {};
  if (apiKey) headers.Authorization = `Bearer ${apiKey}`;
  const response = await requestHttps({
    baseUrl,
    path: '/v1/models',
    method: 'GET',
    headers,
    timeoutMs: 15000
  });
  if (response.statusCode !== 200) return [];
  const entries = Array.isArray(response.jsonBody?.data) ? response.jsonBody.data : [];
  return entries
    .map(item => String(item?.id || '').trim())
    .filter(Boolean);
}

async function callOpenAICompatiblePrompt({ baseUrl, apiKey, prompt, maxTokens, modelCandidates, providerLabel }) {
  let candidates = Array.from(new Set((modelCandidates || []).map(x => String(x || '').trim()).filter(Boolean)));
  const discovered = await fetchOpenAICompatibleModels({ baseUrl, apiKey });
  if (discovered.length > 0) {
    candidates = Array.from(new Set([...candidates, ...discovered]));
  }
  if (candidates.length === 0) {
    return { success: false, error: `${providerLabel} is reachable but no models were returned by /v1/models.` };
  }

  let lastResponse = null;
  for (const modelName of candidates) {
    const payload = JSON.stringify({
      model: modelName,
      max_tokens: maxTokens,
      temperature: 0,
      messages: [{ role: 'user', content: prompt }]
    });

    const headers = { 'Content-Type': 'application/json' };
    if (apiKey) headers.Authorization = `Bearer ${apiKey}`;

    const response = await requestHttps({
      baseUrl,
      path: '/v1/chat/completions',
      method: 'POST',
      headers,
      body: payload,
      timeoutMs: 90000
    });

    lastResponse = response;
    if (response.statusCode !== 200) {
      if (isLikelyModelError(response)) continue;
      return { success: false, error: `${providerLabel} error (${response.statusCode}): ${getErrorDetail(response)}` };
    }

    const content = extractOpenAIText(response.jsonBody?.choices?.[0]?.message?.content);
    if (!content) {
      return { success: false, error: `${providerLabel} response was missing message content.` };
    }
    return { success: true, text: content, modelUsed: modelName };
  }

  return {
    success: false,
    error: `No supported ${providerLabel} model is available. Last error: ${getErrorDetail(lastResponse)}`
  };
}

async function callOpenAIPrompt({ apiKey, prompt, maxTokens }) {
  const modelCandidates = [
    getOpenAIModel(),
    DEFAULT_OPENAI_MODEL,
    'gpt-4.1-mini',
    'gpt-4o-mini',
    'gpt-4.1',
    'gpt-4o'
  ];
  return callOpenAICompatiblePrompt({
    baseUrl: 'https://api.openai.com',
    apiKey,
    prompt,
    maxTokens,
    modelCandidates,
    providerLabel: 'OpenAI'
  });
}

async function callHarveyPrompt({ apiKey, prompt, maxTokens }) {
  const baseUrl = getHarveyBaseUrl();
  const { boundary, body } = buildMultipartFormData({
    prompt,
    mode: 'assist',
    stream: 'false',
    max_tokens: String(maxTokens)
  });

  const response = await requestHttps({
    baseUrl,
    path: '/api/v2/completion?include_citations=false',
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': `multipart/form-data; boundary=${boundary}`
    },
    body,
    timeoutMs: 90000
  });

  if (response.statusCode !== 200) {
    return { success: false, error: `Harvey error (${response.statusCode}): ${getErrorDetail(response)}` };
  }

  const text = response.jsonBody?.response || response.jsonBody?.text;
  if (!text) {
    return { success: false, error: 'Harvey response was missing message content.' };
  }
  return { success: true, text: String(text), modelUsed: 'harvey-assist' };
}

async function callOllamaPrompt({ apiKey, prompt, maxTokens }) {
  const modelCandidates = [
    getOllamaModel(),
    DEFAULT_OLLAMA_MODEL,
    'llama3.1:8b',
    'qwen2.5:7b'
  ];
  return callOpenAICompatiblePrompt({
    baseUrl: getOllamaBaseUrl(),
    apiKey,
    prompt,
    maxTokens,
    modelCandidates,
    providerLabel: 'Ollama'
  });
}

async function callLmStudioPrompt({ apiKey, prompt, maxTokens }) {
  const modelCandidates = [
    getLmStudioModel(),
    DEFAULT_LMSTUDIO_MODEL
  ];
  return callOpenAICompatiblePrompt({
    baseUrl: getLmStudioBaseUrl(),
    apiKey,
    prompt,
    maxTokens,
    modelCandidates,
    providerLabel: 'LM Studio'
  });
}

async function callProviderPrompt({ provider, apiKey, prompt, maxTokens }) {
  const normalized = resolveProviderForApiKey(provider, apiKey);
  if (normalized === 'ollama') {
    return callOllamaPrompt({ apiKey, prompt, maxTokens });
  }
  if (normalized === 'lmstudio') {
    return callLmStudioPrompt({ apiKey, prompt, maxTokens });
  }
  if (normalized === 'openai') {
    return callOpenAIPrompt({ apiKey, prompt, maxTokens });
  }
  if (normalized === 'harvey') {
    return callHarveyPrompt({ apiKey, prompt, maxTokens });
  }
  return callAnthropicPrompt({ apiKey, prompt, maxTokens });
}

function buildProviderHealthEndpoint(baseUrl, requestPath) {
  try {
    return new URL(requestPath, baseUrl).toString();
  } catch (_) {
    return `${String(baseUrl || '').replace(/\/+$/, '')}${requestPath || ''}`;
  }
}

function getProviderConnectionConfig(providerInput, apiKeyInput) {
  const apiKey = String(apiKeyInput || '').trim();
  const requestedProvider = normalizeAIProvider(providerInput || getAIProvider());
  const provider = resolveProviderForApiKey(requestedProvider, apiKey);
  const providerName = getAIProviderDisplayName(provider);

  let baseUrl = 'https://api.anthropic.com';
  let healthPath = '/v1/models';
  let model = getClaudeModel();

  if (provider === 'ollama') {
    baseUrl = getOllamaBaseUrl();
    model = getOllamaModel();
  } else if (provider === 'lmstudio') {
    baseUrl = getLmStudioBaseUrl();
    model = getLmStudioModel();
  } else if (provider === 'openai') {
    baseUrl = 'https://api.openai.com';
    model = getOpenAIModel();
  } else if (provider === 'harvey') {
    baseUrl = getHarveyBaseUrl();
    healthPath = '/api/whoami';
    model = 'harvey-assist';
  }

  return {
    requestedProvider,
    provider,
    providerName,
    providerAutoDetected: provider !== requestedProvider,
    requiresApiKey: providerRequiresApiKey(provider),
    hasApiKey: !!apiKey,
    model,
    baseUrl,
    healthPath,
    healthEndpoint: buildProviderHealthEndpoint(baseUrl, healthPath)
  };
}

async function runProviderHealthCheck(connectionConfig, apiKeyInput) {
  const apiKey = String(apiKeyInput || '').trim();
  if (!connectionConfig) {
    return { ok: false, status: 'invalid_config', detail: 'Provider configuration missing.' };
  }

  if (connectionConfig.requiresApiKey && !apiKey) {
    return {
      ok: false,
      status: 'missing_api_key',
      statusCode: 0,
      endpoint: connectionConfig.healthEndpoint,
      detail: `No ${connectionConfig.providerName} API key configured.`
    };
  }

  const headers = {};
  if (connectionConfig.provider === 'anthropic') {
    headers['anthropic-version'] = '2023-06-01';
    headers['x-api-key'] = apiKey;
  } else if (apiKey) {
    headers.Authorization = `Bearer ${apiKey}`;
  }

  const response = await requestHttps({
    baseUrl: connectionConfig.baseUrl,
    path: connectionConfig.healthPath,
    method: 'GET',
    headers,
    timeoutMs: 15000
  });

  if (response.statusCode === 200) {
    return {
      ok: true,
      status: 'reachable',
      statusCode: response.statusCode,
      endpoint: connectionConfig.healthEndpoint,
      detail: `${connectionConfig.providerName} connection successful.`
    };
  }

  return {
    ok: false,
    status: 'unreachable',
    statusCode: response.statusCode || 0,
    endpoint: connectionConfig.healthEndpoint,
    detail: getErrorDetail(response)
  };
}

const AGENT_TABS = new Set([
  'packets',
  'packetshell',
  'execution',
  'sigblocks',
  'collate',
  'redline',
  'email',
  'timetrack',
  'updatechecklist',
  'punchlist'
]);

const AGENT_TAB_ALIASES = {
  signature_packets: 'packets',
  sig_packets: 'packets',
  sigpacket: 'packets',
  packet_shell: 'packetshell',
  execution_version: 'execution',
  sig_blocks: 'sigblocks',
  checklist: 'updatechecklist',
  update_checklist: 'updatechecklist',
  activity_summary: 'timetrack',
  time_tracking: 'timetrack',
  punch_list: 'punchlist',
  email_search: 'email',
  task_detection: 'email'
};

function normalizeAgentTabName(value) {
  const raw = String(value || '')
    .trim()
    .toLowerCase()
    .replace(/-/g, '_')
    .replace(/\s+/g, '_');
  if (!raw) return null;
  const mapped = AGENT_TAB_ALIASES[raw] || raw;
  return AGENT_TABS.has(mapped) ? mapped : null;
}

function buildAgentFallbackPlan(commandText, attachments) {
  const prompt = String(commandText || '').toLowerCase();
  const files = Array.isArray(attachments) ? attachments : [];
  const hasPdf = files.some(file => file.ext === 'pdf');
  const hasDocLike = files.some(file => file.ext === 'pdf' || file.ext === 'docx' || file.ext === 'doc');

  if (/(signature|sig)\s*(packet|packets|page|pages)|signing set|execution pages?/.test(prompt)) {
    if (hasPdf) {
      return {
        action: 'run_signature_packets',
        target_tab: 'packets',
        run_now: true,
        required_extensions: ['pdf'],
        missing_requirements: [],
        user_message: `Running signature packets with ${files.filter(f => f.ext === 'pdf').length} PDF file(s).`
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'packets',
      run_now: false,
      required_extensions: ['pdf'],
      missing_requirements: ['Attach one or more PDF files to run signature packets.'],
      user_message: 'Opened Create Sig Packets. Attach PDFs, then run again.'
    };
  }

  if (/packet\s*shell|signature\s*shell/.test(prompt)) {
    if (hasDocLike) {
      return {
        action: 'run_packet_shell',
        target_tab: 'packetshell',
        run_now: true,
        required_extensions: ['pdf', 'docx', 'doc'],
        missing_requirements: [],
        user_message: `Running packet shell on ${files.filter(f => f.ext === 'pdf' || f.ext === 'docx' || f.ext === 'doc').length} file(s).`
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'packetshell',
      run_now: false,
      required_extensions: ['pdf', 'docx', 'doc'],
      missing_requirements: ['Attach one or more PDF or Word files to run packet shell.'],
      user_message: 'Opened Create Packet Shell. Attach files, then run again.'
    };
  }

  if (/redline|blackline|compare|comparison/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'redline',
      run_now: false,
      required_extensions: ['docx', 'doc', 'pdf'],
      missing_requirements: [],
      user_message: 'Opened Redline Documents.'
    };
  }

  if (/collat(e|ion)|merge comments|consolidate/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'collate',
      run_now: false,
      required_extensions: ['docx', 'doc'],
      missing_requirements: [],
      user_message: 'Opened Collate Documents.'
    };
  }

  if (/execution version|conformed|insert signed/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'execution',
      run_now: false,
      required_extensions: ['pdf'],
      missing_requirements: [],
      user_message: 'Opened Execution Versions.'
    };
  }

  if (/sig\s*block|signature block|incumbency/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'sigblocks',
      run_now: false,
      required_extensions: ['docx', 'xlsx', 'csv'],
      missing_requirements: [],
      user_message: 'Opened Sig Blocks from Checklist.'
    };
  }

  if (/email|inbox|search email/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'email',
      run_now: false,
      required_extensions: ['csv'],
      missing_requirements: [],
      user_message: 'Opened Email Search.'
    };
  }

  if (/checklist|update checklist/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'updatechecklist',
      run_now: false,
      required_extensions: ['docx', 'csv'],
      missing_requirements: [],
      user_message: 'Opened Update Checklist.'
    };
  }

  if (/punchlist|punch list/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'punchlist',
      run_now: false,
      required_extensions: ['docx'],
      missing_requirements: [],
      user_message: 'Opened Generate Punchlist.'
    };
  }

  if (/time|activity summary|timeline/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'timetrack',
      run_now: false,
      required_extensions: ['csv'],
      missing_requirements: [],
      user_message: 'Opened Activity Summary.'
    };
  }

  return {
    action: 'no_op',
    target_tab: null,
    run_now: false,
    required_extensions: [],
    missing_requirements: [],
    user_message: 'I can route commands for signature packets, packet shell, redline, collate, checklist, punchlist, email search, and time summary.'
  };
}

function sanitizeAgentPlan(rawPlan, fallbackPlan, attachments) {
  const files = Array.isArray(attachments) ? attachments : [];
  const fallback = fallbackPlan || buildAgentFallbackPlan('', files);
  const candidate = rawPlan && typeof rawPlan === 'object' ? rawPlan : {};
  const allowedActions = new Set(['run_signature_packets', 'run_packet_shell', 'open_tab', 'no_op']);
  const actionValue = String(candidate.action || '').trim().toLowerCase().replace(/-/g, '_');
  const action = allowedActions.has(actionValue) ? actionValue : fallback.action;
  let targetTab = normalizeAgentTabName(candidate.target_tab) || fallback.target_tab || null;
  let runNow = Boolean(candidate.run_now);

  const requiredExtensions = Array.isArray(candidate.required_extensions)
    ? candidate.required_extensions.map(x => String(x || '').toLowerCase().replace(/^\./, '')).filter(Boolean).slice(0, 6)
    : (Array.isArray(fallback.required_extensions) ? fallback.required_extensions : []);

  let missingRequirements = Array.isArray(candidate.missing_requirements)
    ? candidate.missing_requirements.map(x => String(x || '').trim()).filter(Boolean).slice(0, 6)
    : (Array.isArray(fallback.missing_requirements) ? fallback.missing_requirements : []);

  const userMessageRaw = String(candidate.user_message || '').trim();
  const userMessage = userMessageRaw || String(fallback.user_message || '').trim();

  if (action === 'run_signature_packets') {
    targetTab = 'packets';
    const hasPdf = files.some(file => file.ext === 'pdf');
    if (!hasPdf) {
      runNow = false;
      missingRequirements = ['Attach one or more PDF files to run signature packets.'];
    }
  } else if (action === 'run_packet_shell') {
    targetTab = 'packetshell';
    const hasSupported = files.some(file => file.ext === 'pdf' || file.ext === 'docx' || file.ext === 'doc');
    if (!hasSupported) {
      runNow = false;
      missingRequirements = ['Attach one or more PDF or Word files to run packet shell.'];
    }
  } else if (action === 'open_tab') {
    runNow = false;
    if (!targetTab) targetTab = fallback.target_tab || 'packets';
  } else {
    targetTab = null;
    runNow = false;
  }

  return {
    action,
    target_tab: targetTab,
    run_now: runNow,
    required_extensions: requiredExtensions,
    missing_requirements: missingRequirements,
    user_message: userMessage
  };
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

  // Initialize Firebase for centralized logging
  const firebaseOk = initFirebase();
  if (firebaseOk) {
    // Flush any logs that were queued while offline
    syncPendingLogs().catch(e => console.error('Pending log sync error:', e.message));
  }

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

// Save a single file to a user-selected location
ipcMain.handle('save-file', async (event, sourcePath, suggestedName) => {
  if (!sourcePath || !fs.existsSync(sourcePath)) {
    return null;
  }

  const sourceExt = path.extname(sourcePath);
  const sourceExtNoDot = sourceExt.replace('.', '').toLowerCase();
  const defaultNameBase = suggestedName || path.basename(sourcePath);
  const defaultName = path.extname(defaultNameBase)
    ? defaultNameBase
    : `${defaultNameBase}${sourceExt}`;

  const dialogOptions = {
    defaultPath: defaultName
  };

  if (sourceExtNoDot) {
    dialogOptions.filters = [{ name: `${sourceExtNoDot.toUpperCase()} File`, extensions: [sourceExtNoDot] }];
  }

  const { filePath } = await dialog.showSaveDialog(mainWindow, dialogOptions);

  if (filePath) {
    fs.copyFileSync(sourcePath, filePath);
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

// ========== LITERA COMPARE INTEGRATION ==========

// Known Litera Compare installation paths (checked in order)
const LITERA_INSTALL_PATHS = [
  'C:\\Program Files (x86)\\Litera\\Compare',
  'C:\\Program Files\\Litera\\Compare',
  'C:\\Program Files (x86)\\Litera Compare',
  'C:\\Program Files\\Litera Compare'
];

const LITERA_EXECUTABLES = {
  auto: 'lcp_auto.exe',
  word: 'lcp_main.exe',
  pdf: 'lcp_pdfcmp.exe',
  powerpoint: 'lcp_ppt.exe',
  excel: 'lcx_main.exe'
};

const LITERA_STYLE_REGISTRY_PATHS = [
  'HKCU\\Software\\Litera2',
  'HKCU\\Software\\WOW6432Node\\Litera2',
  'HKCU\\Software\\Litera',
  'HKCU\\Software\\WOW6432Node\\Litera',
  'HKCU\\Software\\Litera Compare PDF Publisher',
  'HKCU\\Software\\Litera\\Compare',
  'HKCU\\Software\\WOW6432Node\\Litera\\Compare',
  'HKCU\\Software\\Litera\\Workshare\\Compare',
  'HKCU\\Software\\WOW6432Node\\Litera\\Workshare\\Compare',
  'HKCU\\Software\\Workshare\\Compare',
  'HKCU\\Software\\WOW6432Node\\Workshare\\Compare',
  'HKLM\\Software\\Litera2',
  'HKLM\\Software\\WOW6432Node\\Litera2',
  'HKLM\\Software\\Litera',
  'HKLM\\Software\\WOW6432Node\\Litera',
  'HKLM\\Software\\Litera Compare PDF Publisher',
  'HKLM\\Software\\WOW6432Node\\Litera Compare PDF Publisher',
  'HKLM\\Software\\Litera\\Compare',
  'HKLM\\Software\\WOW6432Node\\Litera\\Compare',
  'HKLM\\Software\\Litera\\Workshare\\Compare',
  'HKLM\\Software\\WOW6432Node\\Litera\\Workshare\\Compare',
  'HKLM\\Software\\Workshare\\Compare',
  'HKLM\\Software\\WOW6432Node\\Workshare\\Compare'
];

let literaRegistryStyleHintsCache = null;

function hasAnyLiteraExecutable(dirPath) {
  return Object.values(LITERA_EXECUTABLES)
    .some(exeName => fs.existsSync(path.join(dirPath, exeName)));
}

function scanForLiteraInstallation() {
  const roots = [process.env['ProgramFiles(x86)'], process.env.ProgramFiles]
    .map(value => String(value || '').trim())
    .filter(Boolean);
  const visited = new Set();

  function walk(dirPath, depth) {
    if (!dirPath || visited.has(dirPath) || depth > 3) return null;
    visited.add(dirPath);

    if (hasAnyLiteraExecutable(dirPath)) {
      return dirPath;
    }

    let entries = [];
    try {
      entries = fs.readdirSync(dirPath, { withFileTypes: true });
    } catch (_) {
      return null;
    }

    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const lower = entry.name.toLowerCase();
      if (!(lower.includes('litera') || lower.includes('compare') || depth < 1)) continue;
      const childPath = path.join(dirPath, entry.name);
      const found = walk(childPath, depth + 1);
      if (found) return found;
    }
    return null;
  }

  for (const root of roots) {
    const found = walk(root, 0);
    if (found) return found;
  }
  return null;
}

// Cache the Litera install path once found
let literaInstallPath = null;
let literaChecked = false;

/**
 * Find Litera Compare installation directory.
 * Returns the path if found, null otherwise.
 */
function findLiteraInstallation() {
  if (literaChecked) return literaInstallPath;
  literaChecked = true;

  if (process.platform !== 'win32') {
    console.log('Litera Compare is Windows-only');
    return null;
  }

  for (const installPath of LITERA_INSTALL_PATHS) {
    if (hasAnyLiteraExecutable(installPath)) {
      literaInstallPath = installPath;
      console.log('Found Litera Compare at:', installPath);
      return installPath;
    }
  }

  const scannedInstall = scanForLiteraInstallation();
  if (scannedInstall) {
    literaInstallPath = scannedInstall;
    console.log('Found Litera Compare by scan at:', scannedInstall);
    return scannedInstall;
  }

  console.log('Litera Compare not found');
  return null;
}

function expandWindowsEnvVars(rawValue) {
  return String(rawValue || '').replace(/%([^%]+)%/g, (fullMatch, varName) => {
    const key = String(varName || '').trim();
    if (!key) return fullMatch;
    if (Object.prototype.hasOwnProperty.call(process.env, key)) return process.env[key];
    if (Object.prototype.hasOwnProperty.call(process.env, key.toUpperCase())) return process.env[key.toUpperCase()];
    if (Object.prototype.hasOwnProperty.call(process.env, key.toLowerCase())) return process.env[key.toLowerCase()];
    return fullMatch;
  });
}

function normalizeLiteraStyleHint(value) {
  return expandWindowsEnvVars(value)
    .replace(/^"+|"+$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function hasMonochromeStyleTerm(value) {
  const lower = String(value || '').toLowerCase();
  const blockedTerms = ['black', 'mono', 'monochrome', 'grayscale', 'grey scale', 'gray scale', 'b&w'];
  return blockedTerms.some(term => lower.includes(term));
}

function looksLikeBinaryRegistryBlob(value) {
  const normalized = String(value || '').trim();
  if (!normalized) return true;
  return /^(?:[0-9a-f]{2},){4,}[0-9a-f]{2}$/i.test(normalized) || /^[0-9,\s-]+$/.test(normalized);
}

function parseRegistryStyleHints(registryOutput) {
  const hints = [];
  const lines = String(registryOutput || '').split(/\r?\n/);
  for (const line of lines) {
    const match = line.match(/^\s*([^\s].*?)\s+(REG_\w+)\s+(.+)\s*$/);
    if (!match) continue;
    const valueName = String(match[1] || '').trim().toLowerCase();
    const valueType = String(match[2] || '').trim().toUpperCase();
    const valueData = normalizeLiteraStyleHint(match[3] || '');
    if (!valueData) continue;
    if (!['REG_SZ', 'REG_EXPAND_SZ', 'REG_MULTI_SZ'].includes(valueType)) continue;
    if (looksLikeBinaryRegistryBlob(valueData)) continue;
    if (
      !valueName.includes('style') &&
      !valueName.includes('render') &&
      !/\.(tpx|tpp|tpz)\b/i.test(valueData) &&
      !/(color|colour|blue|red|green|kirkland|default)/i.test(valueData)
    ) {
      continue;
    }
    if (hasMonochromeStyleTerm(valueData)) continue;
    hints.push(valueData);
  }
  return hints;
}

function getLiteraRegistryStyleHints() {
  if (process.platform !== 'win32') return [];
  if (Array.isArray(literaRegistryStyleHintsCache)) {
    return [...literaRegistryStyleHintsCache];
  }

  const hints = [];
  const seen = new Set();

  function addHint(rawHint) {
    const hint = normalizeLiteraStyleHint(rawHint);
    if (!hint) return;
    const key = hint.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    hints.push(hint);
  }

  for (const registryPath of LITERA_STYLE_REGISTRY_PATHS) {
    try {
      const query = spawnSync('reg', ['query', registryPath, '/s'], {
        encoding: 'utf8',
        windowsHide: true,
        maxBuffer: 16 * 1024 * 1024
      });
      if (query.error || query.status !== 0 || !query.stdout) continue;
      for (const hint of parseRegistryStyleHints(query.stdout)) {
        addHint(hint);
      }
    } catch (_) {}
  }

  literaRegistryStyleHintsCache = hints;
  return [...hints];
}

/**
 * Get the appropriate Litera executable for a given file extension.
 * Returns { exe, args_template } or null if unsupported.
 */
function getLiteraTypeForExtension(fileExt) {
  const ext = (fileExt || '').toLowerCase();

  switch (ext) {
    case '.doc':
    case '.docx':
    case '.rtf':
    case '.txt':
    case '.htm':
    case '.html':
    case '.wpd':
      return 'word';
    case '.pdf':
      return 'pdf';
    case '.ppt':
    case '.pps':
    case '.pptx':
    case '.pptm':
    case '.ppsx':
    case '.ppsm':
      return 'powerpoint';
    case '.xls':
    case '.xlsx':
    case '.xlsm':
    case '.xlsb':
      return 'excel';
    case '.png':
    case '.bmp':
    case '.jpg':
    case '.jpeg':
      return 'image';
    default:
      return null;
  }
}

function getLiteraExecutable(fileExt, options = {}) {
  const literaPath = findLiteraInstallation();
  if (!literaPath) return null;

  const type = getLiteraTypeForExtension(fileExt);

  if (!type) {
    return null;
  }

  const preferPerApp = !!options.preferPerApp;

  // Prefer Litera's unified CLI documented in the Litera command-line guide.
  const autoExe = path.join(literaPath, LITERA_EXECUTABLES.auto);
  if (!preferPerApp && fs.existsSync(autoExe)) {
    return {
      exe: autoExe,
      type,
      mode: 'auto'
    };
  }

  // Fallback to per-app executables when lcp_auto is unavailable.
  if (type === 'pdf' || type === 'image') {
    if (fs.existsSync(autoExe)) {
      return {
        exe: autoExe,
        type,
        mode: 'auto'
      };
    }
    return null;
  }

  const perTypeExe = path.join(literaPath, LITERA_EXECUTABLES[type]);
  if (!fs.existsSync(perTypeExe)) {
    if (fs.existsSync(autoExe)) {
      return {
        exe: autoExe,
        type,
        mode: 'auto'
      };
    }
    return null;
  }

  return {
    exe: perTypeExe,
    type,
    mode: type
  };
}

function getLiteraOutputExtension(originalPath, compareOptions = {}) {
  if (compareOptions && compareOptions.change_pages_only) {
    return '.pdf';
  }

  const outputFormat = String((compareOptions && compareOptions.output_format) || 'native').toLowerCase();
  if (outputFormat === 'pdf') {
    return '.pdf';
  }

  const originalExt = path.extname(originalPath || '');
  return originalExt || '.docx';
}

function sanitizeFilenameSegment(name, fallback) {
  const cleaned = String(name || '')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!cleaned) return fallback;
  if (cleaned.length <= 64) return cleaned;
  return `${cleaned.slice(0, 61).trim()}...`;
}

function buildRedlineOutputFilename(originalPath, modifiedPath, outputExtension) {
  const originalBase = sanitizeFilenameSegment(path.parse(originalPath || '').name, 'Original');
  const modifiedBase = sanitizeFilenameSegment(path.parse(modifiedPath || '').name, 'Modified');
  const ext = String(outputExtension || '.docx');
  const normalizedExt = ext.startsWith('.') ? ext : `.${ext}`;
  return `Redline - ${originalBase} v. ${modifiedBase}${normalizedExt}`;
}

function getLiteraStyleExtension(literaType) {
  const expectedExtByType = {
    word: '.tpx',
    powerpoint: '.tpp',
    excel: '.tpz'
  };
  return expectedExtByType[literaType] || null;
}

function normalizeLiteraStyleName(styleName) {
  return String(styleName || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function getLiteraStyleNameVariants(styleName, expectedExt) {
  const variants = [];
  const seen = new Set();

  function add(value) {
    const trimmed = String(value || '').trim();
    if (!trimmed) return;
    const key = trimmed.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    variants.push(trimmed);
  }

  const raw = String(styleName || '').trim();
  if (!raw || !expectedExt) return variants;

  const rawExt = path.extname(raw).toLowerCase();
  const base = rawExt ? raw.slice(0, -rawExt.length).trim() : raw;

  if (rawExt === expectedExt) {
    add(raw);
  } else if (base) {
    add(`${base}${expectedExt}`);
  } else {
    add(raw);
  }

  if (base) {
    add(`${base}${expectedExt}`);
    if (!/^colou?r\s*\(/i.test(base)) {
      add(`Color (${base})${expectedExt}`);
      add(`Colour (${base})${expectedExt}`);
    }
  }

  return variants;
}

function getPreferredLiteraColorStyleNames(literaType, preferredStyleName = null) {
  const expectedExt = getLiteraStyleExtension(literaType);
  if (!expectedExt) return [];

  const preferred = getLiteraStyleNameVariants(preferredStyleName, expectedExt);
  return [
    ...preferred,
    `Blue Red Green${expectedExt}`,
    `Blue-Red-Green${expectedExt}`,
    `Blue Red${expectedExt}`,
    `Blue-Red${expectedExt}`,
    `Red Blue${expectedExt}`,
    `Red-Blue${expectedExt}`,
    `Color (Blue Red Green)${expectedExt}`,
    `Color (Blue Red)${expectedExt}`,
    `Colour (Blue Red Green)${expectedExt}`,
    `Colour (Blue Red)${expectedExt}`,
    `Color (Kirkland Default)${expectedExt}`,
    `Colour (Kirkland Default)${expectedExt}`,
    `Kirkland Default${expectedExt}`,
    `Kirkland${expectedExt}`,
    `Color (Default)${expectedExt}`,
    `Colour (Default)${expectedExt}`,
    `Default${expectedExt}`,
    `Color${expectedExt}`,
    `Colour${expectedExt}`
  ];
}

function resolveLiteraStylePath(renderingStyle, literaType, literaPath) {
  if (!renderingStyle || typeof renderingStyle !== 'string') {
    return null;
  }

  const raw = renderingStyle.trim();
  if (!raw) {
    return null;
  }

  const expectedExt = getLiteraStyleExtension(literaType);
  if (!expectedExt) {
    return null;
  }

  const rawExt = path.extname(raw).toLowerCase();
  let styleToken = raw;
  if (!rawExt) {
    styleToken = `${raw}${expectedExt}`;
  } else if (rawExt !== expectedExt) {
    styleToken = `${raw.slice(0, -rawExt.length)}${expectedExt}`;
  }

  const isAbsolute = path.isAbsolute(styleToken);
  const candidates = isAbsolute
    ? [styleToken]
    : [path.join(literaPath, styleToken), styleToken];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  console.warn(`Litera style file not found: ${styleToken}`);
  return null;
}

function findPreferredLiteraColorStyle(literaType, literaPath, preferredStyleName = null) {
  const expectedExt = getLiteraStyleExtension(literaType);
  if (!expectedExt) {
    return null;
  }

  const preferredFilenames = getPreferredLiteraColorStyleNames(literaType, preferredStyleName);
  const preferredSet = new Set(preferredFilenames.map(name => name.toLowerCase()));
  const preferredStyleBase = normalizeLiteraStyleName(path.parse(String(preferredStyleName || '')).name || preferredStyleName);
  const preferKirkland = !preferredStyleBase || preferredStyleBase.includes('kirkland');
  const blockedTerms = ['black', 'mono', 'monochrome', 'grayscale', 'grey scale', 'gray scale', 'b&w'];

  const startDirs = [literaPath, path.dirname(literaPath)];
  const directSubdirs = ['Styles', 'Style', 'Templates', 'Rendering Styles', 'Comparison Styles'];
  for (const subdir of directSubdirs) {
    const candidate = path.join(literaPath, subdir);
    if (fs.existsSync(candidate)) startDirs.push(candidate);
  }

  // Also look in known Windows data locations where style files are often stored.
  const windowsRoots = [process.env.ProgramData, process.env.APPDATA, process.env.LOCALAPPDATA].filter(Boolean);
  const windowsSuffixes = [
    ['Litera', 'Compare'],
    ['Litera', 'Compare', 'Styles'],
    ['Litera', 'Compare', 'Rendering Styles'],
    ['Litera', 'Workshare', 'Compare'],
    ['Litera', 'Workshare', 'Compare', 'Styles']
  ];
  for (const root of windowsRoots) {
    for (const suffixParts of windowsSuffixes) {
      const candidate = path.join(root, ...suffixParts);
      if (fs.existsSync(candidate)) startDirs.push(candidate);
    }
  }

  const styleFiles = [];
  const visited = new Set();
  const maxDepth = 8;

  function walk(currentDir, depth) {
    if (depth > maxDepth || visited.has(currentDir)) return;
    visited.add(currentDir);

    let entries;
    try {
      entries = fs.readdirSync(currentDir, { withFileTypes: true });
    } catch (_) {
      return;
    }

    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (entry.isDirectory()) {
        const dirName = entry.name.toLowerCase();
        if (depth < maxDepth || dirName.includes('style') || dirName.includes('template')) {
          walk(fullPath, depth + 1);
        }
        continue;
      }

      if (!entry.isFile()) continue;
      const ext = path.extname(entry.name).toLowerCase();
      if (ext === expectedExt) {
        styleFiles.push(fullPath);
      }
    }
  }

  for (const dir of startDirs) {
    walk(dir, 0);
  }

  if (!styleFiles.length) return null;

  let bestPath = null;
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const filePath of styleFiles) {
    const baseName = path.basename(filePath).toLowerCase();
    let score = 0;
    const isBlocked = blockedTerms.some(term => baseName.includes(term));

    if (preferredSet.has(baseName)) score += 120;
    if (preferredStyleBase && baseName.includes(preferredStyleBase)) score += 900;
    if (baseName.includes('color') || baseName.includes('colour')) score += 50;
    if (baseName.includes('default')) score += 20;
    if (baseName.includes('redline')) score += 10;
    if (preferKirkland && baseName.includes('kirkland')) score += 300;
    if (preferKirkland && baseName.includes('kirkland default')) score += 200;
    if (isBlocked) {
      score -= 1000;
    } else {
      score += 10;
    }

    if (score > bestScore) {
      bestScore = score;
      bestPath = filePath;
    }
  }

  if (bestScore <= 0) return null;
  return bestPath;
}

function getPreferredStyleNameFromHints(styleHints = []) {
  let firstNonMonochrome = null;
  for (const hint of styleHints) {
    const cleaned = String(hint || '').trim();
    if (!cleaned) continue;
    if (hasMonochromeStyleTerm(cleaned)) continue;
    const styleName = path.extname(cleaned) ? path.parse(cleaned).name : cleaned;
    if (/(color|colour|blue|red|green|kirkland|default)/i.test(styleName)) {
      return styleName;
    }
    if (!firstNonMonochrome) {
      firstNonMonochrome = styleName;
    }
  }
  return firstNonMonochrome;
}

function getLiteraColorStyleCandidates(literaType, literaPath, options = {}) {
  const candidates = [];
  const seen = new Set();
  const styleHints = Array.isArray(options.styleHints) ? options.styleHints : [];
  const preferredStyleName = String(options.preferredStyleName || '').trim() || null;
  const expectedExt = getLiteraStyleExtension(literaType);

  function addCandidate(value) {
    if (!value) return;
    const normalized = String(value).trim();
    if (!normalized) return;
    if (hasMonochromeStyleTerm(normalized)) return;
    const key = normalized.toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    candidates.push(normalized);
  }

  function addStyleTokenCandidates(styleName) {
    if (!styleName) return;
    const normalized = String(styleName).trim();
    if (!normalized) return;
    if (!expectedExt) {
      addCandidate(normalized);
      return;
    }

    const rawExt = path.extname(normalized);
    const base = rawExt ? normalized.slice(0, -rawExt.length).trim() : normalized;
    if (base) {
      addCandidate(base);
      if (!/^colou?r\s*\(/i.test(base)) {
        addCandidate(`Color (${base})`);
        addCandidate(`Colour (${base})`);
      }
    }

    for (const token of getLiteraStyleNameVariants(normalized, expectedExt)) {
      addCandidate(token);
      const tokenExt = path.extname(token);
      const tokenBase = tokenExt ? token.slice(0, -tokenExt.length).trim() : token;
      if (tokenBase) addCandidate(tokenBase);
    }
  }

  addCandidate(findPreferredLiteraColorStyle(literaType, literaPath, preferredStyleName));
  if (preferredStyleName) {
    addCandidate(resolveLiteraStylePath(preferredStyleName, literaType, literaPath));
    addStyleTokenCandidates(preferredStyleName);
  }

  for (const styleName of getPreferredLiteraColorStyleNames(literaType, preferredStyleName)) {
    const resolved = resolveLiteraStylePath(styleName, literaType, literaPath);
    addCandidate(resolved);
    addStyleTokenCandidates(styleName);
  }

  for (const styleHint of styleHints) {
    addCandidate(resolveLiteraStylePath(styleHint, literaType, literaPath));
    addStyleTokenCandidates(styleHint);
  }

  return candidates;
}

/**
 * Run Litera Compare on a single document pair.
 * Uses Litera's documented command-line interface.
 * Returns a Promise that resolves with the comparison result.
 */
function runLiteraComparison(originalPath, modifiedPath, outputPath, compareOptions = {}) {
  return new Promise((resolve, reject) => {
    const origExt = path.extname(originalPath);
    const originalBaseName = path.parse(originalPath || '').name.toLowerCase();
    const modifiedBaseName = path.parse(modifiedPath || '').name.toLowerCase();
    const normalizedOptions = {
      output_format: String(compareOptions.output_format || 'native').toLowerCase() === 'pdf' ? 'pdf' : 'native',
      change_pages_only: !!compareOptions.change_pages_only
    };
    const literaType = getLiteraTypeForExtension(origExt);
    const prefersPerAppForStyles = ['word', 'powerpoint', 'excel'].includes(literaType || '');
    const literaExe = getLiteraExecutable(origExt, {
      preferPerApp: normalizedOptions.change_pages_only || prefersPerAppForStyles
    });

    if (!literaExe) {
      reject(new Error(`Litera Compare does not support ${origExt} files`));
      return;
    }

    if (!fs.existsSync(literaExe.exe)) {
      reject(new Error(`Litera executable not found: ${literaExe.exe}`));
      return;
    }

    const literaInstallPath = path.dirname(literaExe.exe);
    const hasStyleSupport = !!getLiteraStyleExtension(literaExe.type);
    const registryStyleHints = hasStyleSupport ? getLiteraRegistryStyleHints() : [];
    const preferredStyleName = getPreferredStyleNameFromHints(registryStyleHints);
    const styleCandidates = hasStyleSupport
      ? getLiteraColorStyleCandidates(literaExe.type, literaInstallPath, {
          styleHints: registryStyleHints,
          preferredStyleName
        })
      : [];
    const styleAttempts = hasStyleSupport
      ? (styleCandidates.length ? [...styleCandidates] : [null])
      : [null];
    if (styleCandidates.length) {
      console.log('Litera style candidates:', styleCandidates.map(s => path.basename(String(s))));
    } else if (getLiteraStyleExtension(literaExe.type)) {
      console.warn('No explicit Litera color style found; using Litera default rendering style.');
    }

    if (normalizedOptions.output_format === 'pdf' && !normalizedOptions.change_pages_only && literaType && !['word', 'pdf'].includes(literaType)) {
      reject(new Error('PDF output is currently supported only for Word/PDF comparisons.'));
      return;
    }

    function hasOutputFile(filePath) {
      return fs.existsSync(filePath) && (() => {
        try {
          return fs.statSync(filePath).size > 0;
        } catch (_) {
          return true;
        }
      })();
    }

    function cleanupTempOutput(tempPath, options = {}) {
      if (tempPath && fs.existsSync(tempPath)) {
        const keepIfMatches = options.keepIfMatches ? path.resolve(options.keepIfMatches) : null;
        if (keepIfMatches && path.resolve(tempPath) === keepIfMatches) {
          return;
        }
        try { fs.unlinkSync(tempPath); } catch (_) {}
      }
    }

    function buildFallbackNativeOutputPath(primaryPath, sourcePath) {
      const sourceExt = path.extname(sourcePath || '') || origExt || '.docx';
      const primaryExt = path.extname(primaryPath || '');
      const normalizedSourceExt = sourceExt.startsWith('.') ? sourceExt : `.${sourceExt}`;
      if (primaryExt && primaryExt.toLowerCase() === normalizedSourceExt.toLowerCase()) {
        return primaryPath;
      }
      const basePath = primaryExt
        ? primaryPath.slice(0, -primaryExt.length)
        : primaryPath;
      return `${basePath}${normalizedSourceExt}`;
    }

    function findFallbackChangePagesPdf(primaryPath, startedAtMs) {
      const outputDir = path.dirname(primaryPath);
      let entries = [];
      try {
        entries = fs.readdirSync(outputDir, { withFileTypes: true });
      } catch (_) {
        return null;
      }

      let best = null;
      let bestScore = Number.NEGATIVE_INFINITY;

      for (const entry of entries) {
        if (!entry.isFile()) continue;
        if (path.extname(entry.name).toLowerCase() !== '.pdf') continue;
        const fullPath = path.join(outputDir, entry.name);
        let stat;
        try {
          stat = fs.statSync(fullPath);
        } catch (_) {
          continue;
        }
        if (!stat || stat.size <= 0) continue;
        if (stat.mtimeMs < (startedAtMs - 10000)) continue;

        const lowerName = entry.name.toLowerCase();
        let score = stat.mtimeMs;
        if (lowerName.includes('redp')) score += 2000;
        if (lowerName.includes('change')) score += 1000;
        if (lowerName.includes('changed')) score += 1000;
        if (originalBaseName && lowerName.includes(originalBaseName)) score += 1500;
        if (modifiedBaseName && lowerName.includes(modifiedBaseName)) score += 1500;

        if (score > bestScore) {
          best = fullPath;
          bestScore = score;
        }
      }

      return best;
    }

      function buildArgsForStyle(styleValue) {
        let args = [];
        let primaryOutputPath = outputPath;
        let tempAutoOutputPath = null;
        let warning = null;

      if (literaExe.mode === 'auto') {
        if (normalizedOptions.change_pages_only) {
          warning = 'Change pages only is not supported by this Litera installation; generated a full redline document instead.';
          tempAutoOutputPath = path.join(
            app.getPath('temp'),
            `litera_full_${Date.now()}_${Math.floor(Math.random() * 100000)}${origExt || '.docx'}`
          );
        }
        // Litera CLI: lcp_auto.exe -o <original> -m <modified> -r <redline> [-s <style>]
        args = ['-o', originalPath, '-m', modifiedPath, '-r', tempAutoOutputPath || primaryOutputPath];
        if (styleValue) {
          args.push('-s', styleValue);
        }
      } else if (literaExe.mode === 'word' || literaExe.mode === 'powerpoint') {
        // Per-app CLI fallback:
        // lcp_main.exe/lcp_ppt.exe -org <original> -mod <modified> -auto <redline> -silent
        let autoOutputPath = primaryOutputPath;
        if (normalizedOptions.change_pages_only) {
          tempAutoOutputPath = path.join(
            app.getPath('temp'),
            `litera_full_${Date.now()}_${Math.floor(Math.random() * 100000)}${origExt || '.docx'}`
          );
          autoOutputPath = tempAutoOutputPath;
        }
        args = ['-org', originalPath, '-mod', modifiedPath, '-auto', autoOutputPath, '-silent'];
        if (styleValue) {
          args.push('-style', styleValue);
        }
        if (normalizedOptions.change_pages_only) {
          args.push('-autoredp', primaryOutputPath);
        }
      } else if (literaExe.mode === 'excel') {
        if (normalizedOptions.change_pages_only) {
          throw new Error('Change pages only redline is not supported for Excel comparisons.');
        }
        // Excel silent CLI:
        // lcx_main.exe -s -lorg <original> -lmod <modified> -lres <redline> [-style <style.tpz>]
        args = ['-s', '-lorg', originalPath, '-lmod', modifiedPath, '-lres', primaryOutputPath];
        if (styleValue) {
          args.push('-style', styleValue);
        }
      } else {
        throw new Error(`Unsupported Litera mode: ${literaExe.mode}`);
      }

      return { args, primaryOutputPath, tempAutoOutputPath, warning };
    }

    mainWindow.webContents.send('redline-progress', {
      percent: 30,
      message: `Running Litera Compare (${literaExe.type})...`
    });

    function runAttempt(index) {
      if (index >= styleAttempts.length) {
        reject(new Error('Litera CLI failed to produce output.'));
        return;
      }

      const styleValue = styleAttempts[index];
      const styleLabel = styleValue ? path.basename(String(styleValue)) : 'default';
      let attemptConfig;
      try {
        attemptConfig = buildArgsForStyle(styleValue);
      } catch (err) {
        if (index < styleAttempts.length - 1) {
          console.warn(`Litera style attempt "${styleLabel}" skipped: ${err.message}`);
          runAttempt(index + 1);
          return;
        }
        reject(err);
        return;
      }

      const { args, primaryOutputPath, tempAutoOutputPath, warning } = attemptConfig;
      const attemptStartedAt = Date.now();

      mainWindow.webContents.send('redline-progress', {
        percent: 45,
        message: styleValue
          ? `Running Litera Compare with style: ${styleLabel}`
          : `Running Litera Compare with default style`
      });

      console.log('Running Litera CLI:', literaExe.exe, args);
      console.log('  Original:', originalPath);
      console.log('  Modified:', modifiedPath);
      console.log('  Output:', primaryOutputPath);
      console.log('  Style:', styleLabel);

      const proc = spawn(literaExe.exe, args, {
        cwd: literaInstallPath,
        windowsHide: true
      });

      let stdout = '';
      let stderr = '';

      proc.stdout.on('data', (data) => {
        stdout += data.toString();
        console.log('Litera PS:', data.toString().trim());
      });

      proc.stderr.on('data', (data) => {
        stderr += data.toString();
        console.error('Litera PS stderr:', data.toString().trim());
      });

      proc.on('close', (code) => {
        let resolvedOutputPath = primaryOutputPath;
        let resolvedWarning = warning || null;
        let usedChangePagesFallback = false;
        const styleRejectedPattern = /(style).*(not found|missing|unable|cannot|can't|invalid)/i;
        const styleRejected = !!styleValue && styleRejectedPattern.test(`${stdout}\n${stderr}`);

        if (styleRejected) {
          if (index < styleAttempts.length - 1) {
            console.warn(`Litera style "${styleLabel}" appears invalid. Retrying with next style candidate.`);
            cleanupTempOutput(tempAutoOutputPath);
            runAttempt(index + 1);
            return;
          }
          cleanupTempOutput(tempAutoOutputPath);
          reject(new Error(`Litera could not apply style "${styleLabel}".`));
          return;
        }

        if (!hasOutputFile(resolvedOutputPath) && tempAutoOutputPath && hasOutputFile(tempAutoOutputPath)) {
          const fallbackOutputPath = buildFallbackNativeOutputPath(primaryOutputPath, tempAutoOutputPath);
          try {
            if (path.resolve(fallbackOutputPath) !== path.resolve(tempAutoOutputPath)) {
              fs.copyFileSync(tempAutoOutputPath, fallbackOutputPath);
              resolvedOutputPath = fallbackOutputPath;
            } else {
              resolvedOutputPath = tempAutoOutputPath;
            }
          } catch (_) {
            resolvedOutputPath = tempAutoOutputPath;
          }
          usedChangePagesFallback = true;
          if (!resolvedWarning) {
            resolvedWarning = 'Change pages only output was unavailable; generated a full redline document instead.';
          }
        }

        if (!hasOutputFile(resolvedOutputPath) && normalizedOptions.change_pages_only) {
          const fallbackPdf = findFallbackChangePagesPdf(primaryOutputPath, attemptStartedAt);
          if (fallbackPdf && hasOutputFile(fallbackPdf)) {
            if (path.resolve(fallbackPdf) !== path.resolve(primaryOutputPath)) {
              try {
                fs.copyFileSync(fallbackPdf, primaryOutputPath);
                resolvedOutputPath = primaryOutputPath;
              } catch (_) {
                resolvedOutputPath = fallbackPdf;
              }
            } else {
              resolvedOutputPath = fallbackPdf;
            }
            if (!resolvedWarning) {
              resolvedWarning = 'Used Litera fallback output path for changed-pages PDF.';
            }
          }
        }

        if (hasOutputFile(resolvedOutputPath)) {
          cleanupTempOutput(tempAutoOutputPath, { keepIfMatches: resolvedOutputPath });
          const method = `CLI (${path.basename(literaExe.exe)})`;
          console.log(`Comparison complete via ${method}`);

          mainWindow.webContents.send('redline-progress', {
            percent: 90,
            message: `Comparison complete (${method})`
          });

          resolve({
            success: true,
            engine: 'litera',
            output_path: resolvedOutputPath,
            litera_type: literaExe.type,
            output_format: normalizedOptions.output_format,
            change_pages_only: normalizedOptions.change_pages_only && !usedChangePagesFallback,
            method,
            litera_style: styleValue ? path.basename(String(styleValue)) : 'Litera Default',
            warning: resolvedWarning
          });
          return;
        }

        cleanupTempOutput(tempAutoOutputPath);

        let errMsg = `Litera CLI failed (exit code ${code}).`;
        if (stdout.trim()) errMsg += `\nstdout: ${stdout.trim().substring(0, 800)}`;
        if (stderr.trim()) errMsg += `\nstderr: ${stderr.trim().substring(0, 800)}`;
        errMsg += `\nCommand: ${path.basename(literaExe.exe)} ${args.join(' ')}`;

        if (index < styleAttempts.length - 1) {
          console.warn(`Litera attempt failed with style "${styleLabel}". Retrying...`);
          runAttempt(index + 1);
          return;
        }

        reject(new Error(errMsg));
      });

      proc.on('error', (err) => {
        cleanupTempOutput(tempAutoOutputPath);
        const wrappedErr = new Error(`Failed to launch Litera command line: ${err.message}`);
        if (index < styleAttempts.length - 1) {
          console.warn(`Litera launch failed with style "${styleLabel}". Retrying...`);
          runAttempt(index + 1);
          return;
        }
        reject(wrappedErr);
      });
    }

    runAttempt(0);
  });
}

// IPC handler: Check if Litera Compare is installed
ipcMain.handle('check-litera-installed', async () => {
  const literaPath = findLiteraInstallation();
  if (!literaPath) {
    return { installed: false };
  }

  // Check which executables are available
  const hasAuto = fs.existsSync(path.join(literaPath, LITERA_EXECUTABLES.auto));
  const executables = {
    word: hasAuto || fs.existsSync(path.join(literaPath, LITERA_EXECUTABLES.word)),
    pdf: hasAuto || fs.existsSync(path.join(literaPath, LITERA_EXECUTABLES.pdf)),
    powerpoint: hasAuto || fs.existsSync(path.join(literaPath, LITERA_EXECUTABLES.powerpoint)),
    excel: hasAuto || fs.existsSync(path.join(literaPath, LITERA_EXECUTABLES.excel))
  };

  return {
    installed: true,
    path: literaPath,
    capabilities: executables
  };
});

// ========== REDLINE DOCUMENTS ==========

// Redline documents — routes through Litera Compare when available, falls back to EmmaNeigh table comparison
ipcMain.handle('redline-documents', async (event, config) => {
  const engine = config.engine || 'auto'; // 'auto', 'litera', 'emmaneigh'
  const literaOptions = {
    output_format: String(config.output_format || 'native').toLowerCase() === 'pdf' ? 'pdf' : 'native',
    change_pages_only: !!config.change_pages_only
  };
  const requiresStrictLitera =
    literaOptions.change_pages_only || literaOptions.output_format === 'pdf';

  // Determine if we should use Litera
  let useLitera = false;
  if (engine === 'litera' || engine === 'auto') {
    const literaPath = findLiteraInstallation();
    if (literaPath) {
      // Check if file types are supported by Litera
      const origExt = path.extname(config.original || (config.pairs && config.pairs[0] ? config.pairs[0].original : ''));
      const literaExe = origExt ? getLiteraExecutable(origExt, { preferPerApp: literaOptions.change_pages_only }) : null;
      if (literaExe) {
        useLitera = true;
      } else if (engine === 'litera') {
        throw new Error('Litera Compare does not support this file type');
      }
    } else if (engine === 'litera') {
      throw new Error('Litera Compare is not installed on this machine');
    }
  }

  if (!useLitera && engine === 'auto' && requiresStrictLitera) {
    throw new Error('Selected redline output options require Litera Compare for this file type.');
  }

  // ---- LITERA COMPARE PATH ----
  if (useLitera) {
    mainWindow.webContents.send('redline-progress', {
      percent: 10,
      message: 'Preparing Litera Compare...'
    });

    try {
      if (config.batch && config.pairs) {
        // Batch mode with Litera
        const results = [];
        const outputFolder = config.output_folder || path.dirname(config.pairs[0].original);

        for (let i = 0; i < config.pairs.length; i++) {
          const pair = config.pairs[i];
          const pct = 10 + Math.round((i / config.pairs.length) * 80);
          mainWindow.webContents.send('redline-progress', {
            percent: pct,
            message: `Comparing pair ${i + 1} of ${config.pairs.length} with Litera...`
          });

          const outputExt = getLiteraOutputExtension(pair.original, literaOptions);
          const outputPath = path.join(
            outputFolder,
            buildRedlineOutputFilename(pair.original, pair.modified, outputExt)
          );

          try {
            const result = await runLiteraComparison(
              pair.original,
              pair.modified,
              outputPath,
              literaOptions
            );
            results.push({ ...result, pair_index: i });
          } catch (err) {
            results.push({ success: false, error: err.message, pair_index: i });
          }
        }

        const successful = results.filter(r => r.success).length;
        return {
          type: 'result',
          success: true,
          mode: 'batch',
          engine: 'litera',
          total: config.pairs.length,
          successful,
          output_format: literaOptions.output_format,
          change_pages_only: literaOptions.change_pages_only,
          results
        };
      } else {
        // Single pair with Litera
        const outputExt = getLiteraOutputExtension(config.original, literaOptions);
        const outputPath = config.output || path.join(
          path.dirname(config.original),
          buildRedlineOutputFilename(config.original, config.modified, outputExt)
        );

        mainWindow.webContents.send('redline-progress', {
          percent: 20,
          message: 'Running Litera Compare...'
        });

        const result = await runLiteraComparison(
          config.original,
          config.modified,
          outputPath,
          literaOptions
        );

        mainWindow.webContents.send('redline-progress', {
          percent: 100,
          message: 'Litera comparison complete'
        });

        return {
          type: 'result',
          success: true,
          mode: 'single',
          engine: 'litera',
          output_path: result.output_path,
          litera_type: result.litera_type,
          output_format: result.output_format,
          change_pages_only: result.change_pages_only,
          warning: result.warning || null
        };
      }
    } catch (err) {
      // Do not silently fallback when Litera was selected and started;
      // users expect Litera-style output (including color rendering).
      throw err;
    }
  }

  // ---- EMMANEIGH TABLE COMPARISON PATH (fallback) ----
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
            result.engine = 'emmaneigh'; // Tag the engine
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

      // Check if headers contain an actual attachment filename column
      const hasAttachmentColumn = headers.some(h => h === 'attachments' || h === 'attachment');

      // Extract filenames from text (when no attachment column in CSV)
      function extractFilenamesFromText(text) {
        if (!text) return [];
        const pattern = /[\w\-\.\s]+\.(?:pdf|docx?|xlsx?|pptx?|csv|txt|zip|rar|png|jpg|jpeg|gif|bmp|tiff?|msg|eml|htm|html)\b/gi;
        const matches = text.match(pattern) || [];
        return matches.map(m => m.trim()).filter(m => m.length > 4);
      }

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

        // If no attachment column exists, try to extract filenames from subject/body
        if (!hasAttachmentColumn) {
          const foundFiles = [
            ...extractFilenamesFromText(email.subject),
            ...extractFilenamesFromText(email.body)
          ];
          if (foundFiles.length > 0) {
            // Deduplicate
            const uniqueFiles = [...new Set(foundFiles)];
            email.attachments = uniqueFiles.join('; ');
            email.has_attachments = true;
          }
        }

        emails.push(email);
      }

      const uniqueSenders = new Set(emails.map(e => e.from).filter(f => f));
      const dates = emails.map(e => e.date_sent || e.date_received).filter(d => d).sort();
      const withAttachments = emails.filter(e => e.has_attachments).length;

      const summary = {
        total_emails: emails.length,
        unique_senders: uniqueSenders.size,
        with_attachments: withAttachments,
        date_range: {
          earliest: dates[0] || null,
          latest: dates[dates.length - 1] || null
        },
        found_columns: headers,
        has_attachment_column: hasAttachmentColumn
      };

      if (!hasAttachmentColumn) {
        summary.attachment_note =
          'No dedicated attachment column found in CSV. ' +
          'Filenames were extracted from email body/subject text where possible. ' +
          'For full attachment data, export as MSG files or use Microsoft Graph API.';
      }

      return {
        success: true,
        emails: emails,
        summary: summary
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
  const requestedProvider = normalizeAIProvider(config.provider || getAIProvider());

  // Get the API key
  const apiKey = config.api_key || getApiKey();
  const provider = resolveProviderForApiKey(requestedProvider, apiKey);
  const providerName = getAIProviderDisplayName(provider);

  if (providerRequiresApiKey(provider) && !apiKey) {
    return { success: false, error: `No API key configured. Please add your ${providerName} API key in Settings.` };
  }

  let directError = null;
  try {
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

    const aiResult = await callProviderPrompt({
      provider,
      apiKey,
      prompt,
      maxTokens: 1024
    });

    if (!aiResult.success) {
      directError = aiResult.error || 'Failed to get AI response';
    } else {
      const responseText = String(aiResult.text || '').trim();
      const parsed = extractJsonResponse(responseText) || {
        answer: responseText,
        relevant_email_indices: [],
        confidence: 0.5,
        summary: ''
      };

      return {
        success: true,
        answer: parsed.answer || 'No answer provided',
        relevant_email_indices: parsed.relevant_email_indices || [],
        confidence: parsed.confidence || 0.5,
        summary: parsed.summary || '',
        model_used: aiResult.modelUsed || null,
        provider,
        query
      };
    }
  } catch (e) {
    directError = e.message;
  }

  // Anthropic-only fallback to the packaged Python processor.
  if (provider !== 'anthropic' || !processorPath) {
    return { success: false, error: directError || 'Failed to get AI response' };
  }

  // Production: spawn Python subprocess
  if (app.isPackaged && !fs.existsSync(processorPath)) {
    return { success: false, error: directError || 'NL search processor not found' };
  }

  if (process.platform !== 'win32' && app.isPackaged) {
    try { fs.chmodSync(processorPath, '755'); } catch (e) {}
  }

  const configPath = path.join(app.getPath('temp'), `nl_search_${Date.now()}.json`);
  const configData = {
    emails: config.emails || [],
    query: config.query || '',
    api_key: apiKey,
    provider
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

ipcMain.handle('agent-plan', async (event, payload) => {
  let fallbackPlanForError = null;
  let fallbackProvider = getAIProvider();
  let fallbackProviderName = getAIProviderDisplayName(fallbackProvider);
  let fallbackDiagnostics = null;
  try {
    const prompt = String(payload?.prompt || '').trim();
    if (!prompt) {
      return { success: false, error: 'Please enter a command for Agent Mode.' };
    }

    const attachmentInput = Array.isArray(payload?.attachments) ? payload.attachments : [];
    const attachments = attachmentInput
      .map((item, index) => {
        const itemPath = String(item?.path || '').trim();
        const nameFromPath = itemPath ? path.basename(itemPath) : '';
        const safeName = String(item?.name || nameFromPath || `attachment-${index + 1}`).trim();
        const ext = String(path.extname(safeName || itemPath || '') || '').toLowerCase().replace(/^\./, '');
        return { path: itemPath, name: safeName, ext };
      })
      .filter(item => item.path || item.name);

    const fallbackPlan = buildAgentFallbackPlan(prompt, attachments);
    fallbackPlanForError = fallbackPlan;
    const requestedProvider = normalizeAIProvider(payload?.provider || getAIProvider());
    const apiKey = String(payload?.apiKey || getApiKey() || '').trim();
    const connection = getProviderConnectionConfig(requestedProvider, apiKey);
    const provider = connection.provider;
    const providerName = connection.providerName;
    fallbackProvider = provider;
    fallbackProviderName = providerName;
    fallbackDiagnostics = {
      ...connection,
      health: connection.requiresApiKey && !apiKey
        ? {
            ok: false,
            status: 'missing_api_key',
            statusCode: 0,
            endpoint: connection.healthEndpoint,
            detail: `No ${providerName} API key configured.`
          }
        : null
    };

    if (providerRequiresApiKey(provider) && !apiKey) {
      return {
        success: true,
        source: 'rules',
        provider,
        providerName,
        plan: fallbackPlan,
        warning: `${providerName} API key is missing. Using local routing logic.`,
        diagnostics: fallbackDiagnostics
      };
    }

    const capabilities = [
      'packets: Create Sig Packets',
      'packetshell: Create Packet Shell',
      'execution: Execution Versions',
      'sigblocks: Sig Blocks from Checklist',
      'collate: Collate Documents',
      'redline: Redline Documents',
      'email: Email Search',
      'timetrack: Activity Summary',
      'updatechecklist: Update Checklist',
      'punchlist: Generate Punchlist'
    ];

    const attachmentSummary = attachments.length
      ? attachments.map(file => `${file.name} (${file.ext || 'unknown'})`).join(', ')
      : 'None';

    const plannerPrompt = `You are EmmaNeigh Agent Mode.
Decide how to route the user command into one app action.

Allowed actions:
- run_signature_packets (requires PDF attachments)
- run_packet_shell (requires PDF/DOC/DOCX attachments)
- open_tab
- no_op

Allowed target_tab values:
packets, packetshell, execution, sigblocks, collate, redline, email, timetrack, updatechecklist, punchlist

App capabilities:
${capabilities.join('\n')}

User command:
${prompt}

Attached files:
${attachmentSummary}

Return JSON only with this exact shape:
{
  "action": "run_signature_packets|run_packet_shell|open_tab|no_op",
  "target_tab": "packets|packetshell|execution|sigblocks|collate|redline|email|timetrack|updatechecklist|punchlist|null",
  "run_now": true,
  "required_extensions": ["pdf"],
  "missing_requirements": [],
  "user_message": "short one-sentence instruction for the user"
}

Rules:
- If request is clearly for signature packets and PDFs are attached, choose run_signature_packets and run_now true.
- If request is clearly for packet shell and compatible docs are attached, choose run_packet_shell and run_now true.
- If required files are missing, choose open_tab with run_now false and include missing_requirements.
- Never invent files or claim a workflow ran if run_now is false.
- Output valid JSON only.`;

    const aiResult = await callProviderPrompt({
      provider,
      apiKey,
      prompt: plannerPrompt,
      maxTokens: 450
    });

    if (!aiResult.success) {
      return {
        success: true,
        source: 'rules',
        provider,
        providerName,
        plan: fallbackPlan,
        warning: `LLM request failed for ${providerName}: ${aiResult.error}. Using local routing logic.`,
        diagnostics: {
          ...connection,
          health: {
            ok: false,
            status: 'llm_error',
            statusCode: 0,
            endpoint: connection.healthEndpoint,
            detail: aiResult.error
          }
        }
      };
    }

    const parsed = extractJsonResponse(aiResult.text || '');
    const plan = sanitizeAgentPlan(parsed, fallbackPlan, attachments);
    const modelUsed = aiResult.modelUsed || connection.model || null;
    return {
      success: true,
      source: parsed ? 'llm' : 'rules',
      provider,
      providerName,
      modelUsed,
      plan,
      warning: parsed ? null : 'Could not parse model output. Used rules fallback.',
      diagnostics: {
        ...connection,
        model: modelUsed,
        health: {
          ok: true,
          status: 'reachable',
          statusCode: 200,
          endpoint: connection.healthEndpoint,
          detail: `${providerName} responded successfully.`
        }
      }
    };
  } catch (e) {
    const promptFallback = String(payload?.prompt || '').trim();
    const attachments = Array.isArray(payload?.attachments)
      ? payload.attachments.map(item => ({
          path: String(item?.path || '').trim(),
          name: String(item?.name || path.basename(String(item?.path || ''))).trim(),
          ext: String(path.extname(String(item?.name || item?.path || '')) || '').toLowerCase().replace(/^\./, '')
        }))
      : [];
    const plan = fallbackPlanForError || buildAgentFallbackPlan(promptFallback, attachments);
    return {
      success: true,
      source: 'rules',
      provider: fallbackProvider,
      providerName: fallbackProviderName,
      plan,
      warning: `LLM connection failed: ${e.message}. Using local routing logic.`,
      diagnostics: fallbackDiagnostics
    };
  }
});

ipcMain.handle('agent-llm-diagnostics', async (event, payload) => {
  try {
    const requestedProvider = normalizeAIProvider(payload?.provider || getAIProvider());
    const apiKey = String(payload?.apiKey || getApiKey() || '').trim();
    const connection = getProviderConnectionConfig(requestedProvider, apiKey);
    const health = await runProviderHealthCheck(connection, apiKey);
    return {
      success: true,
      ...connection,
      health
    };
  } catch (e) {
    return {
      success: false,
      error: e.message
    };
  }
});

// ========== TASK DETECTION ==========

ipcMain.handle('detect-tasks', async (event, config) => {
  const taskModuleName = 'task_detector';
  const processorPath = getProcessorPath(taskModuleName);
  const requestedProvider = normalizeAIProvider(config.provider || getAIProvider());

  // Get the API key
  const apiKey = config.api_key || getApiKey();
  const provider = resolveProviderForApiKey(requestedProvider, apiKey);
  const providerName = getAIProviderDisplayName(provider);

  if (providerRequiresApiKey(provider) && !apiKey) {
    return { success: false, error: `No API key configured. Please add your ${providerName} API key in Settings.` };
  }

  const emails = config.emails || [];
  if (!emails.length) {
    return { success: false, error: 'No emails to analyze for tasks.' };
  }

  let directError = null;
  try {
    // Prepare email context (limit to 50 emails, 500 char bodies)
    const emailContext = emails.slice(0, 50).map((email, i) => ({
      index: i,
      from: email.from || 'Unknown',
      to: email.to || '',
      cc: email.cc || '',
      subject: email.subject || '(No Subject)',
      body: (email.body || '').substring(0, 500),
      date: email.date_received || email.date_sent || '',
      attachments: email.attachments || '',
      has_attachments: email.has_attachments || false
    }));

    const taskPrompt = `You are a legal transaction assistant analyzing emails sent to a first-year associate at a law firm. Your job is to identify actionable tasks that the associate needs to perform.

TASK CATEGORIES:
- COLLATE: Merge comments or track changes from multiple document versions into a single document. Keywords: "collate", "merge comments", "consolidate changes", "combine markups", "track changes from all parties"
- REDLINE: Compare two document versions to identify and mark up changes. Keywords: "redline", "compare", "blackline", "show changes", "amended version", "mark up differences"
- SIG_PACKETS: Extract signature pages from closing documents and organize them per signer. Keywords: "signature pages", "sig packets", "closing binder", "execution pages", "signing set"
- EXECUTION_VERSION: Merge signed/executed pages back into the original unsigned agreements. Keywords: "execution version", "conformed copy", "insert signed pages", "final executed"
- REVIEW: Review a document but no automated action is needed. Keywords: "please review", "take a look", "your thoughts on", "comments on"
- NONE: No actionable task detected (informational emails, FYIs, scheduling, etc.)

IMPORTANT RULES:
- Only detect tasks where there is a clear action being requested
- Do NOT classify FYI emails, scheduling emails, or status updates as tasks
- An email saying "attached is the agreement" without asking for action = NONE
- An email saying "please compare the attached against the original" = REDLINE
- Look for action verbs: "please collate", "can you compare", "extract signature pages", etc.

For each task detected, extract:
- task_type: one of COLLATE, REDLINE, SIG_PACKETS, EXECUTION_VERSION, REVIEW, NONE
- source_email_index: the email index this task came from
- documents: list of document filenames or descriptions referenced
- signers: list of person/entity names mentioned as signers (for SIG_PACKETS only)
- deadline: any deadline mentioned (date string or null)
- priority: HIGH (urgent/ASAP/today), MEDIUM (this week/soon), LOW (when you get a chance/no rush)
- summary: one-sentence description of what needs to be done

Respond with a JSON object:
{
    "tasks": [
        {
            "task_type": "COLLATE",
            "source_email_index": 3,
            "documents": ["Credit Agreement v3.docx"],
            "signers": [],
            "deadline": null,
            "priority": "HIGH",
            "summary": "Collate client comments on Credit Agreement"
        }
    ],
    "total_emails_analyzed": ${emailContext.length},
    "emails_with_tasks": 0
}

Respond ONLY with the JSON object, no other text.

EMAILS TO ANALYZE (${emailContext.length} emails):
${JSON.stringify(emailContext, null, 2)}`;

    const aiResult = await callProviderPrompt({
      provider,
      apiKey,
      prompt: taskPrompt,
      maxTokens: 4096
    });
    if (!aiResult.success) {
      directError = aiResult.error || 'Failed to analyze tasks';
    } else {
      // Parse JSON from response
      let result;
      try {
        const rawText = aiResult.text || '';
        let jsonText = extractJsonObjectText(rawText);
        if (!jsonText) jsonText = rawText.trim();
        result = JSON.parse(jsonText);
      } catch (e) {
        result = { tasks: [], total_emails_analyzed: emailContext.length, emails_with_tasks: 0 };
      }

      const tasks = result.tasks || [];
      const actionableTasks = tasks.filter(t => t.task_type !== 'NONE');

      return {
        success: true,
        tasks,
        actionable_tasks: actionableTasks,
        total_emails_analyzed: result.total_emails_analyzed || emailContext.length,
        emails_with_tasks: result.emails_with_tasks || actionableTasks.length,
        provider,
        model_used: aiResult.modelUsed || null
      };
    }
  } catch (e) {
    directError = e.message;
  }

  // Anthropic-only fallback to the packaged Python processor.
  if (provider !== 'anthropic' || !processorPath) {
    return { success: false, error: directError || 'Failed to analyze tasks' };
  }

  // Production: spawn Python subprocess
  if (app.isPackaged && !fs.existsSync(processorPath)) {
    return { success: false, error: directError || 'Task detector processor not found' };
  }

  if (process.platform !== 'win32' && app.isPackaged) {
    try { fs.chmodSync(processorPath, '755'); } catch (e) {}
  }

  const configPath = path.join(app.getPath('temp'), `task_detect_${Date.now()}.json`);
  const configData = {
    emails: emails,
    api_key: apiKey
  };
  fs.writeFileSync(configPath, JSON.stringify(configData));

  return new Promise((resolve, reject) => {
    let args;
    if (app.isPackaged) {
      args = [taskModuleName, configPath];
    } else {
      args = [taskModuleName, path.join(__dirname, 'python', 'task_detector.py'), configPath];
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
            event.sender.send('task-detect-progress', msg);
          }
        } catch (e) {}
      }
    });

    proc.stderr.on('data', (data) => {
      console.error('Task detector stderr:', data.toString());
    });

    proc.on('close', (code) => {
      try { fs.unlinkSync(configPath); } catch (e) {}

      if (code === 0 && result) {
        resolve(result);
      } else if (!result) {
        reject(new Error('Task detection failed with code ' + code));
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

// Test API key for the selected provider.
ipcMain.handle('test-api-key', async (event, payload) => {
  try {
    const parsedPayload = payload && typeof payload === 'object'
      ? payload
      : { apiKey: payload, provider: getAIProvider() };
    const apiKey = String(parsedPayload.apiKey || '').trim();
    const requestedProvider = normalizeAIProvider(parsedPayload.provider || getAIProvider());
    const provider = resolveProviderForApiKey(requestedProvider, apiKey);
    const providerLabel = getAIProviderDisplayName(provider);
    const providerAutoDetected = provider !== requestedProvider;

    let response;
    if (provider === 'ollama') {
      response = await requestHttps({
        baseUrl: getOllamaBaseUrl(),
        path: '/v1/models',
        method: 'GET',
        headers: apiKey ? { Authorization: `Bearer ${apiKey}` } : {},
        timeoutMs: 15000
      });
      if (response.statusCode === 200) {
        return { success: true, message: 'Ollama connection successful', provider };
      }
      return { success: false, error: `Ollama is not reachable at ${getOllamaBaseUrl()}. ${getErrorDetail(response)}` };
    }

    if (provider === 'lmstudio') {
      response = await requestHttps({
        baseUrl: getLmStudioBaseUrl(),
        path: '/v1/models',
        method: 'GET',
        headers: apiKey ? { Authorization: `Bearer ${apiKey}` } : {},
        timeoutMs: 15000
      });
      if (response.statusCode === 200) {
        return { success: true, message: 'LM Studio connection successful', provider };
      }
      return { success: false, error: `LM Studio is not reachable at ${getLmStudioBaseUrl()}. ${getErrorDetail(response)}` };
    }

    if (!apiKey) {
      return { success: false, error: `No ${providerLabel} API key provided` };
    }

    if (provider === 'openai') {
      response = await requestHttps({
        baseUrl: 'https://api.openai.com',
        path: '/v1/models',
        method: 'GET',
        headers: {
          Authorization: `Bearer ${apiKey}`
        },
        timeoutMs: 15000
      });
    } else if (provider === 'harvey') {
      response = await requestHttps({
        baseUrl: getHarveyBaseUrl(),
        path: '/api/whoami',
        method: 'GET',
        headers: {
          Authorization: `Bearer ${apiKey}`
        },
        timeoutMs: 15000
      });
    } else {
      response = await requestHttps({
        baseUrl: 'https://api.anthropic.com',
        path: '/v1/models',
        method: 'GET',
        headers: {
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01'
        },
        timeoutMs: 15000
      });
    }

    if (response.statusCode === 200) {
      return {
        success: true,
        message: `${providerLabel} API key is valid${providerAutoDetected ? ' (auto-detected from key format)' : ''}`,
        provider
      };
    }
    if (response.statusCode === 401 || response.statusCode === 403) {
      return { success: false, error: `Invalid ${providerLabel} API key` };
    }
    return { success: false, error: `${providerLabel} API error (${response.statusCode}): ${getErrorDetail(response)}` };
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

  const normalizedEmail = String(email || '').trim();
  if (normalizedEmail && !normalizedEmail.includes('@')) {
    return { success: false, error: 'Please enter a valid email address or leave it blank' };
  }

  if (!securityQuestion || !securityAnswer || !String(securityAnswer).trim()) {
    return { success: false, error: 'Security question and answer are required for password reset' };
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

    db.run(`INSERT INTO users (id, username, email, password_hash, security_question, security_answer_hash, display_name, two_factor_enabled) VALUES (?, ?, ?, ?, ?, ?, ?, 0)`,
      [id, username.toLowerCase().trim(), normalizedEmail || null, passwordHash, securityQuestion || null, securityAnswerHash, displayName || username]);
    saveDatabase();

    return { success: true, userId: id };
  } catch (e) {
    if (e.message && e.message.includes('UNIQUE')) {
      return { success: false, error: 'Username already exists' };
    }
    return { success: false, error: e.message };
  }
});

// Login with username/password
ipcMain.handle('login-user', async (event, { username, password }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username || !password) {
    return { success: false, error: 'Username and password are required' };
  }

  try {
    const result = db.exec(`SELECT id, username, password_hash, display_name, api_key_encrypted FROM users WHERE username = '${username.toLowerCase().trim()}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const row = result[0].values[0];
    const [id, uname, passwordHash, displayName, apiKeyEnc] = row;

    if (!verifyPassword(password, passwordHash)) {
      return { success: false, error: 'Invalid password' };
    }

    db.run(`UPDATE users SET last_login = datetime('now') WHERE id = '${id}'`);
    saveDatabase();
    logUserActivity(uname, 'login');

    return {
      success: true,
      user: { id, username: uname, displayName: displayName || uname, hasApiKey: !!apiKeyEnc }
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

    // Log to private activity file
    logUserActivity(username, 'login');

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

// Get path to pending logs (for debugging)
ipcMain.handle('get-activity-log-path', async () => {
  return { success: true, path: pendingLogsPath, note: 'Activity logs are centralized via Firebase Firestore' };
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
    if (!enable) {
      return { success: false, error: 'Two-factor authentication is mandatory and cannot be disabled.' };
    }

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
    db.run(`UPDATE users SET email = '${email}', email_verified = 0, two_factor_enabled = 1 WHERE id = '${userId}'`);
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
    const rawValue =
      key === 'claude_model' ? normalizeClaudeModel(value)
      : key === 'ai_provider' ? normalizeAIProvider(value)
      : (value ?? '');
    const settingValue = String(rawValue);
    db.run(`INSERT OR REPLACE INTO app_settings (key, value) VALUES ('${key}', '${settingValue.replace(/'/g, "''")}')`);
    if (key === 'claude_model') {
      process.env.CLAUDE_MODEL = settingValue;
    }
    if (key === 'ai_provider') {
      process.env.AI_PROVIDER = settingValue;
    }
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

function getSettingValue(key, fallbackValue = null) {
  if (!db) return fallbackValue;
  try {
    const result = db.exec(`SELECT value FROM app_settings WHERE key = '${key}'`);
    if (result.length > 0 && result[0].values.length > 0) {
      return result[0].values[0][0];
    }
  } catch (_) {}
  return fallbackValue;
}

function getAIProvider() {
  return normalizeAIProvider(getSettingValue('ai_provider', DEFAULT_AI_PROVIDER));
}

function initAIProvider() {
  process.env.AI_PROVIDER = getAIProvider();
}

function getOpenAIModel() {
  const model = (getSettingValue('openai_model', DEFAULT_OPENAI_MODEL) || '').trim();
  return model || DEFAULT_OPENAI_MODEL;
}

function getOllamaBaseUrl() {
  const raw = (getSettingValue('ollama_base_url', DEFAULT_OLLAMA_BASE_URL) || '').trim();
  if (!raw) return DEFAULT_OLLAMA_BASE_URL;
  return raw.endsWith('/') ? raw.slice(0, -1) : raw;
}

function getLmStudioBaseUrl() {
  const raw = (getSettingValue('lmstudio_base_url', DEFAULT_LMSTUDIO_BASE_URL) || '').trim();
  if (!raw) return DEFAULT_LMSTUDIO_BASE_URL;
  return raw.endsWith('/') ? raw.slice(0, -1) : raw;
}

function getOllamaModel() {
  const model = (getSettingValue('ollama_model', DEFAULT_OLLAMA_MODEL) || '').trim();
  return model || DEFAULT_OLLAMA_MODEL;
}

function getLmStudioModel() {
  const model = (getSettingValue('lmstudio_model', DEFAULT_LMSTUDIO_MODEL) || '').trim();
  return model || DEFAULT_LMSTUDIO_MODEL;
}

function getHarveyBaseUrl() {
  const raw = (getSettingValue('harvey_base_url', DEFAULT_HARVEY_BASE_URL) || '').trim();
  if (!raw) return DEFAULT_HARVEY_BASE_URL;
  return raw.endsWith('/') ? raw.slice(0, -1) : raw;
}

// Get the configured Claude model (or default)
function getClaudeModel() {
  const model = normalizeClaudeModel(getSettingValue('claude_model', DEFAULT_CLAUDE_MODEL));
  process.env.CLAUDE_MODEL = model;
  return model;
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
    const existingStoredPassRaw = String(getSettingValue('smtp_pass', '') || '');
    const hasNewPassword = typeof pass === 'string' && pass.length > 0;

    // Preserve existing SMTP password unless user provides a new one.
    let storedPassValue = existingStoredPassRaw;
    if (hasNewPassword) {
      storedPassValue = pass;
      if (safeStorage.isEncryptionAvailable()) {
        storedPassValue = safeStorage.encryptString(pass).toString('base64');
      }
    }

    const settings = {
      smtp_host: host,
      smtp_port: port,
      smtp_user: user,
      smtp_pass: storedPassValue,
      smtp_from: from || user
    };
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
    const saved = getSmtpSettings() || {};
    const effectiveHost = String(host || '').trim() || String(saved.smtp_host || '').trim();
    const effectivePort = String(port || '').trim() || String(saved.smtp_port || '').trim() || '587';
    const effectiveUser = String(user || '').trim() || String(saved.smtp_user || '').trim();
    const effectivePass = (typeof pass === 'string' && pass.length > 0)
      ? pass
      : String(saved.smtp_pass || '');

    if (!effectiveHost || !effectiveUser || !effectivePass) {
      return { success: false, error: 'SMTP host, user, and password are required. Save SMTP settings first.' };
    }

    const transporter = nodemailer.createTransport({
      host: effectiveHost,
      port: parseInt(effectivePort) || 587,
      secure: parseInt(effectivePort) === 465,
      auth: { user: effectiveUser, pass: effectivePass }
    });

    await transporter.verify();
    return { success: true, message: 'SMTP connection successful' };
  } catch (e) {
    return { success: false, error: `SMTP test failed: ${e.message}` };
  }
});

// ========== FIREBASE CONFIG HANDLERS ==========

// Save Firebase config to local file (not in source control)
ipcMain.handle('save-firebase-config', async (event, config) => {
  try {
    const saved = saveFirebaseConfig(config);
    if (saved) {
      // Re-initialize Firebase with the new config
      const ok = initFirebase();
      if (ok) {
        // Sync any pending logs with the new config
        syncPendingLogs().catch(e => console.error('Pending log sync error:', e.message));
      }
      return { success: true, initialized: ok };
    }
    return { success: false, error: 'Failed to save config file' };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get current Firebase config (for settings UI)
ipcMain.handle('get-firebase-config', async () => {
  const config = loadFirebaseConfig();
  return { success: true, config: config, configured: !!config };
});

// ========== USAGE HISTORY HANDLERS ==========

// Log usage event — writes to both local SQLite and Firebase Firestore
ipcMain.handle('log-usage', async (event, data) => {
  // Write to Firebase Firestore (centralized)
  logUsageToFirestore(data);

  // Also write to local SQLite (offline backup)
  if (!db) return { success: false, error: 'Database not initialized' };

  try {
    db.run(`
      INSERT INTO usage_history (user_name, feature, action, input_count, output_count, duration_ms)
      VALUES (?, ?, ?, ?, ?, ?)
    `, [
      data.user_name || 'unknown',
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

ipcMain.handle('submit-feedback', async (event, data) => {
  try {
    const username = String(data?.username || '').trim();
    const request = String(data?.request || '').trim();
    if (!username) {
      return { success: false, error: 'Username is required.' };
    }
    if (!request) {
      return { success: false, error: 'Feedback cannot be empty.' };
    }
    logFeedbackToFirestore({ username, request });
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
  config = config || {};
  const { checklistPath, emailPath } = config;
  const aiContext = resolveAiCallContext({
    apiKey: config.api_key || getApiKey(),
    requestedProvider: config.provider || getAIProvider()
  });
  const { apiKey, provider, providerName, providerAutoDetectNote } = aiContext;
  const providerModel =
    provider === 'openai' ? getOpenAIModel()
    : provider === 'anthropic' ? getClaudeModel()
    : provider === 'ollama' ? getOllamaModel()
    : provider === 'lmstudio' ? getLmStudioModel()
    : '';
  const providerBaseUrl =
    provider === 'harvey' ? getHarveyBaseUrl()
    : provider === 'ollama' ? getOllamaBaseUrl()
    : provider === 'lmstudio' ? getLmStudioBaseUrl()
    : '';

  if (providerRequiresApiKey(provider) && !apiKey) {
    return {
      success: false,
      error: `No API key configured. Please add your ${providerName} API key in Settings.`
    };
  }
  if (!checklistPath || !fs.existsSync(checklistPath)) {
    return { success: false, error: 'Checklist file not found. Please select a valid .docx checklist.' };
  }
  if (!emailPath || !fs.existsSync(emailPath)) {
    return { success: false, error: 'Email CSV file not found. Please select a valid .csv export.' };
  }

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

    if (app.isPackaged) {
      args = [clModuleName, checklistPath, emailPath, outputFolder, apiKey, provider, providerModel, providerBaseUrl];
    } else {
      args = [clModuleName, path.join(__dirname, 'python', 'checklist_updater.py'), checklistPath, emailPath, outputFolder, apiKey, provider, providerModel, providerBaseUrl];
    }

    if (providerAutoDetectNote) {
      mainWindow.webContents.send('checklist-progress', { message: providerAutoDetectNote, percent: 10 });
    }
    mainWindow.webContents.send('checklist-progress', { message: `Analyzing checklist and emails with ${providerName}...`, percent: 20 });

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

      const result = parseProcessorJsonOutput(stdout);
      if (!result) {
        const stderrDetail = String(stderr || '').trim();
        const stdoutDetail = String(stdout || '').trim().slice(-300);
        const detail = stderrDetail || stdoutDetail;
        resolve({ success: false, error: `Failed to parse checklist result.${detail ? ` Details: ${detail}` : ''}` });
        return;
      }

      resolve({
        success: !!result.success,
        outputPath: result.output_path,
        itemsUpdated: result.items_updated,
        error: result.error
      });
    });

    proc.on('error', (err) => {
      resolve({ success: false, error: err.message });
    });
  });
});

// Generate Punchlist - extract open items from checklist
ipcMain.handle('generate-punchlist', async (event, config) => {
  config = config || {};
  const { checklistPath, statusFilters } = config;
  const aiContext = resolveAiCallContext({
    apiKey: config.api_key || getApiKey(),
    requestedProvider: config.provider || getAIProvider()
  });
  const { apiKey, provider, providerName, providerAutoDetectNote } = aiContext;
  const providerModel =
    provider === 'openai' ? getOpenAIModel()
    : provider === 'anthropic' ? getClaudeModel()
    : provider === 'ollama' ? getOllamaModel()
    : provider === 'lmstudio' ? getLmStudioModel()
    : '';
  const providerBaseUrl =
    provider === 'harvey' ? getHarveyBaseUrl()
    : provider === 'ollama' ? getOllamaBaseUrl()
    : provider === 'lmstudio' ? getLmStudioBaseUrl()
    : '';

  if (providerRequiresApiKey(provider) && !apiKey) {
    return {
      success: false,
      error: `No API key configured. Please add your ${providerName} API key in Settings.`
    };
  }
  if (!checklistPath || !fs.existsSync(checklistPath)) {
    return { success: false, error: 'Checklist file not found. Please select a valid .docx checklist.' };
  }

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

    if (app.isPackaged) {
      args = [plModuleName, checklistPath, outputFolder, filtersJson, apiKey, provider, providerModel, providerBaseUrl];
    } else {
      args = [plModuleName, path.join(__dirname, 'python', 'punchlist_generator.py'), checklistPath, outputFolder, filtersJson, apiKey, provider, providerModel, providerBaseUrl];
    }

    if (providerAutoDetectNote) {
      mainWindow.webContents.send('punchlist-progress', { message: providerAutoDetectNote, percent: 15 });
    }
    mainWindow.webContents.send('punchlist-progress', { message: `Analyzing checklist with ${providerName}...`, percent: 30 });

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

      const result = parseProcessorJsonOutput(stdout);
      if (!result) {
        const stderrDetail = String(stderr || '').trim();
        const stdoutDetail = String(stdout || '').trim().slice(-300);
        const detail = stderrDetail || stdoutDetail;
        resolve({ success: false, error: `Failed to parse punchlist result.${detail ? ` Details: ${detail}` : ''}` });
        return;
      }

      resolve({
        success: !!result.success,
        outputPath: result.output_path,
        itemCount: result.item_count,
        categories: result.categories,
        error: result.error
      });
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
