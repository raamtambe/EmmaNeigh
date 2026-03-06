const { app, BrowserWindow, ipcMain, dialog, shell, safeStorage } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn, spawnSync } = require('child_process');
const archiver = require('archiver');
const crypto = require('crypto');
const nodemailer = require('nodemailer');
const os = require('os');
const { fileURLToPath } = require('url');
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
const feedbackLogPath = path.join(app.getPath('userData'), 'feedback_log.json');
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
const REQUIRE_FIREBASE_TELEMETRY = false;
const FIREBASE_TELEMETRY_PROBE_COLLECTION = 'telemetry_health';
const FIREBASE_STATUS_SUCCESS_CACHE_MS = 30000;
const FIREBASE_STATUS_FAILURE_CACHE_MS = 8000;
const TELEMETRY_INGEST_TIMEOUT_MS = 20000;
const TELEMETRY_INGEST_KEY = 'telemetry_ingest_url';
const TELEMETRY_INGEST_TOKEN_KEY = 'telemetry_ingest_token';
const TELEMETRY_CONFIG_FILENAME = 'telemetry-config.json';
const FEEDBACK_ADMIN_DEFAULTS = ['rtambe', 'raamtambe'];
const ANALYTICS_ADMIN_DEFAULTS = ['rtambe', 'raamtambe'];
const FEEDBACK_ADMIN_USERNAMES = new Set(
  String(process.env.EMMANEIGH_FEEDBACK_ADMINS || FEEDBACK_ADMIN_DEFAULTS.join(','))
    .split(',')
    .map((value) => String(value || '').trim().toLowerCase())
    .filter(Boolean)
);
const ANALYTICS_ADMIN_IDENTIFIERS = new Set(
  String(
    process.env.EMMANEIGH_ANALYTICS_ADMINS ||
    process.env.EMMANEIGH_FEEDBACK_ADMINS ||
    ANALYTICS_ADMIN_DEFAULTS.join(',')
  )
    .split(',')
    .map((value) => String(value || '').trim().toLowerCase())
    .filter(Boolean)
);

const trackedChildProcesses = new Map();
let trackedProcessCounter = 0;
const rawSpawn = spawn;

function spawnTracked(command, args = [], options = {}) {
  const proc = rawSpawn(command, args, options);
  const id = ++trackedProcessCounter;
  trackedChildProcesses.set(id, {
    pid: proc.pid,
    command: String(command || ''),
    startedAt: Date.now(),
    proc
  });
  const cleanup = () => trackedChildProcesses.delete(id);
  proc.once('close', cleanup);
  proc.once('exit', cleanup);
  proc.once('error', cleanup);
  return proc;
}

function terminateTrackedProcess(entry) {
  if (!entry || !entry.proc) return false;
  const target = entry.proc;
  if (target.killed) return true;
  try {
    if (process.platform === 'win32' && entry.pid) {
      // Kill child tree on Windows to ensure PowerShell/Litera/python descendants are terminated.
      spawnSync('taskkill', ['/PID', String(entry.pid), '/T', '/F'], { windowsHide: true });
      return true;
    }
    target.kill('SIGKILL');
    return true;
  } catch (_) {
    return false;
  }
}

function normalizeLocalPath(rawValue) {
  if (rawValue === null || rawValue === undefined) return '';
  let value = Array.isArray(rawValue) ? rawValue[0] : rawValue;
  value = String(value || '').trim();
  if (!value) return '';

  // Strip wrapping quotes from copy-pasted paths.
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    value = value.slice(1, -1).trim();
  }

  if (!value) return '';
  if (/^file:\/\//i.test(value)) {
    try {
      value = fileURLToPath(value);
    } catch (_) {
      value = value.replace(/^file:\/+/, '');
      try {
        value = decodeURIComponent(value);
      } catch (_) {}
    }
  }

  return path.normalize(value);
}

function resolveExistingLocalPath(rawValue) {
  const normalized = normalizeLocalPath(rawValue);
  if (!normalized) return '';
  const candidates = [
    normalized,
    normalized.replace(/\//g, path.sep),
    normalized.replace(/\\/g, path.sep)
  ];
  try {
    const decoded = decodeURIComponent(normalized);
    candidates.push(decoded, decoded.replace(/\//g, path.sep), decoded.replace(/\\/g, path.sep));
  } catch (_) {}

  const seen = new Set();
  for (const candidate of candidates) {
    const c = String(candidate || '').trim();
    if (!c || seen.has(c)) continue;
    seen.add(c);
    if (fs.existsSync(c)) return c;
  }

  return normalized;
}

function normalizeToolLoadedFiles(loadedFiles = []) {
  if (!Array.isArray(loadedFiles)) return [];
  return loadedFiles
    .map((file) => {
      const rawPath = resolveExistingLocalPath(file && file.path);
      const nameFromPath = rawPath ? path.basename(rawPath) : '';
      const rawName = String((file && file.name) || nameFromPath || '').trim();
      const name = rawName || nameFromPath;
      const lowerName = String(name || '').toLowerCase();
      return rawPath ? {
        path: rawPath,
        name,
        lowerName
      } : null;
    })
    .filter(Boolean);
}

function resolveToolFilePath(input = {}, loadedFiles = [], preferredIndex = 0) {
  const explicitPath = resolveExistingLocalPath(input.file_path || input.path || '');
  if (explicitPath && fs.existsSync(explicitPath)) return explicitPath;

  const normalizedLoaded = normalizeToolLoadedFiles(loadedFiles);
  const hintedName = String(input.file_name || input.filename || input.name || '').trim().toLowerCase();
  if (hintedName) {
    const match = normalizedLoaded.find((item) =>
      item.lowerName === hintedName || item.lowerName.includes(hintedName) || hintedName.includes(item.lowerName)
    );
    if (match) return match.path;
  }

  if (normalizedLoaded[preferredIndex]) {
    return normalizedLoaded[preferredIndex].path;
  }

  return '';
}

function normalizeUsername(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidEmail(value) {
  const email = normalizeEmail(value);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function buildUsernameFromEmail(email) {
  const localPart = normalizeEmail(email).split('@')[0] || 'user';
  const sanitized = localPart
    .replace(/[^a-z0-9._-]/g, '_')
    .replace(/^[._-]+|[._-]+$/g, '');
  return sanitized || 'user';
}

function getAvailableUsername(baseUsername) {
  if (!db) return baseUsername || 'user';
  const base = String(baseUsername || 'user').trim() || 'user';
  let candidate = base;
  let index = 1;
  while (index < 100000) {
    const escaped = candidate.replace(/'/g, "''");
    const result = db.exec(`SELECT id FROM users WHERE username = '${escaped}' LIMIT 1`);
    const exists = result.length > 0 && result[0].values.length > 0;
    if (!exists) return candidate;
    candidate = `${base}_${index}`;
    index += 1;
  }
  return `${base}_${Date.now()}`;
}

function parseSqliteDate(value) {
  if (!value) return null;
  const text = String(value).trim();
  if (!text) return null;
  const normalized = text.includes('T') ? text : text.replace(' ', 'T');
  const parsed = new Date(normalized);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed;
}

function categorizeFeatureLabel(featureName) {
  const key = String(featureName || '').trim().toLowerCase();
  if (!key) return 'Other';
  if (key.startsWith('imanage')) return 'iManage';
  if (key.includes('redline')) return 'Redlines';
  if (key === 'email') return 'Email Search';
  if (key === 'collate') return 'Collate';
  if (key === 'timetrack') return 'Time Tracking';
  if (key === 'update_checklist') return 'Checklist Updates';
  if (key === 'generate_punchlist' || key === 'punchlist') return 'Punchlists';
  if (key === 'signature_packets' || key === 'packet_shell' || key === 'execution_version' || key === 'sigblocks') {
    return 'Signature Workflows';
  }
  return key
    .replace(/[_-]+/g, ' ')
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function canAccessFeedbackLog(username) {
  return FEEDBACK_ADMIN_USERNAMES.has(normalizeUsername(username));
}

function canAccessAnalyticsDashboard(identity = {}) {
  const username = normalizeUsername(identity.username || '');
  const email = normalizeEmail(identity.email || '');
  const localPart = email.includes('@') ? email.split('@')[0] : '';
  return (
    (!!username && ANALYTICS_ADMIN_IDENTIFIERS.has(username)) ||
    (!!email && ANALYTICS_ADMIN_IDENTIFIERS.has(email)) ||
    (!!localPart && ANALYTICS_ADMIN_IDENTIFIERS.has(localPart))
  );
}

function resolveRequesterIdentity(requester = {}) {
  const fallback = {
    username: String(requester.username || '').trim(),
    email: String(requester.email || '').trim()
  };
  if (!db) return fallback;

  const userId = String(requester.userId || '').trim();
  if (!userId) return fallback;

  try {
    const escapedUserId = userId.replace(/'/g, "''");
    const result = db.exec(`
      SELECT username, email
      FROM users
      WHERE id = '${escapedUserId}'
      LIMIT 1
    `);
    if (result.length > 0 && result[0].values.length > 0) {
      const [username, email] = result[0].values[0];
      return {
        username: String(username || '').trim(),
        email: String(email || '').trim()
      };
    }
  } catch (_) {}

  return fallback;
}

function canRequesterAccessAnalytics(requester = {}) {
  const identity = resolveRequesterIdentity(requester);
  return canAccessAnalyticsDashboard(identity);
}

function buildSessionUserPayload({ id, username, displayName, email, apiKeyEnc }) {
  const normalizedEmail = normalizeEmail(email || '');
  return {
    id,
    username,
    displayName: displayName || username,
    email: normalizedEmail || null,
    hasApiKey: !!apiKeyEnc,
    isAdmin: canAccessAnalyticsDashboard({ username, email: normalizedEmail })
  };
}

function normalizeIngestUrl(rawUrl) {
  const value = String(rawUrl || '').trim();
  if (!value) return '';
  try {
    const parsed = new URL(value);
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return '';
    return parsed.toString();
  } catch (_) {
    return '';
  }
}

function encodeSecretForSetting(secretValue) {
  const plain = String(secretValue || '').trim();
  if (!plain) return '';
  try {
    if (safeStorage.isEncryptionAvailable()) {
      const encrypted = safeStorage.encryptString(plain).toString('base64');
      return `enc:${encrypted}`;
    }
  } catch (_) {}
  return `plain:${plain}`;
}

function decodeSecretFromSetting(storedValue) {
  const value = String(storedValue || '').trim();
  if (!value) return '';

  if (value.startsWith('enc:')) {
    const payload = value.substring(4);
    try {
      if (safeStorage.isEncryptionAvailable()) {
        return safeStorage.decryptString(Buffer.from(payload, 'base64'));
      }
    } catch (_) {}
    return '';
  }

  if (value.startsWith('plain:')) {
    return value.substring(6);
  }

  // Legacy fallback: value may already be plain text.
  return value;
}

let bundledTelemetryConfigCache = null;
let bundledTelemetryConfigLoaded = false;

function loadBundledTelemetryConfig() {
  if (bundledTelemetryConfigLoaded) {
    return bundledTelemetryConfigCache;
  }

  bundledTelemetryConfigLoaded = true;
  const candidatePaths = [
    String(process.env.EMMANEIGH_TELEMETRY_CONFIG_PATH || '').trim(),
    path.join(process.resourcesPath || '', TELEMETRY_CONFIG_FILENAME),
    path.join(__dirname, TELEMETRY_CONFIG_FILENAME)
  ].filter(Boolean);

  for (const candidatePath of candidatePaths) {
    try {
      if (!fs.existsSync(candidatePath)) continue;
      const raw = fs.readFileSync(candidatePath, 'utf8');
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        bundledTelemetryConfigCache = parsed;
        return bundledTelemetryConfigCache;
      }
    } catch (_) {}
  }

  bundledTelemetryConfigCache = null;
  return null;
}

function getTelemetryIngestUrl() {
  const fromSetting = normalizeIngestUrl(getSettingValue(TELEMETRY_INGEST_KEY, ''));
  if (fromSetting) return fromSetting;

  const fromEnv = normalizeIngestUrl(process.env.EMMANEIGH_TELEMETRY_INGEST_URL || '');
  if (fromEnv) return fromEnv;

  const bundled = loadBundledTelemetryConfig();
  if (!bundled) return '';
  return normalizeIngestUrl(
    bundled.telemetryIngestUrl ||
    bundled.ingestUrl ||
    bundled.url ||
    ''
  );
}

function getTelemetryIngestToken() {
  const fromSetting = decodeSecretFromSetting(getSettingValue(TELEMETRY_INGEST_TOKEN_KEY, ''));
  if (fromSetting) return fromSetting;

  const fromEnv = String(process.env.EMMANEIGH_TELEMETRY_INGEST_TOKEN || '').trim();
  if (fromEnv) return fromEnv;

  const bundled = loadBundledTelemetryConfig();
  if (!bundled) return '';
  return String(
    bundled.telemetryIngestToken ||
    bundled.ingestToken ||
    bundled.token ||
    bundled.apiKey ||
    ''
  ).trim();
}

function invalidateTelemetryStatusCache() {
  firebaseStatusCache.checkedAt = 0;
}

function writeSettingValue(key, value) {
  if (!db) return false;
  const settingValue = String(value ?? '');
  db.run(`INSERT OR REPLACE INTO app_settings (key, value) VALUES ('${key}', '${settingValue.replace(/'/g, "''")}')`);
  return true;
}

async function sendTelemetryToIngestBackend({
  eventType,
  collection,
  payload,
  context = 'telemetry_event',
  urlOverride = '',
  tokenOverride = ''
}) {
  const ingestUrl = normalizeIngestUrl(urlOverride || getTelemetryIngestUrl());
  if (!ingestUrl) {
    return { configured: false, success: false, error: 'Telemetry ingest URL is not configured.' };
  }

  const ingestToken = String(tokenOverride || getTelemetryIngestToken() || '').trim();
  const requestPayload = {
    event_id: crypto.randomUUID(),
    event_type: String(eventType || 'event'),
    collection: String(collection || ''),
    source: 'emmaneigh-desktop',
    app_version: APP_VERSION,
    machine_id: MACHINE_ID,
    context,
    timestamp: new Date().toISOString(),
    data: payload || {}
  };

  const headers = { 'Content-Type': 'application/json' };
  if (ingestToken) {
    headers.Authorization = `Bearer ${ingestToken}`;
    headers['x-emmaneigh-ingest-key'] = ingestToken;
  }

  const response = await requestHttps({
    baseUrl: ingestUrl,
    path: '',
    method: 'POST',
    headers,
    body: JSON.stringify(requestPayload),
    timeoutMs: TELEMETRY_INGEST_TIMEOUT_MS
  });

  if (response.statusCode >= 200 && response.statusCode < 300) {
    return { configured: true, success: true, statusCode: response.statusCode };
  }

  return {
    configured: true,
    success: false,
    statusCode: response.statusCode,
    error: getErrorDetail(response)
  };
}

async function writeTelemetryRecord({ eventType, collection, payload }) {
  const directWrite = async () => {
    const initialized = initFirebase();
    if (!initialized) return false;
    return logToFirestore(collection, {
      ...payload,
      event_type: eventType || null
    });
  };

  const ingestUrl = getTelemetryIngestUrl();
  if (ingestUrl) {
    const backendResult = await sendTelemetryToIngestBackend({
      eventType,
      collection,
      payload,
      context: `ingest_${eventType || 'event'}`
    });
    if (backendResult.success) return true;

    // Resiliency fallback: if direct Firebase is configured, keep telemetry writes flowing.
    return directWrite();
  }

  return directWrite();
}

let firebaseStatusCache = {
  checkedAt: 0,
  status: {
    required: REQUIRE_FIREBASE_TELEMETRY,
    configured: false,
    initialized: false,
    connected: false,
    message: 'Firebase telemetry status has not been checked yet.'
  }
};

function cacheFirebaseStatus(status) {
  firebaseStatusCache = {
    checkedAt: Date.now(),
    status
  };
  return status;
}

async function getFirebaseTelemetryStatus(options = {}) {
  const force = !!options.force;
  const skipProbe = !!options.skipProbe;
  const context = String(options.context || 'runtime_check');
  const mode = String(options.mode || 'any').toLowerCase();
  const now = Date.now();
  const cacheAge = now - firebaseStatusCache.checkedAt;
  const cacheTtl = firebaseStatusCache.status.connected
    ? FIREBASE_STATUS_SUCCESS_CACHE_MS
    : FIREBASE_STATUS_FAILURE_CACHE_MS;

  if (!force && cacheAge >= 0 && cacheAge < cacheTtl) {
    return firebaseStatusCache.status;
  }

  const backendIngestUrl = getTelemetryIngestUrl();
  if (mode !== 'direct' && backendIngestUrl) {
    if (skipProbe) {
      return {
        required: REQUIRE_FIREBASE_TELEMETRY,
        configured: true,
        initialized: true,
        connected: true,
        mode: 'backend_ingest',
        message: `Telemetry ingest backend configured at ${backendIngestUrl}.`
      };
    }

    const backendProbe = await sendTelemetryToIngestBackend({
      eventType: 'telemetry_probe',
      collection: FIREBASE_TELEMETRY_PROBE_COLLECTION,
      payload: {
        mode: 'backend_ingest',
        probe: true
      },
      context
    });

    if (backendProbe.success) {
      return cacheFirebaseStatus({
        required: REQUIRE_FIREBASE_TELEMETRY,
        configured: true,
        initialized: true,
        connected: true,
        mode: 'backend_ingest',
        message: 'Telemetry backend ingest is connected.'
      });
    }

    // If backend ingest is unavailable, attempt direct Firebase fallback in mixed deployments.
    if (mode !== 'backend') {
      const directFallback = await getFirebaseTelemetryStatus({
        force: true,
        skipProbe: false,
        context: `${context}_direct_fallback`,
        mode: 'direct'
      });
      if (directFallback.connected) {
        return cacheFirebaseStatus({
          ...directFallback,
          mode: 'direct_firebase_fallback',
          message: `Telemetry ingest backend is unavailable (${backendProbe.error || 'unknown error'}). Falling back to direct Firebase telemetry.`
        });
      }
    }

    return cacheFirebaseStatus({
      required: REQUIRE_FIREBASE_TELEMETRY,
      configured: true,
      initialized: true,
      connected: false,
      mode: 'backend_ingest',
      message: `Telemetry backend ingest is configured but unreachable: ${backendProbe.error || 'unknown error'}`
    });
  }

  if (mode === 'backend') {
    return cacheFirebaseStatus({
      required: REQUIRE_FIREBASE_TELEMETRY,
      configured: false,
      initialized: false,
      connected: false,
      mode: 'backend_ingest',
      message: 'Telemetry backend ingest URL is not configured.'
    });
  }

  const config = loadFirebaseConfig();
  if (!config) {
    return cacheFirebaseStatus({
      required: REQUIRE_FIREBASE_TELEMETRY,
      configured: false,
      initialized: false,
      connected: false,
      mode: 'direct_firebase',
      message: 'Firebase is not configured. Add Firebase config JSON in Settings.'
    });
  }

  const initialized = initFirebase();
  if (!initialized) {
    return cacheFirebaseStatus({
      required: REQUIRE_FIREBASE_TELEMETRY,
      configured: true,
      initialized: false,
      connected: false,
      mode: 'direct_firebase',
      message: 'Firebase initialization failed. Verify Firebase config JSON.'
    });
  }

  if (skipProbe) {
    return {
      required: REQUIRE_FIREBASE_TELEMETRY,
      configured: true,
      initialized: true,
      connected: true,
      mode: 'direct_firebase',
      message: 'Firebase initialized.'
    };
  }

  const probeSuccess = await logToFirestore(FIREBASE_TELEMETRY_PROBE_COLLECTION, {
    event: 'telemetry_probe',
    context,
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  });

  if (!probeSuccess) {
    return cacheFirebaseStatus({
      required: REQUIRE_FIREBASE_TELEMETRY,
      configured: true,
      initialized: true,
      connected: false,
      mode: 'direct_firebase',
      message: 'Firebase is configured but not reachable for Firestore writes.'
    });
  }

  return cacheFirebaseStatus({
    required: REQUIRE_FIREBASE_TELEMETRY,
    configured: true,
    initialized: true,
    connected: true,
    mode: 'direct_firebase',
    message: 'Firebase telemetry is connected.'
  });
}

async function requireFirebaseTelemetry(actionLabel) {
  if (!REQUIRE_FIREBASE_TELEMETRY) return { success: true };
  const status = await getFirebaseTelemetryStatus({ context: actionLabel || 'required_action' });
  if (status.connected) return { success: true };
  return {
    success: false,
    error: `Firebase telemetry is required for ${actionLabel || 'this action'}. ${status.message}`
  };
}

function telemetryWriteIsMandatory() {
  return REQUIRE_FIREBASE_TELEMETRY;
}

function writeLocalActivityRecord(entry) {
  if (!db) return false;
  try {
    db.run(`
      INSERT INTO user_activity_history (username, email, action, app_version, machine_id)
      VALUES (?, ?, ?, ?, ?)
    `, [
      String(entry.username || 'unknown'),
      entry.email ? String(entry.email) : null,
      String(entry.action || 'activity'),
      String(entry.app_version || APP_VERSION),
      String(entry.machine_id || MACHINE_ID)
    ]);
    saveDatabase();
    return true;
  } catch (e) {
    console.error('Failed to write local activity history:', e.message);
    return false;
  }
}

function getUsageHistoryColumns() {
  if (!db) return new Set();
  try {
    const info = db.exec('PRAGMA table_info(usage_history)');
    if (info.length === 0) return new Set();
    return new Set(
      info[0].values.map((row) => String(row[1] || '').trim().toLowerCase()).filter(Boolean)
    );
  } catch (_) {
    return new Set();
  }
}

function ensureUsageHistorySchema() {
  if (!db) return;
  try {
    const columns = getUsageHistoryColumns();
    if (columns.size === 0) return;

    const addColumnIfMissing = (name, definition) => {
      const normalized = String(name || '').trim().toLowerCase();
      if (!normalized || columns.has(normalized)) return;
      db.run(`ALTER TABLE usage_history ADD COLUMN ${normalized} ${definition}`);
      columns.add(normalized);
    };

    addColumnIfMissing('user_name', 'TEXT');
    addColumnIfMissing('user_email', 'TEXT');
    addColumnIfMissing('feature', 'TEXT');
    addColumnIfMissing('action', 'TEXT');
    addColumnIfMissing('input_count', 'INTEGER DEFAULT 0');
    addColumnIfMissing('output_count', 'INTEGER DEFAULT 0');
    addColumnIfMissing('duration_ms', 'INTEGER DEFAULT 0');
    addColumnIfMissing('timestamp', 'DATETIME');

    if (columns.has('user_id') && columns.has('user_name')) {
      db.run(`
        UPDATE usage_history
        SET user_name = COALESCE(
          NULLIF(TRIM(user_name), ''),
          (SELECT username FROM users WHERE users.id = usage_history.user_id LIMIT 1),
          'unknown'
        )
        WHERE user_name IS NULL OR TRIM(user_name) = ''
      `);
    }

    if (columns.has('user_id') && columns.has('user_email')) {
      db.run(`
        UPDATE usage_history
        SET user_email = COALESCE(
          NULLIF(TRIM(user_email), ''),
          (SELECT email FROM users WHERE users.id = usage_history.user_id LIMIT 1),
          ''
        )
        WHERE user_email IS NULL OR TRIM(user_email) = ''
      `);
    }

    db.run(`CREATE INDEX IF NOT EXISTS idx_usage_history_timestamp ON usage_history(timestamp)`);
    db.run(`CREATE INDEX IF NOT EXISTS idx_usage_history_feature ON usage_history(feature)`);
    db.run(`CREATE INDEX IF NOT EXISTS idx_usage_history_user_email ON usage_history(user_email)`);
  } catch (e) {
    console.error('Failed to migrate usage_history schema:', e.message);
  }
}

// ========== CENTRALIZED ACTIVITY LOGGING (Firebase Firestore) ==========

/**
 * Log user activity to Firebase Firestore.
 * Falls back to offline queue if Firestore is unavailable.
 */
async function logUserActivity(username, action, metadata = {}) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: username || 'unknown',
    action: action,
    email: metadata.email ? normalizeEmail(metadata.email) : null,
    display_name: metadata.displayName || null,
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  // Always record locally (never fails except on DB issues)
  const localSaved = writeLocalActivityRecord(entry);

  // Best-effort Firebase write — never blocks the caller
  const telemetrySaved = await writeTelemetryRecord({
    eventType: 'user_activity',
    collection: 'activity_logs',
    payload: entry
  });

  return localSaved || telemetrySaved;
}

/**
 * Log feature usage to Firebase Firestore.
 * Falls back to offline queue if Firestore is unavailable.
 */
async function logUsageToFirestore(data) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: data.user_name || 'unknown',
    email: data.user_email ? normalizeEmail(data.user_email) : null,
    feature: data.feature || 'unknown',
    action: data.action || 'process',
    engine: data.engine || null,
    input_count: data.input_count || 0,
    output_count: data.output_count || 0,
    duration_ms: data.duration_ms || 0,
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  return writeTelemetryRecord({
    eventType: 'usage',
    collection: 'usage_logs',
    payload: entry
  });
}

function normalizeUsageEventData(data = {}) {
  const normalizedFeature = String(data.feature || 'unknown').trim().toLowerCase() || 'unknown';
  const normalizedAction = String(data.action || 'process').trim().toLowerCase() || 'process';
  return {
    user_name: String(data.user_name || data.username || 'unknown').trim() || 'unknown',
    user_email: normalizeEmail(data.user_email || data.email || ''),
    feature: normalizedFeature,
    action: normalizedAction,
    engine: data.engine ? String(data.engine).trim().toLowerCase() : null,
    input_count: Number.isFinite(Number(data.input_count)) ? Math.max(0, Number(data.input_count)) : 0,
    output_count: Number.isFinite(Number(data.output_count)) ? Math.max(0, Number(data.output_count)) : 0,
    duration_ms: Number.isFinite(Number(data.duration_ms)) ? Math.max(0, Math.round(Number(data.duration_ms))) : 0
  };
}

function writeLocalUsageRecord(data = {}) {
  if (!db) return false;
  try {
    const entry = normalizeUsageEventData(data);
    db.run(`
      INSERT INTO usage_history (user_name, user_email, feature, action, input_count, output_count, duration_ms)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `, [
      entry.user_name,
      entry.user_email || null,
      entry.feature,
      entry.action,
      entry.input_count,
      entry.output_count,
      entry.duration_ms
    ]);
    saveDatabase();
    return true;
  } catch (e) {
    console.error('Failed to write local usage history:', e.message);
    return false;
  }
}

async function recordUsageEvent(data = {}) {
  const entry = normalizeUsageEventData(data);

  // Always write locally first (never fails except on DB issues)
  const localSaved = writeLocalUsageRecord(entry);

  // Best-effort Firebase write — never blocks or fails the caller
  const telemetrySaved = await logUsageToFirestore(entry);

  return {
    success: localSaved || telemetrySaved,
    localLogged: localSaved,
    telemetryLogged: telemetrySaved
  };
}

async function logFeedbackToFirestore(data) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: String(data.username || 'unknown'),
    request: String(data.request || '').trim(),
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  return writeTelemetryRecord({
    eventType: 'feedback',
    collection: 'user_feedback',
    payload: entry
  });
}

/**
 * Log an agent prompt + tool calls to Firestore for usage analytics.
 * Best-effort — never blocks the caller. Logs to 'prompt_logs' collection.
 */
async function logPromptToFirestore(data = {}) {
  const entry = {
    timestamp: new Date().toISOString(),
    username: String(data.username || 'unknown'),
    email: data.email ? normalizeEmail(data.email) : null,
    prompt: String(data.prompt || '').substring(0, 500), // truncate for storage
    active_tab: String(data.activeTab || 'unknown'),
    tools_called: Array.isArray(data.toolsCalled) ? data.toolsCalled : [],
    tool_count: Number(data.toolCount) || 0,
    model_used: String(data.modelUsed || 'unknown'),
    success: !!data.success,
    duration_ms: Number(data.durationMs) || 0,
    app_version: APP_VERSION,
    machine_id: MACHINE_ID
  };

  return writeTelemetryRecord({
    eventType: 'prompt',
    collection: 'prompt_logs',
    payload: entry
  });
}

function readFeedbackLogEntries() {
  try {
    if (!fs.existsSync(feedbackLogPath)) return [];
    const raw = fs.readFileSync(feedbackLogPath, 'utf8').trim();
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.error('Failed to read feedback log:', e.message);
    return [];
  }
}

function appendFeedbackLogEntry(entry) {
  try {
    const existing = readFeedbackLogEntries();
    existing.unshift(entry);
    // Keep file bounded to avoid unbounded growth.
    const bounded = existing.slice(0, 1000);
    fs.writeFileSync(feedbackLogPath, JSON.stringify(bounded, null, 2), 'utf8');
    return true;
  } catch (e) {
    console.error('Failed to append feedback log:', e.message);
    return false;
  }
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
        user_name TEXT,
        user_email TEXT,
        feature TEXT,
        action TEXT,
        input_count INTEGER DEFAULT 0,
        output_count INTEGER DEFAULT 0,
        duration_ms INTEGER DEFAULT 0
      )
    `);

    db.run(`
      CREATE TABLE IF NOT EXISTS user_activity_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
        username TEXT,
        email TEXT,
        action TEXT,
        app_version TEXT,
        machine_id TEXT
      )
    `);

    ensureUsageHistorySchema();

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
  const hasCsv = files.some(file => file.ext === 'csv');
  const docxFiles = files.filter(file => file.ext === 'docx');
  const redlineExts = new Set(['pdf', 'docx', 'doc', 'rtf', 'txt', 'htm', 'html', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'xls', 'xlsx', 'xlsm', 'xlsb']);
  const redlineCandidates = files.filter(file => redlineExts.has(file.ext));
  const hasDocLike = files.some(file => file.ext === 'pdf' || file.ext === 'docx' || file.ext === 'doc');

  if (/(did|does|has|have|who|when|what|where|find|search|check|show|was|were|sent|review|approved|circulated).*(email|inbox|attachment|message)|email.*(did|has|who|when|what|find|search|check)/.test(prompt)) {
    return {
      action: 'run_email_ai_search',
      target_tab: 'email',
      run_now: true,
      required_extensions: [],
      missing_requirements: [],
      user_message: 'Analyzing your emails with AI.'
    };
  }

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

  if (/update\s+checklist|checklist\s+update|refresh\s+checklist/.test(prompt)) {
    if (docxFiles.length > 0 && hasCsv) {
      return {
        action: 'run_update_checklist',
        target_tab: 'updatechecklist',
        run_now: true,
        required_extensions: ['docx', 'csv'],
        missing_requirements: [],
        user_message: 'Running Update Checklist using the attached checklist and email CSV.'
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'updatechecklist',
      run_now: false,
      required_extensions: ['docx', 'csv'],
      missing_requirements: ['Attach one checklist (.docx) and one email export (.csv).'],
      user_message: 'Opened Update Checklist. Attach a checklist and email CSV to run.'
    };
  }

  if (/generate\s+punchlist|create\s+punchlist|\bpunch\s*list\b/.test(prompt)) {
    if (docxFiles.length > 0) {
      return {
        action: 'run_generate_punchlist',
        target_tab: 'punchlist',
        run_now: true,
        required_extensions: ['docx'],
        missing_requirements: [],
        user_message: 'Generating punchlist from the attached checklist.'
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'punchlist',
      run_now: false,
      required_extensions: ['docx'],
      missing_requirements: ['Attach one checklist (.docx) to generate a punchlist.'],
      user_message: 'Opened Generate Punchlist. Attach a checklist document to run.'
    };
  }

  if (/redline|blackline|compare|comparison|markup changes|mark up changes/.test(prompt)) {
    if (redlineCandidates.length >= 2) {
      return {
        action: 'run_redline',
        target_tab: 'redline',
        run_now: true,
        required_extensions: Array.from(redlineExts),
        missing_requirements: [],
        user_message: 'Running redline with the first two compatible attachments.'
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'redline',
      run_now: false,
      required_extensions: Array.from(redlineExts),
      missing_requirements: ['Attach two compatible documents to run redline automatically.'],
      user_message: 'Opened Redline Documents.'
    };
  }

  if (/collat(e|ion)|merge comments|consolidate|combine markups/.test(prompt)) {
    if (docxFiles.length >= 2) {
      return {
        action: 'run_collate',
        target_tab: 'collate',
        run_now: true,
        required_extensions: ['docx'],
        missing_requirements: [],
        user_message: 'Running collate with the first DOCX as base and remaining DOCX files as commented versions.'
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'collate',
      run_now: false,
      required_extensions: ['docx'],
      missing_requirements: ['Attach at least two DOCX files to run collate automatically.'],
      user_message: 'Opened Collate Documents.'
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

  if (prompt.trim()) {
    return {
      action: 'run_general_llm_chat',
      target_tab: null,
      run_now: true,
      required_extensions: [],
      missing_requirements: [],
      user_message: 'Answering with Agent Mode using your selected LLM provider.'
    };
  }

  return {
    action: 'run_general_llm_chat',
    target_tab: null,
    run_now: true,
    required_extensions: [],
    missing_requirements: [],
    user_message: 'Ask a question or give a workflow command and I will route it or answer via the configured LLM.'
  };
}

function sanitizeAgentPlan(rawPlan, fallbackPlan, attachments) {
  const files = Array.isArray(attachments) ? attachments : [];
  const fallback = fallbackPlan || buildAgentFallbackPlan('', files);
  const candidate = rawPlan && typeof rawPlan === 'object' ? rawPlan : {};
  const allowedActions = new Set([
    'run_signature_packets',
    'run_packet_shell',
    'run_redline',
    'run_collate',
    'run_update_checklist',
    'run_generate_punchlist',
    'run_email_ai_search',
    'run_general_llm_chat',
    'open_tab',
    'no_op'
  ]);
  const actionValue = String(candidate.action || '').trim().toLowerCase().replace(/-/g, '_');
  const action = allowedActions.has(actionValue) ? actionValue : fallback.action;
  let targetTab = normalizeAgentTabName(candidate.target_tab) || fallback.target_tab || null;
  let runNow = Boolean(candidate.run_now);
  const docxFiles = files.filter(file => file.ext === 'docx');
  const hasCsv = files.some(file => file.ext === 'csv');
  const redlineExts = new Set(['pdf', 'docx', 'doc', 'rtf', 'txt', 'htm', 'html', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'xls', 'xlsx', 'xlsm', 'xlsb']);
  const redlineCandidates = files.filter(file => redlineExts.has(file.ext));

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
  } else if (action === 'run_redline') {
    targetTab = 'redline';
    if (redlineCandidates.length < 2) {
      runNow = false;
      missingRequirements = ['Attach two compatible documents to run redline automatically.'];
    }
  } else if (action === 'run_collate') {
    targetTab = 'collate';
    if (docxFiles.length < 2) {
      runNow = false;
      missingRequirements = ['Attach at least two DOCX files to run collate automatically.'];
    }
  } else if (action === 'run_update_checklist') {
    targetTab = 'updatechecklist';
    if (!(docxFiles.length > 0 && hasCsv)) {
      runNow = false;
      missingRequirements = ['Attach one checklist (.docx) and one email export (.csv).'];
    }
  } else if (action === 'run_generate_punchlist') {
    targetTab = 'punchlist';
    if (docxFiles.length < 1) {
      runNow = false;
      missingRequirements = ['Attach one checklist (.docx) to generate a punchlist.'];
    }
  } else if (action === 'run_email_ai_search') {
    targetTab = 'email';
  } else if (action === 'run_general_llm_chat') {
    targetTab = null;
    runNow = true;
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

  // Initialize Firebase telemetry status on startup.
  const firebaseStatus = await getFirebaseTelemetryStatus({ force: true, context: 'app_startup' });
  if (!firebaseStatus.connected) {
    const level = REQUIRE_FIREBASE_TELEMETRY ? 'error' : 'warn';
    console[level](`Firebase telemetry unavailable on startup: ${firebaseStatus.message}`);
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

    const proc = spawnTracked(processorPath, [moduleName, ...args]);
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
    const proc = spawnTracked(processorPath, [moduleName, ...args]);
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

    const proc = spawnTracked(processorPath, [moduleName, checklistPath]);
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

    const proc = spawnTracked(processorPath, [moduleName, incPath]);
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

    const proc = spawnTracked(processorPath, [moduleName, configPath]);
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

    const proc = spawnTracked(processorPath, [moduleName, ...args]);
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

    const proc = spawnTracked(processorPath, [moduleName, configPath]);
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

// ========== NATURAL LANGUAGE COMMAND TOOLS (Claude Tool Use) ==========

const COMMAND_TOOLS = [
  {
    name: 'imanage_browse',
    description: 'Open iManage file picker to browse and select documents from the document management system (DMS). Use when the user wants to open, find, or browse files in iManage.',
    input_schema: {
      type: 'object',
      properties: {
        multiple: { type: 'boolean', description: 'Allow selecting multiple files. Default false.' }
      }
    }
  },
  {
    name: 'imanage_save',
    description: 'Save or file a document to iManage. Can save as a new document or as a new version of an existing document. Opens the iManage save dialog for folder selection.',
    input_schema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Local file path of the document to save to iManage' },
        action: { type: 'string', enum: ['new_document', 'new_version'], description: 'Whether to save as a new document or as a new version of an existing document. Default: new_document' },
        show_dialog: { type: 'boolean', description: 'Show iManage save dialog for folder/profile selection. Default true.' }
      },
      required: ['file_path']
    }
  },
  {
    name: 'imanage_get_versions',
    description: 'Get the version history for a document in iManage by its profile ID.',
    input_schema: {
      type: 'object',
      properties: {
        profile_id: { type: 'string', description: 'The iManage profile ID (document number) to look up versions for' }
      },
      required: ['profile_id']
    }
  },
  {
    name: 'imanage_checkout',
    description: 'Check out a document from iManage to work on it locally.',
    input_schema: {
      type: 'object',
      properties: {
        profile_id: { type: 'string', description: 'The iManage profile ID to check out' },
        checkout_path: { type: 'string', description: 'Local folder path to check out to. If omitted, uses default checkout location.' },
        version: { type: 'string', description: 'Optional version number to check out (for example "1" for precedent V1).' }
      },
      required: ['profile_id']
    }
  },
  {
    name: 'imanage_checkin',
    description: 'Check in a document back to iManage after editing.',
    input_schema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Local file path to check in' }
      },
      required: ['file_path']
    }
  },
  {
    name: 'imanage_search',
    description: 'Search for documents in iManage by criteria like matter number, document name, or author.',
    input_schema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query — can be a matter number, document name, author, or description' }
      },
      required: ['query']
    }
  },
  {
    name: 'imanage_test_connection',
    description: 'Test iManage COM connectivity. Run this first if having connection issues. Returns diagnostic info about the COM object, login status, and available methods.',
    input_schema: {
      type: 'object',
      properties: {}
    }
  },
  {
    name: 'run_redline',
    description: 'Compare two documents to produce a redline showing differences. Supports Word, PDF, Excel, and PowerPoint.',
    input_schema: {
      type: 'object',
      properties: {
        original: { type: 'string', description: 'File path of the original document' },
        modified: { type: 'string', description: 'File path of the modified/revised document' },
        engine: { type: 'string', enum: ['auto', 'litera', 'emmaneigh'], description: 'Comparison engine. Default: auto (tries Litera first, falls back to EmmaNeigh).' }
      },
      required: ['original', 'modified']
    }
  },
  {
    name: 'run_checklist_precedent_redlines',
    description: 'For an uploaded checklist, find each document in iManage, retrieve precedent V1 and latest, then run Litera full-document redlines in batch.',
    input_schema: {
      type: 'object',
      properties: {
        checklist_path: { type: 'string', description: 'Path to checklist DOCX. If omitted, uses attached DOCX in Agent Mode.' },
        max_items: { type: 'integer', description: 'Optional cap for number of checklist items to process in one run. Default 40.' }
      }
    }
  },
  {
    name: 'navigate_tab',
    description: 'Switch to a specific tab/feature in EmmaNeigh.',
    input_schema: {
      type: 'object',
      properties: {
        tab: {
          type: 'string',
          enum: ['packets', 'packetshell', 'execution', 'sigblocks', 'timetrack', 'email', 'updatechecklist', 'punchlist', 'collate', 'redline'],
          description: 'The tab to navigate to. packets=Signature Packets, packetshell=Packet Shell, execution=Execution, sigblocks=Signature Blocks, timetrack=Time Tracking, email=Email Search, updatechecklist=Update Checklist, punchlist=Punchlist Generator, collate=Collate Comments, redline=Redline Documents'
        }
      },
      required: ['tab']
    }
  },
  {
    name: 'open_folder',
    description: 'Open a folder or file location in the system file explorer.',
    input_schema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'The folder or file path to open in explorer' }
      },
      required: ['path']
    }
  },
  {
    name: 'search_emails',
    description: 'Search loaded emails using a natural language query.',
    input_schema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Natural language search query for emails' }
      },
      required: ['query']
    }
  },
  // ===== Word COM tools =====
  {
    name: 'word_save_as_pdf',
    description: 'Convert a Word document (.docx/.doc) to PDF using Microsoft Word COM automation.',
    input_schema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Path to the Word document to convert' },
        output_path: { type: 'string', description: 'Optional output PDF path. If omitted, saves alongside the original with .pdf extension.' }
      },
      required: ['file_path']
    }
  },
  {
    name: 'word_find_replace',
    description: 'Find and replace text in a Word document using Word COM. Can replace all occurrences.',
    input_schema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Path to the Word document' },
        find_text: { type: 'string', description: 'Text to find' },
        replace_text: { type: 'string', description: 'Text to replace with' },
        save: { type: 'boolean', description: 'Save the document after replacing. Default true.' }
      },
      required: ['file_path', 'find_text', 'replace_text']
    }
  },
  {
    name: 'word_extract_text',
    description: 'Extract the full text content from a Word document.',
    input_schema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Path to the Word document' }
      },
      required: ['file_path']
    }
  },
  // ===== File system tools =====
  {
    name: 'file_copy',
    description: 'Copy a file from one location to another.',
    input_schema: {
      type: 'object',
      properties: {
        source: { type: 'string', description: 'Source file path' },
        destination: { type: 'string', description: 'Destination file path' }
      },
      required: ['source', 'destination']
    }
  },
  {
    name: 'file_move',
    description: 'Move or rename a file.',
    input_schema: {
      type: 'object',
      properties: {
        source: { type: 'string', description: 'Source file path' },
        destination: { type: 'string', description: 'Destination file path' }
      },
      required: ['source', 'destination']
    }
  },
  {
    name: 'file_list',
    description: 'List files in a directory. Returns file names, sizes, and types.',
    input_schema: {
      type: 'object',
      properties: {
        directory: { type: 'string', description: 'Directory path to list' },
        pattern: { type: 'string', description: 'Optional filter pattern (e.g. "*.docx", "*.pdf"). Default: all files.' }
      },
      required: ['directory']
    }
  },
  {
    name: 'file_create_folder',
    description: 'Create a new folder/directory.',
    input_schema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Path of the folder to create' }
      },
      required: ['path']
    }
  },
  // ===== Outlook COM tools =====
  {
    name: 'outlook_search',
    description: 'Search Outlook emails by subject, sender, date range, or body text. Returns matching emails.',
    input_schema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query — matched against subject, sender, and body' },
        folder: { type: 'string', description: 'Outlook folder to search. Default: Inbox. Examples: Inbox, Sent Items, Drafts' },
        max_results: { type: 'integer', description: 'Maximum results to return. Default: 20.' },
        days_back: { type: 'integer', description: 'Only search emails from the last N days. Default: 30.' }
      },
      required: ['query']
    }
  },
  {
    name: 'outlook_read_email',
    description: 'Read the full content of an Outlook email by its EntryID.',
    input_schema: {
      type: 'object',
      properties: {
        entry_id: { type: 'string', description: 'The Outlook EntryID of the email to read' }
      },
      required: ['entry_id']
    }
  },
  {
    name: 'outlook_save_attachments',
    description: 'Save all attachments from an Outlook email to a local folder.',
    input_schema: {
      type: 'object',
      properties: {
        entry_id: { type: 'string', description: 'The Outlook EntryID of the email' },
        save_folder: { type: 'string', description: 'Local folder to save attachments to. If omitted, uses temp folder.' }
      },
      required: ['entry_id']
    }
  },
  {
    name: 'outlook_send_email',
    description: 'Compose and send an email via Outlook. Use only when the user explicitly asks to send an email.',
    input_schema: {
      type: 'object',
      properties: {
        to: { type: 'string', description: 'Recipient email address(es), semicolon-separated' },
        cc: { type: 'string', description: 'CC recipients, semicolon-separated' },
        subject: { type: 'string', description: 'Email subject line' },
        body: { type: 'string', description: 'Email body text (HTML supported)' },
        attachments: {
          type: 'array',
          items: { type: 'string' },
          description: 'Array of local file paths to attach'
        }
      },
      required: ['to', 'subject', 'body']
    }
  }
];

// ========== iMANAGE COM INTEGRATION ==========

const IMANAGE_PS_BOOTSTRAP = `
# ── iManage COM bootstrap ────────────────────────────────────────────────
# Supports iManage Work 10.x, 9.x, FileSite, and DeskSite COM variants.
# Each function has multiple fallback paths to handle version differences.

function Get-IManageWorkObjectFactory {
  # Ordered from most common (iManage Work 10.x) to legacy (DeskSite/FileSite)
  $progIds = @(
    'iManage.COMAPILib.IManDMS',
    'iManage.Work.WorkObjectFactory',
    'Com.iManage.Work.WorkObjectFactory',
    'WorkSite.Application',
    'iManage.WorkSite.Application',
    'iManage.iwComWrapper.WorkObjectFactory',
    'iManage.WorkSiteObjects.iwComWrapper.WorkObjectFactory',
    'iManage.WorkSiteObjects.WorkObjectFactory',
    'iManage.Integrations.WorkObjectFactory'
  )
  $errors = @()
  foreach ($progId in $progIds) {
    try {
      $obj = New-Object -ComObject $progId
      if ($null -ne $obj) { return $obj }
    } catch {
      $errors += ($progId + ': ' + $_.Exception.Message)
    }
  }
  throw ('Unable to create iManage COM object. Tried: ' + ($progIds -join ', ') + '. Errors: ' + ($errors -join ' | ') + '. Ensure iManage Work Desktop is installed and running.')
}

function Invoke-IManageLogin($wof) {
  # Try multiple login method names across different COM versions
  $methodNames = @('LogIn', 'Login', 'ShowLoginDialog', 'ShowLogInDialog', 'Connect')
  $lastError = $null
  foreach ($methodName in $methodNames) {
    try {
      $method = $wof.PSObject.Methods[$methodName]
      if ($null -eq $method) { continue }
      $wof.$methodName()
      Start-Sleep -Milliseconds 500
      return $true
    } catch {
      $lastError = $_.Exception.Message
      # Try next available login method.
    }
  }
  # Also try Sessions-based login (iManage Work 10.x IManDMS pattern)
  try {
    $sessions = $wof.Sessions
    if ($null -ne $sessions -and $sessions.Count -gt 0) {
      return $true  # Already has an active session
    }
  } catch {
    # Sessions property not available on this COM version.
  }
  return $false
}

function Get-IManageHasLogin($wof, [ref]$hasMember, [ref]$hasLogin) {
  $hasMember.Value = $false
  $hasLogin.Value = $false

  # Check property-based HasLogin
  try {
    $prop = $wof.PSObject.Properties['HasLogin']
    if ($null -ne $prop) {
      $hasMember.Value = $true
      $hasLogin.Value = [bool]$wof.HasLogin
      return
    }
  } catch {
    # Continue to method fallback.
  }

  # Check method-based HasLogin()
  try {
    $method = $wof.PSObject.Methods['HasLogin']
    if ($null -ne $method) {
      $hasMember.Value = $true
      $hasLogin.Value = [bool]($wof.HasLogin())
      return
    }
  } catch {
    # No HasLogin member available.
  }

  # Check Connected property (iManage Work 10.x)
  try {
    $prop = $wof.PSObject.Properties['Connected']
    if ($null -ne $prop) {
      $hasMember.Value = $true
      $hasLogin.Value = [bool]$wof.Connected
      return
    }
  } catch {}

  # Check Sessions count (IManDMS pattern)
  try {
    $sessions = $wof.Sessions
    if ($null -ne $sessions) {
      $hasMember.Value = $true
      $hasLogin.Value = ($sessions.Count -gt 0)
      return
    }
  } catch {}
}

function Ensure-IManageLogin($wof) {
  $hasLoginMember = $false
  $hasLogin = $false
  Get-IManageHasLogin $wof ([ref]$hasLoginMember) ([ref]$hasLogin)

  # If already logged in, return immediately
  if ($hasLoginMember -and $hasLogin) {
    return
  }

  if ($hasLoginMember -and -not $hasLogin) {
    $loginAttempted = Invoke-IManageLogin $wof
    Start-Sleep -Milliseconds 300
    Get-IManageHasLogin $wof ([ref]$hasLoginMember) ([ref]$hasLogin)
    if ($hasLoginMember -and $hasLogin) {
      return
    }
    if ($hasLoginMember -and -not $hasLogin) {
      if ($loginAttempted) {
        throw "iManage login failed. Please ensure iManage Work Desktop is running and you are signed in, then retry."
      }
      throw "iManage login required but no interactive login method is available. Please sign in to iManage Work Desktop first, then retry."
    }
    return
  }

  if (-not $hasLoginMember) {
    # Some COM variants do not expose HasLogin/LogIn methods.
    # Best effort — try to login; downstream calls will surface auth errors.
    [void](Invoke-IManageLogin $wof)
    Start-Sleep -Milliseconds 300
  }
}

function New-IManageQueryFile($queryText) {
  $trimmed = [string]$queryText
  return New-Object PSObject -Property @{
    Number = $trimmed
    Name = $trimmed
    Description = $trimmed
    Author = $trimmed
    __EmmaNeighQuery = $true
  }
}

function Get-IManageComInfo($wof) {
  # Diagnostic: enumerate available methods and properties on the COM object
  $info = @{
    type = $wof.GetType().FullName
    methods = @()
    properties = @()
  }
  try {
    $info.methods = @($wof.PSObject.Methods | ForEach-Object { $_.Name } | Sort-Object -Unique)
  } catch {}
  try {
    $info.properties = @($wof.PSObject.Properties | ForEach-Object { $_.Name } | Sort-Object -Unique)
  } catch {}
  return $info
}
`;

function getIManagePowerShellHosts() {
  const hosts = [];
  if (process.platform !== 'win32') return hosts;

  const systemRoot = process.env.SystemRoot || 'C:\\Windows';
  const preferred = [
    path.join(systemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'),
    path.join(systemRoot, 'SysWOW64', 'WindowsPowerShell', 'v1.0', 'powershell.exe'),
    'powershell.exe'
  ];

  const seen = new Set();
  for (const candidate of preferred) {
    const key = String(candidate || '').toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    if (candidate.includes(path.sep)) {
      if (fs.existsSync(candidate)) hosts.push(candidate);
      continue;
    }
    hosts.push(candidate);
  }
  return hosts.length > 0 ? hosts : ['powershell.exe'];
}

function parseIManagePowerShellOutput(stdout, stderr, code, host) {
  const stdoutText = String(stdout || '').trim();
  const stderrText = String(stderr || '').trim();
  const matches = Array.from(stdoutText.matchAll(/###JSON_START###([\s\S]*?)###JSON_END###/g));
  const jsonMatch = matches.length ? matches[matches.length - 1] : null;

  if (jsonMatch) {
    const rawPayload = String(jsonMatch[1] || '').replace(/^\uFEFF/, '').trim();
    const parseCandidates = [rawPayload];
    const firstJsonToken = rawPayload.search(/[{[]/);
    if (firstJsonToken > 0) {
      parseCandidates.push(rawPayload.slice(firstJsonToken).trim());
    }

    for (const candidate of parseCandidates) {
      if (!candidate) continue;
      try {
        const parsed = JSON.parse(candidate);
        return {
          success: code === 0,
          data: parsed,
          stdout: stdoutText,
          stderr: stderrText,
          host
        };
      } catch (_) {}
    }

    return {
      success: false,
      error: 'Failed to parse PowerShell response payload.',
      stdout: stdoutText,
      stderr: stderrText,
      rawPayload,
      host
    };
  }

  if (code === 0) {
    return { success: true, data: null, stdout: stdoutText, stderr: stderrText, host };
  }

  return {
    success: false,
    error: stderrText || stdoutText || `PowerShell exited with code ${code}`,
    stdout: stdoutText,
    stderr: stderrText,
    host
  };
}

function isComClassFactoryError(result = {}) {
  const blobs = [
    result.error,
    result.stderr,
    result.stdout,
    result.rawPayload,
    result && result.data && result.data.error
  ]
    .map((part) => String(part || '').toLowerCase())
    .join('\n');
  if (!blobs) return false;
  return (
    blobs.includes('retrieving the com class factory') ||
    blobs.includes('class not registered') ||
    blobs.includes('0x80040154') ||
    blobs.includes('80040154') ||
    blobs.includes('invalid class string') ||
    blobs.includes('activex') ||
    blobs.includes('interop type') ||
    blobs.includes('queryinterface') ||
    blobs.includes('no such interface') ||
    blobs.includes('regdb_e_classnotreg') ||
    blobs.includes('co_e_classstring')
  );
}

function isIManageComFactoryError(result = {}) {
  if (isComClassFactoryError(result)) return true;
  const blobs = [
    result.error,
    result.stderr,
    result.stdout,
    result.rawPayload,
    result && result.data && result.data.error
  ]
    .map((part) => String(part || '').toLowerCase())
    .join('\n');
  return blobs.includes('unable to create imanage com object');
}

function formatIManageErrorMessage(message) {
  const text = String(message || '').trim();
  if (!text) return 'Unknown iManage error.';
  if (!isIManageComFactoryError({ error: text })) return text;
  return `iManage COM could not be loaded (CLSID/class-factory issue). EmmaNeigh attempted both 64-bit and 32-bit PowerShell hosts. Please repair or reinstall iManage Work Desktop for this Windows user, then restart and retry. Details: ${text}`;
}

function formatOfficeComErrorMessage(componentLabel, message) {
  const text = String(message || '').trim();
  if (!text) return `${componentLabel} COM error.`;
  if (!isComClassFactoryError({ error: text })) return text;
  return `${componentLabel} COM could not be loaded (CLSID/class-factory issue). EmmaNeigh attempted both 64-bit and 32-bit PowerShell hosts. Please repair or reinstall Microsoft ${componentLabel} for this Windows user, then restart and retry. Details: ${text}`;
}

function normalizeOfficeComFailure(result, componentLabel = 'Office') {
  if (!result || typeof result !== 'object') {
    return { success: false, error: `${componentLabel} COM error.` };
  }
  if (!result.success) {
    const dataError = result.data && result.data.error ? String(result.data.error) : '';
    return {
      ...result,
      success: false,
      error: formatOfficeComErrorMessage(componentLabel, dataError || result.error || result.stderr || result.stdout)
    };
  }
  if (result.data && result.data.error) {
    return { success: false, error: formatOfficeComErrorMessage(componentLabel, result.data.error) };
  }
  return null;
}

function normalizeIManageFailure(result) {
  if (!result || typeof result !== 'object') {
    return { success: false, error: 'Unknown iManage error.' };
  }
  if (!result.success) {
    const dataError = result.data && result.data.error ? String(result.data.error) : '';
    return {
      ...result,
      success: false,
      error: formatIManageErrorMessage(dataError || result.error || result.stderr || result.stdout)
    };
  }
  if (result.data && result.data.error) {
    return { success: false, error: formatIManageErrorMessage(result.data.error) };
  }
  return null;
}

function parseIManageVersionNumber(rawValue) {
  if (rawValue === null || rawValue === undefined) return null;
  const text = String(rawValue).trim();
  if (!text) return null;
  const match = text.match(/\d+/);
  if (!match) return null;
  const parsed = Number(match[0]);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeIManageProfileId(rawValue) {
  const text = String(rawValue || '').trim();
  if (!text) return '';
  const match = text.match(/\d+/g);
  if (!match || !match.length) return text;
  return match.join('');
}

function normalizeIManageArrayPayload(value, fallbackKey = 'files') {
  if (Array.isArray(value)) return value;
  if (value && typeof value === 'object') {
    const nested = value[fallbackKey];
    if (Array.isArray(nested)) return nested;
    return [value];
  }
  return [];
}

function tokenizeForIManageSearch(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .split(' ')
    .map((token) => token.trim())
    .filter((token) => token.length >= 3);
}

function scoreIManageSearchCandidate(query, file = {}) {
  const queryTokens = tokenizeForIManageSearch(query);
  const haystack = [
    file.name,
    file.description,
    file.author,
    file.extension,
    file.number
  ]
    .map((part) => String(part || '').toLowerCase())
    .join(' ');

  let score = 0;
  for (const token of queryTokens) {
    if (haystack.includes(token)) {
      score += 8;
    }
  }

  const lowerQuery = String(query || '').toLowerCase().trim();
  const lowerName = String(file.name || '').toLowerCase().trim();
  if (lowerQuery && lowerName) {
    if (lowerName === lowerQuery) score += 50;
    else if (lowerName.includes(lowerQuery)) score += 24;
    else if (lowerQuery.includes(lowerName)) score += 12;
  }

  const versionNum = parseIManageVersionNumber(file.version);
  if (versionNum !== null) score += Math.min(versionNum, 20);

  return score;
}

function pickBestIManageSearchMatch(query, files = []) {
  if (!Array.isArray(files) || files.length === 0) return null;
  let best = null;
  let bestScore = Number.NEGATIVE_INFINITY;
  for (const file of files) {
    const score = scoreIManageSearchCandidate(query, file);
    if (score > bestScore) {
      best = file;
      bestScore = score;
    }
  }
  return best ? { ...best, _score: bestScore } : null;
}

function pickIManagePrecedentAndCurrentVersions(versions = []) {
  if (!Array.isArray(versions) || versions.length === 0) return null;
  const normalized = versions
    .map((version) => {
      const numeric = parseIManageVersionNumber(version && version.version);
      return numeric === null ? null : { ...version, version_number: numeric };
    })
    .filter(Boolean)
    .sort((a, b) => a.version_number - b.version_number);

  if (!normalized.length) return null;

  const v1 = normalized.find((item) => item.version_number === 1) || normalized[0];
  const latest = normalized[normalized.length - 1];
  return { precedent: v1, latest };
}

function sanitizeFileStem(value, fallback = 'document') {
  const raw = String(value || '').trim();
  const cleaned = raw
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const safe = cleaned
    .replace(/\./g, '_')
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_-]/g, '')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return safe || fallback;
}

function listFilesRecursively(rootDir, depth = 3) {
  const files = [];
  const walk = (currentPath, remainingDepth) => {
    if (!currentPath || !fs.existsSync(currentPath) || remainingDepth < 0) return;
    let entries = [];
    try {
      entries = fs.readdirSync(currentPath, { withFileTypes: true });
    } catch (_) {
      return;
    }
    for (const entry of entries) {
      const fullPath = path.join(currentPath, entry.name);
      if (entry.isFile()) {
        files.push(fullPath);
      } else if (entry.isDirectory()) {
        walk(fullPath, remainingDepth - 1);
      }
    }
  };
  walk(rootDir, depth);
  return files;
}

function resolveIManageCheckoutFilePath(checkoutResult, checkoutDir = '', preferredName = '') {
  const candidates = [];
  const files = Array.isArray(checkoutResult && checkoutResult.files) ? checkoutResult.files : [];
  const preferredLower = String(preferredName || '').toLowerCase();

  for (const file of files) {
    const reportedPath = resolveExistingLocalPath(file && file.path);
    if (reportedPath) candidates.push(reportedPath);
    if (checkoutDir && file && file.name) {
      candidates.push(path.join(checkoutDir, String(file.name)));
    }
  }

  for (const candidate of candidates) {
    if (candidate && fs.existsSync(candidate)) {
      if (!preferredLower) return candidate;
      const base = path.basename(candidate).toLowerCase();
      if (base.includes(preferredLower)) return candidate;
    }
  }

  if (checkoutDir && fs.existsSync(checkoutDir)) {
    const discovered = listFilesRecursively(checkoutDir, 4);
    discovered.sort((a, b) => {
      try {
        return fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs;
      } catch (_) {
        return 0;
      }
    });
    if (preferredLower) {
      const preferred = discovered.find((filePath) => path.basename(filePath).toLowerCase().includes(preferredLower));
      if (preferred) return preferred;
    }
    return discovered[0] || '';
  }

  return '';
}

/**
 * Execute a PowerShell script for iManage COM operations.
 * @param {string} scriptContent — PowerShell script body (IMANAGE_PS_BOOTSTRAP is prepended)
 * @param {object} [options] — options
 * @param {number} [options.timeoutMs=60000] — max execution time before killing the process
 * @param {boolean} [options.allowInteractive=false] — if true, omits -NonInteractive flag (needed for file picker dialogs)
 */
function imanageRunPowerShell(scriptContent, options = {}) {
  const timeoutMs = Number(options.timeoutMs) || 60000;
  const allowInteractive = !!options.allowInteractive;

  return new Promise((resolve) => {
    if (process.platform !== 'win32') {
      resolve({ success: false, error: 'iManage is only available on Windows.' });
      return;
    }

    const tmpDir = app.getPath('temp');
    const scriptPath = path.join(
      tmpDir,
      `emmaneigh_imanage_${Date.now()}_${Math.floor(Math.random() * 100000)}.ps1`
    );
    fs.writeFileSync(scriptPath, `${IMANAGE_PS_BOOTSTRAP}\n${scriptContent}`, 'utf8');

    const hosts = getIManagePowerShellHosts();
    let attemptIndex = 0;
    let lastResult = null;

    const runAttempt = () => {
      const host = hosts[attemptIndex];
      const psArgs = [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-Sta',
        '-File',
        scriptPath
      ];
      // Only add -NonInteractive when dialogs aren't needed
      if (!allowInteractive) {
        psArgs.splice(1, 0, '-NonInteractive');
      }

      const proc = spawnTracked(host, psArgs, { windowsHide: !allowInteractive });

      let stdout = '';
      let stderr = '';
      let timedOut = false;
      let settled = false;

      // Timeout guard — kill process if it hangs
      const timer = setTimeout(() => {
        if (settled) return;
        timedOut = true;
        try { proc.kill('SIGTERM'); } catch (_) {}
        setTimeout(() => {
          try { proc.kill('SIGKILL'); } catch (_) {}
        }, 2000);
      }, timeoutMs);

      proc.stdout.on('data', (data) => { stdout += data.toString(); });
      proc.stderr.on('data', (data) => { stderr += data.toString(); });

      proc.on('close', (code) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);

        if (timedOut) {
          try { fs.unlinkSync(scriptPath); } catch (_) {}
          resolve({
            success: false,
            error: `iManage operation timed out after ${Math.round(timeoutMs / 1000)}s. Ensure iManage Work Desktop is running and responsive.`,
            stdout: String(stdout).trim(),
            stderr: String(stderr).trim(),
            host
          });
          return;
        }

        const parsed = parseIManagePowerShellOutput(stdout, stderr, code, host);
        lastResult = parsed;

        const shouldRetry = isComClassFactoryError(parsed);
        if (shouldRetry && attemptIndex < hosts.length - 1) {
          attemptIndex += 1;
          runAttempt();
          return;
        }

        try { fs.unlinkSync(scriptPath); } catch (_) {}

        if (hosts.length > 1) {
          if (!parsed.success) {
            parsed.error = `${parsed.error} (PowerShell hosts tried: ${hosts.join(', ')})`;
          } else if (parsed.data && parsed.data.error && isComClassFactoryError(parsed)) {
            parsed.data.error = `${parsed.data.error} (PowerShell hosts tried: ${hosts.join(', ')})`;
          }
        }
        resolve(parsed);
      });

      proc.on('error', (err) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);

        lastResult = {
          success: false,
          error: `Failed to launch PowerShell (${host}): ${err.message}`,
          host
        };
        if (attemptIndex < hosts.length - 1) {
          attemptIndex += 1;
          runAttempt();
          return;
        }
        try { fs.unlinkSync(scriptPath); } catch (_) {}
        resolve(lastResult);
      });
    };

    runAttempt();
  });
}

async function imanageBrowseFiles(multiple = false) {
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof

  $files = New-Object System.Collections.Generic.List[System.Object]

  # Try multiple method signatures for file browsing across COM versions
  $browseSuccess = $false
  $browseErrors = @()

  # Method 1: GetFiles with list ref + boolean (WorkObjectFactory pattern)
  if (-not $browseSuccess) {
    try {
      $wof.GetFiles([ref]$files, ${multiple ? '$true' : '$false'})
      $browseSuccess = $true
    } catch {
      $browseErrors += ('GetFiles(ref,bool): ' + $_.Exception.Message)
    }
  }

  # Method 2: GetFiles with list ref only
  if (-not $browseSuccess) {
    try {
      $wof.GetFiles([ref]$files)
      $browseSuccess = $true
    } catch {
      $browseErrors += ('GetFiles(ref): ' + $_.Exception.Message)
    }
  }

  # Method 3: ShowOpen / OpenFileDialog pattern (some COM variants)
  if (-not $browseSuccess) {
    try {
      $dlg = $wof.ShowOpen()
      if ($null -ne $dlg) { $files.Add($dlg) }
      $browseSuccess = $true
    } catch {
      $browseErrors += ('ShowOpen: ' + $_.Exception.Message)
    }
  }

  if (-not $browseSuccess) {
    throw ('Could not open iManage file picker. Methods tried: ' + ($browseErrors -join ' | '))
  }

  $results = @()
  foreach ($f in $files) {
    $info = @{}
    try { $info.name = $f.Name } catch { $info.name = '' }
    try { $info.number = $f.Number } catch { $info.number = '' }
    try { $info.version = $f.Version } catch { $info.version = '' }
    try { $info.extension = $f.Extension } catch { $info.extension = '' }
    try { $info.author = $f.Author } catch { $info.author = '' }
    try { $info.description = $f.Description } catch { $info.description = '' }
    try { $info.path = $f.Path } catch { $info.path = '' }
    $results += $info
  }
  $json = $results | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script, { allowInteractive: true, timeoutMs: 120000 });
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  return { success: true, files: normalizeIManageArrayPayload(result.data, 'files') };
}

async function imanageSaveDocument(filePath, action = 'new_document', showDialog = true) {
  const normalizedAction = String(action || 'new_document').trim().toLowerCase() === 'new_version'
    ? 'new_version'
    : 'new_document';
  const resolvedPath = resolveExistingLocalPath(filePath);
  if (!resolvedPath || !fs.existsSync(resolvedPath)) {
    return { success: false, error: 'Local file was not found for iManage save.' };
  }

  const showDialogBool = showDialog !== false;
  const showDialogPs = showDialogBool ? '$true' : '$false';
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  $sourcePath = '${resolvedPath.replace(/'/g, "''")}'
  if (-not (Test-Path $sourcePath)) {
    $json = @{ error = "File not found: $sourcePath" } | ConvertTo-Json -Compress
    Write-Output "###JSON_START###$json###JSON_END###"
    exit 0
  }
  ${normalizedAction === 'new_version' ? `
  # Save as new version — need profile ID from file path
  $saveErrors = @()
  $profileId = $null
  try { $profileId = $wof.GetProfileIdFromFilePath($sourcePath) } catch { $saveErrors += ('GetProfileIdFromFilePath: ' + $_.Exception.Message) }
  if ($profileId) {
    $files = New-Object System.Collections.Generic.List[System.Object]
    $entry = New-IManageQueryFile($profileId)
    $files.Add($entry)
    $saved = $false
    # Try multiple SaveAsNewVersion signatures
    if (-not $saved) { try { $wof.SaveAsNewVersion([ref]$files, $sourcePath, ${showDialogPs}); $saved = $true } catch { $saveErrors += ('SaveAsNewVersion(ref,path,dialog): ' + $_.Exception.Message) } }
    if (-not $saved) { try { $wof.SaveAsNewVersion([ref]$files, $sourcePath); $saved = $true } catch { $saveErrors += ('SaveAsNewVersion(ref,path): ' + $_.Exception.Message) } }
    if (-not $saved) { try { $wof.SaveAsNewVersion([ref]$files); $saved = $true } catch { $saveErrors += ('SaveAsNewVersion(ref): ' + $_.Exception.Message) } }
    if (-not $saved) { throw ('Could not save new version. Methods tried: ' + ($saveErrors -join ' | ')) }
    $savedFiles = @()
    foreach ($f in $files) {
      $info = @{}
      try { $info.name = $f.Name } catch {}
      try { $info.number = $f.Number } catch {}
      try { $info.version = $f.Version } catch {}
      $savedFiles += $info
    }
    $json = @{ success = $true; action = "new_version"; profileId = $profileId; files = $savedFiles } | ConvertTo-Json -Compress -Depth 4
    Write-Output "###JSON_START###$json###JSON_END###"
  } else {
    Write-Output ('###JSON_START###{"error":"Could not find iManage profile for this file (' + ($saveErrors -join '; ') + '). Try saving as new document instead."}###JSON_END###')
  }
  ` : `
  # Save as new document — try multiple method signatures
  $files = New-Object System.Collections.Generic.List[System.Object]
  $saveSuccess = $false
  $saveErrors = @()
  if (-not $saveSuccess) { try { $wof.SaveAsFiles([ref]$files, $sourcePath, ${showDialogPs}); $saveSuccess = $true } catch { $saveErrors += ('SaveAsFiles(ref,path,dialog): ' + $_.Exception.Message) } }
  if (-not $saveSuccess) { try { $wof.SaveAsFiles([ref]$files, $sourcePath); $saveSuccess = $true } catch { $saveErrors += ('SaveAsFiles(ref,path): ' + $_.Exception.Message) } }
  if (-not $saveSuccess) { try { $wof.SaveAsFiles($sourcePath); $saveSuccess = $true } catch { $saveErrors += ('SaveAsFiles(path): ' + $_.Exception.Message) } }
  if (-not $saveSuccess) { throw ('Could not save to iManage. Methods tried: ' + ($saveErrors -join ' | ')) }
  $results = @()
  foreach ($f in $files) {
    $info = @{}
    try { $info.name = $f.Name } catch {}
    try { $info.number = $f.Number } catch {}
    try { $info.version = $f.Version } catch {}
    $results += $info
  }
  $json = @{ success = $true; action = "new_document"; files = $results } | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
  `}
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script, { allowInteractive: showDialogBool, timeoutMs: 120000 });
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  const data = result.data && typeof result.data === 'object' ? result.data : {};
  if (Array.isArray(data.files)) {
    data.files = normalizeIManageArrayPayload(data.files, 'files');
  }
  return { success: true, ...data };
}

async function imanageGetVersions(profileId) {
  const normalizedProfileId = normalizeIManageProfileId(profileId);
  if (!normalizedProfileId) {
    return { success: false, error: 'A valid iManage profile ID is required.' };
  }
  const escapedId = String(normalizedProfileId).replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  $files = New-Object System.Collections.Generic.List[System.Object]
  $queryFile = New-IManageQueryFile('${escapedId}')
  $files.Add($queryFile)
  try {
    $wof.FindProfiles([ref]$files)
  } catch {
    # Continue with direct profile ID fallback.
  }
  if ($files.Count -eq 0) {
    $files.Add($queryFile)
  }
  try {
    $wof.GetAllVersions([ref]$files)
  } catch {
    $files = New-Object System.Collections.Generic.List[System.Object]
    $fallback = New-Object PSObject -Property @{ Number = '${escapedId}' }
    $files.Add($fallback)
    $wof.GetAllVersions([ref]$files)
  }
  $results = @()
  foreach ($f in $files) {
    if ($f.PSObject.Properties.Match('__EmmaNeighQuery').Count -gt 0) { continue }
    $info = @{}
    try { $info.name = $f.Name } catch { $info.name = '' }
    try { $info.number = $f.Number } catch { $info.number = '' }
    try { $info.version = $f.Version } catch { $info.version = '' }
    try { $info.author = $f.Author } catch { $info.author = '' }
    try { $info.date = $f.Date } catch { $info.date = '' }
    try { $info.description = $f.Description } catch { $info.description = '' }
    $results += $info
  }
  if ($results.Count -eq 0) {
    throw "No versions were returned by iManage for profile ${escapedId}."
  }
  $json = $results | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  return { success: true, versions: normalizeIManageArrayPayload(result.data, 'versions') };
}

async function imanageCheckout(profileId, checkoutPath, version = null) {
  const normalizedProfileId = normalizeIManageProfileId(profileId);
  if (!normalizedProfileId) {
    return { success: false, error: 'A valid iManage profile ID is required for checkout.' };
  }

  const escapedId = String(normalizedProfileId).replace(/'/g, "''");
  const normalizedCheckoutPath = normalizeLocalPath(checkoutPath || '');
  const escapedCheckoutPath = normalizedCheckoutPath ? normalizedCheckoutPath.replace(/'/g, "''") : '';
  const versionNumber = parseIManageVersionNumber(version);
  const versionBlock = versionNumber !== null
    ? `Version = ${versionNumber}`
    : '';
  const applyVersionBlock = versionNumber !== null
    ? `
  foreach ($f in $files) {
    try { $f.Version = ${versionNumber} } catch { }
  }`
    : '';
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  $files = New-Object System.Collections.Generic.List[System.Object]
  $queryFile = New-Object PSObject -Property @{
    Number = '${escapedId}'
    ${versionBlock}
  }
  $files.Add($queryFile)
  try {
    $wof.FindProfiles([ref]$files)
  } catch {
    # Continue with direct profile object fallback.
  }
  if ($files.Count -eq 0) {
    $files.Add($queryFile)
  }
  ${applyVersionBlock}
  $checkoutDir = '${escapedCheckoutPath}'
  if (-not $checkoutDir) { $checkoutDir = [System.IO.Path]::GetTempPath() }
  try {
    $wof.CheckOutFiles([ref]$files, $checkoutDir)
  } catch {
    $files = New-Object System.Collections.Generic.List[System.Object]
    $fallback = New-Object PSObject -Property @{
      Number = '${escapedId}'
      ${versionBlock}
    }
    $files.Add($fallback)
    $wof.CheckOutFiles([ref]$files, $checkoutDir)
  }
  $results = @()
  foreach ($f in $files) {
    if ($f.PSObject.Properties.Match('__EmmaNeighQuery').Count -gt 0) { continue }
    $info = @{}
    try { $info.name = $f.Name } catch { $info.name = '' }
    try { $info.number = $f.Number } catch { $info.number = '' }
    try { $info.path = $f.Path } catch { $info.path = '' }
    try { $info.version = $f.Version } catch { $info.version = '' }
    $results += $info
  }
  if ($results.Count -eq 0) {
    throw "iManage checkout returned no files for profile ${escapedId}."
  }
  $json = @{ success = $true; files = $results; requested_version = ${versionNumber === null ? '""' : versionNumber} } | ConvertTo-Json -Compress -Depth 4
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  const data = result.data && typeof result.data === 'object' ? result.data : {};
  return {
    success: true,
    ...data,
    files: normalizeIManageArrayPayload(data.files, 'files')
  };
}

async function imanageCheckin(filePath) {
  const resolvedPath = resolveExistingLocalPath(filePath);
  if (!resolvedPath || !fs.existsSync(resolvedPath)) {
    return { success: false, error: 'Local file was not found for iManage check-in.' };
  }
  const escapedPath = resolvedPath.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  $sourcePath = '${escapedPath}'
  $files = New-Object System.Collections.Generic.List[System.Object]
  $checkinSuccess = $false
  $checkinErrors = @()
  if (-not $checkinSuccess) { try { $wof.CheckInFiles([ref]$files, $sourcePath); $checkinSuccess = $true } catch { $checkinErrors += ('CheckInFiles(ref,path): ' + $_.Exception.Message) } }
  if (-not $checkinSuccess) { try { $wof.CheckInFiles($sourcePath); $checkinSuccess = $true } catch { $checkinErrors += ('CheckInFiles(path): ' + $_.Exception.Message) } }
  if (-not $checkinSuccess) { throw ('Could not check in file. Methods tried: ' + ($checkinErrors -join ' | ')) }
  $results = @()
  foreach ($f in $files) {
    $info = @{}
    try { $info.name = $f.Name } catch { $info.name = '' }
    try { $info.number = $f.Number } catch { $info.number = '' }
    try { $info.version = $f.Version } catch { $info.version = '' }
    try { $info.path = $f.Path } catch { $info.path = '' }
    $results += $info
  }
  $json = @{ success = $true; message = "File checked in successfully"; files = $results } | ConvertTo-Json -Compress -Depth 4
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  const data = result.data && typeof result.data === 'object' ? result.data : {};
  return {
    success: true,
    message: data.message || 'File checked in successfully.',
    files: normalizeIManageArrayPayload(data.files, 'files')
  };
}

async function imanageSearch(query) {
  const normalizedQuery = String(query || '').trim();
  if (!normalizedQuery) {
    return { success: false, error: 'A search query is required for iManage search.' };
  }
  const escapedQuery = normalizedQuery.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  $files = New-Object System.Collections.Generic.List[System.Object]
  $searchFile = New-IManageQueryFile('${escapedQuery}')
  $files.Add($searchFile)
  # Try multiple search method signatures
  $searchSuccess = $false
  $searchErrors = @()
  if (-not $searchSuccess) { try { $wof.FindProfiles([ref]$files); $searchSuccess = $true } catch { $searchErrors += ('FindProfiles(ref): ' + $_.Exception.Message) } }
  if (-not $searchSuccess) { try { $wof.FindProfiles($files); $searchSuccess = $true } catch { $searchErrors += ('FindProfiles(list): ' + $_.Exception.Message) } }
  if (-not $searchSuccess) { throw ('iManage search failed. Methods tried: ' + ($searchErrors -join ' | ')) }
  $results = @()
  foreach ($f in $files) {
    if ($f.PSObject.Properties.Match('__EmmaNeighQuery').Count -gt 0) { continue }
    $info = @{}
    try { $info.name = $f.Name } catch { $info.name = '' }
    try { $info.number = $f.Number } catch { $info.number = '' }
    try { $info.version = $f.Version } catch { $info.version = '' }
    try { $info.extension = $f.Extension } catch { $info.extension = '' }
    try { $info.author = $f.Author } catch { $info.author = '' }
    try { $info.description = $f.Description } catch { $info.description = '' }
    $results += $info
  }
  $json = $results | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  return { success: true, files: normalizeIManageArrayPayload(result.data, 'files') };
}

/**
 * Diagnostic: test iManage COM connectivity.
 * Returns detailed info about which COM variant was found, login status, and available methods.
 */
async function imanageTestConnection() {
  const script = `
$ErrorActionPreference = 'Stop'
$diag = @{
  platform = $env:OS
  powershell_version = $PSVersionTable.PSVersion.ToString()
  architecture = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture
  com_created = $false
  com_progid = ''
  com_type = ''
  login_status = 'unknown'
  available_methods = @()
  available_properties = @()
  error = ''
}

try {
  $diag.architecture = if ([System.IntPtr]::Size -eq 8) { '64-bit' } else { '32-bit' }
} catch {
  $diag.architecture = 'unknown'
}

# Try to create COM object
$progIds = @(
  'iManage.COMAPILib.IManDMS',
  'iManage.Work.WorkObjectFactory',
  'Com.iManage.Work.WorkObjectFactory',
  'WorkSite.Application',
  'iManage.WorkSite.Application',
  'iManage.iwComWrapper.WorkObjectFactory',
  'iManage.WorkSiteObjects.iwComWrapper.WorkObjectFactory',
  'iManage.WorkSiteObjects.WorkObjectFactory',
  'iManage.Integrations.WorkObjectFactory'
)

$comErrors = @()
$wof = $null
foreach ($progId in $progIds) {
  try {
    $wof = New-Object -ComObject $progId
    if ($null -ne $wof) {
      $diag.com_created = $true
      $diag.com_progid = $progId
      try { $diag.com_type = $wof.GetType().FullName } catch { $diag.com_type = 'unknown' }
      break
    }
  } catch {
    $comErrors += ($progId + ': ' + $_.Exception.Message)
  }
}

if (-not $diag.com_created) {
  $diag.error = 'No iManage COM object could be created. Tried: ' + ($progIds -join ', ') + '. Errors: ' + ($comErrors -join ' | ')
  $json = $diag | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
  exit 0
}

# Enumerate methods and properties
try {
  $diag.available_methods = @($wof.PSObject.Methods | ForEach-Object { $_.Name } | Sort-Object -Unique | Select-Object -First 50)
} catch {}
try {
  $diag.available_properties = @($wof.PSObject.Properties | ForEach-Object { $_.Name } | Sort-Object -Unique | Select-Object -First 50)
} catch {}

# Check login
try {
  $hasLoginMember = $false
  $hasLogin = $false

  try {
    $prop = $wof.PSObject.Properties['HasLogin']
    if ($null -ne $prop) { $hasLoginMember = $true; $hasLogin = [bool]$wof.HasLogin }
  } catch {}
  if (-not $hasLoginMember) {
    try {
      $prop = $wof.PSObject.Properties['Connected']
      if ($null -ne $prop) { $hasLoginMember = $true; $hasLogin = [bool]$wof.Connected }
    } catch {}
  }
  if (-not $hasLoginMember) {
    try {
      $sessions = $wof.Sessions
      if ($null -ne $sessions) { $hasLoginMember = $true; $hasLogin = ($sessions.Count -gt 0) }
    } catch {}
  }

  if ($hasLoginMember) {
    $diag.login_status = if ($hasLogin) { 'logged_in' } else { 'not_logged_in' }
  } else {
    $diag.login_status = 'no_login_property'
  }
} catch {
  $diag.login_status = 'error: ' + $_.Exception.Message
}

$json = $diag | ConvertTo-Json -Compress -Depth 3
Write-Output "###JSON_START###$json###JSON_END###"
  `;

  const result = await imanageRunPowerShell(script, { timeoutMs: 30000 });
  if (!result || !result.success) {
    return {
      success: false,
      error: result ? (result.error || 'Unknown error') : 'PowerShell execution failed',
      diagnostics: {
        platform: process.platform,
        powershellHosts: getIManagePowerShellHosts(),
        stderr: result ? result.stderr : ''
      }
    };
  }
  const data = result.data || {};
  return {
    success: !!data.com_created,
    diagnostics: data,
    message: data.com_created
      ? `iManage COM connected via ${data.com_progid} (${data.login_status}). ${data.available_methods.length} methods available.`
      : `iManage COM could not be loaded. ${data.error || ''}`
  };
}

async function extractChecklistDocumentNames(checklistPath) {
  const resolvedChecklistPath = resolveExistingLocalPath(checklistPath);
  if (!resolvedChecklistPath || !fs.existsSync(resolvedChecklistPath)) {
    return {
      success: false,
      error: 'Checklist file not found. Attach a valid .docx checklist and retry.'
    };
  }

  const moduleName = 'checklist_docname_extractor';
  const processorPath = getProcessorPath(moduleName);
  const usePackagedProcessor = !!(processorPath && app.isPackaged);

  if (usePackagedProcessor && !fs.existsSync(processorPath)) {
    return { success: false, error: `Processor not found: ${processorPath}` };
  }

  return new Promise((resolve) => {
    const command = usePackagedProcessor
      ? processorPath
      : (process.platform === 'win32' ? 'python' : 'python3');
    const args = usePackagedProcessor
      ? [moduleName, resolvedChecklistPath]
      : [path.join(__dirname, 'python', 'checklist_docname_extractor.py'), resolvedChecklistPath];

    const proc = spawnTracked(command, args, { windowsHide: true });
    let stdout = '';
    let stderr = '';

    proc.stdout.on('data', (data) => { stdout += data.toString(); });
    proc.stderr.on('data', (data) => { stderr += data.toString(); });

    proc.on('close', (code) => {
      if (code !== 0 && !stdout.trim()) {
        resolve({
          success: false,
          error: String(stderr || `Checklist extractor exited with code ${code}`).trim()
        });
        return;
      }

      const parsed = parseProcessorJsonOutput(stdout);
      if (!parsed || typeof parsed !== 'object') {
        resolve({
          success: false,
          error: `Failed to parse checklist extraction output.${stderr ? ` ${String(stderr).trim()}` : ''}`
        });
        return;
      }

      if (!parsed.success) {
        resolve({
          success: false,
          error: parsed.error || 'Checklist extractor failed.'
        });
        return;
      }

      const rawItems = Array.isArray(parsed.items) ? parsed.items : [];
      const items = rawItems
        .map((item) => ({
          row_id: Number.isFinite(Number(item && item.row_id)) ? Number(item.row_id) : null,
          document_name: String(item && item.document_name || '').trim(),
          row_context: String(item && item.row_context || '').trim()
        }))
        .filter((item) => item.document_name);

      resolve({
        success: true,
        count: items.length,
        items
      });
    });

    proc.on('error', (err) => {
      resolve({
        success: false,
        error: `Failed to launch checklist extractor: ${err.message}`
      });
    });
  });
}

async function runAgentChecklistPrecedentRedlines(config = {}) {
  if (process.platform !== 'win32') {
    return {
      success: false,
      error: 'This workflow requires Windows because iManage desktop integration is Windows-only.'
    };
  }

  const checklistPath = resolveExistingLocalPath(config.checklistPath || config.checklist_path || '');
  if (!checklistPath || !fs.existsSync(checklistPath)) {
    return { success: false, error: 'Checklist file not found. Attach a checklist DOCX and retry.' };
  }

  const extraction = await extractChecklistDocumentNames(checklistPath);
  if (!extraction.success) {
    return { success: false, error: extraction.error || 'Failed to parse checklist document names.' };
  }

  const rawItems = Array.isArray(extraction.items) ? extraction.items : [];
  if (!rawItems.length) {
    return {
      success: false,
      error: 'No document names were found in the checklist. Confirm the checklist has a document table.'
    };
  }

  const requestedMaxItems = Number(config.maxItems ?? config.max_items ?? 40);
  const maxItems = Number.isFinite(requestedMaxItems)
    ? Math.max(1, Math.min(Math.round(requestedMaxItems), 150))
    : 40;
  const items = rawItems.slice(0, maxItems);

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const defaultOutputRoot = path.join(
    app.getPath('downloads'),
    `EmmaNeigh_Checklist_Redlines_${timestamp}`
  );
  const outputRoot = normalizeLocalPath(config.outputFolder || config.output_folder || defaultOutputRoot);
  fs.mkdirSync(outputRoot, { recursive: true });

  const checkoutRoot = path.join(app.getPath('temp'), `emmaneigh_checklist_imanage_${Date.now()}`);
  fs.mkdirSync(checkoutRoot, { recursive: true });

  const literaPath = findLiteraInstallation();
  if (!literaPath) {
    return {
      success: false,
      error: 'Litera Compare is not installed on this machine. This workflow requires Litera for full-document redlines.'
    };
  }

  const literaOptions = {
    output_format: 'pdf',
    change_pages_only: false
  };

  const results = [];
  let successCount = 0;

  for (let i = 0; i < items.length; i += 1) {
    const item = items[i];
    const docName = String(item.document_name || '').trim();
    const resultBase = {
      row_id: item.row_id,
      document_name: docName,
      index: i + 1
    };

    if (mainWindow && mainWindow.webContents) {
      const pct = Math.min(95, 5 + Math.round((i / Math.max(items.length, 1)) * 85));
      mainWindow.webContents.send('redline-progress', {
        percent: pct,
        message: `Checklist redline ${i + 1}/${items.length}: ${docName || 'Document'}`
      });
    }

    try {
      const searchResult = await imanageSearch(docName);
      if (!searchResult.success) {
        throw new Error(searchResult.error || 'iManage search failed.');
      }

      const bestMatch = pickBestIManageSearchMatch(docName, searchResult.files || []);
      if (!bestMatch || !bestMatch.number) {
        throw new Error('No matching iManage document was found.');
      }

      const profileId = String(bestMatch.number).trim();
      const versionsResult = await imanageGetVersions(profileId);
      if (!versionsResult.success) {
        throw new Error(versionsResult.error || `Failed to retrieve versions for ${profileId}.`);
      }

      const selectedVersions = pickIManagePrecedentAndCurrentVersions(versionsResult.versions || []);
      if (!selectedVersions || !selectedVersions.precedent || !selectedVersions.latest) {
        throw new Error(`No usable versions were found for profile ${profileId}.`);
      }

      if (selectedVersions.precedent.version_number === selectedVersions.latest.version_number) {
        throw new Error(`Only one version exists for profile ${profileId}; cannot redline against V1.`);
      }

      const itemFolder = path.join(checkoutRoot, `${String(i + 1).padStart(3, '0')}_${sanitizeFileStem(docName, 'document')}`);
      const precedentFolder = path.join(itemFolder, 'precedent_v1');
      const currentFolder = path.join(itemFolder, 'current_latest');
      fs.mkdirSync(precedentFolder, { recursive: true });
      fs.mkdirSync(currentFolder, { recursive: true });

      const precedentCheckout = await imanageCheckout(profileId, precedentFolder, selectedVersions.precedent.version_number);
      if (!precedentCheckout.success) {
        throw new Error(precedentCheckout.error || `Failed to check out V${selectedVersions.precedent.version_number}.`);
      }

      const currentCheckout = await imanageCheckout(profileId, currentFolder, selectedVersions.latest.version_number);
      if (!currentCheckout.success) {
        throw new Error(currentCheckout.error || `Failed to check out V${selectedVersions.latest.version_number}.`);
      }

      const precedentPath = resolveIManageCheckoutFilePath(precedentCheckout, precedentFolder, bestMatch.name || docName);
      const currentPath = resolveIManageCheckoutFilePath(currentCheckout, currentFolder, bestMatch.name || docName);
      if (!precedentPath || !fs.existsSync(precedentPath)) {
        throw new Error('Could not locate the checked-out V1 file on disk.');
      }
      if (!currentPath || !fs.existsSync(currentPath)) {
        throw new Error('Could not locate the checked-out latest file on disk.');
      }

      const outputExt = getLiteraOutputExtension(precedentPath, literaOptions);
      const outputPath = path.join(
        outputRoot,
        buildRedlineOutputFilename(precedentPath, currentPath, outputExt)
      );

      const redlineResult = await runLiteraComparison(
        precedentPath,
        currentPath,
        outputPath,
        literaOptions
      );
      if (!redlineResult || !redlineResult.success) {
        throw new Error((redlineResult && redlineResult.error) || 'Litera redline failed.');
      }

      successCount += 1;
      results.push({
        ...resultBase,
        success: true,
        profile_id: profileId,
        profile_name: bestMatch.name || '',
        precedent_version: selectedVersions.precedent.version_number,
        current_version: selectedVersions.latest.version_number,
        output_path: redlineResult.output_path
      });
    } catch (err) {
      results.push({
        ...resultBase,
        success: false,
        error: err.message
      });
    }
  }

  const failedCount = results.length - successCount;
  const summaryMessage = `Completed checklist precedent redlines: ${successCount}/${results.length} succeeded.`;
  if (mainWindow && mainWindow.webContents) {
    mainWindow.webContents.send('redline-progress', {
      percent: 100,
      message: summaryMessage
    });
  }

  return {
    success: successCount > 0,
    output_folder: outputRoot,
    total_items: results.length,
    successful: successCount,
    failed: failedCount,
    results,
    message: summaryMessage
  };
}

// ========== WORD COM TOOLS ==========

async function wordSaveAsPdf(filePath, outputPath) {
  if (process.platform !== 'win32') return { success: false, error: 'Word COM is only available on Windows.' };
  const resolvedInputPath = resolveExistingLocalPath(filePath);
  if (!resolvedInputPath || !fs.existsSync(resolvedInputPath)) return { success: false, error: `File not found: ${filePath}` };
  const defaultPdfPath = resolvedInputPath.replace(/\.(docx?|rtf)$/i, '.pdf');
  const pdfPath = normalizeLocalPath(outputPath || defaultPdfPath);
  const escapedInput = resolvedInputPath.replace(/'/g, "''");
  const escapedOutput = pdfPath.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject "Word.Application"
  $word.Visible = $false
  $doc = $word.Documents.Open('${escapedInput}', $false, $false)
  try {
    $doc.SaveAs2('${escapedOutput}', 17) # 17 = wdFormatPDF
  } catch {
    # Fallback for older Word versions.
    $doc.SaveAs('${escapedOutput}', 17)
  }
  Write-Output '###JSON_START###{"success":true,"output":"${escapedOutput.replace(/\\/g, '\\\\')}"}###JSON_END###'
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($doc -ne $null) { $doc.Close($false) } } catch {}
  try { if ($word -ne $null) { $word.Quit() } } catch {}
  try { if ($doc -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null } } catch {}
  try { if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Word');
  if (failure) return failure;
  return { success: true, message: `Saved PDF: ${pdfPath}`, output_path: pdfPath };
}

async function wordFindReplace(filePath, findText, replaceText, save = true) {
  if (process.platform !== 'win32') return { success: false, error: 'Word COM is only available on Windows.' };
  const resolvedInputPath = resolveExistingLocalPath(filePath);
  if (!resolvedInputPath || !fs.existsSync(resolvedInputPath)) return { success: false, error: `File not found: ${filePath}` };
  const normalizedFind = String(findText || '').trim();
  if (!normalizedFind) return { success: false, error: 'find_text is required for Word find/replace.' };
  const escapedPath = resolvedInputPath.replace(/'/g, "''");
  const escapedFind = normalizedFind.replace(/'/g, "''");
  const escapedReplace = String(replaceText || '').replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject "Word.Application"
  $word.Visible = $false
  $doc = $word.Documents.Open('${escapedPath}', $false, $false)
  $range = $doc.Content
  $find = $range.Find
  $find.ClearFormatting()
  $find.Replacement.ClearFormatting()
  $find.Text = '${escapedFind}'
  $find.Replacement.Text = '${escapedReplace}'
  $find.Forward = $true
  $find.Wrap = 0  # wdFindStop
  $find.Format = $false
  $find.MatchCase = $false
  $find.MatchWholeWord = $false
  $count = 0
  while ($find.Execute()) {
    $range.Text = '${escapedReplace}'
    $count++
    $range.SetRange($range.End, $doc.Content.End)
    $find = $range.Find
    $find.Text = '${escapedFind}'
    $find.Forward = $true
    $find.Wrap = 0
  }
  ${save ? '$doc.Save()' : ''}
  Write-Output "###JSON_START###{\\"success\\":true,\\"replacements\\":$count}###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($doc -ne $null) { $doc.Close($false) } } catch {}
  try { if ($word -ne $null) { $word.Quit() } } catch {}
  try { if ($doc -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null } } catch {}
  try { if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Word');
  if (failure) return failure;
  return { success: true, message: `Replaced ${result.data?.replacements || 0} occurrence(s).`, replacements: result.data?.replacements || 0 };
}

async function wordExtractText(filePath) {
  if (process.platform !== 'win32') return { success: false, error: 'Word COM is only available on Windows.' };
  const resolvedInputPath = resolveExistingLocalPath(filePath);
  if (!resolvedInputPath || !fs.existsSync(resolvedInputPath)) return { success: false, error: `File not found: ${filePath}` };
  const escapedPath = resolvedInputPath.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject "Word.Application"
  $word.Visible = $false
  $doc = $word.Documents.Open('${escapedPath}', $false, $true) # ReadOnly
  $text = $doc.Content.Text
  $encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($text))
  Write-Output "###JSON_START###{\\"success\\":true,\\"text_b64\\":\\"$encoded\\"}###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($doc -ne $null) { $doc.Close($false) } } catch {}
  try { if ($word -ne $null) { $word.Quit() } } catch {}
  try { if ($doc -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null } } catch {}
  try { if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Word');
  if (failure) return failure;
  let text = '';
  if (result.data && result.data.text_b64) {
    text = Buffer.from(result.data.text_b64, 'base64').toString('utf8');
  }
  return { success: true, text: text.substring(0, 50000), truncated: text.length > 50000 };
}

// ========== FILE SYSTEM TOOLS ==========

function fileCopy(source, destination) {
  try {
    if (!fs.existsSync(source)) return { success: false, error: `Source not found: ${source}` };
    const destDir = path.dirname(destination);
    if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });
    fs.copyFileSync(source, destination);
    return { success: true, message: `Copied to ${destination}` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function fileMove(source, destination) {
  try {
    if (!fs.existsSync(source)) return { success: false, error: `Source not found: ${source}` };
    const destDir = path.dirname(destination);
    if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });
    fs.renameSync(source, destination);
    return { success: true, message: `Moved to ${destination}` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function fileList(directory, pattern) {
  try {
    if (!fs.existsSync(directory)) return { success: false, error: `Directory not found: ${directory}` };
    let entries = fs.readdirSync(directory, { withFileTypes: true });
    const files = entries.map(e => {
      const fullPath = path.join(directory, e.name);
      const isDir = e.isDirectory();
      let size = 0;
      try { if (!isDir) size = fs.statSync(fullPath).size; } catch (_) {}
      return { name: e.name, type: isDir ? 'folder' : path.extname(e.name).toLowerCase(), size, path: fullPath };
    });
    if (pattern) {
      const ext = pattern.replace('*', '').toLowerCase();
      return { success: true, files: files.filter(f => f.type === ext || f.name.toLowerCase().endsWith(ext)) };
    }
    return { success: true, files };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function fileCreateFolder(folderPath) {
  try {
    if (fs.existsSync(folderPath)) return { success: true, message: `Folder already exists: ${folderPath}` };
    fs.mkdirSync(folderPath, { recursive: true });
    return { success: true, message: `Created folder: ${folderPath}` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ========== OUTLOOK COM TOOLS ==========

async function outlookSearch(query, folder = 'Inbox', maxResults = 20, daysBack = 30) {
  if (process.platform !== 'win32') return { success: false, error: 'Outlook COM is only available on Windows.' };
  const normalizedQuery = String(query || '').trim();
  if (!normalizedQuery) return { success: false, error: 'query is required for Outlook search.' };
  const normalizedFolder = String(folder || 'Inbox').trim() || 'Inbox';
  const normalizedMaxResults = Math.max(1, Math.min(Math.round(Number(maxResults || 20)), 200));
  const normalizedDaysBack = Math.max(1, Math.min(Math.round(Number(daysBack || 30)), 3650));
  const escapedQuery = normalizedQuery.replace(/'/g, "''");
  const escapedFolder = normalizedFolder.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$ns = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $ns = $outlook.GetNamespace("MAPI")
  try { $null = $ns.Logon("", "", $false, $false) } catch {}
  $folderObj = $null
  $folderName = '${escapedFolder}'
  switch ($folderName) {
    'Inbox' { $folderObj = $ns.GetDefaultFolder(6) }
    'Sent Items' { $folderObj = $ns.GetDefaultFolder(5) }
    'Drafts' { $folderObj = $ns.GetDefaultFolder(16) }
    'Deleted Items' { $folderObj = $ns.GetDefaultFolder(3) }
    default { $folderObj = $ns.GetDefaultFolder(6) }
  }
  if ($null -eq $folderObj) { throw "Unable to resolve Outlook folder: $folderName" }
  $cutoffDate = (Get-Date).AddDays(-${normalizedDaysBack})
  $items = $folderObj.Items
  try { $items.Sort("[ReceivedTime]", $true) } catch {}
  $query = '${escapedQuery}'.ToLower()
  $results = @()
  foreach ($item in $items) {
    if ($results.Count -ge ${normalizedMaxResults}) { break }
    if ($null -eq $item) { continue }
    if ($item.Class -ne 43) { continue } # olMail
    $subj = if ($item.Subject) { $item.Subject } else { '' }
    $sender = if ($item.SenderName) { $item.SenderName } else { '' }
    $body = ''
    try {
      if ($item.Body) { $body = $item.Body.Substring(0, [Math]::Min(500, $item.Body.Length)) }
    } catch {}
    $msgDate = $null
    try { $msgDate = $item.ReceivedTime } catch {}
    if ($null -eq $msgDate) {
      try { $msgDate = $item.SentOn } catch {}
    }
    if ($msgDate -and $msgDate -lt $cutoffDate) { continue }
    if ($subj.ToLower().Contains($query) -or $sender.ToLower().Contains($query) -or $body.ToLower().Contains($query)) {
      $attachmentCount = 0
      try { $attachmentCount = [int]$item.Attachments.Count } catch {}
      $results += @{
        entry_id = $item.EntryID
        subject = $subj
        sender = $sender
        received = if ($msgDate) { $msgDate.ToString("yyyy-MM-dd HH:mm") } else { '' }
        has_attachments = $attachmentCount -gt 0
        attachment_count = $attachmentCount
      }
    }
  }
  $json = $results | ConvertTo-Json -Compress -Depth 3
  if (-not $json) { $json = '[]' }
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($ns -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ns) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;
  const emails = Array.isArray(result.data) ? result.data : (result.data ? [result.data] : []);
  return { success: true, emails, message: `Found ${emails.length} email(s) matching "${normalizedQuery}".` };
}

async function outlookReadEmail(entryId) {
  if (process.platform !== 'win32') return { success: false, error: 'Outlook COM is only available on Windows.' };
  const normalizedEntryId = String(entryId || '').trim();
  if (!normalizedEntryId) return { success: false, error: 'entry_id is required to read Outlook email.' };
  const escapedId = normalizedEntryId.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$ns = $null
$item = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $ns = $outlook.GetNamespace("MAPI")
  try { $null = $ns.Logon("", "", $false, $false) } catch {}
  $item = $ns.GetItemFromID('${escapedId}')
  if ($null -eq $item) { throw "Outlook item not found for EntryID." }
  if ($item.Class -ne 43) { throw "Outlook item is not an email message." }
  $attachNames = @()
  foreach ($att in $item.Attachments) { $attachNames += $att.FileName }
  $bodyText = if ($item.Body) { $item.Body.Substring(0, [Math]::Min(10000, $item.Body.Length)) } else { '' }
  $encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bodyText))
  $msgDate = $null
  try { $msgDate = $item.ReceivedTime } catch {}
  if ($null -eq $msgDate) {
    try { $msgDate = $item.SentOn } catch {}
  }
  $result = @{
    subject = $item.Subject
    sender = $item.SenderName
    sender_email = $item.SenderEmailAddress
    to = $item.To
    cc = $item.CC
    received = if ($msgDate) { $msgDate.ToString("yyyy-MM-dd HH:mm") } else { '' }
    body_b64 = $encoded
    attachments = $attachNames
  }
  $json = $result | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($item -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) | Out-Null } } catch {}
  try { if ($ns -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ns) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;
  if (result.data && result.data.body_b64) {
    result.data.body = Buffer.from(result.data.body_b64, 'base64').toString('utf8');
    delete result.data.body_b64;
  }
  return { success: true, email: result.data };
}

async function outlookSaveAttachments(entryId, saveFolder) {
  if (process.platform !== 'win32') return { success: false, error: 'Outlook COM is only available on Windows.' };
  const normalizedEntryId = String(entryId || '').trim();
  if (!normalizedEntryId) return { success: false, error: 'entry_id is required to save Outlook attachments.' };
  const folder = saveFolder || app.getPath('temp');
  if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
  const escapedId = normalizedEntryId.replace(/'/g, "''");
  const escapedFolder = folder.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$ns = $null
$item = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $ns = $outlook.GetNamespace("MAPI")
  try { $null = $ns.Logon("", "", $false, $false) } catch {}
  $item = $ns.GetItemFromID('${escapedId}')
  if ($null -eq $item) { throw "Outlook item not found for EntryID." }
  if ($item.Class -ne 43) { throw "Outlook item is not an email message." }
  $saved = @()
  foreach ($att in $item.Attachments) {
    $baseName = if ($att.FileName) { $att.FileName } else { "attachment.bin" }
    $savePath = Join-Path '${escapedFolder}' $baseName
    if (Test-Path $savePath) {
      $nameOnly = [System.IO.Path]::GetFileNameWithoutExtension($baseName)
      $extOnly = [System.IO.Path]::GetExtension($baseName)
      $suffix = (Get-Date -Format "yyyyMMdd_HHmmss_fff")
      $savePath = Join-Path '${escapedFolder}' ("$nameOnly-$suffix$extOnly")
    }
    $att.SaveAsFile($savePath)
    $saved += @{ name = $att.FileName; path = $savePath; size = $att.Size }
  }
  $json = $saved | ConvertTo-Json -Compress -Depth 3
  if (-not $json) { $json = '[]' }
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($item -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) | Out-Null } } catch {}
  try { if ($ns -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ns) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;
  const files = Array.isArray(result.data) ? result.data : (result.data ? [result.data] : []);
  return { success: true, files, message: `Saved ${files.length} attachment(s) to ${folder}` };
}

async function outlookSendEmail(to, subject, body, cc, attachments) {
  if (process.platform !== 'win32') return { success: false, error: 'Outlook COM is only available on Windows.' };
  const normalizedTo = String(to || '').trim();
  const normalizedSubject = String(subject || '').trim();
  const normalizedBody = String(body || '');
  if (!normalizedTo) return { success: false, error: 'to is required to send Outlook email.' };
  if (!normalizedSubject) return { success: false, error: 'subject is required to send Outlook email.' };
  const escapedTo = normalizedTo.replace(/'/g, "''");
  const escapedSubject = normalizedSubject.replace(/'/g, "''");
  const escapedCc = cc ? String(cc).replace(/'/g, "''") : '';
  // Body is base64 encoded to avoid escaping issues
  const bodyB64 = Buffer.from(normalizedBody, 'utf8').toString('base64');
  const rawAttachPaths = Array.isArray(attachments) ? attachments : [];
  const attachPaths = [];
  for (const attachPath of rawAttachPaths) {
    const resolvedAttachment = resolveExistingLocalPath(attachPath);
    if (!resolvedAttachment || !fs.existsSync(resolvedAttachment)) {
      return { success: false, error: `Attachment file not found: ${attachPath}` };
    }
    attachPaths.push(resolvedAttachment);
  }
  const attachLines = attachPaths.map(p => `$mail.Attachments.Add('${p.replace(/'/g, "''")}') | Out-Null`).join('\n  ');
  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$mail = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $mail = $outlook.CreateItem(0) # olMailItem
  $mail.To = '${escapedTo}'
  ${escapedCc ? `$mail.CC = '${escapedCc}'` : ''}
  $mail.Subject = '${escapedSubject}'
  $bodyBytes = [Convert]::FromBase64String('${bodyB64}')
  $mail.HTMLBody = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
  ${attachLines}
  $mail.Send()
  Write-Output '###JSON_START###{"success":true}###JSON_END###'
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($mail -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($mail) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;
  return { success: true, message: `Email sent to ${normalizedTo}` };
}

// ========== COMMAND TOOL DISPATCHER ==========

function getToolUsageEvent(toolName, input = {}, toolResult = null) {
  const result = toolResult && typeof toolResult === 'object' ? toolResult : {};
  if (toolName === 'run_checklist_precedent_redlines') {
    const total = Number(result.total_items || 0);
    const successful = Number(result.successful || 0);
    return {
      feature: 'redline',
      action: 'agent_checklist_precedent_batch',
      input_count: total > 0 ? total : 1,
      output_count: successful > 0 ? successful : 0,
      engine: 'litera'
    };
  }

  if (String(toolName || '').startsWith('imanage_')) {
    const action = String(toolName).replace('imanage_', '') || 'action';
    const outputCount = Array.isArray(result.files) ? result.files.length
      : Array.isArray(result.versions) ? result.versions.length
      : result.success ? 1 : 0;
    return {
      feature: 'imanage',
      action: `agent_${action}`,
      input_count: 1,
      output_count: outputCount
    };
  }

  return null;
}

async function dispatchTool(toolName, input, session = {}) {
  const actor = session && typeof session === 'object' ? (session.actor || {}) : {};
  const loadedFiles = session && typeof session === 'object'
    ? normalizeToolLoadedFiles(session.loadedFiles || [])
    : [];
  const safeInput = input && typeof input === 'object' ? input : {};
  const startedAt = Date.now();
  let result;
  switch (toolName) {
    case 'imanage_browse':
      result = await imanageBrowseFiles(safeInput.multiple || false);
      break;

    case 'imanage_save': {
      const filePath = resolveToolFilePath(safeInput, loadedFiles, 0);
      if (!filePath || !fs.existsSync(filePath)) {
        result = {
          success: false,
          error: 'No local file found to save. Attach a document in Agent Mode and retry.'
        };
        break;
      }
      result = await imanageSaveDocument(
        filePath,
        safeInput.action || 'new_document',
        safeInput.show_dialog !== false
      );
      break;
    }

    case 'imanage_get_versions':
      if (!safeInput.profile_id) {
        result = { success: false, error: 'profile_id is required for iManage version lookup.' };
        break;
      }
      result = await imanageGetVersions(safeInput.profile_id);
      break;

    case 'imanage_checkout':
      if (!safeInput.profile_id) {
        result = { success: false, error: 'profile_id is required for iManage checkout.' };
        break;
      }
      result = await imanageCheckout(safeInput.profile_id, safeInput.checkout_path, safeInput.version);
      break;

    case 'imanage_checkin': {
      const filePath = resolveToolFilePath(safeInput, loadedFiles, 0);
      if (!filePath || !fs.existsSync(filePath)) {
        result = {
          success: false,
          error: 'No local file found to check in. Attach a local file in Agent Mode and retry.'
        };
        break;
      }
      result = await imanageCheckin(filePath);
      break;
    }

    case 'imanage_search':
      if (!safeInput.query || !String(safeInput.query).trim()) {
        result = { success: false, error: 'query is required for iManage search.' };
        break;
      }
      result = await imanageSearch(safeInput.query);
      break;

    case 'imanage_test_connection':
      result = await imanageTestConnection();
      break;

    case 'run_redline': {
      const originalPath = resolveToolFilePath({ file_path: safeInput.original }, loadedFiles, 0);
      const modifiedPath = resolveToolFilePath({ file_path: safeInput.modified }, loadedFiles, 1);
      if (!originalPath || !modifiedPath || !fs.existsSync(originalPath) || !fs.existsSync(modifiedPath)) {
        result = {
          success: false,
          error: 'Redline needs two valid local files. Attach both documents in Agent Mode and retry.'
        };
        break;
      }
      // Delegate to existing redline-documents handler
      const engine = safeInput.engine || 'auto';
      const config = {
        engine,
        originalPath,
        modifiedPath,
        literaOptions: {}
      };
      // Send to renderer to trigger the redline
      if (mainWindow && mainWindow.webContents) {
        mainWindow.webContents.send('trigger-redline', config);
      }
      result = { success: true, message: `Redline started: comparing documents with ${engine} engine.` };
      break;
    }

    case 'run_checklist_precedent_redlines': {
      const checklistPath = resolveToolFilePath(
        { file_path: safeInput.checklist_path || safeInput.checklistPath || safeInput.path },
        loadedFiles,
        0
      );
      if (!checklistPath || !fs.existsSync(checklistPath)) {
        result = {
          success: false,
          error: 'Attach a checklist DOCX in Agent Mode (or pass checklist_path) before running batch precedent redlines.'
        };
        break;
      }

      const maxItems = Number(safeInput.max_items || safeInput.maxItems || 40);
      result = await runAgentChecklistPrecedentRedlines({
        checklistPath,
        maxItems: Number.isFinite(maxItems) ? maxItems : 40
      });

      if (result.success && result.output_folder) {
        result.message = `${result.message || 'Checklist precedent redlines completed.'} Output: ${result.output_folder}`;
      }
      break;
    }

    case 'navigate_tab':
      if (mainWindow && mainWindow.webContents) {
        mainWindow.webContents.send('navigate-tab', safeInput.tab);
      }
      result = { success: true, message: `Switched to ${safeInput.tab} tab.` };
      break;

    case 'open_folder': {
      const chosenPath = resolveToolFilePath(safeInput, loadedFiles, 0) || normalizeLocalPath(safeInput.path);
      if (!chosenPath) {
        result = { success: false, error: 'No folder or file path was provided.' };
        break;
      }
      shell.openPath(chosenPath);
      result = { success: true, message: `Opened ${chosenPath}` };
      break;
    }

    case 'search_emails':
      // Delegate to existing nl-search-emails handler — will be called from renderer with email data
      if (mainWindow && mainWindow.webContents) {
        mainWindow.webContents.send('trigger-email-search', { query: safeInput.query });
      }
      result = { success: true, message: `Email search started for: ${safeInput.query}` };
      break;

    // ── Word COM tools ──────────────────────────────────────────────
    case 'word_save_as_pdf': {
      const wordFile = resolveToolFilePath(safeInput, loadedFiles, 0) || normalizeLocalPath(safeInput.file_path);
      if (!wordFile) { result = { success: false, error: 'file_path is required' }; break; }
      const pdfOut = safeInput.output_path ? normalizeLocalPath(safeInput.output_path) : undefined;
      result = await wordSaveAsPdf(wordFile, pdfOut);
      break;
    }

    case 'word_find_replace': {
      const wordFile = resolveToolFilePath(safeInput, loadedFiles, 0) || normalizeLocalPath(safeInput.file_path);
      if (!wordFile) { result = { success: false, error: 'file_path is required' }; break; }
      if (!safeInput.find_text) { result = { success: false, error: 'find_text is required' }; break; }
      result = await wordFindReplace(wordFile, safeInput.find_text, safeInput.replace_text || '', safeInput.save !== false);
      break;
    }

    case 'word_extract_text': {
      const wordFile = resolveToolFilePath(safeInput, loadedFiles, 0) || normalizeLocalPath(safeInput.file_path);
      if (!wordFile) { result = { success: false, error: 'file_path is required' }; break; }
      result = await wordExtractText(wordFile);
      break;
    }

    // ── File system tools ───────────────────────────────────────────
    case 'file_copy': {
      const src = resolveToolFilePath(safeInput, loadedFiles, 0) || normalizeLocalPath(safeInput.source);
      const dst = normalizeLocalPath(safeInput.destination);
      if (!src || !dst) { result = { success: false, error: 'source and destination are required' }; break; }
      result = fileCopy(src, dst);
      break;
    }

    case 'file_move': {
      const src = resolveToolFilePath(safeInput, loadedFiles, 0) || normalizeLocalPath(safeInput.source);
      const dst = normalizeLocalPath(safeInput.destination);
      if (!src || !dst) { result = { success: false, error: 'source and destination are required' }; break; }
      result = fileMove(src, dst);
      break;
    }

    case 'file_list': {
      const dir = normalizeLocalPath(safeInput.directory);
      if (!dir) { result = { success: false, error: 'directory is required' }; break; }
      result = fileList(dir, safeInput.pattern);
      break;
    }

    case 'file_create_folder': {
      const fp = normalizeLocalPath(safeInput.path);
      if (!fp) { result = { success: false, error: 'path is required' }; break; }
      result = fileCreateFolder(fp);
      break;
    }

    // ── Outlook COM tools ───────────────────────────────────────────
    case 'outlook_search': {
      if (!safeInput.query) { result = { success: false, error: 'query is required' }; break; }
      result = await outlookSearch(safeInput.query, safeInput.folder, safeInput.max_results, safeInput.days_back);
      break;
    }

    case 'outlook_read_email': {
      if (!safeInput.entry_id) { result = { success: false, error: 'entry_id is required' }; break; }
      result = await outlookReadEmail(safeInput.entry_id);
      break;
    }

    case 'outlook_save_attachments': {
      if (!safeInput.entry_id) { result = { success: false, error: 'entry_id is required' }; break; }
      const saveDir = safeInput.save_folder ? normalizeLocalPath(safeInput.save_folder) : undefined;
      result = await outlookSaveAttachments(safeInput.entry_id, saveDir);
      break;
    }

    case 'outlook_send_email': {
      if (!safeInput.to || !safeInput.subject) {
        result = { success: false, error: 'to and subject are required' };
        break;
      }
      const sendBody = typeof safeInput.body === 'undefined' || safeInput.body === null
        ? ''
        : safeInput.body;
      let sendAttachments = safeInput.attachments;
      if (!Array.isArray(sendAttachments) && loadedFiles && loadedFiles.length > 0) {
        sendAttachments = loadedFiles.map((file) => file.path).filter(Boolean);
      }
      result = await outlookSendEmail(safeInput.to, safeInput.subject, sendBody, safeInput.cc, sendAttachments);
      break;
    }

    default:
      result = { success: false, error: `Unknown tool: ${toolName}` };
      break;
  }

  const usageEvent = getToolUsageEvent(toolName, input, result);
  if (usageEvent) {
    try {
      const usageLogged = await recordUsageEvent({
        user_name: actor?.username || actor?.name || actor?.displayName || 'unknown',
        user_email: actor?.email || '',
        feature: usageEvent.feature,
        action: usageEvent.action,
        input_count: usageEvent.input_count,
        output_count: usageEvent.output_count,
        duration_ms: Date.now() - startedAt,
        engine: usageEvent.engine || null
      });
      if (!usageLogged.success) {
        console.warn(`Tool usage logging failed for ${toolName}: ${usageLogged.error || 'unknown error'}`);
      }
    } catch (e) {
      console.warn(`Tool usage logging threw for ${toolName}: ${e.message}`);
    }
  }

  return result;
}

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
let literaCustomizationHintsCache = {};

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

function decodeBasicXmlEntities(value) {
  return String(value || '')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function readTextWithEncodingFallback(filePath) {
  const buffer = fs.readFileSync(filePath);
  if (!buffer || !buffer.length) return '';

  if (buffer.length >= 2) {
    const bom0 = buffer[0];
    const bom1 = buffer[1];

    // UTF-16 LE BOM
    if (bom0 === 0xFF && bom1 === 0xFE) {
      return buffer.toString('utf16le');
    }

    // UTF-16 BE BOM
    if (bom0 === 0xFE && bom1 === 0xFF) {
      const swapped = Buffer.allocUnsafe(buffer.length);
      for (let i = 0; i < buffer.length - 1; i += 2) {
        swapped[i] = buffer[i + 1];
        swapped[i + 1] = buffer[i];
      }
      if (buffer.length % 2 === 1) {
        swapped[buffer.length - 1] = buffer[buffer.length - 1];
      }
      return swapped.toString('utf16le');
    }
  }

  const utf8 = buffer.toString('utf8');
  if (utf8.includes('\u0000')) {
    return buffer.toString('utf16le');
  }
  return utf8;
}

function parseCustomizationStringValues(xmlContent) {
  const values = [];
  const pattern = /<([A-Za-z0-9_:-]+)\b[^>]*\bSTRING_VALUE=(?:"([^"]*)"|'([^']*)')[^>]*\/?>/g;
  let match = null;

  while ((match = pattern.exec(String(xmlContent || ''))) !== null) {
    const key = String(match[1] || '').trim();
    const rawValue = match[2] !== undefined ? match[2] : (match[3] || '');
    const value = decodeBasicXmlEntities(rawValue);
    if (!key || !value) continue;
    values.push({ key, value });
  }

  return values;
}

function shouldUseCustomizationKeyForType(rawKey, literaType) {
  const key = String(rawKey || '').toLowerCase();
  const hasPdfScope = key.includes('pdf');
  const hasPptScope = key.includes('ppt') || key.includes('powerpoint');
  const hasExcelScope = key.includes('excel') || key.includes('xls');

  if (literaType === 'word') {
    return !hasPdfScope && !hasPptScope && !hasExcelScope;
  }

  if (literaType === 'powerpoint') {
    if (hasPdfScope || hasExcelScope) return false;
    return hasPptScope || (!hasPptScope && !hasPdfScope && !hasExcelScope);
  }

  if (literaType === 'excel') {
    if (hasPdfScope || hasPptScope) return false;
    return hasExcelScope || (!hasExcelScope && !hasPdfScope && !hasPptScope);
  }

  return true;
}

function collectLiteraCustomizationFiles() {
  const files = [];
  const seen = new Set();
  const roots = [process.env.LOCALAPPDATA, process.env.APPDATA]
    .map(value => String(value || '').trim())
    .filter(Boolean);
  const suffixes = [
    ['Litera', 'Customize', 'customize.xml'],
    ['Litera', 'Customize', 'PPTCustomize.xml'],
    ['Litera', 'Customize', 'XLCustomize.xml'],
    ['Litera', 'Customize', 'PDFCustomize.xml'],
    ['Litera', 'Roaming_Customize', 'customize.xml'],
    ['Litera', 'Roaming_Customize', 'PPTCustomize.xml'],
    ['Litera', 'Roaming_Customize', 'XLCustomize.xml'],
    ['Litera', 'Roaming_Customize', 'PDFCustomize.xml'],
    ['Litera', 'Roaming_Customize', 'UserCustomizations.xml']
  ];

  function addIfFile(filePath) {
    const resolved = path.resolve(filePath);
    if (seen.has(resolved)) return;
    try {
      if (!fs.existsSync(resolved)) return;
      const stat = fs.statSync(resolved);
      if (!stat.isFile()) return;
      seen.add(resolved);
      files.push(resolved);
    } catch (_) {}
  }

  for (const root of roots) {
    for (const suffix of suffixes) {
      addIfFile(path.join(root, ...suffix));
    }
  }

  return files;
}

function getLiteraCustomizationHints(literaType) {
  if (process.platform !== 'win32') {
    return { styleNames: [], stylePaths: [], styleDirs: [] };
  }

  const cacheKey = String(literaType || 'default').toLowerCase();
  if (literaCustomizationHintsCache && literaCustomizationHintsCache[cacheKey]) {
    const cached = literaCustomizationHintsCache[cacheKey];
    return {
      styleNames: [...cached.styleNames],
      stylePaths: [...cached.stylePaths],
      styleDirs: [...cached.styleDirs]
    };
  }

  const styleNames = [];
  const stylePaths = [];
  const styleDirs = [];
  const seenNames = new Set();
  const seenPaths = new Set();

  function addStyleName(rawValue) {
    const value = normalizeLiteraStyleHint(rawValue);
    if (!value || hasMonochromeStyleTerm(value)) return;
    if (/[\\/]/.test(value) || /%[^%]+%/.test(value)) return;
    const key = value.toLowerCase();
    if (seenNames.has(key)) return;
    seenNames.add(key);
    styleNames.push(value);
  }

  function addStylePath(rawValue) {
    const value = normalizeLiteraStyleHint(rawValue);
    if (!value || hasMonochromeStyleTerm(value)) return;
    const key = value.toLowerCase();
    if (seenPaths.has(key)) return;
    seenPaths.add(key);
    stylePaths.push(value);
  }

  function addStyleDir(rawValue) {
    const value = normalizeLiteraStyleHint(rawValue);
    if (!value) return;
    if (!/[\\/]/.test(value) && !/%[^%]+%/.test(value)) return;
    const key = value.toLowerCase();
    if (seenPaths.has(key)) return;
    seenPaths.add(key);
    styleDirs.push(value);
  }

  for (const filePath of collectLiteraCustomizationFiles()) {
    let content = '';
    try {
      content = readTextWithEncodingFallback(filePath);
    } catch (_) {
      continue;
    }
    if (!content) continue;

    for (const { key, value } of parseCustomizationStringValues(content)) {
      if (!shouldUseCustomizationKeyForType(key, literaType)) continue;

      const keyLower = String(key || '').toLowerCase();
      const normalizedValue = normalizeLiteraStyleHint(value);
      if (!normalizedValue) continue;

      const hasStyleTerm = keyLower.includes('style');
      const hasRenderTerm = keyLower.includes('render');
      const isStyleRelated = hasStyleTerm || hasRenderTerm;
      if (!isStyleRelated && !/\.(tpx|tpp|tpz)\b/i.test(normalizedValue)) continue;

      const looksLikeStyleFile = /\.(tpx|tpp|tpz)\b/i.test(normalizedValue);
      const looksLikePath = /[\\/]/.test(normalizedValue) || /%[^%]+%/.test(normalizedValue);
      const isStyleDirHint =
        keyLower.includes('renderstylespath') ||
        keyLower.includes('renderingstylespath') ||
        keyLower.includes('corporaterenderstylespath') ||
        keyLower.includes('personalrenderstylespath') ||
        keyLower.includes('userrenderstylespath');

      if (looksLikeStyleFile) {
        addStylePath(normalizedValue);
      } else if (isStyleDirHint || (isStyleRelated && looksLikePath)) {
        addStyleDir(normalizedValue);
      } else if (isStyleRelated) {
        addStyleName(normalizedValue);
      }
    }
  }

  // Always probe the known Litera rendering-style roots as portable defaults.
  const fallbackStyleDirs = [
    '%PROGRAMDATA%\\Litera\\Change-ProRenderingStyles',
    '%PROGRAMDATA%\\Litera\\Change-ProRenderingStyles\\Corporative_Styles',
    '%PROGRAMDATA%\\Litera\\Change-ProRenderingStyles\\Corporate_Styles',
    '%PROGRAMDATA%\\Litera\\Compare\\Styles',
    '%OneDrive%\\Documents\\LiteraRenderingStyles',
    '%OneDrive%\\Documents\\LiteraRenderingStyles\\Corporative_Styles',
    '%USERPROFILE%\\Documents\\LiteraRenderingStyles'
  ];
  for (const dir of fallbackStyleDirs) {
    addStyleDir(dir);
  }

  const result = { styleNames, stylePaths, styleDirs };
  literaCustomizationHintsCache[cacheKey] = {
    styleNames: [...styleNames],
    stylePaths: [...stylePaths],
    styleDirs: [...styleDirs]
  };
  return result;
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

  const outputFormat = String((compareOptions && compareOptions.output_format) || 'pdf').toLowerCase();
  const literaType = getLiteraTypeForExtension(path.extname(originalPath || ''));
  const originalExt = path.extname(originalPath || '') || '.docx';

  if (outputFormat === 'pdf') {
    if (literaType && !['word', 'pdf'].includes(literaType)) {
      return originalExt;
    }
    return '.pdf';
  }

  if (outputFormat === 'docx') {
    if (literaType === 'word') return '.docx';
    return originalExt;
  }

  return originalExt;
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

function isWindowsAbsolutePath(filePath) {
  const normalized = String(filePath || '').trim();
  return /^[a-zA-Z]:[\\/]/.test(normalized) || /^\\\\[^\\]/.test(normalized);
}

function getLiteraStyleSearchDirectories(literaPath, extraStyleDirs = []) {
  const dirs = [];
  const seen = new Set();

  function addDir(rawPath) {
    const expanded = normalizeLiteraStyleHint(rawPath);
    if (!expanded) return;
    const normalized = expanded.replace(/[\\/]+/g, path.sep);
    const key = normalized.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    dirs.push(normalized);
  }

  addDir(literaPath);
  addDir(path.dirname(literaPath));

  const directSubdirs = [
    'Styles',
    'Style',
    'Templates',
    'Rendering Styles',
    'Comparison Styles',
    'Change-ProRenderingStyles',
    'Corporative_Styles',
    'Corporate_Styles'
  ];
  for (const subdir of directSubdirs) {
    addDir(path.join(literaPath, subdir));
  }

  const windowsRoots = [process.env.ProgramData, process.env.APPDATA, process.env.LOCALAPPDATA, process.env.USERPROFILE].filter(Boolean);
  const windowsSuffixes = [
    ['Litera', 'Compare'],
    ['Litera', 'Compare', 'Styles'],
    ['Litera', 'Compare', 'Rendering Styles'],
    ['Litera', 'Workshare', 'Compare'],
    ['Litera', 'Workshare', 'Compare', 'Styles'],
    ['Litera', 'Change-ProRenderingStyles'],
    ['Litera', 'Change-ProRenderingStyles', 'Corporative_Styles'],
    ['Litera', 'Change-ProRenderingStyles', 'Corporate_Styles'],
    ['Documents', 'LiteraRenderingStyles'],
    ['Documents', 'LiteraRenderingStyles', 'Corporative_Styles'],
    ['OneDrive', 'Documents', 'LiteraRenderingStyles']
  ];
  for (const root of windowsRoots) {
    for (const suffixParts of windowsSuffixes) {
      addDir(path.join(root, ...suffixParts));
    }
  }

  for (const extraDir of extraStyleDirs) {
    addDir(extraDir);
    addDir(path.join(extraDir, 'Corporative_Styles'));
    addDir(path.join(extraDir, 'Corporate_Styles'));
    addDir(path.join(extraDir, 'Styles'));
  }

  return dirs;
}

function resolveLiteraStylePath(renderingStyle, literaType, literaPath, options = {}) {
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

  const searchDirs = Array.isArray(options.searchDirs) ? options.searchDirs : [];
  const allSearchDirs = getLiteraStyleSearchDirectories(literaPath, searchDirs);
  const candidates = [];
  const candidateSeen = new Set();

  function addCandidate(filePath) {
    const normalized = String(filePath || '').trim();
    if (!normalized) return;
    const key = normalized.toLowerCase();
    if (candidateSeen.has(key)) return;
    candidateSeen.add(key);
    candidates.push(normalized);
  }

  const isAbsolute = path.isAbsolute(styleToken) || isWindowsAbsolutePath(styleToken);
  if (isAbsolute) {
    addCandidate(styleToken);
  } else {
    const fileNameOnly = path.basename(styleToken);
    for (const dir of allSearchDirs) {
      addCandidate(path.join(dir, styleToken));
      if (fileNameOnly !== styleToken) {
        addCandidate(path.join(dir, fileNameOnly));
      }
    }
    addCandidate(styleToken);
  }

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  console.warn(`Litera style file not found: ${styleToken}`);
  return null;
}

function findPreferredLiteraColorStyle(literaType, literaPath, options = {}) {
  const expectedExt = getLiteraStyleExtension(literaType);
  if (!expectedExt) {
    return null;
  }

  const preferredStyleName = String(options.preferredStyleName || '').trim() || null;
  const styleDirs = Array.isArray(options.styleDirs) ? options.styleDirs : [];
  const stylePaths = Array.isArray(options.stylePaths) ? options.stylePaths : [];
  const preferredFilenames = getPreferredLiteraColorStyleNames(literaType, preferredStyleName);
  const preferredSet = new Set(preferredFilenames.map(name => name.toLowerCase()));
  const preferredStyleBase = normalizeLiteraStyleName(path.parse(String(preferredStyleName || '')).name || preferredStyleName);
  const blockedTerms = ['black', 'mono', 'monochrome', 'grayscale', 'grey scale', 'gray scale', 'b&w'];
  const startDirs = getLiteraStyleSearchDirectories(literaPath, styleDirs);

  const styleFiles = [];
  const styleFileSeen = new Set();
  const visited = new Set();
  const maxDepth = 8;

  function addStyleFile(filePath) {
    const normalized = String(filePath || '').trim();
    if (!normalized) return;
    if (path.extname(normalized).toLowerCase() !== expectedExt) return;
    const key = normalized.toLowerCase();
    if (styleFileSeen.has(key)) return;
    styleFileSeen.add(key);
    styleFiles.push(normalized);
  }

  for (const hintedStylePath of stylePaths) {
    addStyleFile(resolveLiteraStylePath(hintedStylePath, literaType, literaPath, { searchDirs: styleDirs }));
  }

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
        addStyleFile(fullPath);
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
    if (baseName.includes('corporate') || baseName.includes('corporative')) score += 25;
    if (preferredStyleBase && baseName.includes('kirkland')) score += 100;
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
  const styleDirs = Array.isArray(options.styleDirs) ? options.styleDirs : [];
  const stylePaths = Array.isArray(options.stylePaths) ? options.stylePaths : [];
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

  for (const hintedStylePath of stylePaths) {
    addCandidate(resolveLiteraStylePath(hintedStylePath, literaType, literaPath, { searchDirs: styleDirs }));
    addStyleTokenCandidates(hintedStylePath);
  }

  addCandidate(findPreferredLiteraColorStyle(literaType, literaPath, {
    preferredStyleName,
    styleDirs,
    stylePaths
  }));
  if (preferredStyleName) {
    addCandidate(resolveLiteraStylePath(preferredStyleName, literaType, literaPath, { searchDirs: styleDirs }));
    addStyleTokenCandidates(preferredStyleName);
  }

  for (const styleName of getPreferredLiteraColorStyleNames(literaType, preferredStyleName)) {
    const resolved = resolveLiteraStylePath(styleName, literaType, literaPath, { searchDirs: styleDirs });
    addCandidate(resolved);
    addStyleTokenCandidates(styleName);
  }

  for (const styleHint of styleHints) {
    addCandidate(resolveLiteraStylePath(styleHint, literaType, literaPath, { searchDirs: styleDirs }));
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
    const requestedOutputFormat = String(compareOptions.output_format || 'pdf').toLowerCase();
    const normalizedOutputFormat = requestedOutputFormat === 'docx'
      ? 'docx'
      : requestedOutputFormat === 'pdf'
        ? 'pdf'
        : 'native';
    const normalizedOptions = {
      output_format: normalizedOutputFormat,
      change_pages_only: !!compareOptions.change_pages_only
    };
    const literaType = getLiteraTypeForExtension(origExt);
    let outputFormatWarning = null;

    if (!normalizedOptions.change_pages_only) {
      if (normalizedOptions.output_format === 'pdf' && literaType && !['word', 'pdf'].includes(literaType)) {
        normalizedOptions.output_format = 'native';
        outputFormatWarning = 'PDF output is unavailable for this file type; generated Litera native output instead.';
      } else if (normalizedOptions.output_format === 'docx' && literaType && literaType !== 'word') {
        normalizedOptions.output_format = 'native';
        outputFormatWarning = 'DOCX output is available only for Word comparisons; generated Litera native output instead.';
      }
    }

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
    const customizationHints = hasStyleSupport
      ? getLiteraCustomizationHints(literaExe.type)
      : { styleNames: [], stylePaths: [], styleDirs: [] };
    const registryStyleHints = hasStyleSupport ? getLiteraRegistryStyleHints() : [];
    const allStyleHints = hasStyleSupport
      ? [
          ...customizationHints.styleNames,
          ...customizationHints.stylePaths,
          ...registryStyleHints
        ]
      : [];
    const preferredStyleName = getPreferredStyleNameFromHints(allStyleHints);
    const styleCandidates = hasStyleSupport
      ? getLiteraColorStyleCandidates(literaExe.type, literaInstallPath, {
          styleHints: allStyleHints,
          styleDirs: customizationHints.styleDirs,
          stylePaths: customizationHints.stylePaths,
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
    if (customizationHints.styleNames.length || customizationHints.styleDirs.length || customizationHints.stylePaths.length) {
      console.log('Litera customization style hints:', {
        names: customizationHints.styleNames.slice(0, 5),
        files: customizationHints.stylePaths.slice(0, 5).map(value => path.basename(String(value))),
        dirs: customizationHints.styleDirs.slice(0, 5)
      });
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

      const proc = spawnTracked(literaExe.exe, args, {
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
        let resolvedWarning = warning || outputFormatWarning || null;
        if (warning && outputFormatWarning) {
          resolvedWarning = `${outputFormatWarning} ${warning}`;
        }
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

function normalizeRedlineConfigPaths(config = {}) {
  const normalized = { ...(config || {}) };
  if (!normalized.original && normalized.originalPath) {
    normalized.original = normalized.originalPath;
  }
  if (!normalized.modified && normalized.modifiedPath) {
    normalized.modified = normalized.modifiedPath;
  }
  if (normalized.original) normalized.original = resolveExistingLocalPath(normalized.original);
  if (normalized.modified) normalized.modified = resolveExistingLocalPath(normalized.modified);
  if (normalized.output) normalized.output = normalizeLocalPath(normalized.output);
  if (normalized.output_folder) normalized.output_folder = normalizeLocalPath(normalized.output_folder);
  if (Array.isArray(normalized.pairs)) {
    normalized.pairs = normalized.pairs.map((pair) => ({
      ...(pair || {}),
      original: resolveExistingLocalPath((pair && (pair.original || pair.originalPath)) || ''),
      modified: resolveExistingLocalPath((pair && (pair.modified || pair.modifiedPath)) || '')
    }));
  }
  return normalized;
}

function ensureRedlineInputFilesExist(config = {}) {
  const missing = [];
  const pushIfMissing = (filePath, label) => {
    const resolved = resolveExistingLocalPath(filePath);
    if (!resolved || !fs.existsSync(resolved)) {
      missing.push(`${label}: ${filePath || '(empty path)'}`);
    }
  };

  if (config.batch && Array.isArray(config.pairs)) {
    config.pairs.forEach((pair, index) => {
      pushIfMissing(pair && pair.original, `Pair ${index + 1} original`);
      pushIfMissing(pair && pair.modified, `Pair ${index + 1} modified`);
    });
  } else {
    pushIfMissing(config.original, 'Original document');
    pushIfMissing(config.modified, 'Modified document');
  }

  if (missing.length > 0) {
    throw new Error(`Some selected files do not exist on disk.\n${missing.join('\n')}`);
  }
}

// Redline documents — routes through Litera Compare when available, falls back to EmmaNeigh table comparison
ipcMain.handle('redline-documents', async (event, config) => {
  config = normalizeRedlineConfigPaths(config || {});
  ensureRedlineInputFilesExist(config);

  const engine = config.engine || 'auto'; // 'auto', 'litera', 'emmaneigh'
  const requestedOutputFormat = String(config.output_format || 'pdf').toLowerCase();
  const normalizedOutputFormat = requestedOutputFormat === 'docx'
    ? 'docx'
    : requestedOutputFormat === 'pdf'
      ? 'pdf'
      : 'native';
  const literaOptions = {
    output_format: normalizedOutputFormat,
    change_pages_only: !!config.change_pages_only
  };
  const requiresStrictLitera =
    literaOptions.change_pages_only || literaOptions.output_format === 'pdf' || literaOptions.output_format === 'docx';

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

    const proc = spawnTracked(processorPath, [moduleName, configPath]);
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
    const proc = spawnTracked(processorPath, [emailModuleName, configPath]);
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

    const proc = spawnTracked(processorPath, args);

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
Decide how to route the user command into one app action that EmmaNeigh can execute.

Allowed actions:
- run_signature_packets (requires PDF attachments)
- run_packet_shell (requires PDF/DOC/DOCX attachments)
- run_redline (requires two compatible document attachments)
- run_collate (requires at least two DOCX attachments)
- run_update_checklist (requires checklist DOCX + email CSV)
- run_generate_punchlist (requires checklist DOCX)
- run_email_ai_search (answers question using loaded emails or attached CSV email export)
- run_general_llm_chat (general Q&A / drafting with the selected provider)
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
  "action": "run_signature_packets|run_packet_shell|run_redline|run_collate|run_update_checklist|run_generate_punchlist|run_email_ai_search|run_general_llm_chat|open_tab|no_op",
  "target_tab": "packets|packetshell|execution|sigblocks|collate|redline|email|timetrack|updatechecklist|punchlist|null",
  "run_now": true,
  "required_extensions": ["pdf"],
  "missing_requirements": [],
  "user_message": "short one-sentence instruction for the user"
}

Rules:
- If request is clearly for signature packets and PDFs are attached, choose run_signature_packets and run_now true.
- If request is clearly for packet shell and compatible docs are attached, choose run_packet_shell and run_now true.
- If request is a direct question about emails, choose run_email_ai_search.
- If request asks to update checklist and required files are attached, choose run_update_checklist.
- If request asks to generate punchlist and checklist is attached, choose run_generate_punchlist.
- If request asks to redline/compare and two docs are attached, choose run_redline.
- If request asks to collate/consolidate markups and DOCX files are attached, choose run_collate.
- If request is a general question or drafting request that does not map to a specific workflow, choose run_general_llm_chat.
- If required files are missing, choose open_tab with run_now false and list missing_requirements.
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

ipcMain.handle('agent-general-ask', async (event, payload) => {
  try {
    const prompt = String(payload?.prompt || '').trim();
    if (!prompt) {
      return { success: false, error: 'Please enter a question for Agent Mode.' };
    }

    const aiContext = resolveAiCallContext({
      apiKey: payload?.apiKey || getApiKey(),
      requestedProvider: payload?.provider || getAIProvider()
    });
    const { apiKey, provider, providerName } = aiContext;

    if (providerRequiresApiKey(provider) && !apiKey) {
      return {
        success: false,
        error: `No API key configured. Please add your ${providerName} API key in Settings.`
      };
    }

    const attachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
    const attachmentSummary = attachments
      .map(item => String(item?.name || path.basename(String(item?.path || '')) || '').trim())
      .filter(Boolean)
      .slice(0, 20);

    const assistantPrompt = `You are EmmaNeigh Agent Mode.
You can answer user questions and help draft workflow steps.
When relevant, recommend the correct EmmaNeigh workflow (redline, collate, update checklist, punchlist, sig packets, packet shell, execution version, email search).
Keep answers concise and practical.

Attached file names (if any): ${attachmentSummary.length ? attachmentSummary.join(', ') : 'None'}

User request:
${prompt}`;

    const aiResult = await callProviderPrompt({
      provider,
      apiKey,
      prompt: assistantPrompt,
      maxTokens: 1400
    });

    if (!aiResult.success) {
      return {
        success: false,
        provider,
        providerName,
        error: aiResult.error || `Failed to get a response from ${providerName}.`
      };
    }

    return {
      success: true,
      provider,
      providerName,
      modelUsed: aiResult.modelUsed || null,
      text: String(aiResult.text || '').trim()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// ========== EXECUTE-COMMAND: Natural Language Tool Use ==========
ipcMain.handle('execute-command', async (event, { prompt, context }) => {
  const cmdStartedAt = Date.now();
  const toolsCalledLog = []; // track tool names for prompt logging
  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      return { success: false, error: 'No Anthropic API key configured. Add your API key in Settings to use natural language commands.' };
    }

    // Build context-aware system prompt
    const activeTab = context?.activeTab || 'unknown';
    const loadedFiles = context?.loadedFiles || [];
    const isWindows = process.platform === 'win32';

    const systemPrompt = `You are EmmaNeigh, a legal document processing assistant for Kirkland & Ellis.
You help users manage documents, navigate the application, and interact with iManage (the document management system).

Current state:
- Active tab: ${activeTab}
- Loaded files: ${loadedFiles.length > 0 ? loadedFiles.map(f => f.name || f).join(', ') : 'None'}
- Platform: ${process.platform} (${isWindows ? 'iManage available' : 'iManage not available — Windows only'})
- Available tabs: Signature Packets, Packet Shell, Execution, Signature Blocks, Time Tracking, Email Search, Update Checklist, Punchlist Generator, Collate Comments, Redline Documents

When the user asks you to perform an action (save to iManage, open a file, navigate, run a redline, etc.), use the appropriate tool.
You can use tools across multiple steps in one request until the task is complete.
For checklist-driven precedent workflows, you can use the batch checklist precedent redline tool.
If the user reports iManage connection issues, use imanage_test_connection first to diagnose the problem.
When they ask a question or want advice, respond with helpful text.
Keep responses concise and practical.
${!isWindows ? '\nIMPORTANT: iManage tools are only available on Windows. If the user asks for iManage features, let them know.' : ''}`;

    const session = context && typeof context === 'object'
      ? {
          actor: context.actor || {},
          loadedFiles: context.loadedFiles || []
        }
      : { actor: {}, loadedFiles: [] };

    const formatToolSummary = (toolName, toolResult) => {
      const toolLabel = String(toolName || 'tool')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, (char) => char.toUpperCase());
      if (!toolResult || typeof toolResult !== 'object') {
        return `- ${toolLabel}: no result returned.`;
      }

      if (!toolResult.success) {
        return `- ${toolLabel}: failed (${toolResult.error || 'unknown error'}).`;
      }

      if (toolResult.message) {
        return `- ${toolLabel}: ${toolResult.message}`;
      }
      if (Array.isArray(toolResult.files)) {
        return `- ${toolLabel}: ${toolResult.files.length} file(s) returned.`;
      }
      if (Array.isArray(toolResult.versions)) {
        return `- ${toolLabel}: ${toolResult.versions.length} version(s) returned.`;
      }
      if (toolResult.output_folder) {
        return `- ${toolLabel}: completed. Output folder: ${toolResult.output_folder}`;
      }
      return `- ${toolLabel}: completed.`;
    };

    // Call Claude API with iterative tool use
    const modelCandidates = [
      getClaudeModel(),
      DEFAULT_CLAUDE_MODEL,
      'claude-3-5-sonnet-20241022'
    ];

    let lastResponse = null;
    const maxToolTurns = 8;
    for (const modelName of modelCandidates) {
      let messages = [{ role: 'user', content: prompt }];
      let lastToolName = null;
      let lastToolInput = null;
      let lastToolResult = null;
      const conversationParts = [];
      let modelUnavailable = false;

      for (let turn = 0; turn < maxToolTurns; turn += 1) {
        const response = await requestHttps({
          baseUrl: 'https://api.anthropic.com',
          path: '/v1/messages',
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
          },
          body: JSON.stringify({
            model: modelName,
            max_tokens: 1200,
            system: systemPrompt,
            tools: COMMAND_TOOLS,
            messages
          }),
          timeoutMs: 30000
        });

        lastResponse = response;

        if (response.statusCode !== 200) {
          if (isLikelyModelError(response)) {
            modelUnavailable = true;
            break;
          }
          return {
            success: false,
            error: `Claude API error (${response.statusCode}): ${getErrorDetail(response)}`
          };
        }

        const content = Array.isArray(response.jsonBody?.content) ? response.jsonBody.content : [];
        const textBlocks = content.filter((part) => part && part.type === 'text' && part.text);
        if (textBlocks.length > 0) {
          const text = textBlocks.map((part) => String(part.text || '').trim()).filter(Boolean).join('\n');
          if (text) conversationParts.push(text);
        }

        const toolUses = content.filter((part) => part && part.type === 'tool_use' && part.name);
        if (!toolUses.length) {
          const finalText = conversationParts.join('\n\n').trim();
          // Best-effort prompt logging to Firebase
          logPromptToFirestore({
            username: session.actor?.username || session.actor?.name || 'unknown',
            email: session.actor?.email,
            prompt,
            activeTab,
            toolsCalled: toolsCalledLog,
            toolCount: toolsCalledLog.length,
            modelUsed: modelName,
            success: true,
            durationMs: Date.now() - cmdStartedAt
          }).catch(() => {});
          return {
            success: true,
            type: lastToolName ? 'tool_result' : 'text',
            tool: lastToolName,
            input: lastToolInput,
            toolResult: lastToolResult,
            message: finalText || 'Command received but no actionable response was generated.',
            modelUsed: modelName
          };
        }

        messages.push({ role: 'assistant', content });
        const toolResultBlocks = [];
        const toolSummaries = [];

        for (const toolUse of toolUses) {
          console.log(`[execute-command] Tool call: ${toolUse.name}`, JSON.stringify(toolUse.input));
          const toolResult = await dispatchTool(toolUse.name, toolUse.input, session);
          lastToolName = toolUse.name;
          lastToolInput = toolUse.input;
          lastToolResult = toolResult;
          toolsCalledLog.push(toolUse.name);

          let serialized = '{}';
          try {
            serialized = JSON.stringify(toolResult);
          } catch (_) {
            serialized = JSON.stringify({
              success: false,
              error: 'Tool result could not be serialized.'
            });
          }

          toolResultBlocks.push({
            type: 'tool_result',
            tool_use_id: toolUse.id,
            content: serialized,
            is_error: !toolResult || toolResult.success !== true
          });
          toolSummaries.push(formatToolSummary(toolUse.name, toolResult));
        }

        if (toolSummaries.length > 0) {
          conversationParts.push(`Executed tools:\n${toolSummaries.join('\n')}`);
        }

        messages.push({ role: 'user', content: toolResultBlocks });
      }

      if (modelUnavailable) {
        continue;
      }

      const fallbackText = conversationParts.join('\n\n').trim() || 'Command ran, but the model reached the tool-step limit before finishing.';
      // Best-effort prompt logging to Firebase
      logPromptToFirestore({
        username: session.actor?.username || session.actor?.name || 'unknown',
        email: session.actor?.email,
        prompt,
        activeTab,
        toolsCalled: toolsCalledLog,
        toolCount: toolsCalledLog.length,
        modelUsed: modelName,
        success: true,
        durationMs: Date.now() - cmdStartedAt
      }).catch(() => {});
      return {
        success: true,
        type: lastToolName ? 'tool_result' : 'text',
        tool: lastToolName,
        input: lastToolInput,
        toolResult: lastToolResult,
        message: fallbackText,
        modelUsed: modelName
      };
    }

    // Log failed prompt attempt (no model available)
    logPromptToFirestore({
      username: session.actor?.username || session.actor?.name || 'unknown',
      email: session.actor?.email,
      prompt,
      activeTab,
      toolsCalled: toolsCalledLog,
      toolCount: toolsCalledLog.length,
      modelUsed: 'none',
      success: false,
      durationMs: Date.now() - cmdStartedAt
    }).catch(() => {});
    return {
      success: false,
      error: `No supported Claude model available. Last error: ${getErrorDetail(lastResponse)}`
    };
  } catch (e) {
    console.error('[execute-command] Error:', e);
    return { success: false, error: e.message };
  }
});

ipcMain.handle('agent-checklist-precedent-redlines', async (event, payload) => {
  try {
    const startedAt = Date.now();
    const actor = payload && typeof payload === 'object' ? (payload.actor || {}) : {};
    const result = await runAgentChecklistPrecedentRedlines(payload || {});

    try {
      const usageLogged = await recordUsageEvent({
        user_name: actor?.username || actor?.name || actor?.displayName || 'unknown',
        user_email: actor?.email || '',
        feature: 'redline',
        action: 'agent_checklist_precedent_batch',
        input_count: Number(result && result.total_items ? result.total_items : 0),
        output_count: Number(result && result.successful ? result.successful : 0),
        duration_ms: Date.now() - startedAt,
        engine: 'litera'
      });
      if (!usageLogged.success) {
        console.warn(`Batch precedent redline usage logging failed: ${usageLogged.error || 'unknown error'}`);
      }
    } catch (logErr) {
      console.warn(`Batch precedent redline usage logging threw: ${logErr.message}`);
    }

    return result;
  } catch (e) {
    return {
      success: false,
      error: e.message
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

    const proc = spawnTracked(processorPath, args);
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
    const proc = spawnTracked(processorPath, [timeModuleName, configPath]);
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

// Email-only login (mandatory email identity)
ipcMain.handle('email-login', async (event, { email, displayName }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  const normalizedEmail = normalizeEmail(email);
  if (!isValidEmail(normalizedEmail)) {
    return { success: false, error: 'Please enter a valid email address' };
  }

  const safeEmail = normalizedEmail.replace(/'/g, "''");
  let isNewUser = false;

  try {
    const existing = db.exec(`
      SELECT id, username, display_name, api_key_encrypted
      FROM users
      WHERE lower(trim(email)) = '${safeEmail}'
      ORDER BY created_at ASC
      LIMIT 1
    `);

    let id;
    let username;
    let storedDisplayName;
    let apiKeyEnc;
    const resolvedDisplayName = String(displayName || '').trim();

    if (existing.length > 0 && existing[0].values.length > 0) {
      [id, username, storedDisplayName, apiKeyEnc] = existing[0].values[0];
    } else {
      isNewUser = true;
      id = crypto.randomUUID();
      const usernameSeed = buildUsernameFromEmail(normalizedEmail);
      username = getAvailableUsername(usernameSeed);
      storedDisplayName = resolvedDisplayName || usernameSeed;
      const placeholderPassword = hashPassword(crypto.randomUUID());

      db.run(`
        INSERT INTO users (id, username, email, password_hash, security_question, security_answer_hash, display_name, two_factor_enabled)
        VALUES (?, ?, ?, ?, ?, ?, ?, 0)
      `, [
        id,
        username,
        normalizedEmail,
        placeholderPassword,
        null,
        null,
        storedDisplayName
      ]);
    }

    const finalDisplayName = resolvedDisplayName || storedDisplayName || buildUsernameFromEmail(normalizedEmail);
    db.run(`
      UPDATE users
      SET email = '${safeEmail}', display_name = '${finalDisplayName.replace(/'/g, "''")}', last_login = datetime('now')
      WHERE id = '${id}'
    `);

    // Log activity (best-effort — never blocks login)
    logUserActivity(username, 'login', { email: normalizedEmail, displayName: finalDisplayName }).catch(() => {});
    if (isNewUser) {
      logUserActivity(username, 'account_created', { email: normalizedEmail, displayName: finalDisplayName }).catch(() => {});
    }

    saveDatabase();

    return {
      success: true,
      isNewUser,
      user: buildSessionUserPayload({
        id,
        username,
        displayName: finalDisplayName,
        email: normalizedEmail,
        apiKeyEnc
      })
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Create new user account
ipcMain.handle('create-user', async (event, { username, password, displayName, email, securityQuestion, securityAnswer }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username || !password) {
    return { success: false, error: 'Username and password are required' };
  }

  const normalizedEmail = normalizeEmail(email);
  if (!isValidEmail(normalizedEmail)) {
    return { success: false, error: 'A valid email address is required' };
  }

  if (!securityQuestion || !securityAnswer || !String(securityAnswer).trim()) {
    return { success: false, error: 'Security question and answer are required for password reset' };
  }

  if (password.length < 4) {
    return { success: false, error: 'Password must be at least 4 characters' };
  }

  try {
    const existingEmail = db.exec(`SELECT id FROM users WHERE email = '${normalizedEmail}'`);
    if (existingEmail.length > 0 && existingEmail[0].values.length > 0) {
      return { success: false, error: 'Email address is already in use' };
    }

    const id = crypto.randomUUID();
    const passwordHash = hashPassword(password);

    // Hash security answer if provided
    let securityAnswerHash = null;
    if (securityQuestion && securityAnswer) {
      securityAnswerHash = hashPassword(securityAnswer.toLowerCase().trim());
    }

    db.run(`INSERT INTO users (id, username, email, password_hash, security_question, security_answer_hash, display_name, two_factor_enabled) VALUES (?, ?, ?, ?, ?, ?, ?, 0)`,
      [id, username.toLowerCase().trim(), normalizedEmail, passwordHash, securityQuestion || null, securityAnswerHash, displayName || username]);
    saveDatabase();

    const activityLogged = await logUserActivity(username.toLowerCase().trim(), 'account_created', {
      email: normalizedEmail,
      displayName: displayName || username
    });
    if (!activityLogged) {
      if (telemetryWriteIsMandatory()) {
        db.run(`DELETE FROM users WHERE id = '${id}'`);
        saveDatabase();
        return { success: false, error: 'Firebase telemetry write failed during account creation. Please retry.' };
      }
      console.warn('Telemetry write failed during account creation; continuing because telemetry is not mandatory.');
    }

    return { success: true, userId: id };
  } catch (e) {
    if (e.message && e.message.includes('UNIQUE')) {
      return { success: false, error: 'Username already exists' };
    }
    return { success: false, error: e.message };
  }
});

// Login with username/password
ipcMain.handle('login-user', async (event, { username, password, email }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  if (!username || !password || !email) {
    return { success: false, error: 'Username, password, and email are required' };
  }

  const normalizedEmail = normalizeEmail(email);
  if (!isValidEmail(normalizedEmail)) {
    return { success: false, error: 'Please enter a valid email address' };
  }

  try {
    const result = db.exec(`SELECT id, username, password_hash, display_name, api_key_encrypted, email FROM users WHERE username = '${username.toLowerCase().trim()}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const row = result[0].values[0];
    const [id, uname, passwordHash, displayName, apiKeyEnc, storedEmail] = row;

    if (!verifyPassword(password, passwordHash)) {
      return { success: false, error: 'Invalid password' };
    }

    const currentEmail = normalizeEmail(storedEmail || '');
    if (currentEmail && currentEmail !== normalizedEmail) {
      return { success: false, error: 'Email does not match the email on this account' };
    }

    if (!currentEmail) {
      db.run(`UPDATE users SET email = '${normalizedEmail}' WHERE id = '${id}'`);
    }

    const activityLogged = await logUserActivity(uname, 'login', {
      email: normalizedEmail,
      displayName: displayName || uname
    });
    if (!activityLogged) {
      if (telemetryWriteIsMandatory()) {
        return { success: false, error: 'Firebase telemetry write failed during login. Please retry.' };
      }
      console.warn('Telemetry write failed during login; continuing because telemetry is not mandatory.');
    }

    db.run(`UPDATE users SET last_login = datetime('now') WHERE id = '${id}'`);
    saveDatabase();

    return {
      success: true,
      user: buildSessionUserPayload({
        id,
        username: uname,
        displayName: displayName || uname,
        email: normalizedEmail,
        apiKeyEnc
      })
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Complete login after 2FA verification
ipcMain.handle('complete-2fa-login', async (event, { userId, code }) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  const telemetryGate = await requireFirebaseTelemetry('2FA login');
  if (!telemetryGate.success) return telemetryGate;

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
    const userResult = db.exec(`SELECT id, username, display_name, api_key_encrypted, email FROM users WHERE id = '${userId}'`);

    if (userResult.length === 0 || userResult[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const [id, username, displayName, apiKeyEnc, email] = userResult[0].values[0];
    const normalizedEmail = normalizeEmail(email || '');

    const activityLogged = await logUserActivity(username, 'login', {
      email: normalizedEmail,
      displayName: displayName || username
    });
    if (!activityLogged) {
      if (telemetryWriteIsMandatory()) {
        return { success: false, error: 'Firebase telemetry write failed during login. Please retry.' };
      }
      console.warn('Telemetry write failed during 2FA login; continuing because telemetry is not mandatory.');
    }

    // Update last login
    db.run(`UPDATE users SET last_login = datetime('now') WHERE id = '${id}'`);
    saveDatabase();

    return {
      success: true,
      user: buildSessionUserPayload({
        id,
        username,
        displayName: displayName || username,
        email: normalizedEmail || null,
        apiKeyEnc
      })
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get user statistics (for admin tracking)
ipcMain.handle('can-access-user-stats', async (event, requester = {}) => {
  return { success: true, allowed: canRequesterAccessAnalytics(requester) };
});

ipcMain.handle('get-user-stats', async (event, requester = {}) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  if (!canRequesterAccessAnalytics(requester)) {
    return { success: false, error: 'Not authorized to view user analytics.' };
  }

  try {
    const usersResult = db.exec(`
      SELECT username, display_name, email, last_login, created_at
      FROM users
      WHERE email IS NOT NULL AND TRIM(email) != ''
      ORDER BY datetime(last_login) DESC, datetime(created_at) DESC
    `);

    const uniqueByEmail = new Map();
    const rows = usersResult.length > 0 ? usersResult[0].values : [];
    for (const row of rows) {
      const normalizedEmail = normalizeEmail(row[2] || '');
      if (!isValidEmail(normalizedEmail)) continue;
      if (uniqueByEmail.has(normalizedEmail)) continue;
      uniqueByEmail.set(normalizedEmail, {
        username: row[0],
        displayName: row[1],
        email: normalizedEmail,
        lastLogin: row[3],
        createdAt: row[4]
      });
    }

    const uniqueUsers = Array.from(uniqueByEmail.values());
    const now = Date.now();
    const dayMs = 24 * 60 * 60 * 1000;
    const weeklyActive = uniqueUsers.filter((user) => {
      const parsed = parseSqliteDate(user.lastLogin);
      return !!parsed && (now - parsed.getTime()) <= (7 * dayMs);
    }).length;
    const activeUsers = uniqueUsers.filter((user) => {
      const parsed = parseSqliteDate(user.lastLogin);
      return !!parsed && (now - parsed.getTime()) <= (30 * dayMs);
    }).length;

    const users = uniqueUsers.map((user) => ({
      username: user.username,
      displayName: user.displayName,
      email: user.email ? user.email.replace(/(.{2})(.*)(@.*)/, '$1***$3') : null,
      lastLogin: user.lastLogin,
      createdAt: user.createdAt
    }));

    const featuresResult = db.exec(`
      SELECT feature, COUNT(*) as usage_count
      FROM usage_history
      WHERE feature IS NOT NULL AND TRIM(feature) != ''
      GROUP BY feature
      ORDER BY usage_count DESC
    `);
    const aggregatedFeatureUsage = new Map();
    if (featuresResult.length > 0) {
      for (const row of featuresResult[0].values) {
        const label = categorizeFeatureLabel(row[0]);
        const count = Number(row[1] || 0);
        aggregatedFeatureUsage.set(label, (aggregatedFeatureUsage.get(label) || 0) + count);
      }
    }
    const featureUsage = Array.from(aggregatedFeatureUsage.entries())
      .map(([feature, count]) => ({ feature, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 12);

    const loginResult = db.exec(`
      SELECT COUNT(*)
      FROM user_activity_history
      WHERE action = 'login'
    `);
    const totalLoginEvents = loginResult.length > 0 ? Number(loginResult[0].values[0][0] || 0) : 0;

    return {
      success: true,
      totalUsers: uniqueUsers.length,
      activeUsers,
      weeklyActive,
      users,
      featureUsage,
      totalLoginEvents
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get path to pending logs (for debugging)
ipcMain.handle('get-activity-log-path', async () => {
  return { success: true, path: pendingLogsPath, note: 'Activity logs are centralized via Firebase Firestore' };
});

ipcMain.handle('open-analytics-storage', async (event, requester = {}) => {
  try {
    if (!canRequesterAccessAnalytics(requester)) {
      return { success: false, error: 'Not authorized to open analytics storage.' };
    }
    const analyticsDir = path.dirname(historyDbPath);
    if (!fs.existsSync(analyticsDir)) {
      fs.mkdirSync(analyticsDir, { recursive: true });
    }
    const openResult = await shell.openPath(analyticsDir);
    if (openResult) {
      return { success: false, error: openResult, path: analyticsDir };
    }
    return {
      success: true,
      path: analyticsDir,
      databasePath: historyDbPath,
      feedbackPath: feedbackLogPath
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('cancel-active-operations', async () => {
  try {
    const entries = Array.from(trackedChildProcesses.values());
    let terminated = 0;
    const commands = [];
    for (const entry of entries) {
      if (!entry || !entry.proc) continue;
      if (terminateTrackedProcess(entry)) {
        terminated += 1;
        commands.push(entry.command || 'process');
      }
    }
    return {
      success: true,
      terminated,
      activeBefore: entries.length,
      commands
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
    const result = db.exec(`SELECT id, username, display_name, api_key_encrypted, email FROM users WHERE id = '${userId}'`);

    if (result.length === 0 || result[0].values.length === 0) {
      return { success: false, error: 'User not found' };
    }

    const row = result[0].values[0];
    const [id, uname, displayName, apiKeyEnc, email] = row;
    const normalizedEmail = normalizeEmail(email || '');

    return {
      success: true,
      user: buildSessionUserPayload({
        id,
        username: uname,
        displayName: displayName || uname,
        email: normalizedEmail || null,
        apiKeyEnc
      })
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

  const normalizedEmail = normalizeEmail(email);
  if (!isValidEmail(normalizedEmail)) {
    return { success: false, error: 'Invalid email address' };
  }

  try {
    // Check if email is already used by another user
    const existing = db.exec(`SELECT id FROM users WHERE email = '${normalizedEmail}' AND id != '${userId}'`);
    if (existing.length > 0 && existing[0].values.length > 0) {
      return { success: false, error: 'Email address is already in use' };
    }

    db.run(`UPDATE users SET email = '${normalizedEmail}', email_verified = 0, two_factor_enabled = 0 WHERE id = '${userId}'`);

    const userResult = db.exec(`SELECT username, display_name FROM users WHERE id = '${userId}'`);
    if (userResult.length > 0 && userResult[0].values.length > 0) {
      const [username, displayName] = userResult[0].values[0];
      await logUserActivity(username, 'email_updated', {
        email: normalizedEmail,
        displayName: displayName || username
      });
    }

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
    if (key === TELEMETRY_INGEST_KEY || key === TELEMETRY_INGEST_TOKEN_KEY) {
      invalidateTelemetryStatusCache();
    }
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

ipcMain.handle('get-telemetry-ingest-config', async () => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    const url = getTelemetryIngestUrl();
    const token = getTelemetryIngestToken();
    return {
      success: true,
      url,
      hasToken: !!token
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('save-telemetry-ingest-config', async (event, config = {}) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    const rawUrl = String(config.url || '').trim();
    const normalizedUrl = normalizeIngestUrl(rawUrl);
    if (!normalizedUrl) {
      return { success: false, error: 'Telemetry ingest URL must be a valid http(s) URL.' };
    }

    const rawToken = typeof config.token === 'string' ? config.token.trim() : '';
    const keepExistingToken = !!config.keepExistingToken;
    if (!writeSettingValue(TELEMETRY_INGEST_KEY, normalizedUrl)) {
      return { success: false, error: 'Failed to save telemetry ingest URL.' };
    }

    if (rawToken) {
      if (!writeSettingValue(TELEMETRY_INGEST_TOKEN_KEY, encodeSecretForSetting(rawToken))) {
        return { success: false, error: 'Failed to save telemetry ingest token.' };
      }
    } else if (!keepExistingToken) {
      if (!writeSettingValue(TELEMETRY_INGEST_TOKEN_KEY, '')) {
        return { success: false, error: 'Failed to clear telemetry ingest token.' };
      }
    }

    saveDatabase();
    invalidateTelemetryStatusCache();

    const status = await getFirebaseTelemetryStatus({
      force: true,
      context: 'save_ingest_config',
      mode: 'backend'
    });
    if (!status.connected) {
      return {
        success: true,
        status,
        warning: status.message || 'Ingest config saved, but connectivity test failed.'
      };
    }

    return { success: true, status };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('test-telemetry-ingest-config', async (event, config = {}) => {
  try {
    const url = normalizeIngestUrl(config.url || getTelemetryIngestUrl());
    if (!url) {
      return { success: false, error: 'Telemetry ingest URL must be set before testing.' };
    }

    const tokenInput = typeof config.token === 'string' ? config.token.trim() : '';
    const token = tokenInput || getTelemetryIngestToken();
    const result = await sendTelemetryToIngestBackend({
      eventType: 'telemetry_probe',
      collection: FIREBASE_TELEMETRY_PROBE_COLLECTION,
      payload: { test: true },
      context: 'manual_ingest_test',
      urlOverride: url,
      tokenOverride: token
    });

    if (!result.success) {
      return { success: false, error: result.error || 'Telemetry backend test failed.' };
    }

    return { success: true, message: 'Telemetry backend ingest connection successful.' };
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
      invalidateTelemetryStatusCache();
      const status = await getFirebaseTelemetryStatus({
        force: true,
        context: 'save_firebase_config',
        mode: 'direct'
      });
      if (status.connected) {
        return { success: true, initialized: true, status };
      }
      return { success: false, initialized: false, status, error: status.message };
    }
    return { success: false, error: 'Failed to save config file' };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Get current Firebase config (for settings UI)
ipcMain.handle('get-firebase-config', async () => {
  const config = loadFirebaseConfig();
  const status = await getFirebaseTelemetryStatus({ skipProbe: true });
  return { success: true, config: config, configured: !!config, status };
});

ipcMain.handle('get-firebase-telemetry-status', async (event, options = {}) => {
  const force = !!options.force;
  const context = String(options.context || 'status_request');
  const status = await getFirebaseTelemetryStatus({ force, context });
  return { success: true, status };
});

// ========== USAGE HISTORY HANDLERS ==========

// Log usage event — writes to both local SQLite and Firebase Firestore
ipcMain.handle('log-usage', async (event, data = {}) => {
  const result = await recordUsageEvent(data);
  if (!result.success) return result;
  return {
    success: true,
    localLogged: !!result.localLogged,
    telemetryLogged: !!result.telemetryLogged
  };
});

ipcMain.handle('submit-feedback', async (event, data) => {
  try {
    const telemetryGate = await requireFirebaseTelemetry('feedback logging');
    if (!telemetryGate.success) return telemetryGate;

    const username = String(data?.username || '').trim();
    const request = String(data?.request || '').trim();
    if (!username) {
      return { success: false, error: 'Username is required.' };
    }
    if (!request) {
      return { success: false, error: 'Feedback cannot be empty.' };
    }
    const entry = {
      timestamp: new Date().toISOString(),
      username,
      request,
      app_version: APP_VERSION,
      machine_id: MACHINE_ID
    };
    const localSaved = appendFeedbackLogEntry(entry);
    if (!localSaved) {
      return { success: false, error: 'Failed to save feedback log locally.' };
    }
    const feedbackLogged = await logFeedbackToFirestore(entry);
    if (!feedbackLogged && telemetryWriteIsMandatory()) {
      return { success: false, error: 'Firebase telemetry write failed while submitting feedback.' };
    } else if (!feedbackLogged) {
      console.warn('Telemetry write failed for feedback; continuing because telemetry is not mandatory.');
    }
    return { success: true, path: feedbackLogPath, telemetryLogged: !!feedbackLogged };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('get-feedback', async (event, options = {}) => {
  try {
    const username = String(options?.username || '').trim();
    if (!canAccessFeedbackLog(username)) {
      return { success: false, error: 'Not authorized to view feedback log.', entries: [], path: feedbackLogPath };
    }
    const requestedLimit = Number(options?.limit);
    const limit = Number.isFinite(requestedLimit)
      ? Math.max(1, Math.min(200, Math.floor(requestedLimit)))
      : 25;
    const entries = readFeedbackLogEntries().slice(0, limit);
    return { success: true, entries, path: feedbackLogPath };
  } catch (e) {
    return { success: false, error: e.message, entries: [], path: feedbackLogPath };
  }
});

ipcMain.handle('open-feedback-log', async (event, data = {}) => {
  try {
    const username = String(data?.username || '').trim();
    if (!canAccessFeedbackLog(username)) {
      return { success: false, error: 'Not authorized to open feedback log.' };
    }
    if (!fs.existsSync(feedbackLogPath)) {
      fs.writeFileSync(feedbackLogPath, JSON.stringify([], null, 2), 'utf8');
    }
    const openResult = await shell.openPath(feedbackLogPath);
    if (openResult) {
      return { success: false, error: openResult, path: feedbackLogPath };
    }
    return { success: true, path: feedbackLogPath };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('can-access-feedback-log', async (event, data = {}) => {
  const username = String(data?.username || '').trim();
  return { success: true, allowed: canAccessFeedbackLog(username) };
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

    const proc = spawnTracked(processorPath, args);
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

    const proc = spawnTracked(processorPath, args);
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
