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
const DEFAULT_AI_PROVIDER = 'localai';
const DEFAULT_CLAUDE_MODEL = 'claude-sonnet-4-20250514';
const DEFAULT_OPENAI_MODEL = 'gpt-4o-mini';
const DEFAULT_HARVEY_BASE_URL = 'https://api.harvey.ai';
const DEFAULT_OLLAMA_BASE_URL = 'http://127.0.0.1:11434';
const DEFAULT_LMSTUDIO_BASE_URL = 'http://127.0.0.1:1234';
const DEFAULT_LOCALAI_BASE_URL = 'http://127.0.0.1:11435';
const DEFAULT_OLLAMA_MODEL = 'llama3.1:8b';
const DEFAULT_LMSTUDIO_MODEL = 'local-model';
const DEFAULT_LOCALAI_MODEL = 'qwen2.5:1.5b';
const LOCAL_AI_INSTALL_PS_URL = 'https://ollama.com/install.ps1';
const LOCAL_AI_DOWNLOAD_PAGE_URL = 'https://ollama.com/download';
const LOCAL_AI_STATUS_EVENT = 'local-ai-progress';
const IMANAGE_DRIVE_ROOT_KEY = 'imanage_drive_root';
const LOCAL_AI_PROFILES = {
  'qwen2.5:1.5b': {
    id: 'qwen2.5:1.5b',
    label: 'Qwen 2.5 1.5B',
    size: '986MB',
    summary: 'Fastest free default. Best for routing, summaries, and lighter agent tasks.'
  },
  'qwen2.5:7b': {
    id: 'qwen2.5:7b',
    label: 'Qwen 2.5 7B',
    size: '4.7GB',
    summary: 'Better quality for legal drafting and longer reasoning.'
  },
  'gemma3:1b': {
    id: 'gemma3:1b',
    label: 'Gemma 3 1B',
    size: '815MB',
    summary: 'Smallest install. Good for quick local responses on lighter machines.'
  },
  'gemma3': {
    id: 'gemma3',
    label: 'Gemma 3 4B',
    size: '3.3GB',
    summary: 'Balanced local model for document tasks and chat.'
  }
};
const REQUIRE_FIREBASE_TELEMETRY = false;
const FIREBASE_TELEMETRY_PROBE_COLLECTION = 'telemetry_health';
const FIREBASE_STATUS_SUCCESS_CACHE_MS = 30000;
const FIREBASE_STATUS_FAILURE_CACHE_MS = 8000;
const TELEMETRY_INGEST_TIMEOUT_MS = 20000;
const TELEMETRY_INGEST_KEY = 'telemetry_ingest_url';
const TELEMETRY_INGEST_TOKEN_KEY = 'telemetry_ingest_token';
const AGENT_PROXY_URL_KEY = 'agent_proxy_url';
const AGENT_PROXY_TOKEN_KEY = 'agent_proxy_token';
const ACCESS_POLICY_URL_KEY = 'access_policy_url';
const ACCESS_POLICY_TOKEN_KEY = 'access_policy_token';
const ACCESS_POLICY_FAIL_CLOSED_KEY = 'access_policy_fail_closed';
const ACCESS_POLICY_TIMEOUT_MS = 15000;
const ACCESS_POLICY_CACHE_MS = 45000;
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
let managedLocalAiServerProcess = null;
let localAiBootstrapPromise = null;

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

function encodeSecretToSetting(plainValue) {
  const value = String(plainValue || '').trim();
  if (!value) return '';
  try {
    if (safeStorage.isEncryptionAvailable()) {
      const encrypted = safeStorage.encryptString(value);
      return 'enc:' + encrypted.toString('base64');
    }
  } catch (_) {}
  return 'plain:' + value;
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

function normalizeBooleanSetting(value, fallback = false) {
  if (typeof value === 'boolean') return value;
  const text = String(value ?? '').trim().toLowerCase();
  if (!text) return fallback;
  if (['1', 'true', 'yes', 'on', 'enabled', 'enable'].includes(text)) return true;
  if (['0', 'false', 'no', 'off', 'disabled', 'disable'].includes(text)) return false;
  return fallback;
}

function getAccessPolicyUrl() {
  const fromSetting = normalizeIngestUrl(getSettingValue(ACCESS_POLICY_URL_KEY, ''));
  if (fromSetting) return fromSetting;

  const fromEnv = normalizeIngestUrl(process.env.EMMANEIGH_ACCESS_POLICY_URL || '');
  if (fromEnv) return fromEnv;

  const bundled = loadBundledTelemetryConfig();
  if (!bundled) return '';
  return normalizeIngestUrl(
    bundled.accessPolicyUrl ||
    bundled.access_policy_url ||
    bundled.userAccessPolicyUrl ||
    bundled.user_access_policy_url ||
    bundled.killSwitchUrl ||
    bundled.kill_switch_url ||
    ''
  );
}

function getAccessPolicyToken() {
  const fromSetting = decodeSecretFromSetting(getSettingValue(ACCESS_POLICY_TOKEN_KEY, ''));
  if (fromSetting) return fromSetting;

  const fromEnv = String(process.env.EMMANEIGH_ACCESS_POLICY_TOKEN || '').trim();
  if (fromEnv) return fromEnv;

  const bundled = loadBundledTelemetryConfig();
  if (!bundled) return '';
  return String(
    bundled.accessPolicyToken ||
    bundled.access_policy_token ||
    bundled.userAccessPolicyToken ||
    bundled.user_access_policy_token ||
    bundled.killSwitchToken ||
    bundled.kill_switch_token ||
    ''
  ).trim();
}

function getAccessPolicyFailClosed() {
  const fromSettingRaw = getSettingValue(ACCESS_POLICY_FAIL_CLOSED_KEY, null);
  if (fromSettingRaw !== null && typeof fromSettingRaw !== 'undefined' && String(fromSettingRaw).trim() !== '') {
    return normalizeBooleanSetting(fromSettingRaw, false);
  }

  const envRaw = process.env.EMMANEIGH_ACCESS_POLICY_FAIL_CLOSED;
  if (typeof envRaw !== 'undefined' && String(envRaw).trim() !== '') {
    return normalizeBooleanSetting(envRaw, false);
  }

  const bundled = loadBundledTelemetryConfig();
  if (!bundled) return false;
  const bundledRaw = bundled.accessPolicyFailClosed ?? bundled.access_policy_fail_closed ?? null;
  return normalizeBooleanSetting(bundledRaw, false);
}

function normalizePolicyList(rawValue) {
  if (Array.isArray(rawValue)) {
    return rawValue
      .map((entry) => String(entry || '').trim().toLowerCase())
      .filter(Boolean);
  }
  if (typeof rawValue === 'string') {
    return rawValue
      .split(/[,\n;]+/)
      .map((entry) => String(entry || '').trim().toLowerCase())
      .filter(Boolean);
  }
  return [];
}

function getPolicyList(policy, keys) {
  for (const key of keys) {
    if (policy && Object.prototype.hasOwnProperty.call(policy, key)) {
      const list = normalizePolicyList(policy[key]);
      if (list.length > 0) return list;
    }
  }
  return [];
}

function matchesPolicyEntry(value, entry) {
  const normalizedValue = String(value || '').trim().toLowerCase();
  const normalizedEntry = String(entry || '').trim().toLowerCase();
  if (!normalizedValue || !normalizedEntry) return false;
  if (normalizedEntry === '*' || normalizedEntry === 'all') return true;
  if (normalizedValue === normalizedEntry) return true;

  if (normalizedValue.includes('@')) {
    if (normalizedEntry.startsWith('*@') && normalizedValue.endsWith(normalizedEntry.slice(1))) return true;
    if (normalizedEntry.startsWith('@') && normalizedValue.endsWith(normalizedEntry)) return true;
  }

  if (normalizedEntry.startsWith('*.') && normalizedValue.endsWith(normalizedEntry.slice(1))) return true;
  if (normalizedEntry.startsWith('.')) return normalizedValue.endsWith(normalizedEntry);
  return false;
}

function listHasMatch(list, value) {
  if (!Array.isArray(list) || list.length === 0) return false;
  return list.some((entry) => matchesPolicyEntry(value, entry));
}

function evaluateAccessPolicyPayload(payload, identity = {}) {
  const policyRoot = payload && typeof payload === 'object' ? payload : {};
  const policy = policyRoot.policy && typeof policyRoot.policy === 'object' ? policyRoot.policy : policyRoot;

  const email = normalizeEmail(identity.email || '');
  const username = normalizeUsername(identity.username || '');
  const userId = String(identity.userId || identity.user_id || '').trim().toLowerCase();
  const emailDomain = email.includes('@') ? email.split('@')[1] : '';
  const emailLocalPart = email.includes('@') ? email.split('@')[0] : '';
  const message =
    String(
      policy.message ||
      policy.reason ||
      policy.denied_message ||
      policyRoot.message ||
      policyRoot.reason ||
      ''
    ).trim() || 'Access to EmmaNeigh has been disabled by your administrator.';

  const directAllowRaw =
    policy.allowed ??
    policy.allow ??
    policy.is_allowed ??
    policy.enabled ??
    policy.active;
  const directBlockRaw =
    policy.blocked ??
    policy.block ??
    policy.deny ??
    policy.disabled ??
    policy.revoked ??
    policy.suspended;

  const hasDirectAllow = typeof directAllowRaw !== 'undefined' && directAllowRaw !== null;
  const hasDirectBlock = typeof directBlockRaw !== 'undefined' && directBlockRaw !== null;
  const directAllow = normalizeBooleanSetting(directAllowRaw, true);
  const directBlock = normalizeBooleanSetting(directBlockRaw, false);

  const statusText = String(policy.status || policyRoot.status || '').trim().toLowerCase();
  if (statusText === 'blocked' || statusText === 'disabled' || statusText === 'deny' || statusText === 'revoked') {
    return { allowed: false, message };
  }
  if (normalizeBooleanSetting(policy.block_all ?? policy.disable_all ?? policy.kill_switch ?? false, false)) {
    return { allowed: false, message };
  }
  if (hasDirectBlock && directBlock) {
    return { allowed: false, message };
  }
  if (hasDirectAllow && !directAllow) {
    return { allowed: false, message };
  }

  const blockedEmails = getPolicyList(policy, ['blocked_emails', 'blockedEmails', 'denylist', 'denied_emails', 'banned_emails']);
  const blockedDomains = getPolicyList(policy, ['blocked_domains', 'blockedDomains', 'denied_domains', 'banned_domains']);
  const blockedUsers = getPolicyList(policy, ['blocked_users', 'blockedUsers', 'blocked_usernames', 'blocked_user_ids', 'blockedUserIds']);

  const allowedEmails = getPolicyList(policy, ['allowed_emails', 'allowedEmails', 'allowlist', 'approved_emails']);
  const allowedDomains = getPolicyList(policy, ['allowed_domains', 'allowedDomains', 'approved_domains']);
  const allowedUsers = getPolicyList(policy, ['allowed_users', 'allowedUsers', 'allowed_usernames', 'allowed_user_ids', 'allowedUserIds']);

  if (email && listHasMatch(blockedEmails, email)) return { allowed: false, message };
  if (emailDomain && listHasMatch(blockedDomains, emailDomain)) return { allowed: false, message };
  if ((username && listHasMatch(blockedUsers, username)) || (userId && listHasMatch(blockedUsers, userId))) {
    return { allowed: false, message };
  }

  const hasAllowRules = allowedEmails.length > 0 || allowedDomains.length > 0 || allowedUsers.length > 0;
  if (hasAllowRules) {
    const emailAllowed = email && (listHasMatch(allowedEmails, email) || listHasMatch(allowedEmails, emailLocalPart));
    const domainAllowed = emailDomain && listHasMatch(allowedDomains, emailDomain);
    const userAllowed = (username && listHasMatch(allowedUsers, username)) || (userId && listHasMatch(allowedUsers, userId));
    if (!(emailAllowed || domainAllowed || userAllowed)) {
      return { allowed: false, message };
    }
  }

  return {
    allowed: true,
    message: String(
      policy.allowed_message ||
      policyRoot.allowed_message ||
      'Access granted.'
    ).trim()
  };
}

const accessPolicyCache = new Map();

function clearAccessPolicyCache() {
  accessPolicyCache.clear();
}

function buildAccessPolicyCacheKey(identity = {}, eventType = '') {
  return [
    normalizeEmail(identity.email || ''),
    normalizeUsername(identity.username || ''),
    String(identity.userId || '').trim().toLowerCase(),
    String(eventType || '').trim().toLowerCase()
  ].join('|');
}

function setAccessPolicyCache(key, value) {
  if (!key) return;
  accessPolicyCache.set(key, {
    checkedAt: Date.now(),
    value
  });
  if (accessPolicyCache.size > 500) {
    const firstKey = accessPolicyCache.keys().next().value;
    if (firstKey) accessPolicyCache.delete(firstKey);
  }
}

function getAccessPolicyCache(key, force = false) {
  if (!key || force) return null;
  const cached = accessPolicyCache.get(key);
  if (!cached) return null;
  if ((Date.now() - cached.checkedAt) > ACCESS_POLICY_CACHE_MS) {
    accessPolicyCache.delete(key);
    return null;
  }
  return cached.value;
}

async function evaluateAccessPolicy(identity = {}, options = {}) {
  const email = normalizeEmail(identity.email || '');
  const username = normalizeUsername(identity.username || '');
  const userId = String(identity.userId || identity.user_id || '').trim();
  const eventType = String(options.eventType || options.context || 'session').trim().toLowerCase() || 'session';
  const force = !!options.force;

  const policyUrl = getAccessPolicyUrl();
  if (!policyUrl) {
    return {
      success: true,
      configured: false,
      allowed: true,
      message: 'Access policy is not configured.'
    };
  }

  const cacheKey = buildAccessPolicyCacheKey({ email, username, userId }, eventType);
  const cached = getAccessPolicyCache(cacheKey, force);
  if (cached) return cached;

  const token = getAccessPolicyToken();
  const headers = { Accept: 'application/json' };
  if (token) {
    headers.Authorization = `Bearer ${token}`;
    headers['x-emmaneigh-access-key'] = token;
  }

  const policyPayload = {
    email: email || null,
    username: username || null,
    user_id: userId || null,
    machine_id: MACHINE_ID,
    app_version: APP_VERSION,
    event_type: eventType,
    timestamp: new Date().toISOString()
  };

  let getResponse;
  try {
    const requestUrl = new URL(policyUrl);
    if (email) requestUrl.searchParams.set('email', email);
    if (username) requestUrl.searchParams.set('username', username);
    if (userId) requestUrl.searchParams.set('user_id', userId);
    requestUrl.searchParams.set('event_type', eventType);
    requestUrl.searchParams.set('app_version', APP_VERSION);
    requestUrl.searchParams.set('machine_id', MACHINE_ID);

    getResponse = await requestHttps({
      baseUrl: requestUrl.toString(),
      path: '',
      method: 'GET',
      headers,
      timeoutMs: ACCESS_POLICY_TIMEOUT_MS
    });
  } catch (e) {
    getResponse = { statusCode: 0, networkError: e.message, rawBody: '', jsonBody: null };
  }

  let response = getResponse;
  if (response.statusCode === 405 || response.statusCode === 404 || response.statusCode === 400) {
    response = await requestHttps({
      baseUrl: policyUrl,
      path: '',
      method: 'POST',
      headers: {
        ...headers,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(policyPayload),
      timeoutMs: ACCESS_POLICY_TIMEOUT_MS
    });
  }

  if (!(response.statusCode >= 200 && response.statusCode < 300)) {
    const failClosed = getAccessPolicyFailClosed();
    const failureResult = {
      success: true,
      configured: true,
      allowed: !failClosed,
      message: failClosed
        ? `Access policy service is unavailable (${getErrorDetail(response)}). Access is blocked until policy is reachable.`
        : `Access policy service is unavailable (${getErrorDetail(response)}). Continuing with local access.`,
      source: 'policy_unreachable'
    };
    setAccessPolicyCache(cacheKey, failureResult);
    return failureResult;
  }

  const payload = response.jsonBody && typeof response.jsonBody === 'object'
    ? response.jsonBody
    : parseJsonSafe(response.rawBody || '') || {};
  const decision = evaluateAccessPolicyPayload(payload, { email, username, userId });
  const result = {
    success: true,
    configured: true,
    allowed: !!decision.allowed,
    message: decision.message || (decision.allowed ? 'Access granted.' : 'Access denied by policy.'),
    source: 'policy_remote'
  };
  setAccessPolicyCache(cacheKey, result);
  return result;
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
  if (value === 'local' || value === 'built-in' || value === 'builtin' || value === 'local ai' || value === 'local-ai' || value === 'local pack' || value === 'localpack') return 'localai';
  if (value === 'chatgpt') return 'openai';
  if (value === 'claude') return 'anthropic';
  if (value === 'localai' || value === 'ollama' || value === 'lmstudio' || value === 'anthropic' || value === 'openai' || value === 'harvey') return value;
  if (value === 'lm studio' || value === 'lm-studio') return 'lmstudio';
  return DEFAULT_AI_PROVIDER;
}

function isLocalProvider(provider) {
  const normalized = normalizeAIProvider(provider);
  return normalized === 'localai' || normalized === 'ollama' || normalized === 'lmstudio';
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
    case 'localai':
      return 'Local AI';
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

function getLocalAIProfiles() {
  return Object.values(LOCAL_AI_PROFILES);
}

function getLocalAIProfile(profileId) {
  const normalized = String(profileId || '').trim();
  if (normalized && LOCAL_AI_PROFILES[normalized]) return LOCAL_AI_PROFILES[normalized];
  return LOCAL_AI_PROFILES[DEFAULT_LOCALAI_MODEL];
}

function getLocalProviderBaseUrl(provider) {
  const normalized = normalizeAIProvider(provider);
  if (normalized === 'localai') return getLocalAIBaseUrl();
  if (normalized === 'lmstudio') return getLmStudioBaseUrl();
  return getOllamaBaseUrl();
}

function getLocalProviderModel(provider) {
  const normalized = normalizeAIProvider(provider);
  if (normalized === 'localai') return getLocalAIModel();
  if (normalized === 'lmstudio') return getLmStudioModel();
  return getOllamaModel();
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

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function normalizeProviderBaseUrl(rawValue, fallbackValue) {
  let value = String(rawValue || fallbackValue || '').trim();
  if (!value) value = String(fallbackValue || '').trim();
  if (!value) return '';
  if (!/^[a-z]+:\/\//i.test(value)) {
    value = `http://${value}`;
  }

  try {
    const url = new URL(value);
    url.hash = '';
    url.search = '';
    if (url.pathname === '/') {
      url.pathname = '';
    }
    return url.toString().replace(/\/$/, '');
  } catch (_) {
    return String(fallbackValue || value).trim().replace(/\/+$/, '');
  }
}

function getLocalAiRootDir() {
  return path.join(app.getPath('userData'), 'local-ai');
}

function getLocalAiModelsDir() {
  return path.join(getLocalAiRootDir(), 'models');
}

function ensureDirectoryExists(dirPath) {
  if (!dirPath) return;
  fs.mkdirSync(dirPath, { recursive: true });
}

function buildLoopbackBaseUrlVariants(rawValue, fallbackValue) {
  const primary = normalizeProviderBaseUrl(rawValue, fallbackValue);
  const variants = [];
  const pushVariant = (candidate) => {
    const normalized = normalizeProviderBaseUrl(candidate, primary);
    if (normalized && !variants.includes(normalized)) {
      variants.push(normalized);
    }
  };

  pushVariant(primary);

  try {
    const url = new URL(primary);
    const altHosts = [];
    if (url.hostname === '127.0.0.1') altHosts.push('localhost');
    if (url.hostname === 'localhost') altHosts.push('127.0.0.1');
    if (url.hostname === '0.0.0.0') altHosts.push('127.0.0.1', 'localhost');
    for (const host of altHosts) {
      const alt = new URL(primary);
      alt.hostname = host;
      pushVariant(alt.toString());
    }
  } catch (_) {}

  return variants;
}

function isLocalProviderConnectionError(detail) {
  const text = String(detail || '').trim().toLowerCase();
  if (!text) return false;
  return (
    text.includes('econnrefused') ||
    text.includes('connection refused') ||
    text.includes('actively refused') ||
    text.includes('enotfound') ||
    text.includes('ehostunreach') ||
    text.includes('network error')
  );
}

function findExecutableOnPath(commandName) {
  try {
    const lookupCommand = process.platform === 'win32' ? 'where' : 'which';
    const result = spawnSync(lookupCommand, [commandName], {
      encoding: 'utf8',
      windowsHide: true
    });
    if (result.status !== 0) return '';
    return String(result.stdout || '')
      .split(/\r?\n/)
      .map(line => line.trim())
      .find(Boolean) || '';
  } catch (_) {
    return '';
  }
}

function parseProviderHostPort(baseUrl, fallbackPort) {
  try {
    const parsed = new URL(baseUrl);
    return `${parsed.hostname}:${parsed.port || fallbackPort}`;
  } catch (_) {
    return `127.0.0.1:${fallbackPort}`;
  }
}

function getLocalAiOllamaEnv() {
  const baseUrl = getLocalAIBaseUrl();
  ensureDirectoryExists(getLocalAiRootDir());
  ensureDirectoryExists(getLocalAiModelsDir());
  return {
    ...process.env,
    OLLAMA_HOST: parseProviderHostPort(baseUrl, 11435),
    OLLAMA_MODELS: getLocalAiModelsDir()
  };
}

function emitLocalAiProgress(message, extra = {}) {
  if (!mainWindow || !mainWindow.webContents) return;
  mainWindow.webContents.send(LOCAL_AI_STATUS_EVENT, {
    message: String(message || '').trim(),
    is_error: !!extra.is_error,
    step: extra.step || null,
    timestamp: new Date().toISOString()
  });
}

function findOllamaBinary() {
  const cliPath = findExecutableOnPath('ollama');
  if (cliPath) return cliPath;

  const candidates = process.platform === 'win32'
    ? [
        path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Ollama', 'ollama.exe'),
        path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Ollama', 'Ollama.exe'),
        path.join(process.env.ProgramFiles || '', 'Ollama', 'ollama.exe'),
        path.join(process.env.ProgramFiles || '', 'Ollama', 'Ollama.exe'),
        path.join(process.env['ProgramFiles(x86)'] || '', 'Ollama', 'ollama.exe'),
        path.join(process.env['ProgramFiles(x86)'] || '', 'Ollama', 'Ollama.exe')
      ]
    : [
        '/usr/local/bin/ollama',
        '/opt/homebrew/bin/ollama',
        '/Applications/Ollama.app/Contents/Resources/ollama',
        '/Applications/Ollama.app/Contents/MacOS/Ollama'
      ];

  return candidates.find(candidate => candidate && fs.existsSync(candidate)) || '';
}

function spawnLocalAiServer(commandPath) {
  if (!commandPath) return null;
  try {
    const proc = rawSpawn(commandPath, ['serve'], {
      env: getLocalAiOllamaEnv(),
      windowsHide: true,
      detached: false,
      stdio: 'ignore'
    });
    proc.once('exit', () => {
      if (managedLocalAiServerProcess === proc) {
        managedLocalAiServerProcess = null;
      }
    });
    proc.once('error', () => {
      if (managedLocalAiServerProcess === proc) {
        managedLocalAiServerProcess = null;
      }
    });
    managedLocalAiServerProcess = proc;
    return proc;
  } catch (_) {
    return null;
  }
}

async function startManagedLocalAiServer() {
  const binaryPath = findOllamaBinary();
  if (!binaryPath) {
    return {
      success: false,
      error: process.platform === 'win32'
        ? 'Ollama is not installed yet. Click "Install Local AI" in Settings to install the free local runtime.'
        : `Ollama is not installed yet. Install it from ${LOCAL_AI_DOWNLOAD_PAGE_URL}, then retry.`
    };
  }

  const baseUrl = getLocalAIBaseUrl();
  const alreadyRunning = await probeLocalProviderBaseUrl({
    provider: 'localai',
    baseUrl,
    apiKey: ''
  });
  if (alreadyRunning.reachable) {
    return {
      success: true,
      baseUrl,
      binaryPath,
      note: 'Local AI runtime already running.'
    };
  }

  emitLocalAiProgress('Starting the local AI runtime...', { step: 'start_runtime' });
  const proc = spawnLocalAiServer(binaryPath);
  if (!proc) {
    return { success: false, error: 'EmmaNeigh could not start the local AI runtime.' };
  }

  for (let attempt = 0; attempt < 20; attempt += 1) {
    await sleep(1000);
    const probe = await probeLocalProviderBaseUrl({
      provider: 'localai',
      baseUrl,
      apiKey: ''
    });
    if (probe.reachable) {
      return {
        success: true,
        baseUrl,
        binaryPath,
        note: 'Local AI runtime started.'
      };
    }
  }

  return {
    success: false,
    error: `EmmaNeigh started the local runtime, but it did not respond at ${baseUrl}.`
  };
}

async function installOllamaForLocalAi() {
  if (findOllamaBinary()) {
    return { success: true, installed: true, message: 'Ollama is already installed.' };
  }

  if (process.platform !== 'win32') {
    return {
      success: false,
      installed: false,
      error: `Automatic install is currently supported on Windows only. Install Ollama from ${LOCAL_AI_DOWNLOAD_PAGE_URL}, then retry.`
    };
  }

  emitLocalAiProgress('Installing the free local AI runtime...', { step: 'install_runtime' });
  return await new Promise((resolve) => {
    const installer = rawSpawn('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      `irm ${LOCAL_AI_INSTALL_PS_URL} | iex`
    ], {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe']
    });

    let stderr = '';
    installer.stdout.on('data', (chunk) => {
      const text = String(chunk || '').trim();
      if (text) emitLocalAiProgress(text, { step: 'install_runtime' });
    });
    installer.stderr.on('data', (chunk) => {
      stderr += String(chunk || '');
    });
    installer.on('close', () => {
      const binaryPath = findOllamaBinary();
      if (binaryPath) {
        resolve({ success: true, installed: true, message: 'Ollama installed successfully.', binaryPath });
        return;
      }
      resolve({
        success: false,
        installed: false,
        error: stderr.trim() || 'Ollama installation did not complete successfully.'
      });
    });
    installer.on('error', (err) => {
      resolve({ success: false, installed: false, error: err.message });
    });
  });
}

async function ensureLocalAiModelInstalled(modelName) {
  const model = String(modelName || getLocalAIModel()).trim() || DEFAULT_LOCALAI_MODEL;
  const baseUrl = getLocalAIBaseUrl();
  const currentModels = await fetchOpenAICompatibleModels({ baseUrl, apiKey: '' });
  if (currentModels.includes(model)) {
    return { success: true, model, alreadyInstalled: true };
  }

  const binaryPath = findOllamaBinary();
  if (!binaryPath) {
    return { success: false, error: 'Local AI runtime is not installed yet.' };
  }

  emitLocalAiProgress(`Downloading the local model (${model}). This can take a few minutes...`, { step: 'pull_model' });
  return await new Promise((resolve) => {
    const proc = rawSpawn(binaryPath, ['pull', model], {
      env: getLocalAiOllamaEnv(),
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe']
    });
    let stderr = '';
    proc.stdout.on('data', (chunk) => {
      const text = String(chunk || '').trim();
      if (text) emitLocalAiProgress(text, { step: 'pull_model' });
    });
    proc.stderr.on('data', (chunk) => {
      const text = String(chunk || '').trim();
      if (text) {
        stderr += `${text}\n`;
        emitLocalAiProgress(text, { step: 'pull_model' });
      }
    });
    proc.on('close', async (code) => {
      if (code !== 0) {
        resolve({ success: false, error: stderr.trim() || `Model download failed with exit code ${code}.` });
        return;
      }
      const refreshedModels = await fetchOpenAICompatibleModels({ baseUrl, apiKey: '' });
      if (!refreshedModels.includes(model)) {
        resolve({ success: false, error: `Model download finished, but ${model} is still not available.` });
        return;
      }
      resolve({ success: true, model, models: refreshedModels });
    });
    proc.on('error', (err) => resolve({ success: false, error: err.message }));
  });
}

async function bootstrapManagedLocalAi(profileId) {
  if (localAiBootstrapPromise) return localAiBootstrapPromise;

  localAiBootstrapPromise = (async () => {
    const profile = getLocalAIProfile(profileId || getLocalAIModel());
    emitLocalAiProgress(`Preparing Local AI using ${profile.label}...`, { step: 'prepare' });

    const installResult = await installOllamaForLocalAi();
    if (!installResult.success) {
      emitLocalAiProgress(installResult.error, { step: 'install_runtime', is_error: true });
      return installResult;
    }

    const startResult = await startManagedLocalAiServer();
    if (!startResult.success) {
      emitLocalAiProgress(startResult.error, { step: 'start_runtime', is_error: true });
      return startResult;
    }

    const modelResult = await ensureLocalAiModelInstalled(profile.id);
    if (!modelResult.success) {
      emitLocalAiProgress(modelResult.error, { step: 'pull_model', is_error: true });
      return modelResult;
    }

    emitLocalAiProgress(`Local AI is ready with ${profile.label}.`, { step: 'complete' });
    return {
      success: true,
      provider: 'localai',
      model: profile.id,
      baseUrl: getLocalAIBaseUrl(),
      profile
    };
  })();

  try {
    return await localAiBootstrapPromise;
  } finally {
    localAiBootstrapPromise = null;
  }
}

function launchDetachedProcess(command, args = []) {
  try {
    const child = rawSpawn(command, args, {
      detached: true,
      stdio: 'ignore',
      windowsHide: true
    });
    child.unref();
    return true;
  } catch (_) {
    return false;
  }
}

function getOllamaLaunchCommand() {
  const cliPath = findExecutableOnPath('ollama');
  if (cliPath) {
    return {
      installed: true,
      command: cliPath,
      args: ['serve'],
      display: `${cliPath} serve`
    };
  }

  if (process.platform === 'darwin') {
    const appCandidates = [
      '/Applications/Ollama.app/Contents/MacOS/Ollama',
      path.join(os.homedir(), 'Applications', 'Ollama.app', 'Contents', 'MacOS', 'Ollama')
    ];
    const appBinary = appCandidates.find(candidate => fs.existsSync(candidate));
    if (appBinary) {
      return {
        installed: true,
        command: appBinary,
        args: [],
        display: appBinary
      };
    }
    const openPath = findExecutableOnPath('open');
    if (openPath) {
      return {
        installed: fs.existsSync('/Applications/Ollama.app') || fs.existsSync(path.join(os.homedir(), 'Applications', 'Ollama.app')),
        command: openPath,
        args: ['-a', 'Ollama'],
        display: 'open -a Ollama'
      };
    }
  }

  if (process.platform === 'win32') {
    const appCandidates = [
      path.join(process.env.LOCALAPPDATA || '', 'Programs', 'Ollama', 'Ollama.exe'),
      path.join(process.env.ProgramFiles || '', 'Ollama', 'Ollama.exe'),
      path.join(process.env['ProgramFiles(x86)'] || '', 'Ollama', 'Ollama.exe')
    ].filter(Boolean);
    const appBinary = appCandidates.find(candidate => fs.existsSync(candidate));
    if (appBinary) {
      return {
        installed: true,
        command: appBinary,
        args: [],
        display: appBinary
      };
    }
  }

  return { installed: false, command: '', args: [], display: '' };
}

function getLmStudioLaunchCommand() {
  if (process.platform === 'darwin') {
    const appCandidates = [
      '/Applications/LM Studio.app/Contents/MacOS/LM Studio',
      path.join(os.homedir(), 'Applications', 'LM Studio.app', 'Contents', 'MacOS', 'LM Studio')
    ];
    const appBinary = appCandidates.find(candidate => fs.existsSync(candidate));
    if (appBinary) {
      return {
        installed: true,
        command: appBinary,
        args: [],
        display: appBinary
      };
    }
  }

  if (process.platform === 'win32') {
    const appCandidates = [
      path.join(process.env.LOCALAPPDATA || '', 'Programs', 'LM Studio', 'LM Studio.exe'),
      path.join(process.env.ProgramFiles || '', 'LM Studio', 'LM Studio.exe'),
      path.join(process.env['ProgramFiles(x86)'] || '', 'LM Studio', 'LM Studio.exe')
    ].filter(Boolean);
    const appBinary = appCandidates.find(candidate => fs.existsSync(candidate));
    if (appBinary) {
      return {
        installed: true,
        command: appBinary,
        args: [],
        display: appBinary
      };
    }
  }

  return { installed: false, command: '', args: [], display: '' };
}

function tryLaunchLocalProvider(provider) {
  const normalized = normalizeAIProvider(provider);
  if (normalized === 'localai') {
    const binaryPath = findOllamaBinary();
    if (!binaryPath) {
      return {
        attempted: false,
        started: false,
        installed: false,
        display: ''
      };
    }
    const proc = spawnLocalAiServer(binaryPath);
    return {
      attempted: true,
      started: !!proc,
      installed: true,
      display: `${binaryPath} serve`
    };
  }
  const launchCommand = normalized === 'ollama'
    ? getOllamaLaunchCommand()
    : getLmStudioLaunchCommand();

  if (!launchCommand.command) {
    return {
      attempted: false,
      started: false,
      installed: !!launchCommand.installed,
      display: ''
    };
  }

  const started = launchDetachedProcess(launchCommand.command, launchCommand.args || []);
  return {
    attempted: true,
    started,
    installed: true,
    display: launchCommand.display
  };
}

async function probeLocalProviderBaseUrl({ provider, baseUrl, apiKey }) {
  const headers = {};
  if (apiKey) headers.Authorization = `Bearer ${apiKey}`;

  const response = await requestHttps({
    baseUrl,
    path: '/v1/models',
    method: 'GET',
    headers,
    timeoutMs: 15000
  });

  if (response.statusCode === 200) {
    const models = Array.isArray(response.jsonBody?.data)
      ? response.jsonBody.data.map(item => String(item?.id || '').trim()).filter(Boolean)
      : [];
    return {
      ok: models.length > 0,
      reachable: true,
      baseUrl,
      models,
      endpoint: '/v1/models',
      response,
      note: models.length > 0 ? '' : `${getAIProviderDisplayName(provider)} is reachable but no models are loaded.`
    };
  }

  if (provider === 'ollama' || provider === 'localai') {
    const tagsResponse = await requestHttps({
      baseUrl,
      path: '/api/tags',
      method: 'GET',
      headers: {},
      timeoutMs: 15000
    });
    if (tagsResponse.statusCode === 200) {
      const models = Array.isArray(tagsResponse.jsonBody?.models)
        ? tagsResponse.jsonBody.models.map(item => String(item?.name || '').trim()).filter(Boolean)
        : [];
      return {
        ok: models.length > 0,
        reachable: true,
        baseUrl,
        models,
        endpoint: '/api/tags',
        response: tagsResponse,
        note: models.length > 0
          ? `${getAIProviderDisplayName(provider)} is connected via the native local endpoint.`
          : `${getAIProviderDisplayName(provider)} is running but no models are installed yet.`
      };
    }
  }

  return {
    ok: false,
    reachable: response.statusCode > 0,
    baseUrl,
    models: [],
    endpoint: '/v1/models',
    response
  };
}

function buildLocalProviderFailureMessage({ provider, baseUrl, configuredModel, detail, launchInfo, models }) {
  const label = getAIProviderDisplayName(provider);
  const configured = configuredModel || (
    provider === 'localai' ? DEFAULT_LOCALAI_MODEL
    : provider === 'ollama' ? DEFAULT_OLLAMA_MODEL
    : DEFAULT_LMSTUDIO_MODEL
  );
  const steps = [];

  if (provider === 'localai') {
    if (!launchInfo || launchInfo.installed === false) {
      if (process.platform === 'win32') {
        steps.push('Open Settings and click "Install Local AI" to install the free runtime and model.');
      } else {
        steps.push(`Install Ollama from ${LOCAL_AI_DOWNLOAD_PAGE_URL}, then return to EmmaNeigh and click "Install Local AI".`);
      }
    } else if (launchInfo.attempted && !launchInfo.started) {
      steps.push('EmmaNeigh found the local runtime but could not start it automatically.');
    }
    if (Array.isArray(models) && models.length > 0) {
      steps.push(`Available local models: ${models.slice(0, 5).join(', ')}.`);
    } else {
      const profile = getLocalAIProfile(configured);
      steps.push(`Install a free local model such as ${profile.label} from Settings.`);
    }
    steps.push(`EmmaNeigh uses ${baseUrl} for the built-in local model.`);
  } else if (provider === 'ollama') {
    if (!launchInfo || launchInfo.installed === false) {
      steps.push('Install Ollama from https://ollama.com/download.');
    } else if (launchInfo.attempted && !launchInfo.started) {
      steps.push('EmmaNeigh found Ollama but could not start it automatically.');
    }
    steps.push('Open the Ollama app or run `ollama serve`.');
    if (Array.isArray(models) && models.length > 0) {
      steps.push(`Available local models: ${models.slice(0, 5).join(', ')}.`);
    } else {
      steps.push(`Pull a local model with \`ollama pull ${configured}\`.`);
    }
    steps.push(`Keep the base URL set to ${baseUrl}.`);
  } else {
    if (!launchInfo || launchInfo.installed === false) {
      steps.push('Install LM Studio from https://lmstudio.ai/.');
    } else if (launchInfo.attempted && !launchInfo.started) {
      steps.push('EmmaNeigh found LM Studio but could not launch it automatically.');
    }
    steps.push('Open LM Studio.');
    steps.push('Load a model into the chat/runtime.');
    steps.push('Turn on the OpenAI-compatible local server in LM Studio.');
    steps.push(`Keep the base URL set to ${baseUrl}.`);
  }

  const detailText = detail ? ` Last error: ${detail}.` : '';
  return `${label} is not ready at ${baseUrl}.${detailText} ${steps.join(' ')}`.trim();
}

async function ensureLocalProviderAvailable(provider, apiKey, options = {}) {
  const normalized = normalizeAIProvider(provider);
  if (normalized === 'localai') {
    const configuredBaseUrl = getLocalAIBaseUrl();
    const configuredModel = getLocalAIModel();
    const variants = buildLoopbackBaseUrlVariants(configuredBaseUrl, DEFAULT_LOCALAI_BASE_URL);
    const allowAutoLaunch = options.allowAutoLaunch !== false;
    const allowBootstrap = !!options.allowBootstrap;

    let lastProbe = null;
    for (const candidateBaseUrl of variants) {
      const probe = await probeLocalProviderBaseUrl({
        provider: normalized,
        baseUrl: candidateBaseUrl,
        apiKey: ''
      });
      lastProbe = probe;
      if (probe.ok && probe.models.includes(configuredModel)) {
        return {
          success: true,
          provider: normalized,
          baseUrl: candidateBaseUrl,
          configuredBaseUrl,
          configuredModel,
          models: probe.models,
          endpoint: probe.endpoint,
          message: `${getAIProviderDisplayName(normalized)} is ready at ${candidateBaseUrl}.`
        };
      }
    }

    let launchInfo = null;
    if (allowBootstrap) {
      const bootstrapResult = await bootstrapManagedLocalAi(configuredModel);
      if (bootstrapResult.success) {
        const finalProbe = await probeLocalProviderBaseUrl({
          provider: normalized,
          baseUrl: configuredBaseUrl,
          apiKey: ''
        });
        return {
          success: true,
          provider: normalized,
          baseUrl: configuredBaseUrl,
          configuredBaseUrl,
          configuredModel,
          models: finalProbe.models || [configuredModel],
          endpoint: finalProbe.endpoint || '/api/tags',
          autoStarted: true,
          launchCommand: findOllamaBinary(),
          message: `${getAIProviderDisplayName(normalized)} is ready at ${configuredBaseUrl}.`
        };
      }
      return {
        success: false,
        provider: normalized,
        baseUrl: configuredBaseUrl,
        configuredBaseUrl,
        configuredModel,
        models: [],
        endpoint: '/api/tags',
        autoStarted: false,
        launchCommand: '',
        error: bootstrapResult.error || 'Failed to bootstrap Local AI.'
      };
    }

    const lastDetail = lastProbe ? getErrorDetail(lastProbe.response) : 'Unknown error';
    if (allowAutoLaunch && isLocalProviderConnectionError(lastDetail)) {
      launchInfo = tryLaunchLocalProvider(normalized);
      if (launchInfo.started) {
        for (let attempt = 0; attempt < 12; attempt += 1) {
          await sleep(1000);
          const probe = await probeLocalProviderBaseUrl({
            provider: normalized,
            baseUrl: configuredBaseUrl,
            apiKey: ''
          });
          lastProbe = probe;
          if (probe.ok && probe.models.includes(configuredModel)) {
            return {
              success: true,
              provider: normalized,
              baseUrl: configuredBaseUrl,
              configuredBaseUrl,
              configuredModel,
              models: probe.models,
              endpoint: probe.endpoint,
              autoStarted: true,
              launchCommand: launchInfo.display,
              message: `${getAIProviderDisplayName(normalized)} started and is ready at ${configuredBaseUrl}.`
            };
          }
        }
      }
    }

    const detail = lastProbe ? getErrorDetail(lastProbe.response) : 'Unknown error';
    return {
      success: false,
      provider: normalized,
      baseUrl: configuredBaseUrl,
      configuredBaseUrl,
      configuredModel,
      models: lastProbe?.models || [],
      endpoint: lastProbe?.endpoint || '/api/tags',
      autoStarted: !!launchInfo?.started,
      launchCommand: launchInfo?.display || '',
      error: buildLocalProviderFailureMessage({
        provider: normalized,
        baseUrl: configuredBaseUrl,
        configuredModel,
        detail,
        launchInfo,
        models: lastProbe?.models || []
      })
    };
  }

  const configuredBaseUrl = normalized === 'lmstudio'
    ? getLmStudioBaseUrl()
    : getOllamaBaseUrl();
  const configuredModel = normalized === 'lmstudio'
    ? getLmStudioModel()
    : getOllamaModel();
  const fallbackBaseUrl = normalized === 'lmstudio'
    ? DEFAULT_LMSTUDIO_BASE_URL
    : DEFAULT_OLLAMA_BASE_URL;
  const variants = buildLoopbackBaseUrlVariants(configuredBaseUrl, fallbackBaseUrl);
  const allowAutoLaunch = options.allowAutoLaunch !== false;

  let lastProbe = null;
  for (const candidateBaseUrl of variants) {
    const probe = await probeLocalProviderBaseUrl({
      provider: normalized,
      baseUrl: candidateBaseUrl,
      apiKey
    });
    lastProbe = probe;
    if (probe.ok) {
      return {
        success: true,
        provider: normalized,
        baseUrl: candidateBaseUrl,
        configuredBaseUrl,
        configuredModel,
        models: probe.models,
        endpoint: probe.endpoint,
        message: `${getAIProviderDisplayName(normalized)} is ready at ${candidateBaseUrl}.`
      };
    }
  }

  let launchInfo = null;
  const lastDetail = lastProbe ? getErrorDetail(lastProbe.response) : 'Unknown error';
  if (allowAutoLaunch && isLocalProviderConnectionError(lastDetail)) {
    launchInfo = tryLaunchLocalProvider(normalized);
    if (launchInfo.started) {
      for (let attempt = 0; attempt < 12; attempt += 1) {
        await sleep(1000);
        for (const candidateBaseUrl of variants) {
          const probe = await probeLocalProviderBaseUrl({
            provider: normalized,
            baseUrl: candidateBaseUrl,
            apiKey
          });
          lastProbe = probe;
          if (probe.ok) {
            return {
              success: true,
              provider: normalized,
              baseUrl: candidateBaseUrl,
              configuredBaseUrl,
              configuredModel,
              models: probe.models,
              endpoint: probe.endpoint,
              autoStarted: true,
              launchCommand: launchInfo.display,
              message: `${getAIProviderDisplayName(normalized)} started and is ready at ${candidateBaseUrl}.`
            };
          }
        }
      }
    }
  }

  const detail = lastProbe ? getErrorDetail(lastProbe.response) : 'Unknown error';
  return {
    success: false,
    provider: normalized,
    baseUrl: configuredBaseUrl,
    configuredBaseUrl,
    configuredModel,
    models: lastProbe?.models || [],
    endpoint: lastProbe?.endpoint || '/v1/models',
    autoStarted: !!launchInfo?.started,
    launchCommand: launchInfo?.display || '',
    error: buildLocalProviderFailureMessage({
      provider: normalized,
      baseUrl: configuredBaseUrl,
      configuredModel,
      detail,
      launchInfo,
      models: lastProbe?.models || []
    })
  };
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
  // Try OpenAI-compatible /v1/models first
  const response = await requestHttps({
    baseUrl,
    path: '/v1/models',
    method: 'GET',
    headers,
    timeoutMs: 15000
  });
  if (response.statusCode === 200) {
    const entries = Array.isArray(response.jsonBody?.data) ? response.jsonBody.data : [];
    const models = entries.map(item => String(item?.id || '').trim()).filter(Boolean);
    if (models.length > 0) return models;
  }
  // Fallback: Ollama native /api/tags endpoint (for older Ollama versions)
  try {
    const tagsResponse = await requestHttps({
      baseUrl,
      path: '/api/tags',
      method: 'GET',
      headers: {},
      timeoutMs: 15000
    });
    if (tagsResponse.statusCode === 200 && Array.isArray(tagsResponse.jsonBody?.models)) {
      return tagsResponse.jsonBody.models.map(m => String(m?.name || '').trim()).filter(Boolean);
    }
  } catch (_) {}
  return [];
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
      stream: false,
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

async function callLocalAIPrompt({ prompt, maxTokens }) {
  const localProvider = await ensureLocalProviderAvailable('localai', '', {
    allowAutoLaunch: true,
    allowBootstrap: true
  });
  if (!localProvider.success) {
    return { success: false, error: localProvider.error };
  }

  const configuredModel = getLocalAIModel();
  const modelCandidates = Array.from(new Set([
    configuredModel,
    DEFAULT_LOCALAI_MODEL,
    'qwen2.5:1.5b',
    'qwen2.5:7b',
    'gemma3:1b',
    'gemma3'
  ]));

  const openaiResult = await callOpenAICompatiblePrompt({
    baseUrl: localProvider.baseUrl,
    apiKey: '',
    prompt,
    maxTokens,
    modelCandidates,
    providerLabel: 'Local AI'
  });
  if (openaiResult.success) return openaiResult;

  const discovered = await fetchOpenAICompatibleModels({ baseUrl: localProvider.baseUrl, apiKey: '' });
  const allCandidates = Array.from(new Set([...modelCandidates, ...discovered]));
  let lastError = openaiResult.error || 'Unknown error';
  for (const modelName of allCandidates) {
    const response = await requestHttps({
      baseUrl: localProvider.baseUrl,
      path: '/api/chat',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: modelName,
        stream: false,
        messages: [{ role: 'user', content: prompt }],
        options: maxTokens ? { num_predict: maxTokens } : {}
      }),
      timeoutMs: 90000
    });
    if (response.statusCode === 200 && response.jsonBody?.message?.content) {
      return { success: true, text: String(response.jsonBody.message.content).trim(), modelUsed: modelName };
    }
    if (response.statusCode === 404) continue;
    lastError = getErrorDetail(response);
  }

  return {
    success: false,
    error: `Local AI could not respond. ${lastError}`
  };
}

async function callOllamaPrompt({ apiKey, prompt, maxTokens }) {
  const localProvider = await ensureLocalProviderAvailable('ollama', apiKey, { allowAutoLaunch: true });
  if (!localProvider.success) {
    return { success: false, error: localProvider.error };
  }
  const ollamaBase = localProvider.baseUrl;
  const modelCandidates = [
    getOllamaModel(),
    DEFAULT_OLLAMA_MODEL,
    'llama3.1:8b',
    'qwen2.5:7b',
    'gemma3:4b'
  ];

  // Try OpenAI-compatible endpoint first (preferred, works with Ollama 0.1.24+)
  const openaiResult = await callOpenAICompatiblePrompt({
    baseUrl: ollamaBase,
    apiKey,
    prompt,
    maxTokens,
    modelCandidates,
    providerLabel: 'Ollama'
  });
  if (openaiResult.success) return openaiResult;

  // Fallback: Ollama native /api/chat endpoint (for older versions or when /v1 fails)
  const discovered = await fetchOpenAICompatibleModels({ baseUrl: ollamaBase, apiKey });
  const allCandidates = Array.from(new Set([...modelCandidates, ...discovered]));
  let lastError = openaiResult.error || 'Unknown error';

  for (const modelName of allCandidates) {
    try {
      const response = await requestHttps({
        baseUrl: ollamaBase,
        path: '/api/chat',
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: modelName,
          stream: false,
          messages: [{ role: 'user', content: prompt }],
          options: maxTokens ? { num_predict: maxTokens } : {}
        }),
        timeoutMs: 90000
      });
      if (response.statusCode === 200 && response.jsonBody?.message?.content) {
        return { success: true, text: String(response.jsonBody.message.content).trim(), modelUsed: modelName };
      }
      if (response.statusCode === 404) continue; // Model not found, try next
      lastError = getErrorDetail(response);
    } catch (e) {
      lastError = e.message;
    }
  }

  return { success: false, error: `Ollama is not responding. Ensure Ollama is running (ollama serve) and a model is pulled. Last error: ${lastError}` };
}

async function callLmStudioPrompt({ apiKey, prompt, maxTokens }) {
  const localProvider = await ensureLocalProviderAvailable('lmstudio', apiKey, { allowAutoLaunch: true });
  if (!localProvider.success) {
    return { success: false, error: localProvider.error };
  }
  const modelCandidates = [
    getLmStudioModel(),
    DEFAULT_LMSTUDIO_MODEL
  ];
  return callOpenAICompatiblePrompt({
    baseUrl: localProvider.baseUrl,
    apiKey,
    prompt,
    maxTokens,
    modelCandidates,
    providerLabel: 'LM Studio'
  });
}

async function callProviderPrompt({ provider, apiKey, prompt, maxTokens }) {
  const normalized = resolveProviderForApiKey(provider, apiKey);
  if (normalized === 'localai') {
    return callLocalAIPrompt({ prompt, maxTokens });
  }
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

  if (provider === 'localai') {
    baseUrl = getLocalAIBaseUrl();
    model = getLocalAIModel();
  } else if (provider === 'ollama') {
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

  if (isLocalProvider(connectionConfig.provider)) {
    const localProvider = await ensureLocalProviderAvailable(connectionConfig.provider, apiKey, { allowAutoLaunch: true });
    if (localProvider.success) {
      return {
        ok: true,
        status: 'reachable',
        statusCode: 200,
        endpoint: buildProviderHealthEndpoint(localProvider.baseUrl, localProvider.endpoint || '/v1/models'),
        detail: localProvider.message
      };
    }
    return {
      ok: false,
      status: 'unreachable',
      statusCode: 0,
      endpoint: connectionConfig.healthEndpoint,
      detail: localProvider.error
    };
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
    if (docxFiles.length > 0) {
      return {
        action: 'run_update_checklist',
        target_tab: 'updatechecklist',
        run_now: true,
        required_extensions: ['docx'],
        missing_requirements: [],
        user_message: 'Running Update Checklist using the attached checklist and Outlook email folders.'
      };
    }
    return {
      action: 'open_tab',
      target_tab: 'updatechecklist',
      run_now: false,
      required_extensions: ['docx'],
      missing_requirements: ['Attach one checklist (.docx) to run Update Checklist. EmmaNeigh will scan Outlook automatically.'],
      user_message: 'Opened Update Checklist. Attach a checklist document to run.'
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
      required_extensions: [],
      missing_requirements: [],
      user_message: 'Opened Email Search. EmmaNeigh can scan Outlook email folders directly.'
    };
  }

  if (/checklist|update checklist/.test(prompt)) {
    return {
      action: 'open_tab',
      target_tab: 'updatechecklist',
      run_now: false,
      required_extensions: ['docx'],
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
    if (docxFiles.length < 1) {
      runNow = false;
      missingRequirements = ['Attach one checklist (.docx) to run Update Checklist. EmmaNeigh will scan Outlook automatically.'];
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

function ensureSignaturePacketProcessor(processorPath) {
  if (!processorPath) {
    throw new Error('Development mode - please build the app first');
  }

  if (!fs.existsSync(processorPath)) {
    throw new Error('Processor not found: ' + processorPath);
  }

  if (process.platform === 'darwin') {
    try { fs.chmodSync(processorPath, '755'); } catch (e) {}
  }
}

function buildSignaturePacketArgs(input, extraConfig = {}) {
  const config = { ...extraConfig };

  if (typeof input === 'string') {
    config.folder = input;
  } else if (input && input.folder) {
    config.folder = input.folder;
    if (input.output_format) {
      config.output_format = input.output_format;
    }
  } else if (input && input.files) {
    config.files = input.files;
    config.output_format = input.output_format || 'preserve';
  } else {
    throw new Error('Invalid input: must be folder path or { folder } or { files }');
  }

  const configPath = path.join(
    app.getPath('temp'),
    `packets-config-${Date.now()}-${Math.floor(Math.random() * 100000)}.json`
  );
  fs.writeFileSync(configPath, JSON.stringify(config));
  return { args: ['--config', configPath], configPath };
}

function runSignaturePacketProcessor(input, extraConfig = {}, onMessage = null) {
  return new Promise((resolve, reject) => {
    const moduleName = 'signature_packets';
    const processorPath = getProcessorPath(moduleName);
    let configRef = null;
    let proc = null;
    let result = null;
    let settled = false;

    const cleanup = () => {
      if (!configRef || !configRef.configPath) return;
      try { fs.unlinkSync(configRef.configPath); } catch (e) {}
    };

    const fail = (err) => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(err instanceof Error ? err : new Error(String(err || 'Signature packet processing failed.')));
    };

    const succeed = (payload) => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(payload);
    };

    try {
      ensureSignaturePacketProcessor(processorPath);
      configRef = buildSignaturePacketArgs(input, extraConfig);
      proc = spawnTracked(processorPath, [moduleName, ...configRef.args]);
    } catch (err) {
      fail(err);
      return;
    }

    proc.stdout.on('data', (data) => {
      const lines = data.toString().split('\n').filter(l => l.trim());
      for (const line of lines) {
        try {
          const msg = JSON.parse(line);
          if (onMessage) {
            onMessage(msg);
          }
          if (msg.type === 'result') {
            result = msg;
          } else if (msg.type === 'error') {
            fail(new Error(msg.message));
            return;
          }
        } catch (e) {}
      }
    });

    proc.stderr.on('data', (data) => {
      console.error('stderr:', data.toString());
    });

    proc.on('close', (code) => {
      if (settled) return;
      if (code === 0 && result) {
        succeed(result);
      } else if (!result) {
        fail(new Error('Processing failed with code ' + code));
      }
    });

    proc.on('error', (err) => {
      fail(err);
    });
  });
}


// Process signature packets
// Accepts either a folder path string (legacy) or an object { folder: string } or { files: string[] }
ipcMain.handle('process-folder', async (event, input) => {
  mainWindow.webContents.send('progress', { percent: 5, message: 'Initializing signature packet processor...' });
  return runSignaturePacketProcessor(input, {}, (msg) => {
    if (msg.type === 'progress') {
      mainWindow.webContents.send('progress', msg);
    }
  });
});

ipcMain.handle('preflight-signature-packets', async (event, input) => {
  mainWindow.webContents.send('progress', { percent: 2, message: 'Checking document structure...' });
  return runSignaturePacketProcessor(input, {
    preflight_long_annex_check: true,
    long_annex_threshold: 100
  }, (msg) => {
    if (msg.type === 'progress') {
      mainWindow.webContents.send('progress', msg);
    }
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
    description: 'Test iManage COM connectivity. Run this first if having connection issues. Returns diagnostic info about the COM object, login status, available methods, and detected API type (IManDMS vs WorkObjectFactory).',
    input_schema: {
      type: 'object',
      properties: {}
    }
  },
  {
    name: 'imanage_create_folder',
    description: 'Create a new folder/workspace in iManage.',
    input_schema: {
      type: 'object',
      properties: {
        folder_name: { type: 'string', description: 'Name of the folder to create' },
        parent_path: { type: 'string', description: 'Optional parent folder path or ID to create the folder under' }
      },
      required: ['folder_name']
    }
  },
  {
    name: 'imanage_redline_versions',
    description: 'Redline two versions of an iManage document using Litera Compare and optionally email the result. Checks out both versions from iManage, runs a Litera full-document comparison, and can send the redline via Outlook. Use when the user wants to compare version 1 to version 2 (or any two versions) of an iManage document.',
    input_schema: {
      type: 'object',
      properties: {
        profile_id: { type: 'string', description: 'The iManage profile ID (document number) to redline' },
        version_1: { type: 'string', description: 'Version number for the original/older side (e.g. "1"). If omitted, auto-detects the earliest version.' },
        version_2: { type: 'string', description: 'Version number for the modified/newer side (e.g. "2"). If omitted, auto-detects the latest version.' },
        email_to: { type: 'string', description: 'Email recipient(s) to send the redline to (semicolon-separated). If omitted, the redline is saved locally without emailing.' },
        email_cc: { type: 'string', description: 'CC recipients (semicolon-separated)' },
        email_subject: { type: 'string', description: 'Custom email subject. If omitted, auto-generates from document name and versions.' },
        email_body: { type: 'string', description: 'Custom email body (HTML). If omitted, auto-generates.' },
        output_format: { type: 'string', enum: ['pdf', 'docx'], description: 'Output format for the redline. Default: pdf.' },
        output_folder: { type: 'string', description: 'Local folder to save the redline to. Default: Downloads/EmmaNeigh_Redlines.' },
        change_pages_only: { type: 'boolean', description: 'If true, generates a Change Pages Only (CPO) redline — only pages with differences are included. Default: false (full document redline).' }
      },
      required: ['profile_id']
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
        engine: { type: 'string', enum: ['auto', 'litera', 'emmaneigh'], description: 'Comparison engine. Default: auto (tries Litera first, falls back to EmmaNeigh).' },
        change_pages_only: { type: 'boolean', description: 'If true, generates a Change Pages Only (CPO) redline — only pages with differences. Default: false.' }
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
          enum: ['packets', 'packetshell', 'execution', 'sigblocks', 'email', 'updatechecklist', 'punchlist', 'collate', 'redline'],
          description: 'The tab to navigate to. packets=Signature Packets, packetshell=Packet Shell, execution=Execution, sigblocks=Signature Blocks, email=Email Search, updatechecklist=Update Checklist, punchlist=Punchlist Generator, collate=Collate Comments, redline=Redline Documents'
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
  },
  {
    name: 'outlook_reply_email',
    description: 'Reply (or reply-all) to an existing Outlook email thread. Use when the user wants to respond to a specific email. Requires the entry_id from a previous outlook_search or outlook_read_email result.',
    input_schema: {
      type: 'object',
      properties: {
        entry_id: { type: 'string', description: 'The Outlook EntryID of the email to reply to' },
        body: { type: 'string', description: 'Reply body text (HTML supported). This will be prepended above the original message.' },
        reply_all: { type: 'boolean', description: 'If true, reply to all recipients. Default: false (reply to sender only).' },
        attachments: {
          type: 'array',
          items: { type: 'string' },
          description: 'Array of local file paths to attach to the reply'
        }
      },
      required: ['entry_id', 'body']
    }
  }
];

// ========== iMANAGE COM INTEGRATION ==========

const IMANAGE_PS_BOOTSTRAP = `
# ── iManage COM bootstrap ────────────────────────────────────────────────
# Supports iManage Work 10.x (IManDMS), 9.x (WorkObjectFactory), FileSite, and DeskSite COM variants.
# Detects COM type and provides unified + type-specific helpers.

# Global COM type tracking — set by Get-IManageWorkObjectFactory
$global:IManageProgId = ''
$global:IManageCOMType = ''  # 'IManDMS', 'WorkObjectFactory', 'WorkSite'
$global:IManageCandidateDiagnostics = @()
$global:IManageSelectionDetails = @{}

function Get-IManageMethodNames($obj) {
  try {
    return @($obj.PSObject.Methods | ForEach-Object { $_.Name } | Sort-Object -Unique)
  } catch {
    return @()
  }
}

function Get-IManagePropertyNames($obj) {
  try {
    return @($obj.PSObject.Properties | ForEach-Object { $_.Name } | Sort-Object -Unique)
  } catch {
    return @()
  }
}

function Test-IManageDMSCapabilities($obj) {
  $methods = Get-IManageMethodNames $obj
  $properties = Get-IManagePropertyNames $obj

  $hasCreateProfileSearch = ($methods -contains 'CreateProfileSearchParameters')
  $hasDmsSessionSignals = (($properties -contains 'Sessions') -or ($properties -contains 'Databases'))
  $hasLegacyFactorySignals =
    ($methods -contains 'FindProfiles') -or
    ($methods -contains 'GetFiles') -or
    ($methods -contains 'SaveAsFiles') -or
    ($methods -contains 'CheckOutFiles')

  # Some iManage Work 10 COM wrappers expose DMS APIs even under WorkObjectFactory ProgIDs.
  if ($hasCreateProfileSearch -and $hasDmsSessionSignals) { return $true }
  if ($hasCreateProfileSearch -and -not $hasLegacyFactorySignals) { return $true }
  return $false
}

function Resolve-IManageCOMType($obj, $progId) {
  try {
    if (Test-IManageDMSCapabilities $obj) { return 'IManDMS' }
  } catch {}

  if ($progId -match 'WorkSite') { return 'WorkSite' }
  return 'WorkObjectFactory'
}

function Get-IManageCapabilitySnapshot($obj, $progId) {
  $methods = Get-IManageMethodNames $obj
  $properties = Get-IManagePropertyNames $obj

  $searchMethodSet = @('FindProfiles', 'SearchProfiles', 'CreateProfileSearchParameters', 'GetFiles')
  $writeMethodSet = @('SaveAsFiles', 'SaveAsNewVersion', 'SaveFiles', 'ImportDocument', 'AddDocument', 'CheckInFiles')
  $folderMethodSet = @('CreateFolder', 'MakeFolder', 'NewFolder', 'AddFolder', 'CreateWorkspace')
  $keyMethodSet = @(
    'CreateProfileSearchParameters',
    'FindProfiles',
    'GetFiles',
    'SaveAsFiles',
    'SaveAsNewVersion',
    'SaveFiles',
    'CheckOutFiles',
    'CheckInFiles',
    'ImportDocument',
    'AddDocument',
    'CreateFolder',
    'MakeFolder',
    'NewFolder',
    'AddFolder',
    'CreateWorkspace'
  )
  $keyPropertySet = @('Sessions', 'Databases', 'Connected', 'HasLogin')

  $keyMethods = @($methods | Where-Object { $_ -in $keyMethodSet })
  $keyProperties = @($properties | Where-Object { $_ -in $keyPropertySet })

  $hasDmsCapability = $false
  try { $hasDmsCapability = Test-IManageDMSCapabilities $obj } catch { $hasDmsCapability = $false }

  $hasSearchOps = @($methods | Where-Object { $_ -in $searchMethodSet }).Count -gt 0
  $hasWriteOps = @($methods | Where-Object { $_ -in $writeMethodSet }).Count -gt 0
  $hasFolderOps = @($methods | Where-Object { $_ -in $folderMethodSet }).Count -gt 0

  $score = 0
  if ($hasDmsCapability) { $score += 500 }
  if ($progId -match 'IManDMS') { $score += 250 }
  if ($methods -contains 'CreateProfileSearchParameters') { $score += 150 }
  if (($properties -contains 'Sessions') -or ($properties -contains 'Databases')) { $score += 120 }
  if ($hasSearchOps) { $score += 90 }
  if ($hasWriteOps) { $score += 120 }
  if ($hasFolderOps) { $score += 80 }
  if ($methods.Count -ge 20) { $score += 60 }
  if ($methods.Count -le 8) { $score -= 220 }
  if ($progId -match 'iwComWrapper') { $score -= 200 }

  return @{
    score = [int]$score
    method_count = [int]$methods.Count
    property_count = [int]$properties.Count
    has_dms_capability = [bool]$hasDmsCapability
    has_search_ops = [bool]$hasSearchOps
    has_write_ops = [bool]$hasWriteOps
    has_folder_ops = [bool]$hasFolderOps
    key_methods = $keyMethods
    key_properties = $keyProperties
  }
}

function Test-IManageLimitedApiSurface($snapshot, $apiType, $progId) {
  if ($null -eq $snapshot) { return $true }
  if ($apiType -eq 'IManDMS') { return $false }
  if ($snapshot.has_write_ops -or $snapshot.has_folder_ops -or $snapshot.has_search_ops) {
    if ($snapshot.method_count -gt 8) { return $false }
  }
  if (($progId -match 'iwComWrapper') -and ($snapshot.method_count -le 10)) { return $true }
  if ($snapshot.method_count -le 8) { return $true }
  return $false
}

function Get-IManageCreateObjectCandidates($factoryObj, $factoryProgId) {
  $candidates = @()
  $createMethod = $null
  try { $createMethod = $factoryObj.PSObject.Methods['CreateObject'] } catch {}
  if ($null -eq $createMethod) { return @() }

  $targets = @(
    'iManage.COMAPILib.IManDMS',
    'iManage.Work.WorkObjectFactory',
    'Com.iManage.Work.WorkObjectFactory',
    'iManage.WorkSiteObjects.WorkObjectFactory',
    'WorkSite.Application',
    'iManage.WorkSite.Application'
  )

  foreach ($target in $targets) {
    try {
      $obj = $factoryObj.CreateObject($target)
      if ($null -eq $obj) { continue }
      $candidates += @{
        obj = $obj
        candidate_id = ($factoryProgId + '::CreateObject(' + $target + ')')
        source_prog_id = $factoryProgId
        create_target = $target
      }
    } catch {}
  }

  return $candidates
}

function Use-IManageDMSApi($wof) {
  if ($global:IManageCOMType -eq 'IManDMS') { return $true }
  try {
    if (Test-IManageDMSCapabilities $wof) {
      $global:IManageCOMType = 'IManDMS'
      return $true
    }
  } catch {}
  return $false
}

function Assert-IManageOperationSupport($wof, $operationName) {
  if (Use-IManageDMSApi $wof) { return }

  $methods = Get-IManageMethodNames $wof
  $requirements = @{
    browse = @('FindProfiles', 'GetFiles', 'SearchProfiles')
    search = @('FindProfiles', 'SearchProfiles')
    save = @('SaveAsFiles', 'SaveAsNewVersion', 'SaveFiles', 'ImportDocument', 'AddDocument', 'CheckInFiles')
    checkout = @('GetFiles', 'CheckOutFiles', 'CopyFiles')
    checkin = @('CheckInFiles', 'SaveFiles', 'SaveAsFiles', 'ImportDocument', 'AddDocument')
    versions = @('GetDocumentVersions', 'FindProfiles', 'GetFiles')
    create_folder = @('CreateFolder', 'MakeFolder', 'NewFolder', 'AddFolder', 'CreateWorkspace')
  }

  $requiredMethods = $requirements[$operationName]
  if ($null -eq $requiredMethods -or $requiredMethods.Count -eq 0) { return }

  $availableRequiredMethods = @($methods | Where-Object { $_ -in $requiredMethods })
  if ($availableRequiredMethods.Count -gt 0) { return }

  $progId = if ($global:IManageProgId) { $global:IManageProgId } else { 'unknown' }
  $previewMethods = @($methods | Select-Object -First 12)
  $preview = if ($previewMethods.Count -gt 0) { $previewMethods -join ', ' } else { 'none' }
  throw ("iManage COM is connected through a limited API surface ($progId) that does not support '$operationName'. Available methods: $preview. Open iManage Work Desktop and sign in, then retry. If this persists, repair/reinstall iManage Work Desktop COM integration so IManDMS or full WorkObjectFactory methods are available.")
}

function Get-IManageWorkObjectFactory {
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
  $candidates = @()
  $bestObj = $null
  $bestMeta = $null
  $global:IManageCandidateDiagnostics = @()
  $global:IManageSelectionDetails = @{}

  foreach ($progId in $progIds) {
    try {
      $obj = New-Object -ComObject $progId
      if ($null -ne $obj) {
        $candidateObjects = @(
          @{
            obj = $obj
            candidate_id = $progId
            source_prog_id = ''
            create_target = ''
          }
        )
        $viaFactoryObjects = Get-IManageCreateObjectCandidates $obj $progId
        foreach ($extraCandidate in $viaFactoryObjects) {
          $candidateObjects += $extraCandidate
        }

        foreach ($candidate in $candidateObjects) {
          $candidateObj = $candidate.obj
          $candidateDisplayId = [string]$candidate.candidate_id
          $createTarget = [string]$candidate.create_target
          $resolveProgId = if ($createTarget) { $createTarget } else { $candidateDisplayId }
          $apiType = Resolve-IManageCOMType $candidateObj $resolveProgId
          $snapshot = Get-IManageCapabilitySnapshot $candidateObj $resolveProgId
          $limitedSurface = Test-IManageLimitedApiSurface $snapshot $apiType $resolveProgId
          $candidateMeta = @{
            prog_id = $candidateDisplayId
            resolved_prog_id = $resolveProgId
            source_prog_id = [string]$candidate.source_prog_id
            via_create_object = [bool]$createTarget
            create_target = $createTarget
            api_type = $apiType
            score = [int]$snapshot.score
            method_count = [int]$snapshot.method_count
            property_count = [int]$snapshot.property_count
            key_methods = $snapshot.key_methods
            key_properties = $snapshot.key_properties
            limited_surface = [bool]$limitedSurface
          }
          $candidates += $candidateMeta

          $useCandidate = $false
          if ($null -eq $bestMeta) {
            $useCandidate = $true
          } elseif ($snapshot.score -gt $bestMeta.score) {
            $useCandidate = $true
          } elseif (($snapshot.score -eq $bestMeta.score) -and ($snapshot.method_count -gt $bestMeta.method_count)) {
            $useCandidate = $true
          }

          if ($useCandidate) {
            $bestObj = $candidateObj
            $bestMeta = $candidateMeta
          }
        }
      }
    } catch {
      $errors += ($progId + ': ' + $_.Exception.Message)
    }
  }

  $global:IManageCandidateDiagnostics = $candidates
  if ($null -ne $bestObj) {
    $global:IManageProgId = $bestMeta.prog_id
    $global:IManageCOMType = $bestMeta.api_type
    $global:IManageSelectionDetails = $bestMeta
    return $bestObj
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
    comType = $global:IManageCOMType
    progId = $global:IManageProgId
  }
  try {
    $info.methods = @($wof.PSObject.Methods | ForEach-Object { $_.Name } | Sort-Object -Unique)
  } catch {}
  try {
    $info.properties = @($wof.PSObject.Properties | ForEach-Object { $_.Name } | Sort-Object -Unique)
  } catch {}
  return $info
}

# ── IManDMS session-based helpers (iManage Work 10.x) ──────────────────
# Uses the actual iManage COM API: IManDMS → Sessions → Database → SearchDocuments

# ── iManage COM enum constants ──
# imProfileAttributeID enum values (for CreateProfileSearchParameters.Add)
$script:imProfileDocnum     = 1
$script:imProfileVersion    = 2
$script:imProfileName       = 3
$script:imProfileDescription = 4
$script:imProfileAuthor     = 5
# imGetCopyOptions enum
$script:imNativeFormat      = 0
# imCheckinDisposition enum
$script:imCheckinNewVersion = 1
$script:imCheckinNewDocument = 0
# imCheckinOptions enum
$script:imDontKeepCheckedOut = 0
$script:imKeepCheckedOut     = 1
# Try to load actual enum types from the COM type library
try { $script:imProfileDocnum = [int][iManage.imProfileAttributeID]::imProfileDocnum } catch {}
try { $script:imProfileName = [int][iManage.imProfileAttributeID]::imProfileName } catch {}
try { $script:imProfileDescription = [int][iManage.imProfileAttributeID]::imProfileDescription } catch {}
try { $script:imNativeFormat = [int][iManage.imGetCopyOptions]::imNativeFormat } catch {}
try { $script:imCheckinNewVersion = [int][iManage.imCheckinDisposition]::imCheckinNewVersion } catch {}

function Get-IManDMSSession($dms) {
  $errors = @()
  # Try 1: Enumerate existing sessions (iManage Desktop should provide one)
  try {
    $sessions = $dms.Sessions
    if ($null -ne $sessions) {
      # Try PowerShell-native foreach enumeration first (most reliable for COM)
      try {
        foreach ($s in $sessions) {
          if ($null -ne $s) { return $s }
        }
      } catch { $errors += ('foreach Sessions: ' + $_.Exception.Message) }
      # Try 1-based COM index
      try {
        $cnt = $sessions.Count
        if ($cnt -gt 0) {
          for ($i = 1; $i -le $cnt; $i++) {
            try {
              $s = $sessions.Item($i)
              if ($null -ne $s) { return $s }
            } catch {}
          }
        }
      } catch { $errors += ('Sessions.Item: ' + $_.Exception.Message) }
    }
  } catch { $errors += ('Sessions access: ' + $_.Exception.Message) }
  # Try 2: Add a new session (may need server name — try empty first)
  try {
    $session = $dms.Sessions.Add('')
    if ($null -ne $session) {
      try { $session.TrustedLogin() } catch {}
      return $session
    }
  } catch { $errors += ('Sessions.Add(empty): ' + $_.Exception.Message) }
  try {
    $session = $dms.Sessions.Add()
    if ($null -ne $session) {
      try { $session.TrustedLogin() } catch {}
      return $session
    }
  } catch { $errors += ('Sessions.Add(): ' + $_.Exception.Message) }
  throw ('Could not get or create iManage DMS session. Methods tried: ' + ($errors -join ' | ') + '. Ensure iManage Work Desktop is running and you are signed in.')
}

function Get-IManDMSDatabase($session) {
  $errors = @()
  # Try 1: Enumerate Databases collection with foreach
  try {
    $dbs = $session.Databases
    if ($null -ne $dbs) {
      try {
        foreach ($db in $dbs) {
          if ($null -ne $db) { return $db }
        }
      } catch { $errors += ('foreach Databases: ' + $_.Exception.Message) }
      # Try 1-based COM index
      try {
        $cnt = $dbs.Count
        if ($cnt -gt 0) {
          for ($i = 1; $i -le $cnt; $i++) {
            try {
              $db = $dbs.Item($i)
              if ($null -ne $db) { return $db }
            } catch {}
          }
        }
      } catch { $errors += ('Databases.Item: ' + $_.Exception.Message) }
    }
  } catch { $errors += ('Databases access: ' + $_.Exception.Message) }
  # Try 2: PreferredDatabase property
  try {
    $pdb = $session.PreferredDatabase
    if ($null -ne $pdb) { return $pdb }
  } catch { $errors += ('PreferredDatabase: ' + $_.Exception.Message) }
  # Try 3: WorkArea property
  try {
    $wa = $session.WorkArea
    if ($null -ne $wa) { return $wa }
  } catch { $errors += ('WorkArea: ' + $_.Exception.Message) }
  throw ('Could not access iManage database. Methods tried: ' + ($errors -join ' | ') + '. Ensure you are connected to an iManage library.')
}

function Get-IManDMSDocById($dms, $docNumber, $versionNum) {
  $session = Get-IManDMSSession $dms
  $db = Get-IManDMSDatabase $session
  $errors = @()
  # Try 1: Search using CreateProfileSearchParameters on DMS object (correct API)
  try {
    $params = $dms.CreateProfileSearchParameters()
    if ($null -ne $params) {
      $params.Add($script:imProfileDocnum, [string]$docNumber)
      $found = $db.SearchDocuments($params, $true)
      if ($null -ne $found) {
        foreach ($doc in $found) {
          if ($null -ne $doc) {
            # If specific version requested, try to access it
            if ($null -ne $versionNum -and $versionNum -gt 0) {
              try {
                $versions = $doc.Versions
                if ($null -ne $versions) {
                  foreach ($v in $versions) {
                    try {
                      if ($null -ne $v -and $v.Version -eq $versionNum) { return $v }
                    } catch {}
                  }
                }
              } catch {}
            }
            return $doc
          }
        }
      }
    }
  } catch { $errors += ('CreateProfileSearchParameters+SearchDocuments: ' + $_.Exception.Message) }
  # Try 2: GetDocument on database (older API, may work on some installations)
  if ($null -ne $versionNum -and $versionNum -gt 0) {
    try {
      $doc = $db.GetDocument([int]$docNumber, [int]$versionNum)
      if ($null -ne $doc) { return $doc }
    } catch { $errors += ('GetDocument(num,ver): ' + $_.Exception.Message) }
  }
  try {
    $doc = $db.GetDocument([int]$docNumber)
    if ($null -ne $doc) { return $doc }
  } catch { $errors += ('GetDocument(int): ' + $_.Exception.Message) }
  try {
    $doc = $db.GetDocument([string]$docNumber)
    if ($null -ne $doc) { return $doc }
  } catch { $errors += ('GetDocument(str): ' + $_.Exception.Message) }
  throw ('Could not retrieve document ' + $docNumber + '. Errors: ' + ($errors -join ' | '))
}

function Search-IManDMSDocs($dms, $queryText, $maxResults) {
  if (-not $maxResults) { $maxResults = 25 }
  $session = Get-IManDMSSession $dms
  $db = Get-IManDMSDatabase $session
  $errors = @()
  # Try 1: CreateProfileSearchParameters on DMS object (correct API)
  try {
    $params = $dms.CreateProfileSearchParameters()
    if ($null -ne $params) {
      # Add search criteria — try description and name fields
      try { $params.Add($script:imProfileDescription, $queryText) } catch {}
      try { $params.Add($script:imProfileName, $queryText) } catch {}
      $found = $db.SearchDocuments($params, $true)
      if ($null -ne $found) {
        $results = @()
        foreach ($doc in $found) { if ($null -ne $doc) { $results += $doc } }
        if ($results.Count -gt 0) { return $results }
      }
    }
  } catch { $errors += ('DMS.CreateProfileSearchParameters: ' + $_.Exception.Message) }
  # Try 2: CreateSearchParameters on Database (some older versions)
  try {
    $params = $db.CreateSearchParameters()
    if ($null -ne $params) {
      try { $params.Add($script:imProfileDescription, $queryText) } catch {
        try { $params.Add('DESCRIPTION', $queryText) } catch {}
      }
      try { $params.Add($script:imProfileName, $queryText) } catch {
        try { $params.Add('NAME', $queryText) } catch {}
      }
      $found = $db.SearchDocuments($params, $true)
      if ($null -ne $found) {
        $results = @()
        foreach ($doc in $found) { if ($null -ne $doc) { $results += $doc } }
        if ($results.Count -gt 0) { return $results }
      }
    }
  } catch { $errors += ('db.CreateSearchParameters: ' + $_.Exception.Message) }
  # Try 3: QuickSearch on database
  try {
    $found = $db.QuickSearch($queryText, $maxResults)
    if ($null -ne $found) {
      $results = @()
      foreach ($doc in $found) { if ($null -ne $doc) { $results += $doc } }
      if ($results.Count -gt 0) { return $results }
    }
  } catch { $errors += ('QuickSearch: ' + $_.Exception.Message) }
  if ($errors.Count -gt 0) {
    throw ('iManage search failed. Methods tried: ' + ($errors -join ' | '))
  }
  return @()
}

function Get-IManDMSDocCopy($doc, $destPath) {
  # Download a document to a local path using the correct COM API
  $errors = @()
  # Try 1: GetCopy with native format option (correct API)
  try {
    $doc.GetCopy($destPath, $script:imNativeFormat)
    if (Test-Path $destPath) { return $true }
  } catch { $errors += ('GetCopy(path,format): ' + $_.Exception.Message) }
  # Try 2: GetCopy with single param
  try {
    $doc.GetCopy($destPath)
    if (Test-Path $destPath) { return $true }
  } catch { $errors += ('GetCopy(path): ' + $_.Exception.Message) }
  # Try 3: CheckOut methods (may work on some versions)
  try {
    $doc.CheckOut($destPath, 1)
    if (Test-Path $destPath) { return $true }
  } catch { $errors += ('CheckOut(path,mode): ' + $_.Exception.Message) }
  try {
    $doc.CheckOut($destPath)
    if (Test-Path $destPath) { return $true }
  } catch { $errors += ('CheckOut(path): ' + $_.Exception.Message) }
  # Try 4: Copy method
  try {
    $doc.Copy($destPath)
    if (Test-Path $destPath) { return $true }
  } catch { $errors += ('Copy: ' + $_.Exception.Message) }
  throw ('Could not download document. Methods tried: ' + ($errors -join ' | '))
}

function Set-IManDMSDocCheckin($doc, $sourcePath, $comment) {
  # Check in a document using the correct COM API
  if (-not $comment) { $comment = 'Checked in via EmmaNeigh' }
  $errors = @()
  # Try 1: CheckInWithResults (correct iManage COM API)
  try {
    $result = $doc.CheckInWithResults($sourcePath, $script:imCheckinNewVersion, $script:imDontKeepCheckedOut)
    return $true
  } catch { $errors += ('CheckInWithResults: ' + $_.Exception.Message) }
  # Try 2: CheckInEx (extended checkin with full params)
  try {
    $checkInResults = $null
    $doc.CheckInEx($sourcePath, $script:imCheckinNewVersion, $script:imDontKeepCheckedOut, 0, 'EmmaNeigh', $comment, $sourcePath, [ref]$checkInResults)
    return $true
  } catch { $errors += ('CheckInEx: ' + $_.Exception.Message) }
  # Try 3: Simple CheckIn (older API variants)
  try {
    $doc.CheckIn($sourcePath, $comment)
    return $true
  } catch { $errors += ('CheckIn(path,comment): ' + $_.Exception.Message) }
  try {
    $doc.CheckIn($sourcePath)
    return $true
  } catch { $errors += ('CheckIn(path): ' + $_.Exception.Message) }
  # Try 4: UpdateFromFile
  try {
    $doc.UpdateFromFile($sourcePath)
    return $true
  } catch { $errors += ('UpdateFromFile: ' + $_.Exception.Message) }
  throw ('Could not check in document. Methods tried: ' + ($errors -join ' | '))
}

function Import-IManDMSDoc($dms, $localPath) {
  $session = Get-IManDMSSession $dms
  $db = Get-IManDMSDatabase $session
  $errors = @()
  # Try 1: CreateDocument + set properties + CheckInWithResults
  try {
    $doc = $db.CreateDocument()
    if ($null -ne $doc) {
      $fileName = [System.IO.Path]::GetFileNameWithoutExtension($localPath)
      $ext = [System.IO.Path]::GetExtension($localPath).TrimStart('.')
      try { $doc.SetAttributeByID($script:imProfileName, $fileName) } catch {
        try { $doc.Name = $fileName } catch {}
      }
      try { $doc.Type = $ext } catch {
        try { $doc.Extension = $ext } catch {}
      }
      try {
        Set-IManDMSDocCheckin $doc $localPath 'Imported via EmmaNeigh'
        return $doc
      } catch { $errors += ('CreateDocument+Checkin: ' + $_.Exception.Message) }
    }
  } catch { $errors += ('CreateDocument: ' + $_.Exception.Message) }
  # Try 2: ImportDocument on database
  try {
    $doc = $db.ImportDocument($localPath)
    if ($null -ne $doc) { return $doc }
  } catch { $errors += ('db.ImportDocument: ' + $_.Exception.Message) }
  # Try 3: AddDocument
  try {
    $doc = $db.AddDocument($localPath)
    if ($null -ne $doc) { return $doc }
  } catch { $errors += ('AddDocument: ' + $_.Exception.Message) }
  throw ('Could not import document to iManage. Methods tried: ' + ($errors -join ' | '))
}

function Extract-IManDocInfo($doc) {
  $info = @{}
  try { $info.name = $doc.Name } catch { $info.name = '' }
  try { $info.number = $doc.Number } catch {
    try { $info.number = $doc.DocNumber } catch { $info.number = '' }
  }
  try { $info.version = $doc.Version } catch { $info.version = '' }
  try { $info.extension = $doc.Extension } catch {
    try { $info.extension = $doc.Type } catch { $info.extension = '' }
  }
  try { $info.author = $doc.Author } catch { $info.author = '' }
  try { $info.description = $doc.Description } catch { $info.description = '' }
  try { $info.path = $doc.Path } catch { $info.path = '' }
  try { $info.class_name = $doc.Class } catch { $info.class_name = '' }
  try { $info.database = $doc.Database } catch { $info.database = '' }
  return $info
}

function Create-IManDMSFolder($dms, $folderName, $parentPath) {
  $session = Get-IManDMSSession $dms
  $db = Get-IManDMSDatabase $session
  $errors = @()
  $parent = $null
  if ($parentPath) {
    try { $parent = $db.GetFolder($parentPath) } catch {}
    try { if ($null -eq $parent) { $parent = $db.GetFolderByPath($parentPath) } } catch {}
  }
  if ($null -ne $parent) {
    try {
      $folder = $parent.SubFolders.Add($folderName)
      if ($null -ne $folder) { return $folder }
    } catch { $errors += ('parent.SubFolders.Add: ' + $_.Exception.Message) }
    try {
      $folder = $parent.CreateSubFolder($folderName)
      if ($null -ne $folder) { return $folder }
    } catch { $errors += ('parent.CreateSubFolder: ' + $_.Exception.Message) }
  }
  try {
    $folder = $db.CreateFolder($folderName)
    if ($null -ne $folder) { return $folder }
  } catch { $errors += ('db.CreateFolder: ' + $_.Exception.Message) }
  try {
    $folder = $db.Folders.Add($folderName)
    if ($null -ne $folder) { return $folder }
  } catch { $errors += ('db.Folders.Add: ' + $_.Exception.Message) }
  try {
    $root = $db.RootFolder
    if ($null -ne $root) {
      $folder = $root.SubFolders.Add($folderName)
      if ($null -ne $folder) { return $folder }
    }
  } catch { $errors += ('RootFolder.SubFolders.Add: ' + $_.Exception.Message) }
  throw ('Could not create folder. Methods tried: ' + ($errors -join ' | '))
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
  if (/^drive:/i.test(text)) return text;
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

function isIManageDriveProfileId(rawValue) {
  return /^drive:/i.test(String(rawValue || '').trim());
}

function encodeIManageDriveProfileId(filePath) {
  const normalized = normalizeLocalPath(filePath);
  return normalized ? `drive:${normalized}` : '';
}

function decodeIManageDriveProfileId(profileId) {
  const text = String(profileId || '').trim();
  if (!/^drive:/i.test(text)) return '';
  return normalizeLocalPath(text.replace(/^drive:/i, ''));
}

function isDirectoryPath(candidatePath) {
  try {
    return !!candidatePath && fs.existsSync(candidatePath) && fs.statSync(candidatePath).isDirectory();
  } catch (_) {
    return false;
  }
}

function getConfiguredIManageDriveRoot() {
  const configured = resolveExistingLocalPath(
    getSettingValue(IMANAGE_DRIVE_ROOT_KEY, process.env.IMANAGE_DRIVE_ROOT || '')
  );
  return isDirectoryPath(configured) ? path.resolve(configured) : '';
}

function collectNamedIManageDriveRoots(baseDir, output, seen) {
  if (!isDirectoryPath(baseDir)) return;
  let entries = [];
  try {
    entries = fs.readdirSync(baseDir, { withFileTypes: true });
  } catch (_) {
    return;
  }
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (!/imanage/i.test(entry.name)) continue;
    const fullPath = path.resolve(path.join(baseDir, entry.name));
    const key = fullPath.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(fullPath);
  }
}

function detectIManageDriveRoots() {
  const roots = [];
  const seen = new Set();

  const addRoot = (candidate) => {
    const resolved = resolveExistingLocalPath(candidate);
    if (!isDirectoryPath(resolved)) return;
    const fullPath = path.resolve(resolved);
    const key = fullPath.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    roots.push(fullPath);
  };

  const configured = getConfiguredIManageDriveRoot();
  if (configured) addRoot(configured);

  const homeDir = os.homedir();
  const oneDriveRoots = [
    process.env.OneDrive,
    process.env.OneDriveCommercial,
    process.env.OneDriveConsumer,
    path.join(homeDir, 'OneDrive'),
    path.join(homeDir, 'Library', 'CloudStorage')
  ].filter(Boolean);
  const baseDirs = [
    homeDir,
    path.join(homeDir, 'Documents'),
    ...oneDriveRoots,
    ...oneDriveRoots.map((rootDir) => path.join(rootDir, 'Documents'))
  ];

  for (const candidate of [
    path.join(homeDir, 'iManage'),
    path.join(homeDir, 'iManage Drive'),
    path.join(homeDir, 'Documents', 'iManage'),
    path.join(homeDir, 'Documents', 'iManage Drive')
  ]) {
    addRoot(candidate);
  }

  for (const baseDir of baseDirs) {
    collectNamedIManageDriveRoots(baseDir, roots, seen);
  }

  return roots;
}

async function promptForIManageDriveRoot(existingRoots = []) {
  const defaultPath = existingRoots[0] || getConfiguredIManageDriveRoot() || path.join(os.homedir(), 'Documents');
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select your iManage Drive sync folder',
    defaultPath,
    properties: ['openDirectory']
  });
  if (result.canceled || !result.filePaths || !result.filePaths[0]) return '';
  const selected = resolveExistingLocalPath(result.filePaths[0]);
  if (!isDirectoryPath(selected)) return '';
  writeSettingValue(IMANAGE_DRIVE_ROOT_KEY, selected);
  try { saveDatabase(); } catch (_) {}
  return path.resolve(selected);
}

async function getAvailableIManageDriveRoots(options = {}) {
  const allowPrompt = !!options.allowPrompt;
  const roots = detectIManageDriveRoots();
  if (roots.length || !allowPrompt) return roots;
  const selected = await promptForIManageDriveRoot(roots);
  return selected ? [selected] : [];
}

function isWithinIManageDriveRoot(filePath, rootPath) {
  const resolvedFile = resolveExistingLocalPath(filePath);
  const resolvedRoot = resolveExistingLocalPath(rootPath);
  if (!resolvedFile || !resolvedRoot) return false;
  try {
    const rel = path.relative(path.resolve(resolvedRoot), path.resolve(resolvedFile));
    return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
  } catch (_) {
    return false;
  }
}

function getIManageDriveRootForPath(filePath, roots = null) {
  const candidates = Array.isArray(roots) && roots.length ? roots : detectIManageDriveRoots();
  return candidates.find((rootPath) => isWithinIManageDriveRoot(filePath, rootPath)) || '';
}

function buildIManageDriveFileRecord(filePath, rootPath = '') {
  const resolvedPath = resolveExistingLocalPath(filePath);
  const matchedRoot = rootPath || getIManageDriveRootForPath(resolvedPath) || '';
  let stats = null;
  try { stats = fs.statSync(resolvedPath); } catch (_) {}
  return {
    mode: 'drive',
    profile_id: encodeIManageDriveProfileId(resolvedPath),
    path: resolvedPath,
    root_path: matchedRoot,
    name: path.basename(resolvedPath),
    description: matchedRoot ? path.relative(matchedRoot, resolvedPath) : resolvedPath,
    version: 'local-sync',
    modified_at: stats ? new Date(stats.mtimeMs).toISOString() : '',
    size: stats ? stats.size : 0
  };
}

function pickUniqueIManageDriveTargetPath(targetPath) {
  if (!fs.existsSync(targetPath)) return targetPath;
  const dirName = path.dirname(targetPath);
  const ext = path.extname(targetPath);
  const stem = path.basename(targetPath, ext);
  for (let i = 2; i <= 999; i += 1) {
    const candidate = path.join(dirName, `${stem} (${i})${ext}`);
    if (!fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return targetPath;
}

async function imanageDriveBrowseFiles(multiple = false) {
  const roots = await getAvailableIManageDriveRoots({ allowPrompt: true });
  if (!roots.length) {
    return {
      success: false,
      error: 'No iManage Drive sync folder was found. Select your synced iManage folder and retry.'
    };
  }
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Choose files from your iManage Drive sync folder',
    defaultPath: roots[0],
    properties: multiple ? ['openFile', 'multiSelections'] : ['openFile']
  });
  if (result.canceled || !result.filePaths || !result.filePaths.length) {
    return { success: false, error: 'iManage Drive browse canceled.' };
  }
  const validFilePaths = result.filePaths.filter((filePath) => !!getIManageDriveRootForPath(filePath, roots));
  if (!validFilePaths.length) {
    return {
      success: false,
      error: 'Choose files inside your synced iManage Drive folder.'
    };
  }
  return {
    success: true,
    mode: 'drive',
    files: validFilePaths.map((filePath) => buildIManageDriveFileRecord(filePath, getIManageDriveRootForPath(filePath, roots)))
  };
}

async function imanageDriveSearch(query, options = {}) {
  const roots = await getAvailableIManageDriveRoots({ allowPrompt: !!options.allowPrompt });
  if (!roots.length) {
    return {
      success: false,
      error: 'No iManage Drive sync folder was found. Select your synced iManage folder and retry.'
    };
  }

  const queryText = String(query || '').trim();
  if (!queryText) {
    return { success: false, error: 'query is required for iManage Drive search.' };
  }

  const matches = [];
  for (const rootPath of roots) {
    const files = listFilesRecursively(rootPath, 8);
    for (const filePath of files) {
      const record = buildIManageDriveFileRecord(filePath, rootPath);
      const score = scoreIManageSearchCandidate(queryText, record);
      if (score <= 0) continue;
      matches.push({ ...record, _score: score });
    }
  }

  matches.sort((a, b) => b._score - a._score);
  return {
    success: true,
    mode: 'drive',
    files: matches.slice(0, Number(options.maxResults) || 25)
  };
}

async function imanageDriveSaveDocument(filePath, action = 'new_document', showDialog = true) {
  const resolvedPath = resolveExistingLocalPath(filePath);
  if (!resolvedPath || !fs.existsSync(resolvedPath)) {
    return { success: false, error: 'Local file not found for iManage Drive save.' };
  }

  const roots = await getAvailableIManageDriveRoots({ allowPrompt: !!showDialog });
  if (!roots.length) {
    return {
      success: false,
      error: 'No iManage Drive sync folder was found. Select your synced iManage folder and retry.'
    };
  }

  const currentRoot = getIManageDriveRootForPath(resolvedPath, roots);
  if (action === 'new_version' && currentRoot) {
    return {
      success: true,
      mode: 'drive',
      profile_id: encodeIManageDriveProfileId(resolvedPath),
      path: resolvedPath,
      files: [buildIManageDriveFileRecord(resolvedPath, currentRoot)],
      warning: 'File is already inside your synced iManage Drive folder. Saving changes in place will sync back through iManage Drive.'
    };
  }

  let targetPath = '';
  if (showDialog) {
    const defaultPath = path.join(roots[0], path.basename(resolvedPath));
    const dialogResult = await dialog.showSaveDialog(mainWindow, {
      title: action === 'new_version' ? 'Save updated version to iManage Drive' : 'Save document to iManage Drive',
      defaultPath
    });
    if (dialogResult.canceled || !dialogResult.filePath) {
      return { success: false, error: 'iManage Drive save canceled.' };
    }
    targetPath = resolveExistingLocalPath(dialogResult.filePath);
  } else {
    targetPath = path.join(roots[0], path.basename(resolvedPath));
    if (action !== 'new_version') {
      targetPath = pickUniqueIManageDriveTargetPath(targetPath);
    }
  }

  const targetRoot = getIManageDriveRootForPath(targetPath, roots);
  if (!targetRoot) {
    return {
      success: false,
      error: 'Choose a target location inside your iManage Drive sync folder.'
    };
  }

  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.copyFileSync(resolvedPath, targetPath);
  return {
    success: true,
    mode: 'drive',
    profile_id: encodeIManageDriveProfileId(targetPath),
    path: targetPath,
    files: [buildIManageDriveFileRecord(targetPath, targetRoot)],
    warning: action === 'new_version'
      ? 'Saved into your synced iManage Drive folder. Version behavior depends on your firm’s iManage Drive configuration.'
      : 'Saved into your synced iManage Drive folder.'
  };
}

async function imanageDriveCreateFolder(folderName, parentPath = '') {
  const safeName = String(folderName || '').trim().replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!safeName) {
    return { success: false, error: 'folder_name is required.' };
  }

  const roots = await getAvailableIManageDriveRoots({ allowPrompt: true });
  if (!roots.length) {
    return {
      success: false,
      error: 'No iManage Drive sync folder was found. Select your synced iManage folder and retry.'
    };
  }

  const parentCandidate = resolveExistingLocalPath(parentPath);
  const baseDir = parentCandidate && isDirectoryPath(parentCandidate) && getIManageDriveRootForPath(parentCandidate, roots)
    ? parentCandidate
    : roots[0];
  const targetDir = path.join(baseDir, safeName);
  fs.mkdirSync(targetDir, { recursive: true });
  return {
    success: true,
    mode: 'drive',
    folder: {
      name: safeName,
      path: targetDir,
      root_path: getIManageDriveRootForPath(targetDir, roots) || roots[0]
    }
  };
}

async function imanageDriveCheckout(profileId, checkoutPath = '', version = null) {
  const sourcePath = decodeIManageDriveProfileId(profileId);
  if (!sourcePath || !fs.existsSync(sourcePath)) {
    return { success: false, error: 'Synced iManage Drive file not found.' };
  }
  const rootPath = getIManageDriveRootForPath(sourcePath);
  if (!rootPath) {
    return {
      success: false,
      error: 'This file is not inside a synced iManage Drive folder.'
    };
  }
  return {
    success: true,
    mode: 'drive',
    path: sourcePath,
    version_requested: version,
    files: [buildIManageDriveFileRecord(sourcePath, rootPath)],
    warning: version !== null && version !== undefined
      ? 'iManage Drive fallback exposes the synced local file only, not DMS version history.'
      : 'Using the synced iManage Drive file directly. Edit this file in place and iManage Drive will sync it.'
  };
}

async function imanageDriveCheckin(filePath) {
  const resolvedPath = resolveExistingLocalPath(filePath);
  if (!resolvedPath || !fs.existsSync(resolvedPath)) {
    return { success: false, error: 'Local file not found for iManage Drive check-in.' };
  }
  const rootPath = getIManageDriveRootForPath(resolvedPath);
  if (!rootPath) {
    return {
      success: false,
      error: 'This file is not inside a synced iManage Drive folder. Save it into your synced folder first so Drive can upload it.'
    };
  }
  return {
    success: true,
    mode: 'drive',
    path: resolvedPath,
    files: [buildIManageDriveFileRecord(resolvedPath, rootPath)],
    warning: 'File is already inside your synced iManage Drive folder. iManage Drive should sync the updated file automatically.'
  };
}

function shouldTryIManageDriveFallback(errorLike) {
  const text = String(
    errorLike && typeof errorLike === 'object'
      ? (errorLike.error || errorLike.stderr || errorLike.stdout || '')
      : (errorLike || '')
  )
    .toLowerCase()
    .trim();
  if (!text) return false;
  return (
    text.includes('limited api surface') ||
    text.includes('class-factory') ||
    text.includes('class not registered') ||
    text.includes('unable to create imanage com object') ||
    text.includes('retrieving the com class factory') ||
    text.includes('only available on windows') ||
    text.includes('iManage COM could not be loaded'.toLowerCase())
  );
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
  Assert-IManageOperationSupport $wof 'browse'

  if (Use-IManageDMSApi $wof) {
    # IManDMS: browse documents using correct COM API
    $session = Get-IManDMSSession $wof
    $db = Get-IManDMSDatabase $session
    $docs = @()
    $browseErrors = @()
    # Try 1: CreateProfileSearchParameters on DMS + SearchDocuments (correct API)
    try {
      $params = $wof.CreateProfileSearchParameters()
      if ($null -ne $params) {
        try { $params.Add($script:imProfileName, '*') } catch {
          try { $params.Add(3, '*') } catch {}
        }
        $found = $db.SearchDocuments($params, $true)
        if ($null -ne $found) { foreach ($d in $found) { if ($null -ne $d) { $docs += $d } } }
      }
    } catch { $browseErrors += ('DMS.CreateProfileSearchParameters: ' + $_.Exception.Message) }
    # Try 2: RecentDocuments (property on database)
    if ($docs.Count -eq 0) {
      try {
        $recent = $db.RecentDocuments
        if ($null -ne $recent) { foreach ($d in $recent) { if ($null -ne $d) { $docs += $d } } }
      } catch { $browseErrors += ('RecentDocuments: ' + $_.Exception.Message) }
    }
    # Try 3: Fallback CreateSearchParameters on database (older API)
    if ($docs.Count -eq 0) {
      try {
        $params = $db.CreateSearchParameters()
        if ($null -ne $params) {
          try { $params.Add($script:imProfileName, '*') } catch {
            try { $params.Add('NAME', '*') } catch {}
          }
          $found = $db.SearchDocuments($params, $true)
          if ($null -ne $found) { foreach ($d in $found) { if ($null -ne $d) { $docs += $d } } }
        }
      } catch { $browseErrors += ('db.CreateSearchParameters: ' + $_.Exception.Message) }
    }
    if ($docs.Count -eq 0 -and $browseErrors.Count -gt 0) {
      throw ('Could not browse iManage documents. Methods tried: ' + ($browseErrors -join ' | '))
    }
    $results = @()
    foreach ($d in $docs) { $results += (Extract-IManDocInfo $d) }
    $json = $results | ConvertTo-Json -Compress -Depth 3
    Write-Output "###JSON_START###$json###JSON_END###"
  } else {
    # WorkObjectFactory / WorkSite path
    $files = New-Object System.Collections.Generic.List[System.Object]
    $browseSuccess = $false
    $browseErrors = @()

    if (-not $browseSuccess) {
      try {
        $wof.GetFiles([ref]$files, ${multiple ? '$true' : '$false'})
        $browseSuccess = $true
      } catch {
        $browseErrors += ('GetFiles(ref,bool): ' + $_.Exception.Message)
      }
    }
    if (-not $browseSuccess) {
      try {
        $wof.GetFiles([ref]$files)
        $browseSuccess = $true
      } catch {
        $browseErrors += ('GetFiles(ref): ' + $_.Exception.Message)
      }
    }
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
  }
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script, { allowInteractive: true, timeoutMs: 120000 });
  const failure = normalizeIManageFailure(result);
  if (failure) {
    if (shouldTryIManageDriveFallback(failure)) {
      return imanageDriveBrowseFiles(multiple);
    }
    return failure;
  }
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
  Assert-IManageOperationSupport $wof 'save'
  $sourcePath = '${resolvedPath.replace(/'/g, "''")}'
  if (-not (Test-Path $sourcePath)) {
    $json = @{ error = "File not found: $sourcePath" } | ConvertTo-Json -Compress
    Write-Output "###JSON_START###$json###JSON_END###"
    exit 0
  }

  if (Use-IManageDMSApi $wof) {
    # IManDMS session-based save
    ${normalizedAction === 'new_version' ? `
    # New version via IManDMS — find document, then use Set-IManDMSDocCheckin helper
    $saveErrors = @()
    $docSaved = $false
    # Try to get document number from file name pattern (e.g. 12345_1.doc)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)
    $docNumMatch = [regex]::Match($baseName, '(\\d{4,})')
    if ($docNumMatch.Success) {
      $docNum = $docNumMatch.Groups[1].Value
      try {
        $doc = Get-IManDMSDocById $wof $docNum $null
        if ($null -ne $doc) {
          try {
            Set-IManDMSDocCheckin $doc $sourcePath 'New version via EmmaNeigh'
            $docSaved = $true
            try { $doc.Refresh() } catch {}
          } catch { $saveErrors += ('Set-IManDMSDocCheckin: ' + $_.Exception.Message) }
          if ($docSaved) {
            $info = Extract-IManDocInfo $doc
            $json = @{ success = $true; action = "new_version"; files = @($info) } | ConvertTo-Json -Compress -Depth 4
            Write-Output "###JSON_START###\$json###JSON_END###"
            exit 0
          }
        }
      } catch { $saveErrors += ('GetDoc: ' + $_.Exception.Message) }
    }
    if (-not $docSaved) {
      # Fall back to import as new document
      $doc = Import-IManDMSDoc $wof $sourcePath
      $info = Extract-IManDocInfo $doc
      $json = @{ success = $true; action = "new_document"; note = "Saved as new document (could not save as new version). Errors: " + ($saveErrors -join " | "); files = @($info) } | ConvertTo-Json -Compress -Depth 4
      Write-Output "###JSON_START###\$json###JSON_END###"
    }
    ` : `
    # New document via IManDMS — import using bootstrap helper
    $doc = Import-IManDMSDoc $wof $sourcePath
    $info = Extract-IManDocInfo $doc
    $json = @{ success = $true; action = "new_document"; files = @($info) } | ConvertTo-Json -Compress -Depth 4
    Write-Output "###JSON_START###\$json###JSON_END###"
    `}
  } else {
    # WorkObjectFactory / WorkSite path
    ${normalizedAction === 'new_version' ? `
    $saveErrors = @()
    $profileId = $null
    try { $profileId = $wof.GetProfileIdFromFilePath($sourcePath) } catch { $saveErrors += ('GetProfileIdFromFilePath: ' + $_.Exception.Message) }
    if ($profileId) {
      $files = New-Object System.Collections.Generic.List[System.Object]
      $entry = New-IManageQueryFile($profileId)
      $files.Add($entry)
      $saved = $false
      $savedAction = 'new_version'
      $savedNote = ''
      if (-not $saved) { try { $wof.SaveAsNewVersion([ref]$files, $sourcePath, ${showDialogPs}); $saved = $true } catch { $saveErrors += ('SaveAsNewVersion(ref,path,dialog): ' + $_.Exception.Message) } }
      if (-not $saved) { try { $wof.SaveAsNewVersion([ref]$files, $sourcePath); $saved = $true } catch { $saveErrors += ('SaveAsNewVersion(ref,path): ' + $_.Exception.Message) } }
      if (-not $saved) { try { $wof.SaveAsNewVersion([ref]$files); $saved = $true } catch { $saveErrors += ('SaveAsNewVersion(ref): ' + $_.Exception.Message) } }
      if (-not $saved) { try { $wof.CheckInFiles([ref]$files, $sourcePath); $saved = $true } catch { $saveErrors += ('CheckInFiles(ref,path): ' + $_.Exception.Message) } }
      if (-not $saved) { try { $wof.SaveFiles([ref]$files, $sourcePath); $saved = $true } catch { $saveErrors += ('SaveFiles(ref,path): ' + $_.Exception.Message) } }
      if (-not $saved) {
        try {
          $savedDoc = $wof.ImportDocument($sourcePath)
          if ($null -ne $savedDoc) {
            $files.Add($savedDoc)
            $saved = $true
            $savedAction = 'new_document'
            $savedNote = 'Saved as new document (new-version methods unavailable for this iManage COM API).'
          }
        } catch { $saveErrors += ('ImportDocument(path): ' + $_.Exception.Message) }
      }
      if (-not $saved) {
        try {
          $savedDoc = $wof.AddDocument($sourcePath)
          if ($null -ne $savedDoc) {
            $files.Add($savedDoc)
            $saved = $true
            $savedAction = 'new_document'
            $savedNote = 'Saved as new document (new-version methods unavailable for this iManage COM API).'
          }
        } catch { $saveErrors += ('AddDocument(path): ' + $_.Exception.Message) }
      }
      if (-not $saved) { throw ('Could not save new version. Methods tried: ' + ($saveErrors -join ' | ')) }
      $savedFiles = @()
      foreach ($f in $files) {
        $info = @{}
        try { $info.name = $f.Name } catch {}
        try { $info.number = $f.Number } catch {}
        try { $info.version = $f.Version } catch {}
        $savedFiles += $info
      }
      $json = @{ success = $true; action = $savedAction; note = $savedNote; profileId = $profileId; files = $savedFiles } | ConvertTo-Json -Compress -Depth 4
      Write-Output "###JSON_START###$json###JSON_END###"
    } else {
      Write-Output ('###JSON_START###{"error":"Could not find iManage profile for this file (' + ($saveErrors -join '; ') + '). Try saving as new document instead."}###JSON_END###')
    }
    ` : `
    $files = New-Object System.Collections.Generic.List[System.Object]
    $saveSuccess = $false
    $saveErrors = @()
    if (-not $saveSuccess) { try { $wof.SaveAsFiles([ref]$files, $sourcePath, ${showDialogPs}); $saveSuccess = $true } catch { $saveErrors += ('SaveAsFiles(ref,path,dialog): ' + $_.Exception.Message) } }
    if (-not $saveSuccess) { try { $wof.SaveAsFiles([ref]$files, $sourcePath); $saveSuccess = $true } catch { $saveErrors += ('SaveAsFiles(ref,path): ' + $_.Exception.Message) } }
    if (-not $saveSuccess) { try { $wof.SaveAsFiles($sourcePath); $saveSuccess = $true } catch { $saveErrors += ('SaveAsFiles(path): ' + $_.Exception.Message) } }
    if (-not $saveSuccess) { try { $savedDoc = $wof.ImportDocument($sourcePath); if ($null -ne $savedDoc) { $files.Add($savedDoc) }; $saveSuccess = $true } catch { $saveErrors += ('ImportDocument(path): ' + $_.Exception.Message) } }
    if (-not $saveSuccess) { try { $savedDoc = $wof.AddDocument($sourcePath); if ($null -ne $savedDoc) { $files.Add($savedDoc) }; $saveSuccess = $true } catch { $saveErrors += ('AddDocument(path): ' + $_.Exception.Message) } }
    if (-not $saveSuccess) { try { $wof.SaveDocument($sourcePath); $saveSuccess = $true } catch { $saveErrors += ('SaveDocument(path): ' + $_.Exception.Message) } }
    if (-not $saveSuccess) { try { $wof.SaveFile($sourcePath); $saveSuccess = $true } catch { $saveErrors += ('SaveFile(path): ' + $_.Exception.Message) } }
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
  }
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script, { allowInteractive: showDialogBool, timeoutMs: 120000 });
  const failure = normalizeIManageFailure(result);
  if (failure) {
    if (shouldTryIManageDriveFallback(failure)) {
      return imanageDriveSaveDocument(resolvedPath, normalizedAction, showDialogBool);
    }
    return failure;
  }
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
  if (isIManageDriveProfileId(normalizedProfileId)) {
    return {
      success: false,
      error: 'iManage Drive fallback does not expose DMS version history. Use COM or iManage REST/OAuth for version lookup.'
    };
  }
  const escapedId = String(normalizedProfileId).replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  Assert-IManageOperationSupport $wof 'versions'

  if (Use-IManageDMSApi $wof) {
    # IManDMS: get document and enumerate Versions collection
    $doc = Get-IManDMSDocById $wof '${escapedId}' $null
    $results = @()
    $versionsErrors = @()

    # Try to get latest version number for reference
    $latestVersionNum = $null
    try { $latestVersionNum = $doc.LatestVersion } catch {}
    if ($null -eq $latestVersionNum) { try { $latestVersionNum = $doc.Version } catch {} }

    # Try Versions collection (primary approach)
    try {
      $versions = $doc.Versions
      if ($null -ne $versions) {
        foreach ($v in $versions) {
          $info = Extract-IManDocInfo $v
          try { $info.date = $v.EditDate } catch { try { $info.date = $v.EditProfileDate } catch { try { $info.date = $v.Date } catch { $info.date = '' } } }
          try { $info.is_latest = ($v.Version -eq $latestVersionNum) } catch { $info.is_latest = $false }
          $results += $info
        }
      }
    } catch { $versionsErrors += ('doc.Versions: ' + $_.Exception.Message) }

    # Fallback: try GetVersions method
    if ($results.Count -eq 0) {
      try {
        $versions = $doc.GetVersions()
        if ($null -ne $versions) {
          foreach ($v in $versions) {
            $info = Extract-IManDocInfo $v
            try { $info.date = $v.EditDate } catch { try { $info.date = $v.Date } catch { $info.date = '' } }
            try { $info.is_latest = ($v.Version -eq $latestVersionNum) } catch { $info.is_latest = $false }
            $results += $info
          }
        }
      } catch { $versionsErrors += ('GetVersions(): ' + $_.Exception.Message) }
    }

    # Fallback: search for all versions by document number
    if ($results.Count -eq 0) {
      try {
        $session = Get-IManDMSSession $wof
        $db = Get-IManDMSDatabase $session
        $params = $wof.CreateProfileSearchParameters()
        $params.Add($script:imProfileDocnum, '${escapedId}')
        $searchResults = $db.SearchDocuments($params, $true)
        if ($null -ne $searchResults) {
          foreach ($v in $searchResults) {
            $info = Extract-IManDocInfo $v
            try { $info.date = $v.EditDate } catch { try { $info.date = $v.Date } catch { $info.date = '' } }
            $results += $info
          }
        }
      } catch { $versionsErrors += ('SearchDocuments fallback: ' + $_.Exception.Message) }
    }

    # Final fallback: just return the document itself as a single version
    if ($results.Count -eq 0) {
      $info = Extract-IManDocInfo $doc
      try { $info.date = $doc.EditDate } catch { $info.date = '' }
      $info.is_latest = $true
      $results += $info
    }

    $output = @{ versions = $results }
    if ($versionsErrors.Count -gt 0) { $output.diagnostics = $versionsErrors }
    if ($null -ne $latestVersionNum) { $output.latest_version = $latestVersionNum }
    $json = $output | ConvertTo-Json -Compress -Depth 4
    Write-Output "###JSON_START###$json###JSON_END###"
  } else {
    # WorkObjectFactory path
    $files = New-Object System.Collections.Generic.List[System.Object]
    $queryFile = New-IManageQueryFile('${escapedId}')
    $files.Add($queryFile)
    try {
      $wof.FindProfiles([ref]$files)
    } catch {}
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
  }
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) return failure;
  const versions = normalizeIManageArrayPayload(result.data, 'versions');
  const output = { success: true, versions };
  if (result.data && result.data.latest_version != null) output.latest_version = result.data.latest_version;
  if (result.data && Array.isArray(result.data.diagnostics) && result.data.diagnostics.length > 0) output.diagnostics = result.data.diagnostics;
  return output;
}

async function imanageCheckout(profileId, checkoutPath, version = null) {
  const normalizedProfileId = normalizeIManageProfileId(profileId);
  if (!normalizedProfileId) {
    return { success: false, error: 'A valid iManage profile ID is required for checkout.' };
  }
  if (isIManageDriveProfileId(normalizedProfileId)) {
    return imanageDriveCheckout(normalizedProfileId, checkoutPath, version);
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
  Assert-IManageOperationSupport $wof 'checkout'

  if (Use-IManageDMSApi $wof) {
    # IManDMS: find document via search, then download with GetCopy
    $doc = Get-IManDMSDocById $wof '${escapedId}' ${versionNumber === null ? '$null' : versionNumber}
    $checkoutDir = '${escapedCheckoutPath}'
    if (-not $checkoutDir) { $checkoutDir = [System.IO.Path]::GetTempPath() }
    $docName = ''
    try { $docName = $doc.Name } catch {}
    $docExt = ''
    try { $docExt = $doc.Extension } catch { try { $docExt = $doc.Type } catch {} }
    if (-not $docName) { $docName = '${escapedId}' }
    if ($docExt -and -not $docName.EndsWith(".$docExt")) { $docName = "$docName.$docExt" }
    $destPath = [System.IO.Path]::Combine($checkoutDir, $docName)
    if (-not (Test-Path $checkoutDir)) { New-Item -ItemType Directory -Path $checkoutDir -Force | Out-Null }
    # Use Get-IManDMSDocCopy (tries GetCopy, CheckOut, Copy in correct order)
    Get-IManDMSDocCopy $doc $destPath
    $info = Extract-IManDocInfo $doc
    $info.path = $destPath
    $json = @{ success = $true; files = @($info); requested_version = ${versionNumber === null ? '""' : versionNumber} } | ConvertTo-Json -Compress -Depth 4
    Write-Output "###JSON_START###$json###JSON_END###"
  } else {
    # WorkObjectFactory path
    $files = New-Object System.Collections.Generic.List[System.Object]
    $queryFile = New-Object PSObject -Property @{
      Number = '${escapedId}'
      ${versionBlock}
    }
    $files.Add($queryFile)
    try {
      $wof.FindProfiles([ref]$files)
    } catch {}
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
  }
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
  const driveRoot = getIManageDriveRootForPath(resolvedPath);
  if (driveRoot) {
    return imanageDriveCheckin(resolvedPath);
  }
  const escapedPath = resolvedPath.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  Assert-IManageOperationSupport $wof 'checkin'
  $sourcePath = '${escapedPath}'

  if (Use-IManageDMSApi $wof) {
    # IManDMS: extract doc number from filename, find document, check in with correct API
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)
    $docNumMatch = [regex]::Match($baseName, '(\d{4,})')
    $ciSuccess = $false
    if ($docNumMatch.Success) {
      $docNum = $docNumMatch.Groups[1].Value
      try {
        $doc = Get-IManDMSDocById $wof $docNum $null
        if ($null -ne $doc) {
          # Use Set-IManDMSDocCheckin (tries CheckInWithResults, CheckInEx, CheckIn, UpdateFromFile)
          Set-IManDMSDocCheckin $doc $sourcePath 'Checked in via EmmaNeigh'
          $ciSuccess = $true
          try { $doc.Refresh() } catch {}
          $info = Extract-IManDocInfo $doc
          $json = @{ success = $true; message = "File checked in successfully"; files = @($info) } | ConvertTo-Json -Compress -Depth 4
          Write-Output "###JSON_START###$json###JSON_END###"
          exit 0
        }
      } catch {}
    }
    if (-not $ciSuccess) {
      # Fallback: import as new document
      $doc = Import-IManDMSDoc $wof $sourcePath
      $info = Extract-IManDocInfo $doc
      $json = @{ success = $true; message = "Imported as new document (could not match existing document for check-in)"; files = @($info) } | ConvertTo-Json -Compress -Depth 4
      Write-Output "###JSON_START###$json###JSON_END###"
    }
  } else {
    # WorkObjectFactory path
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
  }
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) {
    if (shouldTryIManageDriveFallback(failure)) {
      return imanageDriveCheckin(resolvedPath);
    }
    return failure;
  }
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
  Assert-IManageOperationSupport $wof 'search'

  if (Use-IManageDMSApi $wof) {
    # IManDMS session-based search
    $docs = Search-IManDMSDocs $wof '${escapedQuery}' 25
    $results = @()
    foreach ($d in $docs) { $results += (Extract-IManDocInfo $d) }
    $json = $results | ConvertTo-Json -Compress -Depth 3
    Write-Output "###JSON_START###$json###JSON_END###"
  } else {
    # WorkObjectFactory path
    $files = New-Object System.Collections.Generic.List[System.Object]
    $searchFile = New-IManageQueryFile('${escapedQuery}')
    $files.Add($searchFile)
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
  }
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeIManageFailure(result);
  if (failure) {
    if (shouldTryIManageDriveFallback(failure)) {
      return imanageDriveSearch(normalizedQuery, { allowPrompt: true, maxResults: 25 });
    }
    return failure;
  }
  return { success: true, files: normalizeIManageArrayPayload(result.data, 'files') };
}

/**
 * Diagnostic: test iManage COM connectivity.
 * Returns detailed info about which COM variant was found, login status, and available methods.
 */
async function imanageTestConnection() {
  const script = `
$ErrorActionPreference = 'Continue'

function Get-ObjectMembers($obj, $label, $maxMembers) {
  if (-not $maxMembers) { $maxMembers = 80 }
  $info = @{ label = $label; methods = @(); properties = @(); error = '' }
  if ($null -eq $obj) { $info.error = 'null object'; return $info }
  try {
    $info.methods = @($obj.PSObject.Methods | ForEach-Object { $_.Name } | Sort-Object -Unique | Select-Object -First $maxMembers)
  } catch { $info.error = 'methods: ' + $_.Exception.Message }
  try {
    $info.properties = @($obj.PSObject.Properties | ForEach-Object { $_.Name } | Sort-Object -Unique | Select-Object -First $maxMembers)
  } catch { if (-not $info.error) { $info.error = 'properties: ' + $_.Exception.Message } }
  return $info
}

$diag = @{
  platform = $env:OS
  powershell_version = $PSVersionTable.PSVersion.ToString()
  architecture = 'unknown'
  com_created = $false
  com_progid = ''
  com_type = ''
  com_api_type = ''
  com_candidates = @()
  selected_candidate = @{}
  limited_api_surface = $false
  login_status = 'unknown'
  dms_members = @{}
  session_members = @{}
  database_members = @{}
  document_members = @{}
  sessions_info = @{ count = 0; error = '' }
  databases_info = @{ count = 0; names = @(); error = '' }
  error = ''
}

try {
  $diag.architecture = if ([System.IntPtr]::Size -eq 8) { '64-bit' } else { '32-bit' }
} catch {}

# ── Step 1: Create COM object ──
$wof = $null
try {
  $wof = Get-IManageWorkObjectFactory
  if ($null -ne $wof) {
    $diag.com_created = $true
    $diag.com_progid = $global:IManageProgId
    try { $diag.com_type = $wof.GetType().FullName } catch { $diag.com_type = 'unknown' }
    $diag.com_api_type = $global:IManageCOMType
    try { $diag.selected_candidate = $global:IManageSelectionDetails } catch {}
    try { $diag.limited_api_surface = [bool]$global:IManageSelectionDetails.limited_surface } catch {}
  }
} catch {
  $diag.error = $_.Exception.Message
}
try {
  if ($global:IManageCandidateDiagnostics) {
    $diag.com_candidates = $global:IManageCandidateDiagnostics
  }
} catch {}
if (-not $diag.com_created -and -not $diag.error) {
  $diag.error = 'No iManage COM object created.'
}
if (-not $diag.com_created) {
  $json = $diag | ConvertTo-Json -Compress -Depth 4
  Write-Output "###JSON_START###$json###JSON_END###"
  exit 0
}

# ── Step 2: Enumerate DMS object methods/properties ──
$diag.dms_members = Get-ObjectMembers $wof 'DMS'

# ── Step 3: Check login / sessions ──
try {
  $hasLogin = $false
  try { if ($wof.PSObject.Properties['HasLogin']) { $hasLogin = [bool]$wof.HasLogin } }
  catch {}
  if (-not $hasLogin) {
    try { if ($wof.PSObject.Properties['Connected']) { $hasLogin = [bool]$wof.Connected } }
    catch {}
  }
  if (-not $hasLogin) {
    try {
      $sess = $wof.Sessions
      if ($null -ne $sess -and $sess.Count -gt 0) { $hasLogin = $true }
    } catch {}
  }
  $diag.login_status = if ($hasLogin) { 'logged_in' } else { 'not_logged_in' }
} catch { $diag.login_status = 'error: ' + $_.Exception.Message }

# ── Step 4: Enumerate Sessions ──
$session = $null
try {
  $sessions = $wof.Sessions
  if ($null -ne $sessions) {
    try { $diag.sessions_info.count = $sessions.Count } catch {}
    # Get first session
    try { foreach ($s in $sessions) { if ($null -ne $s) { $session = $s; break } } }
    catch {
      try {
        if ($sessions.Count -gt 0) { $session = $sessions.Item(1) }
      } catch {}
    }
  }
} catch { $diag.sessions_info.error = $_.Exception.Message }

if ($null -ne $session) {
  $diag.session_members = Get-ObjectMembers $session 'Session'
}

# ── Step 5: Enumerate Databases ──
$db = $null
if ($null -ne $session) {
  try {
    $dbs = $session.Databases
    if ($null -ne $dbs) {
      try { $diag.databases_info.count = $dbs.Count } catch {}
      $dbNames = @()
      try {
        foreach ($d in $dbs) {
          if ($null -ne $d) {
            if ($null -eq $db) { $db = $d }
            try { $dbNames += $d.Name } catch {}
          }
        }
      } catch {
        try {
          $cnt = $dbs.Count
          for ($i = 1; $i -le $cnt; $i++) {
            try {
              $d = $dbs.Item($i)
              if ($null -ne $d) {
                if ($null -eq $db) { $db = $d }
                try { $dbNames += $d.Name } catch {}
              }
            } catch {}
          }
        } catch {}
      }
      $diag.databases_info.names = $dbNames
    }
  } catch { $diag.databases_info.error = $_.Exception.Message }
}

if ($null -ne $db) {
  $diag.database_members = Get-ObjectMembers $db 'Database'
}

# ── Step 6: Try to find a Document to enumerate ──
if ($null -ne $db -and $null -ne $wof) {
  $sampleDoc = $null
  # Try CreateProfileSearchParameters on DMS
  try {
    $params = $wof.CreateProfileSearchParameters()
    if ($null -ne $params) {
      $diag.has_CreateProfileSearchParameters = $true
      try { $params.Add(3, '*') } catch { try { $params.Add('NAME', '*') } catch {} }
      try {
        $found = $db.SearchDocuments($params, $true)
        if ($null -ne $found) {
          foreach ($doc in $found) { if ($null -ne $doc) { $sampleDoc = $doc; break } }
        }
      } catch {}
    }
  } catch { $diag.has_CreateProfileSearchParameters = $false }
  # Fallback: try RecentDocuments
  if ($null -eq $sampleDoc) {
    try {
      $recent = $db.RecentDocuments
      if ($null -ne $recent) {
        foreach ($doc in $recent) { if ($null -ne $doc) { $sampleDoc = $doc; break } }
      }
    } catch {}
  }
  if ($null -ne $sampleDoc) {
    $diag.document_members = Get-ObjectMembers $sampleDoc 'Document'
    try { $diag.sample_doc_name = $sampleDoc.Name } catch {}
    try { $diag.sample_doc_number = $sampleDoc.Number } catch {}
  }
}

# ── Step 7: Check for key methods ──
$diag.key_checks = @{
  dms_capability_detected = (Use-IManageDMSApi $wof)
  dms_CreateProfileSearchParameters = ($diag.dms_members.methods -contains 'CreateProfileSearchParameters')
  db_SearchDocuments = ($diag.database_members.methods -contains 'SearchDocuments')
  db_CreateDocument = ($diag.database_members.methods -contains 'CreateDocument')
  db_GetDocument = ($diag.database_members.methods -contains 'GetDocument')
  doc_GetCopy = ($diag.document_members.methods -contains 'GetCopy')
  doc_CheckInWithResults = ($diag.document_members.methods -contains 'CheckInWithResults')
  doc_CheckInEx = ($diag.document_members.methods -contains 'CheckInEx')
  doc_CheckOut = ($diag.document_members.methods -contains 'CheckOut')
  doc_Versions = ($diag.document_members.properties -contains 'Versions')
  doc_LatestVersion = ($diag.document_members.properties -contains 'LatestVersion')
}

$json = $diag | ConvertTo-Json -Compress -Depth 5
Write-Output "###JSON_START###$json###JSON_END###"
  `;

  const result = await imanageRunPowerShell(script, { timeoutMs: 45000 });
  const driveRoots = detectIManageDriveRoots();
  if (!result || !result.success) {
    return {
      success: false,
      error: result ? (result.error || 'Unknown error') : 'PowerShell execution failed',
      diagnostics: {
        platform: process.platform,
        powershellHosts: getIManagePowerShellHosts(),
        stderr: result ? result.stderr : '',
        drive_roots: driveRoots,
        drive_fallback_available: driveRoots.length > 0
      }
    };
  }
  const data = result.data || {};
  const checks = data.key_checks || {};
  const checkSummary = Object.entries(checks)
    .map(([k, v]) => `${k}: ${v ? 'YES' : 'no'}`)
    .join(', ');
  return {
    success: !!data.com_created,
    diagnostics: {
      ...data,
      drive_roots: driveRoots,
      drive_fallback_available: driveRoots.length > 0
    },
    message: data.com_created
      ? `iManage COM connected via ${data.com_progid} [API: ${data.com_api_type || 'unknown'}] (${data.login_status}). Sessions: ${data.sessions_info?.count || 0}. Databases: ${data.databases_info?.count || 0} (${(data.databases_info?.names || []).join(', ')}). Key methods: ${checkSummary}`
      : `iManage COM could not be loaded. ${data.error || ''}`
  };
}

/**
 * Create a folder in iManage.
 */
async function imanageCreateFolder(folderName, parentPath = '') {
  const normalizedName = String(folderName || '').trim();
  if (!normalizedName) {
    return { success: false, error: 'A folder name is required.' };
  }
  const escapedName = normalizedName.replace(/'/g, "''");
  const escapedParent = String(parentPath || '').trim().replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
try {
  $wof = Get-IManageWorkObjectFactory
  Ensure-IManageLogin $wof
  Assert-IManageOperationSupport $wof 'create_folder'

  if (Use-IManageDMSApi $wof) {
    $folder = Create-IManDMSFolder $wof '${escapedName}' '${escapedParent}'
    $info = @{}
    try { $info.name = $folder.Name } catch { $info.name = '${escapedName}' }
    try { $info.id = $folder.FolderNumber } catch {
      try { $info.id = $folder.FolderId } catch { $info.id = '' }
    }
    try { $info.path = $folder.Path } catch { $info.path = '' }
    $json = @{ success = $true; folder = $info } | ConvertTo-Json -Compress -Depth 3
    Write-Output "###JSON_START###$json###JSON_END###"
  } else {
    # WorkObjectFactory — try CreateFolder / MakeFolder methods
    $createErrors = @()
    $created = $false
    if (-not $created) { try { $wof.CreateFolder('${escapedName}'); $created = $true } catch { $createErrors += ('CreateFolder: ' + $_.Exception.Message) } }
    if (-not $created) { try { $wof.MakeFolder('${escapedName}'); $created = $true } catch { $createErrors += ('MakeFolder: ' + $_.Exception.Message) } }
    if (-not $created) { try { $wof.NewFolder('${escapedName}'); $created = $true } catch { $createErrors += ('NewFolder: ' + $_.Exception.Message) } }
    if (-not $created) { try { $wof.AddFolder('${escapedName}'); $created = $true } catch { $createErrors += ('AddFolder: ' + $_.Exception.Message) } }
    if (-not $created) { try { $wof.CreateWorkspace('${escapedName}'); $created = $true } catch { $createErrors += ('CreateWorkspace(name): ' + $_.Exception.Message) } }
    if (-not $created -and '${escapedParent}') {
      try {
        $files = New-Object System.Collections.Generic.List[System.Object]
        $queryFile = New-IManageQueryFile('${escapedParent}')
        $files.Add($queryFile)
        $wof.FindProfiles([ref]$files)
        if ($files.Count -gt 0) {
          try { $wof.CreateFolder('${escapedName}', $files[0]); $created = $true } catch { $createErrors += ('CreateFolder(name,parent): ' + $_.Exception.Message) }
          if (-not $created) { try { $wof.MakeFolder('${escapedName}', $files[0]); $created = $true } catch { $createErrors += ('MakeFolder(name,parent): ' + $_.Exception.Message) } }
          if (-not $created) { try { $wof.CreateWorkspace('${escapedName}', $files[0]); $created = $true } catch { $createErrors += ('CreateWorkspace(name,parent): ' + $_.Exception.Message) } }
        }
      } catch { $createErrors += ('FindParent: ' + $_.Exception.Message) }
    }
    if (-not $created) { throw ('Could not create folder. Methods tried: ' + ($createErrors -join ' | ')) }
    $json = @{ success = $true; folder = @{ name = '${escapedName}' } } | ConvertTo-Json -Compress -Depth 3
    Write-Output "###JSON_START###$json###JSON_END###"
  }
} catch {
  $json = @{ error = $_.Exception.Message } | ConvertTo-Json -Compress
  Write-Output "###JSON_START###$json###JSON_END###"
}
  `;
  const result = await imanageRunPowerShell(script, { timeoutMs: 60000 });
  const failure = normalizeIManageFailure(result);
  if (failure) {
    if (shouldTryIManageDriveFallback(failure)) {
      return imanageDriveCreateFolder(normalizedName, parentPath);
    }
    return failure;
  }
  const data = result.data && typeof result.data === 'object' ? result.data : {};
  return { success: true, folder: data.folder || { name: normalizedName } };
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

// ========== iMANAGE VERSION REDLINE (LITERA + EMAIL) ==========

/**
 * Checkout two versions of an iManage document, run a Litera redline, and
 * optionally email the result via Outlook COM.
 *
 * @param {string} profileId     iManage document number / profile ID
 * @param {number|string} v1     Version number for the "original" (older) side
 * @param {number|string} v2     Version number for the "modified" (newer) side
 * @param {object}  opts
 * @param {string}  [opts.email_to]       Recipients for the redline (semicolon-separated)
 * @param {string}  [opts.email_cc]       CC recipients
 * @param {string}  [opts.email_subject]  Custom subject (default auto-generated)
 * @param {string}  [opts.email_body]     Custom body (default auto-generated)
 * @param {string}  [opts.output_format]  'pdf' (default) or 'docx'
 * @param {string}  [opts.output_folder]  Custom output folder (default: Downloads)
 */
async function imanageRedlineVersions(profileId, v1, v2, opts = {}) {
  if (process.platform !== 'win32') {
    return { success: false, error: 'iManage + Litera redlining requires Windows.' };
  }

  const normalizedProfileId = normalizeIManageProfileId(profileId);
  if (!normalizedProfileId) {
    return { success: false, error: 'A valid iManage profile ID (document number) is required.' };
  }

  const version1 = parseIManageVersionNumber(v1);
  const version2 = parseIManageVersionNumber(v2);

  // If versions are not explicitly provided, auto-detect V1 and latest
  let autoDetected = false;
  let resolvedV1 = version1;
  let resolvedV2 = version2;

  if (resolvedV1 === null || resolvedV2 === null) {
    const versionsResult = await imanageGetVersions(normalizedProfileId);
    if (!versionsResult.success) {
      return {
        success: false,
        error: versionsResult.error || `Failed to retrieve version history for document ${normalizedProfileId}.`
      };
    }
    const versionList = versionsResult.versions || [];
    if (versionList.length < 2) {
      return {
        success: false,
        error: `Document ${normalizedProfileId} only has ${versionList.length} version(s). Need at least 2 versions to redline.`
      };
    }
    const selected = pickIManagePrecedentAndCurrentVersions(versionList);
    if (!selected || !selected.precedent || !selected.latest) {
      return { success: false, error: 'Could not determine which versions to compare.' };
    }
    if (resolvedV1 === null) resolvedV1 = selected.precedent.version_number;
    if (resolvedV2 === null) resolvedV2 = selected.latest.version_number;
    autoDetected = true;
  }

  if (resolvedV1 === resolvedV2) {
    return { success: false, error: `Both versions are the same (V${resolvedV1}). Provide two different version numbers.` };
  }

  // Ensure V1 < V2 so the diff direction is correct (older → newer)
  const origVersion = Math.min(resolvedV1, resolvedV2);
  const modVersion = Math.max(resolvedV1, resolvedV2);

  // Create temp checkout folders
  const timestamp = Date.now();
  const checkoutRoot = path.join(app.getPath('temp'), `emmaneigh_version_redline_${timestamp}`);
  const v1Folder = path.join(checkoutRoot, `v${origVersion}`);
  const v2Folder = path.join(checkoutRoot, `v${modVersion}`);
  fs.mkdirSync(v1Folder, { recursive: true });
  fs.mkdirSync(v2Folder, { recursive: true });

  // Step 1: Checkout version 1 (older)
  const v1Checkout = await imanageCheckout(normalizedProfileId, v1Folder, origVersion);
  if (!v1Checkout.success) {
    return {
      success: false,
      error: `Failed to check out V${origVersion}: ${v1Checkout.error || 'Unknown error'}`
    };
  }

  // Step 2: Checkout version 2 (newer)
  const v2Checkout = await imanageCheckout(normalizedProfileId, v2Folder, modVersion);
  if (!v2Checkout.success) {
    return {
      success: false,
      error: `Failed to check out V${modVersion}: ${v2Checkout.error || 'Unknown error'}`
    };
  }

  // Resolve actual file paths from checkout results
  const docNameHint = (v1Checkout.files && v1Checkout.files[0] && v1Checkout.files[0].name) || normalizedProfileId;
  const v1Path = resolveIManageCheckoutFilePath(v1Checkout, v1Folder, docNameHint);
  const v2Path = resolveIManageCheckoutFilePath(v2Checkout, v2Folder, docNameHint);

  if (!v1Path || !fs.existsSync(v1Path)) {
    return { success: false, error: `Could not locate the checked-out V${origVersion} file on disk.` };
  }
  if (!v2Path || !fs.existsSync(v2Path)) {
    return { success: false, error: `Could not locate the checked-out V${modVersion} file on disk.` };
  }

  // Step 3: Check Litera is installed
  const literaPath = findLiteraInstallation();
  if (!literaPath) {
    return {
      success: false,
      error: 'Litera Compare is not installed. Install Litera Compare to enable version redlines.',
      v1_path: v1Path,
      v2_path: v2Path
    };
  }

  // Determine output path
  const outputFormat = String(opts.output_format || 'pdf').toLowerCase();
  const cpo = !!opts.change_pages_only;
  const literaOptions = {
    output_format: outputFormat === 'docx' ? 'docx' : 'pdf',
    change_pages_only: cpo
  };
  const outputExt = getLiteraOutputExtension(v1Path, literaOptions);
  const defaultOutputFolder = normalizeLocalPath(opts.output_folder || '')
    || path.join(app.getPath('downloads'), 'EmmaNeigh_Redlines');
  fs.mkdirSync(defaultOutputFolder, { recursive: true });
  const outputFileName = buildRedlineOutputFilename(v1Path, v2Path, outputExt);
  const outputPath = path.join(defaultOutputFolder, outputFileName);

  // Step 4: Run Litera comparison
  let redlineResult;
  try {
    redlineResult = await runLiteraComparison(v1Path, v2Path, outputPath, literaOptions);
  } catch (err) {
    return {
      success: false,
      error: `Litera comparison failed: ${err.message}`,
      v1_path: v1Path,
      v2_path: v2Path
    };
  }

  if (!redlineResult || !redlineResult.success) {
    return {
      success: false,
      error: (redlineResult && redlineResult.error) || 'Litera comparison did not produce output.',
      v1_path: v1Path,
      v2_path: v2Path
    };
  }

  const redlinePath = redlineResult.output_path || outputPath;
  const resultPayload = {
    success: true,
    profile_id: normalizedProfileId,
    version_original: origVersion,
    version_modified: modVersion,
    auto_detected_versions: autoDetected,
    v1_path: v1Path,
    v2_path: v2Path,
    redline_path: redlinePath,
    output_format: literaOptions.output_format,
    change_pages_only: cpo
  };

  // Step 5: Optionally email the redline
  const emailTo = String(opts.email_to || '').trim();
  if (emailTo) {
    const docDisplayName = docNameHint || normalizedProfileId;
    const subject = opts.email_subject
      || `Redline: ${docDisplayName} V${origVersion} → V${modVersion}`;
    const body = opts.email_body
      || `<p>Please find attached the redline comparison of <b>${docDisplayName}</b>, ` +
         `comparing Version ${origVersion} to Version ${modVersion}.</p>` +
         `<p>Generated by EmmaNeigh (Litera Compare).</p>`;
    const cc = opts.email_cc || '';

    if (!fs.existsSync(redlinePath)) {
      resultPayload.email_sent = false;
      resultPayload.email_error = 'Redline output file not found for attachment.';
    } else {
      const emailResult = await outlookSendEmail(emailTo, subject, body, cc, [redlinePath]);
      if (emailResult.success) {
        resultPayload.email_sent = true;
        resultPayload.email_to = emailTo;
        resultPayload.message = `Redline V${origVersion}→V${modVersion} created and emailed to ${emailTo}.`;
      } else {
        resultPayload.email_sent = false;
        resultPayload.email_error = emailResult.error || 'Failed to send email via Outlook.';
        resultPayload.message = `Redline created at ${redlinePath} but email failed: ${resultPayload.email_error}`;
      }
    }
  } else {
    resultPayload.message = `Redline V${origVersion}→V${modVersion} created: ${redlinePath}`;
  }

  return resultPayload;
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

/**
 * Reply to an existing Outlook email thread (or reply-all).
 *
 * @param {string} entryId     The Outlook EntryID of the email to reply to
 * @param {string} body        Reply body (HTML supported)
 * @param {boolean} replyAll   If true, use Reply All instead of Reply
 * @param {string[]} attachments  Optional file paths to attach
 */
async function outlookReplyEmail(entryId, body, replyAll = false, attachments = []) {
  if (process.platform !== 'win32') return { success: false, error: 'Outlook COM is only available on Windows.' };
  const normalizedEntryId = String(entryId || '').trim();
  if (!normalizedEntryId) return { success: false, error: 'entry_id is required to reply to an Outlook email.' };
  const normalizedBody = String(body || '');
  if (!normalizedBody.trim()) return { success: false, error: 'body is required for the reply.' };
  const escapedId = normalizedEntryId.replace(/'/g, "''");
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
  const attachLines = attachPaths.map(p => `$reply.Attachments.Add('${p.replace(/'/g, "''")}') | Out-Null`).join('\n  ');
  const replyMethod = replyAll ? 'ReplyAll' : 'Reply';
  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$ns = $null
$item = $null
$reply = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $ns = $outlook.GetNamespace("MAPI")
  try { $null = $ns.Logon("", "", $false, $false) } catch {}
  $item = $ns.GetItemFromID('${escapedId}')
  if ($null -eq $item) { throw "Outlook item not found for EntryID." }
  if ($item.Class -ne 43) { throw "Outlook item is not an email message." }
  $reply = $item.${replyMethod}()
  $bodyBytes = [Convert]::FromBase64String('${bodyB64}')
  $newBody = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
  $reply.HTMLBody = $newBody + $reply.HTMLBody
  ${attachLines}
  $reply.Send()
  $result = @{
    success = $true
    replied_to_subject = $item.Subject
    reply_type = '${replyMethod.toLowerCase()}'
  }
  $json = $result | ConvertTo-Json -Compress -Depth 3
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($reply -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($reply) | Out-Null } } catch {}
  try { if ($item -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) | Out-Null } } catch {}
  try { if ($ns -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ns) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;
  const result = await imanageRunPowerShell(script);
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;
  const data = result.data && typeof result.data === 'object' ? result.data : {};
  return {
    success: true,
    replied_to_subject: data.replied_to_subject || '',
    reply_type: data.reply_type || replyMethod.toLowerCase(),
    message: `${replyAll ? 'Reply-all' : 'Reply'} sent to thread: "${data.replied_to_subject || 'email'}"`
  };
}

function splitEmailAttachmentNames(rawValue) {
  if (Array.isArray(rawValue)) {
    return rawValue
      .map((value) => String(value || '').trim())
      .filter(Boolean);
  }
  return String(rawValue || '')
    .split(/[;,]/)
    .map((value) => value.trim())
    .filter(Boolean);
}

function normalizeLoadedEmail(record = {}) {
  const attachments = splitEmailAttachmentNames(record.attachments);
  return {
    subject: String(record.subject || '').trim(),
    from: String(record.from || record.sender || record.sender_name || record.sender_email || '').trim(),
    to: String(record.to || '').trim(),
    cc: String(record.cc || '').trim(),
    date_sent: String(record.date_sent || record.sent || '').trim(),
    date_received: String(record.date_received || record.received || '').trim(),
    body: String(record.body || '').trim(),
    attachments: attachments.join('; '),
    has_attachments: Boolean(record.has_attachments) || attachments.length > 0,
    folder: String(record.folder || '').trim(),
    entry_id: String(record.entry_id || '').trim()
  };
}

function buildLoadedEmailSummary(emails = []) {
  const normalized = Array.isArray(emails) ? emails : [];
  const uniqueSenders = new Set(
    normalized
      .map((email) => String(email.from || '').trim().toLowerCase())
      .filter(Boolean)
  );
  const dateValues = normalized
    .map((email) => String(email.date_received || email.date_sent || '').trim())
    .filter(Boolean)
    .map((value) => {
      const parsed = Date.parse(value);
      return Number.isFinite(parsed) ? { raw: value, ts: parsed } : null;
    })
    .filter(Boolean)
    .sort((a, b) => a.ts - b.ts);

  return {
    total_emails: normalized.length,
    unique_senders: uniqueSenders.size,
    with_attachments: normalized.filter((email) => email.has_attachments).length,
    date_range: {
      earliest: dateValues.length ? dateValues[0].raw : null,
      latest: dateValues.length ? dateValues[dateValues.length - 1].raw : null
    }
  };
}

function writeTempEmailDataset(emails = []) {
  const outputPath = path.join(
    app.getPath('temp'),
    `emmaneigh_email_activity_${Date.now()}_${Math.floor(Math.random() * 100000)}.json`
  );
  fs.writeFileSync(outputPath, JSON.stringify({ emails }, null, 2), 'utf8');
  return outputPath;
}

async function listOutlookFolders(options = {}) {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Outlook COM is only available on Windows.' };
  }

  const roots = Array.isArray(options.roots) && options.roots.length
    ? options.roots.map((value) => String(value || '').trim()).filter(Boolean)
    : ['Inbox', 'Sent Items'];
  const includeSubfolders = options.includeSubfolders !== false;
  const escapedRoots = roots.map((value) => `'${value.replace(/'/g, "''")}'`).join(', ');

  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$ns = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $ns = $outlook.GetNamespace("MAPI")
  try { $null = $ns.Logon("", "", $false, $false) } catch {}

  $rootNames = @(${escapedRoots})
  $folders = New-Object System.Collections.Generic.List[Object]

  function Add-FolderTargets($folderObj, $rootName, $targetList, [bool]$includeChildren) {
    if ($null -eq $folderObj) { return }
    $folderPath = ''
    try { $folderPath = $folderObj.FolderPath } catch {}
    if (-not $folderPath) { $folderPath = $rootName }
    $folderDisplayName = $folderPath
    try {
      if ($folderObj.Name) { $folderDisplayName = $folderObj.Name }
    } catch {}
    $targetList.Add(@{
      path = $folderPath
      root = $rootName
      name = $folderDisplayName
    }) | Out-Null
    if (-not $includeChildren) { return }
    try {
      foreach ($childFolder in $folderObj.Folders) {
        Add-FolderTargets $childFolder $rootName $targetList $includeChildren
      }
    } catch {}
  }

  foreach ($rootName in $rootNames) {
    $folderObj = $null
    switch ($rootName) {
      'Inbox' { $folderObj = $ns.GetDefaultFolder(6) }
      'Sent Items' { $folderObj = $ns.GetDefaultFolder(5) }
      'Drafts' { $folderObj = $ns.GetDefaultFolder(16) }
      'Deleted Items' { $folderObj = $ns.GetDefaultFolder(3) }
      default { $folderObj = $ns.GetDefaultFolder(6) }
    }
    if ($null -eq $folderObj) { continue }
    Add-FolderTargets $folderObj $rootName $folders ${includeSubfolders ? '$true' : '$false'}
  }

  $json = $folders | Sort-Object root, path | ConvertTo-Json -Compress -Depth 4
  if (-not $json) { $json = '[]' }
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($ns -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ns) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;

  const result = await imanageRunPowerShell(script, { timeoutMs: 60000 });
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;
  const folders = Array.isArray(result.data) ? result.data : (result.data ? [result.data] : []);
  return { success: true, folders };
}

async function collectOutlookEmails(options = {}) {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Outlook COM is only available on Windows.' };
  }

  const folders = Array.isArray(options.folders) && options.folders.length
    ? options.folders.map((value) => String(value || '').trim()).filter(Boolean)
    : ['Inbox', 'Sent Items'];
  const folderPaths = Array.isArray(options.folderPaths)
    ? options.folderPaths.map((value) => String(value || '').trim()).filter(Boolean)
    : [];
  const includeSubfolders = options.includeSubfolders !== false;
  const daysBack = Math.max(1, Math.min(Math.round(Number(options.daysBack || 30)), 3650));
  const maxResultsPerFolder = Math.max(25, Math.min(Math.round(Number(options.maxResultsPerFolder || 250)), 1000));
  const maxTotalResults = Math.max(50, Math.min(Math.round(Number(options.maxTotalResults || 600)), 5000));
  const bodyCharLimit = Math.max(400, Math.min(Math.round(Number(options.bodyCharLimit || 4000)), 20000));
  const escapedFolders = folders.map((value) => `'${value.replace(/'/g, "''")}'`).join(', ');
  const escapedFolderPaths = folderPaths.map((value) => `'${value.replace(/'/g, "''")}'`).join(', ');

  const script = `
$ErrorActionPreference = 'Stop'
$outlook = $null
$ns = $null
try {
  $outlook = New-Object -ComObject "Outlook.Application"
  $ns = $outlook.GetNamespace("MAPI")
  try { $null = $ns.Logon("", "", $false, $false) } catch {}

  $targetFolders = @(${escapedFolders})
  $selectedFolderPaths = @(${escapedFolderPaths})
  $cutoffDate = (Get-Date).AddDays(-${daysBack})
  $emails = New-Object System.Collections.Generic.List[Object]
  $folderTargets = New-Object System.Collections.Generic.List[Object]

  function Add-FolderTargets($folderObj, $rootName, $targetList, [bool]$includeChildren) {
    if ($null -eq $folderObj) { return }
    $folderPath = ''
    try { $folderPath = $folderObj.FolderPath } catch {}
    if (-not $folderPath) { $folderPath = $rootName }
    $targetList.Add(@{
      Name = $folderPath
      Root = $rootName
      Folder = $folderObj
    }) | Out-Null
    if (-not $includeChildren) { return }
    try {
      foreach ($childFolder in $folderObj.Folders) {
        Add-FolderTargets $childFolder $rootName $targetList $includeChildren
      }
    } catch {}
  }

  foreach ($folderName in $targetFolders) {
    $folderObj = $null
    switch ($folderName) {
      'Inbox' { $folderObj = $ns.GetDefaultFolder(6) }
      'Sent Items' { $folderObj = $ns.GetDefaultFolder(5) }
      'Drafts' { $folderObj = $ns.GetDefaultFolder(16) }
      'Deleted Items' { $folderObj = $ns.GetDefaultFolder(3) }
      default { $folderObj = $ns.GetDefaultFolder(6) }
    }
    if ($null -eq $folderObj) { continue }
    Add-FolderTargets $folderObj $folderName $folderTargets ${includeSubfolders ? '$true' : '$false'}
  }

  $totalCount = 0
  foreach ($folderInfo in $folderTargets) {
    if ($totalCount -ge ${maxTotalResults}) { break }
    if ($selectedFolderPaths.Count -gt 0 -and -not ($selectedFolderPaths -contains $folderInfo.Name)) { continue }
    $folderObj = $folderInfo.Folder
    $folderName = if ($folderInfo.Name) { $folderInfo.Name } else { $folderInfo.Root }
    $rootFolderName = $folderInfo.Root
    $items = $folderObj.Items
    try {
      if ($rootFolderName -eq 'Sent Items') {
        $items.Sort("[SentOn]", $true)
      } else {
        $items.Sort("[ReceivedTime]", $true)
      }
    } catch {}

    $count = 0
    foreach ($item in $items) {
      if ($count -ge ${maxResultsPerFolder} -or $totalCount -ge ${maxTotalResults}) { break }
      if ($null -eq $item) { continue }
      if ($item.Class -ne 43) { continue }

      $sentOn = $null
      $receivedTime = $null
      try { $sentOn = $item.SentOn } catch {}
      try { $receivedTime = $item.ReceivedTime } catch {}
      $msgDate = if ($rootFolderName -eq 'Sent Items') { $sentOn } else { $receivedTime }
      if ($null -eq $msgDate) { $msgDate = $sentOn }
      if ($null -eq $msgDate) { $msgDate = $receivedTime }
      if ($null -ne $msgDate -and $msgDate -lt $cutoffDate) { continue }

      $bodyText = ''
      try {
        if ($item.Body) {
          $bodyText = $item.Body.Substring(0, [Math]::Min(${bodyCharLimit}, $item.Body.Length))
        }
      } catch {}
      $bodyEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bodyText))

      $attachmentNames = @()
      try {
        foreach ($att in $item.Attachments) {
          if ($att.FileName) { $attachmentNames += $att.FileName }
        }
      } catch {}

      $senderName = ''
      try { if ($item.SenderName) { $senderName = $item.SenderName } } catch {}
      $senderEmail = ''
      try { if ($item.SenderEmailAddress) { $senderEmail = $item.SenderEmailAddress } } catch {}
      if (-not $senderName -and $rootFolderName -eq 'Sent Items') { $senderName = 'Me' }

      $emails.Add(@{
        entry_id = $item.EntryID
        folder = $folderName
        root_folder = $folderInfo.Root
        subject = if ($item.Subject) { $item.Subject } else { '' }
        from = $senderName
        sender_email = $senderEmail
        to = if ($item.To) { $item.To } else { '' }
        cc = if ($item.CC) { $item.CC } else { '' }
        date_sent = if ($sentOn) { $sentOn.ToString("o") } else { '' }
        date_received = if ($receivedTime) { $receivedTime.ToString("o") } else { '' }
        body_b64 = $bodyEncoded
        attachments = $attachmentNames
        has_attachments = ($attachmentNames.Count -gt 0)
      }) | Out-Null
      $count += 1
      $totalCount += 1
    }
  }

  $json = $emails | ConvertTo-Json -Compress -Depth 6
  if (-not $json) { $json = '[]' }
  Write-Output "###JSON_START###$json###JSON_END###"
} catch {
  $errMsg = $_.Exception.Message -replace '"', '\\"'
  Write-Output "###JSON_START###{\\"error\\":\\"$errMsg\\"}###JSON_END###"
} finally {
  try { if ($ns -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ns) | Out-Null } } catch {}
  try { if ($outlook -ne $null) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) | Out-Null } } catch {}
}`;

  const result = await imanageRunPowerShell(script, { timeoutMs: Math.max(60000, daysBack > 90 ? 120000 : 90000) });
  const failure = normalizeOfficeComFailure(result, 'Outlook');
  if (failure) return failure;

  const rawEmails = Array.isArray(result.data) ? result.data : (result.data ? [result.data] : []);
  const seen = new Set();
  const emails = rawEmails
    .map((record) => {
      const normalized = { ...record };
      if (normalized.body_b64) {
        try {
          normalized.body = Buffer.from(String(normalized.body_b64), 'base64').toString('utf8');
        } catch (_) {
          normalized.body = '';
        }
        delete normalized.body_b64;
      }
      return normalizeLoadedEmail(normalized);
    })
    .filter((email) => {
      const dedupeKey = `${email.entry_id || ''}|${email.folder || ''}`;
      if (dedupeKey !== '|' && seen.has(dedupeKey)) return false;
      if (dedupeKey !== '|') seen.add(dedupeKey);
      return true;
    })
    .sort((a, b) => {
      const aTs = Date.parse(a.date_received || a.date_sent || '') || 0;
      const bTs = Date.parse(b.date_received || b.date_sent || '') || 0;
      return bTs - aTs;
    });

  return {
    success: true,
    emails,
    summary: buildLoadedEmailSummary(emails),
    daysBack,
    folders,
    folderPaths
  };
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

    case 'imanage_create_folder':
      if (!safeInput.folder_name || !String(safeInput.folder_name).trim()) {
        result = { success: false, error: 'folder_name is required.' };
        break;
      }
      result = await imanageCreateFolder(safeInput.folder_name, safeInput.parent_path || '');
      break;

    case 'imanage_redline_versions': {
      if (!safeInput.profile_id) {
        result = { success: false, error: 'profile_id is required for iManage version redline.' };
        break;
      }
      result = await imanageRedlineVersions(
        safeInput.profile_id,
        safeInput.version_1 || null,
        safeInput.version_2 || null,
        {
          email_to: safeInput.email_to,
          email_cc: safeInput.email_cc,
          email_subject: safeInput.email_subject,
          email_body: safeInput.email_body,
          output_format: safeInput.output_format,
          output_folder: safeInput.output_folder,
          change_pages_only: !!safeInput.change_pages_only
        }
      );
      break;
    }

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
        literaOptions: {
          change_pages_only: !!safeInput.change_pages_only
        }
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

    case 'outlook_reply_email': {
      if (!safeInput.entry_id) {
        result = { success: false, error: 'entry_id is required to reply to an Outlook email.' };
        break;
      }
      if (!safeInput.body) {
        result = { success: false, error: 'body is required for the reply.' };
        break;
      }
      let replyAttachments = safeInput.attachments;
      if (!Array.isArray(replyAttachments) && loadedFiles && loadedFiles.length > 0) {
        replyAttachments = loadedFiles.map((file) => file.path).filter(Boolean);
      }
      result = await outlookReplyEmail(safeInput.entry_id, safeInput.body, !!safeInput.reply_all, replyAttachments);
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

ipcMain.handle('load-outlook-emails', async (event, config = {}) => {
  try {
    if (mainWindow && mainWindow.webContents) {
      mainWindow.webContents.send('email-progress', { message: 'Connecting to Outlook...', percent: 10 });
    }
    const result = await collectOutlookEmails(config || {});
    if (!result.success) {
      return result;
    }
    if (mainWindow && mainWindow.webContents) {
      mainWindow.webContents.send('email-progress', {
      message: `Loaded ${result.emails.length} Outlook email(s) from Inbox, Sent Items, and subfolders.`,
        percent: 100
      });
    }
    return {
      success: true,
      emails: result.emails,
      summary: result.summary,
      daysBack: result.daysBack,
      folders: result.folders,
      folderPaths: result.folderPaths,
      source: 'outlook'
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('list-outlook-folders', async (event, config = {}) => {
  try {
    const result = await listOutlookFolders(config || {});
    if (!result.success) {
      return result;
    }
    return {
      success: true,
      folders: Array.isArray(result.folders) ? result.folders : []
    };
  } catch (e) {
    return { success: false, error: e.message, folders: [] };
  }
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
- run_update_checklist (requires checklist DOCX; Outlook Inbox, Sent Items, and subfolders are scanned automatically)
- run_generate_punchlist (requires checklist DOCX)
- run_email_ai_search (answers question using loaded emails or by scanning Outlook folders)
- run_general_llm_chat (general Q&A / drafting with the selected provider)
- open_tab
- no_op

Allowed target_tab values:
packets, packetshell, execution, sigblocks, collate, redline, email, updatechecklist, punchlist

App capabilities:
${capabilities.join('\n')}

User command:
${prompt}

Attached files:
${attachmentSummary}

Return JSON only with this exact shape:
{
  "action": "run_signature_packets|run_packet_shell|run_redline|run_collate|run_update_checklist|run_generate_punchlist|run_email_ai_search|run_general_llm_chat|open_tab|no_op",
  "target_tab": "packets|packetshell|execution|sigblocks|collate|redline|email|updatechecklist|punchlist|null",
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

// Convert Anthropic-format tool definitions to OpenAI-format
function convertToolsToOpenAIFormat(tools) {
  return tools.map(t => ({
    type: 'function',
    function: {
      name: t.name,
      description: t.description || '',
      parameters: t.input_schema || { type: 'object', properties: {} }
    }
  }));
}

// Resolve which provider + config to use for the agent command bar
function resolveAgentProvider() {
  const provider = getAIProvider();
  const apiKey = getApiKey();
  const isLocal = provider === 'localai' || provider === 'ollama' || provider === 'lmstudio';

  // Local providers (Ollama, LM Studio) bypass proxy — they run on localhost
  if (!isLocal) {
    const proxyUrl = getAgentProxyUrl();
    if (proxyUrl && providerRequiresApiKey(provider) && !apiKey) {
      // Fallback proxy mode — only used when a direct API-key provider is selected but no key is configured.
      return { mode: 'proxy', proxyUrl, proxyToken: getAgentProxyToken() };
    }
  }

  if (provider === 'anthropic' || (!isLocal && apiKey && apiKey.startsWith('sk-ant-'))) {
    return { mode: 'anthropic', apiKey, provider: 'anthropic' };
  }
  if (provider === 'openai' || (!isLocal && apiKey && apiKey.startsWith('sk-'))) {
    return {
      mode: 'openai',
      apiKey,
      provider: 'openai',
      baseUrl: 'https://api.openai.com',
      models: [getOpenAIModel(), DEFAULT_OPENAI_MODEL, 'gpt-4o-mini', 'gpt-4o']
    };
  }
  if (provider === 'harvey') {
    return {
      mode: 'harvey',
      apiKey,
      provider: 'harvey',
      baseUrl: getHarveyBaseUrl(),
      models: ['harvey-assist']
    };
  }
  if (provider === 'localai') {
    return {
      mode: 'openai',
      apiKey: '',
      provider: 'localai',
      baseUrl: getLocalAIBaseUrl(),
      models: [getLocalAIModel(), DEFAULT_LOCALAI_MODEL, 'qwen2.5:1.5b', 'gemma3:1b']
    };
  }
  if (provider === 'lmstudio') {
    return {
      mode: 'openai',
      apiKey: '',
      provider: 'lmstudio',
      baseUrl: getLmStudioBaseUrl(),
      models: [getLmStudioModel(), DEFAULT_LMSTUDIO_MODEL]
    };
  }
  // Default: Ollama (free, local, no key needed)
  return {
    mode: 'openai',
    apiKey: '',
    provider: 'ollama',
    baseUrl: getOllamaBaseUrl(),
    models: [getOllamaModel(), DEFAULT_OLLAMA_MODEL, 'llama3.1:8b', 'qwen2.5:7b']
  };
}

ipcMain.handle('execute-command', async (event, { prompt, context, conversationHistory }) => {
  const cmdStartedAt = Date.now();
  const toolsCalledLog = [];
  const runId = `run_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const threadId = String(context?.thread_id || '').trim();
  const progressSender = event && event.sender ? event.sender : null;
  let progressFinalized = false;

  const emitAgentProgress = (status, message, extra = {}) => {
    if (!progressSender || !message) return;
    progressSender.send('agent-execution-progress', {
      run_id: runId,
      thread_id: threadId,
      status: String(status || '').trim() || 'update',
      message: String(message || '').trim(),
      is_error: !!extra.is_error,
      tool: extra.tool || null,
      timestamp: new Date().toISOString()
    });
  };

  const finalizeAgentProgress = (result) => {
    if (progressFinalized) return;
    progressFinalized = true;
    if (result && result.success) {
      const completedMessage = String(result.message || '').trim() || 'Agent run completed.';
      emitAgentProgress('complete', completedMessage);
    } else {
      const failedMessage = String((result && result.error) || 'Agent run failed.').trim();
      emitAgentProgress('error', failedMessage, { is_error: true });
    }
  };
  try {
    const agentConfig = resolveAgentProvider();
    emitAgentProgress('start', 'Agent is planning your request...');

    if (agentConfig.provider === 'harvey') {
      const unsupportedResult = {
        success: false,
        error: 'Harvey is configured for planning and general chat, but multi-step tool execution is not available yet. EmmaNeigh will use planner mode instead.'
      };
      finalizeAgentProgress(unsupportedResult);
      return unsupportedResult;
    }

    if (providerRequiresApiKey(agentConfig.provider) && agentConfig.mode !== 'proxy' && !agentConfig.apiKey) {
      const missingKeyResult = {
        success: false,
        error: `No ${getAIProviderDisplayName(agentConfig.provider)} API key configured. Add one in Settings, switch to a local model, or configure a backend proxy.`
      };
      finalizeAgentProgress(missingKeyResult);
      return missingKeyResult;
    }

    const activeTab = context?.activeTab || 'unknown';
    const loadedFiles = context?.loadedFiles || [];
    const isWindows = process.platform === 'win32';

    if (agentConfig.provider === 'localai' || agentConfig.provider === 'ollama' || agentConfig.provider === 'lmstudio') {
      const localProvider = await ensureLocalProviderAvailable(agentConfig.provider, agentConfig.apiKey, {
        allowAutoLaunch: true,
        allowBootstrap: agentConfig.provider === 'localai'
      });
      if (!localProvider.success) {
        const unavailableResult = {
          success: false,
          error: localProvider.error
        };
        finalizeAgentProgress(unavailableResult);
        return unavailableResult;
      }
      agentConfig.baseUrl = localProvider.baseUrl;
      agentConfig.models = Array.from(new Set([
        ...(agentConfig.models || []),
        ...(localProvider.models || [])
      ].filter(Boolean)));
    }

    const systemPrompt = `You are EmmaNeigh, a legal document processing assistant for Kirkland & Ellis.
You help users manage documents, navigate the application, and interact with iManage (the document management system).

Current state:
- Active tab: ${activeTab}
- Loaded files: ${loadedFiles.length > 0 ? loadedFiles.map(f => f.name || f).join(', ') : 'None'}
- Platform: ${process.platform} (${isWindows ? 'iManage available' : 'iManage not available — Windows only'})
- Available tabs: Signature Packets, Packet Shell, Execution, Signature Blocks, Time Tracking, Email Search, Update Checklist, Punchlist Generator, Collate Comments, Redline Documents

When the user asks you to perform an action (save to iManage, open a file, navigate, run a redline, etc.), use the appropriate tool.
You can chain multiple tools across steps in one request until the full task is complete.

Common multi-step workflows you should handle:
- "Save this and redline against precedent" → imanage_save → use the returned profile_id → imanage_redline_versions (auto-detects V1 vs latest)
- "Redline V1 to V2 and email to someone" → imanage_redline_versions with email_to parameter (handles redline + email in one tool call)
- "Redline V1 to V2, CPO, and reply to that email thread" → imanage_redline_versions with change_pages_only=true → outlook_reply_email with the redline as attachment
- "Find document X, check out V3, and compare to V1" → imanage_search → imanage_redline_versions with version_1 and version_2
- "Run redlines for the whole checklist" → run_checklist_precedent_redlines (batch mode)

Key terminology:
- "CPO" = Change Pages Only — a Litera mode that outputs only pages with differences. Set change_pages_only=true on imanage_redline_versions or run_redline.
- "Redline against precedent" = compare V1 (original/precedent) to the latest version.

For imanage_redline_versions: you can omit version_1 and version_2 to auto-detect (V1 vs latest), or specify them explicitly.
If the user wants to email the result, pass email_to directly to imanage_redline_versions — it handles checkout, Litera comparison, and Outlook email in one call.
If the user wants to reply to an existing email thread with the redline, first use outlook_search to find the thread, then imanage_redline_versions to create the redline, then outlook_reply_email with the redline_path as an attachment.

If the user reports iManage connection issues, use imanage_test_connection first to diagnose the problem.
When they ask a question or want advice, respond with helpful text.
Keep responses concise and practical.
${!isWindows ? '\nIMPORTANT: iManage tools are only available on Windows. If the user asks for iManage features, let them know.' : ''}`;

    const session = context && typeof context === 'object'
      ? { actor: context.actor || {}, loadedFiles: context.loadedFiles || [] }
      : { actor: {}, loadedFiles: [] };

    const formatToolSummary = (toolName, toolResult) => {
      const toolLabel = String(toolName || 'tool').replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
      if (!toolResult || typeof toolResult !== 'object') return `- ${toolLabel}: no result returned.`;
      if (!toolResult.success) return `- ${toolLabel}: failed (${toolResult.error || 'unknown error'}).`;
      if (toolResult.message) return `- ${toolLabel}: ${toolResult.message}`;
      if (Array.isArray(toolResult.files)) return `- ${toolLabel}: ${toolResult.files.length} file(s) returned.`;
      if (Array.isArray(toolResult.versions)) return `- ${toolLabel}: ${toolResult.versions.length} version(s) returned.`;
      if (toolResult.output_folder) return `- ${toolLabel}: completed. Output folder: ${toolResult.output_folder}`;
      return `- ${toolLabel}: completed.`;
    };

    const logAndReturn = (result, modelUsed) => {
      finalizeAgentProgress(result);
      logPromptToFirestore({
        username: session.actor?.username || session.actor?.name || 'unknown',
        email: session.actor?.email,
        prompt,
        activeTab,
        toolsCalled: toolsCalledLog,
        toolCount: toolsCalledLog.length,
        modelUsed: modelUsed || 'unknown',
        success: !!result.success,
        durationMs: Date.now() - cmdStartedAt
      }).catch(() => {});
      return result;
    };

    const maxToolTurns = 12;

    // Build initial messages with conversation memory
    const priorHistory = Array.isArray(conversationHistory)
      ? conversationHistory
          .filter(m => m && (m.role === 'user' || m.role === 'assistant') && m.content)
          .slice(-10)
      : [];
    const buildInitialMessages = () => {
      const msgs = [];
      for (const m of priorHistory) {
        msgs.push({ role: m.role, content: String(m.content) });
      }
      msgs.push({ role: 'user', content: prompt });
      return msgs;
    };

    // ── PROXY MODE ──────────────────────────────────────────────────
    // Send the full request to your backend proxy which holds the API key.
    // The proxy is expected to accept the same body as Anthropic /v1/messages
    // and return the same response format.
    if (agentConfig.mode === 'proxy') {
      let messages = buildInitialMessages();
      let lastToolName = null, lastToolInput = null, lastToolResult = null;
      const conversationParts = [];

      for (let turn = 0; turn < maxToolTurns; turn += 1) {
        const headers = { 'Content-Type': 'application/json' };
        if (agentConfig.proxyToken) headers['Authorization'] = `Bearer ${agentConfig.proxyToken}`;
        headers['X-EmmaNeigh-Version'] = APP_VERSION;
        headers['X-EmmaNeigh-Machine'] = MACHINE_ID;

        const response = await requestHttps({
          baseUrl: agentConfig.proxyUrl,
          path: '/v1/agent',
          method: 'POST',
          headers,
          body: JSON.stringify({
            system: systemPrompt,
            tools: COMMAND_TOOLS,
            messages,
            max_tokens: 1200
          }),
          timeoutMs: 60000
        });

        if (response.statusCode !== 200) {
          if (isLikelyModelError(response)) break;
          return logAndReturn({ success: false, error: `Proxy error (${response.statusCode}): ${getErrorDetail(response)}` }, 'proxy');
        }

        const content = Array.isArray(response.jsonBody?.content) ? response.jsonBody.content : [];
        const textBlocks = content.filter(p => p && p.type === 'text' && p.text);
        if (textBlocks.length > 0) {
          const text = textBlocks.map(p => String(p.text || '').trim()).filter(Boolean).join('\n');
          if (text) conversationParts.push(text);
        }

        const toolUses = content.filter(p => p && p.type === 'tool_use' && p.name);
        if (!toolUses.length) {
          return logAndReturn({
            success: true,
            type: lastToolName ? 'tool_result' : 'text',
            tool: lastToolName, input: lastToolInput, toolResult: lastToolResult,
            message: conversationParts.join('\n\n').trim() || 'Command received.',
            modelUsed: response.jsonBody?.model || 'proxy'
          }, response.jsonBody?.model || 'proxy');
        }

        messages.push({ role: 'assistant', content });
        const toolResultBlocks = [];
        const toolSummaries = [];
        for (const toolUse of toolUses) {
          console.log(`[execute-command:proxy] Tool call: ${toolUse.name}`, JSON.stringify(toolUse.input));
          emitAgentProgress('tool_start', `Running tool: ${toolUse.name.replace(/_/g, ' ')}`, { tool: toolUse.name });
          const toolResult = await dispatchTool(toolUse.name, toolUse.input, session);
          emitAgentProgress(
            toolResult && toolResult.success ? 'tool_complete' : 'tool_error',
            formatToolSummary(toolUse.name, toolResult),
            { tool: toolUse.name, is_error: !(toolResult && toolResult.success) }
          );
          lastToolName = toolUse.name; lastToolInput = toolUse.input; lastToolResult = toolResult;
          toolsCalledLog.push(toolUse.name);
          let serialized = '{}';
          try { serialized = JSON.stringify(toolResult); } catch (_) { serialized = '{"success":false,"error":"Serialize error"}'; }
          toolResultBlocks.push({ type: 'tool_result', tool_use_id: toolUse.id, content: serialized, is_error: !toolResult || toolResult.success !== true });
          toolSummaries.push(formatToolSummary(toolUse.name, toolResult));
        }
        if (toolSummaries.length > 0) conversationParts.push(`Executed tools:\n${toolSummaries.join('\n')}`);
        messages.push({ role: 'user', content: toolResultBlocks });
      }

      return logAndReturn({
        success: true, type: lastToolName ? 'tool_result' : 'text',
        tool: lastToolName, input: lastToolInput, toolResult: lastToolResult,
        message: conversationParts.join('\n\n').trim() || 'Reached tool-step limit.',
        modelUsed: 'proxy'
      }, 'proxy');
    }

    // ── ANTHROPIC MODE ──────────────────────────────────────────────
    if (agentConfig.mode === 'anthropic') {
      const modelCandidates = [getClaudeModel(), DEFAULT_CLAUDE_MODEL, 'claude-3-5-sonnet-20241022'];
      let lastResponse = null;

      for (const modelName of modelCandidates) {
        let messages = buildInitialMessages();
        let lastToolName = null, lastToolInput = null, lastToolResult = null;
        const conversationParts = [];
        let modelUnavailable = false;

        for (let turn = 0; turn < maxToolTurns; turn += 1) {
          const response = await requestHttps({
            baseUrl: 'https://api.anthropic.com',
            path: '/v1/messages',
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'x-api-key': agentConfig.apiKey,
              'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({ model: modelName, max_tokens: 1200, system: systemPrompt, tools: COMMAND_TOOLS, messages }),
            timeoutMs: 30000
          });
          lastResponse = response;

          if (response.statusCode !== 200) {
            if (isLikelyModelError(response)) { modelUnavailable = true; break; }
            return logAndReturn({ success: false, error: `Claude API error (${response.statusCode}): ${getErrorDetail(response)}` }, modelName);
          }

          const content = Array.isArray(response.jsonBody?.content) ? response.jsonBody.content : [];
          const textBlocks = content.filter(p => p && p.type === 'text' && p.text);
          if (textBlocks.length > 0) {
            const text = textBlocks.map(p => String(p.text || '').trim()).filter(Boolean).join('\n');
            if (text) conversationParts.push(text);
          }

          const toolUses = content.filter(p => p && p.type === 'tool_use' && p.name);
          if (!toolUses.length) {
            return logAndReturn({
              success: true, type: lastToolName ? 'tool_result' : 'text',
              tool: lastToolName, input: lastToolInput, toolResult: lastToolResult,
              message: conversationParts.join('\n\n').trim() || 'Command received but no actionable response was generated.',
              modelUsed: modelName
            }, modelName);
          }

          messages.push({ role: 'assistant', content });
          const toolResultBlocks = [];
          const toolSummaries = [];
          for (const toolUse of toolUses) {
            console.log(`[execute-command:anthropic] Tool call: ${toolUse.name}`, JSON.stringify(toolUse.input));
            emitAgentProgress('tool_start', `Running tool: ${toolUse.name.replace(/_/g, ' ')}`, { tool: toolUse.name });
            const toolResult = await dispatchTool(toolUse.name, toolUse.input, session);
            emitAgentProgress(
              toolResult && toolResult.success ? 'tool_complete' : 'tool_error',
              formatToolSummary(toolUse.name, toolResult),
              { tool: toolUse.name, is_error: !(toolResult && toolResult.success) }
            );
            lastToolName = toolUse.name; lastToolInput = toolUse.input; lastToolResult = toolResult;
            toolsCalledLog.push(toolUse.name);
            let serialized = '{}';
            try { serialized = JSON.stringify(toolResult); } catch (_) { serialized = '{"success":false,"error":"Serialize error"}'; }
            toolResultBlocks.push({ type: 'tool_result', tool_use_id: toolUse.id, content: serialized, is_error: !toolResult || toolResult.success !== true });
            toolSummaries.push(formatToolSummary(toolUse.name, toolResult));
          }
          if (toolSummaries.length > 0) conversationParts.push(`Executed tools:\n${toolSummaries.join('\n')}`);
          messages.push({ role: 'user', content: toolResultBlocks });
        }

        if (modelUnavailable) continue;

        return logAndReturn({
          success: true, type: lastToolName ? 'tool_result' : 'text',
          tool: lastToolName, input: lastToolInput, toolResult: lastToolResult,
          message: conversationParts.join('\n\n').trim() || 'Reached tool-step limit.',
          modelUsed: modelName
        }, modelName);
      }

      return logAndReturn({ success: false, error: `No supported Claude model available. Last error: ${getErrorDetail(lastResponse)}` }, 'none');
    }

    // ── OPENAI-COMPATIBLE MODE (Local AI, Ollama, LM Studio, OpenAI) ──────────
    const openaiTools = convertToolsToOpenAIFormat(COMMAND_TOOLS);
    const isLocalOllama = agentConfig.provider === 'localai' || agentConfig.provider === 'ollama' || agentConfig.provider === 'lmstudio';

    // Discover available models from the provider (Ollama, LM Studio, OpenAI all support /v1/models)
    let modelCandidates = agentConfig.models || [DEFAULT_OLLAMA_MODEL];
    if (isLocalOllama) {
      try {
        const discovered = await fetchOpenAICompatibleModels({ baseUrl: agentConfig.baseUrl, apiKey: agentConfig.apiKey });
        if (discovered.length > 0) {
          // Merge: preferred models first, then discovered ones
          modelCandidates = Array.from(new Set([...modelCandidates, ...discovered]));
        }
      } catch (_) {}
    }
    let lastResponse = null;

    for (const modelName of modelCandidates) {
      const initialMsgs = buildInitialMessages();
      let messages = [
        { role: 'system', content: systemPrompt },
        ...initialMsgs
      ];
      let lastToolName = null, lastToolInput = null, lastToolResult = null;
      const conversationParts = [];
      let modelUnavailable = false;
      let toolsDisabled = false; // Track if tools were disabled due to model not supporting them

      for (let turn = 0; turn < maxToolTurns; turn += 1) {
        const headers = { 'Content-Type': 'application/json' };
        if (agentConfig.apiKey) headers['Authorization'] = `Bearer ${agentConfig.apiKey}`;

        const requestBody = {
          model: modelName,
          max_tokens: 1200,
          temperature: 0,
          stream: false, // Explicit: ensure non-streaming response from Ollama/LM Studio
          messages
        };
        // Only include tools if not disabled (some local models don't support tool calling)
        if (!toolsDisabled) {
          requestBody.tools = openaiTools;
        }

        const response = await requestHttps({
          baseUrl: agentConfig.baseUrl,
          path: '/v1/chat/completions',
          method: 'POST',
          headers,
          body: JSON.stringify(requestBody),
          timeoutMs: 90000
        });
        lastResponse = response;

        if (response.statusCode !== 200) {
          // If tools caused the error (common with local models), retry without tools
          if (!toolsDisabled && isLocalOllama && (response.statusCode === 400 || response.statusCode === 500)) {
            const errDetail = getErrorDetail(response).toLowerCase();
            if (errDetail.includes('tool') || errDetail.includes('function') || errDetail.includes('not supported') || errDetail.includes('invalid') || errDetail.includes('schema')) {
              console.log(`[execute-command:${agentConfig.provider}] Tool calling not supported by ${modelName}, retrying without tools`);
              toolsDisabled = true;
              continue; // Retry same turn without tools
            }
          }
          if (isLikelyModelError(response)) { modelUnavailable = true; break; }
          return logAndReturn({
            success: false,
            error: `${agentConfig.provider} error (${response.statusCode}): ${getErrorDetail(response)}`
          }, modelName);
        }

        const choice = response.jsonBody?.choices?.[0];
        const assistantMsg = choice?.message;
        if (!assistantMsg) {
          return logAndReturn({ success: false, error: `${agentConfig.provider} response missing message.` }, modelName);
        }

        // Extract text
        const text = String(assistantMsg.content || '').trim();
        if (text) conversationParts.push(text);

        // Extract tool calls (OpenAI format)
        const toolCalls = Array.isArray(assistantMsg.tool_calls) ? assistantMsg.tool_calls.filter(tc => tc && tc.function) : [];
        if (!toolCalls.length) {
          return logAndReturn({
            success: true, type: lastToolName ? 'tool_result' : 'text',
            tool: lastToolName, input: lastToolInput, toolResult: lastToolResult,
            message: conversationParts.join('\n\n').trim() || 'Command received.',
            modelUsed: modelName
          }, modelName);
        }

        // Add assistant message to conversation (with tool_calls)
        messages.push(assistantMsg);

        const toolSummaries = [];
        for (const tc of toolCalls) {
          const fnName = tc.function?.name || '';
          let fnArgs = {};
          try { fnArgs = JSON.parse(tc.function?.arguments || '{}'); } catch (_) {}

          console.log(`[execute-command:${agentConfig.provider}] Tool call: ${fnName}`, JSON.stringify(fnArgs));
          emitAgentProgress('tool_start', `Running tool: ${fnName.replace(/_/g, ' ')}`, { tool: fnName });
          const toolResult = await dispatchTool(fnName, fnArgs, session);
          emitAgentProgress(
            toolResult && toolResult.success ? 'tool_complete' : 'tool_error',
            formatToolSummary(fnName, toolResult),
            { tool: fnName, is_error: !(toolResult && toolResult.success) }
          );
          lastToolName = fnName; lastToolInput = fnArgs; lastToolResult = toolResult;
          toolsCalledLog.push(fnName);

          let serialized = '{}';
          try { serialized = JSON.stringify(toolResult); } catch (_) { serialized = '{"success":false,"error":"Serialize error"}'; }

          // OpenAI tool result format: role=tool, tool_call_id, content
          // Generate a fallback ID if Ollama doesn't provide one (some versions omit tc.id)
          const toolCallId = tc.id || `call_${fnName}_${turn}_${Date.now()}`;
          messages.push({
            role: 'tool',
            tool_call_id: toolCallId,
            content: serialized
          });
          toolSummaries.push(formatToolSummary(fnName, toolResult));
        }
        if (toolSummaries.length > 0) conversationParts.push(`Executed tools:\n${toolSummaries.join('\n')}`);
      }

      if (modelUnavailable) continue;

      return logAndReturn({
        success: true, type: lastToolName ? 'tool_result' : 'text',
        tool: lastToolName, input: lastToolInput, toolResult: lastToolResult,
        message: conversationParts.join('\n\n').trim() || 'Reached tool-step limit.',
        modelUsed: modelName
      }, modelName);
    }

    return logAndReturn({
      success: false,
      error: `No supported ${agentConfig.provider} model available. Ensure ${agentConfig.provider === 'localai'
        ? 'Local AI is installed from Settings'
        : agentConfig.provider === 'ollama'
          ? 'Ollama is running (ollama serve) and a model is pulled (ollama pull llama3.1:8b)'
          : 'the provider is reachable'}. Last error: ${getErrorDetail(lastResponse)}`
    }, 'none');
  } catch (e) {
    console.error('[execute-command] Error:', e);
    finalizeAgentProgress({ success: false, error: e.message });
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

ipcMain.handle('get-local-ai-profiles', async () => {
  return {
    success: true,
    profiles: getLocalAIProfiles(),
    selectedModel: getLocalAIModel(),
    baseUrl: getLocalAIBaseUrl()
  };
});

ipcMain.handle('get-local-ai-status', async () => {
  try {
    const status = await ensureLocalProviderAvailable('localai', '', {
      allowAutoLaunch: true,
      allowBootstrap: false
    });
    return {
      success: true,
      ready: !!status.success,
      configuredModel: getLocalAIModel(),
      baseUrl: getLocalAIBaseUrl(),
      models: status.models || [],
      message: status.success ? status.message : status.error
    };
  } catch (e) {
    return {
      success: false,
      ready: false,
      configuredModel: getLocalAIModel(),
      baseUrl: getLocalAIBaseUrl(),
      models: [],
      message: e.message
    };
  }
});

ipcMain.handle('bootstrap-local-ai', async (event, payload) => {
  try {
    const profileId = String(payload?.model || getLocalAIModel()).trim() || DEFAULT_LOCALAI_MODEL;
    const result = await bootstrapManagedLocalAi(profileId);
    if (!result.success) {
      return { success: false, error: result.error || 'Failed to install Local AI.' };
    }
    return {
      success: true,
      model: result.model,
      baseUrl: result.baseUrl,
      profile: result.profile,
      message: `Local AI installed successfully using ${result.profile?.label || result.model}.`
    };
  } catch (e) {
    return { success: false, error: e.message };
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
    if (provider === 'localai') {
      const localProvider = await ensureLocalProviderAvailable('localai', '', {
        allowAutoLaunch: true,
        allowBootstrap: true
      });
      if (!localProvider.success) {
        return { success: false, error: localProvider.error };
      }
      const configuredModel = getLocalAIModel();
      return {
        success: true,
        message: `${localProvider.message} Active model: ${configuredModel}. Available models: ${localProvider.models.slice(0, 5).join(', ')}${localProvider.models.length > 5 ? '...' : ''}`,
        provider,
        models: localProvider.models,
        configuredModel,
        configuredModelAvailable: localProvider.models.includes(configuredModel),
        baseUrl: localProvider.baseUrl
      };
    }

    if (provider === 'ollama') {
      const localProvider = await ensureLocalProviderAvailable('ollama', apiKey, { allowAutoLaunch: true });
      if (!localProvider.success) {
        return { success: false, error: localProvider.error };
      }
      const configuredModel = getOllamaModel();
      return {
        success: true,
        message: `${localProvider.message} Available models: ${localProvider.models.slice(0, 5).join(', ')}${localProvider.models.length > 5 ? '...' : ''}`,
        provider,
        models: localProvider.models,
        configuredModel,
        configuredModelAvailable: localProvider.models.includes(configuredModel),
        baseUrl: localProvider.baseUrl
      };
    }

    if (provider === 'lmstudio') {
      const localProvider = await ensureLocalProviderAvailable('lmstudio', apiKey, { allowAutoLaunch: true });
      if (!localProvider.success) {
        return { success: false, error: localProvider.error };
      }
      return {
        success: true,
        message: `${localProvider.message} Available models: ${localProvider.models.slice(0, 5).join(', ')}${localProvider.models.length > 5 ? '...' : ''}`,
        provider,
        models: localProvider.models,
        configuredModel: getLmStudioModel(),
        baseUrl: localProvider.baseUrl
      };
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

// ── Agent proxy settings ────────────────────────────────────────────
ipcMain.handle('get-agent-proxy', async () => {
  return {
    success: true,
    url: getAgentProxyUrl(),
    hasToken: !!getAgentProxyToken()
  };
});

ipcMain.handle('set-agent-proxy', async (event, { url, token }) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    const normalizedUrl = String(url || '').trim();
    db.run(`INSERT OR REPLACE INTO app_settings (key, value) VALUES ('${AGENT_PROXY_URL_KEY}', '${normalizedUrl.replace(/'/g, "''")}')`);
    if (token !== undefined) {
      const encoded = encodeSecretToSetting(String(token || '').trim());
      db.run(`INSERT OR REPLACE INTO app_settings (key, value) VALUES ('${AGENT_PROXY_TOKEN_KEY}', '${encoded.replace(/'/g, "''")}')`);
    }
    saveDatabase();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('get-agent-config', async () => {
  const config = resolveAgentProvider();
  return {
    success: true,
    mode: config.mode,
    provider: config.provider || config.mode,
    hasProxy: config.mode === 'proxy',
    hasApiKey: config.mode === 'anthropic' ? !!config.apiKey : true,
    models: config.models || []
  };
});

// ========== USER ACCOUNT HANDLERS ==========

// Email-only login (mandatory email identity)
ipcMain.handle('email-login', async (event, { email, displayName }) => {
  if (!db) return { success: false, error: 'Database not initialized' };

  const normalizedEmail = normalizeEmail(email);
  if (!isValidEmail(normalizedEmail)) {
    return { success: false, error: 'Please enter a valid email address' };
  }

  const loginAccess = await evaluateAccessPolicy(
    { email: normalizedEmail, username: buildUsernameFromEmail(normalizedEmail) },
    { eventType: 'login' }
  );
  if (!loginAccess.allowed) {
    return {
      success: false,
      blocked: true,
      error: loginAccess.message || 'Access has been disabled by your administrator.'
    };
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

  const createAccess = await evaluateAccessPolicy(
    { email: normalizedEmail, username: username || buildUsernameFromEmail(normalizedEmail) },
    { eventType: 'account_create' }
  );
  if (!createAccess.allowed) {
    return {
      success: false,
      blocked: true,
      error: createAccess.message || 'Account creation is currently disabled for this user.'
    };
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

  const loginAccess = await evaluateAccessPolicy(
    { email: normalizedEmail, username },
    { eventType: 'login' }
  );
  if (!loginAccess.allowed) {
    return {
      success: false,
      blocked: true,
      error: loginAccess.message || 'Access has been disabled by your administrator.'
    };
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

    const loginAccess = await evaluateAccessPolicy(
      { email: normalizedEmail, username, userId: id },
      { eventType: 'login_2fa' }
    );
    if (!loginAccess.allowed) {
      return {
        success: false,
        blocked: true,
        error: loginAccess.message || 'Access has been disabled by your administrator.'
      };
    }

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

    const sessionAccess = await evaluateAccessPolicy(
      { email: normalizedEmail, username: uname, userId: id },
      { eventType: 'session_restore' }
    );
    if (!sessionAccess.allowed) {
      return {
        success: false,
        blocked: true,
        error: sessionAccess.message || 'Access has been disabled by your administrator.'
      };
    }

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
    if (key === ACCESS_POLICY_URL_KEY || key === ACCESS_POLICY_TOKEN_KEY || key === ACCESS_POLICY_FAIL_CLOSED_KEY) {
      clearAccessPolicyCache();
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

ipcMain.handle('get-access-policy-config', async () => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    return {
      success: true,
      url: getAccessPolicyUrl(),
      hasToken: !!getAccessPolicyToken(),
      failClosed: getAccessPolicyFailClosed()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('save-access-policy-config', async (event, config = {}) => {
  if (!db) return { success: false, error: 'Database not initialized' };
  try {
    const rawUrl = String(config.url || '').trim();
    const normalizedUrl = normalizeIngestUrl(rawUrl);
    if (!normalizedUrl) {
      return { success: false, error: 'Access policy URL must be a valid http(s) URL.' };
    }

    const rawToken = typeof config.token === 'string' ? config.token.trim() : '';
    const keepExistingToken = !!config.keepExistingToken;
    const failClosed = normalizeBooleanSetting(config.failClosed, false);

    if (!writeSettingValue(ACCESS_POLICY_URL_KEY, normalizedUrl)) {
      return { success: false, error: 'Failed to save access policy URL.' };
    }

    if (rawToken) {
      if (!writeSettingValue(ACCESS_POLICY_TOKEN_KEY, encodeSecretForSetting(rawToken))) {
        return { success: false, error: 'Failed to save access policy token.' };
      }
    } else if (!keepExistingToken) {
      if (!writeSettingValue(ACCESS_POLICY_TOKEN_KEY, '')) {
        return { success: false, error: 'Failed to clear access policy token.' };
      }
    }

    if (!writeSettingValue(ACCESS_POLICY_FAIL_CLOSED_KEY, failClosed ? '1' : '0')) {
      return { success: false, error: 'Failed to save fail-closed setting.' };
    }

    saveDatabase();
    clearAccessPolicyCache();

    const testResult = await evaluateAccessPolicy(
      { email: config.email || 'admin@example.com', username: config.username || 'admin' },
      { eventType: 'policy_test', force: true }
    );

    return {
      success: true,
      status: testResult
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('test-access-policy-config', async (event, config = {}) => {
  try {
    const identity = config.identity && typeof config.identity === 'object' ? config.identity : {};
    const result = await evaluateAccessPolicy(
      {
        email: identity.email || '',
        username: identity.username || '',
        userId: identity.userId || ''
      },
      { eventType: String(config.eventType || 'policy_test'), force: true }
    );
    return { success: true, status: result };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('check-access-policy', async (event, identity = {}) => {
  try {
    const result = await evaluateAccessPolicy(
      {
        email: identity.email || '',
        username: identity.username || '',
        userId: identity.userId || ''
      },
      {
        eventType: String(identity.eventType || identity.context || 'session'),
        force: !!identity.force
      }
    );
    return {
      success: true,
      configured: !!result.configured,
      allowed: !!result.allowed,
      message: result.message || '',
      source: result.source || 'unknown'
    };
  } catch (e) {
    return { success: false, allowed: true, message: e.message };
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

function getLocalAIBaseUrl() {
  return normalizeProviderBaseUrl(getSettingValue('localai_base_url', DEFAULT_LOCALAI_BASE_URL), DEFAULT_LOCALAI_BASE_URL);
}

function getLocalAIModel() {
  const model = (getSettingValue('localai_model', DEFAULT_LOCALAI_MODEL) || '').trim();
  return model || DEFAULT_LOCALAI_MODEL;
}

function getOllamaBaseUrl() {
  return normalizeProviderBaseUrl(getSettingValue('ollama_base_url', DEFAULT_OLLAMA_BASE_URL), DEFAULT_OLLAMA_BASE_URL);
}

function getLmStudioBaseUrl() {
  return normalizeProviderBaseUrl(getSettingValue('lmstudio_base_url', DEFAULT_LMSTUDIO_BASE_URL), DEFAULT_LMSTUDIO_BASE_URL);
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

// ── Agent proxy settings ────────────────────────────────────────────
function getAgentProxyUrl() {
  const raw = (getSettingValue(AGENT_PROXY_URL_KEY, '') || '').trim();
  if (!raw) return '';
  return raw.endsWith('/') ? raw.slice(0, -1) : raw;
}

function getAgentProxyToken() {
  return decodeSecretFromSetting(getSettingValue(AGENT_PROXY_TOKEN_KEY, '')) || '';
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
    : provider === 'localai' ? getLocalAIModel()
    : provider === 'ollama' ? getOllamaModel()
    : provider === 'lmstudio' ? getLmStudioModel()
    : '';
  let providerBaseUrl =
    provider === 'harvey' ? getHarveyBaseUrl()
    : provider === 'localai' ? getLocalAIBaseUrl()
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
  if (provider === 'localai' || provider === 'ollama' || provider === 'lmstudio') {
    const localProvider = await ensureLocalProviderAvailable(provider, apiKey, {
      allowAutoLaunch: true,
      allowBootstrap: provider === 'localai'
    });
    if (!localProvider.success) {
      return { success: false, error: localProvider.error };
    }
    providerBaseUrl = localProvider.baseUrl;
  }

  let emailInputPath = '';
  let emailSource = 'outlook';
  let cleanupEmailInputPath = '';
  let collectedEmailCount = 0;
  let collectedEmailDaysBack = Number(config.daysBack || 30);

  if (Array.isArray(config.emails) && config.emails.length) {
    const normalizedEmails = config.emails.map((email) => normalizeLoadedEmail(email)).filter((email) => (
      email.subject || email.body || email.from || email.to || email.attachments
    ));
    if (!normalizedEmails.length) {
      return { success: false, error: 'No email activity was provided for checklist analysis.' };
    }
    cleanupEmailInputPath = writeTempEmailDataset(normalizedEmails);
    emailInputPath = cleanupEmailInputPath;
    emailSource = 'memory';
    collectedEmailCount = normalizedEmails.length;
  } else if (emailPath && fs.existsSync(emailPath)) {
    emailInputPath = emailPath;
    emailSource = 'csv';
  } else {
    if (mainWindow && mainWindow.webContents) {
      mainWindow.webContents.send('checklist-progress', {
        message: 'Scanning Outlook email folders...',
        percent: 12
      });
    }
    const outlookResult = await collectOutlookEmails({
      daysBack: collectedEmailDaysBack,
      folders: Array.isArray(config.folders) && config.folders.length ? config.folders : ['Inbox', 'Sent Items'],
      folderPaths: Array.isArray(config.folderPaths) ? config.folderPaths : [],
      maxResultsPerFolder: config.maxResultsPerFolder || 250,
      bodyCharLimit: config.bodyCharLimit || 4000
    });
    if (!outlookResult.success) {
      return {
        success: false,
        error: outlookResult.error || 'Failed to scan Outlook email folders.'
      };
    }
    if (!Array.isArray(outlookResult.emails) || !outlookResult.emails.length) {
      return {
        success: false,
        error: `No Outlook emails were found in Inbox or Sent Items for the last ${outlookResult.daysBack || collectedEmailDaysBack} days.`
      };
    }
    cleanupEmailInputPath = writeTempEmailDataset(outlookResult.emails);
    emailInputPath = cleanupEmailInputPath;
    emailSource = 'outlook';
    collectedEmailCount = outlookResult.emails.length;
    collectedEmailDaysBack = outlookResult.daysBack || collectedEmailDaysBack;
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
      args = [clModuleName, checklistPath, emailInputPath, outputFolder, apiKey, provider, providerModel, providerBaseUrl];
    } else {
      args = [clModuleName, path.join(__dirname, 'python', 'checklist_updater.py'), checklistPath, emailInputPath, outputFolder, apiKey, provider, providerModel, providerBaseUrl];
    }

    if (providerAutoDetectNote) {
      mainWindow.webContents.send('checklist-progress', { message: providerAutoDetectNote, percent: 10 });
    }
    mainWindow.webContents.send('checklist-progress', {
      message: `Analyzing checklist and ${emailSource === 'outlook' ? 'Outlook activity' : 'email activity'} with ${providerName}...`,
      percent: 20
    });

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
      if (cleanupEmailInputPath) {
        try { fs.unlinkSync(cleanupEmailInputPath); } catch (_) {}
      }
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
        emailsProcessed: result.emails_processed || collectedEmailCount,
        emailSource,
        daysBack: emailSource === 'outlook' ? collectedEmailDaysBack : null,
        folderPaths: emailSource === 'outlook' ? (Array.isArray(config.folderPaths) ? config.folderPaths : []) : [],
        error: result.error
      });
    });

    proc.on('error', (err) => {
      if (cleanupEmailInputPath) {
        try { fs.unlinkSync(cleanupEmailInputPath); } catch (_) {}
      }
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
    : provider === 'localai' ? getLocalAIModel()
    : provider === 'ollama' ? getOllamaModel()
    : provider === 'lmstudio' ? getLmStudioModel()
    : '';
  let providerBaseUrl =
    provider === 'harvey' ? getHarveyBaseUrl()
    : provider === 'localai' ? getLocalAIBaseUrl()
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
  if (provider === 'localai' || provider === 'ollama' || provider === 'lmstudio') {
    const localProvider = await ensureLocalProviderAvailable(provider, apiKey, {
      allowAutoLaunch: true,
      allowBootstrap: provider === 'localai'
    });
    if (!localProvider.success) {
      return { success: false, error: localProvider.error };
    }
    providerBaseUrl = localProvider.baseUrl;
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
