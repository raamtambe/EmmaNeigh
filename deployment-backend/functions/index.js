const { onRequest } = require('firebase-functions/v2/https');
const { defineSecret } = require('firebase-functions/params');
const logger = require('firebase-functions/logger');
const admin = require('firebase-admin');
const crypto = require('crypto');

admin.initializeApp();
const db = admin.firestore();

const TELEMETRY_TOKEN = defineSecret('EMMANEIGH_TELEMETRY_TOKEN');
const ACCESS_POLICY_TOKEN = defineSecret('EMMANEIGH_ACCESS_POLICY_TOKEN');
const GROQ_API_KEY = defineSecret('EMMANEIGH_GROQ_API_KEY');
const SESSION_SIGNING_KEY = defineSecret('EMMANEIGH_SESSION_SIGNING_KEY');

const ALLOWED_TELEMETRY_COLLECTIONS = new Set([
  'telemetry_health',
  'usage_history',
  'user_feedback',
  'prompt_logs',
  'user_activity_history'
]);

const DEFAULT_MANAGED_AI_PROVIDER = String(process.env.EMMANEIGH_MANAGED_AI_PROVIDER || 'groq').trim().toLowerCase() || 'groq';
const DEFAULT_MANAGED_AI_MODEL = String(process.env.EMMANEIGH_MANAGED_AI_MODEL || 'qwen/qwen3-32b').trim() || 'qwen/qwen3-32b';
const GROQ_OPENAI_BASE_URL = 'https://api.groq.com/openai/v1';
const MANAGED_AI_SESSION_ISSUER = 'emmaneigh-managed-ai';

function getManagedAiSessionTtlSeconds() {
  const parsed = Number.parseInt(process.env.EMMANEIGH_MANAGED_AI_SESSION_TTL_SECONDS || '3600', 10);
  if (!Number.isFinite(parsed)) return 3600;
  return Math.max(300, Math.min(parsed, 86400));
}

function parseJsonBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (!req.rawBody) return {};
  try {
    return JSON.parse(Buffer.from(req.rawBody).toString('utf8'));
  } catch (_) {
    return {};
  }
}

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeUsername(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeDocId(value) {
  return encodeURIComponent(String(value || '').trim().toLowerCase());
}

function normalizeBoolean(value, fallbackValue = false) {
  if (value === null || value === undefined || value === '') return fallbackValue;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return fallbackValue;
  if (['1', 'true', 'yes', 'y', 'on', 'enabled', 'allow', 'allowed', 'active'].includes(normalized)) return true;
  if (['0', 'false', 'no', 'n', 'off', 'disabled', 'deny', 'denied', 'blocked', 'inactive'].includes(normalized)) return false;
  return fallbackValue;
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

function evaluateAccessPolicyDecision(policyRoot, identity = {}) {
  const policy = policyRoot && typeof policyRoot === 'object' ? policyRoot : {};
  const email = normalizeEmail(identity.email || '');
  const username = normalizeUsername(identity.username || '');
  const userId = String(identity.user_id || identity.userId || '').trim().toLowerCase();
  const emailDomain = email.includes('@') ? email.split('@')[1] : '';
  const emailLocalPart = email.includes('@') ? email.split('@')[0] : '';
  const message =
    String(
      policy.message ||
      policy.reason ||
      policy.denied_message ||
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
  const directAllow = normalizeBoolean(directAllowRaw, true);
  const directBlock = normalizeBoolean(directBlockRaw, false);

  const statusText = String(policy.status || '').trim().toLowerCase();
  if (['blocked', 'disabled', 'deny', 'revoked'].includes(statusText)) {
    return { allowed: false, isAdmin: false, message };
  }
  if (normalizeBoolean(policy.block_all ?? policy.disable_all ?? policy.kill_switch ?? false, false)) {
    return { allowed: false, isAdmin: false, message };
  }
  if (hasDirectBlock && directBlock) {
    return { allowed: false, isAdmin: false, message };
  }
  if (hasDirectAllow && !directAllow) {
    return { allowed: false, isAdmin: false, message };
  }

  const blockedEmails = getPolicyList(policy, ['blocked_emails', 'blockedEmails', 'denylist', 'denied_emails', 'banned_emails']);
  const blockedDomains = getPolicyList(policy, ['blocked_domains', 'blockedDomains', 'denied_domains', 'banned_domains']);
  const blockedUsers = getPolicyList(policy, ['blocked_users', 'blockedUsers', 'blocked_usernames', 'blocked_user_ids', 'blockedUserIds']);
  const allowedEmails = getPolicyList(policy, ['allowed_emails', 'allowedEmails', 'allowlist', 'approved_emails']);
  const allowedDomains = getPolicyList(policy, ['allowed_domains', 'allowedDomains', 'approved_domains']);
  const allowedUsers = getPolicyList(policy, ['allowed_users', 'allowedUsers', 'allowed_usernames', 'allowed_user_ids', 'allowedUserIds']);
  const adminEmails = getPolicyList(policy, ['admin_emails', 'adminEmails']);
  const adminDomains = getPolicyList(policy, ['admin_domains', 'adminDomains']);
  const adminUsers = getPolicyList(policy, ['admin_users', 'adminUsers', 'admin_usernames', 'adminUsernames']);
  const hasAdminRules = adminEmails.length > 0 || adminDomains.length > 0 || adminUsers.length > 0;
  const directAdminRaw = policy.is_admin ?? policy.isAdmin ?? policy.admin;
  const roleText = String(policy.role || '').trim().toLowerCase();
  const isAdminFromRules = (
    (!!email && (listHasMatch(adminEmails, email) || listHasMatch(adminEmails, emailLocalPart))) ||
    (!!emailDomain && listHasMatch(adminDomains, emailDomain)) ||
    (!!username && listHasMatch(adminUsers, username)) ||
    (!!userId && listHasMatch(adminUsers, userId))
  );
  const isAdmin = hasAdminRules
    ? isAdminFromRules
    : (normalizeBoolean(directAdminRaw, false) || roleText === 'admin' || roleText === 'owner');

  if (email && listHasMatch(blockedEmails, email)) return { allowed: false, isAdmin, message };
  if (emailDomain && listHasMatch(blockedDomains, emailDomain)) return { allowed: false, isAdmin, message };
  if ((username && listHasMatch(blockedUsers, username)) || (userId && listHasMatch(blockedUsers, userId))) {
    return { allowed: false, isAdmin, message };
  }

  const hasAllowRules = allowedEmails.length > 0 || allowedDomains.length > 0 || allowedUsers.length > 0;
  if (hasAllowRules) {
    const emailAllowed = email && (listHasMatch(allowedEmails, email) || listHasMatch(allowedEmails, emailLocalPart));
    const domainAllowed = emailDomain && listHasMatch(allowedDomains, emailDomain);
    const userAllowed = (username && listHasMatch(allowedUsers, username)) || (userId && listHasMatch(allowedUsers, userId));
    if (!(emailAllowed || domainAllowed || userAllowed)) {
      return { allowed: false, isAdmin, message };
    }
  }

  return {
    allowed: true,
    isAdmin,
    message: String(policy.allowed_message || 'Access granted.').trim()
  };
}

function isManagedAiAllowedByPolicy(policy) {
  const allowedProviders = getPolicyList(policy, ['allowed_ai_providers', 'allowedAiProviders']);
  if (allowedProviders.length === 0) return true;
  const provider = getManagedProvider();
  return (
    listHasMatch(allowedProviders, 'managed') ||
    listHasMatch(allowedProviders, provider) ||
    listHasMatch(allowedProviders, '*') ||
    listHasMatch(allowedProviders, 'all')
  );
}

function base64UrlEncode(value) {
  const raw = typeof value === 'string' ? value : JSON.stringify(value);
  return Buffer.from(raw, 'utf8').toString('base64url');
}

function decodeBase64UrlJson(value) {
  const decoded = Buffer.from(String(value || ''), 'base64url').toString('utf8');
  return JSON.parse(decoded);
}

function safeTimingEqual(a, b) {
  const left = Buffer.from(String(a || ''), 'utf8');
  const right = Buffer.from(String(b || ''), 'utf8');
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function issueManagedAiSessionToken(identity, options = {}, signingKey) {
  const key = String(signingKey || '').trim();
  if (!key) {
    throw new Error('Managed AI session signing key is not configured.');
  }

  const nowSeconds = Math.floor(Date.now() / 1000);
  const ttlSeconds = getManagedAiSessionTtlSeconds();
  const payload = {
    iss: MANAGED_AI_SESSION_ISSUER,
    aud: 'emmaneigh-desktop',
    iat: nowSeconds,
    exp: nowSeconds + ttlSeconds,
    jti: crypto.randomUUID(),
    sub: String(identity.user_id || identity.email || identity.username || 'emmaneigh-user').trim(),
    email: normalizeEmail(identity.email || ''),
    username: normalizeUsername(identity.username || ''),
    user_id: String(identity.user_id || identity.userId || '').trim(),
    machine_id: String(identity.machine_id || '').trim(),
    app_version: String(identity.app_version || '').trim(),
    display_name: String(identity.display_name || '').trim(),
    is_admin: !!options.isAdmin,
    managed_provider: getManagedProvider(),
    managed_model: getManagedModel(options.requestedModel || '')
  };

  const headerPart = base64UrlEncode({ alg: 'HS256', typ: 'JWT' });
  const payloadPart = base64UrlEncode(payload);
  const signature = crypto.createHmac('sha256', key).update(`${headerPart}.${payloadPart}`).digest('base64url');

  return {
    token: `${headerPart}.${payloadPart}.${signature}`,
    payload,
    ttlSeconds,
    expiresAt: new Date(payload.exp * 1000).toISOString()
  };
}

function verifyManagedAiSessionToken(token, signingKey) {
  const rawToken = String(token || '').trim();
  const key = String(signingKey || '').trim();
  if (!rawToken) {
    return { ok: false, statusCode: 401, code: 'missing_token', error: 'Managed AI session token is required.' };
  }
  if (!key) {
    return { ok: false, statusCode: 500, code: 'missing_signing_key', error: 'Managed AI session signing key is not configured.' };
  }

  const parts = rawToken.split('.');
  if (parts.length !== 3) {
    return { ok: false, statusCode: 401, code: 'invalid_token', error: 'Managed AI session token is malformed.' };
  }

  const [headerPart, payloadPart, signaturePart] = parts;
  const expectedSignature = crypto.createHmac('sha256', key).update(`${headerPart}.${payloadPart}`).digest('base64url');
  if (!safeTimingEqual(signaturePart, expectedSignature)) {
    return { ok: false, statusCode: 401, code: 'invalid_signature', error: 'Managed AI session token signature is invalid.' };
  }

  let payload;
  try {
    payload = decodeBase64UrlJson(payloadPart);
  } catch (_) {
    return { ok: false, statusCode: 401, code: 'invalid_payload', error: 'Managed AI session token payload is invalid.' };
  }

  const nowSeconds = Math.floor(Date.now() / 1000);
  if (String(payload.iss || '').trim() !== MANAGED_AI_SESSION_ISSUER) {
    return { ok: false, statusCode: 401, code: 'invalid_issuer', error: 'Managed AI session token issuer is invalid.' };
  }
  if (!payload.exp || Number(payload.exp) <= nowSeconds) {
    return { ok: false, statusCode: 401, code: 'session_expired', error: 'Managed AI session has expired.' };
  }

  return { ok: true, payload };
}

function buildManagedAiSessionIdentity(req) {
  const body = parseJsonBody(req);
  return {
    email: req.query.email || body.email || '',
    username: req.query.username || body.username || '',
    user_id: req.query.user_id || body.user_id || body.userId || '',
    display_name: req.query.display_name || body.display_name || body.displayName || '',
    machine_id: req.query.machine_id || body.machine_id || body.machineId || '',
    app_version: req.query.app_version || body.app_version || body.appVersion || ''
  };
}

function buildManagedAiUnauthorizedPayload(result) {
  return {
    success: false,
    error: result.error || 'Unauthorized',
    code: result.code || 'unauthorized',
    session_expired: result.code === 'session_expired'
  };
}

function readBearerToken(req) {
  const authHeader = String(req.get('authorization') || '').trim();
  if (authHeader.toLowerCase().startsWith('bearer ')) {
    return authHeader.slice(7).trim();
  }
  return String(
    req.get('x-emmaneigh-telemetry-key') ||
    req.get('x-emmaneigh-access-key') ||
    req.get('x-emmaneigh-ai-key') ||
    ''
  ).trim();
}

function requireSharedSecret(req, expectedSecretValue) {
  const expected = String(expectedSecretValue || '').trim();
  if (!expected) return true;
  const provided = readBearerToken(req);
  return !!provided && provided === expected;
}

function sanitizeTelemetryCollection(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) return 'usage_history';
  return ALLOWED_TELEMETRY_COLLECTIONS.has(normalized) ? normalized : 'usage_history';
}

async function readPolicyDoc(collectionName, docId) {
  if (!docId) return {};
  const snap = await db.collection(collectionName).doc(docId).get();
  return snap.exists ? (snap.data() || {}) : {};
}

function mergePolicy(basePolicy, overridePolicy) {
  return {
    ...(basePolicy || {}),
    ...(overridePolicy || {})
  };
}

async function resolvePolicy(identity = {}) {
  const email = normalizeEmail(identity.email || '');
  const username = normalizeUsername(identity.username || '');
  const userId = String(identity.user_id || identity.userId || '').trim().toLowerCase();

  let merged = await readPolicyDoc('emmaneigh_policy', 'global');

  if (email) {
    merged = mergePolicy(merged, await readPolicyDoc('emmaneigh_policy_overrides', `email:${normalizeDocId(email)}`));
  }
  if (username) {
    merged = mergePolicy(merged, await readPolicyDoc('emmaneigh_policy_overrides', `username:${normalizeDocId(username)}`));
  }
  if (userId) {
    merged = mergePolicy(merged, await readPolicyDoc('emmaneigh_policy_overrides', `userid:${normalizeDocId(userId)}`));
  }

  return merged;
}

function jsonResponse(res, statusCode, payload) {
  res.status(statusCode).set('Cache-Control', 'no-store').json(payload);
}

function getManagedProvider() {
  return DEFAULT_MANAGED_AI_PROVIDER;
}

function getManagedModel(requestedModel) {
  const overrideAllowed = String(process.env.EMMANEIGH_MANAGED_AI_ALLOW_MODEL_OVERRIDE || '').trim().toLowerCase() === 'true';
  const requested = String(requestedModel || '').trim();
  if (overrideAllowed && requested) return requested;
  return DEFAULT_MANAGED_AI_MODEL;
}

function buildManagedAuthHeaders(apiKey) {
  return {
    'Content-Type': 'application/json',
    Authorization: `Bearer ${apiKey}`
  };
}

async function callGroq(pathname, payload, apiKey) {
  const response = await fetch(`${GROQ_OPENAI_BASE_URL}${pathname}`, {
    method: 'POST',
    headers: buildManagedAuthHeaders(apiKey),
    body: JSON.stringify(payload)
  });
  const rawText = await response.text();
  let parsed = null;
  try {
    parsed = rawText ? JSON.parse(rawText) : null;
  } catch (_) {
    parsed = null;
  }
  return {
    statusCode: response.status,
    ok: response.ok,
    rawText,
    json: parsed
  };
}

async function fetchGroqModels(apiKey) {
  const response = await fetch(`${GROQ_OPENAI_BASE_URL}/models`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${apiKey}`
    }
  });
  const rawText = await response.text();
  let parsed = null;
  try {
    parsed = rawText ? JSON.parse(rawText) : null;
  } catch (_) {
    parsed = null;
  }
  return {
    statusCode: response.status,
    ok: response.ok,
    rawText,
    json: parsed
  };
}

function extractProviderError(result) {
  if (!result) return 'Unknown provider error';
  if (result.json && typeof result.json === 'object') {
    const error = result.json.error;
    if (error && typeof error === 'object' && error.message) return String(error.message);
    if (result.json.message) return String(result.json.message);
  }
  return String(result.rawText || '').trim().slice(0, 400) || 'Unknown provider error';
}

function anthropicToolsToOpenAI(tools) {
  if (!Array.isArray(tools)) return [];
  return tools
    .filter(tool => tool && tool.name)
    .map(tool => ({
      type: 'function',
      function: {
        name: String(tool.name || '').trim(),
        description: String(tool.description || '').trim(),
        parameters: tool.input_schema && typeof tool.input_schema === 'object'
          ? tool.input_schema
          : { type: 'object', properties: {} }
      }
    }));
}

function textFromAnthropicBlocks(blocks) {
  if (!Array.isArray(blocks)) return '';
  return blocks
    .filter(block => block && block.type === 'text' && block.text)
    .map(block => String(block.text || '').trim())
    .filter(Boolean)
    .join('\n');
}

function anthropicMessagesToOpenAI(systemPrompt, messages) {
  const openAiMessages = [];
  const system = String(systemPrompt || '').trim();
  if (system) {
    openAiMessages.push({ role: 'system', content: system });
  }

  const items = Array.isArray(messages) ? messages : [];
  for (const message of items) {
    const role = String(message && message.role ? message.role : 'user').trim() || 'user';
    const content = message ? message.content : '';

    if (typeof content === 'string') {
      const text = content.trim();
      if (text) openAiMessages.push({ role, content: text });
      continue;
    }

    if (!Array.isArray(content)) continue;

    const text = textFromAnthropicBlocks(content);
    const toolUses = content.filter(block => block && block.type === 'tool_use' && block.name);
    const toolResults = content.filter(block => block && block.type === 'tool_result');

    if (role === 'assistant') {
      if (text || toolUses.length > 0) {
        const assistantMessage = {
          role: 'assistant',
          content: text || ''
        };
        if (toolUses.length > 0) {
          assistantMessage.tool_calls = toolUses.map((toolUse, index) => ({
            id: String(toolUse.id || `tool_${Date.now()}_${index}`),
            type: 'function',
            function: {
              name: String(toolUse.name || '').trim(),
              arguments: JSON.stringify(toolUse.input && typeof toolUse.input === 'object' ? toolUse.input : {})
            }
          }));
        }
        openAiMessages.push(assistantMessage);
      }
      continue;
    }

    if (text) {
      openAiMessages.push({ role: 'user', content: text });
    }

    for (const toolResult of toolResults) {
      const resultContent = typeof toolResult.content === 'string'
        ? toolResult.content
        : JSON.stringify(toolResult.content || {});
      openAiMessages.push({
        role: 'tool',
        tool_call_id: String(toolResult.tool_use_id || toolResult.id || ''),
        content: resultContent
      });
    }
  }

  return openAiMessages;
}

function extractOpenAITextContent(content) {
  if (typeof content === 'string') return content.trim();
  if (!Array.isArray(content)) return '';
  return content
    .map(item => {
      if (typeof item === 'string') return item;
      if (item && typeof item === 'object' && typeof item.text === 'string') return item.text;
      return '';
    })
    .join('\n')
    .trim();
}

function openAIMessageToAnthropicContent(message) {
  const content = [];
  if (!message || typeof message !== 'object') return content;

  const text = extractOpenAITextContent(message.content);
  if (text) {
    content.push({ type: 'text', text });
  }

  const toolCalls = Array.isArray(message.tool_calls) ? message.tool_calls : [];
  for (const toolCall of toolCalls) {
    const functionData = toolCall && toolCall.function ? toolCall.function : {};
    let parsedArgs = {};
    try {
      parsedArgs = functionData.arguments ? JSON.parse(functionData.arguments) : {};
    } catch (_) {
      parsedArgs = {};
    }
    content.push({
      type: 'tool_use',
      id: String(toolCall.id || ''),
      name: String(functionData.name || '').trim(),
      input: parsedArgs && typeof parsedArgs === 'object' ? parsedArgs : {}
    });
  }

  if (content.length === 0) {
    content.push({ type: 'text', text: '' });
  }
  return content;
}

async function callManagedChatCompletion({ system, messages, tools, maxTokens, requestedModel, apiKey }) {
  const provider = getManagedProvider();
  const model = getManagedModel(requestedModel);

  if (provider !== 'groq') {
    throw new Error(`Unsupported managed AI provider: ${provider}`);
  }

  const payload = {
    model,
    temperature: 0,
    max_tokens: Number(maxTokens) > 0 ? Number(maxTokens) : 1200,
    messages: anthropicMessagesToOpenAI(system, messages)
  };

  const openAiTools = anthropicToolsToOpenAI(tools);
  if (openAiTools.length > 0) {
    payload.tools = openAiTools;
    payload.tool_choice = 'auto';
  }

  const result = await callGroq('/chat/completions', payload, apiKey);
  if (!result.ok) {
    throw new Error(extractProviderError(result));
  }

  const choice = Array.isArray(result.json && result.json.choices) ? result.json.choices[0] : null;
  const message = choice && choice.message ? choice.message : null;
  if (!message) {
    throw new Error('Managed AI response was missing a message.');
  }

  return {
    provider,
    model,
    content: openAIMessageToAnthropicContent(message)
  };
}

async function callManagedPrompt({ prompt, system, maxTokens, requestedModel, apiKey }) {
  const result = await callManagedChatCompletion({
    system,
    messages: [{ role: 'user', content: String(prompt || '') }],
    tools: [],
    maxTokens,
    requestedModel,
    apiKey
  });

  const text = textFromAnthropicBlocks(result.content);
  if (!text) {
    throw new Error('Managed AI response was empty.');
  }

  return {
    provider: result.provider,
    model: result.model,
    text
  };
}

exports.telemetryIngest = onRequest(
  {
    region: 'us-central1',
    cors: true,
    secrets: [TELEMETRY_TOKEN]
  },
  async (req, res) => {
    if (req.method === 'OPTIONS') {
      res.status(204).send('');
      return;
    }
    if (req.method !== 'POST') {
      jsonResponse(res, 405, { success: false, error: 'Method not allowed' });
      return;
    }
    if (!requireSharedSecret(req, TELEMETRY_TOKEN.value())) {
      jsonResponse(res, 401, { success: false, error: 'Unauthorized' });
      return;
    }

    const body = parseJsonBody(req);
    const collection = sanitizeTelemetryCollection(body.collection);
    const payload = body.payload && typeof body.payload === 'object' ? body.payload : {};

    try {
      await db.collection('emmaneigh_events').add({
        collection,
        event_type: String(body.event_type || body.eventType || 'event').trim(),
        app_version: String(body.app_version || '').trim(),
        machine_id: String(body.machine_id || '').trim(),
        client_timestamp: String(body.timestamp || '').trim(),
        payload,
        received_at: admin.firestore.FieldValue.serverTimestamp(),
        user_agent: String(req.get('user-agent') || '').trim()
      });

      jsonResponse(res, 200, { success: true });
    } catch (error) {
      logger.error('Telemetry ingest failed', error);
      jsonResponse(res, 500, { success: false, error: 'Telemetry ingest failed' });
    }
  }
);

exports.accessPolicy = onRequest(
  {
    region: 'us-central1',
    cors: true,
    secrets: [ACCESS_POLICY_TOKEN]
  },
  async (req, res) => {
    if (req.method === 'OPTIONS') {
      res.status(204).send('');
      return;
    }
    if (req.method !== 'GET' && req.method !== 'POST') {
      jsonResponse(res, 405, { success: false, error: 'Method not allowed' });
      return;
    }
    if (!requireSharedSecret(req, ACCESS_POLICY_TOKEN.value())) {
      jsonResponse(res, 401, { success: false, error: 'Unauthorized' });
      return;
    }

    const body = parseJsonBody(req);
    const identity = {
      email: req.query.email || body.email || '',
      username: req.query.username || body.username || '',
      user_id: req.query.user_id || body.user_id || body.userId || ''
    };

    try {
      const policy = await resolvePolicy(identity);
      jsonResponse(res, 200, {
        policy,
        generated_at: new Date().toISOString()
      });
    } catch (error) {
      logger.error('Access policy lookup failed', error);
      jsonResponse(res, 500, { success: false, error: 'Policy lookup failed' });
    }
  }
);

exports.managedAiSession = onRequest(
  {
    region: 'us-central1',
    cors: true,
    secrets: [SESSION_SIGNING_KEY]
  },
  async (req, res) => {
    if (req.method === 'OPTIONS') {
      res.status(204).send('');
      return;
    }
    if (req.method !== 'POST') {
      jsonResponse(res, 405, { success: false, error: 'Method not allowed' });
      return;
    }

    const body = parseJsonBody(req);
    const identity = buildManagedAiSessionIdentity(req);
    const email = normalizeEmail(identity.email || '');
    if (!email) {
      jsonResponse(res, 400, { success: false, error: 'A valid email is required to start a managed AI session.' });
      return;
    }

    try {
      const policy = await resolvePolicy(identity);
      const decision = evaluateAccessPolicyDecision(policy, identity);
      if (!decision.allowed) {
        jsonResponse(res, 403, {
          success: false,
          error: decision.message || 'Access denied by policy.',
          code: 'access_denied'
        });
        return;
      }
      if (!isManagedAiAllowedByPolicy(policy)) {
        jsonResponse(res, 403, {
          success: false,
          error: 'Managed AI is not enabled for this user by policy.',
          code: 'provider_not_allowed'
        });
        return;
      }

      const session = issueManagedAiSessionToken(identity, {
        isAdmin: !!decision.isAdmin,
        requestedModel: body.model || req.query.model || ''
      }, SESSION_SIGNING_KEY.value());

      jsonResponse(res, 200, {
        success: true,
        token: session.token,
        expires_at: session.expiresAt,
        expires_in_seconds: session.ttlSeconds,
        provider: 'managed',
        backend_provider: getManagedProvider(),
        model: session.payload.managed_model,
        is_admin: !!decision.isAdmin
      });
    } catch (error) {
      logger.error('Managed AI session issuance failed', error);
      jsonResponse(res, 500, { success: false, error: 'Managed AI session issuance failed' });
    }
  }
);

exports.managedAiHealth = onRequest(
  {
    region: 'us-central1',
    cors: true,
    secrets: [SESSION_SIGNING_KEY, GROQ_API_KEY]
  },
  async (req, res) => {
    if (req.method === 'OPTIONS') {
      res.status(204).send('');
      return;
    }
    if (req.method !== 'GET' && req.method !== 'POST') {
      jsonResponse(res, 405, { success: false, error: 'Method not allowed' });
      return;
    }
    const auth = verifyManagedAiSessionToken(readBearerToken(req), SESSION_SIGNING_KEY.value());
    if (!auth.ok) {
      jsonResponse(res, auth.statusCode || 401, buildManagedAiUnauthorizedPayload(auth));
      return;
    }

    try {
      const provider = getManagedProvider();
      const model = getManagedModel(req.query.model || parseJsonBody(req).model || '');
      const groqResult = await fetchGroqModels(GROQ_API_KEY.value());
      if (!groqResult.ok) {
        jsonResponse(res, 502, {
          success: false,
          provider,
          model,
          error: extractProviderError(groqResult)
        });
        return;
      }

      const models = Array.isArray(groqResult.json && groqResult.json.data)
        ? groqResult.json.data.map(item => String(item && item.id || '').trim()).filter(Boolean)
        : [];

      jsonResponse(res, 200, {
        success: true,
        provider,
        model,
        model_available: models.includes(model),
        models
      });
    } catch (error) {
      logger.error('Managed AI health check failed', error);
      jsonResponse(res, 500, { success: false, error: 'Managed AI health check failed' });
    }
  }
);

exports.managedAiPrompt = onRequest(
  {
    region: 'us-central1',
    cors: true,
    secrets: [SESSION_SIGNING_KEY, GROQ_API_KEY]
  },
  async (req, res) => {
    if (req.method === 'OPTIONS') {
      res.status(204).send('');
      return;
    }
    if (req.method !== 'POST') {
      jsonResponse(res, 405, { success: false, error: 'Method not allowed' });
      return;
    }
    const auth = verifyManagedAiSessionToken(readBearerToken(req), SESSION_SIGNING_KEY.value());
    if (!auth.ok) {
      jsonResponse(res, auth.statusCode || 401, buildManagedAiUnauthorizedPayload(auth));
      return;
    }

    const body = parseJsonBody(req);
    const prompt = String(body.prompt || '').trim();
    if (!prompt) {
      jsonResponse(res, 400, { success: false, error: 'Prompt is required' });
      return;
    }

    try {
      const result = await callManagedPrompt({
        prompt,
        system: String(body.system || '').trim(),
        maxTokens: body.max_tokens || body.maxTokens || 1400,
        requestedModel: body.model,
        apiKey: GROQ_API_KEY.value()
      });
      jsonResponse(res, 200, {
        success: true,
        provider: result.provider,
        model: result.model,
        text: result.text
      });
    } catch (error) {
      logger.error('Managed AI prompt failed', error);
      jsonResponse(res, 502, { success: false, error: String(error.message || error) });
    }
  }
);

exports.managedAiAgent = onRequest(
  {
    region: 'us-central1',
    cors: true,
    secrets: [SESSION_SIGNING_KEY, GROQ_API_KEY]
  },
  async (req, res) => {
    if (req.method === 'OPTIONS') {
      res.status(204).send('');
      return;
    }
    if (req.method !== 'POST') {
      jsonResponse(res, 405, { success: false, error: 'Method not allowed' });
      return;
    }
    const auth = verifyManagedAiSessionToken(readBearerToken(req), SESSION_SIGNING_KEY.value());
    if (!auth.ok) {
      jsonResponse(res, auth.statusCode || 401, buildManagedAiUnauthorizedPayload(auth));
      return;
    }

    const body = parseJsonBody(req);
    try {
      const result = await callManagedChatCompletion({
        system: body.system,
        messages: Array.isArray(body.messages) ? body.messages : [],
        tools: Array.isArray(body.tools) ? body.tools : [],
        maxTokens: body.max_tokens || body.maxTokens || 1200,
        requestedModel: body.model,
        apiKey: GROQ_API_KEY.value()
      });
      jsonResponse(res, 200, {
        model: result.model,
        provider: result.provider,
        content: result.content
      });
    } catch (error) {
      logger.error('Managed AI agent call failed', error);
      jsonResponse(res, 502, { success: false, error: String(error.message || error) });
    }
  }
);
