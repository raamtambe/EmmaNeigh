const GROQ_OPENAI_BASE_URL = 'https://api.groq.com/openai/v1';
const MANAGED_AI_SESSION_ISSUER = 'emmaneigh-managed-ai';
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function jsonResponse(payload, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-EmmaNeigh-Version, X-EmmaNeigh-Machine',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      ...extraHeaders
    }
  });
}

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeUsername(value) {
  return String(value || '').trim().toLowerCase();
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

function splitCsv(value) {
  return String(value || '')
    .split(/[,\n;]+/)
    .map((entry) => String(entry || '').trim().toLowerCase())
    .filter(Boolean);
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

function getManagedProvider(env) {
  return String(env.MANAGED_AI_PROVIDER || 'groq').trim().toLowerCase() || 'groq';
}

function getManagedModel(env, requestedModel = '') {
  const overrideAllowed = normalizeBoolean(env.MANAGED_AI_ALLOW_MODEL_OVERRIDE, false);
  const requested = String(requestedModel || '').trim();
  if (overrideAllowed && requested) return requested;
  return String(env.MANAGED_AI_MODEL || 'qwen/qwen3-32b').trim() || 'qwen/qwen3-32b';
}

function getManagedAiSessionTtlSeconds(env) {
  const parsed = Number.parseInt(env.MANAGED_AI_SESSION_TTL_SECONDS || '3600', 10);
  if (!Number.isFinite(parsed)) return 3600;
  return Math.max(300, Math.min(parsed, 86400));
}

async function parseJsonBody(request) {
  const contentType = String(request.headers.get('content-type') || '').toLowerCase();
  if (!contentType.includes('application/json')) return {};
  try {
    return await request.json();
  } catch (_) {
    return {};
  }
}

function readBearerToken(request) {
  const authHeader = String(request.headers.get('authorization') || '').trim();
  if (authHeader.toLowerCase().startsWith('bearer ')) {
    return authHeader.slice(7).trim();
  }
  return '';
}

function getAccessPolicy(env) {
  return {
    allowedEmails: splitCsv(env.ALLOWED_EMAILS),
    allowedDomains: splitCsv(env.ALLOWED_DOMAINS),
    allowedUsers: splitCsv(env.ALLOWED_USERS),
    blockedEmails: splitCsv(env.BLOCKED_EMAILS),
    blockedDomains: splitCsv(env.BLOCKED_DOMAINS),
    blockedUsers: splitCsv(env.BLOCKED_USERS),
    adminEmails: splitCsv(env.ADMIN_EMAILS),
    adminDomains: splitCsv(env.ADMIN_DOMAINS),
    adminUsers: splitCsv(env.ADMIN_USERS),
    allowedAiProviders: splitCsv(env.ALLOWED_AI_PROVIDERS),
    blockAll: normalizeBoolean(env.BLOCK_ALL, false),
    message: String(env.ACCESS_POLICY_MESSAGE || 'Access to EmmaNeigh is managed by your administrator.').trim()
  };
}

function evaluatePolicy(identity, env) {
  const policy = getAccessPolicy(env);
  const email = normalizeEmail(identity.email || '');
  const username = normalizeUsername(identity.username || '');
  const userId = String(identity.user_id || identity.userId || '').trim().toLowerCase();
  const emailDomain = email.includes('@') ? email.split('@')[1] : '';
  const emailLocalPart = email.includes('@') ? email.split('@')[0] : '';

  const isAdmin = (
    (!!email && (listHasMatch(policy.adminEmails, email) || listHasMatch(policy.adminEmails, emailLocalPart))) ||
    (!!emailDomain && listHasMatch(policy.adminDomains, emailDomain)) ||
    (!!username && listHasMatch(policy.adminUsers, username)) ||
    (!!userId && listHasMatch(policy.adminUsers, userId))
  );

  if (policy.blockAll) {
    return { allowed: false, isAdmin, message: policy.message };
  }
  if (email && listHasMatch(policy.blockedEmails, email)) return { allowed: false, isAdmin, message: policy.message };
  if (emailDomain && listHasMatch(policy.blockedDomains, emailDomain)) return { allowed: false, isAdmin, message: policy.message };
  if ((username && listHasMatch(policy.blockedUsers, username)) || (userId && listHasMatch(policy.blockedUsers, userId))) {
    return { allowed: false, isAdmin, message: policy.message };
  }

  const hasAllowRules = policy.allowedEmails.length > 0 || policy.allowedDomains.length > 0 || policy.allowedUsers.length > 0;
  if (hasAllowRules) {
    const emailAllowed = email && (listHasMatch(policy.allowedEmails, email) || listHasMatch(policy.allowedEmails, emailLocalPart));
    const domainAllowed = emailDomain && listHasMatch(policy.allowedDomains, emailDomain);
    const userAllowed = (username && listHasMatch(policy.allowedUsers, username)) || (userId && listHasMatch(policy.allowedUsers, userId));
    if (!(emailAllowed || domainAllowed || userAllowed)) {
      return { allowed: false, isAdmin, message: policy.message };
    }
  }

  if (policy.allowedAiProviders.length > 0) {
    const backendProvider = getManagedProvider(env);
    const aiAllowed = (
      listHasMatch(policy.allowedAiProviders, 'managed') ||
      listHasMatch(policy.allowedAiProviders, backendProvider) ||
      listHasMatch(policy.allowedAiProviders, '*') ||
      listHasMatch(policy.allowedAiProviders, 'all')
    );
    if (!aiAllowed) {
      return { allowed: false, isAdmin, message: 'Managed AI is not enabled for this user by policy.' };
    }
  }

  return { allowed: true, isAdmin, message: 'Access granted.' };
}

function base64UrlEncodeBytes(bytes) {
  let binary = '';
  const uint8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  for (const byte of uint8) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function base64UrlEncodeJson(value) {
  return base64UrlEncodeBytes(encoder.encode(JSON.stringify(value)));
}

function base64UrlDecodeToBytes(value) {
  const normalized = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function decodeBase64UrlJson(value) {
  return JSON.parse(decoder.decode(base64UrlDecodeToBytes(value)));
}

async function signHmac(value, secret) {
  const key = await crypto.subtle.importKey(
    'raw',
    encoder.encode(String(secret || '')),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const signature = await crypto.subtle.sign('HMAC', key, encoder.encode(String(value || '')));
  return base64UrlEncodeBytes(new Uint8Array(signature));
}

function timingSafeEquals(a, b) {
  const left = encoder.encode(String(a || ''));
  const right = encoder.encode(String(b || ''));
  if (left.length !== right.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i += 1) {
    diff |= left[i] ^ right[i];
  }
  return diff === 0;
}

async function issueManagedAiSessionToken(identity, options, env) {
  const nowSeconds = Math.floor(Date.now() / 1000);
  const ttlSeconds = getManagedAiSessionTtlSeconds(env);
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
    managed_provider: getManagedProvider(env),
    managed_model: getManagedModel(env, options.requestedModel || '')
  };

  const headerPart = base64UrlEncodeJson({ alg: 'HS256', typ: 'JWT' });
  const payloadPart = base64UrlEncodeJson(payload);
  const signature = await signHmac(`${headerPart}.${payloadPart}`, env.SESSION_SIGNING_KEY);

  return {
    token: `${headerPart}.${payloadPart}.${signature}`,
    payload,
    ttlSeconds,
    expiresAt: new Date(payload.exp * 1000).toISOString()
  };
}

async function verifyManagedAiSessionToken(token, env) {
  const rawToken = String(token || '').trim();
  if (!rawToken) {
    return { ok: false, statusCode: 401, code: 'missing_token', error: 'Managed AI session token is required.' };
  }
  if (!env.SESSION_SIGNING_KEY) {
    return { ok: false, statusCode: 500, code: 'missing_signing_key', error: 'SESSION_SIGNING_KEY is not configured.' };
  }

  const parts = rawToken.split('.');
  if (parts.length !== 3) {
    return { ok: false, statusCode: 401, code: 'invalid_token', error: 'Managed AI session token is malformed.' };
  }

  const [headerPart, payloadPart, signaturePart] = parts;
  const expectedSignature = await signHmac(`${headerPart}.${payloadPart}`, env.SESSION_SIGNING_KEY);
  if (!timingSafeEquals(signaturePart, expectedSignature)) {
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

function buildUnauthorizedPayload(result) {
  return {
    success: false,
    error: result.error || 'Unauthorized',
    code: result.code || 'unauthorized',
    session_expired: result.code === 'session_expired'
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

async function callGroq(pathname, payload, apiKey) {
  const response = await fetch(`${GROQ_OPENAI_BASE_URL}${pathname}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`
    },
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

function anthropicToolsToOpenAI(tools) {
  if (!Array.isArray(tools)) return [];
  return tools
    .filter((tool) => tool && tool.name)
    .map((tool) => ({
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
    .filter((block) => block && block.type === 'text' && block.text)
    .map((block) => String(block.text || '').trim())
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
    const toolUses = content.filter((block) => block && block.type === 'tool_use' && block.name);
    const toolResults = content.filter((block) => block && block.type === 'tool_result');

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
    .map((item) => {
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

async function callManagedChatCompletion({ system, messages, tools, maxTokens, requestedModel, env }) {
  const provider = getManagedProvider(env);
  const model = getManagedModel(env, requestedModel);

  if (provider !== 'groq') {
    throw new Error(`Unsupported managed AI provider: ${provider}`);
  }
  if (!env.GROQ_API_KEY) {
    throw new Error('GROQ_API_KEY is not configured.');
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

  const result = await callGroq('/chat/completions', payload, env.GROQ_API_KEY);
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

async function callManagedPrompt({ prompt, system, maxTokens, requestedModel, env }) {
  const result = await callManagedChatCompletion({
    system,
    messages: [{ role: 'user', content: String(prompt || '') }],
    tools: [],
    maxTokens,
    requestedModel,
    env
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

async function handleSession(request, env) {
  if (request.method !== 'POST') {
    return jsonResponse({ success: false, error: 'Method not allowed' }, 405);
  }

  const body = await parseJsonBody(request);
  const identity = {
    email: body.email || '',
    username: body.username || '',
    user_id: body.user_id || body.userId || '',
    display_name: body.display_name || body.displayName || '',
    machine_id: body.machine_id || body.machineId || '',
    app_version: body.app_version || body.appVersion || ''
  };

  const email = normalizeEmail(identity.email || '');
  if (!email) {
    return jsonResponse({ success: false, error: 'A valid email is required to start a managed AI session.' }, 400);
  }

  const decision = evaluatePolicy(identity, env);
  if (!decision.allowed) {
    return jsonResponse({
      success: false,
      error: decision.message || 'Access denied by policy.',
      code: 'access_denied'
    }, 403);
  }

  const session = await issueManagedAiSessionToken(identity, {
    isAdmin: !!decision.isAdmin,
    requestedModel: body.model || ''
  }, env);

  return jsonResponse({
    success: true,
    token: session.token,
    expires_at: session.expiresAt,
    expires_in_seconds: session.ttlSeconds,
    provider: 'managed',
    backend_provider: getManagedProvider(env),
    model: session.payload.managed_model,
    is_admin: !!decision.isAdmin
  });
}

async function requireManagedSession(request, env) {
  const auth = await verifyManagedAiSessionToken(readBearerToken(request), env);
  if (!auth.ok) {
    return { response: jsonResponse(buildUnauthorizedPayload(auth), auth.statusCode || 401) };
  }
  return { payload: auth.payload };
}

async function handleHealth(request, env) {
  if (request.method !== 'GET' && request.method !== 'POST') {
    return jsonResponse({ success: false, error: 'Method not allowed' }, 405);
  }
  const auth = await requireManagedSession(request, env);
  if (auth.response) return auth.response;

  const body = request.method === 'POST' ? await parseJsonBody(request) : {};
  const model = getManagedModel(env, new URL(request.url).searchParams.get('model') || body.model || '');
  const provider = getManagedProvider(env);
  const groqResult = await fetchGroqModels(env.GROQ_API_KEY);
  if (!groqResult.ok) {
    return jsonResponse({
      success: false,
      provider,
      model,
      error: extractProviderError(groqResult)
    }, 502);
  }

  const models = Array.isArray(groqResult.json && groqResult.json.data)
    ? groqResult.json.data.map((item) => String(item && item.id || '').trim()).filter(Boolean)
    : [];

  return jsonResponse({
    success: true,
    provider,
    model,
    model_available: models.includes(model),
    models
  });
}

async function handlePrompt(request, env) {
  if (request.method !== 'POST') {
    return jsonResponse({ success: false, error: 'Method not allowed' }, 405);
  }
  const auth = await requireManagedSession(request, env);
  if (auth.response) return auth.response;

  const body = await parseJsonBody(request);
  const prompt = String(body.prompt || '').trim();
  if (!prompt) {
    return jsonResponse({ success: false, error: 'Prompt is required' }, 400);
  }

  try {
    const result = await callManagedPrompt({
      prompt,
      system: String(body.system || '').trim(),
      maxTokens: body.max_tokens || body.maxTokens || 1400,
      requestedModel: body.model,
      env
    });
    return jsonResponse({
      success: true,
      provider: result.provider,
      model: result.model,
      text: result.text
    });
  } catch (error) {
    return jsonResponse({ success: false, error: String(error.message || error) }, 502);
  }
}

async function handleAgent(request, env) {
  if (request.method !== 'POST') {
    return jsonResponse({ success: false, error: 'Method not allowed' }, 405);
  }
  const auth = await requireManagedSession(request, env);
  if (auth.response) return auth.response;

  const body = await parseJsonBody(request);
  try {
    const result = await callManagedChatCompletion({
      system: body.system,
      messages: Array.isArray(body.messages) ? body.messages : [],
      tools: Array.isArray(body.tools) ? body.tools : [],
      maxTokens: body.max_tokens || body.maxTokens || 1200,
      requestedModel: body.model,
      env
    });
    return jsonResponse({
      model: result.model,
      provider: result.provider,
      content: result.content
    });
  } catch (error) {
    return jsonResponse({ success: false, error: String(error.message || error) }, 502);
  }
}

function normalizePathname(pathname) {
  const normalized = String(pathname || '').replace(/\/+$/, '') || '/';
  return normalized;
}

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return jsonResponse({ success: true }, 204);
    }

    const url = new URL(request.url);
    const pathname = normalizePathname(url.pathname);

    if (pathname === '/' || pathname === '/healthz') {
      return jsonResponse({
        success: true,
        service: 'emmaneigh-managed-ai',
        provider: getManagedProvider(env),
        model: getManagedModel(env, ''),
        endpoints: ['/session', '/health', '/prompt', '/agent']
      });
    }

    if (pathname === '/session' || pathname === '/managed-ai/session') {
      return handleSession(request, env);
    }
    if (pathname === '/health' || pathname === '/managed-ai/health') {
      return handleHealth(request, env);
    }
    if (pathname === '/prompt' || pathname === '/managed-ai/prompt') {
      return handlePrompt(request, env);
    }
    if (pathname === '/agent' || pathname === '/managed-ai/agent') {
      return handleAgent(request, env);
    }

    return jsonResponse({ success: false, error: 'Not found' }, 404);
  }
};
