const { onRequest } = require('firebase-functions/v2/https');
const { defineSecret } = require('firebase-functions/params');
const logger = require('firebase-functions/logger');
const admin = require('firebase-admin');

admin.initializeApp();
const db = admin.firestore();

const TELEMETRY_TOKEN = defineSecret('EMMANEIGH_TELEMETRY_TOKEN');
const ACCESS_POLICY_TOKEN = defineSecret('EMMANEIGH_ACCESS_POLICY_TOKEN');

const ALLOWED_TELEMETRY_COLLECTIONS = new Set([
  'telemetry_health',
  'usage_history',
  'user_feedback',
  'prompt_logs',
  'user_activity_history'
]);

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

function readBearerToken(req) {
  const authHeader = String(req.get('authorization') || '').trim();
  if (authHeader.toLowerCase().startsWith('bearer ')) {
    return authHeader.slice(7).trim();
  }
  return String(req.get('x-emmaneigh-telemetry-key') || req.get('x-emmaneigh-access-key') || '').trim();
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
      res.status(405).json({ success: false, error: 'Method not allowed' });
      return;
    }
    if (!requireSharedSecret(req, TELEMETRY_TOKEN.value())) {
      res.status(401).json({ success: false, error: 'Unauthorized' });
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

      res.json({ success: true });
    } catch (error) {
      logger.error('Telemetry ingest failed', error);
      res.status(500).json({ success: false, error: 'Telemetry ingest failed' });
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
      res.status(405).json({ success: false, error: 'Method not allowed' });
      return;
    }
    if (!requireSharedSecret(req, ACCESS_POLICY_TOKEN.value())) {
      res.status(401).json({ success: false, error: 'Unauthorized' });
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
      res.json({
        policy,
        generated_at: new Date().toISOString()
      });
    } catch (error) {
      logger.error('Access policy lookup failed', error);
      res.status(500).json({ success: false, error: 'Policy lookup failed' });
    }
  }
);
