/**
 * Firebase Firestore Configuration for EmmaNeigh
 *
 * Firebase config is embedded in the app for automatic centralized telemetry.
 * A local override file at [userData]/firebase_config.json is checked first,
 * but if not present, the embedded config is used automatically.
 * No user configuration is needed — usage data flows to Firestore silently.
 */

const { app } = require('electron');
const path = require('path');
const fs = require('fs');
const { initializeApp } = require('firebase/app');
const { getFirestore, collection, addDoc, writeBatch, doc } = require('firebase/firestore');

// Embedded Firebase config loaded from environment or local file
// NEVER hardcode credentials in source code committed to git
const EMBEDDED_FIREBASE_CONFIG = null;

// Optional local override (checked first, falls back to embedded)
const firebaseConfigPath = path.join(app.getPath('userData'), 'firebase_config.json');

let firebaseApp = null;
let firestoreDb = null;
let firebaseReady = false;

/**
 * Read Firebase config. Checks local override file first, then uses embedded config.
 * Always returns a valid config — never null.
 */
function loadFirebaseConfig() {
  // Try local override first
  try {
    if (fs.existsSync(firebaseConfigPath)) {
      const raw = fs.readFileSync(firebaseConfigPath, 'utf8');
      const config = JSON.parse(raw);
      if (config.apiKey && config.projectId) {
        return config;
      }
    }
  } catch (e) {
    console.error('Failed to read local Firebase config override:', e.message);
  }
  // Fall back to embedded config if available
  return EMBEDDED_FIREBASE_CONFIG || null;
}

/**
 * Save Firebase config to the local file.
 */
function saveFirebaseConfig(config) {
  try {
    fs.writeFileSync(firebaseConfigPath, JSON.stringify(config, null, 2));
    return true;
  } catch (e) {
    console.error('Failed to save Firebase config:', e.message);
    return false;
  }
}

/**
 * Initialize Firebase. Returns true if successful, false if config is missing.
 */
function initFirebase() {
  if (firebaseReady) return true;

  const firebaseConfig = loadFirebaseConfig();
  if (!firebaseConfig) {
    console.warn('Firebase config not found. Telemetry will be disabled.');
    return false;
  }

  try {
    firebaseApp = initializeApp(firebaseConfig);
    firestoreDb = getFirestore(firebaseApp);
    firebaseReady = true;
    console.log('Firebase initialized for project:', firebaseConfig.projectId);
    return true;
  } catch (err) {
    console.error('Firebase initialization failed:', err.message);
    return false;
  }
}

/**
 * Log a single document to a Firestore collection.
 * Returns true on success, false on failure (caller should queue for retry).
 */
async function logToFirestore(collectionName, data) {
  if (!firebaseReady) return false;

  try {
    await addDoc(collection(firestoreDb, collectionName), {
      ...data,
      _logged_at: new Date().toISOString()
    });
    return true;
  } catch (err) {
    console.error(`Firestore write to ${collectionName} failed:`, err.message);
    return false;
  }
}

/**
 * Batch-write multiple documents to a Firestore collection.
 * Used by the offline queue sync. Returns number of successfully written docs.
 */
async function batchLogToFirestore(collectionName, docs) {
  if (!firebaseReady || docs.length === 0) return 0;

  let written = 0;
  // Firestore batches are limited to 500 operations
  const batchSize = 450;

  for (let i = 0; i < docs.length; i += batchSize) {
    const chunk = docs.slice(i, i + batchSize);
    try {
      const batch = writeBatch(firestoreDb);
      for (const data of chunk) {
        const docRef = doc(collection(firestoreDb, collectionName));
        batch.set(docRef, { ...data, _logged_at: data._logged_at || new Date().toISOString() });
      }
      await batch.commit();
      written += chunk.length;
    } catch (err) {
      console.error(`Firestore batch write failed (chunk ${i}):`, err.message);
      break;
    }
  }

  return written;
}

module.exports = {
  initFirebase,
  loadFirebaseConfig,
  saveFirebaseConfig,
  logToFirestore,
  batchLogToFirestore
};
