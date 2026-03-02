/**
 * Firebase Firestore Configuration for EmmaNeigh
 *
 * Centralizes user activity and usage logging to Firebase Firestore.
 * To set up:
 *   1. Go to https://console.firebase.google.com/
 *   2. Create a new project (e.g., "emmaneigh-analytics")
 *   3. Enable Firestore Database (start in test mode or configure rules)
 *   4. Go to Project Settings → General → Your apps → Add web app
 *   5. Copy the firebaseConfig object and paste it below
 */

const { initializeApp } = require('firebase/app');
const { getFirestore, collection, addDoc, writeBatch, doc } = require('firebase/firestore');

// ============================================================
// PASTE YOUR FIREBASE CONFIG HERE
// ============================================================
const firebaseConfig = {
  apiKey: "AIzaSyAXl5s1asHqXpuVi3K_JKGsMxRuU7frJFE",
  authDomain: "emmaneigh-7f845.firebaseapp.com",
  projectId: "emmaneigh-7f845",
  storageBucket: "emmaneigh-7f845.firebasestorage.app",
  messagingSenderId: "383420892084",
  appId: "1:383420892084:web:cd78a7a242c2c514b9a8e1",
  measurementId: "G-DZLTQM7B23"
};

// ============================================================

let firebaseApp = null;
let firestoreDb = null;
let firebaseReady = false;

/**
 * Initialize Firebase. Returns true if successful, false if config is missing.
 */
function initFirebase() {
  if (firebaseReady) return true;

  // Check if config has been filled in
  if (!firebaseConfig.apiKey || !firebaseConfig.projectId) {
    console.warn('Firebase config not set — activity logging will be local-only.');
    console.warn('Edit desktop-app/firebase-config.js to enable centralized logging.');
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
  logToFirestore,
  batchLogToFirestore
};
