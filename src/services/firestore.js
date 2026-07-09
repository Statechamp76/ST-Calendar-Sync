const admin = require('firebase-admin');

let initialized = false;
let db = null;

function getFirestore() {
  if (db) {
    return db;
  }

  if (!initialized) {
    const projectId = String(process.env.FIRESTORE_PROJECT_ID || '').trim();
    const options = projectId ? { projectId } : undefined;
    admin.initializeApp(options);
    initialized = true;
  }

  db = admin.firestore();
  return db;
}

function serverTimestamp() {
  return admin.firestore.FieldValue.serverTimestamp();
}

module.exports = {
  getFirestore,
  serverTimestamp,
};
