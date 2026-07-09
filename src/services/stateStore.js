const crypto = require('node:crypto');
const { DateTime } = require('luxon');
const sheets = require('./sheets');
const firestoreService = require('./firestore');
const { normalizeUpn } = require('../utils/upn');

const LOCKS_COLLECTION = 'locks';
const EVENT_MAP_COLLECTION = 'eventMap';
const DELTA_STATE_COLLECTION = 'deltaState';
const DEFAULT_LOCK_TTL_SECONDS = 600;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getRetryableCode(error) {
  if (error && typeof error.code === 'number') return error.code;
  const status = error?.response?.status;
  if (typeof status === 'number') return status;
  return null;
}

function isRetryableFirestoreError(error) {
  const code = getRetryableCode(error);
  if (code === 429) return true;
  if (code >= 500 && code <= 599) return true;

  const known = ['aborted', 'deadline-exceeded', 'unavailable', 'resource-exhausted', 'internal'];
  const msg = String(error?.message || '').toLowerCase();
  return known.some((k) => msg.includes(k));
}

async function withFirestoreRetry(fn, label) {
  const maxAttempts = 6;
  let lastError = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;
      const retryable = isRetryableFirestoreError(error);
      if (!retryable || attempt === maxAttempts) {
        throw new Error(`Firestore ${label} failed: ${error.message}`);
      }
      const base = Math.min(10_000, 250 * (2 ** (attempt - 1)));
      const jitter = Math.floor(Math.random() * 250);
      await sleep(base + jitter);
    }
  }
  throw lastError;
}

function isFirestoreEnabled() {
  return String(process.env.FIRESTORE_ENABLED || '').trim().toLowerCase() === 'true';
}

function hashStableKey(stableKey) {
  return crypto.createHash('sha256').update(String(stableKey || '')).digest('hex').slice(0, 16);
}

function normalizeMailboxKey(mailboxKey) {
  return normalizeUpn(mailboxKey);
}

function lockTtlSeconds() {
  const parsed = Number.parseInt(String(process.env.LOCK_TTL_SECONDS || DEFAULT_LOCK_TTL_SECONDS), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) return DEFAULT_LOCK_TTL_SECONDS;
  return parsed;
}

function mapFirestoreDoc(doc) {
  if (!doc || !doc.exists) return null;
  const data = doc.data() || {};
  return { id: doc.id, ...data };
}

async function acquireLock(mailboxKey, owner, ttlSeconds = lockTtlSeconds()) {
  const key = normalizeMailboxKey(mailboxKey);
  const holder = String(owner || '').trim() || 'unknown-holder';
  const ttl = Number.parseInt(String(ttlSeconds || ''), 10);

  if (!key) throw new Error('Missing mailboxKey');
  if (!Number.isFinite(ttl) || ttl <= 0) throw new Error('Invalid ttlSeconds');

  if (!isFirestoreEnabled()) {
    return sheets.tryAcquireLock(key, holder, ttl);
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    const lockRef = db.collection(LOCKS_COLLECTION).doc(key);
    const now = DateTime.utc();
    const expiresAt = now.plus({ seconds: ttl });

    const acquired = await db.runTransaction(async (tx) => {
      const snap = await tx.get(lockRef);
      const current = snap.exists ? (snap.data() || {}) : {};
      const currentExpiresUtc = String(current.expiresAtUtc || '').trim();
      const currentExpires = currentExpiresUtc ? DateTime.fromISO(currentExpiresUtc, { zone: 'utc' }) : null;
      if (currentExpires && currentExpires.isValid && currentExpires > now) {
        return false;
      }

      tx.set(lockRef, {
        owner: holder,
        acquiredAtUtc: now.toISO(),
        expiresAtUtc: expiresAt.toISO(),
      }, { merge: true });
      return true;
    });

    if (!acquired) return { acquired: false, reason: 'locked' };
    return { acquired: true };
  }, `acquire lock ${key}`);
}

async function releaseLock(mailboxKey, owner) {
  const key = normalizeMailboxKey(mailboxKey);
  const holder = String(owner || '').trim() || 'unknown-holder';

  if (!key) throw new Error('Missing mailboxKey');

  if (!isFirestoreEnabled()) {
    return sheets.releaseLock(key, holder);
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    const lockRef = db.collection(LOCKS_COLLECTION).doc(key);
    const now = DateTime.utc().toISO();

    const released = await db.runTransaction(async (tx) => {
      const snap = await tx.get(lockRef);
      if (!snap.exists) return { released: false, reason: 'missing' };
      const current = snap.data() || {};
      if (String(current.owner || '') !== holder) return { released: false, reason: 'not_holder' };
      tx.set(lockRef, { expiresAtUtc: now, acquiredAtUtc: now }, { merge: true });
      return { released: true };
    });

    return released;
  }, `release lock ${key}`);
}

async function getDeltaState(mailboxKey) {
  const key = normalizeMailboxKey(mailboxKey);
  if (!key) throw new Error('Missing mailboxKey');

  if (!isFirestoreEnabled()) {
    const s = await sheets.getDeltaState(key);
    return {
      mailboxKey: key,
      deltaToken: s.delta_link || null,
      lastDeltaSyncUtc: s.last_run_utc || null,
      rowIndex: s.rowIndex || null,
    };
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    const snap = await db.collection(DELTA_STATE_COLLECTION).doc(key).get();
    if (!snap.exists) {
      return {
        mailboxKey: key,
        deltaToken: null,
        lastDeltaSyncUtc: null,
        lastFullReconcileUtc: null,
      };
    }
    const data = snap.data() || {};
    return {
      mailboxKey: key,
      deltaToken: data.deltaToken || null,
      lastDeltaSyncUtc: data.lastDeltaSyncUtc || null,
      lastFullReconcileUtc: data.lastFullReconcileUtc || null,
    };
  }, `get delta state ${key}`);
}

async function setDeltaState(mailboxKey, deltaToken, existingRowIndex = null) {
  const key = normalizeMailboxKey(mailboxKey);
  if (!key) throw new Error('Missing mailboxKey');

  if (!isFirestoreEnabled()) {
    await sheets.updateDeltaState(key, deltaToken, existingRowIndex);
    return;
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    await db.collection(DELTA_STATE_COLLECTION).doc(key).set({
      deltaToken: deltaToken || null,
      lastDeltaSyncUtc: DateTime.utc().toISO(),
    }, { merge: true });
  }, `set delta state ${key}`);
}

async function setLastFullReconcileUtc(mailboxKey, valueIso = null) {
  const key = normalizeMailboxKey(mailboxKey);
  if (!key) throw new Error('Missing mailboxKey');
  const value = valueIso || DateTime.utc().toISO();

  if (!isFirestoreEnabled()) {
    return;
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    await db.collection(DELTA_STATE_COLLECTION).doc(key).set({
      lastFullReconcileUtc: value,
    }, { merge: true });
  }, `set lastFullReconcileUtc ${key}`);
}

function makeEventMapDocId(stableKey, mailboxKey = null) {
  const base = String(stableKey || '').trim();
  if (!base) return '';
  const mailbox = mailboxKey ? normalizeMailboxKey(mailboxKey) : '';
  if (!mailbox) return base;
  return `${mailbox}::${base}`;
}

function toLegacyLikeMapping(data) {
  if (!data) return null;
  return {
    mailbox_key: data.mailboxKey || '',
    outlook_event_id: data.stableKey || '',
    st_nonjob_ids_json: JSON.stringify(data.stNonJobIds || []),
    last_hash: data.lastHash || '',
    last_synced_utc: data.lastSyncUtc || '',
    status: data.status || '',
  };
}

async function getEventMap(stableKey, mailboxKey = null) {
  const key = makeEventMapDocId(stableKey);
  const mailbox = mailboxKey ? normalizeMailboxKey(mailboxKey) : null;
  if (!key) return null;

  if (!isFirestoreEnabled()) {
    if (!mailbox) throw new Error('mailboxKey required for sheets fallback');
    return sheets.findEventMapping(mailbox, key);
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    // Primary lookup: mailbox-scoped key to prevent cross-mailbox collisions.
    if (mailbox) {
      const scopedKey = makeEventMapDocId(stableKey, mailbox);
      const scopedSnap = await db.collection(EVENT_MAP_COLLECTION).doc(scopedKey).get();
      const scopedMapped = mapFirestoreDoc(scopedSnap);
      if (scopedMapped) {
        return toLegacyLikeMapping(scopedMapped);
      }
    }

    // Backward compatibility: legacy unscoped doc id.
    const legacySnap = await db.collection(EVENT_MAP_COLLECTION).doc(key).get();
    const legacyMapped = mapFirestoreDoc(legacySnap);
    if (!legacyMapped) return null;
    if (mailbox && normalizeMailboxKey(legacyMapped.mailboxKey) !== mailbox) return null;
    return toLegacyLikeMapping(legacyMapped);
  }, `get event map ${key}`);
}

async function getEventMapByGraphId(mailboxKey, graphEventId) {
  const mailbox = normalizeMailboxKey(mailboxKey);
  const gid = String(graphEventId || '').trim();
  if (!mailbox || !gid) return null;

  if (!isFirestoreEnabled()) {
    return sheets.findEventMappingByGraphId(mailbox, gid);
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    const q = await db.collection(EVENT_MAP_COLLECTION)
      .where('mailboxKey', '==', mailbox)
      .where('graphEventId', '==', gid)
      .limit(1)
      .get();
    if (q.empty) return null;
    return toLegacyLikeMapping({ stableKey: q.docs[0].id, ...(q.docs[0].data() || {}) });
  }, `get event map by graph id ${mailbox}`);
}

async function upsertEventMap(
  mailboxKey,
  stableKey,
  stNonJobIds,
  lastHash,
  status = 'ACTIVE',
  graphEventId = null,
  metadata = {},
) {
  const mailbox = normalizeMailboxKey(mailboxKey);
  const key = makeEventMapDocId(stableKey);
  const scopedKey = makeEventMapDocId(stableKey, mailbox);
  const ids = Array.isArray(stNonJobIds) ? stNonJobIds.map((x) => String(x)) : [];
  const nowUtc = DateTime.utc().toISO();

  if (!mailbox || !key) throw new Error('Missing mailboxKey/stableKey');

  if (!isFirestoreEnabled()) {
    const existing = await sheets.findEventMapping(mailbox, key);
    await sheets.updateEventMapping(
      mailbox,
      key,
      ids,
      lastHash || '',
      status,
      existing ? existing.rowIndex : null,
    );
    return;
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    const payload = {
      stableKey: String(stableKey || '').trim(),
      mailboxKey: mailbox,
      stNonJobIds: ids,
      stNonJobId: ids[0] || null,
      status,
      lastSeenUtc: nowUtc,
      lastSyncUtc: nowUtc,
      lastHash: lastHash || '',
      hashKey: hashStableKey(key),
      graphEventId: graphEventId ? String(graphEventId) : null,
      startUtcIso: metadata.startUtcIso || null,
      endUtcIso: metadata.endUtcIso || null,
      iCalUId: metadata.iCalUId || null,
    };
    await db.collection(EVENT_MAP_COLLECTION).doc(scopedKey).set(payload, { merge: true });
  }, `upsert event map ${scopedKey}`);
}

async function markEventMapDeleted(mailboxKey, stableKey, lastHash = '') {
  const mailbox = normalizeMailboxKey(mailboxKey);
  const key = makeEventMapDocId(stableKey);
  const scopedKey = makeEventMapDocId(stableKey, mailbox);
  if (!mailbox || !key) throw new Error('Missing mailboxKey/stableKey');

  if (!isFirestoreEnabled()) {
    const existing = await sheets.findEventMapping(mailbox, key);
    if (!existing) return;
    await sheets.updateEventMapping(
      mailbox,
      key,
      [],
      lastHash || existing.last_hash || '',
      'DELETED',
      existing.rowIndex,
    );
    return;
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    await db.collection(EVENT_MAP_COLLECTION).doc(scopedKey).set({
      mailboxKey: mailbox,
      stableKey: String(stableKey || '').trim(),
      stNonJobIds: [],
      stNonJobId: null,
      status: 'DELETED',
      lastSyncUtc: DateTime.utc().toISO(),
      lastHash: lastHash || '',
      hashKey: hashStableKey(String(stableKey || '').trim()),
    }, { merge: true });
  }, `mark deleted ${scopedKey}`);
}

async function markEventMapRolledBack(mailboxKey, stableKey, message = 'rollback') {
  const mailbox = normalizeMailboxKey(mailboxKey);
  const key = makeEventMapDocId(stableKey);
  const scopedKey = makeEventMapDocId(stableKey, mailbox);
  if (!mailbox || !key) return;

  if (!isFirestoreEnabled()) {
    return;
  }

  return withFirestoreRetry(async () => {
    const db = firestoreService.getFirestore();
    await db.collection(EVENT_MAP_COLLECTION).doc(scopedKey).set({
      mailboxKey: mailbox,
      stableKey: String(stableKey || '').trim(),
      status: 'ROLLED_BACK',
      rollbackMessage: String(message || '').slice(0, 500),
      lastSyncUtc: DateTime.utc().toISO(),
      hashKey: hashStableKey(String(stableKey || '').trim()),
    }, { merge: true });
  }, `mark rolled back ${scopedKey}`);
}

module.exports = {
  isFirestoreEnabled,
  acquireLock,
  releaseLock,
  getDeltaState,
  setDeltaState,
  setLastFullReconcileUtc,
  getEventMap,
  getEventMapByGraphId,
  upsertEventMap,
  markEventMapDeleted,
  markEventMapRolledBack,
};
