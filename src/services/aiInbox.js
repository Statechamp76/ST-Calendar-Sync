const crypto = require('crypto');
const sheets = require('./sheets');
const firestoreService = require('./firestore');

const SPREADSHEET_ID = '1DHL_hHvduwxUxm0uAtmCIgolwv6cqDSHUe0hJwW6na8';
const INBOX_SHEET = 'AI_INBOX';
const OUTBOX_SHEET = 'AI_OUTBOX';
const INBOX_HEADERS = ['id', 'created_utc', 'project', 'type', 'title', 'tags', 'content', 'source'];
const OUTBOX_WIDTH = 7;
const LOCK_COLLECTION = 'system_locks';
const LOCK_DOC_ID = 'aiInboxLock';
const DEFAULT_LOCK_TTL_SECONDS = 120;

function isoNow() {
  return new Date().toISOString();
}

function toSafeString(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

function parseTags(value) {
  const raw = toSafeString(value);
  if (!raw) return [];

  if (raw.startsWith('[') && raw.endsWith(']')) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return parsed.map((tag) => toSafeString(tag)).filter(Boolean);
      }
    } catch {
      // Fall back to CSV split below.
    }
  }

  return raw
    .split(',')
    .map((tag) => tag.trim())
    .filter(Boolean);
}

function makeOutboxId() {
  return `outbox_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`;
}

function makeLockHolder() {
  return `aiInbox_${process.pid}_${Date.now()}_${crypto.randomBytes(3).toString('hex')}`;
}

function getLockTtlSeconds() {
  const raw = Number.parseInt(String(process.env.AI_INBOX_LOCK_TTL_SECONDS || DEFAULT_LOCK_TTL_SECONDS), 10);
  if (!Number.isFinite(raw) || raw <= 0) return DEFAULT_LOCK_TTL_SECONDS;
  return raw;
}

function parseInboxRows(rawRows, { project = 'ST-Calendar-Sync', limit = 50 } = {}) {
  const capped = Math.min(Math.max(Number.parseInt(String(limit || 50), 10) || 50, 1), 200);
  const normalizedProject = toSafeString(project).toLowerCase();
  const result = [];

  for (let idx = 0; idx < rawRows.length; idx += 1) {
    const row = rawRows[idx] || [];
    const rowIndex = idx + 2; // Header row is 1.
    const id = toSafeString(row[0]);
    if (!id) continue;

    const rowProject = toSafeString(row[2]);
    if (!rowProject || rowProject.toLowerCase() !== normalizedProject) continue;

    result.push({
      rowIndex,
      id,
      createdUtc: toSafeString(row[1]),
      project: rowProject,
      type: toSafeString(row[3]),
      title: toSafeString(row[4]),
      tags: parseTags(row[5]),
      content: String(row[6] || ''),
      source: toSafeString(row[7]),
    });

    if (result.length >= capped) break;
  }

  return result;
}

async function readInbox(project = 'ST-Calendar-Sync', limit = 50, deps = {}) {
  const sheetsApi = deps.sheets || sheets;
  const rawRows = await sheetsApi.readSheetRows(`${INBOX_SHEET}!A2:H`, SPREADSHEET_ID);
  return parseInboxRows(rawRows, { project, limit });
}

async function writeOutbox(rows, deps = {}) {
  const sheetsApi = deps.sheets || sheets;
  const outboxRows = Array.isArray(rows) ? rows : [];
  if (outboxRows.length === 0) return;
  await sheetsApi.appendSheetRows(`${OUTBOX_SHEET}!A:G`, outboxRows, SPREADSHEET_ID);
}

async function clearInboxRows(rows, deps = {}) {
  const sheetsApi = deps.sheets || sheets;
  const inboxRows = Array.isArray(rows) ? rows : [];
  if (inboxRows.length === 0) return;

  const data = inboxRows.map((row) => ({
    range: `${INBOX_SHEET}!A${row.rowIndex}:H${row.rowIndex}`,
    values: [new Array(INBOX_HEADERS.length).fill('')],
  }));
  await sheetsApi.batchUpdateSheetRanges(data, SPREADSHEET_ID);
}

function buildOutboxRowFromInbox(row) {
  const titleBase = row.title || '(untitled)';
  try {
    if (!row.type) throw new Error('missing_type');
    if (!row.title) throw new Error('missing_title');
    if (!String(row.content || '').trim()) throw new Error('missing_content');

    const source = row.source || 'unknown';
    const tags = Array.isArray(row.tags) ? row.tags.join(', ') : '';
    const message = [
      'status=consumed',
      `source=${source}`,
      tags ? `tags=${tags}` : '',
      row.content ? `content=${row.content}` : '',
    ].filter(Boolean).join('\n');

    return [
      makeOutboxId(),
      isoNow(),
      row.project || 'ST-Calendar-Sync',
      row.type || 'note',
      `Consumed: ${titleBase}`.slice(0, 500),
      message,
      row.id,
    ];
  } catch (error) {
    return [
      makeOutboxId(),
      isoNow(),
      row.project || 'ST-Calendar-Sync',
      'incident',
      `Error consuming: ${titleBase}`.slice(0, 500),
      `status=error\nmessage=${error.message || 'malformed_row'}`,
      row.id,
    ];
  }
}

async function acquireProcessingLock(deps = {}) {
  const firestore = deps.firestore || firestoreService.getFirestore();
  const holder = makeLockHolder();
  const nowMs = Date.now();
  const ttlSeconds = getLockTtlSeconds();
  const expiresAtMs = nowMs + (ttlSeconds * 1000);
  const lockRef = firestore.collection(LOCK_COLLECTION).doc(LOCK_DOC_ID);

  const acquired = await firestore.runTransaction(async (tx) => {
    const snap = await tx.get(lockRef);
    const current = snap.exists ? (snap.data() || {}) : {};
    const currentExpiresAtMs = Number.parseInt(String(current.expiresAtMs || '0'), 10) || 0;
    if (currentExpiresAtMs > nowMs) {
      return false;
    }

    tx.set(lockRef, {
      holder,
      expiresAtMs,
      updatedAt: firestoreService.serverTimestamp(),
    }, { merge: true });
    return true;
  });

  return {
    acquired,
    holder,
    expiresAtMs,
  };
}

async function releaseProcessingLock(holder, deps = {}) {
  if (!holder) return;
  const firestore = deps.firestore || firestoreService.getFirestore();
  const lockRef = firestore.collection(LOCK_COLLECTION).doc(LOCK_DOC_ID);
  const nowMs = Date.now();

  await firestore.runTransaction(async (tx) => {
    const snap = await tx.get(lockRef);
    if (!snap.exists) return;

    const current = snap.data() || {};
    if (String(current.holder || '') !== String(holder)) return;

    tx.set(lockRef, {
      expiresAtMs: nowMs - 1,
      updatedAt: firestoreService.serverTimestamp(),
    }, { merge: true });
  });
}

async function processInbox({ project = 'ST-Calendar-Sync', limit = 50 } = {}, deps = {}) {
  const sheetsApi = deps.sheets || sheets;
  const lock = await acquireProcessingLock(deps);
  if (!lock.acquired) {
    return {
      project,
      read: 0,
      outboxWritten: 0,
      cleared: 0,
      skipped: true,
      reason: 'locked',
    };
  }

  try {
    const inboxRows = await readInbox(project, limit, { sheets: sheetsApi });

    if (inboxRows.length === 0) {
      return {
        project,
        read: 0,
        outboxWritten: 0,
        cleared: 0,
      };
    }

    const outboxRows = inboxRows.map((row) => buildOutboxRowFromInbox(row));
    // Required ordering: outbox append first; clear only after append succeeds.
    await writeOutbox(outboxRows, { sheets: sheetsApi });
    await clearInboxRows(inboxRows, { sheets: sheetsApi });

    return {
      project,
      read: inboxRows.length,
      outboxWritten: outboxRows.length,
      cleared: inboxRows.length,
    };
  } finally {
    await releaseProcessingLock(lock.holder, deps);
  }
}

module.exports = {
  SPREADSHEET_ID,
  INBOX_SHEET,
  OUTBOX_SHEET,
  OUTBOX_WIDTH,
  parseInboxRows,
  readInbox,
  writeOutbox,
  clearInboxRows,
  processInbox,
  acquireProcessingLock,
  releaseProcessingLock,
};
