const firestoreService = require('./firestore');

function timestampToIso(value) {
  if (!value || typeof value.toDate !== 'function') return null;
  try {
    return value.toDate().toISOString();
  } catch {
    return null;
  }
}

async function ensureProject(project) {
  const db = firestoreService.getFirestore();
  const projectRef = db.collection('ai_projects').doc(project);
  const snapshot = await projectRef.get();
  const now = firestoreService.serverTimestamp();
  const doc = {
    projectId: project,
    updatedAt: now,
  };
  if (!snapshot.exists) {
    doc.createdAt = now;
  }
  await projectRef.set(doc, { merge: true });
  return projectRef;
}

async function addEntry(payload) {
  const projectRef = await ensureProject(payload.project);
  const now = firestoreService.serverTimestamp();
  const entryData = {
    project: payload.project,
    type: payload.type,
    title: payload.title,
    tags: payload.tags,
    content: payload.content,
    source: payload.source,
    createdAt: now,
    updatedAt: now,
    hash: payload.hash || null,
    redactionFlags: Array.isArray(payload.redactionFlags) ? payload.redactionFlags : [],
  };

  const entryRef = await projectRef.collection('entries').add(entryData);
  const createdSnapshot = await entryRef.get();
  return {
    id: entryRef.id,
    project: payload.project,
    createdAt: timestampToIso(createdSnapshot.get('createdAt')),
  };
}

async function listEntries({ project, limit, type, tag }) {
  const db = firestoreService.getFirestore();
  let query = db
    .collection('ai_projects')
    .doc(project)
    .collection('entries')
    .orderBy('createdAt', 'desc')
    .limit(limit);

  if (type) {
    query = query.where('type', '==', type);
  }
  if (tag) {
    query = query.where('tags', 'array-contains', tag);
  }

  const snapshot = await query.get();
  return snapshot.docs.map((doc) => {
    const data = doc.data();
    return {
      id: doc.id,
      project: data.project,
      type: data.type,
      title: data.title,
      tags: Array.isArray(data.tags) ? data.tags : [],
      content: data.content,
      source: data.source,
      hash: data.hash || null,
      redactionFlags: Array.isArray(data.redactionFlags) ? data.redactionFlags : [],
      createdAt: timestampToIso(data.createdAt),
      updatedAt: timestampToIso(data.updatedAt),
    };
  });
}

async function getState(project) {
  const entries = await listEntries({ project, limit: 100, type: null, tag: null });
  const countsByType = {};
  for (const entry of entries) {
    countsByType[entry.type] = (countsByType[entry.type] || 0) + 1;
  }

  return {
    project,
    recentWindowCount: entries.length,
    countsByType,
    latestTitles: entries.slice(0, 5).map((entry) => ({
      id: entry.id,
      type: entry.type,
      title: entry.title,
      createdAt: entry.createdAt,
    })),
  };
}

module.exports = {
  addEntry,
  listEntries,
  getState,
};
