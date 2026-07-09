const test = require('node:test');
const assert = require('node:assert/strict');
const firestoreService = require('../src/services/firestore');
const aiBridgeStore = require('../src/services/aiBridgeStore');

test('aiBridgeStore.addEntry writes to Firestore entries.add', async () => {
  let addCalls = 0;

  const createdAtValue = {
    toDate: () => new Date('2026-02-12T00:00:00.000Z'),
  };

  const fakeEntryRef = {
    id: 'entry-firestore-1',
    get: async () => ({
      get: (field) => (field === 'createdAt' ? createdAtValue : null),
    }),
  };

  const fakeProjectRef = {
    get: async () => ({ exists: false }),
    set: async () => {},
    collection: () => ({
      add: async () => {
        addCalls += 1;
        return fakeEntryRef;
      },
    }),
  };

  const fakeDb = {
    collection: () => ({
      doc: () => fakeProjectRef,
    }),
  };

  const originalGetFirestore = firestoreService.getFirestore;
  const originalServerTimestamp = firestoreService.serverTimestamp;
  firestoreService.getFirestore = () => fakeDb;
  firestoreService.serverTimestamp = () => ({ __serverTimestamp: true });

  try {
    const result = await aiBridgeStore.addEntry({
      project: 'ST-Calendar-Sync',
      type: 'decision',
      title: 'Test title',
      tags: [],
      content: 'Test content',
      source: 'codex',
      hash: 'abc123',
      redactionFlags: [],
    });

    assert.equal(addCalls, 1);
    assert.equal(result.id, 'entry-firestore-1');
    assert.equal(result.project, 'ST-Calendar-Sync');
    assert.equal(result.createdAt, '2026-02-12T00:00:00.000Z');
  } finally {
    firestoreService.getFirestore = originalGetFirestore;
    firestoreService.serverTimestamp = originalServerTimestamp;
  }
});
