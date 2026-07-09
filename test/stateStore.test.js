const test = require('node:test');
const assert = require('node:assert/strict');
const firestoreService = require('../src/services/firestore');
const stateStore = require('../src/services/stateStore');

function makeMockFirestore() {
  const docs = new Map();

  function keyFor(collection, id) {
    return `${collection}/${id}`;
  }

  return {
    collection: (name) => ({
      doc: (id) => ({ __collection: name, __id: id }),
    }),
    runTransaction: async (fn) => {
      const tx = {
        get: async (ref) => {
          const data = docs.get(keyFor(ref.__collection, ref.__id));
          return {
            exists: Boolean(data),
            data: () => data || {},
          };
        },
        set: (ref, payload, options = {}) => {
          const k = keyFor(ref.__collection, ref.__id);
          const existing = docs.get(k) || {};
          docs.set(k, options.merge ? { ...existing, ...payload } : payload);
        },
      };
      return fn(tx);
    },
  };
}

test('Firestore lock acquire/release semantics', async () => {
  const originalEnabled = process.env.FIRESTORE_ENABLED;
  const originalGetFirestore = firestoreService.getFirestore;
  const originalServerTimestamp = firestoreService.serverTimestamp;
  process.env.FIRESTORE_ENABLED = 'true';
  const mock = makeMockFirestore();
  firestoreService.getFirestore = () => mock;
  firestoreService.serverTimestamp = () => ({ __ts: true });

  try {
    const a = await stateStore.acquireLock('user@example.com', 'owner-a', 60);
    assert.equal(a.acquired, true);

    const b = await stateStore.acquireLock('user@example.com', 'owner-b', 60);
    assert.equal(b.acquired, false);

    const relWrong = await stateStore.releaseLock('user@example.com', 'owner-b');
    assert.equal(relWrong.released, false);

    const rel = await stateStore.releaseLock('user@example.com', 'owner-a');
    assert.equal(rel.released, true);

    const c = await stateStore.acquireLock('user@example.com', 'owner-c', 60);
    assert.equal(c.acquired, true);
  } finally {
    process.env.FIRESTORE_ENABLED = originalEnabled;
    firestoreService.getFirestore = originalGetFirestore;
    firestoreService.serverTimestamp = originalServerTimestamp;
  }
});
