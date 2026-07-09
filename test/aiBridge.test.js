const test = require('node:test');
const assert = require('node:assert/strict');
const express = require('express');
const request = require('supertest');
const { createAiBridgeRouter } = require('../src/routes/aiBridge');

function makeApp(storeOverride = {}) {
  const app = express();
  app.use(express.json());
  app.use('/ai', createAiBridgeRouter({ store: storeOverride }));
  return app;
}

function setBaseEnv() {
  process.env.AI_BRIDGE_ENABLED = 'true';
  process.env.AI_BRIDGE_API_KEY = 'test-api-key';
}

function validPayload() {
  return {
    project: 'ST-Calendar-Sync',
    type: 'decision',
    title: 'Sample Decision',
    tags: ['ops', 'sync'],
    content: 'Use per-user lock and delta sync.',
    source: 'codex',
  };
}

test('POST /ai/entries rejects missing API key', async () => {
  setBaseEnv();
  const app = makeApp({
    addEntry: async () => ({ id: 'x', project: 'ST-Calendar-Sync', createdAt: null }),
  });

  const response = await request(app).post('/ai/entries').send(validPayload());
  assert.equal(response.status, 401);
});

test('POST /ai/entries rejects missing required fields', async () => {
  setBaseEnv();
  const app = makeApp({
    addEntry: async () => ({ id: 'x', project: 'ST-Calendar-Sync', createdAt: null }),
  });

  const response = await request(app)
    .post('/ai/entries')
    .set('X-API-Key', 'test-api-key')
    .send({ project: 'ST-Calendar-Sync' });

  assert.equal(response.status, 400);
});

test('POST /ai/entries rejects content > 50k', async () => {
  setBaseEnv();
  const app = makeApp({
    addEntry: async () => ({ id: 'x', project: 'ST-Calendar-Sync', createdAt: null }),
  });

  const payload = validPayload();
  payload.content = 'a'.repeat(50_001);

  const response = await request(app)
    .post('/ai/entries')
    .set('X-API-Key', 'test-api-key')
    .send(payload);

  assert.equal(response.status, 400);
});

test('POST /ai/entries rejects disallowed type', async () => {
  setBaseEnv();
  const app = makeApp({
    addEntry: async () => ({ id: 'x', project: 'ST-Calendar-Sync', createdAt: null }),
  });

  const payload = validPayload();
  payload.type = 'calendar_dump';

  const response = await request(app)
    .post('/ai/entries')
    .set('X-API-Key', 'test-api-key')
    .send(payload);

  assert.equal(response.status, 400);
});

test('POST /ai/entries rejects payload containing blocked keys', async () => {
  setBaseEnv();
  const app = makeApp({
    addEntry: async () => ({ id: 'x', project: 'ST-Calendar-Sync', createdAt: null }),
  });

  const payload = validPayload();
  payload.raw = { attendees: [{ email: 'a@example.com' }] };

  const response = await request(app)
    .post('/ai/entries')
    .set('X-API-Key', 'test-api-key')
    .send(payload);

  assert.equal(response.status, 400);
});

test('POST /ai/entries accepts valid payload and calls addEntry', async () => {
  setBaseEnv();
  let addCalls = 0;
  const app = makeApp({
    addEntry: async () => {
      addCalls += 1;
      return { id: 'entry-123', project: 'ST-Calendar-Sync', createdAt: null };
    },
  });

  const response = await request(app)
    .post('/ai/entries')
    .set('X-API-Key', 'test-api-key')
    .send(validPayload());

  assert.equal(response.status, 201);
  assert.equal(response.body.id, 'entry-123');
  assert.equal(addCalls, 1);
});
