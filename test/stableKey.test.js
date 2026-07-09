const test = require('node:test');
const assert = require('node:assert/strict');
const { makeStableKeyNormal, makeStableKeyHuddle, hashStableKey } = require('../src/utils/stableKey');

test('makeStableKeyNormal is deterministic and includes tenant+tech+id+times', () => {
  const k1 = makeStableKeyNormal({
    tenantId: 'T1',
    technicianId: '223',
    iCalUId: 'ical-xyz',
    graphId: 'graph-abc',
    startUtcIso: '2026-02-03T16:00:00.000Z',
    endUtcIso: '2026-02-03T18:00:00.000Z',
  });

  const k2 = makeStableKeyNormal({
    tenantId: 'T1',
    technicianId: '223',
    iCalUId: 'ical-xyz',
    graphId: 'graph-zzz',
    startUtcIso: '2026-02-03T16:00:00.000Z',
    endUtcIso: '2026-02-03T18:00:00.000Z',
  });

  assert.equal(k1, k2);
  assert.match(k1, /^T1:223:ical-xyz:2026-02-03T16:00:00\.000Z:2026-02-03T18:00:00\.000Z$/);
});

test('makeStableKeyHuddle uses HUDDLE marker and slot key', () => {
  const k = makeStableKeyHuddle({
    tenantId: 'T1',
    huddleSlotKey: 'tue_1000_1200',
    startUtcIso: '2026-02-03T16:00:00.000Z',
    endUtcIso: '2026-02-03T18:00:00.000Z',
  });

  assert.equal(k, 'T1:HUDDLE:tue_1000_1200:2026-02-03T16:00:00.000Z:2026-02-03T18:00:00.000Z');
});

test('hashStableKey is stable and redacts key contents', () => {
  const h = hashStableKey('T1:223:ical-xyz:2026-02-03T16:00:00.000Z:2026-02-03T18:00:00.000Z');
  assert.equal(h.length, 16);
  assert.match(h, /^[0-9a-f]{16}$/);
});

test('makeStableKeyNormal normalizes offset ISO timestamps to UTC Z', () => {
  const k = makeStableKeyNormal({
    tenantId: 'T1',
    technicianId: '223',
    iCalUId: 'ical-utc',
    startUtcIso: '2026-02-03T10:00:00-06:00',
    endUtcIso: '2026-02-03T12:00:00-06:00',
  });
  assert.match(k, /^T1:223:ical-utc:2026-02-03T16:00:00\.000Z:2026-02-03T18:00:00\.000Z$/);
});
