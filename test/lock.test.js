const test = require('node:test');
const assert = require('node:assert/strict');
const { makeLockHolderId } = require('../src/utils/lock');

test('makeLockHolderId returns hex and is different across calls', () => {
  const a = makeLockHolderId();
  const b = makeLockHolderId();
  assert.match(a, /^[0-9a-f]{16}$/);
  assert.match(b, /^[0-9a-f]{16}$/);
  assert.notEqual(a, b);
});

