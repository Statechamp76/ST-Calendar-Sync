const test = require('node:test');
const assert = require('node:assert/strict');
const { normalizeUpn } = require('../src/utils/upn');

test('normalizeUpn trims and lowercases mailbox identifiers', () => {
  assert.equal(normalizeUpn('  MBrennan@ElevatedRoofing.com  '), 'mbrennan@elevatedroofing.com');
});

test('normalizeUpn handles empty input safely', () => {
  assert.equal(normalizeUpn(null), '');
  assert.equal(normalizeUpn(undefined), '');
  assert.equal(normalizeUpn('   '), '');
});
