const test = require('node:test');
const assert = require('node:assert/strict');
const { getHuddleSlotKey } = require('../src/utils/huddle');

test('getHuddleSlotKey detects Tuesday 10-12 CT huddle from UTC timestamps', () => {
  const slot = getHuddleSlotKey({
    subject: 'Sales Huddle',
    // 10:00-12:00 America/Chicago is 16:00-18:00 UTC on 2026-02-03 (winter time).
    start: '2026-02-03T16:00:00.000Z',
    end: '2026-02-03T18:00:00.000Z',
  });

  assert.equal(slot, 'tue_1000_1200');
});

test('getHuddleSlotKey detects Monday 08:30-09:00 CT huddle', () => {
  const slot = getHuddleSlotKey({
    subject: 'Sales Huddle',
    // 08:30-09:00 America/Chicago is 14:30-15:00 UTC on 2026-02-02.
    start: '2026-02-02T14:30:00.000Z',
    end: '2026-02-02T15:00:00.000Z',
  });

  assert.equal(slot, 'mon_0830_0900');
});

