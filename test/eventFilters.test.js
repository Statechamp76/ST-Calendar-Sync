const test = require('node:test');
const assert = require('node:assert/strict');
const { shouldExcludeAllDayEvent } = require('../src/utils/eventFilters');

test('shouldExcludeAllDayEvent excludes birthday all-day events', () => {
  const event = { isAllDay: true, subject: 'John Birthday' };
  assert.equal(shouldExcludeAllDayEvent(event, []), true);
});

test('shouldExcludeAllDayEvent excludes configured holiday keywords', () => {
  const event = { isAllDay: true, subject: 'Company Holiday - Memorial Day' };
  assert.equal(shouldExcludeAllDayEvent(event, ['holiday', 'pto']), true);
});

test('shouldExcludeAllDayEvent keeps normal all-day events', () => {
  const event = { isAllDay: true, subject: 'Sales Kickoff' };
  assert.equal(shouldExcludeAllDayEvent(event, ['holiday']), false);
});
