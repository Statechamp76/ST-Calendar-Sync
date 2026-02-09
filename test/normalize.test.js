const test = require('node:test');
const assert = require('node:assert/strict');
const { normalizeGraphEvent, getEventDedupeKey } = require('../src/utils/normalize');

test('normalizeGraphEvent returns expected normalized shape', () => {
  const event = {
    id: 'abc123',
    subject: 'Team Meeting',
    start: { dateTime: '2026-02-01T10:00:00', timeZone: 'America/Chicago' },
    end: { dateTime: '2026-02-01T11:00:00', timeZone: 'America/Chicago' },
    isAllDay: false,
    showAs: 'busy',
    location: { displayName: 'Conference Room A' },
    attendees: [
      { emailAddress: { address: 'a@example.com', name: 'A User' }, type: 'required' },
    ],
    bodyPreview: 'Agenda',
    lastModifiedDateTime: '2026-02-01T15:30:00Z',
    sensitivity: 'private',
  };

  const normalized = normalizeGraphEvent(event);

  assert.equal(normalized.id, 'abc123');
  assert.equal(normalized.subject, 'Team Meeting');
  assert.equal(normalized.isAllDay, false);
  assert.equal(normalized.showAs, 'busy');
  assert.equal(normalized.isPrivate, true);
  assert.equal(normalized.location, 'Conference Room A');
  assert.equal(normalized.attendees.length, 1);
  assert.equal(normalized.attendees[0].email, 'a@example.com');
  assert.match(normalized.start, /Z$/);
  assert.match(normalized.end, /Z$/);
  assert.equal(normalized.bodyPreview, 'Agenda');
  assert.equal(normalized.lastModifiedDateTime, '2026-02-01T15:30:00.000Z');
});

test('getEventDedupeKey uses id + lastModifiedDateTime', () => {
  const key = getEventDedupeKey({
    id: 'event-id',
    lastModifiedDateTime: '2026-02-01T15:30:00.000Z',
  });

  assert.equal(key, 'event-id:2026-02-01T15:30:00.000Z');
});
