const { DateTime } = require('luxon');

const DEDUPE_KEY_VERSION = 'v3';

function normalizeGraphEvent(event) {
  const start = normalizeDateTime(event.start);
  const end = normalizeDateTime(event.end);

  return {
    id: event.id,
    iCalUId: event.iCalUId || '',
    subject: event.subject || '',
    start,
    end,
    isAllDay: Boolean(event.isAllDay),
    showAs: (event.showAs || 'busy').toLowerCase(),
    isPrivate: String(event.sensitivity || '').toLowerCase() === 'private',
    location: event.location?.displayName || '',
    attendees: (event.attendees || []).map((attendee) => ({
      email: attendee?.emailAddress?.address || '',
      name: attendee?.emailAddress?.name || '',
      type: attendee?.type || '',
    })),
    bodyPreview: event.bodyPreview || '',
    lastModifiedDateTime: normalizeDateTimeValue(event.lastModifiedDateTime),
    isRemoved: Boolean(event['@removed']),
  };
}

function normalizeDateTime(dateObj) {
  if (!dateObj || !dateObj.dateTime) {
    return null;
  }

  const zone = dateObj.timeZone || 'UTC';
  const dt = DateTime.fromISO(dateObj.dateTime, { zone });
  if (!dt.isValid) {
    return null;
  }
  return dt.toUTC().toISO();
}

function normalizeDateTimeValue(value) {
  if (!value) {
    return null;
  }

  const dt = DateTime.fromISO(value);
  if (!dt.isValid) {
    return null;
  }
  return dt.toUTC().toISO();
}

function getEventDedupeKey(event) {
  // Include key fields so logic changes (e.g., private masking) can force an update even when
  // lastModifiedDateTime is unchanged.
  return [
    DEDUPE_KEY_VERSION,
    event.id,
    event.iCalUId || '',
    event.lastModifiedDateTime || '',
    event.isPrivate ? 'P' : 'N',
    event.showAs || '',
    event.isAllDay ? 'A' : 'T',
    event.start || '',
    event.end || '',
    // Include the subject so changes propagate to ST (non-private only, since private events are masked).
    event.isPrivate ? '' : (event.subject || ''),
  ].join(':');
}

function getStableEventKey(event) {
  // Used for EventMap keying so calendarView and delta views map to the same occurrence.
  // Tombstones often lack start/end; callers must handle that separately.
  const uid = String(event.iCalUId || '').trim();
  const start = event.start || '';
  const end = event.end || '';
  if (uid && start && end) {
    return `${uid}:${start}:${end}`;
  }
  // Fallback: Graph id + times.
  return `${event.id || ''}:${start}:${end}`;
}

module.exports = {
  normalizeGraphEvent,
  getEventDedupeKey,
  getStableEventKey,
};
