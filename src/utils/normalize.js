const { DateTime } = require('luxon');

function normalizeGraphEvent(event) {
  const start = normalizeDateTime(event.start);
  const end = normalizeDateTime(event.end);

  return {
    id: event.id,
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
    event.id,
    event.lastModifiedDateTime || '',
    event.isPrivate ? 'P' : 'N',
    event.showAs || '',
    event.isAllDay ? 'A' : 'T',
    event.start || '',
    event.end || '',
  ].join(':');
}

module.exports = {
  normalizeGraphEvent,
  getEventDedupeKey,
};
