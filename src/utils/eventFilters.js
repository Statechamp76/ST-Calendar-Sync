function includesCi(haystack, needle) {
  return String(haystack || '').toLowerCase().includes(String(needle || '').toLowerCase());
}

function isBirthdaySubject(subject) {
  return includesCi(subject, 'birthday');
}

function matchesHolidayKeyword(subject, keywords) {
  const s = String(subject || '');
  const list = Array.isArray(keywords) ? keywords : [];
  return list.some((k) => {
    const kw = String(k || '').trim();
    if (!kw) return false;
    return includesCi(s, kw);
  });
}

function shouldExcludeAllDayEvent(event, holidayKeywords) {
  if (!event || !event.isAllDay) return false;
  const subject = String(event.subject || '');
  if (isBirthdaySubject(subject)) return true;
  return matchesHolidayKeyword(subject, holidayKeywords);
}

module.exports = {
  isBirthdaySubject,
  matchesHolidayKeyword,
  shouldExcludeAllDayEvent,
};
