const { DateTime } = require('luxon');
const { TIMEZONE } = require('./time');

function isSalesHuddleSubject(subject) {
  const subj = String(subject || '').trim().toLowerCase();
  return subj === 'sales huddle';
}

function getHuddleSlotKey(event) {
  if (!event || !event.start || !event.end) return null;
  const startLocal = DateTime.fromISO(event.start, { zone: 'utc' }).setZone(TIMEZONE);
  const endLocal = DateTime.fromISO(event.end, { zone: 'utc' }).setZone(TIMEZONE);
  if (!startLocal.isValid || !endLocal.isValid) return null;

  const weekday = startLocal.weekday; // 1=Mon ... 7=Sun
  const sh = startLocal.hour;
  const sm = startLocal.minute;
  const eh = endLocal.hour;
  const em = endLocal.minute;

  // Weekly patterns (local time).
  if (weekday === 1 && sh === 8 && sm === 30 && eh === 9 && em === 0) return 'mon_0830_0900';
  if (weekday === 4 && sh === 8 && sm === 30 && eh === 9 && em === 0) return 'thu_0830_0900';
  if (weekday === 2 && sh === 10 && sm === 0 && eh === 12 && em === 0) return 'tue_1000_1200';

  // Subject-matched huddles at unexpected times still count, but get a deterministic slot key.
  if (isSalesHuddleSubject(event.subject)) {
    const day = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'][weekday - 1] || 'unk';
    const pad2 = (n) => String(n).padStart(2, '0');
    return `subj_${day}_${pad2(sh)}${pad2(sm)}_${pad2(eh)}${pad2(em)}`;
  }

  return null;
}

module.exports = {
  isSalesHuddleSubject,
  getHuddleSlotKey,
};

