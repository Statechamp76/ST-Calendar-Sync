const crypto = require('node:crypto');
const { DateTime } = require('luxon');

function normalizeIsoUtc(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const dt = DateTime.fromISO(s, { zone: 'utc' });
  if (!dt.isValid) return '';
  return dt.toUTC().toISO();
}

function makeStableKeyNormal(options = {}) {
  const {
    tenantId,
    technicianId,
    iCalUId,
    graphId,
    startUtcIso,
    endUtcIso,
  } = options;

  const t = String(tenantId || '').trim();
  const tech = String(technicianId || '').trim();
  const uid = String(iCalUId || '').trim();
  const gid = String(graphId || '').trim();
  const start = normalizeIsoUtc(startUtcIso);
  const end = normalizeIsoUtc(endUtcIso);

  if (!t) throw new Error('stableKey: missing tenantId');
  if (!tech) throw new Error('stableKey: missing technicianId');
  if (!start || !end) throw new Error('stableKey: missing start/end');

  const bestId = uid || gid;
  if (!bestId) throw new Error('stableKey: missing iCalUId/graphId');

  // Format invariant: {tenant}:{technicianId}:{iCalUId|graphId}:{startUtc}:{endUtc}
  return `${t}:${tech}:${bestId}:${start}:${end}`;
}

function makeStableKeyHuddle(options = {}) {
  const {
    tenantId,
    huddleSlotKey,
    startUtcIso,
    endUtcIso,
  } = options;

  const t = String(tenantId || '').trim();
  const slot = String(huddleSlotKey || '').trim();
  const start = normalizeIsoUtc(startUtcIso);
  const end = normalizeIsoUtc(endUtcIso);

  if (!t) throw new Error('stableKey: missing tenantId');
  if (!slot) throw new Error('stableKey: missing huddleSlotKey');
  if (!start || !end) throw new Error('stableKey: missing start/end');

  // Format invariant: {tenant}:HUDDLE:{slot}:{startUtc}:{endUtc}
  return `${t}:HUDDLE:${slot}:${start}:${end}`;
}

function hashStableKey(stableKey) {
  const s = String(stableKey || '');
  if (!s) return '';
  return crypto.createHash('sha256').update(s).digest('hex').slice(0, 16);
}

module.exports = {
  makeStableKeyNormal,
  makeStableKeyHuddle,
  hashStableKey,
};
