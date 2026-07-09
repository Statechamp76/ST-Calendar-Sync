const { PubSub } = require('@google-cloud/pubsub');
const graph = require('../api/graph');
const servicetitan = require('../api/servicetitan');
const sheets = require('./sheets');
const stateStore = require('./stateStore');
const { getSecrets } = require('../utils/secrets');
const { normalizeGraphEvent, getEventDedupeKey, getStableEventKey } = require('../utils/normalize');
const { mapEventToServiceTitanPayloads } = require('./mapping');
const { notifyFailure } = require('./alerts');
const { loadConfig } = require('../config');
const { DateTime } = require('luxon');
const { TIMEZONE } = require('../utils/time');
const { makeStableKeyNormal, hashStableKey } = require('../utils/stableKey');
const { getHuddleSlotKey } = require('../utils/huddle');
const { extractIntegrationExternalId } = require('../utils/externalData');
const { makeLockHolderId } = require('../utils/lock');
const { normalizeUpn } = require('../utils/upn');
const { shouldExcludeAllDayEvent } = require('../utils/eventFilters');

const config = loadConfig();
const LOCK_HOLDER_ID = makeLockHolderId();
const USER_LOCK_TTL_SECONDS = config.lockTtlSeconds || (10 * 60);

function createSummary() {
  return {
    startedAt: new Date().toISOString(),
    finishedAt: null,
    calendarsProcessed: 0,
    eventsFetched: 0,
    eventsUpserted: 0,
    eventsSkipped: 0,
    errors: [],
  };
}

function parseJsonArray(value) {
  if (!value) {
    return [];
  }
  try {
    return JSON.parse(value);
  } catch {
    return [];
  }
}

function isAvailabilityEvent(showAs) {
  const value = (showAs || '').toLowerCase();
  return value === 'free' || value === 'available';
}

function shouldSyncEvent(showAs) {
  // SYNC_MODE=EXCLUDE_FREE_ONLY: sync everything except free/available.
  const mode = String(config.syncMode || 'EXCLUDE_FREE_ONLY').trim().toUpperCase();
  if (mode === 'EXCLUDE_FREE_ONLY') {
    return !isAvailabilityEvent(showAs);
  }
  return !isAvailabilityEvent(showAs);
}

function normalizeUserConfigUpn(userConfig) {
  if (!userConfig) return userConfig;
  return {
    ...userConfig,
    outlook_upn: normalizeUpn(userConfig.outlook_upn),
  };
}

function isIgnoredTuesday10amEvent(event) {
  if (!config.ignoreTuesday10am) return false;
  if (!event || !event.start) return false;
  const dt = DateTime.fromISO(event.start, { zone: 'utc' }).setZone(TIMEZONE);
  if (!dt.isValid) return false;
  // Luxon weekday: 1=Mon, 2=Tue ... 7=Sun
  return dt.weekday === 2 && dt.hour === 10 && dt.minute === 0;
}

function computePayloadStartEndUtc(payload) {
  const start = DateTime.fromISO(payload.start);
  if (!start.isValid) throw new Error('Invalid payload.start');

  // duration is HH:mm:ss
  const parts = String(payload.duration || '').split(':').map((x) => Number.parseInt(x, 10));
  if (parts.length !== 3 || parts.some((n) => !Number.isFinite(n) || n < 0)) {
    throw new Error('Invalid payload.duration');
  }

  const [hh, mm, ss] = parts;
  const end = start.plus({ hours: hh, minutes: mm, seconds: ss });

  return {
    startUtcIso: start.toUTC().toISO(),
    endUtcIso: end.toUTC().toISO(),
  };
}

async function findExistingNonJobByStableKey(userConfig, payload, stableKey) {
  // Fall back path when Sheets mapping is missing/stale: scan ST non-jobs near this time and match our stableKey.
  // This is intentionally conservative to prevent duplicates even under sheet failures.
  const { startUtcIso, endUtcIso } = computePayloadStartEndUtc(payload);
  const stableKeyHash = hashStableKey(stableKey);

  // Search window: +/- 12 hours around the appointment. Keeps list sizes reasonable.
  const startWin = DateTime.fromISO(startUtcIso, { zone: 'utc' }).minus({ hours: 12 }).toISO();
  const endWin = DateTime.fromISO(endUtcIso, { zone: 'utc' }).plus({ hours: 12 }).toISO();

  const appts = await servicetitan.listNonJobs({
    technicianId: String(payload.technicianId),
    startsOnOrAfter: startWin,
    startsOnOrBefore: endWin,
    page: 1,
    pageSize: 200,
  });

  for (const appt of appts || []) {
    const id = appt && appt.id !== undefined ? String(appt.id) : null;
    if (!id) continue;

    try {
      const detail = await servicetitan.getNonJob(id);
      const ext = extractIntegrationExternalId(detail, config.stIntegrationApplicationGuid);
      if (ext && ext === stableKey) {
        console.log('sync.idempotency.hit', {
          technicianId: String(payload.technicianId),
          stableKeyHash,
          appointmentId: id,
        });
        return id;
      }
    } catch (e) {
      console.warn('sync.idempotency.lookup_detail_failed', {
        appointmentId: id,
        technicianId: String(payload.technicianId),
        stableKeyHash,
        message: e.message,
      });
    }
  }

  return null;
}

async function deleteMappedEvent(userUpn, outlookEventId, existingMapping) {
  const existingIds = parseJsonArray(existingMapping.st_nonjob_ids_json);
  for (const appointmentId of existingIds) {
    await servicetitan.deleteNonJob(appointmentId);
  }
  await stateStore.markEventMapDeleted(userUpn, outlookEventId, existingMapping.last_hash || '');
}

async function upsertOneNonJobByStableKey(payload, stableKey) {
  const stableKeyHash = hashStableKey(stableKey);
  const payloadWithIntegration = {
    ...payload,
    integrationApplicationGuid: config.stIntegrationApplicationGuid,
    integrationExternalId: stableKey,
  };

  const existingId = await findExistingNonJobByStableKey(null, payloadWithIntegration, stableKey);
  if (existingId) {
    await servicetitan.updateNonJob(existingId, payloadWithIntegration);
    return existingId;
  }

  const createdId = await servicetitan.createNonJob(payloadWithIntegration);
  console.log('sync.huddle.create', {
    technicianId: String(payload.technicianId),
    appointmentId: String(createdId),
    stableKeyHash,
  });
  return createdId;
}

async function deleteOneNonJobByStableKey(payload, stableKey) {
  const stableKeyHash = hashStableKey(stableKey);
  const existingId = await findExistingNonJobByStableKey(null, payload, stableKey);
  if (!existingId) return false;
  await servicetitan.deleteNonJob(existingId);
  console.log('sync.huddle.delete', {
    technicianId: String(payload.technicianId),
    appointmentId: String(existingId),
    stableKeyHash,
  });
  return true;
}

async function upsertServiceTitanAppointments(userConfig, event, existingMapping) {
  const payloads = mapEventToServiceTitanPayloads(event, userConfig);
  const previousIds = existingMapping ? parseJsonArray(existingMapping.st_nonjob_ids_json) : [];
  const currentIds = [];

  for (let index = 0; index < payloads.length; index += 1) {
    const payload = payloads[index];
    const { startUtcIso, endUtcIso } = computePayloadStartEndUtc(payload);
    const stableKey = makeStableKeyNormal({
      tenantId: config.serviceTitanTenantId,
      technicianId: payload.technicianId,
      iCalUId: event.iCalUId,
      graphId: event.id,
      startUtcIso,
      endUtcIso,
    });
    const stableKeyHash = hashStableKey(stableKey);

    const payloadWithIntegration = {
      ...payload,
      integrationApplicationGuid: config.stIntegrationApplicationGuid,
      integrationExternalId: stableKey,
    };

    let appointmentId = previousIds[index];

    if (!appointmentId) {
      appointmentId = await findExistingNonJobByStableKey(userConfig, payloadWithIntegration, stableKey);
    }

    if (appointmentId) {
      try {
        await servicetitan.updateNonJob(appointmentId, payloadWithIntegration);
      } catch (error) {
        const previousId = appointmentId;
        appointmentId = await servicetitan.createNonJob(payloadWithIntegration);
        try {
          await servicetitan.deleteNonJob(previousId);
        } catch (deleteError) {
          console.warn('sync.upsert.reconcile.delete_previous_failed', {
            appointmentId: previousId,
            message: deleteError.message,
          });
        }
      }
    } else {
      appointmentId = await servicetitan.createNonJob(payloadWithIntegration);
    }

    // Never log subjects; stableKey contains iCalUId, so log hash only.
    console.log('sync.upsert.block.complete', {
      technicianId: String(payload.technicianId),
      appointmentId: String(appointmentId),
      stableKeyHash,
      startUtcIso,
      endUtcIso,
    });
    currentIds.push(appointmentId);
  }

  for (let index = payloads.length; index < previousIds.length; index += 1) {
    await servicetitan.deleteNonJob(previousIds[index]);
  }

  return currentIds;
}

async function processNormalizedEvent(userConfig, normalizedEvent, summary) {
  const stableKey = getStableEventKey(normalizedEvent);
  const existingMapping = await stateStore.getEventMap(stableKey, userConfig.outlook_upn);
  const dedupeKey = getEventDedupeKey(normalizedEvent);

  // Graph delta tombstones (`@removed`) must delete any mapped ST records.
  if (normalizedEvent.isRemoved) {
    // Tombstones usually only include Graph id; lookup via status gid=... marker.
    const mappingByGid = await stateStore.getEventMapByGraphId(userConfig.outlook_upn, normalizedEvent.id);
    if (mappingByGid) {
      await deleteMappedEvent(userConfig.outlook_upn, mappingByGid.outlook_event_id, mappingByGid);
    }
    summary.eventsSkipped += 1;
    return;
  }

  if (!normalizedEvent.start || !normalizedEvent.end) {
    summary.eventsSkipped += 1;
    return;
  }

  // Include all-day events except birthdays or configured holiday keywords.
  if (shouldExcludeAllDayEvent(normalizedEvent, config.holidayExcludeKeywords)) {
    if (existingMapping) {
      await deleteMappedEvent(userConfig.outlook_upn, stableKey, existingMapping);
    }
    summary.eventsSkipped += 1;
    return;
  }

  // Group huddle collapsing: only the canonical mailbox is allowed to drive these events.
  // For other mailboxes, we skip and also try to delete any previously-created huddle non-jobs for that tech.
  const huddleSlotKey = getHuddleSlotKey(normalizedEvent);
  if (huddleSlotKey) {
    // If huddles are created directly in ServiceTitan, do not create/update them from Outlook at all.
    // We only clean up any integration-created artifacts (externalData-marked) and remove sheet mappings.
    if (config.disableHuddleSync) {
      if (existingMapping) {
        await deleteMappedEvent(userConfig.outlook_upn, stableKey, existingMapping);
      }

      const payloads = mapEventToServiceTitanPayloads(normalizedEvent, userConfig);
      for (const payload of payloads) {
        payload.name = normalizedEvent.isPrivate ? 'Private' : 'Sales Huddle';
        const { startUtcIso, endUtcIso } = computePayloadStartEndUtc(payload);
        const nonJobStableKey = makeStableKeyNormal({
          tenantId: config.serviceTitanTenantId,
          technicianId: payload.technicianId,
          iCalUId: normalizedEvent.iCalUId,
          graphId: normalizedEvent.id,
          startUtcIso,
          endUtcIso,
        });
        try {
          await deleteOneNonJobByStableKey(payload, nonJobStableKey);
        } catch (e) {
          console.warn('sync.huddle.disabled.cleanup_failed', {
            technicianId: String(payload.technicianId),
            stableKeyHash: hashStableKey(nonJobStableKey),
            message: e.message,
          });
        }
      }

      summary.eventsSkipped += 1;
      return;
    }

    const canonicalUpn = String(config.canonicalHuddleMailbox || '').trim().toLowerCase();
    const thisUpn = String(userConfig.outlook_upn || '').trim().toLowerCase();

    if (thisUpn !== canonicalUpn) {
      if (existingMapping) {
        await deleteMappedEvent(userConfig.outlook_upn, stableKey, existingMapping);
      }

      const payloads = mapEventToServiceTitanPayloads(normalizedEvent, userConfig);
      for (const payload of payloads) {
        payload.name = normalizedEvent.isPrivate ? 'Private' : 'Sales Huddle';
        const { startUtcIso, endUtcIso } = computePayloadStartEndUtc(payload);
        const nonJobStableKey = makeStableKeyNormal({
          tenantId: config.serviceTitanTenantId,
          technicianId: payload.technicianId,
          iCalUId: normalizedEvent.iCalUId,
          graphId: normalizedEvent.id,
          startUtcIso,
          endUtcIso,
        });
        try {
          await deleteOneNonJobByStableKey(payload, nonJobStableKey);
        } catch (e) {
          console.warn('sync.huddle.non_canonical.cleanup_failed', {
            technicianId: String(payload.technicianId),
            stableKeyHash: hashStableKey(nonJobStableKey),
            message: e.message,
          });
        }
      }

      summary.eventsSkipped += 1;
      return;
    }

    const techMap = await sheets.getTechMap();
    const targetUpns = (config.salesHuddleUserUpns && config.salesHuddleUserUpns.length > 0)
      ? config.salesHuddleUserUpns.map((u) => String(u).trim().toLowerCase()).filter(Boolean)
      : [canonicalUpn];

    const targets = techMap
      .filter((u) => u.enabled)
      .filter((u) => targetUpns.includes(String(u.outlook_upn || '').trim().toLowerCase()))
      .filter((u) => u.st_technician_id);

    for (const target of targets) {
      const payloads = mapEventToServiceTitanPayloads(normalizedEvent, target);
      for (const payload of payloads) {
        // Force a consistent, non-sensitive title for huddles.
        payload.name = normalizedEvent.isPrivate ? 'Private' : 'Sales Huddle';

        const { startUtcIso, endUtcIso } = computePayloadStartEndUtc(payload);
        const nonJobStableKey = makeStableKeyNormal({
          tenantId: config.serviceTitanTenantId,
          technicianId: payload.technicianId,
          iCalUId: normalizedEvent.iCalUId,
          graphId: normalizedEvent.id,
          startUtcIso,
          endUtcIso,
        });

        await upsertOneNonJobByStableKey(payload, nonJobStableKey);
      }
    }

    summary.eventsUpserted += 1;
    return;
  }

  // SYNC_MODE=EXCLUDE_FREE_ONLY.
  if (!shouldSyncEvent(normalizedEvent.showAs)) {
    if (existingMapping) {
      await deleteMappedEvent(userConfig.outlook_upn, stableKey, existingMapping);
    }
    summary.eventsSkipped += 1;
    return;
  }

  if (existingMapping && existingMapping.last_hash === dedupeKey) {
    summary.eventsSkipped += 1;
    return;
  }

  const appointmentIds = await upsertServiceTitanAppointments(userConfig, normalizedEvent, existingMapping);
  try {
    await stateStore.upsertEventMap(
      userConfig.outlook_upn,
      stableKey,
      appointmentIds,
      dedupeKey,
      'ACTIVE',
      normalizedEvent.id,
      {
        iCalUId: normalizedEvent.iCalUId || null,
        startUtcIso: normalizedEvent.start || null,
        endUtcIso: normalizedEvent.end || null,
      },
    );
  } catch (error) {
    // If we created ST records but could not record the mapping, delete the ST records so we don't
    // create orphan duplicates on the next run.
    if (!existingMapping) {
      for (const appointmentId of appointmentIds) {
        try {
          await servicetitan.deleteNonJob(appointmentId);
        } catch (deleteError) {
          console.warn('sync.mapping_failed.rollback.delete_failed', {
            appointmentId,
            message: deleteError.message,
          });
        }
      }
    }
    await stateStore.markEventMapRolledBack(userConfig.outlook_upn, stableKey, error.message);
    throw error;
  }
  summary.eventsUpserted += 1;
}

async function processUserEvents(userConfig, rawEvents, summary) {
  const seen = new Set();
  const normalizedEvents = rawEvents.map(normalizeGraphEvent).filter((event) => Boolean(event.id));

  for (const event of normalizedEvents) {
    const dedupeKey = getEventDedupeKey(event);
    if (seen.has(dedupeKey)) {
      summary.eventsSkipped += 1;
      continue;
    }
    seen.add(dedupeKey);

    try {
      await processNormalizedEvent(userConfig, event, summary);
    } catch (error) {
      summary.errors.push({
        userUpn: userConfig.outlook_upn,
        eventId: event.id,
        message: error.message,
      });
    }
  }
}

async function runDeltaSyncForUser(userUpn, userConfigOverride = null) {
  const normalizedUpn = normalizeUpn(userUpn);
  const summary = createSummary();
  console.log('sync.delta.start', { userUpn: normalizedUpn });

  const lock = await stateStore.acquireLock(normalizedUpn, LOCK_HOLDER_ID, USER_LOCK_TTL_SECONDS);
  if (!lock.acquired) {
    console.warn('sync.delta.skipped.locked', { userUpn: normalizedUpn, reason: lock.reason });
    summary.finishedAt = new Date().toISOString();
    return summary;
  }

  let userConfig = normalizeUserConfigUpn(userConfigOverride);
  if (!userConfig) {
    const techMap = await sheets.getTechMap();
    userConfig = techMap.find((user) => normalizeUpn(user.outlook_upn) === normalizedUpn && user.enabled);
  }

  if (!userConfig) {
    console.log('sync.delta.skipped.user_not_enabled', { userUpn: normalizedUpn });
    summary.finishedAt = new Date().toISOString();
    await stateStore.releaseLock(normalizedUpn, LOCK_HOLDER_ID);
    return summary;
  }

  try {
    const deltaState = await stateStore.getDeltaState(normalizedUpn);
    const graphResponse = await graph.getDeltaEvents(normalizedUpn, deltaState.deltaToken, {
      pastDays: config.syncWindowPastDays,
      futureDays: config.syncWindowFutureDays,
    });
    const { events, nextDeltaLink } = graphResponse;
    summary.calendarsProcessed = 1;
    summary.eventsFetched = events.length;

    await processUserEvents(userConfig, events, summary);
    await stateStore.setDeltaState(normalizedUpn, nextDeltaLink, deltaState.rowIndex);
    summary.finishedAt = new Date().toISOString();

    console.log('sync.delta.complete', summary);
    return summary;
  } finally {
    await stateStore.releaseLock(normalizedUpn, LOCK_HOLDER_ID);
  }
}

async function runBackfillLast30DaysForUser(userUpn, userConfigOverride = null) {
  const normalizedUpn = normalizeUpn(userUpn);
  const summary = createSummary();
  console.log('sync.backfill30.start', { userUpn: normalizedUpn });

  const lock = await stateStore.acquireLock(normalizedUpn, LOCK_HOLDER_ID, USER_LOCK_TTL_SECONDS);
  if (!lock.acquired) {
    console.warn('sync.backfill30.skipped.locked', { userUpn: normalizedUpn, reason: lock.reason });
    summary.finishedAt = new Date().toISOString();
    return summary;
  }

  let userConfig = normalizeUserConfigUpn(userConfigOverride);
  if (!userConfig) {
    const techMap = await sheets.getTechMap();
    userConfig = techMap.find((user) => normalizeUpn(user.outlook_upn) === normalizedUpn && user.enabled);
  }

  if (!userConfig) {
    console.log('sync.backfill30.skipped.user_not_enabled', { userUpn: normalizedUpn });
    summary.finishedAt = new Date().toISOString();
    await stateStore.releaseLock(normalizedUpn, LOCK_HOLDER_ID);
    return summary;
  }

  try {
    // Full pull (not delta): last 30 days only.
    const events = await graph.getCalendarWindowEvents(normalizedUpn, 30, 0);
    summary.calendarsProcessed = 1;
    summary.eventsFetched = events.length;

    await processUserEvents(userConfig, events, summary);
    summary.finishedAt = new Date().toISOString();

    console.log('sync.backfill30.complete', summary);
    return summary;
  } finally {
    await stateStore.releaseLock(normalizedUpn, LOCK_HOLDER_ID);
  }
}

async function runBackfillNext90DaysForUser(userUpn, userConfigOverride = null) {
  const normalizedUpn = normalizeUpn(userUpn);
  const summary = createSummary();
  console.log('sync.backfill90.start', { userUpn: normalizedUpn });

  const lock = await stateStore.acquireLock(normalizedUpn, LOCK_HOLDER_ID, USER_LOCK_TTL_SECONDS);
  if (!lock.acquired) {
    console.warn('sync.backfill90.skipped.locked', { userUpn: normalizedUpn, reason: lock.reason });
    summary.finishedAt = new Date().toISOString();
    return summary;
  }

  let userConfig = normalizeUserConfigUpn(userConfigOverride);
  if (!userConfig) {
    const techMap = await sheets.getTechMap();
    userConfig = techMap.find((user) => normalizeUpn(user.outlook_upn) === normalizedUpn && user.enabled);
  }

  if (!userConfig) {
    console.log('sync.backfill90.skipped.user_not_enabled', { userUpn: normalizedUpn });
    summary.finishedAt = new Date().toISOString();
    await stateStore.releaseLock(normalizedUpn, LOCK_HOLDER_ID);
    return summary;
  }

  try {
    // Full pull (not delta): temporary test window controlled by config.
    const startUtc = DateTime.fromISO(config.backfillNextStartUtc, { zone: 'utc' });
    if (!startUtc.isValid) {
      throw new Error(`Invalid BACKFILL_NEXT_START_UTC: ${config.backfillNextStartUtc}`);
    }
    const endUtc = startUtc.plus({ days: config.backfillNextDays });
    const events = await graph.getCalendarEventsBetween(normalizedUpn, startUtc.toISO(), endUtc.toISO());
    summary.calendarsProcessed = 1;
    summary.eventsFetched = events.length;

    await processUserEvents(userConfig, events, summary);
    summary.finishedAt = new Date().toISOString();

    console.log('sync.backfill90.complete', summary);
    return summary;
  } finally {
    await stateStore.releaseLock(normalizedUpn, LOCK_HOLDER_ID);
  }
}

async function runBackfillNext90DaysAllUsers() {
  const summary = createSummary();
  console.log('sync.backfill90.all.start');

  const techMap = await sheets.getTechMap();
  const enabledUsers = techMap.filter((user) => user.enabled);

  for (const userConfig of enabledUsers) {
    try {
      const userSummary = await runBackfillNext90DaysForUser(userConfig.outlook_upn, userConfig);
      summary.calendarsProcessed += userSummary.calendarsProcessed;
      summary.eventsFetched += userSummary.eventsFetched;
      summary.eventsUpserted += userSummary.eventsUpserted;
      summary.eventsSkipped += userSummary.eventsSkipped;
      summary.errors.push(...userSummary.errors);
      console.log('sync.backfill90.user.complete', {
        userUpn: userConfig.outlook_upn,
        eventsFetched: userSummary.eventsFetched,
        eventsUpserted: userSummary.eventsUpserted,
        eventsSkipped: userSummary.eventsSkipped,
      });
    } catch (error) {
      summary.errors.push({
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
      console.error('sync.backfill90.user.error', {
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
    }
  }

  summary.finishedAt = new Date().toISOString();
  console.log('sync.backfill90.all.complete', summary);
  return summary;
}

async function runBackfillLast30DaysAllUsers() {
  const summary = createSummary();
  console.log('sync.backfill30.all.start');

  const techMap = await sheets.getTechMap();
  const enabledUsers = techMap.filter((user) => user.enabled);

  for (const userConfig of enabledUsers) {
    try {
      const userSummary = await runBackfillLast30DaysForUser(userConfig.outlook_upn, userConfig);
      summary.calendarsProcessed += userSummary.calendarsProcessed;
      summary.eventsFetched += userSummary.eventsFetched;
      summary.eventsUpserted += userSummary.eventsUpserted;
      summary.eventsSkipped += userSummary.eventsSkipped;
      summary.errors.push(...userSummary.errors);
      console.log('sync.backfill30.user.complete', {
        userUpn: userConfig.outlook_upn,
        eventsFetched: userSummary.eventsFetched,
        eventsUpserted: userSummary.eventsUpserted,
        eventsSkipped: userSummary.eventsSkipped,
      });
    } catch (error) {
      summary.errors.push({
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
      console.error('sync.backfill30.user.error', {
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
    }
  }

  summary.finishedAt = new Date().toISOString();
  console.log('sync.backfill30.all.complete', summary);
  return summary;
}

async function runFullSyncForAllUsers() {
  console.log('sync.full.enqueue.start');
  const techMap = await sheets.getTechMap();
  const pubsub = new PubSub();
  const topicName = config.pubsubTopic;

  for (const userConfig of techMap) {
    if (!userConfig.enabled) {
      continue;
    }
    const upn = normalizeUpn(userConfig.outlook_upn);
    await pubsub.topic(topicName).publishMessage({ json: { upn }, orderingKey: upn });
  }

  console.log('sync.full.enqueue.complete');
}

async function renewGraphSubscriptions() {
  console.log('sync.subscriptions.renew.start');
  const techMap = await sheets.getTechMap();
  const secrets = await getSecrets(['GRAPH_WEBHOOK_URL', 'GRAPH_CLIENT_STATE']);
  const errors = [];

  for (const userConfig of techMap) {
    if (!userConfig.enabled) {
      continue;
    }
    try {
      await graph.createOrRenewSubscription(
        normalizeUpn(userConfig.outlook_upn),
        secrets.GRAPH_WEBHOOK_URL,
        secrets.GRAPH_CLIENT_STATE,
      );
    } catch (error) {
      errors.push({
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
      console.error('sync.subscriptions.renew.error', JSON.stringify({
        userUpn: userConfig.outlook_upn,
        message: error.message,
      }));
    }
  }
  if (errors.length > 0) {
    await notifyFailure('ST Calendar Sync: subscription renewal errors', {
      errorCount: errors.length,
      sample: errors.slice(0, 5),
    });
  }
  console.log('sync.subscriptions.renew.complete');
}

async function runSyncCycle() {
  const summary = createSummary();
  console.log('sync.cycle.start', { mode: 'delta' });

  const techMap = await sheets.getTechMap();
  const enabledUsers = techMap.filter((user) => user.enabled);

  for (const userConfig of enabledUsers) {
    try {
      const userSummary = await runDeltaSyncForUser(
        userConfig.outlook_upn,
        userConfig,
      );
      summary.calendarsProcessed += userSummary.calendarsProcessed;
      summary.eventsFetched += userSummary.eventsFetched;
      summary.eventsUpserted += userSummary.eventsUpserted;
      summary.eventsSkipped += userSummary.eventsSkipped;
      summary.errors.push(...userSummary.errors);
      console.log('sync.cycle.user.complete', {
        userUpn: userConfig.outlook_upn,
        eventsFetched: userSummary.eventsFetched,
        eventsUpserted: userSummary.eventsUpserted,
        eventsSkipped: userSummary.eventsSkipped,
      });
    } catch (error) {
      summary.errors.push({
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
      console.error('sync.cycle.user.error', {
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
    }
  }

  summary.finishedAt = new Date().toISOString();
  console.log('sync.cycle.complete', summary);
  return summary;
}

module.exports = {
  runDeltaSyncForUser,
  runBackfillLast30DaysForUser,
  runBackfillLast30DaysAllUsers,
  runBackfillNext90DaysForUser,
  runBackfillNext90DaysAllUsers,
  runFullSyncForAllUsers,
  renewGraphSubscriptions,
  runSyncCycle,
};
