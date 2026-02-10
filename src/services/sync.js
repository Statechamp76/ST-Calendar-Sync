const { PubSub } = require('@google-cloud/pubsub');
const graph = require('../api/graph');
const servicetitan = require('../api/servicetitan');
const sheets = require('./sheets');
const { getSecrets } = require('../utils/secrets');
const { normalizeGraphEvent, getEventDedupeKey, getStableEventKey } = require('../utils/normalize');
const { mapEventToServiceTitanPayloads } = require('./mapping');
const { notifyFailure } = require('./alerts');
const { loadConfig } = require('../config');

const config = loadConfig();

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

function isSyncableBusyOrOof(showAs) {
  const value = (showAs || '').toLowerCase();
  return value === 'busy' || value === 'oof';
}

async function deleteMappedEvent(userUpn, outlookEventId, existingMapping) {
  const existingIds = parseJsonArray(existingMapping.st_nonjob_ids_json);
  for (const appointmentId of existingIds) {
    await servicetitan.deleteNonJob(appointmentId);
  }
  await sheets.deleteEventMapping(userUpn, outlookEventId, existingMapping);
}

async function upsertServiceTitanAppointments(userConfig, event, existingMapping) {
  const payloads = mapEventToServiceTitanPayloads(event, userConfig);
  const previousIds = existingMapping ? parseJsonArray(existingMapping.st_nonjob_ids_json) : [];
  const currentIds = [];

  for (let index = 0; index < payloads.length; index += 1) {
    const payload = payloads[index];
    let appointmentId = previousIds[index];

    if (appointmentId) {
      try {
        await servicetitan.updateNonJob(appointmentId, payload);
      } catch (error) {
        const previousId = appointmentId;
        appointmentId = await servicetitan.createNonJob(payload);
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
      appointmentId = await servicetitan.createNonJob(payload);
    }

    currentIds.push(appointmentId);
  }

  for (let index = payloads.length; index < previousIds.length; index += 1) {
    await servicetitan.deleteNonJob(previousIds[index]);
  }

  return currentIds;
}

async function processNormalizedEvent(userConfig, normalizedEvent, summary) {
  const stableKey = getStableEventKey(normalizedEvent);
  const existingMapping = await sheets.findEventMapping(userConfig.outlook_upn, stableKey);
  const dedupeKey = getEventDedupeKey(normalizedEvent);

  // Graph delta tombstones (`@removed`) must delete any mapped ST records.
  if (normalizedEvent.isRemoved) {
    // Tombstones usually only include Graph id; lookup via status gid=... marker.
    const mappingByGid = await sheets.findEventMappingByGraphId(userConfig.outlook_upn, normalizedEvent.id);
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

  // Do not sync available/free events; remove existing ST mapping if present.
  if (isAvailabilityEvent(normalizedEvent.showAs)) {
    if (existingMapping) {
      await deleteMappedEvent(userConfig.outlook_upn, stableKey, existingMapping);
    }
    summary.eventsSkipped += 1;
    return;
  }

  // Only sync events that are explicitly Busy or Out of Office. If we previously created an ST
  // appointment for an event that no longer matches this policy, remove it.
  if (!isSyncableBusyOrOof(normalizedEvent.showAs)) {
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
    await sheets.updateEventMapping(
      userConfig.outlook_upn,
      stableKey,
      appointmentIds,
      dedupeKey,
      `SYNCED|gid=${normalizedEvent.id}`,
      existingMapping ? existingMapping.rowIndex : null,
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
  const summary = createSummary();
  console.log('sync.delta.start', { userUpn });

  let userConfig = userConfigOverride;
  if (!userConfig) {
    const techMap = await sheets.getTechMap();
    userConfig = techMap.find((user) => user.outlook_upn === userUpn && user.enabled);
  }

  if (!userConfig) {
    console.log('sync.delta.skipped.user_not_enabled', { userUpn });
    summary.finishedAt = new Date().toISOString();
    return summary;
  }

  const deltaState = await sheets.getDeltaState(userUpn);
  const graphResponse = await graph.getDeltaEvents(userUpn, deltaState.delta_link, {
    pastDays: config.syncWindowPastDays,
    futureDays: config.syncWindowFutureDays,
  });
  const { events, nextDeltaLink } = graphResponse;
  summary.calendarsProcessed = 1;
  summary.eventsFetched = events.length;

  await processUserEvents(userConfig, events, summary);
  await sheets.updateDeltaState(userUpn, nextDeltaLink, deltaState.rowIndex);
  summary.finishedAt = new Date().toISOString();

  console.log('sync.delta.complete', summary);
  return summary;
}

async function runBackfillLast30DaysForUser(userUpn, userConfigOverride = null) {
  const summary = createSummary();
  console.log('sync.backfill30.start', { userUpn });

  let userConfig = userConfigOverride;
  if (!userConfig) {
    const techMap = await sheets.getTechMap();
    userConfig = techMap.find((user) => user.outlook_upn === userUpn && user.enabled);
  }

  if (!userConfig) {
    console.log('sync.backfill30.skipped.user_not_enabled', { userUpn });
    summary.finishedAt = new Date().toISOString();
    return summary;
  }

  // Full pull (not delta): last 30 days only.
  const events = await graph.getCalendarWindowEvents(userUpn, 30, 0);
  summary.calendarsProcessed = 1;
  summary.eventsFetched = events.length;

  await processUserEvents(userConfig, events, summary);
  summary.finishedAt = new Date().toISOString();

  console.log('sync.backfill30.complete', summary);
  return summary;
}

async function runBackfillNext90DaysForUser(userUpn, userConfigOverride = null) {
  const summary = createSummary();
  console.log('sync.backfill90.start', { userUpn });

  let userConfig = userConfigOverride;
  if (!userConfig) {
    const techMap = await sheets.getTechMap();
    userConfig = techMap.find((user) => user.outlook_upn === userUpn && user.enabled);
  }

  if (!userConfig) {
    console.log('sync.backfill90.skipped.user_not_enabled', { userUpn });
    summary.finishedAt = new Date().toISOString();
    return summary;
  }

  // Full pull (not delta): next 90 days only.
  const events = await graph.getCalendarWindowEvents(userUpn, 0, 90);
  summary.calendarsProcessed = 1;
  summary.eventsFetched = events.length;

  await processUserEvents(userConfig, events, summary);
  summary.finishedAt = new Date().toISOString();

  console.log('sync.backfill90.complete', summary);
  return summary;
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
    await pubsub.topic(topicName).publishMessage({ json: { upn: userConfig.outlook_upn } });
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
        userConfig.outlook_upn,
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
