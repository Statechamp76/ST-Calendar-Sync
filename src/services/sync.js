const { PubSub } = require('@google-cloud/pubsub');
const graph = require('../api/graph');
const servicetitan = require('../api/servicetitan');
const sheets = require('./sheets');
const { getSecrets } = require('../utils/secrets');
const { normalizeGraphEvent, getEventDedupeKey } = require('../utils/normalize');
const { mapEventToServiceTitanPayloads } = require('./mapping');
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
        appointmentId = await servicetitan.createNonJob(payload);
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
  const existingMapping = await sheets.findEventMapping(userConfig.outlook_upn, normalizedEvent.id);
  const dedupeKey = getEventDedupeKey(normalizedEvent);

  if (!normalizedEvent.start || !normalizedEvent.end) {
    summary.eventsSkipped += 1;
    return;
  }

  if (normalizedEvent.showAs === 'free') {
    if (existingMapping) {
      const existingIds = parseJsonArray(existingMapping.st_nonjob_ids_json);
      for (const appointmentId of existingIds) {
        await servicetitan.deleteNonJob(appointmentId);
      }
      await sheets.deleteEventMapping(userConfig.outlook_upn, normalizedEvent.id);
    }
    summary.eventsSkipped += 1;
    return;
  }

  if (existingMapping && existingMapping.last_hash === dedupeKey) {
    summary.eventsSkipped += 1;
    return;
  }

  const appointmentIds = await upsertServiceTitanAppointments(userConfig, normalizedEvent, existingMapping);
  await sheets.updateEventMapping(
    userConfig.outlook_upn,
    normalizedEvent.id,
    appointmentIds,
    dedupeKey,
    'SYNCED',
    existingMapping ? existingMapping.rowIndex : null,
  );
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

async function runDeltaSyncForUser(userUpn) {
  console.log('sync.delta.start', { userUpn });

  const techMap = await sheets.getTechMap();
  const userConfig = techMap.find((user) => user.outlook_upn === userUpn && user.enabled);
  if (!userConfig) {
    console.log('sync.delta.skipped.user_not_enabled', { userUpn });
    return;
  }

  const deltaState = await sheets.getDeltaState(userUpn);
  const graphResponse = await graph.getDeltaEvents(userUpn, deltaState.delta_link);
  const { events, nextDeltaLink } = graphResponse;
  const summary = createSummary();
  summary.calendarsProcessed = 1;
  summary.eventsFetched = events.length;

  await processUserEvents(userConfig, events, summary);
  await sheets.updateDeltaState(userUpn, nextDeltaLink, deltaState.rowIndex);
  summary.finishedAt = new Date().toISOString();

  console.log('sync.delta.complete', summary);
}

async function runFullSyncForAllUsers() {
  console.log('sync.full.enqueue.start');
  const techMap = await sheets.getTechMap();
  const pubsub = new PubSub();
  const topicName = 'graph-notifications';

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
      console.error('sync.subscriptions.renew.error', {
        userUpn: userConfig.outlook_upn,
        message: error.message,
      });
    }
  }
  console.log('sync.subscriptions.renew.complete');
}

async function runSyncCycle() {
  const summary = createSummary();
  console.log('sync.cycle.start', {
    windowPastDays: config.syncWindowPastDays,
    windowFutureDays: config.syncWindowFutureDays,
  });

  const techMap = await sheets.getTechMap();
  const enabledUsers = techMap.filter((user) => user.enabled);

  for (const userConfig of enabledUsers) {
    summary.calendarsProcessed += 1;
    try {
      const events = await graph.getCalendarWindowEvents(
        userConfig.outlook_upn,
        config.syncWindowPastDays,
        config.syncWindowFutureDays,
      );
      summary.eventsFetched += events.length;
      await processUserEvents(userConfig, events, summary);
      console.log('sync.cycle.user.complete', {
        userUpn: userConfig.outlook_upn,
        eventsFetched: events.length,
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
  runFullSyncForAllUsers,
  renewGraphSubscriptions,
  runSyncCycle,
};
