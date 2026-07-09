function parsePositiveInt(value, defaultValue, keyName) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }

  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error(`Invalid ${keyName}: expected a non-negative integer`);
  }
  return parsed;
}

function splitCsv(value) {
  if (!value) {
    return [];
  }
  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function loadConfig() {
  const config = {
    port: Number.parseInt(process.env.PORT || '8080', 10),
    syncWindowPastDays: parsePositiveInt(process.env.SYNC_WINDOW_PAST_DAYS, 30, 'SYNC_WINDOW_PAST_DAYS'),
    syncWindowFutureDays: parsePositiveInt(process.env.SYNC_WINDOW_FUTURE_DAYS, 90, 'SYNC_WINDOW_FUTURE_DAYS'),
    runSyncAudience: process.env.RUN_SYNC_AUDIENCE || '',
    pubsubTopic: (process.env.PUBSUB_TOPIC || 'outlook-change-notifications').trim() || 'outlook-change-notifications',
    maintenanceMode: (String(process.env.MAINTENANCE_MODE || '').trim().toLowerCase() === 'true'),
    ignoreTuesday10am: (String(process.env.IGNORE_TUESDAY_10AM || 'true').trim().toLowerCase() === 'true'),
    // Default to disabling huddle sync when huddles are managed directly in ServiceTitan.
    // When true, we will not create/update huddle non-jobs; we only clean up any previously integration-created ones.
    disableHuddleSync: (String(process.env.DISABLE_HUDDLE_SYNC || 'true').trim().toLowerCase() === 'true'),
    disableBackfillLast30: (String(process.env.DISABLE_BACKFILL_LAST_30 || 'true').trim().toLowerCase() === 'true'),
    backfillNextStartUtc: (process.env.BACKFILL_NEXT_START_UTC || '2026-02-14T00:00:00Z').trim(),
    backfillNextDays: parsePositiveInt(process.env.BACKFILL_NEXT_DAYS, 7, 'BACKFILL_NEXT_DAYS'),
    stIntegrationApplicationGuid: (process.env.ST_INTEGRATION_APPLICATION_GUID || '8b4a4e3b-3c1f-4a05-9c7c-1b8a00c1c2b0').trim(),
    canonicalHuddleMailbox: (process.env.CANONICAL_HUDDLE_MAILBOX || 'MBrennan@elevatedroofing.com').trim(),
    salesHuddleUserUpns: splitCsv(process.env.SALES_HUDDLE_USER_UPNS),
    graphClientId: process.env.GRAPH_CLIENT_ID || '',
    graphClientSecret: process.env.GRAPH_CLIENT_SECRET || '',
    graphTenantId: process.env.GRAPH_TENANT_ID || '',
    serviceTitanClientId: process.env.SERVICETITAN_CLIENT_ID || '',
    serviceTitanClientSecret: process.env.SERVICETITAN_CLIENT_SECRET || '',
    serviceTitanTenantId: process.env.SERVICETITAN_TENANT_ID || '',
    googleSpreadsheetId: process.env.GOOGLE_SPREADSHEET_ID || '',
    outlookUserUpns: splitCsv(process.env.OUTLOOK_USER_UPNS),
    graphWebhookUrl: process.env.GRAPH_WEBHOOK_URL || '',
    graphClientState: process.env.GRAPH_CLIENT_STATE || '',
    firestoreEnabled: (String(process.env.FIRESTORE_ENABLED || 'false').trim().toLowerCase() === 'true'),
    firestoreProjectId: (process.env.FIRESTORE_PROJECT_ID || '').trim(),
    syncMode: (process.env.SYNC_MODE || 'EXCLUDE_FREE_ONLY').trim(),
    holidayExcludeKeywords: splitCsv(process.env.HOLIDAY_EXCLUDE_KEYWORDS || 'holiday,vacation'),
    lockTtlSeconds: parsePositiveInt(process.env.LOCK_TTL_SECONDS, 600, 'LOCK_TTL_SECONDS'),
    reconcileDaysAhead: parsePositiveInt(process.env.RECONCILE_DAYS_AHEAD, 90, 'RECONCILE_DAYS_AHEAD'),
  };

  const requiredKeys = [
    ['RUN_SYNC_AUDIENCE', config.runSyncAudience],
    ['GRAPH_CLIENT_ID', config.graphClientId],
    ['GRAPH_CLIENT_SECRET', config.graphClientSecret],
    ['GRAPH_TENANT_ID', config.graphTenantId],
    ['SERVICETITAN_CLIENT_ID', config.serviceTitanClientId],
    ['SERVICETITAN_CLIENT_SECRET', config.serviceTitanClientSecret],
    ['SERVICETITAN_TENANT_ID', config.serviceTitanTenantId],
    ['GOOGLE_SPREADSHEET_ID', config.googleSpreadsheetId],
  ];

  const missing = requiredKeys
    .filter(([, value]) => !value)
    .map(([key]) => key);

  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }

  return config;
}

module.exports = {
  loadConfig,
};
